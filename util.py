"""Standalone data helpers extracted from quantium for monthly_report.py.

The report example should not import quantium/vnpy or read cached pickle files.
These helpers keep the required database initialization logic local to the
example. Database connection settings used by this report are embedded below,
with optional environment-variable overrides such as WIND_DB_HOST.
"""

from __future__ import annotations

import datetime
import gc
import os
import time

import numpy as np
import pandas as pd

try:
    from tqdm import tqdm
except ImportError:

    def tqdm(iterable, **kwargs):
        return iterable

try:
    import pymssql
except ImportError:
    pymssql = None

try:
    import pymysql
except ImportError:
    pymysql = None

try:
    import cx_Oracle
except ImportError:
    cx_Oracle = None

try:
    import clickhouse_driver as clickhouse
    from clickhouse_driver import connect as clickhouse_connect
except ImportError:
    clickhouse = None
    clickhouse_connect = None

os.environ.setdefault("NLS_LANG", "AMERICAN_AMERICA.ZHS16GBK")

EMBEDDED_DB_CONFIG = {
    "WIND_DB": {
        "db_type": "ORACLE",
        "user": "chaxun",
        "password": "chaxun123",
        "host": "10.3.80.206",
        "port": 1521,
        "database": "winddata",
    },
    "JY_DB": {
        "db_type": "MSSQL",
        "host": "10.3.80.201",
        "user": "chaxun",
        "password": "chaxun123",
        "database": "JYDB",
    },
    "SAM_CH1": {
        "db_type": "CLICKHOUSE",
        "host": "10.3.8.1",
        "port": 9000,
        "user": "default",
        "password": "",
    },
}


def _drivers():
    drivers = {}
    if pymssql is not None:
        drivers["MSSQL"] = pymssql.connect
    if pymysql is not None:
        drivers["MYSQL"] = pymysql.Connect
    if cx_Oracle is not None:
        drivers["ORACLE"] = cx_Oracle.connect
    if clickhouse is not None:
        drivers["CLICKHOUSE"] = clickhouse.connect
    return drivers


def read_db_config(section, filename=None, db_name=None, db_type=None):
    if filename is not None:
        raise ValueError("monthly_report 的独立 util.py 不再读取外部配置文件，请直接修改 EMBEDDED_DB_CONFIG。")

    section = section.upper()
    prefix = section.upper()
    if section in EMBEDDED_DB_CONFIG:
        db = EMBEDDED_DB_CONFIG[section].copy()
    else:
        db = {}

    env_overrides = {
        "host": os.environ.get(f"{prefix}_HOST"),
        "user": os.environ.get(f"{prefix}_USER"),
        "password": os.environ.get(f"{prefix}_PASSWORD"),
        "port": os.environ.get(f"{prefix}_PORT"),
        "database": os.environ.get(f"{prefix}_DATABASE"),
        "db_type": os.environ.get(f"{prefix}_DB_TYPE"),
    }
    for key, value in env_overrides.items():
        if value not in (None, ""):
            db[key] = int(value) if isinstance(value, str) and value.isdigit() else value

    if not db or "host" not in db:
        raise KeyError(f"未找到内置数据库配置: {section}")

    if db_name is not None:
        db["database"] = db_name
    db_type = db.get("db_type", "MYSQL").upper() if db_type is None else db_type.upper()
    db.pop("db_type", None)
    if db_type not in _drivers():
        raise ValueError(f"不支持或未安装数据库驱动: {db_type}")
    return db, db_type


def db_connect(section, db_name=None, db_type=None, filename=None, clickhouse_type="connect"):
    db_config, db_type = read_db_config(section, filename, db_name, db_type)
    if db_type == "CLICKHOUSE":
        if clickhouse_type == "connect":
            if not db_config.get("dsn"):
                db_config["dsn"] = (
                    f"clickhouse://{db_config.get('host')}:{db_config.get('port')}/"
                    f"?user={db_config.get('user', '')}&password={db_config.get('password', '')}"
                )
            return clickhouse.connect(**db_config)
        return clickhouse.Client(**db_config)
    if db_type == "ORACLE":
        db_config["dsn"] = (
            f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={db_config.get('host')})"
            f"(PORT={db_config.get('port')}))(CONNECT_DATA=(SID={db_config.get('database')})))"
        )
        db_config.pop("host", None)
        db_config.pop("port", None)
        db_config.pop("database", None)
    return _drivers()[db_type](**db_config)


def add_tail(number, returns):
    for suffix in ["SZ", "SH", "OF", "BJ", "HK"]:
        wind_code = f"{number}.{suffix}"
        if wind_code in returns.columns:
            return wind_code
    return number


def quarter_cuts():
    dates = {
        "-12-31": "-01-24",
        "-03-31": "-04-26",
        "-06-30": "-07-22",
        "-09-30": "-10-29",
    }
    result = {}
    for year in range(2010, time.localtime().tm_year + 1):
        for end_date, ann_date in dates.items():
            ann_year = year + 1 if end_date == "-12-31" else year
            result[f"{year}{end_date}"] = f"{ann_year}{ann_date}"
    result = pd.Series(result).sort_index()
    result.index = pd.to_datetime(result.index)
    result = pd.to_datetime(result)
    return result.loc[result.index <= datetime.datetime.today()]


def get_fund_returns(start_date="20100101"):
    with db_connect("WIND_DB") as conn:
        a_price = pd.read_sql(
            f"""
            SELECT F_INFO_WINDCODE, PRICE_DATE, F_NAV_ADJUSTED
            FROM ChinaMutualFundNAV
            WHERE to_date(PRICE_DATE, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')
            """,
            conn,
        )
        hk_price = pd.read_sql(
            f"""
            SELECT F_INFO_WINDCODE, PRICE_DATE, F_NAV_ADJUSTED
            FROM CHFundNAV
            WHERE to_date(PRICE_DATE, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')
            """,
            conn,
        )
        price = pd.concat([a_price, hk_price])
        price = price.pivot(index="PRICE_DATE", columns="F_INFO_WINDCODE", values="F_NAV_ADJUSTED")
        price.index = pd.to_datetime(price.index)
        price.sort_index(inplace=True)
        returns = price.pct_change(fill_method="ffill").replace(0, np.nan)

        inner_data = pd.read_sql(
            f"""
            SELECT S_INFO_WINDCODE, TRADE_DT, S_DQ_PCTCHANGE, S_DQ_ADJCLOSE
            FROM ChinaClosedFundEODPrice
            WHERE to_date(TRADE_DT, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')
            """,
            conn,
        )
        inner_data["TRADE_DT"] = pd.to_datetime(inner_data["TRADE_DT"])
        inner_data["S_DQ_PCTCHANGE"] /= 100
        inner_price = inner_data.pivot(index="TRADE_DT", columns="S_INFO_WINDCODE", values="S_DQ_ADJCLOSE")
        inner_returns = inner_data.pivot(index="TRADE_DT", columns="S_INFO_WINDCODE", values="S_DQ_PCTCHANGE").replace(0, np.nan)

        fof_funds = pd.read_sql(
            """
            SELECT F_INFO_WINDCODE
            FROM ChinaMutualFundDescription
            WHERE F_INFO_FULLNAME LIKE '%FOF%'
            """,
            conn,
        )["F_INFO_WINDCODE"]

        common_columns = returns.columns.intersection(inner_returns.columns).drop(fof_funds, errors="ignore")
        returns[common_columns] = inner_returns[common_columns]
        returns = pd.concat([returns, inner_returns[inner_returns.columns.difference(returns.columns)]], axis=1)

        price[common_columns] = inner_price[common_columns]
        price = pd.concat([price, inner_price[inner_price.columns.difference(price.columns)]], axis=1)

    return returns.dropna(how="all"), price.dropna(how="all")


def get_fund_category():
    category = {}
    with db_connect("WIND_DB") as conn:
        for substr_len, level_num, category_name in [(8, "4", "WIND_1"), (10, "5", "WIND_2")]:
            df = pd.read_sql(
                f"""
                SELECT a.F_INFO_WINDCODE, b.INDUSTRIESNAME
                FROM ChinaMutualFundSector a
                INNER JOIN AShareIndustriesCode b
                ON SUBSTR(a.S_INFO_SECTOR, 1, {substr_len}) = SUBSTR(b.INDUSTRIESCODE, 1, {substr_len})
                AND a.cur_sign = '1'
                AND b.levelnum = '{level_num}'
                AND SUBSTR(a.S_INFO_SECTOR, 1, 6) = '200101'
                ORDER BY 1
                """,
                conn,
            )
            category[category_name] = df.set_index("F_INFO_WINDCODE")["INDUSTRIESNAME"].to_dict()
        hk = pd.read_sql(
            "SELECT F_INFO_WINDCODE, F_INFO_FIRSTINVESTTYPE FROM CHFundDescription",
            conn,
        )
    hk.set_index("F_INFO_WINDCODE", inplace=True)
    hk["WIND_1"] = "中港互认基金"
    hk.rename(columns={"F_INFO_FIRSTINVESTTYPE": "WIND_2"}, inplace=True)
    hk["WIND_2"] = "中港互认" + hk["WIND_2"]
    return pd.concat([pd.DataFrame(category), hk])


def get_fund_describe():
    conn = db_connect("WIND_DB")
    describe = pd.read_sql(
        "SELECT F_INFO_WINDCODE, F_INFO_NAME, F_INFO_CORP_FUNDMANAGEMENTCOMP, F_INFO_SETUPDATE, F_INFO_ISINITIAL FROM ChinaMutualFundDescription",
        conn,
    )
    hk_describe = pd.read_sql(
        "SELECT F_INFO_WINDCODE, F_INFO_NAME, F_INFO_CORP_FUNDMANAGEMENTCOMP, F_INFO_SETUPDATE, F_INFO_ISINITIAL FROM CHFundDescription",
        conn,
    )
    describe = pd.concat([describe, hk_describe], ignore_index=True)
    describe.set_index("F_INFO_WINDCODE", inplace=True)
    describe["F_INFO_SETUPDATE"] = pd.to_datetime(describe["F_INFO_SETUPDATE"])
    describe["F_INFO_ISINITIAL"] = describe["F_INFO_ISINITIAL"] == 1
    describe.index.name = "基金代码"
    describe = describe.rename(
        columns={
            "F_INFO_NAME": "基金名称",
            "F_INFO_CORP_FUNDMANAGEMENTCOMP": "基金管理人",
            "F_INFO_SETUPDATE": "基金成立日期",
            "F_INFO_ISINITIAL": "是否主份额",
        }
    ).replace("nan", np.nan)
    describe["万得二级分类"] = get_fund_category()["WIND_2"]
    describe.loc[describe["基金名称"].str.contains("REIT", na=False), "万得二级分类"] = (
        describe.loc[describe["基金名称"].str.contains("REIT", na=False), "万得二级分类"].fillna("REITs")
    )
    return describe


def fund_position(start_date="20230101"):
    conn = db_connect("WIND_DB")
    position = pd.read_sql(
        f"""
        SELECT S_INFO_WINDCODE, F_PRT_ENDDATE, S_INFO_STOCKWINDCODE, F_PRT_STKVALUE, F_PRT_STKVALUETONAV, ANN_DATE, STOCK_PER
        FROM ChinaMutualFundStockPortfolio
        WHERE to_date(ANN_DATE, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')
        """,
        conn,
    )
    position["ANN_DATE"] = pd.to_datetime(position["ANN_DATE"])
    position["F_PRT_ENDDATE"] = pd.to_datetime(position["F_PRT_ENDDATE"])
    position.sort_values("ANN_DATE", inplace=True)
    return position.drop_duplicates()


def _normalize_listed_fund_codes(df):
    mapper = df.loc[
        df["S_INFO_WINDCODE"].str.endswith("SZ") | df["S_INFO_WINDCODE"].str.endswith("SH"),
        "S_INFO_WINDCODE",
    ].unique()
    mapper = {i[:-2] + "OF": i for i in mapper}
    df = df.copy()
    df["S_INFO_WINDCODE"] = df["S_INFO_WINDCODE"].apply(lambda x: x[:-2] + "OF")
    return df, mapper


def quarter_report(position, start_date="2023-01-01"):
    conn = db_connect("WIND_DB")
    cache, mapper = _normalize_listed_fund_codes(position)
    ann = pd.read_sql(
        f"""
        SELECT S_INFO_WINDCODE, ANN_DT, N_INFO_TITLE, N_INFO_FCODE, COLLECT_TIME
        FROM FundAnnInf
        WHERE N_INFO_FCODE IN ('5602030000')
        AND ANN_DT >= to_date('{start_date.replace("-", "")}', 'yyyyMMdd')
        """,
        conn,
    )
    ann["ANN_DT"] = pd.to_datetime(ann["ANN_DT"])
    ann = ann[ann["ANN_DT"] >= start_date].sort_values("ANN_DT")
    ann = ann[(ann["N_INFO_TITLE"].str.contains("季度")) | (ann["N_INFO_TITLE"].str.contains("季报"))]
    for keyword in ["更正", "提示", "补充", "组合情况", "更新"]:
        ann = ann[~ann["N_INFO_TITLE"].str.contains(keyword)]
    result = pd.merge(
        cache,
        ann,
        left_on=["S_INFO_WINDCODE", "ANN_DATE"],
        right_on=["S_INFO_WINDCODE", "ANN_DT"],
    )[cache.columns].drop_duplicates()
    result["S_INFO_WINDCODE"] = result["S_INFO_WINDCODE"].apply(lambda x: mapper[x] if x in mapper else x)
    result["F_PRT_STKVALUE"].fillna(0, inplace=True)
    result.sort_values("F_PRT_STKVALUE", inplace=True)
    result = result.groupby(["S_INFO_WINDCODE", "F_PRT_ENDDATE"]).tail(10)
    ann_cut = (
        [f"{i}-12-31" for i in range(1995, 2030)]
        + [f"{i}-06-30" for i in range(1995, 2030)]
        + [f"{i}-03-31" for i in range(1995, 2030)]
        + [f"{i}-09-30" for i in range(1995, 2030)]
    )
    return result[result["F_PRT_ENDDATE"].isin(ann_cut)].sort_values("ANN_DATE")


def detail_report(position, start_date="2023-01-01"):
    conn = db_connect("WIND_DB")
    quarters = quarter_report(position, start_date)
    cache, mappers = _normalize_listed_fund_codes(position)
    ann = pd.read_sql(
        f"""
        SELECT S_INFO_WINDCODE, ANN_DT, N_INFO_TITLE, N_INFO_FCODE, COLLECT_TIME
        FROM FundAnnInf
        WHERE N_INFO_FCODE IN ('5602010000', '5602020000')
        AND ANN_DT >= to_date('{start_date.replace("-", "")}', 'yyyyMMdd')
        """,
        conn,
    )
    ann["ANN_DT"] = pd.to_datetime(ann["ANN_DT"])
    ann = ann[ann["ANN_DT"] >= start_date].sort_values("ANN_DT")
    ann = ann[
        (ann["N_INFO_TITLE"].str.contains("年度"))
        | (ann["N_INFO_TITLE"].str.contains("半年"))
        | (ann["N_INFO_TITLE"].str.contains("年报"))
        | (ann["N_INFO_TITLE"].str.contains("中期报告"))
    ]
    for keyword in ["摘要", "更正", "提示", "补充", "公告", "审计", "展望"]:
        ann = ann[~ann["N_INFO_TITLE"].str.contains(keyword)]
    half_year = pd.merge(
        cache,
        ann,
        left_on=["S_INFO_WINDCODE", "ANN_DATE"],
        right_on=["S_INFO_WINDCODE", "ANN_DT"],
    )[cache.columns].drop_duplicates()
    half_year["S_INFO_WINDCODE"] = half_year["S_INFO_WINDCODE"].apply(lambda x: mappers[x] if x in mappers else x)
    mapper = half_year.groupby(["S_INFO_WINDCODE", "F_PRT_ENDDATE"])["ANN_DATE"].unique()
    quarters["ANN_DATE"] = quarters.apply(
        lambda x: (
            mapper.loc[x["S_INFO_WINDCODE"], x["F_PRT_ENDDATE"]][0]
            if (x["S_INFO_WINDCODE"], x["F_PRT_ENDDATE"]) in mapper.index
            else np.nan
        ),
        axis=1,
    )
    half_year = pd.concat([half_year, quarters.dropna()], ignore_index=True)
    ann_cut = [f"{i}-12-31" for i in range(1995, 2030)] + [f"{i}-06-30" for i in range(1995, 2030)]
    return half_year[half_year["F_PRT_ENDDATE"].isin(ann_cut)].sort_values("ANN_DATE")


def get_allocation(start_date="20230101"):
    conn = db_connect("WIND_DB")
    allocation = pd.read_sql(
        f"""
        SELECT S_INFO_WINDCODE, F_PRT_ENDDATE, F_PRT_STOCKTONAV, F_PRT_CASHTONAV, F_PRT_COVERTBONDTONAV, F_PRT_BONDTONAV, F_PRT_FUNDTONAV, F_PRT_OTHERTONAV
        FROM ChinaMutualFundAssetPortfolio
        WHERE to_date(F_PRT_ENDDATE, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')
        """,
        conn,
    )
    allocation["F_PRT_ENDDATE"] = pd.to_datetime(allocation["F_PRT_ENDDATE"])
    allocation = allocation.loc[allocation["F_PRT_ENDDATE"].isin(quarter_cuts().index)]
    allocation.fillna(0, inplace=True)
    allocation = allocation.set_index(["S_INFO_WINDCODE", "F_PRT_ENDDATE"]).sort_index()
    allocation.columns = ["stock", "cash", "convert", "bond", "fund", "other"]
    allocation["bond"] -= allocation["convert"]
    return allocation / 100


def get_benches():
    conn = db_connect("WIND_DB")
    benches = pd.read_sql(
        "SELECT S_INFO_WINDCODE, S_INFO_INDEXWINDCODE, REMOVE_DT FROM ChinaMutualFundTrackingIndex",
        conn,
    )
    benches = benches[pd.isna(benches["REMOVE_DT"])].set_index("S_INFO_WINDCODE")["S_INFO_INDEXWINDCODE"]
    return benches[~benches.index.str.endswith("XT")]


def add_important_index(index_returns, start_date="20230101"):
    interested = {
        3210: "SPX.GI",
        3204: "IXIC.GI",
        3205: "NDX.GI",
        3173: "FCHI.GI",
        3176: "GDAXI.GI",
        3162: "N225.GI",
        272750: "VN30.GI",
        324687: "WISAUNT.FI",
        3199: "IBOVESPA.GI",
        204262: "935600.MI",
        3171: "FTSE.GI",
        140392: "DWRTF.GI",
        564844: "INSYBUE.GI",
        643693: "SXXP.GI",
        21054: "990100.MI",
        14078: "SPXEWTR.SPI",
        3206: "DJI.GI",
        491518: "NDXTMC.GI",
        12776: "SPOEXEUP.SPI",
        227230: "TPX.GI",
        6181: "891800.MI",
    }
    conn = db_connect("JY_DB")
    df = pd.read_sql(
        "SELECT IndexCode, TradingDay, ChangePCT FROM QT_OSIndexQuote a WHERE IndexCode IN ("
        + str(list(interested.keys()))[1:-1]
        + f") AND TradingDay >= '{start_date}'",
        conn,
    )
    world_index_return = df.pivot(index="TradingDay", columns="IndexCode", values="ChangePCT").rename(columns=interested) / 100
    world_index_return.index = pd.to_datetime(world_index_return.index)
    index_returns = index_returns.copy()
    for col in world_index_return.columns:
        index_returns[col] = world_index_return[col]
    return index_returns


def add_world_index(index_returns, start_date="20230101"):
    if clickhouse_connect is None:
        raise ImportError("需要安装 clickhouse_driver 才能读取彭博指数")
    cur = clickhouse_connect("clickhouse://ckreader:ckreader123456@10.3.64.7:9000/ods_ths")
    world_index = pd.read_sql(
        f"SELECT * FROM ods_ths.all_index WHERE TIME >= '{pd.to_datetime(start_date).strftime('%Y-%m-%d')}'",
        cur,
    )
    world_index = world_index.groupby(["TIME", "INDEX_NAME"])["PX_LAST"].last().unstack()
    world_index = world_index.replace("", np.nan).astype(float).clip(lower=0.0)
    world_index_returns = world_index.pct_change(fill_method=None).replace(0, np.nan)
    world_index_returns = world_index_returns[
        [
            "XIN9I INDEX",
            "SPOEXEUP INDEX",
            "SPXEWTR INDEX",
            "LUACTRUU Index",
            "LF98TRUU Index",
            "SPBDU1BT INDEX",
            "SPBDUS3T INDEX",
            "GOLDS Comdty",
            "HG1 Comdty",
            "EREE Index",
            "ERGLU INDEX",
            "CO1 Comdty",
            "CL1 Comdty",
        ]
    ]
    world_index_returns.index = pd.to_datetime(world_index_returns.index)
    # world_index_returns.loc[["2020-04-20", "2020-04-21"], "CL1 Comdty"] = np.nan
    return pd.merge(index_returns, world_index_returns, left_index=True, right_index=True, how="outer")


def get_multiple_index_return(index_codes=None, start_date="20230101"):
    conn = db_connect("WIND_DB")
    start_ts = pd.to_datetime(start_date)

    def format_codes(codes):
        return ",".join(f"'{c}'" for c in codes)

    def read_pctchange_sql(table, pct_col="S_DQ_PCTCHANGE", extra_where="", codes=None):
        sql = f"SELECT S_INFO_WINDCODE, TRADE_DT, {pct_col} FROM {table}"
        where = [f"to_date(TRADE_DT, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')"]
        if codes:
            where.append(f"S_INFO_WINDCODE IN ({format_codes(codes)})")
        if extra_where:
            where.append(extra_where)
        sql += " WHERE " + " AND ".join(where)
        df = pd.read_sql(sql, conn)
        df["TRADE_DT"] = pd.to_datetime(df["TRADE_DT"])
        return df

    def read_close_prec_sql(table, codes=None):
        sql = f"SELECT S_INFO_WINDCODE, TRADE_DT, S_DQ_PRECLOSE, S_DQ_CLOSE FROM {table}"
        where = [f"to_date(TRADE_DT, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')"]
        if codes:
            where.append(f"S_INFO_WINDCODE IN ({format_codes(codes)})")
        sql += " WHERE " + " AND ".join(where)
        df = pd.read_sql(sql, conn)
        df["S_DQ_PCTCHANGE"] = df["S_DQ_CLOSE"] / df["S_DQ_PRECLOSE"] - 1
        df["TRADE_DT"] = pd.to_datetime(df["TRADE_DT"])
        return df

    def read_close_sql(table, codes=None, close_col="S_DQ_INDEXVALUE"):
        sql = f"SELECT S_INFO_WINDCODE, TRADE_DT, {close_col} FROM {table}"
        where = [f"to_date(TRADE_DT, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')"]
        if codes:
            where.append(f"S_INFO_WINDCODE IN ({format_codes(codes)})")
        sql += " WHERE " + " AND ".join(where)
        df = pd.read_sql(sql, conn)
        df["TRADE_DT"] = pd.to_datetime(df["TRADE_DT"])
        df = pd.pivot(df, index="TRADE_DT", columns="S_INFO_WINDCODE", values=close_col).sort_index()
        return df.pct_change(fill_method=None)

    def pivot_pctchange(df):
        return pd.pivot(df, index="TRADE_DT", columns="S_INFO_WINDCODE", values="S_DQ_PCTCHANGE").sort_index()

    stock_index_return = pivot_pctchange(read_pctchange_sql("AIndexEODPrices", codes=index_codes)) / 100
    sw_index_return = pivot_pctchange(read_close_prec_sql("ASWSIndexEOD", codes=index_codes))
    citic_index_return = pivot_pctchange(read_close_prec_sql("AIndexIndustriesEODCITICS", codes=index_codes))
    wind_index_return = pivot_pctchange(read_close_prec_sql("AIndexWindIndustriesEOD", codes=index_codes))
    zx_index_return = pivot_pctchange(read_close_prec_sql("ASPCITICIndexEOD", codes=index_codes))
    hk_index_return = pivot_pctchange(read_pctchange_sql("HKIndexEODPrices", codes=index_codes)) / 100
    bond_index_return = read_pctchange_sql("CBIndexEODPrices", codes=index_codes)
    bond_index_return = bond_index_return.groupby(["S_INFO_WINDCODE", "TRADE_DT"])["S_DQ_PCTCHANGE"].last().unstack().T / 100

    wind_bond_index_return = pd.read_sql(
        "SELECT S_INFO_WINDCODE, TRADE_DT, S_DQ_CLOSE FROM CBIndexEODPricesWIND"
        + " WHERE to_date(TRADE_DT, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')".format(start_date=start_date)
        + (f" AND S_INFO_WINDCODE IN ({format_codes(index_codes)})" if index_codes else ""),
        conn,
    )
    wind_bond_index_return["TRADE_DT"] = pd.to_datetime(wind_bond_index_return["TRADE_DT"])
    wind_bond_index_return["S_DQ_PCTCHANGE"] = wind_bond_index_return.groupby("S_INFO_WINDCODE")["S_DQ_CLOSE"].pct_change(fill_method=None).fillna(0)
    wind_bond_index_return = pivot_pctchange(wind_bond_index_return)

    fund_index_return = pivot_pctchange(read_pctchange_sql("CMFIndexEOD", codes=index_codes)) / 100
    csi_index_return = pivot_pctchange(read_close_prec_sql("CSIAIndexEODPrices", codes=index_codes)) / 100
    csi_index_return2 = read_close_sql("CSITotalBondIndeEODPrice", codes=index_codes)
    cbond_index_return = read_close_sql("CBondIndexEODCNBD", codes=index_codes, close_col="S_DQ_CLOSE")
    neeq_index_return = pivot_pctchange(read_close_prec_sql("NEEQIndexEODPrices", codes=index_codes))

    msci_sql = (
        "SELECT S_INFO_WINDCODE, TRADE_DT, S_DQ_CLOSE_STD_ FROM AMSCIIndexEOD WHERE S_DQ_ISLOCALCRNCY = 0"
        + f" AND to_date(TRADE_DT, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')"
        + (f" AND S_INFO_WINDCODE IN ({format_codes(index_codes)})" if index_codes else "")
    )
    msci_index_return = pd.read_sql(msci_sql, conn)
    msci_index_return["TRADE_DT"] = pd.to_datetime(msci_index_return["TRADE_DT"])
    msci_index_return = msci_index_return.groupby(["S_INFO_WINDCODE", "TRADE_DT"])["S_DQ_CLOSE_STD_"].last().unstack().T
    msci_index_return = msci_index_return.pct_change(fill_method=None).replace(0, np.nan).dropna(how="all").dropna(how="all", axis=1)

    future_index_return = pd.read_sql(
        "SELECT S_INFO_WINDCODE, TRADE_DT, S_DQ_CLOSE FROM ThirdPartyIndexEOD"
        + f" WHERE to_date(TRADE_DT, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')"
        + (f" AND S_INFO_WINDCODE IN ({format_codes(index_codes)})" if index_codes else ""),
        conn,
    )
    future_index_return["TRADE_DT"] = pd.to_datetime(future_index_return["TRADE_DT"])
    future_index_return = future_index_return.groupby(["S_INFO_WINDCODE", "TRADE_DT"])["S_DQ_CLOSE"].last().unstack().T.replace(0, np.nan)
    future_index_return = future_index_return.pct_change(fill_method=None).replace(0, np.nan).dropna(how="all").dropna(how="all", axis=1)

    commody_return = pd.read_sql(
        "SELECT S_INFO_WINDCODE, TRADE_DT, S_DQ_PRESETTLE, S_DQ_SETTLE FROM CCommodityFuturesEODPrices"
        + f" WHERE to_date(TRADE_DT, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')"
        + (f" AND S_INFO_WINDCODE IN ({format_codes(index_codes)})" if index_codes else ""),
        conn,
    )
    commody_return["S_DQ_PCTCHANGE"] = commody_return["S_DQ_SETTLE"] / commody_return["S_DQ_PRESETTLE"] - 1
    commody_return["TRADE_DT"] = pd.to_datetime(commody_return["TRADE_DT"])
    commody_return = pd.pivot(commody_return, index="TRADE_DT", columns="S_INFO_WINDCODE", values="S_DQ_PCTCHANGE").sort_index()
    commody_return = commody_return[commody_return.columns[~commody_return.columns.str.contains("0|1|2|3|4|5|6|7|8|9")]]

    commody_index_return = pd.read_sql(
        "SELECT S_INFO_WINDCODE, TRADE_DT, S_DQ_PRECLOSE, S_DQ_CLOSE FROM CFutureIndexEODPrices"
        + f" WHERE to_date(TRADE_DT, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')"
        + (f" AND S_INFO_WINDCODE IN ({format_codes(index_codes)})" if index_codes else ""),
        conn,
    )
    commody_index_return["S_DQ_PCTCHANGE"] = commody_index_return["S_DQ_CLOSE"] / commody_index_return["S_DQ_PRECLOSE"] - 1
    commody_index_return["TRADE_DT"] = pd.to_datetime(commody_index_return["TRADE_DT"])
    commody_index_return = pd.pivot(commody_index_return, index="TRADE_DT", columns="S_INFO_WINDCODE", values="S_DQ_PCTCHANGE").sort_index()

    if index_codes and "Au9999.SGE" not in index_codes:
        au = pd.DataFrame(columns=["Au9999.SGE"])
    else:
        au = pd.read_sql(
            f"""
            SELECT TRADE_DT, S_PCT_CHG
            FROM CGoldSpotEODPrices
            WHERE S_INFO_WINDCODE = 'Au9999.SGE'
            AND to_date(TRADE_DT, 'yyyyMMdd') >= to_date('{start_date}', 'yyyyMMdd')
            """,
            conn,
        )
        au["TRADE_DT"] = pd.to_datetime(au["TRADE_DT"])
        au = (au.set_index("TRADE_DT")["S_PCT_CHG"] / 100).rename("Au9999.SGE").to_frame().sort_index()

    cj_index_return = pivot_pctchange(read_close_prec_sql("AIndexIndustriesEODCJZQ", codes=index_codes))

    index_return = pd.concat(
        [
            stock_index_return,
            sw_index_return,
            citic_index_return,
            wind_index_return,
            zx_index_return,
            hk_index_return,
            bond_index_return,
            wind_bond_index_return,
            fund_index_return,
            csi_index_return,
            csi_index_return2,
            cbond_index_return,
            neeq_index_return,
            msci_index_return,
            future_index_return,
            commody_index_return,
            commody_return,
            au,
            cj_index_return,
        ],
        axis=1,
    )
    index_return = index_return.loc[:, ~index_return.columns.duplicated()].sort_index()
    if not index_codes:
        index_return = add_important_index(index_return, start_date)
        index_return = add_world_index(index_return, start_date)
    return index_return.loc[index_return.index >= start_ts]


def get_index_category():
    conn = db_connect("WIND_DB")
    queries = {
        "A股": [
            "SELECT DISTINCT S_INFO_WINDCODE FROM AIndexEODPrices",
            "SELECT DISTINCT S_INFO_WINDCODE FROM ASWSIndexEOD",
            "SELECT DISTINCT S_INFO_WINDCODE FROM AIndexIndustriesEODCITICS",
            "SELECT DISTINCT S_INFO_WINDCODE FROM AIndexWindIndustriesEOD",
            "SELECT DISTINCT S_INFO_WINDCODE FROM ASPCITICIndexEOD",
            "SELECT DISTINCT S_INFO_WINDCODE FROM CSIAIndexEODPrices",
            "SELECT DISTINCT S_INFO_WINDCODE FROM AMSCIIndexEOD",
        ],
        "港股": ["SELECT DISTINCT S_INFO_WINDCODE FROM HKIndexEODPrices"],
        "A债": [
            "SELECT DISTINCT S_INFO_WINDCODE FROM CBIndexEODPrices",
            "SELECT DISTINCT S_INFO_WINDCODE FROM CBIndexEODPricesWIND",
            "SELECT DISTINCT S_INFO_WINDCODE FROM CSITotalBondIndeEODPrice",
            "SELECT DISTINCT S_INFO_WINDCODE FROM CBondIndexEODCNBD",
        ],
        "基金": ["SELECT DISTINCT S_INFO_WINDCODE FROM CMFIndexEOD"],
    }

    def fetch_indices(query_list):
        result = []
        for query in query_list:
            result.extend(pd.read_sql(query, conn)["S_INFO_WINDCODE"].dropna().tolist())
        return result

    stock_index = fetch_indices(queries["A股"])
    hk_index = fetch_indices(queries["港股"])
    bond_index = fetch_indices(queries["A债"])
    fund_index = fetch_indices(queries["基金"])
    missing_bench = pd.read_sql(
        "SELECT DISTINCT S_INFO_INDEXWINDCODE FROM ChinaMutualFundBenchMark WHERE CUR_SIGN = 1 AND S_INFO_INDEXWINDCODE IS NOT NULL",
        conn,
    )["S_INFO_INDEXWINDCODE"].dropna()
    missing_bench = missing_bench[~missing_bench.isin(stock_index + hk_index + bond_index + fund_index)]

    categories = {
        "A股": stock_index + ["XIN9I INDEX", "SPCQVCP.SPI", "GPCCH003.FI", "DJCN88.GI", "SPCADMCP.SPI"],
        "港股": hk_index + ["SPACEVCP.SPI", "FCAH50.FI", "GPCSP006.FI", "SPAHLVHP.SPI", "HSHDYI.HI"],
        "A债": bond_index + ["SPBCNCOTP.SPI", "830101.XI", "SPCCBI.SPI", "830102.XI", "CBM00701.CS", "CI0251106.WI"],
        "基金": fund_index + ["801615.SI", "801639.SI", "930868.CSI", "930842.CSI"],
        "货币": missing_bench[missing_bench.str.contains("DEPO", na=False)].tolist()
        + ["DEMAND.WI", "CALLDEPO7D.WI", "DEPO1Y.WI", "DEPO6M.WI", "DEPO3M.WI", "HIBOR3M.IR", "SHIBOR3M.IR", "h11025.CSI"],
        "可转债": ["000832.CSI", "931078.CSI"],
        "美股": ["NDX.GI", "SPX.GI", "XNDX.O", "SPTR500N.SPI", "NDXTMC.GI", "SPXEWTR.SPI", "S5INFT.SPI", "DJINET.GI", "SOX.GI", "DJI.GI", "SPOEXEUP.SPI", "SPOEXEUP INDEX", "SPXEWTR INDEX", "SPX Index", "891800.MI", "990100.MI", "SPN.SPI"],
        "海外债券": missing_bench[missing_bench.str.contains("YRNOTE", na=False)].tolist()
        + ["INSYBUE.GI", "IOVROV.GI", "UST10Y.GBM", "SBWGL.CIT", "LUACTRUU Index", "LF98TRUU Index", "SPBDU1BT INDEX", "SPBDUS3T INDEX"],
        "黄金": ["Au9999.SGE", "SHAU.SGE", "SPTAUUSDOZ.IDC", "GC1 Comdty", "GOLDS Comdty"],
        "能化": ["000201.CZC"],
        "白银": ["AG.SHF"],
        "豆粕": ["DCESMFI.DCE"],
        "有色金属": ["IMCI.SHF", "HG1 Comdty"],
        "英国": ["FTSE.GI", "UKX INDEX"],
        "境外REITs": ["128456.MI", "ERGLU.FI", "DWRTF.GI", "TERGLU.FI", "EREE Index", "REI INDEX", "RMZ INDEX", "ERGLU INDEX"],
        "境内REITs": ["932006.CSI"],
        "日本": ["N225.GI", "TPX.GI"],
        "印度": ["935600.MI", "CIS51001.WI", "NIFTY Index"],
        "德国": ["GDAXI.GI"],
        "法国": ["FCHI.GI"],
        "越南": ["VHINDEX.GI", "VNINDEX Index", "VN30.GI"],
        "原油": ["CL.NYM", "SPGSBRTR.SPI", "T.IPE", "B.IPE", "CL1 Comdty", "CO1 Comdty", "SPSIOP.SPI", "SPGOGUP.SPI", "SPGSCITR.SPI"],
        "东南亚": ["ASIATECP.SG", "EMASIAUP.SG"],
        "商品": ["h30009.CSI", "NH0100.NHF", "CCFI.WI", "000001.CCI"],
        "海外权益": ["990100.MI"],
        "巴西": ["IBOVESPA.GI"],
        "沙特": ["WISAUNT.FI"],
        "股票多空": ["885078.WI"],
    }
    result = pd.concat([pd.Series(category, index=indices) for category, indices in categories.items()])
    return result.groupby(result.index).last()


def get_bench_universe(index_returns, date=datetime.datetime.today().strftime("%Y-%m-%d")):
    index_returns = index_returns.loc[:date].dropna(how="all", axis=1)
    start = index_returns.apply(lambda x: x.dropna().index[0])
    end = index_returns.apply(lambda x: x.dropna().index[-1])
    latest = pd.to_datetime(date) - pd.DateOffset(years=2)
    earliest = pd.to_datetime(date) - pd.DateOffset(months=1)
    bench_universe = start[start <= latest].index.intersection(end[end >= earliest].index)
    category = get_index_category()
    bench_universe = bench_universe.intersection(category[category != "基金"].index)

    bond_index = bench_universe.intersection(category[category == "A债"].index)
    conn = db_connect("WIND_DB")
    existing_bench = pd.read_sql(
        "SELECT DISTINCT S_INFO_INDEXWINDCODE FROM ChinaMutualFundBenchMark WHERE CUR_SIGN = 1 AND S_INFO_INDEXWINDCODE IS NOT NULL",
        conn,
    )["S_INFO_INDEXWINDCODE"]
    bench_universe = pd.Index([i for i in bench_universe if i not in [j for j in bond_index if j not in existing_bench.values]])

    cur = db_connect("SAM_CH1")
    with_weight = pd.read_sql("SELECT DISTINCT index_code FROM factor.index_port", cur)["index_code"]
    equity_index = category[category.isin(["A股", "港股"])].index.intersection(bench_universe)
    missing_equity = [i for i in equity_index if i not in with_weight.values]
    return pd.Index([i for i in bench_universe if i not in missing_equity])


def calc_fund_basic(fund_basic, returns, min_len=253):
    fund_basic = fund_basic.loc[fund_basic.index.intersection(returns.columns)]
    passive = get_benches()
    length = returns[fund_basic.index].apply(lambda x: x.count()).sort_values()
    fund_basic = fund_basic.loc[length[length >= min_len].index]
    return fund_basic.loc[[i for i in fund_basic.index if i not in passive.index]]


def calc_tracking_error(fund_basic_cal, returns, bench_universe, index_returns):
    if fund_basic_cal.empty or returns.empty or index_returns.empty or bench_universe.empty:
        return pd.DataFrame()
    fund_returns = returns.loc[:, fund_basic_cal.index]
    valid_benchmarks = bench_universe.intersection(index_returns.columns)
    common_dates = fund_returns.index.intersection(index_returns.index)
    if len(valid_benchmarks) == 0 or len(common_dates) == 0:
        return pd.DataFrame()
    aligned_returns = fund_returns.loc[common_dates]
    result_data = {}
    for benchmark in tqdm(valid_benchmarks, desc="串行计算跟踪误差"):
        benchmark_returns = index_returns.loc[common_dates, benchmark]
        excess_returns = aligned_returns.sub(benchmark_returns, axis=0)
        result_data[benchmark] = excess_returns.std() * np.sqrt(252)
        del excess_returns, benchmark_returns
        gc.collect()
    result_df = pd.DataFrame(result_data)
    result_df.index.name = "F_INFO_WINDCODE"
    return result_df.dropna(how="all", axis=1).astype(float)


def calc_fund_bench(te, index_category, fund_basic):
    benches = te.idxmin(axis=1)
    industry = fund_basic["万得二级分类"].copy()

    equity_funds = industry[industry.isin(["偏股混合型基金", "普通股票型基金"])].index
    bond_funds = industry[
        industry.isin(["偏债混合型基金", "混合债券型二级基金", "短期纯债型基金", "中长期纯债型基金", "混合债券型一级基金"])
    ].index
    equity_index = index_category[index_category.isin(["A股", "港股"])].index.intersection(te.columns)
    bond_index = index_category[index_category == "A债"].index
    if len(equity_index) > 0:
        equity_benches = te.loc[te.index.intersection(equity_funds), equity_index].idxmin(axis=1)
        benches.loc[equity_benches.index] = equity_benches
    if len(bond_index.intersection(te.columns)) > 0:
        bond_benches = te.loc[te.index.intersection(bond_funds), te.columns.intersection(bond_index)].idxmin(axis=1)
        benches.loc[bond_benches.index] = bond_benches

    bench_table = benches.reset_index().set_index(0)
    bench_table["资产各类"] = index_category
    other = bench_table[
        (~bench_table["F_INFO_WINDCODE"].isin(equity_funds))
        & (~bench_table["F_INFO_WINDCODE"].isin(bond_funds))
    ]
    other_equity = other[other["资产各类"].isin(["A股", "港股"])]
    if not other_equity.empty and len(equity_index) > 0:
        other_equity_benches = te.loc[other_equity["F_INFO_WINDCODE"], equity_index].idxmin(axis=1)
        benches.loc[other_equity_benches.index] = other_equity_benches

    string_result = pd.Series(index=fund_basic.index, dtype=object)
    fund_names = fund_basic["基金名称"].fillna("")
    string_result.loc[fund_basic[fund_names.str.contains("货币|现金") & (~fund_names.str.contains("现金流"))].index] = "h11025.CSI"
    string_result.loc[fund_basic[fund_names.str.contains("转债|转换")].index] = "931078.CSI"
    string_result.loc[fund_basic[fund_names.str.contains("REIT")].index] = "932006.CSI"
    string_result.loc[fund_basic[industry == "股票多空"].index] = "885078.WI"
    string_result.dropna(inplace=True)
    benches = benches.reindex(benches.index.union(string_result.index))
    benches.loc[string_result.index] = string_result

    qdii_bond_funds = industry[industry.isin(["国际(QDII)债券型基金", "中港互认债券型基金"])].index
    qdii_bond_index = index_category[index_category == "海外债券"].index
    qdii_bond_te = te.loc[te.index.intersection(qdii_bond_funds), te.columns.intersection(qdii_bond_index)]
    if not qdii_bond_te.empty:
        qdii_bond_benches = qdii_bond_te.idxmin(axis=1)
        benches = benches.reindex(benches.index.union(qdii_bond_benches.index))
        benches.loc[qdii_bond_benches.index] = qdii_bond_benches

    passive = get_benches()
    benches = benches.reindex(benches.index.union(passive.index))
    benches.loc[passive.index] = passive

    missing2 = [i for i in fund_basic.index if i not in benches.index]
    conn = db_connect("WIND_DB")
    missing_bench = pd.read_sql(
        "SELECT S_INFO_WINDCODE, S_INFO_INDEXWINDCODE, S_INFO_INDEXWEG FROM ChinaMutualFundBenchMark WHERE CUR_SIGN = 1 AND S_INFO_INDEXWINDCODE IS NOT NULL",
        conn,
    )
    missing_bench = missing_bench[missing_bench["S_INFO_WINDCODE"].isin(missing2)]
    if not missing_bench.empty:
        missing_bench = pd.pivot(missing_bench, index="S_INFO_WINDCODE", columns="S_INFO_INDEXWINDCODE", values="S_INFO_INDEXWEG")
        for_append = missing_bench.idxmax(axis=1)
        cache = index_category.loc[index_category.index.intersection(for_append.unique())]
        cache = cache[cache.isin(["A股", "港股"])]
        for_append.replace({i: "000300.SH" for i in cache.index if i not in te.columns}, inplace=True)
        benches = pd.concat([benches, for_append])

    benches.loc["022512.OF"] = "SPBDUS3T INDEX"
    benches.loc["968130.OF"] = "SPBDUS3T INDEX"
    benches.loc["968163.OF"] = "SPBDUS3T INDEX"
    benches.loc["968153.OF"] = "SPBDUS3T INDEX"
    benches.loc["3110.HK"] = "HSHDYI.HI"
    benches.loc["501018.SH"] = "CL1 Comdty"
    benches.loc["511880.SH"] = "h11025.CSI"
    benches.loc["968168.OF"] = "SPX.GI"
    benches.loc["968121.OF"] = "INSYBUE.GI"
    benches.loc["968155.OF"] = "INSYBUE.GI"
    benches.loc["968157.OF"] = "SPX.GI"
    benches.loc["2800.HK"] = "HSI.HI"
    return benches.groupby(benches.index).last()


def build_fund_benchmark_result(index_return, returns, fund_basic, index_category):
    bench_universe = get_bench_universe(index_return)
    fund_basic_cal = calc_fund_basic(fund_basic, returns)
    te = calc_tracking_error(fund_basic_cal, returns, bench_universe, index_return)
    result = calc_fund_bench(te, index_category, fund_basic)
    result.loc["third_foreword"] = "000300.SH"
    result.loc["third_backword"] = "000300.SH"
    return result


def build_fund_mapper(result, index_category, fund_basic, passive_returns):
    fund_category = result.replace(index_category)
    passive = passive_returns.replace(index_category)
    mapper = pd.Series(index=fund_category.index, dtype=object)
    mapper.loc[fund_category[fund_category == "黄金"].index.intersection(mapper.index)] = "黄金ETF"
    mapper.loc[fund_category[fund_category.isin(["白银", "有色金属", "原油", "豆粕", "能化"])].index.intersection(mapper.index)] = "其他商品"
    mapper.loc[fund_category[fund_category.isin(["境外REITs", "境内REITs"])].index.intersection(mapper.index)] = "REITS"
    mapper.loc[fund_category[fund_category.isin(["货币"])].index.intersection(mapper.index)] = "货币基金"
    mapper.loc[fund_category[fund_category.isin(["海外债券"])].index.intersection(mapper.index)] = "境外固收"
    mapper.loc[fund_basic[fund_basic["万得二级分类"].isin(["中长期纯债型基金", "短期纯债型基金"])].index.intersection(mapper.index)] = "纯债债基"
    mapper.loc[fund_basic[fund_basic["万得二级分类"].isin(["混合债券型二级基金", "混合债券型一级基金"])].index.intersection(mapper.index)] = "二级债基"
    mapper.loc[fund_basic[fund_basic["万得二级分类"].isin(["偏债混合型基金", "可转换债券型基金"])].index.intersection(mapper.index)] = "偏债混合"
    mapper.loc[fund_basic[fund_basic["万得二级分类"].isin(["被动指数型债券基金", "增强指数型债券基金"])].index.intersection(mapper.index)] = "债券ETF"
    mapper.loc[fund_category[fund_category.isin(["股票多空"])].index.intersection(mapper.index)] = "其他"
    mapper.loc[fund_category[fund_category.isin(["A股", "港股", "美股", "德国", "日本", "印度", "法国", "英国", "越南"])].index.intersection(mapper.index)] = "主动权益"
    mapper.loc[fund_basic[fund_basic["万得二级分类"].isin(["被动指数型基金", "增强指数型基金"])].index.intersection(mapper.index)] = "被动权益"
    mapper.loc[passive[passive.isin(["A股", "港股", "美股", "德国", "日本", "印度", "法国", "英国", "越南"])].index.intersection(mapper.index)] = "被动权益"
    mapper.loc[passive[passive.isin(["A债"])].index.intersection(mapper.index)] = "债券ETF"
    for code in ["002411.OF", "004047.OF", "015572.OF", "015779.OF", "519197.OF"]:
        mapper.loc[code] = "主动权益"
    mapper.fillna("偏债混合", inplace=True)
    mapper.index = mapper.index.str.slice(0, -3)
    mapper.loc["159223"] = "被动权益"
    mapper.loc["513950"] = "被动权益"
    mapper.loc["159519"] = "被动权益"
    mapper.loc["005010"] = "纯债债基"
    mapper.loc["378546"] = "其他商品"
    mapper = mapper.groupby(mapper.index).last()
    return mapper.rename("资产类型").to_frame()


def single_asset_category(row, half_year, allo, result, index_category):
    start = row.name - pd.offsets.QuarterEnd(2)

    allo_cache = allo.loc[row.index, start: row.name, :].reset_index()
    maximum = allo_cache.groupby("S_INFO_WINDCODE")["F_PRT_ENDDATE"].max()
    maximum = maximum[maximum >= maximum.max() - pd.DateOffset(months=6) - pd.Timedelta(days=1)]
    allo_cache = allo_cache.set_index(["S_INFO_WINDCODE", "F_PRT_ENDDATE"]).loc[list(zip(maximum.index, maximum.values))]
    allo_cache.index = allo_cache.index.droplevel(1)
    allo_cache = allo_cache.div(allo_cache.sum(axis=1), axis=0)

    neutral = result[result.isin(index_category[index_category.isin(["股票多空", "境内REITs"])].index)].index
    allo_cache = allo_cache.loc[allo_cache.index.difference(neutral)]
    big_miss = row.index.difference(allo_cache.index.unique())

    fund_port_convert = allo_cache["convert"].mul(row, fill_value=0).sum()
    rename = result.loc[allo_cache.index].replace(index_category)
    rename = rename[rename.isin(["白银", "黄金", "原油", "豆粕", "有色金属", "能化"])]
    fund_port_commody = allo_cache.sum(axis=1).loc[rename.index].mul(row).rename(rename).dropna()
    fund_port_commody = fund_port_commody.groupby(fund_port_commody.index).sum()
    allo_cache.drop(index=rename.index, inplace=True)
    fund_port_cash = allo_cache["cash"].mul(row, fill_value=0).sum()
    fund_port_other = allo_cache["other"].mul(row, fill_value=0).sum()

    fund_port_fund = allo_cache.copy()
    fund_port_fund["fund"] *= row
    fund_port_fund = fund_port_fund.rename(result).rename(index_category)
    fund_port_fund = fund_port_fund.groupby(fund_port_fund.index)["fund"].sum()

    cache = half_year[
        half_year["S_INFO_WINDCODE"].isin(row.index)
        & (half_year["F_PRT_ENDDATE"] >= start)
        & (half_year["F_PRT_ENDDATE"] <= row.name)
    ]
    maximum = cache.groupby("S_INFO_WINDCODE")["F_PRT_ENDDATE"].max()
    maximum = maximum[maximum >= maximum.max() - pd.DateOffset(months=6) - pd.Timedelta(days=1)]
    cache = cache.set_index(["S_INFO_WINDCODE", "F_PRT_ENDDATE"]).loc[list(zip(maximum.index, maximum.values))]

    hk = cache[cache["S_INFO_STOCKWINDCODE"].str.endswith("HK")]
    hk = hk.groupby(hk.index.get_level_values(0))["F_PRT_STKVALUETONAV"].sum() / 100
    a_stock = cache[
        cache["S_INFO_STOCKWINDCODE"].str.endswith("SZ")
        | cache["S_INFO_STOCKWINDCODE"].str.endswith("SH")
        | cache["S_INFO_STOCKWINDCODE"].str.endswith("BJ")
        | cache["S_INFO_STOCKWINDCODE"].str.endswith("NQ")
    ]
    a_stock = a_stock.groupby(a_stock.index.get_level_values(0))["F_PRT_STKVALUETONAV"].sum() / 100

    stock_alloc = allo_cache["stock"].mul(row, fill_value=0)
    disclosed_a = a_stock.reindex(allo_cache.index).fillna(0).mul(row, fill_value=0)
    disclosed_hk = hk.reindex(allo_cache.index).fillna(0).mul(row, fill_value=0)
    disclosed_total = disclosed_a.add(disclosed_hk, fill_value=0)
    disclosed_scale = (stock_alloc / disclosed_total).replace([np.inf, -np.inf], np.nan).clip(upper=1).fillna(1)
    disclosed_a = disclosed_a.mul(disclosed_scale, fill_value=0)
    disclosed_hk = disclosed_hk.mul(disclosed_scale, fill_value=0)

    residual_equity = stock_alloc.sub(disclosed_a, fill_value=0).sub(disclosed_hk, fill_value=0).clip(lower=0)
    fund_port_equity = residual_equity.rename(result).rename(index_category)
    fund_port_equity = fund_port_equity.groupby(fund_port_equity.index).sum()
    fund_port_equity.loc["A股"] = fund_port_equity.get("A股", 0) + disclosed_a.sum()
    fund_port_equity.loc["港股"] = fund_port_equity.get("港股", 0) + disclosed_hk.sum()

    fund_port_bond = allo_cache.copy()
    fund_port_bond["bond"] *= row
    fund_port_bond = fund_port_bond.rename(result).rename(index_category)
    fund_port_bond.rename({i: "A债" for i in fund_port_bond.index if i != "海外债券"}, inplace=True)
    fund_port_bond = fund_port_bond.groupby(fund_port_bond.index)["bond"].sum()

    fund_port = fund_port_bond.add(fund_port_equity, fill_value=0).add(fund_port_commody, fill_value=0)
    fund_port.loc["货币"] = fund_port_cash
    fund_port.loc["其他"] = fund_port_other
    fund_port.loc["可转债"] = fund_port_convert

    big_miss = row.loc[big_miss].rename(index=result.loc[big_miss].replace(index_category))
    big_miss = big_miss.groupby(big_miss.index).sum()
    fund_port = fund_port.add(big_miss, fill_value=0).add(fund_port_fund, fill_value=0)
    fund_port = fund_port.groupby(fund_port.index).sum()
    fund_port /= fund_port.sum()
    return fund_port.replace(0, np.nan).dropna().sort_values(ascending=False)
