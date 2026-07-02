# 招行月报

这个目录用于生成招商渠道 FOF 月度运作报告：

- `monthly_report.py`：通用版月报，需要显式传入基金目录名称。
- `qinyang.py`：孟清扬版月报，默认基金为 `华夏保守养老`。
- `shaoqiang.py`：卢少强版月报，默认基金为 `华夏聚源优选`。

## 目录结构

```text
招行月报/
├── monthly_report.py                 # 通用版入口
├── qinyang.py                        # 孟清扬入口
├── shaoqiang.py                      # 国内权益拆分版入口
├── util.py                           # 本地数据读取、数据库连接和资产分类辅助函数
├── requirements.txt                  # Python 第三方依赖
├── README.md                         # 使用说明
└── data/                             # 本地数据目录，已被 .gitignore 忽略
    ├── 基准指数.xlsx                  # 基准指数映射配置
    └── YYYY-MM-DD/
        └── 基金目录名称/
            ├── *持仓收益分析*YYYY-MM-DD_YYYY-MM-DD*.xlsx
            └── *持仓收益分析*YYYY-MM-01_YYYY-MM-DD*.xlsx
```

脚本默认读取 `data/<date>/<fund>/` 下的招商持仓收益分析 Excel，并根据文件名里的起止日期自动识别周度、月度和后续区间数据。`data/` 已写入 `.gitignore`，上传 GitHub 时只上传代码，不上传本地数据。

## 环境安装

建议使用 Python 3.9 或更新版本。

```bash
cd 招行月报
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Windows PowerShell：

```powershell
cd 招行月报
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

`cx-Oracle` 需要本机已安装并能加载 Oracle Instant Client。`util.py` 会设置：

```python
os.environ.setdefault("NLS_LANG", "AMERICAN_AMERICA.ZHS16GBK")
```

如果连接 Oracle 时报动态库错误，请先检查 Oracle Client、系统环境变量和 Python 位数是否匹配。

## 数据库依赖

数据库连接配置集中在 `util.py` 的 `EMBEDDED_DB_CONFIG` 中。默认使用以下内部数据源：

```text
WIND_DB    Oracle
JY_DB      SQL Server
SAM_CH1    ClickHouse
```

运行环境需要能访问对应内部网络和数据库服务。

如果不想直接改代码，可以用环境变量覆盖连接配置，例如：

```powershell
$env:WIND_DB_HOST = "10.3.80.206"
$env:WIND_DB_USER = "chaxun"
$env:WIND_DB_PASSWORD = "******"
$env:WIND_DB_PORT = "1521"
$env:WIND_DB_DATABASE = "winddata"
```

可覆盖字段包括 `HOST`、`USER`、`PASSWORD`、`PORT`、`DATABASE`、`DB_TYPE`。

## 运行方式

先做语法检查：

```bash
python -m py_compile monthly_report.py qinyang.py shaoqiang.py util.py
```

### VSCode 调试

如果希望在 VSCode 里打断点调试，不需要手动输入命令行参数：

1. 打开本目录 `招行月报`。
2. 先打开 `build_report_cache.py`，按“运行-启动调试”生成缓存。
3. 再打开 `monthly_report.py`，按“运行-启动调试”生成报告。

断点可以打在 `build_report_cache.py`、`monthly_report.py` 或 `util.py` 中。

### 先取数缓存

公司服务器上建议先单独运行缓存脚本。默认只取 2023 年以来的数据：

```bash
python build_report_cache.py
```

缓存会生成在：

```text
data/cache/
```

之后再运行月报时，`monthly_report.py` 默认只读取这些 pickle 缓存，不会再连数据库取数。如果缓存缺失，会提示先运行 `build_report_cache.py`。

从本目录运行通用版：

```bash
python monthly_report.py 2026-03-31 华夏盈泰稳健
```

也可以直接运行 `python monthly_report.py`，此时会使用 `monthly_report.py` 顶部的 `DEFAULT_REPORT_DATE` 和 `DEFAULT_REPORT_FUND`。

如果临时想从其他日期开始重新取数：

```bash
python build_report_cache.py --start-date 20240101
```

运行孟清扬版：

```bash
python qinyang.py 2026-03-31
```

运行卢少强版：

```bash
python shaoqiang.py 2026-03-31
```

如果数据不在默认的 `data/` 目录，请显式传入数据根目录：

```bash
python monthly_report.py 2026-03-31 华夏盈泰稳健 --data-root /path/to/data
```

## 输入与输出

运行前需要准备：

- `data/基准指数.xlsx`：包含资产类型、基准指数名称和基准指数代码映射。
- `data/<date>/<fund>/`：每只基金一个目录，目录内放招商导出的持仓收益分析 Excel。

脚本会在当前工作目录生成 Word 文件：

```text
<基金名称>FOF运作报告模板.docx
```

因此建议先 `cd 招行月报` 再运行，避免报告输出到项目根目录或其他目录。

## 注意事项

- `util.py` 不再读取项目级配置文件，也不依赖预生成的 pickle 缓存。
- `monthly_report.py`、`qinyang.py`、`shaoqiang.py` 都会在解析命令行参数后才初始化数据库数据，因此 `--help` 不会触发连库。
- `monthly_report_domestic_abroad.py` 保留为历史参考脚本，不纳入本目录 `requirements.txt` 的独立运行承诺。
