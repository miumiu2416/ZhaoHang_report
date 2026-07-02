"""Fetch database-backed report data and save it to data/cache/*.pkl.

Run this first on the company server. After it finishes, monthly_report.py can
generate reports from the cached pickle files without querying databases.
"""

import argparse

import monthly_report as report


def parse_args():
    parser = argparse.ArgumentParser(description="缓存月报所需数据库数据。")
    parser.add_argument(
        "--data-root",
        default=str(report.DEFAULT_DATA_ROOT),
        help="数据目录根路径，缓存会写入 <data-root>/cache/。默认 data/。",
    )
    parser.add_argument(
        "--start-date",
        default=report.DEFAULT_CACHE_START_DATE,
        help="取数起始日期，格式 YYYYMMDD。默认 20230101。",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    report.init_report_data(
        args.data_root,
        use_cache=True,
        refresh_cache=True,
        save_cache=True,
        cache_only=False,
        cache_start_date=args.start_date,
    )
    print(f"缓存完成：{report.get_cache_dir(args.data_root)}")


if __name__ == "__main__":
    main()
