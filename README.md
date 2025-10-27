# hanair-data

海航随心飞航线数据自动化抓取工具。

本仓库提供一个 Python 脚本，用于从海航官网公告页面（默认链接为
<https://m.hnair.com/cms/me/plus/info/202508/t20250808_78914.html>）抓取表格数据，
并将结果写入 Excel 工作簿。脚本会在工作簿中维护一个“Latest”工作表来保存最新
一次抓取的完整表格，同时默认会生成以当天日期命名的历史工作表，便于每日更新。

## 快速开始

1. 安装依赖（需要 Python 3.9+）：

   ```bash
   pip install -r requirements.txt
   ```

2. 运行抓取脚本：

   ```bash
   python -m hanair_data.table_updater
   ```

   默认会在当前目录生成 `hnair_table.xlsx`，其中包含最新抓取的数据和一份以当天
   日期命名的历史副本。如果只想更新“Latest”工作表，可添加 `--skip-history` 参数。

3. 常用参数：

   ```bash
   python -m hanair_data.table_updater \
       --output data/hnair_table.xlsx \
       --latest-sheet-name 最新 \
       --history-sheet-name "2024年08月08日"
   ```

   - `--url`：指定其它页面链接。
   - `--table-index`：当页面有多个表格时可指定抓取第几个表格（从 0 开始）。
   - `--skip-history`：仅更新最新工作表，不保留历史记录。
   - `--timeout`：设置网络请求的超时时间（秒）。

## 每日自动更新

可以结合系统的计划任务每日执行脚本，例如在 Linux/macOS 上使用 cron：

```bash
0 8 * * * /usr/bin/python -m hanair_data.table_updater --output /path/to/hnair_table.xlsx >> /path/to/hnair_table.log 2>&1
```

上述命令表示每天早上 8 点运行脚本，并将日志追加写入 `hnair_table.log`。

在 Windows 上，可以通过“任务计划程序”创建每日任务，执行命令：

```
python -m hanair_data.table_updater --output C:\\path\\to\\hnair_table.xlsx
```

## 运行结果

脚本执行成功后会在终端打印抓取的行数和抓取时间，同时 Excel 文件的 `Latest`
工作表会被刷新，历史表（默认命名为 `YYYY-MM-DD`）也会更新为当天的数据。

如需将数据进一步处理或分析，可直接使用 Excel 或其它支持 xlsx 的工具继续操作。
