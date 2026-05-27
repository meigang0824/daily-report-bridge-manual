# 日报桥接接口手动启动包

这个目录可以手动启动 Dify 日报桥接接口。它包含：

- `bridge/daily_report_bridge.py`：Dify 调用的 HTTP 接口。
- `legacy/main.py`：报表生成核心逻辑的脱敏副本。
- `bridge/requirements.txt`：Python 依赖。
- `workflow/每日店铺数据日报_完整迁移_Dify工作流.yml`：当前 Dify 工作流 DSL。
- `fixtures/daily_query_fixture.json`：仅供本地回归验证使用的测试数据。

## 启动步骤

```bash
cd "/path/to/daily-report-bridge-manual"
python3 -m venv .venv
source .venv/bin/activate
pip install -r bridge/requirements.txt
cp .env.example .env
# 编辑 .env，填写数据库与 Dify Dataset API Key
chmod +x start.sh
./start.sh
```

启动成功后访问：

```bash
curl http://localhost:8767/health
```

## 接管当前端口

目前已有服务监听 `8767`。你准备改为手动启动时，先停止原常驻进程：

```bash
screen -S daily-report-bridge -X quit
lsof -nP -iTCP:8767 -sTCP:LISTEN
```

如果第二条命令仍显示 Python 进程，执行 `kill <显示的PID>` 后再运行 `./start.sh`。

## 接口说明

工作流应使用直接返回 Excel 的接口：

```text
POST http://host.docker.internal:8767/v1/workflow/run-files-download
```

该 POST 接口的 `multipart/form-data` 输入：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `data_date` | text | 日报日期，如 `2026-02-06` |
| `shop_list_file` | file | 文件名必须为 `{运营姓名}-负责店铺列表.xlsx` |
| `supplement_files` | file-list | 一个或多个带日期的补单 Excel 文件 |

## 重要说明

- `run-files-download` 会直接响应生成后的 `.xlsx` 文件，并同步替换 Dify 知识库中该运营的日报。
- 不要同时启动多个监听 `8767` 端口的服务。
- `.env` 含数据库密码与 Dify API Key，不要对外发送。
- 包内 `legacy/main.py` 已移除原有硬编码凭证，必须使用 `.env` 提供数据库与知识库配置。
