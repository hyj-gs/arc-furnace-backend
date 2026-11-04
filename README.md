# 🔥 电弧炉炼钢过程记录 - 多人协作版

## 快速开始

### Mac / Linux
```bash
./start.sh
```

### Windows
```
双击 start.bat
```

启动后访问 http://localhost:3000

## 功能
- ✅ 6 张工艺记录表（配料/装料/熔化氧化/LF精炼/VD真空/浇注脱模）
- ✅ 按炉次组织数据，自动生成炉号
- ✅ 多人同时填写，数据实时共享
- ✅ 一键导出 Excel 汇总（7个 Sheet）
- ✅ SQLite 持久化存储
- ✅ 移动端友好界面

## API 列表
| 方法 | 路径 | 说明 |
|------|------|------|
| GET | /api/health | 健康检查 |
| GET | /api/stats | 统计 |
| GET | /api/templates | 6张表字段模板 |
| GET | /api/heats | 所有炉次列表 |
| POST | /api/heats | 创建炉次 |
| GET | /api/heats/:id | 单个炉次详情 |
| PUT | /api/heats/:id | 更新炉次 |
| DELETE | /api/heats/:id | 删除炉次 |
| POST | /api/heats/:id/forms/:no | 保存某张表数据 |
| GET | /api/export | 导出 Excel |

## 部署到公网
Glitch 一键：https://glitch.com/import/repo?REPO_URL
Render 一键：https://render.com/deploy?repo=REPO_URL
Railway 一键：https://railway.app/new/template?repo=REPO_URL

## 数据文件
SQLite 数据库：`data/arc-furnace.db`
