# 💒 婚礼宾客统计系统

婚礼宾客名单管理工具：会场桌位平面图、拖拽排座、宾客名单管理、Excel 导入导出、人数统计。

后端 Python（FastAPI + SQLite），前端零依赖单页 HTML。

## 快速开始

```bash
./local-up.sh       # 启动（macOS / Linux）
./local-down.sh     # 停止

# Windows PowerShell
./local-up.ps1
./local-down.ps1
```

## 功能

| 页签 | 说明 |
|---|---|
| 🪑 桌位图 | 会场按比例平面图，半圆舞台 + 中线通道，圆桌按实际尺寸绘制；拖动圆桌摆放位置；拖拽宾客名牌到桌上入座；右侧宾客分布明细（按桌分组）+ 未安排宾客区 |
| 📋 宾客名单 | 宾客增删改查、搜索筛选、排序；行内修改邀请/确认状态；批量操作（换桌、设状态、删除）；Excel 导入导出 |
| 🛠 桌子管理 | 新增/编辑/删除桌子，批量新增，桌间宾客交换 |
| ⚙️ 设置 | 婚礼日期、默认桌容量、人数预算、会场尺寸、桌径、桌间距 |

## 项目结构

```
backend/
  app.py        # FastAPI 路由 + 业务逻辑
  storage.py    # SQLite 存储层
frontend/
  index.html    # 单页应用（原生 JS）
data/           # 运行时数据 wedding.db（首次启动自动生成）
```
