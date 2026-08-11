# 中芯核销拆票工具

纯前端Web应用，业务人员上传Excel文件即可自动完成拆票分组、高亮标注、生成拆分理由，并下载结果。

## 功能

- **自动拆票**：按账册号、业务申报表号、客户类型、征税/免表、131豁免、PO分组等维度自动分票
- **永芯40行限制**：一票最多40行，超出自动拆分子票
- **征税规则修正**：原产国精确为"中国"才拆征税，"中国台湾"不算
- **高亮标注**：同票同产品编号但以下任一维度不同时，整行黄色背景+不一致单元格红色字体：大PO / 小PO / 备案单位 / 原产国 / 出库备注 / 产品名称 / 客户料号（共7个维度，配置见 `js/config.js` 的 `highlightDimensions`）
- **uodId数据完整性校验**：拆票后自动比对原始数据与结果的uodId，检测行丢失/多余/重复/字段错位
- **拆分理由**：每票生成详细的分票理由说明
- **配置化**：业务规则变更只需修改 `js/config.js`，无需动引擎代码

## 使用方法

1. 打开网页
2. 拖拽或点击上传Excel文件（需含"中芯"sheet）
3. 自动拆票并展示结果
4. 查看"数据校验"标签页确认数据完整性
5. 下载结果Excel

## 技术栈

- 前端：纯 HTML/CSS/JavaScript，零后端依赖（Excel 处理用本地 `libs/xlsx.bundle.js`，已本地化，国内可直接打开）
- 可选后端：`server/audit_mail.py`（纯 Python 标准库）用于把留底自动发到管理员邮箱
- 可部署到 GitHub Pages / CloudStudio

## 留底说明（防扯皮）

每次业务人员上传并拆票后，系统会自动：

1. **本地留底**：结果文件 + 原始上传文件存入浏览器 IndexedDB，并自动下载到本机"下载"文件夹（命名含时间戳，可一一对应）。在"留底记录"标签页可查看/重新下载。
2. **邮件推送（可选）**：若配置了 `js/config.js` 的 `auditUploadUrl`，前端会把两份文件 POST 到该后端代理，由后端用 SMTP 发到管理员邮箱。

> 本地留底只存在各人自己的浏览器/电脑。要让管理员也能收到所有人的留底，必须部署 `server/audit_mail.py` 并填好 `auditUploadUrl`（见下文）。

## 部署留底邮件后端（管理员收到所有人的留底）

后端 `server/audit_mail.py` 是纯标准库实现，无需 `pip install`。凭据**只通过环境变量传入，绝不写进代码**。

```bash
cd server
SMTP_HOST=smtp.qiye.aliyun.com \
SMTP_USER=jiajun_cai@hmglog.com \
SMTP_PASS='你的授权码' \
RECIPIENT=jiajun_cai@hmglog.com \
ALLOWED_ORIGIN='https://你的网站域名' \
PORT=8000 \
python audit_mail.py
```

- 把后端部署到**公网可访问**的地址（如 阿里云函数计算 Custom Runtime + HTTP 触发器、或任意可跑 Python 的服务器），并配置 `ALLOWED_ORIGIN` 为你的网站域名（限制来源，防滥用）。
- 然后在 `js/config.js` 把 `auditUploadUrl` 设为该地址（如 `https://你的后端/audit`），重新部署前端即可生效。
- 前端每次拆票会最佳努力推送；推送失败不影响本地留底与下载。

## 文件结构

```
├── index.html          # 主页面
├── css/style.css       # 样式
├── js/config.js        # 业务规则配置（改规则/配邮件推送只改这个文件）
├── js/splitEngine.js   # 拆票引擎 + uodId校验
├── js/app.js           # UI交互逻辑（含留底与邮件推送）
├── libs/               # 本地化依赖（xlsx.bundle.js）
└── server/
    └── audit_mail.py   # 可选：留底邮件代理（纯标准库）
```
