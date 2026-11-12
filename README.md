# 财务与外贸自动化技能套件

三个可独立安装的文件技能，覆盖 **PI 转出口合同、财务发票台账、合同打印整理**。本仓库发布已经完成的技能及配套脚本，业务规则沿用原技能；合同生成与打印整理分别按用户请求执行。

版本：**1.0.0** · 2026-11-12

| 技能 | 能力 | 示例请求 | 运行条件 |
|---|---|---|---|
| [PI 转出口合同](skills/pi-to-export-contract/SKILL.md) | 核对 PI、订单确认书与购货合同，保留原布局生成合同副本 | 处理这些 PI，制作出口合同 | Python 3.10+；XLS 原生编辑需 Windows 与 Excel/WPS |
| [财务发票台账](skills/finance-invoice-ledger/SKILL.md) | 发票识别、按开票月份去重归档、追加 Excel 台账、异常分流与备份 | 处理待处理目录中的财务发票 | Python 3.10+；可访问的兼容财务台账服务及 finance 账户 |
| [合同打印整理](skills/contract-print-layout/SKILL.md) | 调整打印范围、缩放、分页，交付 Excel 打印版与 PDF 打印版 | 把这些合同整理为清晰可打印的版本 | Windows、Python 3.10+、Excel/WPS；需逐页视觉核验 |

## 安装

将 `skills/` 下需要的**整个技能文件夹**复制到客户端支持的技能目录。例如 Codex 的 `%USERPROFILE%\.codex\skills\`。更新已有技能前先备份、比对本机修改；新会话确认技能已加载。

安装所选技能目录中的 `requirements.txt`，或在独立 Python 环境运行：

```powershell
python -m pip install -r requirements.txt
```

单独分发包位于 [dist/upload](dist/upload/)。每个 ZIP 根层直接包含一个 `SKILL.md` 及其引用的资料、脚本和依赖清单；解压到以技能标识命名的目录。使用支持本机脚本与文件的客户端。SkillHub 对 `.ps1` 等类型的实际接受情况需要现场验证，不能视为已上架。

## 使用与配置

- PI 转合同只按已确认的标题、右上角编号标签、日期规则编辑，其余内容与样式保留；无法唯一关联的输入进入待核对。正文识别与关联由助手结合真实资料完成，脚本不猜订单和日期。
- 财务技能首次使用时询问建立“台账处理”的位置，然后配置本机实际可访问的服务、finance 账户与凭据位置。**本仓库包含客户端，不包含台账服务端、OCR/模型服务、账户或 API 密钥**；仅安装技能不等于服务已可用。详见 [首次设置](skills/finance-invoice-ledger/references/workspace-setup.md) 与 [连接条件](docs/CONNECTIONS.md)。
- 打印整理先生成候选，逐页核对 PDF 后再发布。优先 A4 纵向清晰的一页，较长时两页或更多页；需要时采用窄边距或横向。不会自动发送实体打印任务。

原文件保留策略按各技能执行：PI 转合同及打印整理保留源文件；财务入口只在归档、入账或异常保存及必要备份等校验完成后，清理本轮“待处理”内内容未变化的输入副本，外部附件原路径保留。

## 自检与分发

```powershell
python -m pip install -r requirements-dev.txt
python -B -m pytest tests -q
python -B tools/build_release.py
python -B tools/verify_release.py
```

验证范围与限制见 [SELFTEST.md](SELFTEST.md)，安装包结构见 [MANIFEST.md](MANIFEST.md)，发布说明见 [docs/PUBLISHING.md](docs/PUBLISHING.md)。本仓库不含真实发票、合同、台账、本机绑定、凭据及运行记录。许可证沿用团队项目采用的 [Apache-2.0](LICENSE)。
