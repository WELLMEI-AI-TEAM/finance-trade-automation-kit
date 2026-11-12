# Excel 原生定点修改

## 首选：XLSX/XLSM 原包定点修改

对于 `.xlsx/.xlsm`，使用 `../scripts/apply_ooxml_plan.py`。仅依赖 Python 标准库，无须启动 Excel/WPS 或安装 Python 第三方包。计划格式见下文，与原生脚本一致；`format` 采用下述常见日期格式，不支持的复杂格式会拒绝，不自行改格式。

```powershell
& '<可用Python绝对路径>' '<skill绝对路径>\scripts\apply_ooxml_plan.py' --plan '<本次计划.json>'
```

Python 可用 Codex 的工作区依赖运行时（通过 load_workspace_dependencies 查询）或已有 Python。脚本保留 ZIP 中每个条目，只有标题/日期所在工作表内指定单元格的 payload/type 属性改变；其他单元格 XML、所有样式、行高列宽、图片、打印设置、公式和缓存均保持原字节。共享字符串标题会转为仅该单元格的 inlineStr，原共享字符串表保留，避免波及其他引用。无须重算公式。无需修改时输出与原 PI 完全相同的副本。签名包、公式目标、富文本目标、不支持的结构或日期均拒绝处理并列待核。

脚本保存后逐包条目重读核验，并输出 `.audit.json`，其中 `all_other_package_parts_identical=true` 有逐条字节验证依据。随后读取成品核对标题/日期并渲染涉及区域；为了查看而打开 Excel/WPS 时使用另一个预览副本，不能再另存覆盖已验证成品，否则需重新核验。

## XLS 原生处理及独立检查

自带 `../scripts/apply_excel_plan.ps1` 用独立、不可见的 Excel 或 WPS 实例，禁用宏、外链更新和自动计算；仅将 PI 的隔离副本交给原生引擎，验证后定点修改标题、日期，保存重开后比较。要求 Windows 和已安装的 Excel 或 WPS。Excel 无法打开或运行受阻时按用户要求使用 WPS，不要关闭用户正在运行的进程。脚本用实例 HWND/PID 与调用前进程清单判断所有权，复用已有用户进程即拒绝；退出前有非脚本工作簿则不 Quit。已有用户进程时不以关闭用户程序作为重试办法。

原生另存可能重新编码字体或行高列宽，即使原生应用里的快照看似相同，也不代表原始文件完全保留。因此 `.xls` 的计划 `output` 先使用独立草稿路径，保存后运行：

```powershell
& '<可用Python绝对路径>' '<skill绝对路径>\scripts\audit_xls.py' --plan '<同一计划.json>'
```

只读 BIFF 检查依赖 `xlrd`。若运行时没有，在本次输出目录隔离安装（`python -m pip install --target <本次依赖目录> xlrd`），仅给该次 Python 设置 PYTHONPATH，不修改系统依赖或固定桌面目录。检查实际值、字体/格式、混合字符样式、行列原始尺寸、合并、调色板、定义名称、公式 token 及打印/绘图记录。返回 `failed` 时不能关闭检查：判断差异是否影响要求，无法证明完全保真则列待核。仅 BIFF 检查、对象专项检查和视觉检查通过后才将草稿副本复制到最终输出目录，并按最终路径补充来源与哈希。`preservation.json` 会保留具体不通过项；不覆盖旧报告。

## 先读取输入

正文识别与关联由调用 skill 的助手完成；脚本不会自动从任意订单表、扫描件中猜编号或日期。Excel/WPS 只读 COM 仅打开隔离副本，不将业务原件交给原生引擎；Python openpyxl、xlrd 可辅助只读提取。PDF 先提取文字，扫描件需渲染/OCR并查看签订日期位置。发现无法确定的信息就列待核，不伪造证据来满足计划字段。

根据实际附件生成 UTF-8 JSON 计划，绝对路径仅限本次输入和独立输出目录：

```json
{
  "source": "C:/work/PI.xls",
  "output": "C:/work/output/出口合同/WMDS26001 ABC26001.xls",
  "source_sha256": "本次读取源PI的真实SHA256",
  "association": {
    "status": "confirmed",
    "order_confirmation": {
      "path": "C:/work/订单确认书.xlsx",
      "locator": "TOTAL!A18:L18",
      "evidence": "WMDS26001 → ABC26001 → PO391422，与PI内部订单号/PO一致"
    },
    "purchase_contract": {
      "path": "C:/work/购货合同.pdf",
      "locator": "第1页签订日期及合同编号",
      "evidence": "WMDS26001，签订日期2026-05-04",
      "signing_date": "2026-05-04"
    }
  },
  "title": {"sheet": "PI", "cell": "A1", "expected": "PROFORMA INVOICE"},
  "invoice_label": {"sheet": "PI", "cell": "T2", "expected": "Invoice No."},
  "date": {
    "sheet": "PI", "cell": "D3",
    "expected": "Date: May 05, 2026",
    "pi_date": "2026-05-05",
    "format": "MMMM dd, yyyy",
    "prefix": "Date: ", "suffix": "", "text_case": "none"
  }
}
```

- `title: null` 表示已检查，确无单据标题；不是“没找到就跳过”。已有 CONTRACT 标题可给 expected CONTRACT，不重复修改。
- `output` 的最终文件名按 `购货合同编号 我方订单号.原扩展名`，两个编号都必须来自本次正文核对，保留拆单后缀，不加“_出口合同”。XLS草稿可使用隔离工作路径，通过核验后再按该命名发布；同步记录最终路径和SHA256。
- `invoice_label` 定位**右上角**的 `Invoice No.` 标签；有则提供准确原值，未出现则显式设为 `null`。目标固定改为 `No.`，编号值及冒号/空白保留；例如同格 `Invoice No.: ABC1` 变为 `No.: ABC1`。不要将 PI No.或其他位置的同名标签传入此项。标题、标签、日期必须为不同单元格，字段共格或混合字体需另行保真处理或存入待核对。脚本省略该键只用于兼容旧计划；新计划必须先核对后填入或设null。
- 日期仅定位 **Date 后面的值**，先结合标签和排版核实，不能选 Shipment Date/Shipping Date/Ship Date/ETD。英文日期由助手读取原文解释为 ISO 日期，并用原格式完整复原后才修改。英文序数写法使用 `format: "MMMM d{ordinal}, yyyy"`，例如 `May 1st, 2026`；`ordinal_case: "upper"` 可保留 ST/ND/RD/TH，默认lower。序数后缀按目标日计算，不能把原st/th作为固定字面量沿用。11/12/13使用th。已有 `text_case` 仍控制整体月份等的大小写。
- `date` 必须指向已核实单据日期，`expected` 是准确的原 Value2。数值日期用 JSON number，文本日期用 JSON string，禁止将 Excel 序列号当作普通文字。`pi_date`、`signing_date` 是已核实的 ISO 日期；脚本自行比较并选择，不接受任意目标日期。
- 数值日期保留 NumberFormat 和时间小数，自动适配 1900/1904 日期系统。文本日期用 .NET 不变文化日期格式，例如 `yyyy-MM-dd`、`yyyy.M.d`、`dd/MM/yyyy`、`MMM dd, yyyy`、`MMMM dd, yyyy`。中文可用 `yyyy年M月d日`。`prefix/suffix` 原样保留；`text_case` 可为 `none`/`upper`/`lower`。脚本先用该格式复原原文本，必须完全一致后才允许编辑。日月顺序必须先从文件证据确定，不得以“某格式能解析”为判定依据。
- 单个单元格只能是完整标题或带前后空白的标题。标题和其他内容共用单元格、富文本/公式标题、日期公式/混合样式、图片内字段需另行保真处理，不能改 expected 绕过检查。
- 合并单元格指定左上角；明确有多个重复标题/日期时需逐个核对，自带单计划脚本覆盖一个标题、一个右上角编号标签和一个日期。不要遗漏后声称全部完成，也不借循环反复修改同一源文件覆盖前次结果。

## 执行

使用现有 PowerShell 7（`pwsh`）。原生核验可能较慢，按表大小提供进度；不要为提速跳过核验。默认先不导出 PDF，以免网络打印机导致等待；需要原生打印预览时可另行加入 `-ExportPreview`，隔离运行并设置合理超时，失败则使用其他渲染方式并如实记录分页检查未完成：

```powershell
& '<skill绝对路径>\scripts\apply_excel_plan.ps1' -PlanPath '<本次计划.json>'

# Excel 不可用时使用 WPS；使用新的输出名，保留先前失败诊断。
& '<skill绝对路径>\scripts\apply_excel_plan.ps1' -PlanPath '<本次计划.json>' -Engine WPS
```

脚本没有批量猜测功能。每份已确认对应的 PI 使用一个计划；不同文件可顺序处理。`-ExportPreview` 导出原件/成品的原生打印 PDF，需实际渲染查看后再报告视觉核验结果。PDF 导出失败会列在 audit 中，不等于修改失败，但视觉核验仍未完成。

输出同扩展名合同和 `.audit.json`。Excel 比对每张表的 SpreadsheetML 范围快照（只屏蔽实际获准改动的 Data 内容）；WPS 的该 XML 不含完整数据，所以改为逐格比较值、公式、NumberFormat、整格字体/边框/填充/对齐/合并、行高列宽及隐藏状态。两者均检查定义名称（包括打印区域/标题）、图形几何/文字、原件哈希和实际修改目标。`checked_snapshot_fields_unchanged` 仅表示这些检查通过，不表示全面保真。非目标单元格的混合字符样式仍需专项检查；不能以整格 Font 检查代替逐字符保真。打印机相关 PageSetup 属性在 Excel/WPS 中都可能等待网络打印机，脚本不调用这些属性；`unchecked_features` 列出的图片二进制、图表、条件格式、批注/超链接、页面设置等需专项检查及预览。含数据连接/查询表的工作簿拒绝直接编辑，需先调查刷新行为。存在原生数字签名时修改会使签名失效，保留原件并如实说明，不宣称签名仍有效。

失败不发布成品；若需调整计划，先理解失败原因，只修真实定位/格式错误。`association.status` 必须来自正文核对，不得为通过脚本随意设为 confirmed。不要用关闭检查、重建表格、去掉宏/图片来处理失败。
