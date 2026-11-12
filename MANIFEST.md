# 发布清单

版本与三个技能标识以 [suite.json](suite.json) 为准。

| 目录 | 内容 |
|---|---|
| skills/pi-to-export-contract | SKILL.md、agents、references、合同编辑和批次备份脚本 |
| skills/finance-invoice-ledger | SKILL.md、agents、references、工作区初始化与财务台账客户端脚本 |
| skills/contract-print-layout | SKILL.md、agents、打印候选及保真检查脚本 |
| tests | 使用临时文件与模拟服务的可复现回归测试 |
| tools | 显式白名单打包及独立包核验 |
| docs | 连接条件、安装与发布说明 |
| dist/upload | 三个独立 ZIP、逐文件哈希清单、ZIP 校验和 |

三个技能独立安装，无彼此目录依赖。财务技能对外依赖兼容服务端；合同编辑与打印整理对原生 Office 的要求详见各自 SKILL.md。

分发时不包含 profiles、binding.local.json、.managed-source.json、凭据、业务文件、缓存或运行记录。仓库根目录的 LICENSE 原文以 LICENSE.txt 随各 ZIP 分发。
