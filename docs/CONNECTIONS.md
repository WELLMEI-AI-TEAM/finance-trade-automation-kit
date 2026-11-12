# 财务服务连接条件

财务发票技能通过现有 HTTP 服务识别发票、验证身份和生成台账。服务必须支持 finance 账户、显式发票 ID 导出、模板规范化、源文件清单、归档恢复和复核接口。旧版能力不足时客户端停止发布结果。

本项目交付客户端技能及本地归档、备份、异常与恢复逻辑，**不交付服务端或模型密钥**。安装者需要由服务维护者提供实际服务地址、财务用户名、凭据文件路径或已设置的密码环境变量名；不使用他人的测试账户。

先阅读 [工作区设置](../skills/finance-invoice-ledger/references/workspace-setup.md)，目录初始化不联网，也不会生成虚假账本。配置详情与旧版兼容见 [configuration.md](../skills/finance-invoice-ledger/references/configuration.md)。凭据与 binding.local.json 只保存在使用者本机，不上传仓库或放入 ZIP。
