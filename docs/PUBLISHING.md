# 安装与发布说明

参照团队现有项目采用中文 README、skills 源码、可运行脚本、回归测试、自检范围、独立分发包和 SHA256 清单。

1. 安装 requirements-dev.txt，运行 `python -B -m pytest tests -q`。
2. 运行 `python -B tools/build_release.py`，按明确资源目录白名单构建三个 ZIP。
3. 运行 `python -B tools/verify_release.py`，独立检查压缩包完整性、SKILL.md 根入口、文件清单和逐文件字节。
4. 分发 dist/upload 下对应的 ZIP。ZIP 根层为技能内容，安装时解压到对应技能标识目录，避免多套一层目录。

三个文件技能沿用原 name，不为打包改写业务规则；财务本机绑定、测试 profiles 及实际账户不分发。构建可重复，版本以 suite.json 为准。

上传 GitHub 不等于已导入任何技能市场。本次不声称已经完成 SkillHub 审核、跨客户端菜单验收、新电脑服务连通或真实业务批次处理。需要 Windows Excel/WPS 的功能必须在目标环境按实际文件验收。
