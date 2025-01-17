# 标题重复检查工具

一个用于检查Excel文件中标题重复词的工具。

## 主要功能

- 支持多语言标题检查（英语、法语、西班牙语、意大利语、德语）
- 自动识别和过滤停用词（介词、冠词、连词等）
- 自动过滤度量单位（cm、kg、mm等）
- 支持重音字符处理
- 支持单复数形式识别
- 提供详细的日志记录

## 配置说明

配置文件 `config.yaml` 包含以下设置：

- 重复次数阈值
- 测试模式设置
- 列名配置
- 语言检测设置

## 使用方法

1. 准备Excel文件，确保包含标题列
2. 根据需要修改 `config.yaml` 配置
3. 运行程序
4. 查看结果：重复词将在Excel中以红色标记显示 