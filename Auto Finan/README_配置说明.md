# 配置文件说明

## 概述

程序现在支持从JSON配置文件读取Excel文件路径和工作表名称，使配置更加灵活。

## 配置文件

### 文件名
`config.json`

### 查找路径
程序会按以下顺序查找配置文件：
1. 当前目录：`config.json`
2. 当前工作目录：`{当前工作目录}/config.json`
3. 上级目录：`{当前工作目录}/../config.json`
4. 上上级目录：`{当前工作目录}/../../config.json`
5. 上上上级目录：`{当前工作目录}/../../../config.json`
6. 上上上上级目录：`{当前工作目录}/../../../../config.json`

### 配置项说明

```json
{
  "ExcelFilePath": "报销信息.xlsx",
  "MappingFilePath": "标题-ID.xlsx", 
  "SheetName": "ChaiLv_sheet",
  "MappingSheetName": "Sheet1"
}
```

- **ExcelFilePath**: 主要Excel文件名
- **MappingFilePath**: 标题-ID映射表文件名
- **SheetName**: 主要Excel文件中的工作表名
- **MappingSheetName**: 标题-ID映射表中的工作表名

## 默认配置

如果找不到配置文件，程序会使用以下默认值：
- ExcelFilePath: "报销信息.xlsx"
- MappingFilePath: "标题-ID.xlsx"
- SheetName: "ChaiLv_sheet"
- MappingSheetName: "Sheet1"

## 使用方法

1. 创建`config.json`文件
2. 根据需要修改配置项
3. 将配置文件放在程序能找到的目录中
4. 运行程序，会自动读取配置

## 示例

### 基本配置
```json
{
  "ExcelFilePath": "报销信息.xlsx",
  "MappingFilePath": "标题-ID.xlsx",
  "SheetName": "ChaiLv_sheet",
  "MappingSheetName": "Sheet1"
}
```

### 自定义配置
```json
{
  "ExcelFilePath": "我的报销数据.xlsx",
  "MappingFilePath": "字段映射表.xlsx",
  "SheetName": "报销明细",
  "MappingSheetName": "映射表"
}
```

## 注意事项

1. JSON文件必须使用UTF-8编码
2. 配置项名称区分大小写
3. 如果配置文件格式错误，程序会使用默认配置
4. Excel文件的查找路径与配置文件查找路径保持一致

## 测试

运行`test_config.bat`来测试配置功能。







