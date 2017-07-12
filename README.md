# ExcelUtil
封装NPOI实现的excel导入工具集,支持excel2003和excel2007，无需安装excel

## 获取工作表名称列表

* 接口 
```
/// <summary>
/// 获取工作表名称列表
/// </summary>
/// <returns></returns>
List<string> GetSheets();
```
* 示例 
```
var fileName = Path.Combine(folder, "2003.xls");
var excel = ExcelHelper.NewExcelInstance(fileName);
var sheets = excel.GetSheets();
```

## 将DataTable写入Excel
* 接口 
```
/// <summary>
/// 将DataTable写入Excel
/// </summary>
/// <param name="dataTable">要写入的DataTable</param>
/// <param name="sheetName">工作表表名</param>
/// <param name="ifContainCaption">是否将列标题写入</param>
/// <returns></returns>
int WirteDataTable(DataTable dataTable, string sheetName, bool ifContainCaption);
```
* 示例 
```
var table = GetTestDataTable();
var fileName = Path.Combine(folder, "2003.xls");
var excel = ExcelHelper.NewExcelInstance(fileName);

var count = excel.WirteDataTable(table, "sheet1", true);
```


## 从Excel读取DataTable
* 接口 
```
/// <summary>
/// 从Excel读取DataTable
/// </summary>
/// <param name="sheetName">工作表名</param>
/// <param name="ifFirstRowCaption">首行是否当做列标题</param>
/// <returns></returns>
DataTable ReadDataTable(string sheetName, bool ifFirstRowCaption);
```
* 示例 
```
var fileName = Path.Combine(folder, "2003.xls");
var excel = ExcelHelper.NewExcelInstance(fileName);
var data = excel.ReadDataTable("sheet1", true);
```
