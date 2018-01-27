using System;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelUtil
{
    /// <summary>
    /// Excel功能基础类
    /// </summary>
    internal abstract class ExcelBase : IExcel, IDisposable
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="fileName">文件名</param>
        protected ExcelBase(string fileName)
        {
            FileName = fileName;
        }

        /// <summary>
        /// 工作表
        /// </summary>
        protected IWorkbook workbook;

        /// <summary>
        /// 文件流
        /// </summary>
        private FileStream fileStream;

        /// <summary>
        /// 资源释放标记
        /// </summary>
        private bool disposed;

        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 初始化工作表
        /// </summary>
        /// <param name="fs"></param>
        /// <returns></returns>
        protected virtual void InitWorkbook(FileStream fs)
        {
        }

        /// <summary>
        /// 获取工作表,表名不存在的时候获取第一个工作表
        /// </summary>
        /// <param name="sheetName">工作表名</param>
        /// <returns></returns>
        private ISheet GetSheet(string sheetName = null)
        {
            ISheet sheet;
            if (sheetName != null)
                sheet = workbook.GetSheet(sheetName) ?? workbook.GetSheetAt(0);
            else
                sheet = workbook.GetSheetAt(0);
            return sheet;
        }

        /// <summary>
        /// 获取工作表名称列表
        /// </summary>
        /// <returns></returns>
        public List<string> GetSheets()
        {
            var sheets = new List<string>();
            fileStream = new FileStream(FileName, FileMode.Open, FileAccess.Read);
            InitWorkbook(fileStream);
            for (var i = 0; i < workbook.NumberOfSheets; i++)
            {
                sheets.Add(workbook.GetSheetName(i));
            }
            return sheets;
        }

        /// <summary>
        /// 将DataTable写入Excel
        /// </summary>
        /// <param name="dataTable">要写入的DataTable</param>
        /// <param name="sheetName">工作表表名</param>
        /// <param name="ifContainCaption">是否将列标题写入</param>
        /// <returns></returns>
        public int WirteDataTable(DataTable dataTable, string sheetName, bool ifContainCaption)
        {
            var memoryStream = new MemoryStream();    //创建内存流用于写入文件       
            InitWorkbook(null);

            var sheet = workbook.GetSheet(sheetName) ?? workbook.CreateSheet(sheetName);

            var totalCount = 0;
            if (ifContainCaption) //写入DataTable的列名
            {
                var row = sheet.CreateRow(0);
                for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; ++columnIndex)
                {
                    var caption = dataTable.Columns[columnIndex].Caption;
                    caption = string.IsNullOrWhiteSpace(caption)
                        ? dataTable.Columns[columnIndex].ColumnName
                        : caption;
                    row.CreateCell(columnIndex).SetCellValue(caption);
                }
            }

            int rowIndex;
            for (rowIndex = 0; rowIndex < dataTable.Rows.Count; ++rowIndex)
            {
                IRow row = sheet.CreateRow(totalCount);
                for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; ++columnIndex)
                {
                    var cellValue = dataTable.Rows[rowIndex][columnIndex];
                    var value = cellValue == null || cellValue == DBNull.Value ? "" : cellValue.ToString();
                    row.CreateCell(columnIndex).SetCellValue(value);
                }
                ++totalCount;
            }
            workbook.Write(memoryStream);
            memoryStream.Flush();
            memoryStream.Position = 0;

            FileStream dumpFile = new FileStream(FileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
            memoryStream.WriteTo(dumpFile);//将流写入文件

            return totalCount;
        }

        /// <summary>
        /// 从Excel读取DataTable
        /// </summary>
        /// <param name="sheetName">工作表名</param>
        /// <param name="ifFirstRowCaption">首行是否当做列标题</param>
        /// <returns></returns>
        public DataTable ReadDataTable(string sheetName, bool ifFirstRowCaption)
        {
            var data = new DataTable();

            fileStream = new FileStream(FileName, FileMode.Open, FileAccess.Read);
            InitWorkbook(fileStream);

            var sheet = GetSheet(sheetName);
            if (sheet != null)
            {
                Dictionary<DataColumn, int> columnMap = new Dictionary<DataColumn, int>();
                IRow firstRow = sheet.GetRow(sheet.FirstRowNum);
                int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                int startRow;
                if (ifFirstRowCaption)
                {
                    for (int rowIndex = firstRow.FirstCellNum; rowIndex < cellCount; ++rowIndex)
                    {
                        var cell = firstRow.GetCell(rowIndex);
                        if (cell != null)
                        {
                            var cellValue = cell.ToString().Trim();
                            if (!string.IsNullOrWhiteSpace(cellValue))
                            {
                                var column = new DataColumn(cellValue);
                                data.Columns.Add(column);
                                columnMap.Add(column, rowIndex);
                            }
                        }
                    }
                    startRow = sheet.FirstRowNum + 1;
                }
                else
                {
                    startRow = sheet.FirstRowNum;
                }

                int rowCount = sheet.LastRowNum;
                for (int i = startRow; i <= rowCount; ++i)
                {
                    var row = sheet.GetRow(i);
                    if (row == null)
                        continue;

                    DataRow dataRow = data.NewRow();
                    foreach (var map in columnMap)
                    {
                        var cell = row.GetCell(map.Value);
                        {
                            dataRow[map.Key] = cell != null ? cell.ToString().Trim() : "";
                        }
                    }
                    data.Rows.Add(dataRow);
                }
            }

            return data;
        }



        /// <summary>
        /// 执行与释放或重置非托管资源相关的应用程序定义的任务。
        /// </summary>
        /// <filterpriority>2</filterpriority>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 执行与释放或重置非托管资源相关的应用程序定义的任务。
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    if (fileStream != null)
                        fileStream.Close();
                }

                fileStream = null;
                disposed = true;
            }
        }
    }
}