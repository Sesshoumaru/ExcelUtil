using System;
using System.Data;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelUtil.Test
{
    [TestClass]
    public class ExcelUtilUnitTest
    {
        const string folder = @"E:\ExcelTest";

        private DataTable GetTestDataTable()
        {
            var table = new DataTable();
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Value", typeof(double));
            table.Columns.Add("CreateTime", typeof(DateTime));

            table.Rows.Add(1, "语文", 80, DateTime.Now.AddDays(-90));
            table.Rows.Add(2, "数学", 76, DateTime.Now.AddDays(-10));
            table.Rows.Add(3, "英语", 92, DateTime.Now.AddDays(-40));

            return table;
        }

        [TestMethod]
        public void Excel2003_GetSheets_Test()
        {
            var fileName = Path.Combine(folder, "2003.xls");
            var excel = ExcelHelper.NewExcelInstance(fileName);
            var sheets = excel.GetSheets();
            var expected = 4;

            Assert.IsTrue(expected == sheets.Count, "expected==sheets.Count");
        }

        [TestMethod]
        public void Excel2003_ReadDataTable_Test()
        {
            var fileName = Path.Combine(folder, "2003.xls");
            var excel = ExcelHelper.NewExcelInstance(fileName);
            var data = excel.ReadDataTable("sheet1", true);
            var expected = 5;

            Assert.IsTrue(expected == data.Rows.Count, "expected==data.Rows.Count");
        }


        [TestMethod]
        public void Excel2007_GetSheets_Test()
        {
            var fileName = Path.Combine(folder, "2007.xlsx");
            var excel = ExcelHelper.NewExcelInstance(fileName);
            var sheets = excel.GetSheets();
            var expected = 4;

            Assert.IsTrue(expected == sheets.Count, "expected==sheets.Count");
        }

        [TestMethod]
        public void Excel2007_ReadDataTable_Test()
        {
            var fileName = Path.Combine(folder, "2007.xlsx");
            var excel = ExcelHelper.NewExcelInstance(fileName);
            var data = excel.ReadDataTable("sheet1", true);
            var expected = 5;

            Assert.IsTrue(expected == data.Rows.Count, "expected==data.Rows.Count");
        }

        [TestMethod]
        public void Excel2003_WirteDataTable_Test()
        {
            var table = GetTestDataTable();
            var fileName = Path.Combine(folder, "2003.xls");
            var excel = ExcelHelper.NewExcelInstance(fileName);

            var count = excel.WirteDataTable(table, "sheet1", true);
            var expected = 3;

            Assert.IsTrue(expected == count, "expected==count");
        }
    }
}
