namespace ExcelUtil
{
    /// <summary>
    /// Excel帮助类
    /// </summary>
    public class ExcelHelper
    {
        /// <summary>
        /// 创建Excel操作实例
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <returns></returns>
        public static IExcel NewExcelInstance(string fileName)
        {
            var excel = fileName.IndexOf(".xlsx") > 0 ? (IExcel)new Excel2007(fileName) : new Excel2003(fileName);
            return excel;
        }
    }
}