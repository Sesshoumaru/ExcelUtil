using System.Collections.Generic;
using System.Data;

namespace ExcelUtil
{
    /// <summary>
    /// Excel功能接口
    /// </summary>
    public interface IExcel
    {
        /// <summary>
        /// 文件名
        /// </summary>
        string FileName { get; set; }

        /// <summary>
        /// 获取工作表名称列表
        /// </summary>
        /// <returns></returns>
        List<string> GetSheets();

        /// <summary>
        /// 将DataTable写入Excel
        /// </summary>
        /// <param name="dataTable">要写入的DataTable</param>
        /// <param name="sheetName">工作表表名</param>
        /// <param name="ifContainCaption">是否将列标题写入</param>
        /// <returns></returns>
        int WirteDataTable(DataTable dataTable, string sheetName, bool ifContainCaption);

        /// <summary>
        /// 从Excel读取DataTable
        /// </summary>
        /// <param name="sheetName">工作表名</param>
        /// <param name="ifFirstRowCaption">首行是否当做列标题</param>
        /// <returns></returns>
        DataTable ReadDataTable(string sheetName, bool ifFirstRowCaption);
    }
}
