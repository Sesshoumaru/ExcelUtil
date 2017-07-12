using System.Linq;
using NPOI.SS.UserModel;

namespace ExcelUtil
{
    /// <summary>
    /// 行扩展方法
    /// </summary>
    public static class RowExtension
    {
        /// <summary>
        /// 获取Sheet中列索引为columnIndex的单元格
        /// </summary>
        /// <param name="row"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public static ICell GetCell1(this IRow row, int columnIndex)
        {
            return row.Cells.FirstOrDefault(c => c.ColumnIndex == columnIndex);
        }
    }
}