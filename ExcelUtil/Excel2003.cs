using System.IO;
using NPOI.HSSF.UserModel;

namespace ExcelUtil
{
    /// <summary>
    /// 2003版本Excel功能类
    /// </summary>
    internal class Excel2003 : ExcelBase
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="fileName">文件名</param>
        public Excel2003(string fileName) : base(fileName)
        {

        }

        /// <summary>
        /// 初始化工作表
        /// </summary>
        /// <param name="fs"></param>
        /// <returns></returns>
        protected override void InitWorkbook(FileStream fs)
        {
            workbook = fs!= null ? new HSSFWorkbook(fs) : new HSSFWorkbook();
        }
    }
}