using System.Data;

namespace ExcelNpoi
{
    //*****************************************
    /// <summary>
    /// NPOI操作Excel 接口
    /// @author:tengxiaohui
    /// time:2017-08-03
    /// </summary>
    //*****************************************
    public interface IExcelNpoi
    {
        /// <summary>
        /// Npoi读取文件
        /// </summary>
        /// <param name="fileName">文件</param>
        /// <returns>数据表</returns>
        DataTable ReadFileFromExcel(string file);

        /// <summary>
        /// 导出数据
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="fileName">导出文件名</param>
        /// <param name="objs"></param>
        void ExportToExcel(DataTable dt, string file, params object[] objs);

        /// <summary>
        /// 导出数据
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="fileName">导出文件名</param>
        /// <param name="header">数据头</param>
        /// <param name="commnet">批注内容</param>
        /// <param name="headRows">数据头行</param>
        void ExportToExcel(DataTable dt, string file, string header = null, string commnet = null, int headRows = 0);
    }
}



