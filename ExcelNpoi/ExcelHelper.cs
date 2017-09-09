using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.OpenXmlFormats.Dml.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;


namespace ExcelNpoi
{
    /// <summary>
    /// Excel操作方法
    /// </summary>
    public class ExcelHelper
    {
        #region 公有变量

        /// <summary>
        /// Excel文件路径
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 默认工作表名称
        /// </summary>
        public string SheetName { get; set; }

        #endregion

        #region 私有变量

        //工作薄
        private IWorkbook workbook = null;

        //sheet表
        private ISheet sheet = null;

        //文件流
        private FileStream filestream = null;

        //操作Excel文件的方式
        private ExcelOperateMode operatetype;

        //Excel类型
        private ExcelType exceltype;

        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="strFileName">Excel文件路径</param>
        /// <param name="strSheetName">Sheet名称</param>
        /// <param name="OperateType">操作Excel方式（打开、创建）</param>
        public ExcelHelper(string strFileName, string strSheetName, ExcelOperateMode OperateType)
        {
            FileName = strFileName;
            SheetName = strSheetName;
            operatetype = OperateType;
            exceltype = ExcelType.DEFAULT;
        }
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="strFileName">Excel文件路径</param>
        /// <param name="strSheetName">Sheet名称</param>
        /// <param name="OperateType">操作Excel方式（打开、创建）</param>
        /// <param name="pExcelType">Excel文件类型（两种：XLSX，XLS）</param>
        public ExcelHelper(string strFileName, string strSheetName, ExcelOperateMode OperateType, ExcelType pExcelType)
        {
            FileName = strFileName;
            SheetName = strSheetName;
            operatetype = OperateType;
            exceltype = pExcelType;
        }
        #endregion

        #region Excel基本操作（新建、打开、保存、另存为、关闭）
        /// <summary>
        /// 打开Excel文件
        /// </summary>
        /// <returns>打开成功返回true，打开失败返回false</returns>
        public bool Open()
        {
            bool OK = false;
            try
            {
                //1、以文件流的方式打开Excel
                FileStream fileStream = new FileStream(FileName, FileMode.Open, FileAccess.ReadWrite);
                //2、初始化工作薄
                InitializeWorkbook(fileStream);
                //3、获取sheet
                sheet = workbook.GetSheet(SheetName);
                //4、关闭文件流
                fileStream.Close();
                OK = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return OK;
        }
        /// <summary>
        /// 初始化工作薄（open）
        /// </summary>
        /// <param name="fileStream">文件流</param>
        private void InitializeWorkbook(FileStream fileStream)
        {
            try
            {
                switch (exceltype)
                {
                    case ExcelType.DEFAULT:
                        string strExtension = Path.GetExtension(FileName).ToLower();
                        switch (strExtension)
                        {
                            case ".xlsx":
                                workbook = new XSSFWorkbook(fileStream);
                                break;
                            case ".xls":
                                workbook = new HSSFWorkbook(fileStream);
                                break;
                            default:
                                workbook = null;
                                break;
                        }
                        break;
                    case ExcelType.XLSX:
                        workbook = new XSSFWorkbook(fileStream);
                        break;
                    case ExcelType.XLS:
                        workbook = new HSSFWorkbook(fileStream);
                        break;
                    default:
                        workbook = null;
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// 初始化工作薄（Create）
        /// </summary>
        private void InitializeWorkbook()
        {
            try
            {
                switch (exceltype)
                {
                    case ExcelType.DEFAULT:
                        string strExtension = Path.GetExtension(FileName).ToLower();
                        switch (strExtension)
                        {
                            case ".xlsx":
                                workbook = new XSSFWorkbook();
                                break;
                            case ".xls":
                                workbook = new HSSFWorkbook();
                                break;
                            default:
                                workbook = null;
                                break;
                        }
                        break;
                    case ExcelType.XLSX:
                        workbook = new XSSFWorkbook();
                        break;
                    case ExcelType.XLS:
                        workbook = new HSSFWorkbook();
                        break;
                    default:
                        workbook = null;
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// 创建Excel文件
        /// </summary>
        /// <returns>创建成功返回true，创建失败返回false</returns>
        public bool Create()
        {
            bool OK = false;
            try
            {
                //1、初始化工作薄
                InitializeWorkbook();
                //2、创建sheet
                sheet = workbook.CreateSheet(SheetName);
                OK = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return OK;
        }
        /// <summary>
        /// 保存Excel文件
        /// </summary>
        public bool Save()
        {
            bool OK = false;
            try
            {
                //分为两种：新建和打开
                if (operatetype == ExcelOperateMode.create)
                {
                    filestream = new FileStream(FileName, FileMode.Create);
                }
                else
                {
                    File.Delete(FileName);
                    filestream = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                }
                workbook.Write(filestream);
                OK = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return OK;
        }
        /// <summary>
        /// 另存为Excel文件
        /// </summary>
        /// <param name="strFileName">另存为路径</param>
        public bool SaveAs(string strFileName)
        {
            bool OK = false;
            try
            {
                filestream = new FileStream(strFileName, FileMode.Create);
                workbook.Write(filestream);
                OK = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return OK;
        }
        /// <summary>
        /// 关闭Excel文件
        /// </summary>
        public void Close()
        {
            filestream.Close();
            workbook = null;
            sheet = null;
        }
        #endregion

        #region 读取Excel

        /// <summary>
        /// 读取一定范围内的内容
        /// </summary>
        /// <param name="col1">最小列号</param>
        /// <param name="minrow">最小行号</param>
        /// <param name="maxcol">最大列号</param>
        /// <param name="maxrow">最大行号</param>
        /// <returns>在此区间的内容</returns>
        public DataTable ReadRegionValue(int mincol, int minrow, int maxcol, int maxrow)
        {
            IRow row = null;
            DataRow dr = null;
            DataColumn column = null;
            ICell cell = null;
            DataTable dtData = new DataTable();
            try
            {
                if (maxcol < mincol && maxrow < minrow)
                    return null;
                int intFirstRow = sheet.FirstRowNum;
                int intLastRow = sheet.LastRowNum;
                IRow firstrow = sheet.GetRow(intFirstRow);
                for (int i = mincol; i < maxcol + 1; i++)
                {
                    string strColumnName = firstrow.GetCell(i).ToString();
                    column = new DataColumn(strColumnName);
                    dtData.Columns.Add(column);
                }
                for (int i = minrow; i < maxrow + 1; i++)
                {
                    row = sheet.GetRow(i);
                    dr = dtData.NewRow();
                    for (int j = 0; j < maxcol - mincol + 1; j++)
                    {
                        cell = row.GetCell(j);
                        dr[j] = cell.ToString();
                    }
                    dtData.Rows.Add(dr);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return dtData;
        }

        /// <summary>
        /// 读取单元格内容
        /// </summary>
        /// <param name="intRowIndex">行号</param>
        /// <param name="intColumnIndex">列号</param>
        /// <returns>单元格内容</returns>
        public string ReadCellValue(int intRowIndex, int intColumnIndex)
        {
            string strCellValue = string.Empty;
            try
            {
                IRow row = sheet.GetRow(intRowIndex);
                if (row == null)
                {
                    return strCellValue;
                }
                ICell cell = row.GetCell(intColumnIndex);
                if (cell == null)
                {
                    return strCellValue;
                }
                strCellValue = cell.StringCellValue;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return strCellValue;
        }

        /// <summary>
        /// 读取Excel数据至DataTable中（读取sheet中的全部数据）
        /// </summary>
        /// <returns>成功返回datatable，失败返回null</returns>
        public DataTable ReadExcelToDatatable()
        {
            DataTable dtData = GetTableFromSheet();
            return dtData;
        }

        /// <summary>
        /// 读取Excel中的数据，生成DataSet，每个sheet为一个DataTable
        /// </summary>
        /// <returns>DataSet</returns>
        public DataSet ReadExcelToDataSet()
        {
            string strFileName = Path.GetFileNameWithoutExtension(FileName);
            DataSet ds = new DataSet(strFileName);
            DataTable dtData = null;
            try
            {
                int iSheetCount = workbook.NumberOfSheets;      //获取Excel中Sheet数量
                if (iSheetCount == 0)
                    return null;
                for (int i = 0; i < iSheetCount; i++)
                {
                    sheet = workbook.GetSheetAt(i);
                    dtData = GetTableFromSheet();
                    ds.Tables.Add(dtData);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return ds;
        }

        /// <summary>
        /// 将sheet转为DatatTable
        /// </summary>
        /// <returns>成功返回DataTable，失败返回null</returns>
        private DataTable GetTableFromSheet()
        {
            IRow row = null;
            DataRow dr = null;
            DataColumn column = null;
            DataTable dtData = new DataTable(sheet.SheetName);
            IEnumerator rows = sheet.GetRowEnumerator();
            try
            {
                int intFirstRow = sheet.FirstRowNum;
                IRow firstrow = sheet.GetRow(intFirstRow);
                //添加列名
                int intFirstCellNum = firstrow.FirstCellNum;
                int intLastCellNum = 6;
                for (int i = intFirstCellNum; i < intLastCellNum; i++)
                {
                    string strColumnName = firstrow.GetCell(i).ToString();
                    column = new DataColumn(strColumnName);
                    dtData.Columns.Add(column);
                }
                //添加行
                while (rows.MoveNext())
                {
                    row = (IRow)rows.Current;
                    if (row.RowNum == intFirstRow)
                    {
                        continue;
                    }
                    dr = dtData.NewRow();
                    ICell cell = null;
                    for (int i = 0; i < 6; i++)
                    {
                        cell = row.GetCell(i);
                        dr[i] = cell.ToString();
                    }
                    dtData.Rows.Add(dr);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return dtData;
        }

        /// <summary>
        /// 将sheet转为string
        /// </summary>
        /// <returns>成功返回List<string></returns>
        public List<string> GetListFromSheet()
        {
            List<string> list = new List<string>();
            IRow row = null;
            IEnumerator rows = sheet.GetRowEnumerator();
            try
            {
                int intFirstRow = sheet.FirstRowNum;
                IRow firstrow = sheet.GetRow(intFirstRow);

                //添加行
                while (rows.MoveNext())
                {
                    row = (IRow)rows.Current;
                    StringBuilder sb = new StringBuilder();
                    ICell cell = null;
                    for (int i = 0; i < 6; i++)
                    {
                        cell = row.GetCell(i);
                        if (cell != null)
                            sb.Append(cell.ToString() + ",");
                    }
                    if (!string.IsNullOrEmpty(sb.ToString()))
                    {
                        int index = sb.ToString().LastIndexOf(',');
                        list.Add(sb.ToString().Substring(0, index));
                    }

                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return list;
        }

        #endregion

        #region 导出、修改Excel
        /// <summary>
        /// 将DataTable中的数据导为Excel文件
        /// </summary>
        /// <param name="dt">DataTable数据</param>
        public bool ExportDataTableToExcel(DataTable dt)
        {
            bool OK = false;
            int count = 0;
            IRow row = null;
            ICell cell = null;
            try
            {
                //单元格样式
                ICellStyle Cellstyle = SetRowCellStyle();
                ICellStyle Columnstyle = SetColumnCellStyle();
                //添加列名
                row = sheet.CreateRow(0);                                           //新建行
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    string strColumnName = dt.Columns[i].ColumnName;                  //获取列名
                    int intCellWidth = strColumnName.ToCharArray().Length * 2 * 256;   //获取列宽
                    sheet.SetColumnWidth(i, intCellWidth);
                    cell = row.CreateCell(i);
                    cell.SetCellValue(strColumnName);

                    cell.CellStyle = Columnstyle;
                }
                count = 1;
                //遍历Datatable，获取单元格内容写入Cell中
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    row = sheet.CreateRow(count);                                        //行高
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        string strCellValue = dt.Rows[i][j].ToString();
                        cell = row.CreateCell(j);
                        cell.SetCellValue(strCellValue);
                        cell.CellStyle = Cellstyle;

                        if (strCellValue.ToCharArray().Length < 255)
                        {
                            int intCellWidth = strCellValue.ToCharArray().Length * 256;
                            if (intCellWidth > sheet.GetColumnWidth(j))
                            {
                                sheet.SetColumnWidth(j, intCellWidth);
                            }
                        }
                    }
                    count++;
                }
                OK = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return OK;
        }

        /// <summary>
        /// 设置单元格样式（“行”单元格）
        /// </summary>
        /// <returns></returns>
        private ICellStyle SetRowCellStyle()
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            //1、内容位置（居中）
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            //2、字体设置
            IFont font = workbook.CreateFont();
            font.FontName = "微软雅黑";                          //字体名称
            font.FontHeightInPoints = 10;                        //字号
            font.Color = HSSFColor.Black.Index;                  //字体颜色
            //font.Underline = FontUnderlineType.None;              //下划线
            font.IsStrikeout = false;                            //删除线
            //3、背景色、前景色
            cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;   //图案颜色
            cellStyle.FillPattern = FillPattern.NoFill;                        //图案样式
            cellStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;   //背景色
            //4、设置边框样式
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            //5、设置边框颜色
            cellStyle.TopBorderColor = HSSFColor.Black.Index;
            cellStyle.BottomBorderColor = HSSFColor.Black.Index;
            cellStyle.LeftBorderColor = HSSFColor.Black.Index;
            cellStyle.RightBorderColor = HSSFColor.Black.Index;

            cellStyle.SetFont(font);
            return cellStyle;
        }
        /// <summary>
        /// 设置单元格样式（“表头”单元格）
        /// </summary>
        /// <returns></returns>
        private ICellStyle SetColumnCellStyle()
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();
            //1、内容位置（居中）
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            //2、字体设置
            IFont font = workbook.CreateFont();
            font.FontName = "微软雅黑";                          //字体名称
            font.FontHeightInPoints = 11;                        //字号
            font.Color = HSSFColor.Black.Index;                  //字体颜色
            font.Boldweight = 30;

            //3、背景色、前景色（不知道是什么原因在这里设置不管用）
            cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.White.Index;   //图案颜色
            cellStyle.FillPattern = FillPattern.SolidForeground;                  //图案样式
            cellStyle.FillBackgroundColor = HSSFColor.Red.Index;                  //背景色
            //4、设置边框样式
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            //5、设置边框颜色
            cellStyle.TopBorderColor = HSSFColor.Black.Index;
            cellStyle.BottomBorderColor = HSSFColor.Black.Index;
            cellStyle.LeftBorderColor = HSSFColor.Black.Index;
            cellStyle.RightBorderColor = HSSFColor.Black.Index;

            cellStyle.SetFont(font);
            return cellStyle;
        }

        /// <summary>
        /// 修改单元格内容
        /// </summary>
        /// <param name="intRowIndex">行号</param>
        /// <param name="intColumnIndex">列号</param>
        /// <param name="objCellValue">单元格值</param>
        /// <param name="celltype">单元格类型</param>
        /// <returns>修改成功返回true，失败返回false</returns>
        public bool WriteCellValue(int intRowIndex, int intColumnIndex, object objCellValue, CellType celltype)
        {
            bool OK = false;
            try
            {
                //获取行、单元格
                IRow row = sheet.GetRow(intRowIndex);
                ICell cell = row.GetCell(intColumnIndex);
                //设置单元格类型
                cell.SetCellType(celltype);
                switch (celltype)
                {
                    case CellType.String:
                        cell.SetCellValue(objCellValue.ToString());
                        break;
                    case CellType.Boolean:
                        cell.SetCellValue((Boolean)objCellValue);
                        break;
                    case CellType.Formula:
                        DateTime Datime = Convert.ToDateTime(objCellValue);
                        cell.SetCellValue(Datime);
                        //设置样式
                        ICellStyle cellStyle = workbook.CreateCellStyle();
                        IDataFormat format = workbook.CreateDataFormat();
                        cellStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
                        cell.CellStyle = cellStyle;
                        break;
                    case CellType.Numeric:
                        double dblValue = Convert.ToDouble(objCellValue);
                        cell.SetCellValue(dblValue);
                        break;
                }
                OK = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return OK;
        }

        /// <summary>
        /// 往Excel中添加图片
        /// </summary>
        /// <param name="strPicturePath">图片路径（PNG图片）</param>
        /// <param name="strColumnIndex">将图片添加至第几行（图片的位置，左对齐）</param>
        public bool InsertPicture(string strPicturePath, int intRowIndex)
        {
            bool OK = false;
            try
            {
                if (!File.Exists(strPicturePath))
                {
                    return OK;
                }
                //加载图片
                FileStream file = new FileStream(strPicturePath, FileMode.Open, FileAccess.Read);
                byte[] buffer = new byte[file.Length];
                file.Read(buffer, 0, (int)file.Length);
                int intPictureIndex = workbook.AddPicture(buffer, PictureType.PNG);
                switch (exceltype)
                {
                    case ExcelType.DEFAULT:
                        string strExtension = Path.GetExtension(FileName).ToLower();
                        if (strExtension == ".xls")
                        {
                            HSSFCreatePicture(intRowIndex, intPictureIndex);
                        }
                        else if (strExtension == ".xlsx")
                        {
                            XSSFCreatePicture(intRowIndex, intPictureIndex);
                        }
                        break;
                    case ExcelType.XLSX:
                        XSSFCreatePicture(intRowIndex, intPictureIndex);
                        break;
                    case ExcelType.XLS:
                        HSSFCreatePicture(intRowIndex, intPictureIndex);
                        break;
                }
                OK = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return OK;
        }
        /// <summary>
        /// XLS类型Excle文件添加图片
        /// </summary>
        /// <param name="intRowIndex">插入图片的行数</param>
        /// <param name="pictureIndex">图片的顺序</param>
        private void HSSFCreatePicture(int intRowIndex, int pictureIndex)
        {
            HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
            HSSFClientAnchor anchor = new HSSFClientAnchor();
            anchor.Row1 = intRowIndex;
            //anchor.AnchorType = 2;
            HSSFPicture picture = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIndex);
            picture.LineStyle = LineStyle.Solid;
            picture.Resize();//显示图片的原始尺寸
        }

        /// <summary>
        /// XLSX类型Excle文件添加图片
        /// </summary>
        /// <param name="intRowIndex">插入图片的行数</param>
        /// <param name="pictureIndex">图片的顺序</param>
        private void XSSFCreatePicture(int intRowIndex, int pictureIndex)
        {
            XSSFDrawing xssfDrawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFClientAnchor anchor = new XSSFClientAnchor();
            anchor.Row1 = intRowIndex;
            //anchor.AnchorType =  AnchorType2;
            XSSFPicture picture = (XSSFPicture)xssfDrawing.CreatePicture(anchor, pictureIndex);
            picture.LineStyle = LineStyle.Solid;
            picture.Resize();//显示图片的原始尺寸
        }

        #endregion
    }

    /// <summary>
    /// Excel文件操作方式
    /// </summary>
    public enum ExcelOperateMode
    {
        /// <summary>
        /// 打开已有文件
        /// </summary>
        open,
        /// <summary>
        /// 创建新文件
        /// </summary>
        create
    }

    /// <summary>
    /// Excel类型
    /// </summary>
    public enum ExcelType
    {
        /// <summary>
        /// 默认，不做指定
        /// </summary>
        DEFAULT,
        /// <summary>
        /// Excel工作薄
        /// </summary>
        XLSX,
        /// <summary>
        /// Excel 97-2003 工作薄
        /// </summary>
        XLS
    }

}
