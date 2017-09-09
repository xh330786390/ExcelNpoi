using System;
using System.Collections.Generic;

using NPOI.SS.UserModel;
using System.Data;
using System.IO;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;

namespace ExcelNpoi
{
    //*****************************************
    /// <summary>
    /// NPOI操作Excel 2007类
    /// @author:tengxiaohui
    /// time:2017-08-03
    /// </summary>
    //*****************************************
    public class ExcelNpoi2007 : IExcelNpoi
    {
        private const string YYYY_MM_DD = @"\d{4}-\d{2}-\d{2}";
        private const string MM_DD = @"([0-1][0-9])-([0-3][0-9])";
        private const string YYYY年 = @"\d{4}年";
        private const string MM月 = @"[1-9]月|[0-1][0-9]月";

        #region 字段声明
        /// <summary>
        /// 日期列宽
        /// </summary>
        private const int DATEWIDTH = 15;
        /// <summary>
        /// 数据列宽
        /// </summary>
        private const int DATAWIDTH = 25;

        /// <summary>
        /// 工作表
        /// </summary>
        private XSSFWorkbook _workBook = null;

        /// <summary>
        /// 数据头
        /// </summary>
        private string header = null;

        /// <summary>
        /// 批注
        /// </summary>
        private string commnet = null;

        /// <summary>
        /// 数据头行数
        /// </summary>
        private int headRows = 0;
        #endregion

        #region 单元格样式
        /// <summary>
        /// 标题样式
        /// </summary>
        private ICellStyle _titleStyle = null;
        private ICellStyle TitleStyle
        {
            get
            {
                if (_titleStyle == null)
                {
                    _titleStyle = _workBook.CreateCellStyle();
                    _titleStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;

                    NPOI.SS.UserModel.IFont _fontTitle = _workBook.CreateFont();
                    _fontTitle.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                    _fontTitle.Color = NPOI.HSSF.Util.HSSFColor.Red.Index;
                    _titleStyle.SetFont(_fontTitle);
                }
                return _titleStyle;
            }
        }

        /// <summary>
        /// 数据头样式
        /// </summary>
        private ICellStyle _headStyle = null;
        private NPOI.SS.UserModel.ICellStyle HeadStyle
        {
            get
            {
                if (_headStyle == null)
                {
                    _headStyle = _workBook.CreateCellStyle();
                    _headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;

                    NPOI.SS.UserModel.IFont _fontHead = _workBook.CreateFont();
                    _fontHead.Color = NPOI.HSSF.Util.HSSFColor.Black.Index;
                    _headStyle.SetFont(_fontHead);
                }
                return _headStyle;
            }
        }

        /// <summary>
        ///日期样式
        /// </summary>
        private ICellStyle _dateStyle_yyyy_MM_dd = null;
        private ICellStyle DateStyle_yyyy_MM_dd
        {
            get
            {
                if (_dateStyle_yyyy_MM_dd == null)
                {
                    _dateStyle_yyyy_MM_dd = _workBook.CreateCellStyle();
                    _dateStyle_yyyy_MM_dd.WrapText = true;
                    _dateStyle_yyyy_MM_dd.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                    _dateStyle_yyyy_MM_dd.DataFormat = _workBook.CreateDataFormat().GetFormat("yyyy-MM-dd");

                    NPOI.SS.UserModel.IFont _fontDate = _workBook.CreateFont();
                    _fontDate.Color = NPOI.HSSF.Util.HSSFColor.Black.Index;
                    _dateStyle_yyyy_MM_dd.SetFont(_fontDate);
                }
                return _dateStyle_yyyy_MM_dd;
            }
        }

        /// <summary>
        ///数据样式
        /// </summary>
        private ICellStyle _dataStyle = null;
        private ICellStyle DataStyle
        {
            get
            {
                if (_dataStyle == null)
                {
                    _dataStyle = _workBook.CreateCellStyle();
                    _dataStyle.WrapText = true;
                    _dataStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;

                    NPOI.SS.UserModel.IFont _fontData = _workBook.CreateFont();
                    _fontData.Color = NPOI.HSSF.Util.HSSFColor.Black.Index;
                    _dataStyle.SetFont(_fontData);
                }
                return _dataStyle;
            }
        }
        #endregion

        #region [公共方法]
        /// <summary>
        /// Npoi读取文件
        /// </summary>
        /// <param name="fileName">文件</param>
        /// <returns>数据表</returns>
        public DataTable ReadFileFromExcel(string file)
        {
            DataTable dt = new DataTable();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                _workBook = new XSSFWorkbook(fs);
                ISheet sheet = _workBook.GetSheetAt(0);

                //@1.获取表头
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueTypeForXls(header.GetCell(i) as XSSFCell);
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                    }
                    else
                    {
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    }
                    columns.Add(i);
                }

                //@1.获取数据
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueTypeForXls(sheet.GetRow(i).GetCell(j) as XSSFCell);
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                            break;
                        }
                    }

                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// 导出数据
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="fileName">导出文件名</param>
        /// <param name="objs"></param>
        public void ExportToExcel(DataTable dt, string file, params object[] objs)
        {

        }

        /// <summary>
        /// 导出数据
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="fileName">导出文件名</param>
        /// <param name="header">数据头</param>
        /// <param name="commnet">批注内容</param>
        public void ExportToExcel(DataTable dt, string file, string header = null, string commnet = null, int headRows = 0)
        {
            this.header = header;
            this.commnet = commnet;
            this.headRows = headRows;

            try
            {
                byte[] bytes = Export(dt);

                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(bytes, 0, bytes.Length);
                    fs.Flush();
                }
            }
            catch (Exception er) { Console.WriteLine(er.ToString()); }
        }
        #endregion

        #region [私有方法]
        /// <summary>
        /// DataTable 数据至二进制
        /// </summary>
        /// <param name="dt">数据表</param>
        /// <param name="header"></param>
        /// <param name="commnet"></param>
        /// <returns>二进制数组</returns>
        public byte[] Export(DataTable dt)
        {
            _workBook = new XSSFWorkbook();
            ISheet sheet = _workBook.CreateSheet();
            if (header == "钢联数据")
            {
                _workBook.SetSheetName(0, "Sheet1");
            }
            else
            {
                _workBook.SetSheetName(0, "数据");
            }

            //设置列宽
            setColumnWidth(dt, sheet);

            int rowIndex = 0;

            foreach (DataRow dr in dt.Rows)
            {
                //新建表，填充表头，填充列头，样式
                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = _workBook.CreateSheet();
                    }

                    if (!string.IsNullOrEmpty(header))
                    {
                        IRow titleRow = sheet.CreateRow(0);

                        createTitleComment(sheet, titleRow);
                        rowIndex = 1;
                    }
                }

                //设置单元格数据
                setCellsValue(sheet, dt, dr, rowIndex);
                rowIndex++;
            }

            using (MemoryStream ms = new MemoryStream())
            {
                _workBook.Write(ms);
                ms.Flush();
                _workBook.Clear();
                _workBook = null;
                return ms.ToArray();
            }
        }

        /// <summary>
        /// 设置单元值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dt"></param>
        /// <param name="dr"></param>
        /// <param name="rowIndex"></param>
        private void setCellsValue(ISheet sheet, DataTable dt, DataRow dr, int rowIndex)
        {
            IRow dataRow = sheet.CreateRow(rowIndex);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell newCell = dataRow.CreateCell(dt.Columns[i].Ordinal);
                string drValue = dr[i].ToString();

                DateTime dtTimeOut;
                double doubleTemp;
                if (drValue != string.Empty && Double.TryParse(drValue.Replace(",", string.Empty), out doubleTemp))
                {
                    newCell.SetCellValue(doubleTemp);
                    newCell.CellStyle = DataStyle;
                }
                else if (drValue != string.Empty && DateTime.TryParse(drValue, out dtTimeOut))
                {
                    if (Regex.IsMatch(drValue, YYYY_MM_DD))
                    {
                        newCell.SetCellValue(dtTimeOut.Date);
                        newCell.CellStyle = DateStyle_yyyy_MM_dd;
                    }
                    else
                    {
                        newCell.SetCellValue(drValue);
                    }
                }
                else
                {
                    if (rowIndex <= headRows)
                    {
                        newCell.CellStyle = HeadStyle;
                    }
                    newCell.SetCellValue(drValue);
                }
            }
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="sheet"></param>
        private void setColumnWidth(DataTable dt, ISheet sheet)
        {
            foreach (DataColumn column in dt.Columns)
            {
                if (column.Ordinal == 0)
                {
                    sheet.SetColumnWidth(column.Ordinal, DATEWIDTH * 256);
                }

                else
                {
                    sheet.SetColumnWidth(column.Ordinal, DATAWIDTH * 256);
                }
            }
        }

        /// <summary>
        /// 创建标题批注
        /// </summary>
        /// <param name="sheet">表单</param>
        /// <param name="titleRow">行</param>
        /// <param name="header"></param>
        /// <param name="commnet"></param>
        private void createTitleComment(ISheet sheet, IRow titleRow)
        {
            ICell newCellHeader = titleRow.CreateCell(0);

            newCellHeader.SetCellValue(header);
            newCellHeader.CellStyle = TitleStyle;

            if (string.IsNullOrEmpty(this.header) || string.IsNullOrEmpty(this.commnet)) return;

            XSSFDrawing xdraw = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFComment commnet = (XSSFComment)xdraw.CreateCellComment(new XSSFClientAnchor(0, 0, 0, 0, 1, 2, 4, 4));
            commnet.SetString(this.commnet);
            newCellHeader.CellComment = commnet;
        }

        /// <summary>  
        /// 获取单元格类型(xls)  
        /// </summary>  
        /// <param name="cell"></param>  
        /// <returns></returns>  
        private object GetValueTypeForXls(XSSFCell cell)
        {
            if (cell == null) return null;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return null;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Error:
                    return cell.ErrorCellValue;
                case CellType.Formula:
                default:
                    return "=" + cell.CellFormula;
            }
        }
        #endregion



        static DataTable RenderFromExcel(Stream excelFileStream)
        {
            using (excelFileStream)
            {
                IWorkbook workbook = new XSSFWorkbook(excelFileStream);
                {
                    ISheet sheet = workbook.GetSheetAt(0);
                    {
                        DataTable table = new DataTable();

                        IRow headerRow = sheet.GetRow(0);//第一行为标题行  
                        int cellCount = headerRow.LastCellNum;//LastCellNum = PhysicalNumberOfCells  
                        int rowCount = sheet.LastRowNum;//LastRowNum = PhysicalNumberOfRows - 1  

                        //handling header.  
                        for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                        {
                            DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                            table.Columns.Add(column);
                        }

                        for (int i = (sheet.FirstRowNum + 1); i <= rowCount; i++)
                        {
                            IRow row = sheet.GetRow(i);
                            DataRow dataRow = table.NewRow();

                            if (row != null)
                            {
                                for (int j = row.FirstCellNum; j < cellCount; j++)
                                {
                                    //if (row.GetCell(j) != null)
                                    //    dataRow[j] = GetCellValue(row.GetCell(j));
                                }
                            }

                            table.Rows.Add(dataRow);
                        }
                        return table;

                    }
                }
            }
        }


        //public static bool HasData(Stream excelFileStream)
        //{
        //    using (excelFileStream)
        //    {
        //        using (IWorkbook workbook = new HSSFWorkbook(excelFileStream))
        //        {
        //            if (workbook.NumberOfSheets > 0)
        //            {
        //                using (ISheet sheet = workbook.GetSheetAt(0))
        //                {
        //                    return sheet.PhysicalNumberOfRows > 0;
        //                }
        //            }
        //        }
        //    }
        //    return false;
        //}  

    }



}
