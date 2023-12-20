using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data;
using NPOI.HSSF.Util;

namespace ExcelTableSchemaApp
{
    /// <summary>
    /// NPOI 輸出 Excel 的函式
    /// 2014.05.22 Jason
    /// </summary>
    static public class ExcelExport
    {

     

        /// <summary>
        /// 寫入 Excel 欄位值
        /// </summary>
        /// <param name="db">資料庫存取物件參考</param>
        /// <param name="workbook">Excel</param>
        /// <param name="sheet_new">Excel 工作表</param>
        /// <param name="DT_TableAttr">欲填入 Excel 的schema資訊</param>
        /// <param name="Author">填表人</param>
        /// <param name="info">表格資訊</param>
        static public void WriteExcelValue(DBController db, HSSFWorkbook workbook, HSSFSheet sheet_new, DataTable DT_TableAttr, string Author, TableInfo info)
        {            
            string TableName = info.Name;

            //寫Table Name
            HSSFRow erow = sheet_new.GetRow(1) as HSSFRow;
            erow.GetCell(1).SetCellValue(TableName);


            //資料表名稱(中文)    2014.5.20 Jason
            erow = sheet_new.GetRow(0) as HSSFRow;
            string desc = db.getTableDesc(TableName);
            erow.GetCell(2).SetCellValue(desc);
            info.Desc = desc;

            //寫填表人姓名
            erow = sheet_new.GetRow(0) as HSSFRow;
            erow.GetCell(4).SetCellValue(Author);

            #region 回首頁連結  2014.6.26
           

            erow = sheet_new.GetRow(0) as HSSFRow;
            HSSFCell cell = erow.GetCell(7) as HSSFCell;

            //  設定超連結 style
            setLinkStyle(workbook, cell);
                     
            //  設定超連結
            HSSFHyperlink link = new HSSFHyperlink(HyperlinkType.Document);
            link.Address = string.Format("'{0}'!A1", "TableList");
            cell.Hyperlink = (link);
            cell.SetCellValue("回首頁");

            #endregion


            //寫PK值
            erow = sheet_new.GetRow(2) as HSSFRow;
            erow.GetCell(1).SetCellValue(db.getPKValue(TableName));

            //寫FK值
            erow = sheet_new.GetRow(3) as HSSFRow;
            erow.GetCell(1).SetCellValue(db.getFKValue(TableName));

            int i = 5;
            foreach (DataRow row in DT_TableAttr.Rows)
            {
                //   判斷是否 sample 已建立資料列
                erow = sheet_new.GetRow(i) as HSSFRow;
                if (erow != null)
                {
                    //   已建立，則直接取得參考並設值

                    erow.GetCell(0).SetCellValue(row["COLUMN_NAME"].ToString());
                    erow.GetCell(1).SetCellValue("");
                    if (row["Description"] != DBNull.Value)
                    {
                        erow.GetCell(1).SetCellValue(row["Description"].ToString());
                    }


                    //  2014.08.06 Jason 改為另外處理
                    erow.GetCell(2).SetCellValue(processDataTypeString(row));

                    //   2014.05.20 Jason 填入整數
                    if (row["CHARACTER_MAXIMUM_LENGTH"] != DBNull.Value)
                    {
                        int n = 0;
                        if (Int32.TryParse(row["CHARACTER_MAXIMUM_LENGTH"].ToString(), out n))
                            erow.GetCell(3).SetCellValue(n);
                    }

                    if (row["IS_NULLABLE"].ToString() == "YES")
                    {
                        erow.GetCell(4).SetCellValue("V");
                    }

                    erow.GetCell(5).SetCellValue("");
                    //if (row["Description"] != DBNull.Value)
                    //{
                    //    erow.GetCell(5).SetCellValue(row["Description"].ToString());
                    //}
                }
                else
                {
                    //   尚未建立，則建立新列，並繪製邊線

                    erow = sheet_new.CreateRow(i) as HSSFRow;
                    erow.CreateCell(0).SetCellValue(row["COLUMN_NAME"].ToString());
                    FillBorder(erow.GetCell(0));

                    erow.CreateCell(1).SetCellValue("");
                    FillBorder(erow.GetCell(1));
                    if (row["Description"] != DBNull.Value)
                    {
                        erow.GetCell(1).SetCellValue(row["Description"].ToString());
                    }

                    erow.CreateCell(2).SetCellValue(row["DATA_TYPE"].ToString());
                    FillBorder(erow.GetCell(2));


                    erow.CreateCell(3);

                    FillBorder(erow.GetCell(3));

                    if (row["CHARACTER_MAXIMUM_LENGTH"] != DBNull.Value)
                    {
                        int n = 0;
                        if (Int32.TryParse(row["CHARACTER_MAXIMUM_LENGTH"].ToString(), out n))
                            erow.GetCell(3).SetCellValue(n);
                    }

                    erow.CreateCell(4).SetCellValue("");
                    FillBorder(erow.GetCell(4));
                    if (row["IS_NULLABLE"].ToString() == "YES")
                    {
                        erow.GetCell(4).SetCellValue("V");
                    }

                    erow.CreateCell(5).SetCellValue("");
                    FillBorder(erow.GetCell(5));
                    //if (row["Description"] != DBNull.Value)
                    //{
                    //    erow.GetCell(5).SetCellValue(row["Description"].ToString());
                    //}

                }


                i++;
            }

        }

        /// <summary>
        /// 處理 Decimal/Numeric 這類 type 的 NUMERIC_PRECISION / NUMERIC_SCALE 顯示
        /// 2014.08.06 Jason
        /// </summary>
        /// <param name="row">資料列</param>
        /// <returns>顯示字串</returns>
        static public string processDataTypeString(DataRow row)
        {
            //  處理 Decimal/Numeric 這類 type 的 NUMERIC_PRECISION / NUMERIC_SCALE 顯示

            if (row["DATA_TYPE"] == DBNull.Value)
                return "Data Type Wrong!";
                     

            string sDataType = row["DATA_TYPE"].ToString();

            //  只處理 Numeric, Decimal
            if ("NUMERIC" != sDataType.ToUpper() && "DECIMAL" != sDataType.ToUpper())
                return sDataType;
            

            if (row["NUMERIC_PRECISION"] != DBNull.Value && row["NUMERIC_SCALE"] != DBNull.Value)
            {
                sDataType = string.Format("{0}({1},{2})", sDataType, row["NUMERIC_PRECISION"], row["NUMERIC_SCALE"]);
            }

            return sDataType;
        }

        /// <summary>
        ///   寫入Table List 工作表
        ///   2014.05.22 Jason
        /// </summary>
        /// <param name="lstTableInfo">Table資訊串列</param>
        /// <param name="workbook">整份工作表物件</param>
        static public void WriteTableListSheet(List<TableInfo> lstTableInfo, HSSFWorkbook workbook)
        {
            //  取得 Table List 工作表
            HSSFSheet sheet = workbook.GetSheet("TableList") as HSSFSheet;


                     

            //  寫入各表格名稱和說明
            int index = 0;
            foreach (TableInfo item in lstTableInfo)
            {

                HSSFRow erow = sheet.GetRow(index+1) as HSSFRow;
                if (erow != null)
                {
                    //  項次
                    erow.GetCell(0).SetCellValue(item.No);
                    //  資料表名稱 (加上文內連結)
                    //  設定超連結 style
                    HSSFCell cell = erow.GetCell(1) as HSSFCell;
                    setLinkStyle(workbook, cell);


                    HSSFHyperlink link = new HSSFHyperlink(HyperlinkType.Document);
                    link.Address = string.Format("'{0}'!A1", item.Name);
                    cell.Hyperlink = (link);
                    cell.SetCellValue(item.Name);
                    FillBorder(cell);

                    //  資料表說明
                    erow.GetCell(2).SetCellValue(item.Desc);
                }
                else
                {
                    erow = sheet.CreateRow(index+1) as HSSFRow;
                    //  項次
                    erow.CreateCell(0).SetCellValue(item.No);
                    //  資料表名稱 (加上文內連結)
                    HSSFCell cell = erow.CreateCell(1) as HSSFCell;
                    setLinkStyle(workbook, cell);
                    HSSFHyperlink link = new HSSFHyperlink(HyperlinkType.Document);
                    link.Address = string.Format("'{0}'!A1", item.Name);
                    cell.Hyperlink = (link);
                    cell.SetCellValue(item.Name);
                    FillBorder(cell);
                }
                index++;
            }   //  end foreach



        }



        /// <summary>
        /// 加上 Cell 框線
        /// </summary>
        /// <param name="ecell"></param>
        static public void FillBorder(ICell ecell)
        {
            ecell.CellStyle.BorderBottom = BorderStyle.Thin;
            ecell.CellStyle.BorderLeft = BorderStyle.Thin;
            ecell.CellStyle.BorderRight = BorderStyle.Thin;
            ecell.CellStyle.BorderTop = BorderStyle.Thin;
        }

        /// <summary>
        /// 設定超連結 style
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cell"></param>
        static public void setLinkStyle(HSSFWorkbook workbook, HSSFCell cell)
        {
            //  設定 style
            HSSFCellStyle hlink_style = workbook.CreateCellStyle() as HSSFCellStyle;
            HSSFFont hlink_font = workbook.CreateFont() as HSSFFont;
            hlink_font.Underline = FontUnderlineType.Single;
            hlink_font.Color = HSSFColor.Blue.Index;
            hlink_style.SetFont(hlink_font);

            cell.CellStyle = (hlink_style);
        }


    }

    /// <summary>
    /// 資料表資訊類別
    /// 2014.05.22  Jason
    /// </summary>
    public class TableInfo
    {
        /// <summary>
        /// 項次
        /// </summary>
        public int No { get; set; }
        /// <summary>
        /// 資料表名稱
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 資料表說明
        /// </summary>
        public string Desc { get; set; }
    }

}
