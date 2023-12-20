using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.HSSF.UserModel; //建立 Excel 內容
using System.IO;
using System.Configuration;
using System.Threading;

namespace ExcelTableSchemaApp
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// 資料庫存取類別
        /// </summary>
        DBController m_db;
        /// <summary>
        /// 表格資訊串列
        /// </summary>
        List<TableInfo> m_lstTableInfo;

        /// <summary>
        /// 是否正在匯出
        /// </summary>
        bool m_bIsExporting = false;


        public Form1()
        {
            InitializeComponent();
        }

      

        #region 內部函式


        /// <summary>
        /// 初始化
        /// </summary>
        private void init()
        {
            txtConnString.Text = ConfigurationManager.ConnectionStrings["iscomString"].ConnectionString;
            m_lstTableInfo = new List<TableInfo>();
            txtOutputFile.Text = "OutPut"; 
        }

        /// <summary>
        /// 載入資料庫
        /// </summary>
        private void loadDB()
        {
            m_db = new DBController(txtConnString.Text);

            try
            {
                chklistTable.Items.Clear();

                //取清單Table
                BindTableList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            txtexpfilepath.Text = Application.StartupPath;
            if (File.Exists(Application.StartupPath + "\\sample.xls") == false)
            {
                MessageBox.Show("來源sample檔不存在，請先確認!");
                Close();
            }
        }

        /// <summary>
        /// 取得 Table 清單，並綁定至清單
        /// </summary>
        private void BindTableList()
        {

            DataTable dt = m_db.getAllTableSchema();

            foreach (DataRow row in dt.Rows)
            {
                chklistTable.Items.Add(row[0].ToString());
            }

        }

        /// <summary>
        /// 匯出EXCEL
        /// </summary>
        private void exportExcel()
        {            
            if (chklistTable.CheckedItems.Count == 0)
            {
                MessageBox.Show("請選擇Table名稱!");
                return;
            }
            if (string.IsNullOrEmpty(txtOutputFile.Text))
            {
                MessageBox.Show("請輸入匯出檔名!");
                return;
            }

            //  將 check list 被勾選的項目，加入至 table 串列中
            m_lstTableInfo.Clear();
            for (int i = 0; i < chklistTable.CheckedItems.Count; i++)
            {
                TableInfo info = new TableInfo
                {
                    Name = chklistTable.CheckedItems[i].ToString(), 
                    No = (i + 1)
                };

                m_lstTableInfo.Add(info);
            }   //  end for

            //  GUI 操作
            btnexport.Enabled = false;
            progressBar1.Value = 0; 
            progressBar1.Step = 1;
            progressBar1.Maximum = chklistTable.CheckedItems.Count;

            HSSFWorkbook workbook = null;
            FileStream file = null;
            
            string TableName = "";
            HSSFSheet sheet = null;

            try
            {

                //  載入 excel 檔
                workbook = new HSSFWorkbook(new FileStream(Application.StartupPath + "\\sample.xls", FileMode.Open));
                // 集中寫在同一檔
                file = new FileStream(txtexpfilepath.Text + "\\" + txtOutputFile.Text + ".xls", FileMode.Create);

                // 取得示範用的 sheet。   
                sheet = workbook.GetSheet("SheetSample999") as HSSFSheet;
            }
            catch (IOException exIO)
            {
                MessageBox.Show(string.Format("讀取檔案錯誤！，錯誤訊息：{0}", exIO.Message));
                btnexport.Enabled = true;
                return;
            }


            try
            {

                int i = 0;
                //  針對每一個勾選的項目，進行寫檔
                foreach (TableInfo item in m_lstTableInfo)
                {
                    TableName = item.Name;

                    //  讀取本 Table schema 資料
                    DataTable DT_TableAttr = m_db.getTableAttr(TableName);
                    //  copy 新的 sheet
                    sheet.CopyTo(workbook, TableName, true, true);
                    HSSFSheet sheetNew = workbook.GetSheet(TableName) as HSSFSheet;
                    //  複製失敗則不處理
                    if (sheetNew == null)
                        continue;
                    //  變更 sheet 名稱 (不需要了)
                    //workbook.SetSheetName( i + 1,TableName);


                    //寫Excel欄位值
                    ExcelExport.WriteExcelValue(m_db, workbook, sheetNew, DT_TableAttr, txtcreuser.Text, item);


                    progressBar1.PerformStep();
                    lblProgress.Text = string.Format("{0}/{1}", i + 1, m_lstTableInfo.Count);

                    Application.DoEvents();
                    i++;
                }   //  end for


                //  移除樣版 Sheet1
                workbook.RemoveSheetAt(1);
                //  寫入第一頁 TableList
                ExcelExport.WriteTableListSheet(m_lstTableInfo, workbook);

                //寫入檔案    
                workbook.Write(file);
                MessageBox.Show("檔案產生完成!");
            }
            catch (IOException exIO)
            {
                MessageBox.Show(string.Format("寫入檔案錯誤！，錯誤訊息：{0}", exIO.Message));
            }
            catch (Exception ex)
            {
                //  非 IO 錯誤，則立刻寫入檔案
                workbook.Write(file);
                MessageBox.Show(string.Format("輸出 Excel 錯誤！資料表：{0} ，錯誤訊息：{1}", TableName, ex.Message));
            }
            finally
            {
                //  最後再 close
                file.Close();
                workbook = null;

                btnexport.Enabled = true;
            }
        }


        
        #endregion

   
    



        #region 事件處理常式


        private void Form1_Load(object sender, EventArgs e)
        {

            init();


        }


        //全選
        private void linkchkall_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            for (int i = 0; i < chklistTable.Items.Count; i++)
            {
                chklistTable.SetItemChecked(i, true);
            }
        }

        //全不選
        private void linkunchkall_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            for (int i = 0; i < chklistTable.Items.Count; i++)
            {
                chklistTable.SetItemChecked(i, false);
            }
        }
        //  檔案對話盒
        private void btnBrowser_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtexpfilepath.Text = folderBrowserDialog1.SelectedPath;
            }

        }

        /// <summary>
        ///  匯出EXCEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnexport_Click(object sender, EventArgs e)
        {
            m_bIsExporting = true;
            exportExcel();
            m_bIsExporting = false;
 
        }

        //  載入資料庫
        private void btnLoadDB_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtConnString.Text))
            {
                MessageBox.Show("請輸入連線字串!");
                return;
            }


            loadDB();
        }

        //  視窗關閉
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (m_bIsExporting)
            {
                MessageBox.Show("匯出中，請稍候再關閉!");
            }
        }

        #endregion

     

        

    }
}
