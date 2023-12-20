/*
 *  Author          :   Fisher   
 *  Function        :   多資料庫管理
 *  (2013.10.14，複製至Component-Cloud專案，用於 Message Sender 的修改版本)
*/
using System;
using System.Data;
using MySql.Data.Types;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.Common;
using System.Collections;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;
using System.Configuration;

namespace DBManage
{
    #region 參數相關
    // 參數結構
    public class DBParam
    {
        public String m_strKey;
        public Object m_Value;
        public DbType m_dbType;
        public ParameterDirection m_Direction;

        public DBParam()
        {
            this.m_strKey = null;
            this.m_Value = null;
            this.m_dbType = DbType.Object;
            this.m_Direction = ParameterDirection.Input;
        }

        // 參數結?
        public DBParam(String Key, DbType dbType, object Value)
        {
            this.m_strKey = Key;
            this.m_Value = Value;
            this.m_dbType = dbType;
            this.m_Direction = ParameterDirection.Input;
        }
    }

    // 參數操作
    public class BatchProc
    {
        // 存储过程名称
        private String m_strProcName = "";

        // 存储过程参数数组
        private DBParam[] m_ProcParam = null;

        // 默认构造函数
        public BatchProc(String strProcName, DBParam[] ProcParam)
        {
            if (null != ProcParam)
            {
                m_strProcName = strProcName;

                m_ProcParam = ProcParam;
            }
        }

        // 屬性：參數名稱
        public String ProcName
        {
            get
            {
                return m_strProcName;
            }
            set
            {
                m_strProcName = value;
            }
        }

        // 屬性：參數?列
        public DBParam[] ProcParam
        {
            get
            {
                return m_ProcParam;
            }
            set
            {
                m_ProcParam = value;
            }
        }

        // 傳回參數?列
        public DBParam[] GetProcParam()
        {
            return m_ProcParam;
        }

        // 傳回參數名稱
        public String GetProcName()
        {
            return m_strProcName;
        }
    }

    #endregion

    // 資料庫操作?別
    public class DBManage
    {
        #region 建構式

        /// <summary>
        /// 建構式，傳入config 的 key，取得連線字串。
        /// 2013.10.14 Jason
        /// </summary>
        /// <param name="ConnStringKey">key值</param>
        public DBManage(string ConnStringKey)
        {
           m_strConnection =  ConfigurationManager.ConnectionStrings[ConnStringKey].ConnectionString;
        }

        #endregion

        #region -----------------------公用函式------------------------


        /// <summary>
        /// 執行SQL敘述，傳回是否成功
        /// </summary>
        /// <param name="strSQL"></param>
        /// <returns></returns>
        public bool ExecuteSQL(String strSQL)
        {
            if (InitDataBase(m_strConnection, m_strDBType))
            {
                ExecuteOracleFormat();

                //IDbTransaction dbTransaction = null;

                try
                {
                    // 交易開始
                    //dbTransaction = m_DBConnection.BeginTransaction();

                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.CommandType = CommandType.Text;
                    //m_DBCommand.Transaction = dbTransaction;
                    m_DBCommand.CommandText = SQLProcess(strSQL);

                    // 執行SQL
                    m_DBCommand.ExecuteNonQuery();

                    // 交易 commit
                    //dbTransaction.Commit();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute SQL Successful";
                }
                catch (Exception e)
                {
                    //if (dbTransaction != null)
                    //{
                    //    // 交易撤回
                    //    dbTransaction.Rollback();
                    //}

                    m_bIsSuccess = false;
                    m_strOpeInfo = e.Message;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }

        
        /// <summary>
        /// 執行多SQL敘述，傳回是否成功
        /// </summary>
        /// <param name="strSQL"></param>
        /// <returns></returns>
        public bool ExecuteSQL(String[] strSQL)
        {
            if (InitDataBase(m_strConnection, m_strDBType))
            {
                ExecuteOracleFormat();

                IDbTransaction dbTransaction = null;

                try
                {
                    // 交易開始
                    dbTransaction = m_DBConnection.BeginTransaction();

                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.Transaction = dbTransaction;

                    // 開始剖析 SQL 敘述
                    foreach (String str in strSQL)
                    {
                        if (str.Trim() != "")
                        {
                            m_DBCommand.CommandText = SQLProcess(str);
                            m_DBCommand.CommandType = CommandType.Text;

                            m_DBCommand.ExecuteNonQuery();
                        }
                    }

                    // commit 交易
                    dbTransaction.Commit();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute SQL Array Successful";

                }
                catch (Exception e)
                {
                    if (dbTransaction != null)
                    {
                        // 交易撤回
                        dbTransaction.Rollback();
                    }

                    m_bIsSuccess = false;
                    m_strOpeInfo = e.Message;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }

        /// <summary>
        /// 執行SQL?列, 返回结果到DS(累加字段)
        /// </summary>
        /// <param name="strSQL"></param>
        /// <param name="dataset"></param>
        /// <returns></returns>
        public bool ExecuteSQL(String[] strSQL, ref DataSet dataset)
        {
            if (InitDataBase(m_strConnection, m_strDBType))
            {
                if (null == dataset)
                {
                    dataset = new DataSet();
                }

                dataset.Clear();

                ExecuteOracleFormat();

                m_DBCommand = m_DBConnection.CreateCommand();

                try
                {
                    Int32 iIdxName = 0;
                    foreach (String str in strSQL)
                    {
                        DataSet tempDS = new DataSet();

                        m_DBCommand.CommandText = SQLProcess(str);

                        m_DBDataAdapter.SelectCommand = m_DBCommand;

                        m_DBDataAdapter.Fill(dataset);

                        dataset.Tables[iIdxName].TableName = "table" + iIdxName.ToString();

                        ++iIdxName;
                    }

                    m_bIsSuccess = true;

                    m_strOpeInfo = "Execute SQL Array Into DataSet Correct";
                }
                catch (Exception e)
                {
                    dataset = null;

                    m_bIsSuccess = false;

                    m_strOpeInfo = e.Message;
                }
                finally
                {
                    Close();
                }
            }
            else
            {
                dataset = null;
            }

            return m_bIsSuccess;
        }

      

        /// <summary>
        ///  執行SQL?列, 傳回 DataSet ,bool 傳回 true, 則ds不清空
        /// </summary>
        /// <param name="strSQL"></param>
        /// <param name="dataset"></param>
        /// <param name="bRestore"></param>
        /// <returns></returns>
        public bool ExecuteSQL(String[] strSQL, ref DataSet dataset, bool bRestore)
        {
            if (InitDataBase(m_strConnection, m_strDBType))
            {
                if (null == dataset)
                {
                    dataset = new DataSet();
                }

                if (!bRestore)
                {
                    dataset.Clear();
                }

                ExecuteOracleFormat();

                m_DBCommand = m_DBConnection.CreateCommand();

                try
                {
                    Int32 iIdxName = dataset.Tables.Count;
                    foreach (String str in strSQL)
                    {
                        DataSet tempDS = new DataSet();

                        m_DBCommand.CommandText = SQLProcess(str);

                        m_DBDataAdapter.SelectCommand = m_DBCommand;

                        m_DBDataAdapter.Fill(dataset);

                        dataset.Tables[iIdxName].TableName = "table" + iIdxName.ToString();

                        ++iIdxName;
                    }

                    m_bIsSuccess = true;

                    m_strOpeInfo = "Execute SQL Array Into DataSet Correct";
                }
                catch (Exception e)
                {
                    dataset = null;

                    m_bIsSuccess = false;

                    m_strOpeInfo = e.Message;
                }
                finally
                {
                    Close();
                }
            }
            else
            {
                dataset = null;
            }

            return m_bIsSuccess;
        }

        /// <summary>
        /// 執行SQL，取得字串值
        /// </summary>
        /// <param name="strSQL"></param>
        /// <returns></returns>
        public String ExecuteForValue(String strSQL)
        {
            String strValue = "-1";

            if (InitDataBase(m_strConnection, m_strDBType))
            {
                ExecuteOracleFormat();

                try
                {
                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.CommandText = SQLProcess(strSQL);

                    strValue = m_DBCommand.ExecuteScalar().ToString();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute For Value Success";
                }
                catch (Exception e)
                {
                    m_strOpeInfo = e.Message;
                    m_bIsSuccess = false;
                    strValue = "-1";
                }
            }

            return strValue;
        }

        /// <summary>
        /// 批次執行SQL串列，傳回執行成功的資料筆數
        /// </summary>
        /// <param name="strArrayList"></param>
        /// <param name="nMaxCount"></param>
        /// <returns></returns>
        public int ExecuteSQLArray(ArrayList strArrayList, Int64 nMaxCount)
        {
            // 成功筆數
            int iSuccessCount = 0;

            if (null != strArrayList)
            {
                foreach (String[] strArray in strArrayList)
                {
                    if (ExecuteSQL(strArray))
                    {
                        iSuccessCount++;
                    }

                    if (iSuccessCount >= nMaxCount)
                    {
                        break;
                    }
                }
            }

            return iSuccessCount;
        }


        /// <summary>
        /// 批次執行SQL串列，傳回執行成功的資料筆數
        /// </summary>
        /// <param name="strArrayList"></param>
        /// <returns></returns>
        public int ExecuteSQLArray(ArrayList strArrayList)
        {
            // 成功筆數
            int iSuccessCount = 0;

            if (null != strArrayList)
            {
                foreach (String[] strArray in strArrayList)
                {
                    if (ExecuteSQL(strArray))
                    {
                        iSuccessCount++;
                    }
                }
            }

            return iSuccessCount;
        }


        /// <summary>
        /// 比次執行SQL?列,返回成功筆數(Add Log)
        /// </summary>
        /// <param name="strArrayList"></param>
        /// <param name="nMaxCount"></param>
        /// <param name="strMailList"></param>
        /// <param name="strNameList"></param>
        /// <param name="DT_ImportLogTable"></param>
        /// <param name="iInsertSeq"></param>
        /// <param name="successIndex"></param>
        /// <returns></returns>
        public int ExecuteSQLArrayForLog(ArrayList strArrayList, Int64 nMaxCount, ArrayList strMailList, ArrayList strNameList, ref DataTable DT_ImportLogTable, ref int iInsertSeq, ArrayList successIndex)
        {
            String strSuccess = XMLManage.GetValue("//Member//DataTrans//LogError14");
            String strFailure = XMLManage.GetValue("//Member//DataTrans//LogError15");


            // 成功比數
            int iSuccessCount = 0;
            int iSeqNo = 0;
            if (null != strArrayList)
            {
                foreach (String[] strArray in strArrayList)
                {

                    //foreach (String str in strArray)
                    //{
                    //    if (str.Trim() != "")
                    //    {
                    //        DataRow dr1 = DT_ImportLogTable.NewRow();
                    //        dr1["Email"] = strMailList[iSeqNo];
                    //        dr1["Name"] = strNameList[iSeqNo];
                    //        dr1["ImportTime"] = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                    //        dr1["Status"] = "TEST";
                    //        dr1["Error"] = str.Trim();
                    //        DT_ImportLogTable.Rows.Add(dr1);
                    //    }
                    //}

                    iInsertSeq += 1;
                    if (nMaxCount != -1 && iSuccessCount > nMaxCount)
                    {
                        DataRow dr = DT_ImportLogTable.NewRow();
                        dr["Email"] = strMailList[iSeqNo];
                        dr["Name"] = strNameList[iSeqNo];
                        dr["ImportTime"] = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                        dr["Status"] = strFailure;
                        dr["Error"] = XMLManage.GetValue("//Member//DataTrans//LogError8") + "; ";
                        DT_ImportLogTable.Rows.Add(dr);
                    }
                    else
                    {
                        if (ExecuteSQL(strArray))
                        {
                            iSuccessCount++;

                            DataRow dr = DT_ImportLogTable.NewRow();
                            dr["Email"] = strMailList[(int)successIndex[iSeqNo]];
                            dr["Name"] = strNameList[(int)successIndex[iSeqNo]];
                            dr["ImportTime"] = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                            dr["Status"] = strSuccess;
                            dr["Error"] = "";
                            DT_ImportLogTable.Rows.Add(dr);
                        }
                        else
                        {
                            DataRow dr = DT_ImportLogTable.NewRow();
                            dr["Email"] = strMailList[(int)successIndex[iSeqNo]];
                            dr["Name"] = strNameList[(int)successIndex[iSeqNo]];
                            dr["ImportTime"] = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                            dr["Status"] = strFailure;
                            dr["Error"] = XMLManage.GetValue("//Member//DataTrans//LogError9") + ": "
                            + m_strOpeInfo + "; ";
                            DT_ImportLogTable.Rows.Add(dr);
                        }
                    }

                    iSeqNo += 1;
                }
            }

            return iSuccessCount;
        }


        /// <summary>
        /// 執行SQL，接受參數，傳回是否成功
        /// </summary>
        /// <param name="strSQL"></param>
        /// <param name="prams"></param>
        /// <returns></returns>
        public bool ExecuteSQL(String strSQL, DBParam[] prams)
        {
            if (InitDataBase(m_strConnection, m_strDBType))
            {
                ExecuteOracleFormat();

                IDbTransaction dbTransaction = null;

                try
                {
                    // 开始事务
                    dbTransaction = m_DBConnection.BeginTransaction();

                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.CommandType = CommandType.Text;
                    m_DBCommand.Transaction = dbTransaction;
                    m_DBCommand.CommandText = SQLProcess(strSQL);

                    CreateNewInParameter(m_strDBType, prams);

                    m_DBCommand.ExecuteNonQuery();

                    // 事务提交
                    dbTransaction.Commit();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute SQL Successful";
                }
                catch (Exception e)
                {
                    if (dbTransaction != null)
                    {
                        // 事务回滚
                        dbTransaction.Rollback();
                    }

                    m_strOpeInfo = e.Message;
                    m_bIsSuccess = false;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }

        // 执行SQL，返回是否成功，同时返回SQL操作结果
        public bool ExecuteSQL(String strSQL, ref Object sqlReVal)
        {
            if (InitDataBase(m_strConnection, m_strDBType))
            {
                ExecuteOracleFormat();

                m_DBCommand = m_DBConnection.CreateCommand();
                m_DBCommand.CommandType = CommandType.Text;
                m_DBCommand.CommandText = SQLProcess(strSQL);

                try
                {
                    sqlReVal = m_DBCommand.ExecuteScalar();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute SQL Successful";
                }
                catch (Exception e)
                {
                    m_bIsSuccess = false;
                    m_strOpeInfo = e.Message;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }

        // 执行SQL，返回是否成功，同时返回DATASET
        public bool ExecuteSQL(String strSQL, ref DataSet dataset)
        {
            if (dataset == null)
            {
                dataset = new DataSet();
            }

            if (InitDataBase(m_strConnection, m_strDBType))
            {
                ExecuteOracleFormat();

                IDbTransaction dbTransaction = null;

                try
                {
                    // 开始事务
                    dbTransaction = m_DBConnection.BeginTransaction();

                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.CommandType = CommandType.Text;
                    m_DBCommand.Transaction = dbTransaction;
                    m_DBCommand.CommandText = SQLProcess(strSQL);

                    dataset.Clear();
                    m_DBDataAdapter.SelectCommand = m_DBCommand;
                    m_DBDataAdapter.Fill(dataset);

                    // 事务提交
                    dbTransaction.Commit();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute SQL Successful";
                }
                catch (Exception e)
                {
                    if (dbTransaction != null)
                    {
                        // 事务回滚
                        dbTransaction.Rollback();
                    }

                    m_strOpeInfo = e.Message;
                    m_bIsSuccess = false;
                    dataset = null;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }

        // 执行SQL，返回是否成功，同时返回DATASET,包含多表
        public bool ExecuteSQL(String strSQL, ref DataSet dataset, bool isMulti)
        {
            if (dataset == null)
            {
                dataset = new DataSet();
            }

            if (InitDataBase(m_strConnection, m_strDBType))
            {
                ExecuteOracleFormat();

                IDbTransaction dbTransaction = null;

                try
                {
                    // 开始事务
                    dbTransaction = m_DBConnection.BeginTransaction();

                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.CommandType = CommandType.Text;
                    m_DBCommand.Transaction = dbTransaction;
                    m_DBCommand.CommandText = SQLProcess(strSQL);

                    m_DBDataAdapter.SelectCommand = m_DBCommand;
                    m_DBDataAdapter.Fill(dataset);

                    // 事务提交
                    dbTransaction.Commit();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute SQL Successful";
                }
                catch (Exception e)
                {
                    if (dbTransaction != null)
                    {
                        // 事务回滚
                        dbTransaction.Rollback();
                    }

                    m_strOpeInfo = e.Message;
                    m_bIsSuccess = false;
                    dataset = null;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }

        // 执行SQL，接受参数，返回是否成功，同时返回DATASET
        public bool ExecuteSQL(String strSQL, DBParam[] prams, ref DataSet dataset)
        {
            if (dataset == null)
            {
                dataset = new DataSet();
            }
            if (InitDataBase(m_strConnection, m_strDBType))
            {
                ExecuteOracleFormat();

                IDbTransaction dbTransaction = null;

                try
                {
                    // 开始事务
                    dbTransaction = m_DBConnection.BeginTransaction();

                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.CommandType = CommandType.Text;
                    m_DBCommand.Transaction = dbTransaction;
                    m_DBCommand.CommandText = SQLProcess(strSQL);


                    CreateNewInParameter(m_strDBType, prams);
                    dataset.Clear();
                    m_DBDataAdapter.SelectCommand = m_DBCommand;
                    m_DBDataAdapter.Fill(dataset);

                    // 事务提交
                    dbTransaction.Commit();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute SQL Successful";
                }
                catch (Exception e)
                {
                    if (dbTransaction != null)
                    {
                        // 事务回滚
                        dbTransaction.Rollback();
                    }

                    m_strOpeInfo = e.Message;
                    m_bIsSuccess = false;
                    dataset = null;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }



        // 执行存储过程，返回是否成功
        public bool ExecuteProc(String strProcName)
        {
            if (InitDataBase(m_strConnection, m_strDBType))
            {
                IDbTransaction dbTransaction = null;

                try
                {
                    // 开始事务
                    dbTransaction = m_DBConnection.BeginTransaction();

                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.CommandText = strProcName;
                    m_DBCommand.Transaction = dbTransaction;
                    m_DBCommand.CommandType = CommandType.StoredProcedure;

                    m_DBCommand.ExecuteNonQuery();

                    // 事务提交
                    dbTransaction.Commit();

                    m_strOpeInfo = "Execute Procedure Successful";
                    m_bIsSuccess = true;
                }
                catch (Exception e)
                {
                    if (dbTransaction != null)
                    {
                        // 事务回滚
                        dbTransaction.Rollback();
                    }

                    m_strOpeInfo = e.Message;
                    m_bIsSuccess = false;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }

        // 執行儲存過程數組,接受 SP名稱,參數列表 數組,返回是否成功
        public bool ExecuteProc(BatchProc[] arrayBatchProc)
        {
            if (InitDataBase(m_strConnection, m_strDBType))
            {
                IDbTransaction dbTransaction = null;

                try
                {
                    // 開始事務
                    dbTransaction = m_DBConnection.BeginTransaction();

                    // 開始解析Proc數組
                    foreach (BatchProc bp in arrayBatchProc)
                    {
                        m_DBCommand = m_DBConnection.CreateCommand();
                        m_DBCommand.Transaction = dbTransaction;
                        m_DBCommand.CommandText = bp.GetProcName();
                        m_DBCommand.CommandType = CommandType.StoredProcedure;

                        if (null == bp.GetProcParam())
                        {
                            // 无参数 
                        }
                        else
                        {
                            // 參數構造添加
                            DBParam[] prams = bp.GetProcParam();

                            CreateNewInParameter(m_strDBType, prams);
                        }

                        m_DBCommand.ExecuteNonQuery();
                    }

                    // 提交事务
                    dbTransaction.Commit();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute Proc Successful";
                }
                catch (Exception e)
                {
                    if (dbTransaction != null)
                    {
                        // 事务回滚
                        dbTransaction.Rollback();
                    }

                    m_bIsSuccess = false;
                    m_strOpeInfo = e.Message;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }


        /// <summary>
        /// 執行預儲程序，接受參數陣列，且傳回是否成功
        /// </summary>
        /// <param name="strProcName">預儲程序名稱</param>
        /// <param name="prams">參數陣列</param>
        /// <returns>是否成功</returns>
        public bool ExecuteProc(String strProcName, DBParam[] prams)
        {
            if (InitDataBase(m_strConnection, m_strDBType))
            {
                IDbTransaction dbTransaction = null;

                try
                {
                    // 开始事务
                    dbTransaction = m_DBConnection.BeginTransaction();

                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.CommandText = strProcName;
                    m_DBCommand.Transaction = dbTransaction;
                    m_DBCommand.CommandType = CommandType.StoredProcedure;

                    CreateNewInParameter(m_strDBType, prams);
                    m_DBCommand.ExecuteNonQuery();

                    // 事务提交
                    dbTransaction.Commit();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute Procedure Successful";
                }
                catch (Exception e)
                {
                    if (dbTransaction != null)
                    {
                        // 事务回滚
                        dbTransaction.Rollback();
                    }

                    m_strOpeInfo = e.Message;
                    m_bIsSuccess = false;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }

        /// <summary>
        /// 執行預儲程序，且傳回 DataSet
        /// </summary>
        /// <param name="strProcName"></param>
        /// <param name="dataset"></param>
        /// <returns>是否成功</returns>
        public bool ExecuteProc(String strProcName, ref DataSet dataset)
        {
            if (dataset == null)
            {
                dataset = new DataSet();
            }

            if (InitDataBase(m_strConnection, m_strDBType))
            {
                IDbTransaction dbTransaction = null;

                try
                {
                    // 开始事务
                    dbTransaction = m_DBConnection.BeginTransaction();

                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.CommandText = strProcName;
                    m_DBCommand.Transaction = dbTransaction;
                    m_DBCommand.CommandType = CommandType.StoredProcedure;

                    dataset.Clear();
                    m_DBDataAdapter.SelectCommand = m_DBCommand;
                    m_DBDataAdapter.Fill(dataset);

                    // 事务提交
                    dbTransaction.Commit();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute Procedure Successful";
                }
                catch (Exception e)
                {
                    if (dbTransaction != null)
                    {
                        // 事务回滚
                        dbTransaction.Rollback();
                    }

                    m_strOpeInfo = e.Message;
                    m_bIsSuccess = false;
                    dataset = null;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }

        // 执行存储过程，返回是否成功，同时返回DATASET，可包含多表
        public bool ExecuteProc(String strProcName, ref DataSet dataset, bool isMuti)
        {
            if (dataset == null)
            {
                dataset = new DataSet();
            }
            if (InitDataBase(m_strConnection, m_strDBType))
            {
                IDbTransaction dbTransaction = null;

                try
                {
                    // 开始事务
                    dbTransaction = m_DBConnection.BeginTransaction();

                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.CommandText = GetProcNameFromDBType(strProcName);
                    m_DBCommand.Transaction = dbTransaction;
                    m_DBCommand.CommandType = CommandType.StoredProcedure;


                    m_DBDataAdapter.SelectCommand = m_DBCommand;
                    m_DBDataAdapter.Fill(dataset);

                    // 事务提交
                    dbTransaction.Commit();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute Procedure Successful";
                }
                catch (Exception e)
                {
                    if (dbTransaction != null)
                    {
                        // 事务回滚
                        dbTransaction.Rollback();
                    }

                    m_strOpeInfo = e.Message;
                    m_bIsSuccess = false;
                    dataset = null;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }

        // 执行存储过程，接受参数列表，返回是否成功，同时返回DATASET
        public bool ExecuteProc(String strProcName, DBParam[] prams, ref DataSet dataset)
        {
            if (dataset == null)
            {
                dataset = new DataSet();
            }

            if (InitDataBase(m_strConnection, m_strDBType))
            {
                IDbTransaction dbTransaction = null;

                try
                {
                    // 开始事务
                    dbTransaction = m_DBConnection.BeginTransaction();

                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.CommandText = GetProcNameFromDBType(strProcName);
                    m_DBCommand.Transaction = dbTransaction;
                    m_DBCommand.CommandType = CommandType.StoredProcedure;

                    CreateNewInParameter(m_strDBType, prams);

                    dataset.Clear();
                    m_DBDataAdapter.SelectCommand = m_DBCommand;
                    m_DBDataAdapter.Fill(dataset);

                    // 事务提交
                    dbTransaction.Commit();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute Procedure Successful";
                }
                catch (Exception e)
                {
                    if (dbTransaction != null)
                    {
                        // 事务回滚
                        dbTransaction.Rollback();
                    }

                    m_strOpeInfo = e.Message;
                    m_bIsSuccess = false;
                    dataset = null;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }

        // 执行存储过程，接受参数列表，返回是否成功，同时返回SQL操作结果
        public bool ExecuteProc(String strProcName, DBParam[] prams, ref Int64 strOutSign)
        {
            if (InitDataBase(m_strConnection, m_strDBType))
            {
                IDbTransaction dbTransaction = null;

                try
                {
                    // 开始事务
                    dbTransaction = m_DBConnection.BeginTransaction();

                    m_DBCommand = m_DBConnection.CreateCommand();
                    m_DBCommand.CommandText = GetProcNameFromDBType(strProcName);
                    m_DBCommand.Transaction = dbTransaction;
                    m_DBCommand.CommandType = CommandType.StoredProcedure;

                    DbParameter dbParam = null;

                    CreateNewInParameter(m_strDBType, prams);

                    switch (m_strDBType)
                    {
                        case "MYSQL":
                            dbParam = new MySqlParameter("?RETURNVALUE", MySqlDbType.Int64);
                            break;
                        case "MSSQL":
                            dbParam = new SqlParameter("RETURNVALUE", SqlDbType.BigInt);
                            break;
                        //case "ORACLE":
                        //    dbParam = new OracleParameter("RETURNVALUE", OracleType.Number);
                        //    break;
                        default:
                            break;
                    }

                    // 将输出参数加入到参数列表
                    dbParam.Direction = ParameterDirection.Output;
                    m_DBCommand.Parameters.Add(dbParam);

                    m_DBCommand.ExecuteNonQuery();

                    // 事务提交
                    dbTransaction.Commit();

                    m_bIsSuccess = true;
                    m_strOpeInfo = "Execute Procedure Successful";

                    // 存储过程返回值
                    strOutSign = Int64.Parse((dbParam.Value).ToString());
                }

                catch (Exception e)
                {
                    if (dbTransaction != null)
                    {
                        // 事务回滚
                        dbTransaction.Rollback();
                    }

                    m_strOpeInfo = e.Message;
                    m_bIsSuccess = false;

                    strOutSign = -1;
                }
                finally
                {
                    Close();
                }
            }

            return m_bIsSuccess;
        }


        // 返回上次操作的结果
        public bool GetLastResult(out String strErrorOut)
        {
            strErrorOut = m_strOpeInfo;

            return GetLastResult();
        }
        public bool GetLastResult()
        {
            return m_bIsSuccess;
        }

        // 返回当前数据库连接字符串
        public String GetConn()
        {
            return m_strConnection;
        }

        // 返回当前数据库类型
        public String GetDBType()
        {
            return m_strDBType;
        }

        // 检测数据库是否被正确设置
        public bool IsCorrectConfig()
        {
            return m_bIsCorrectConfig;
        }

        // 数据库连线测试
        public bool ConnectTest(String strConnection, String strDBType)
        {
            InitDataBase(strConnection, strDBType);

            return m_bIsSuccess;
        }

        // 更新数据库连接字符串信息
        public bool UpdateDBConfig(String strConnection, String strDBType)
        {
            if (strConnection == m_strConnection && strDBType == m_strDBType)
            {
                m_bIsSuccess = true;
                m_strOpeInfo = "No Change Of Config Setting";
            }
            else
            {
                if (ConnectTest(strConnection, strDBType))
                {
                    m_strConnection = strConnection;
                    m_strDBType = strDBType;
                    m_bIsCorrectConfig = true;
                }
            }
            if (m_bIsSuccess == false)
                FileManage.logToFile("DB Connection Invalid: " + strConnection + " ");

            return m_bIsSuccess;
        }

        #endregion -----------------------公用接口定义------------------------


        #region -----------------------私有函数定义------------------------

        // 初始化数据库连接对象
        private bool InitDataBase(String strConnection, String strDBType)
        {
            bool bFlag = false;

            try
            {
                if (strDBType.Equals("MSSQL"))
                {
                    m_DBConnection = new SqlConnection(strConnection);
                    m_DBDataAdapter = new SqlDataAdapter();
                    m_DBParameter = new SqlParameter();
                    bFlag = true;
                }
                //else if (strDBType.Equals("ORACLE"))
                //{
                //    m_DBConnection = new OracleConnection(strConnection);
                //    m_DBDataAdapter = new OracleDataAdapter();
                //    m_DBParameter = new OracleParameter();
                //    bFlag = true;
                //}
                else if (strDBType.Equals("MYSQL"))
                {
                    m_DBConnection = new MySqlConnection(strConnection);
                    m_DBDataAdapter = new MySqlDataAdapter();
                    m_DBParameter = new MySqlParameter();
                    bFlag = true;
                }

                if (bFlag)
                {
                    m_bIsSuccess = true;
                    m_strOpeInfo = "InitDataBase Success";

                    // 建立数据连接
                    Open();
                }
                else
                {
                    m_bIsSuccess = false;
                    m_strOpeInfo = "InitDataBase Error";
                }
            }
            catch (Exception e)
            {
                Close();
                m_bIsSuccess = false;
                m_strOpeInfo = e.Message;
            }

            return m_bIsSuccess;
        }

        // 建立数据库连接
        private bool Open()
        {
            if (m_DBConnection != null)
            {
                try
                {
                    if (m_DBConnection.State == ConnectionState.Closed)
                    {
                        m_DBConnection.Open();
                    }
                    else if (m_DBConnection.State == ConnectionState.Broken)
                    {
                        m_DBConnection.Close();
                        m_DBConnection.Open();
                    }
                    m_strOpeInfo = "Connect To DataBase Successful";
                    m_bIsSuccess = true;
                }
                catch (Exception e)
                {
                    m_strOpeInfo = e.Message;
                    m_bIsSuccess = false;
                }
            }
            else
            {
                m_bIsSuccess = false;
                m_strOpeInfo = "Fail To Connect To DataBase";
            }

            return m_bIsSuccess;
        }

        // 关闭数据库连接
        private void Close()
        {
            if (m_DBConnection != null)
            {
                if (m_DBConnection.State == ConnectionState.Open)
                {
                    m_DBConnection.Close();
                }
            }

            Dispose();
        }

        // 释放数据库资源
        private void Dispose()
        {
            if (m_DBConnection != null)
            {
                m_DBConnection.Dispose();
                m_DBConnection = null;
            }
            if (m_DBDataAdapter != null)
            {
                m_DBDataAdapter = null;
            }
            if (m_DBParameter != null)
            {
                m_DBParameter = null;
            }
            if (m_DBCommand != null)
            {
                m_DBCommand.Dispose();
                m_DBCommand = null;
            }
        }

        // 执行Oracle时间格式化
        private bool ExecuteOracleFormat()
        {
            if (m_strDBType == "ORACLE")
            {
                String strSQL = "alter session set nls_date_format = 'yyyy-mm-dd hh24:mi:ss'";

                //IDbConnection DBConnectionForOracle = new OracleConnection(m_strConnection);

                m_DBCommandForOracle = m_DBConnection.CreateCommand();

                m_DBCommandForOracle.CommandType = CommandType.Text;

                m_DBCommandForOracle.CommandText = strSQL;

                m_DBCommandForOracle.ExecuteNonQuery();

                return true;
            }

            return false;
        }

        // SQL命令字符处理
        private String SQLProcess(String strSQL)
        {
            String strReturn = null;

            //DB special string process
            try
            {
                switch (m_strDBType)
                {
                    case "MSSQL":
                        {

                        }
                        break;

                    case "MYSQL":
                        {
                            strSQL = strSQL.Replace("getdate()", "now()");
                            strSQL = strSQL.Replace("Replace(convert(char(20),mailtask.mt_sendTime,111),'/','-')", "Date_FORMAT(mailtask.mt_sendTime,'%Y-%m-%d') ");
                        }
                        break;

                    case "ORACLE":
                        {

                        }
                        break;
                }
            }
            catch (DbException)
            {

            }


            if (null != strSQL && String.Empty != strSQL)
            {
                char[] cstr = { ';' };

                strReturn = strSQL.TrimEnd(cstr);
            }

            return strReturn;
        }

        // 为输入构造参数
        private bool CreateNewInParameter(String strType, DBParam[] dbParams)
        {
            bool bFlag = false;
            try
            {
                IDbDataParameter pushdbParam = null;

                switch (strType)
                {
                    case "MSSQL":
                        {
                            foreach (DBParam dbParam in dbParams)
                            {
                                pushdbParam = new SqlParameter();

                                pushdbParam.ParameterName = dbParam.m_strKey;
                                pushdbParam.DbType = dbParam.m_dbType;
                                pushdbParam.Value = dbParam.m_Value;
                                pushdbParam.Direction = dbParam.m_Direction;

                                m_DBCommand.Parameters.Add(pushdbParam);
                            }

                            bFlag = true;
                        }
                        break;

                    case "MYSQL":
                        {
                            foreach (DBParam dbParam in dbParams)
                            {
                                pushdbParam = new MySqlParameter();

                                pushdbParam.ParameterName = "?" + dbParam.m_strKey;
                                pushdbParam.DbType = dbParam.m_dbType;
                                pushdbParam.Value = dbParam.m_Value;
                                pushdbParam.Direction = dbParam.m_Direction;

                                m_DBCommand.Parameters.Add(pushdbParam);
                            }

                            bFlag = true;
                        }
                        break;

                    //case "ORACLE":
                    //    {
                    //        foreach (DBParam dbParam in dbParams)
                    //        {
                    //            pushdbParam = new OracleParameter();

                    //            if (dbParam.m_dbType == DbType.Int64 || dbParam.m_dbType == DbType.Int32)
                    //            {
                    //                pushdbParam = new OracleParameter(dbParam.m_strKey, OracleType.Number);
                    //            }
                    //            else
                    //            {
                    //                pushdbParam.ParameterName = dbParam.m_strKey;
                    //                pushdbParam.DbType = dbParam.m_dbType;
                    //            }

                    //            pushdbParam.Value = dbParam.m_Value;
                    //            pushdbParam.Direction = dbParam.m_Direction;

                    //            m_DBCommand.Parameters.Add(pushdbParam);
                    //        }

                    //        bFlag = true;
                    //    }
                    //    break;

                    case "ACCESS":
                    case "EXCEL":
                        pushdbParam = new OdbcParameter();
                        break;

                    default:
                        pushdbParam = new SqlParameter();
                        break;
                }
            }
            catch (DbException)
            {
                bFlag = false;
            }


            return bFlag;
        }

        #endregion -----------------------私有函数定义------------------------


        #region -----------------------成员变量定义------------------------

        // 数据库操作对象
        private IDbCommand m_DBCommand = null;
        private IDbConnection m_DBConnection = null;
        private IDbDataAdapter m_DBDataAdapter = null;
        private IDbDataParameter m_DBParameter = null;
        private IDbCommand m_DBCommandForOracle = null;

        // 操作提示信息
        private String m_strOpeInfo = "";

        // 操作结果标识
        private bool m_bIsSuccess = false;

        // 数据库连接字符串
        //private static String m_strConnection = "Data Source =192.168.1.11; Initial Catalog =new_DMS; User Id = test; Password = 1234";
        private String m_strConnection = @"Data Source = 192.168.1.43\SQL2008_CI; Initial Catalog = DMSdb; User Id = sa; Password = iscom@7o598966";
        // 数据库类型
        private static String m_strDBType = "MSSQL";

        // 数据库是否被正确设置
        private static bool m_bIsCorrectConfig = false;

        #endregion -----------------------成员变量定义------------------------


        #region -----------------------其它资料来源------------------------

        private IDbCommand m_otherDBComm = null;
        private IDbConnection m_otherDBConn = null;
        private IDbDataAdapter m_otherDBAdapter = null;

     

        // 指定连接字符串,获得DS
        public DataSet GetResultDS(String strType, String strConn, String strSQL)
        {
            bool bFlag = SetInit(strConn, strType);

            DataSet ds = new DataSet();

            try
            {
                m_otherDBComm = m_otherDBConn.CreateCommand();
                m_otherDBComm.CommandText = strSQL;
                m_otherDBAdapter.SelectCommand = m_otherDBComm;
                m_otherDBAdapter.Fill(ds);
            }
            catch (Exception)
            {
                ds = null;
            }

            finally
            {
                if (m_otherDBConn != null)
                {
                    m_otherDBConn.Close();
                    m_otherDBConn.Dispose();
                    m_otherDBConn = null;
                }
                if (m_otherDBComm != null)
                {
                    m_otherDBComm = null;
                }
                if (m_otherDBAdapter != null)
                {
                    m_otherDBAdapter = null;
                }
            }

            return ds;
        }

        // 指定连接字符,获取ArrayList (目前主要用于获得数据库的表字段)
        public ArrayList GetArrayList(String strType, String strConn, String strSQL)
        {
            DataSet ds = GetResultDS(strType, strConn, strSQL);

            ArrayList arrayList = null;

            if (null != ds && ds.Tables[0].Rows.Count > 0)
            {
                arrayList = new ArrayList();

                foreach (DataColumn dc in ds.Tables[0].Columns)
                {
                    arrayList.Add(dc.ColumnName.ToString());
                }
            }

            return arrayList;
        }

        // 指定用户名,密码,表名等信息获得DS
        public DataSet GetResultDS(String strType, String strAddr, String strUser, String strPass, String strDB, String strSQL)
        {
            String strConn = "";

            if (strType == "MYSQL" || strType == "MSSQL")
            {
                strConn = "Data Source=" + strAddr + ";DataBase=" + strDB + ";User ID=" + strUser + ";Password=" + strPass;
            }

            if (strType == "ORACLE")
            {
                // strConn = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" + strAddr +
                //          ")(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=" + strDB + ")));User Id=" +
                //        strUser + ";Password=" + strPass;
                strConn = "Data Source=" + strDB + ";User ID=" + strUser + ";Password=" + strPass + ";";
            }

            return GetResultDS(strType, strConn, strSQL);
        }

        // 指定用户名,密码,表名等信息获得ArrayList
        public ArrayList GetArrayList(String strType, String strAddr, String strUser, String strPass, String strDB, String strSQL)
        {
            String strConn = "";

            if (strType == "MYSQL" || strType == "MSSQL")
            {
                strConn = "Data Source=" + strAddr + ";DataBase=" + strDB + ";User ID=" + strUser + ";Password=" + strPass;
            }

            if (strType == "ORACLE")
            {
                //strConn = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" + strAddr +
                //        ")(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=" + strDB + ")));User Id=" +
                //      strUser + ";Password=" + strPass;
                strConn = "Data Source=" + strDB + ";User ID=" + strUser + ";Password=" + strPass + ";";
            }

            return GetArrayList(strType, strConn, strSQL);
        }

        // 初始化及连接建立
        private bool SetInit(String strConnection, String strDBType)
        {
            try
            {
                if (strDBType.Equals("MSSQL"))
                {
                    m_otherDBConn = new SqlConnection(strConnection);
                    m_otherDBAdapter = new SqlDataAdapter();
                }
                //else if (strDBType.Equals("ORACLE"))
                //{
                //    m_otherDBConn = new OracleConnection(strConnection);
                //    m_otherDBAdapter = new OracleDataAdapter();
                //}
                else if (strDBType.Equals("MYSQL"))
                {
                    m_otherDBConn = new MySqlConnection(strConnection);
                    m_otherDBAdapter = new MySqlDataAdapter();
                }

                m_otherDBConn.Open();
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        #endregion
        private string GetProcNameFromDBType(string strProcName)
        {
            string getProcName = "";
            switch (m_strDBType)
            {
                case "MYSQL":
                    getProcName = strProcName;
                    break;
                case "MSSQL":
                    getProcName = strProcName;
                    break;
                case "ORACLE":
                    getProcName = "SYSTEM." + strProcName;
                    break;
                default:
                    getProcName = strProcName;
                    break;
            }
            return getProcName;
        }
    }
}