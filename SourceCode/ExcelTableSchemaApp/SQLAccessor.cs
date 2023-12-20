using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Diagnostics;

namespace ExcelTableSchemaApp
{
    /// <summary>
    /// 資料庫存取類別簡易版，只留必要的 select 函式
    /// 2014.05.20 Jason
    /// </summary>
	public class SQLAccessor : IDisposable
	{
		#region 建構式

		/// <summary>
		/// 建構式
        /// </summary>
        /// <param name="ConnString">連線字串</param>
        public SQLAccessor(string ConnString)
        {
            m_sConnString = ConnString;

    
        }


        /// <summary>
		/// 第2建構式
        /// </summary>
        /// <param name="builder">連線字串 builder</param>
		public SQLAccessor(SqlConnectionStringBuilder builder)
        {
            m_sConnString = builder.ConnectionString;

        }



	

		#endregion


		#region Member

		/// <summary>
		/// 連線字串，此處只用於建構子設定到 SqlConnection 物件中。
		/// </summary>
		private string m_sConnString;



		/// <summary>
		/// 連線字串 (唯讀)，由 SqlConnection 物件取出的字串
		/// </summary>
		public string ConnString
		{
			get { return m_conn.ConnectionString; }
		}
		/// <summary>
		/// 提供外界資料庫名稱
		/// 2010.11.11
		/// </summary>
		public string DB
		{
			get { return m_conn.Database; }
		}
		/// <summary>
		/// 提供外界資料庫 IP
		/// </summary>
		public string DBIP
		{
			get { return m_conn.DataSource; }
		}
		/// <summary>
		/// 連線物件
		/// </summary>
		private SqlConnection m_conn;

        /// <summary>
        /// 取得連線物件 (緩式評估)
        /// 2014.3.27 Jason
        /// </summary>
        /// <param name="connString">連線字串，若為空白，表示採用目前已設定的連線字串</param>
        /// <returns></returns>
        public SqlConnection getConnection(string connString="")
        {
                      
           if (string.IsNullOrEmpty(connString))
           {
               if (m_conn == null)
                   m_conn = new SqlConnection(m_sConnString);

           }
           else
           {
               m_conn = new SqlConnection(connString);
           }

           return m_conn;
        }

	

		/// <summary>
		/// 提供 DataMapper 關閉連線的功能
		/// 2011.2.1
		/// </summary>
		public void Close()
		{

			if (m_conn.State == ConnectionState.Open)
				m_conn.Close();

		}

		/// <summary>
		/// 釋放資源的 Dispose 函式。
		/// </summary>
		public void Dispose()
		{
			if (m_conn != null)
			{
				m_conn.Close();
				m_conn = null;
			}
		}

		/// <summary>
		/// SqlCommand 的逾時設定
		/// </summary>
		private static int s_nCommandTimeout;
		
		/// <summary>
		/// SqlCommand 的逾時設定。小於等於0 時表示不指定。
		/// </summary>
		public static int CommandTimeout
		{
			get
			{
				return (s_nCommandTimeout);
			}

			set
			{
				s_nCommandTimeout = value;
			}
		}

	

		#endregion

		


		#region 基本資料庫操作

		#region Select
		/// <summary>
		/// Select
		/// </summary>
		/// <remarks>
		/// 參數由 Parameter 串列提供。
		/// </remarks>
		/// <example>select 的使用範例
		/// <code>
		/// 
		/// </code>
		/// </example>
		/// <param name="cmd">SQL 命令字串</param>
		/// <returns></returns>
		public DataTable select(string cmd)
		{
			DataTable dt = new DataTable();

			try
			{
				
				dt = executeDataTable(getConnection(), false, CommandType.Text, cmd, null);

			}
			catch (Exception ex)
			{
				// 拋至外層
				throw new Exception(ex.Message);
			}
			return dt;
		}


		/// <summary>
		/// 取得單一欄位值。
		/// </summary>
		/// <remarks>
		/// 使用泛型，減低呼叫者的 type-cast 錯誤。
		/// </remarks>
		/// <example>selectOne 的使用範例。
		/// <code>
		/// 
		///   int ret = db.selectOne{int}("Select TB86F05 From [dbo].[TB86]");   //  sorry 我沒辦法在註解中用 "大於" 和 "小於"，會造成輸出說明文件的錯誤，所以用"大括號"代替。
		///   
		///   或使用參數：
		///  
		///    db.Parameters.Clear();
		///    db.Parameters.Add(new Parameter("@param0", 99));
		///    int ret = db.selectOne{int}("Select TB86F05 From [dbo].[TB86] WHERE TB86F01=@param0");
		/// </code>
		/// </example>
		/// <typeparam name="T">型別</typeparam>
		/// <param name="cmdText">SQL 字串</param>
		/// <returns>傳回的欄位</returns>
		public T selectOne<T>(string cmdText)
		{
			T ret = default(T);
			try
			{
                ret = executeScalar<T>(getConnection(), false, CommandType.Text, cmdText, null);
			}
			catch (Exception ex)
			{
				// 拋至外層
				throw new Exception(ex.Message);
			}
			return ret;
		}

		#endregion

		

	


	
		

		#endregion








		#region Implementations 底層函式

		

		/// <summary>
		/// 處理 command 的共通設定，由各 execute 函式呼叫。
		/// </summary>
		/// <param name="cmd"></param>
		/// <param name="conn"></param>
		/// <param name="trans">交易，若非交易則傳入null</param>
		/// <param name="cmdType">command 種類：Text: 文字命令,  StoredProcedure: 預儲程序</param>
		/// <param name="cmdText">SQL 命令字串</param>
		/// <param name="cmdParms">參數陣列，若無參數則傳入 null</param>
		private void processCmd(SqlCommand cmd, SqlConnection conn, SqlTransaction trans, CommandType cmdType, string cmdText, SqlParameter[] cmdParms)
		{
			if (conn.State != ConnectionState.Open)
			{
				conn.Open();

			}
			cmd.Connection = conn;
			cmd.CommandText = cmdText;
			// 有指定 CommandTimeout 時，進行設定
			if (CommandTimeout > 0)
				cmd.CommandTimeout = CommandTimeout;

			if (trans != null)
				cmd.Transaction = trans;

			cmd.CommandType = cmdType;
			if (cmdParms != null)
			{
				foreach (SqlParameter parm in cmdParms)
					cmd.Parameters.Add(parm);

			}   // end if

		}

		/// <summary>
		/// 不傳回資料的執行動作。重載版，增加 auto close。
		/// </summary>
		/// <param name="conn"></param>
		/// <param name="tran"></param>
		/// <param name="autoClose">是否要在此次存取結束就關閉連線。</param>
		/// <param name="cmdType"></param>
		/// <param name="cmdText"></param>
		/// <param name="cmdParms"></param>
		/// <returns></returns>
		private int executeNonQuery(SqlConnection conn, SqlTransaction tran, bool autoClose, CommandType cmdType, string cmdText, params SqlParameter[] cmdParms)
		{
			SqlCommand cmd = new SqlCommand();
			processCmd(cmd, conn, tran, cmdType, cmdText, cmdParms);
			int ret = 0;





			try
			{
				ret = cmd.ExecuteNonQuery();
			}
			catch (Exception ex)
			{

				// 拋至外層
				throw new Exception(ex.Message);
			}
			finally
			{
				cmd.Dispose();
				if (autoClose)
					conn.Close();
			}



			cmd.Parameters.Clear();
			return ret;
		}


	

		/// <summary>
		/// 只傳回單一欄位。
		/// 2011.03.13 增加是否為null的判斷，否則資料傳回空值時將出現問題。
		/// 2011.10.18 同上，增加 DateTime 的處理。
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="conn"></param>
		/// <param name="autoClose">是否要在此次存取結束就關閉連線。</param>
		/// <param name="cmdType"></param>
		/// <param name="cmdText"></param>
		/// <param name="cmdParms"></param>
		/// <returns></returns>
		private T executeScalar<T>(SqlConnection conn, bool autoClose, CommandType cmdType, string cmdText, params SqlParameter[] cmdParms)
		{
			SqlCommand cmd = new SqlCommand();
			processCmd(cmd, conn, null, cmdType, cmdText, cmdParms);

			
			object ret = null;
			T tRet;



			try
			{

				ret = cmd.ExecuteScalar();
				if (ret == null || ret == DBNull.Value)
				{
					if (typeof(T) == System.Type.GetType("System.String"))
						ret = (string)"";
					else if (typeof(T) == System.Type.GetType("System.DateTime"))
						ret = DateTime.MinValue;
					else
						ret = (object)0;
				}
				tRet = (T)ret;
			}
			catch (Exception ex)
			{

				// 拋至外層
				throw new Exception(ex.Message);
			}
			finally
			{
				cmd.Dispose();
				if (autoClose)
					conn.Close();
			}


			cmd.Parameters.Clear();

			return tRet;
		}



		/// <summary>
		/// 傳回 DataTable 的查詢動作。
		/// </summary>
		/// <param name="conn"></param>
		/// <param name="autoClose">是否要在此次存取結束就關閉連線。</param>
		/// <param name="cmdType"></param>
		/// <param name="cmdText"></param>
		/// <param name="cmdParms"></param>
		/// <returns></returns>
		private DataTable executeDataTable(SqlConnection conn, bool autoClose, CommandType cmdType, string cmdText, params SqlParameter[] cmdParms)
		{
			SqlCommand cmd = new SqlCommand();
			processCmd(cmd, conn, null, cmdType, cmdText, cmdParms);
			DataTable dt = new DataTable();


			try
			{
				SqlDataAdapter sda = new System.Data.SqlClient.SqlDataAdapter(cmd);
				sda.Fill(dt);
				sda.Dispose();
			}
			catch (Exception ex)
			{

				// 拋至外層
				throw new Exception(ex.Message);
			}
			finally
			{
				cmd.Dispose();
				if (autoClose)
					conn.Close();
			}




			return dt;


		}



		#endregion


		



	}
}
