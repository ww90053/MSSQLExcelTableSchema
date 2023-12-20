using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ExcelTableSchemaApp
{
    /// <summary>
    /// 資料庫存取函式
    /// 2014.05.20   重構  by Jason
    /// </summary>
     public class DBController
    {
         SQLAccessor m_db;
         public DBController(string sConnString)
         {
             m_db = new SQLAccessor(sConnString);
         }
         

        /// <summary>
        /// 取得所有資料表
        /// </summary>
        /// <returns></returns>
         public DataTable  getAllTableSchema()
        {

            string sql = " SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES " +
                         "  order by TABLE_NAME ";
            return m_db.select(sql);
        
        }

         /// <summary>
         /// 取得 Table 描述
         /// 2014.05.20 Jason
         /// 2014.07.14 Jason 增加轉型，將傳回的資料全轉為字串
         /// </summary>
         /// <param name="TableName">表格名稱</param>
         /// <returns></returns>
         public string getTableDesc(string TableName)
         {
             string sql = string.Format("SELECT CONVERT(varchar,value) AS value FROM ::fn_listextendedproperty(NULL, 'user', 'dbo', 'table', '{0}', NULL, NULL)", TableName);
             return m_db.selectOne<string>(sql);

         }


        /// <summary>
        ///  取得Table PK值
        /// </summary>
        /// <param name="TableName"></param>
        /// <returns></returns>
         public string getPKValue(string TableName)
        {
            //  取得 PK
            string pk = "";
            string sql = "";
            sql = " SELECT COLUMN_NAME " +
                    " FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE " +
                    " WHERE TABLE_NAME = '" + TableName + "' " +
                    " 	and CONSTRAINT_NAME like 'PK%' ";


            DataTable tbl = m_db.select(sql);
        
                      

            int i = 0;
            foreach (DataRow row in tbl.Rows)
            {
                if (i == 0)
                {
                    pk = row["COLUMN_NAME"].ToString();
                }
                else
                {
                    pk += "," + row["COLUMN_NAME"].ToString();
                }
                i++;
            }

            return pk;
        }


        /// <summary>
        ///  取得Table FK值
        /// </summary>
        /// <param name="TableName"></param>
        /// <returns></returns>
         public string getFKValue(string TableName)
        {
            //  取FK
            string fk = "";
            string sql = "";
            sql = " SELECT COLUMN_NAME " +
                    " FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE " +
                    " WHERE TABLE_NAME = '" + TableName + "' " +
                    " 	and CONSTRAINT_NAME like 'FK%' ";



            DataTable tbl = m_db.select(sql);

            int i = 0;
            foreach (DataRow row in tbl.Rows)
            {
                if (i == 0)
                {
                    fk = row["COLUMN_NAME"].ToString();
                }
                else
                {
                    fk += "," + row["COLUMN_NAME"].ToString();
                }
                i++;
            }

            return fk;
        }


        /// <summary>
        ///  取得 Table 屬性資料
        /// </summary>
        /// <param name="TableName"></param>
        /// <returns></returns>
         public  DataTable getTableAttr(string TableName)
        {


            string sql = "";
            sql = " SELECT COLUMN_NAME  " +
                "    	   ,DATA_TYPE  " +
                " 		,case DATA_TYPE when 'ntext' then ''  " +
                " 		 when 'datetime' then '' " +
                " 		 when 'float' then '' " +
                " 		 when 'decimal' then cASt(NUMERIC_PRECISION AS nvarchar(20))+','+cASt(NUMERIC_SCALE AS nvarchar(20)) " +
                " 		 when 'nvarchar' then cASe CHARACTER_MAXIMUM_LENGTH when -1 then 'max' ELSE cASt(CHARACTER_MAXIMUM_LENGTH AS nvarchar(20))  END " +
                " 		 when 'varchar' then cASe CHARACTER_MAXIMUM_LENGTH when -1 then 'max' ELSE cASt(CHARACTER_MAXIMUM_LENGTH AS nvarchar(20))  END " +
                " 		 when 'nchar' then cASe CHARACTER_MAXIMUM_LENGTH when -1 then 'max' ELSE cASt(CHARACTER_MAXIMUM_LENGTH AS nvarchar(20))  END " +
                " 		 when 'char' then cASe CHARACTER_MAXIMUM_LENGTH when -1 then 'max' ELSE cASt(CHARACTER_MAXIMUM_LENGTH AS nvarchar(20))  END " +
                "          ELSE '' END AS CHARACTER_MAXIMUM_LENGTH " +
                "    		,NUMERIC_PRECISION  " +
                "         ,NUMERIC_SCALE " +
                "         ,A.IS_NULLABLE " +
                "         ,C.value AS [Description] " +
                "   FROM INFORMATION_SCHEMA.COLUMNS A " +
                "   INNER JOIN sys.columns B ON OBJECT_NAME(B.object_id) = A.TABLE_NAME AND B.name = A.COLUMN_NAME  " +
                "   LEFT OUTER JOIN sys.extended_properties C ON C.major_id = B.object_id AND C.minor_id = B.column_id AND C.name = 'MS_Description'  " +
                " WHERE 1=1 " +
                "   AND OBJECTPROPERTY(B.object_id, 'IsMsShipped')=0  " +
                "   AND OBJECT_NAME(B.object_id) = '" + TableName + "' ORDER BY ORDINAL_POSITION";

            return m_db.select(sql);

        }



    }



}
