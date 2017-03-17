using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace 入库工具
{
    class SqlHelper
    {
        public static readonly string CnStr = "server=suznt004;user id=andy;password=123;database=EMsystem;Connect Timeout=10";
       

        public static DataSet ExcuteDataSet(string Sql)
        {
            try
            {
                SqlConnection cn = new SqlConnection(CnStr);
                SqlCommand cmd = new SqlCommand(Sql, cn);
                SqlDataAdapter dp = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                cn.Open();
                dp.Fill(ds);
                cn.Close();
                cn.Dispose();
                return ds;
            }
            catch (Exception)
            {
                return null;
            }
        }
      

       
        public static string ExcuteStr(string Sql)
        {
            try
            {
                SqlConnection cn = new SqlConnection(CnStr);
                cn.Open();
                SqlCommand cmd = new SqlCommand(Sql, cn);
                SqlDataReader dr = cmd.ExecuteReader();
                string str = null;
                while (dr.Read())
                {
                    str = dr[0].ToString();
                }
                cn.Close();
                cn.Dispose();
                return str;
            }
            catch (Exception)
            {
                return null;
            }
        }
        public static bool ExecuteNonQuery(string Sql)
        {
            try
            {
                SqlConnection cn = new SqlConnection(CnStr);
                SqlCommand cmd = new SqlCommand(Sql, cn);
                cn.Open();
                cmd.ExecuteNonQuery();
                cn.Close();
                cn.Dispose();
                return true;
            }
            catch (Exception)
            {
                return false;
            }


        }
    }
}
