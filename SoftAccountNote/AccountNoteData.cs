using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftAccountNote
{
    class AccountNoteData
    {
        static string s1 = Directory.GetCurrentDirectory();
        OleDbConnection oleDb = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + s1 + "/" + "AccountNote.accdb;Persist Security Info=True");

        public AccountNoteData() //构造函数
        {
            oleDb.Open();
        }
        public bool HandleSQL(string sqlstr)
        {
            
            OleDbCommand oleDbCommand = new OleDbCommand(sqlstr, oleDb);
            int i = oleDbCommand.ExecuteNonQuery(); //返回被修改的数目
            return i > 0;
        }
        public DataTable Query(string tableName)
        {
            string sql = "select * from " +tableName + "";
            //获取表1的内容
            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, oleDb); //创建适配对象
            DataTable dt = new DataTable(); //新建表对象
            dbDataAdapter.Fill(dt); //用适配对象填充表对象
            return dt;
        }

        public DataTable MonQuery(string tableName, string MonDate)
        {
            string sql = "select'" + MonDate + "'from " + tableName + "";
            //获取表1的内容
            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, oleDb); //创建适配对象
            DataTable dt = new DataTable(); //新建表对象
            dbDataAdapter.Fill(dt); //用适配对象填充表对象

            return dt;
        }

        public bool Add(string consume, string kind, string remark, string date)
        {
            string sql = "insert into MainData (花费,分类,备注,日期) values ('" + consume + "','" + kind + "','" + remark + "','" + date + "')";
            //往表1添加一条记录，昵称是LanQ，账号是2545493686
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDb);
            int i = oleDbCommand.ExecuteNonQuery(); //返回被修改的数目
            return i > 0;
          
        }
        public bool AddBuget(string MonDate, string MonBuget)
        {
            string sql = "insert into BaseData (月份,预算) values ('" + MonDate +  "','" + MonBuget + "')";
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDb);
            int i = oleDbCommand.ExecuteNonQuery(); //返回被修改的数目
            return i > 0;
        }
        public bool Del()
        {
            string sqlOne = "delete * from BaseData";
            string sqlTwo = "delete * from MainData";
            OleDbCommand oleDbCommand1 = new OleDbCommand(sqlOne, oleDb);
            OleDbCommand oleDbCommand2 = new OleDbCommand(sqlTwo, oleDb);
            int i = oleDbCommand1.ExecuteNonQuery();
            int j = oleDbCommand2.ExecuteNonQuery();
            return i+j > 0;
        }
        public bool Change()
        {
            string sql = "update MainData set 密码='233333' where 昵称='你好'";
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDb);
            int i = oleDbCommand.ExecuteNonQuery();
            return i > 0;
        }

        public bool ChangeBuget(string MonDate, string MonBuget)
        {
            string sql = "update BaseData set 预算='" + MonBuget + "'where 月份='" + MonDate + "'";
            OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDb);
            int i = oleDbCommand.ExecuteNonQuery();
            return i > 0;
        }
        public void Close()
        {
            oleDb.Close();
            oleDb.Dispose();
        }
    }
}
