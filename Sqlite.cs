using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
namespace Excel_Toolkit
{
    public class Sqlite
    {
        public string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Excel_Toolkit.db");
        public SQLiteCommand cmd;
        public SQLiteDataReader reader;
        public SQLiteConnection conn;

        public Sqlite()
        {
            conn = new SQLiteConnection("Data Source=" + dbPath + ";Version=3;");
            conn.Open();
        }
        #region 获取函数数据
        public string[] GetLevel1()
        {
            string sql = @"SELECT MIN(id) as id, level1 FROM 内容 GROUP BY level1 ORDER BY id ASC";
            List<Dictionary<string, object>> list = GetData(sql);
            string[] level1 = new string[list.Count];
            for (int i = 0; i < list.Count; i++)
            {
                level1[i] = list[i]["level1"].ToString();
            }
            return level1;
        }

        public List<Dictionary<string, object>> GetLevel2(string level1)
        {
            string sql = $"SELECT level1,level2,info FROM 内容 WHERE level1 = '{level1}'";
            return GetData(sql);
        }

        public string GetContent(string level1, string level2)
        {
            string sql = $"SELECT content FROM 内容 WHERE level1 = '{level1}' AND level2 = '{level2}'";
            List<Dictionary<string, object>> list = GetData(sql);
            return list[0]["content"].ToString();
        }

        public List<Dictionary<string, object>> Search(
            string key,
            bool content = true,
            bool info = true,
            bool level2 = true
            )
        {
            if (string.IsNullOrEmpty(key) || !(content || info || level2))
            {
                Console.WriteLine("参数错误");
                return new List<Dictionary<string, object>>();
            }
            string sql = "SELECT level1,level2,info FROM 内容 WHERE 1>1";
            if (content) sql += $" OR content LIKE '%{key}%'";
            if (info) sql += $" OR info LIKE '%{key}%'";
            if (level2) sql += $" OR level2 LIKE '%{key}%'";
            return GetData(sql);
        }

        public List<Dictionary<string, object>> GetTableAll(string tableName)
        {
            string sql = $"SELECT * FROM {tableName}";
            return GetData(sql);
        }
        #endregion

        public List<Dictionary<string, object>> GetData(string sql)
        {
            var list = new List<Dictionary<string, object>>();
            cmd = new SQLiteCommand(sql, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var row = new Dictionary<string, object>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    string columnName = reader.GetName(i);
                    object value = reader.GetValue(i);
                    row[columnName] = value;
                    Console.WriteLine(value);
                }
                list.Add(row);
            }
            reader.Close();
            cmd.Dispose();
            return list;
        }

        public void Close()
        {
            if (conn != null)
            {
                conn.Close();
                conn.Dispose();
            }
        }
    }
}
