using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

namespace Excel_Toolkit
{
    public class Sqlite
    {
        public SQLiteCommand cmd;
        public SQLiteConnection conn;
        public string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Excel_Toolkit.db");
        public SQLiteDataReader reader;

        public Sqlite()
        {
            conn = new SQLiteConnection($"Data Source={dbPath};Version=3;");
            conn.Open();
        }

        #region 查询数据表

        public List<Dictionary<string, object>> GetTableAll(string tableName)
        {
            var sql = $"SELECT * FROM {tableName}";
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
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    var columnName = reader.GetName(i);
                    var value = reader.GetValue(i);
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
            conn.Close();
            conn.Dispose();
        }

        #region 获取函数数据

        public string[] GetLevel1()
        {
            const string sql = @"SELECT MIN(id) as id, level1 FROM 内容 GROUP BY level1 ORDER BY id ASC";
            var list = GetData(sql);
            var level1 = new string[list.Count];
            for (var i = 0; i < list.Count; i++) level1[i] = list[i]["level1"].ToString();
            return level1;
        }

        public List<Dictionary<string, object>> GetLevel2(string level1)
        {
            var sql = $"SELECT level1,level2,info FROM 内容 WHERE level1 = '{level1}'";
            return GetData(sql);
        }

        public string GetContent(string level1, string level2)
        {
            var sql = $"SELECT content FROM 内容 WHERE level1 = '{level1}' AND level2 = '{level2}'";
            var list = GetData(sql);
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
                return new List<Dictionary<string, object>>();
            var sql = "SELECT level1,level2,info FROM 内容 WHERE 1>1";
            if (content) sql += $" OR content LIKE '%{key}%'";
            if (info) sql += $" OR info LIKE '%{key}%'";
            if (level2) sql += $" OR level2 LIKE '%{key}%'";
            return GetData(sql);
        }

        #endregion
    }
}