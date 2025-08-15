using System;
using System.Runtime.InteropServices;
using System.Text.Encodings.Web;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Toolkit
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Layout
    {
        public ThisAddIn addIn;
        public Excel.Application app;
        public string appPath = AppDomain.CurrentDomain.BaseDirectory;
        public string level1Active;
        public Ribbon ribbon;
        public Sqlite sql;
        public WebView webView;

        public Layout()
        {
            sql = new Sqlite();
            addIn = Globals.ThisAddIn;
        }

        public Layout(bool initSql = true)
        {
            if (initSql) sql = new Sqlite();

            addIn = Globals.ThisAddIn;
        }

        //#region 功能区用户控件交互
        //public void QueryClick(string level1)
        //{
        //    WebView.Pane pane = WebView.Pane.Query;
        //    if (webView == null) webView = new WebView(pane);
        //    if (webView.pane != pane) webView.LoadHtml(pane);
        //    if (level1 == level1Active)
        //    {
        //        webView.Visible = !webView.Visible;
        //    }
        //    else
        //    {
        //        webView.RunJavaScript($"InitLevel2('{level1}')");
        //        level1Active = level1;
        //        webView.Visible = true;
        //    }
        //}
        //public void NameClick()
        //{
        //    WebView.Pane pane = WebView.Pane.Name;
        //    level1Active = null;
        //    if (webView == null) webView = new WebView(pane);
        //    if (webView.pane != pane) { 
        //        webView.LoadHtml(pane); 
        //        webView.Visible = true;
        //    }
        //    else
        //    {
        //        webView.Visible = !webView.Visible;
        //    }
        //}
        //#endregion

        #region 传递Sqlite查询

        #region 函数查询部分

        public string[] GetLevel1()
        {
            var data = sql.GetLevel1();
            return data;
        }

        public string GetLevel2(string level1)
        {
            var data = sql.GetLevel2(level1);
            return DataToJson(data);
        }

        public string Search(string key, bool content = true, bool info = true, bool level2 = true)
        {
            var data = sql.Search(key, content, info, level2);
            return DataToJson(data);
        }

        public string GetContent(string level1, string level2)
        {
            return sql.GetContent(level1, level2);
        }

        #endregion

        #region 名称管理器查询部分

        public string GetNameTable()
        {
            var data = sql.GetTableAll("名称管理器");
            foreach (var row in data)
            {
                row["isInApp"] = row["isInApp"].ToString() == "1";
                row["isInBook"] = row["isInBook"].ToString() == "1";
                row["isInSheet"] = row["isInSheet"].ToString() == "1";
            }

            return DataToJson(data);
        }

        #endregion

        #region Power Query模块

        public string GetPowerQueryTable()
        {
            var data = sql.GetTableAll("PQ");
            var json = DataToJson(data);
            return json;
        }

        #endregion

        private string DataToJson(object data)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = null,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            return JsonSerializer.Serialize(data, options);
        }

        #endregion
    }
}