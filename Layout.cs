using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace Excel_Toolkit
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Layout
    {
        public ThisAddIn addIn;
        public Sqlite sql;
        public Ribbon ribbon;
        public WebView webView;
        public string[] level1;
        public string level1Active;
        public Excel.Application app;
        public Layout()
        {
            sql = new Sqlite();
            level1 = sql.GetLevel1();
            addIn = Globals.ThisAddIn;
        }


        #region 功能区用户控件交互
        public void Level1ButtonClick(string level1)
        {
            WebView.Pane pane = WebView.Pane.query;
            if (webView == null) webView = new WebView(pane);
            if (webView.pane != pane) webView.LoadHTML(pane);
            if (level1 == level1Active)
            {
                if (webView.controlTaskPane.Visible) webView.controlTaskPane.Visible = false;
            }
            else
            {
                webView.RunJS($"InitLevel2('{level1}')");
                level1Active = level1;
                webView.controlTaskPane.Visible = true;
            }
        }
        #endregion

        #region 传递Sqlite查询
        #region 函数查询部分
        public string[] GetLevel1()
        {
            string[] data = sql.GetLevel1();
            return data;
        }
        public string GetLevel2(string level1)
        {
            List<Dictionary<string, object>> data = sql.GetLevel2(level1);
            return DataToJson(data);
        }
        public string Search(string key, bool content = true, bool info = true, bool level2 = true)
        {
            List<Dictionary<string, object>> data = sql.Search(key, content, info, level2);
            return DataToJson(data);
        }
        public string GetContent(string level1, string level2)
        {
            return sql.GetContent(level1, level2);
        }
        #endregion

        private string DataToJson(object data)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = null,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            return JsonSerializer.Serialize(data, options);
        }
        #endregion
    }
}
