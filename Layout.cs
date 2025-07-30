using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
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
        public Layout()
        {
            sql = new Sqlite();
            level1 = sql.GetLevel1();
        }

        #region 功能区用户控件交互
        public async void Level1ButtonClick(string level1)
        {
            if (level1Active == null)
            {
                webView = new WebView();
                string jsCode = $"layout.InitLevel2('{level1}')";
                await webView.RunJS(jsCode);
                webView.ShowControl(true);
                level1Active = level1;
            }
            else if (level1 != level1Active)
            {
                string jsCode = $"layout.InitLevel2('{level1}')";
                await webView.RunJS(jsCode);
                webView.ShowControl(true);
                level1Active = level1;
            }
            else if (webView.controlTaskPane.Visible == false)
            {
                webView.ShowControl(true);
            }
            else
            {
                webView.ShowControl(false);
            }
        }
        #endregion

        #region 传递Sqlite查询
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
