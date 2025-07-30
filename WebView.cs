using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
namespace Excel_Toolkit
{
    public class WebView : Microsoft.Web.WebView2.WinForms.WebView2
    {
        public UserControl control;
        public Microsoft.Office.Tools.CustomTaskPane controlTaskPane;
        public Layout layout = Globals.ThisAddIn.layout;
        public ThisAddIn addIn = Globals.ThisAddIn;
        public string userDataPath = @"C:\temp\MyWebView2";
        public string htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Excel_Toolkit.html");
        public Task initTask;
        public WebView()
        {
            control = new UserControl();
            Dock = DockStyle.Fill;
            controlTaskPane = addIn.CustomTaskPanes.Add(control, "Excel Toolkit");
            initTask = InitWebView();
            control.Controls.Add(this);
            controlTaskPane.Width = 1200;
            layout.webView = this;
        }
        async Task InitWebView()
        {
            var env = await CoreWebView2Environment.CreateAsync(null, userDataPath);
            await EnsureCoreWebView2Async(env);
            CoreWebView2.AddHostObjectToScript("Layout", layout);
            string html = File.ReadAllText(htmlPath);
            NavigateToString(html);
        }
        public async Task RunJS(string js)
        {
            await initTask;
            await CoreWebView2.ExecuteScriptAsync(js);
        }
        public void ShowControl(bool isShow)
        {
            controlTaskPane.Visible = isShow;
        }
    }
}
