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
        public Microsoft.Office.Tools.CustomTaskPane controlTaskPane;
        public Layout layout = Globals.ThisAddIn.layout;
        public string appPath = AppDomain.CurrentDomain.BaseDirectory;
        public UserControl control;
        public string htmlPath;
        public Task initTask;
        public Task paneTask;
        public Pane pane;
        public WebView(Pane pane)
        {
            control = new UserControl();
            Dock = DockStyle.Fill;
            controlTaskPane = layout.addIn.CustomTaskPanes.Add(control, "Excel Toolkit");
            control.Controls.Add(this);
            controlTaskPane.Width = 1200;
            layout.webView = this;
            initTask = InitWebView(pane);
          
        }
        public enum Pane
        {
            name,
            query
        }
        async Task InitWebView(Pane pane)
        {
            this.pane = pane;
            var env = await CoreWebView2Environment.CreateAsync(null, @"C:\temp\MyWebView2");
            await EnsureCoreWebView2Async(env);
            CoreWebView2.AddHostObjectToScript("Layout", layout);
            this.paneTask = LoadHTML(pane);
        }

        public Task LoadHTML(Pane pane)
        {
            this.pane = pane;
            htmlPath = Path.Combine(appPath, "HTML", pane.ToString() + ".html");
            string html = File.ReadAllText(htmlPath);
            NavigateToString(html);
            return Task.CompletedTask;
        }
        public async void RunJS(string jsCode)
        {
            await initTask;
            await paneTask;
            await CoreWebView2.ExecuteScriptAsync($"layout.{jsCode}");
        }
        public void ShowControl()
        {
            controlTaskPane.Visible = !controlTaskPane.Visible;
        }
        public void SetSize(int width)
        {
            controlTaskPane.Width = width;
        }
    }
}
