using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace Excel_Toolkit
{
    public class WebView : WebView2
    {
        public enum Pane
        {
            Name,
            Query,
            PowerQuery
        }

        public string appPath = AppDomain.CurrentDomain.BaseDirectory;
        public UserControl control;
        public CustomTaskPane controlTaskPane;
        public string htmlPath;
        public Task initTask;
        public Layout layout = Globals.ThisAddIn.layout;
        public Pane pane;
        public Task paneTask;

        public WebView(Pane pane)
        {
            layout.webView = this;
            control = new UserControl();
            // ReSharper disable once VirtualMemberCallInConstructor
            Dock = DockStyle.Fill;
            controlTaskPane = layout.addIn.CustomTaskPanes.Add(control, "Excel Toolkit");
            control.Controls.Add(this);
            SetSize(1200);
            initTask = InitWebView(pane);
            controlTaskPane.Visible = false;
        }

        public new bool Visible
        {
            get => controlTaskPane.Visible;
            set => controlTaskPane.Visible = value;
        }

        private async Task InitWebView(Pane pane)
        {
            this.pane = pane;
            var env = await CoreWebView2Environment.CreateAsync(null, @"C:\temp\MyWebView2");
            await EnsureCoreWebView2Async(env);
            paneTask = LoadHtml(pane);
        }

        public Task LoadHtml(Pane pane)
        {
            CoreWebView2.AddHostObjectToScript("Layout", layout);
            this.pane = pane;
            htmlPath = Path.Combine(appPath, "HTML", pane + ".html");
            var html = File.ReadAllText(htmlPath);
            NavigateToString(html);
            return Task.CompletedTask;
        }

        public async Task RunJavaScript(string jsCode)
        {
            await initTask;
            await paneTask;
            await CoreWebView2.ExecuteScriptAsync($"layout.{jsCode}");
        }

        public void SetSize(int width)
        {
            controlTaskPane.Width = width;
        }
    }
}