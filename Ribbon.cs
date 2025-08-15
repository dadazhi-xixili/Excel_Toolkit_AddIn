using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace Excel_Toolkit
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        public Layout layout;
        public string[] level1;
        public string level1Active;
        public Office.IRibbonUI ribbon;
        public WebView webView;
        public Xml xml;
        public Ribbon()
        {
            
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            if (layout == null)
            {
                layout = Globals.ThisAddIn.layout;
                layout.ribbon = this;
                level1 = layout.GetLevel1();
                var groupLevel1 = Xml.Group.ButtonsGroupFromArr(level1, "QueryClick", "函数分组", "函数分组");
                Xml.IControl nameButton =  new Xml.Button("名称", "名称", "NameClick", "large", "NameDefine") ;
                var groupName = new Xml.Group("名称管理器", "名称管理器", nameButton);
                this.xml = new Xml(new Xml.Tab("Toolkit", "Toolkit", new[] { groupLevel1, groupName }));
            }
            return xml.ToXml();
        }

        #endregion

        #region 功能区回调

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        private void LoadWebView(WebView.Pane pane)
        {
            if (webView != null) return;
            webView = new WebView(pane);
            layout.webView = webView;
        }

        public void QueryClick(Office.IRibbonControl control)
        {
            const WebView.Pane pane = WebView.Pane.Query;
            LoadWebView(pane);
            if (webView.pane != pane) webView.LoadHtml(pane);
            if (control.Id == level1Active)
            {
                webView.Visible = !webView.Visible;
            }
            else
            {
                webView.RunJavaScript($"InitLevel2('{control.Id}')");
                level1Active = control.Id;
                webView.Visible = true;
            }
        }

        public void NameClick(Office.IRibbonControl control)
        {
            const WebView.Pane pane = WebView.Pane.Name;
            level1Active = null;
            LoadWebView(pane);
            if (webView.pane != pane)
            {
                webView.LoadHtml(pane);
                webView.Visible = true;
            }
            else
            {
                webView.Visible = !webView.Visible;
            }
        }

        #endregion
    }
}