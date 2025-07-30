using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
namespace Excel_Toolkit
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        public string[] level1;
        public Layout layout;
        public Ribbon()
        {

        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            System.Diagnostics.Debug.WriteLine($"GetCustomUI called with ribbonID: {ribbonID}");
            layout = Globals.ThisAddIn.layout;
            layout.ribbon = this;
            level1 = layout.level1;
            string buttons = "";
            for (int i = 0; i < level1.Length; i++)
            {
                string label = level1[i];
                buttons += $"<button id=\"{label}\" label=\"{label}\" onAction=\"Level1ButtonClick\"/>\n";
            }
            string xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                 "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"Ribbon_Load\">" +
                 "<ribbon>" +
                 "<tabs>" +
                 "<tab id=\"Toolkit\" label=\"Toolkit\">" +
                 "<group id=\"Group1\" label=\"功能组\">" +
                 buttons +
                 "</group>" +
                 "</tab>" +
                 "</tabs>" +
                 "</ribbon>" +
                 "</customUI>";
            return xml;
        }
        #endregion

        #region 功能区回调
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        public void Level1ButtonClick(Office.IRibbonControl control)
        {
            layout.Level1ButtonClick(control.Id);
        }
        #endregion
    }
}
