using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Excel_Toolkit
{
    [Guid("2E79EE15-C38A-4262-A29F-FA5535495903")]
    [ProgId("Contoso.ExcelTookit.AddIn")]
    public partial class ThisAddIn
    {
        public Excel.Application app;
        public Layout layout = new Layout();
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            layout.addIn = this;
            layout.app = Application;
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            layout.sql.Close();
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
