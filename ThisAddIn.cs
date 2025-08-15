using System;
using Microsoft.Office.Core;

namespace Excel_Toolkit
{
    //[Guid("6F26193E-E505-4648-9A48-B762990EDDC6")]
    //[ProgId("XXL.ExcelToolkit.AddIn")]
    public partial class ThisAddIn
    {
        public Layout layout = new Layout();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            layout.addIn = this;
            layout.app = Application;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            layout.sql.Close();
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        #region VSTO 生成的代码

        /// <summary>
        ///     设计器支持所需的方法 - 不要修改
        ///     使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}