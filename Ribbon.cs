using Microsoft.Office.Interop.Excel;
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
            layout = Globals.ThisAddIn.layout;
            layout.ribbon = this;
            level1 = layout.level1;
            Xml.Control[] buttons = level1.Select(label => (Xml.Control)new Xml.Button(label, label, "Level1ButtonClick")).ToArray();
            Xml.Group[] groups = { new Xml.Group("函数分组", "函数分组", buttons) };
            Xml.Tab[] tabs = { new Xml.Tab("Toolkit", "Toolkit", groups) };
            Xml xml = new Xml(tabs);
            return xml.ToXml();
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

        public class Xml
        {
            public string header = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"Ribbon_Load\"><ribbon><tabs>";
            public string footer = "</tabs></ribbon></customUI>";
            public Tab[] tabs;
            public Xml(Tab[] tabs)
            {
                this.tabs = tabs;
            }
            public string ToXml()
            {
                string xml = header;
                for (int i = 0; i < this.tabs.Length; i++)
                {
                    xml += this.tabs[i].ToXml();
                }
                xml += footer;
                return xml;
            }
            public struct Tab
            {
                public string id;
                public string label;
                Group[] groups;
                public Tab(string id, string label, Group[] groups)
                {
                    this.id = id;
                    this.label = label;
                    this.groups = groups;
                }
                public string ToXml()
                {
                    string xml = $@"<tab id=""{id}"" label=""{label}"">";
                    foreach (Group group in this.groups)
                    {
                        xml += group.ToXml();
                    }
                    xml += "</tab>";
                    return xml;
                }
                public Tab AddGroup(Group group)
                {
                    Group[] newGroups = new Group[this.groups.Length + 1];
                    for (int i = 0; i < this.groups.Length; i++)
                    {
                        newGroups[i] = this.groups[i];
                    }
                    groups[this.groups.Length] = group;
                    this.groups = newGroups;
                    return this;
                }
            }
            public struct Group
            {
                public string id;
                public string label;
                Control[] controls;
                public Group(string id, string label, Control[] controls)
                {
                    this.id = id;
                    this.label = label;
                    this.controls = controls;
                }
                public Group AddControl(Control control)
                {
                    Control[] newControls = new Control[this.controls.Length + 1];
                    for (int i = 0; i < this.controls.Length; i++)
                    {
                        newControls[i] = this.controls[i];
                    }
                    controls[this.controls.Length] = control;
                    this.controls = newControls;
                    return this;
                }
                public Group AddControls(Control[] controls)
                {
                    Control[] newControls = new Control[this.controls.Length + controls.Length];
                    Array.Copy(this.controls, newControls, this.controls.Length);
                    Array.Copy(controls, 0, newControls, this.controls.Length, controls.Length);
                    this.controls = newControls;
                    return this;
                }
                public string ToXml()
                {
                    string xml = $@"<group id=""{id}"" label=""{label}"">";
                    foreach (Control control in this.controls)
                    {
                        xml += control.ToXml();
                    }
                    xml += "</group>";
                    return xml;
                }
            }
            public interface Control
            {
                string ToXml();
            }
            public struct Button : Control
            {
                public string id;
                public string label;
                public string onAction;
                public Button(string id, string label, string onAction)
                {
                    this.id = id;
                    this.label = label;
                    this.onAction = onAction;
                }
                public string ToXml()
                {
                    return $"<button id=\"{id}\" label=\"{label}\" onAction=\"{onAction}\" />";
                }
            }
            public struct Separator : Control
            {
                public string id;
                public Separator(string id)
                {
                    this.id = id;
                }
                public string ToXml()
                {
                    return $"<separator id=\"{id}\" />";
                }
            }
        }
    }
}
