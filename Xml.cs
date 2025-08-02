using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Toolkit
{
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
                newGroups[this.groups.Length] = group; 
                this.groups = newGroups;
                return this;
            }
        }
        public struct Group
        {
            public string id;
            public string label;
            IControl[] controls;
            public Group(string id, string label, IControl[] controls)
            {
                this.id = id;
                this.label = label;
                this.controls = controls;
            }
            public Group(string id, string label, params IControl[][] controlss)
            {
                this.id = id;
                this.label = label;
                this.controls = new IControl[controlss.Sum(arr => arr?.Length ?? 0)];
                int offset = 0;
                foreach (var arr in controlss)
                {
                    if (arr == null) continue;
                    Array.Copy(arr, 0, this.controls, offset, arr.Length);
                    offset += arr.Length;
                }
            }
            public Group AddControl(IControl control)
            {
                IControl[] newControls = new IControl[this.controls.Length + 1];
                for (int i = 0; i < this.controls.Length; i++)
                {
                    newControls[i] = this.controls[i];
                }
                controls[this.controls.Length] = control;
                this.controls = newControls;
                return this;
            }
            public Group AddControls(IControl[] controls)
            {
                IControl[] newControls = new IControl[this.controls.Length + controls.Length];
                Array.Copy(this.controls, newControls, this.controls.Length);
                Array.Copy(controls, 0, newControls, this.controls.Length, controls.Length);
                this.controls = newControls;
                return this;
            }
            public string ToXml()
            {
                string xml = $@"<group id=""{id}"" label=""{label}"">";
                foreach (IControl control in this.controls)
                {
                    xml += control.ToXml();
                }
                xml += "</group>";
                return xml;
            }
            public static Group ButtonsGroup(string[] strings, string func, string id, string label)
            {
                IControl[] buttons = strings.Select(item => (IControl)new Button(item, item, func)).ToArray();
                Group group = new Group(id, label, buttons);
                return group;
            }
        }
        public interface IControl
        {
            string ToXml();
        }
        public struct Button : IControl
        {
            public string id;
            public string label;
            public string onAction;
            public string imageMso;
            public string size;
            public Button(string id, string label, string onAction, string size = null, string imageMso = null)
            {
                this.id = id;
                this.label = label;
                this.onAction = onAction;
                this.size = size ?? "small";
                this.imageMso = imageMso;
            }
            public string ToXml()
            {
                if (imageMso != null)
                {
                    return $"<button id=\"{id}\" label=\"{label}\" onAction=\"{onAction}\" size=\"{size}\" imageMso=\"{imageMso}\"/>";
                }
                else
                {
                    return $"<button id=\"{id}\" label=\"{label}\" onAction=\"{onAction}\" />\"";
                }

            }
        }
        public struct Separator : IControl
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
