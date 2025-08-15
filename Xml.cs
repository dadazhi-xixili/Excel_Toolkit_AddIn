using System;
using System.Linq;
using System.Text;

namespace Excel_Toolkit
{
    public class Xml
    {
        public string footer = "</tabs></ribbon></customUI>";

        public string header =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"Ribbon_Load\"><ribbon><tabs>";

        public Tab[] tabs;

        public Xml(Tab[] tabs)
        {
            this.tabs = tabs;
        }

        public Xml(Tab tab)
        {
            tabs = new[] { tab };
        }

        public string ToXml()
        {
            StringBuilder builder = new StringBuilder(2048*tabs.Length);
            builder.Append(header);
            foreach (Tab tab in tabs) builder.Append(tab.ToXml());
            builder.Append(footer);
            return builder.ToString();
        }

        public struct Tab
        {
            public string id;
            public string label;
            public Group[] groups;

            public Tab(string id, string label, Group[] groups)
            {
                this.id = id;
                this.label = label;
                this.groups = groups;
            }

            public string ToXml()
            {
                StringBuilder builder = new StringBuilder(512*groups.Length);
                builder.Append($@"<tab id=""{id}"" label=""{label}"">");
                foreach (Group group in groups) builder.Append(group.ToXml());
                builder.Append("</tab>");
                return builder.ToString();
            }

            public Tab AddGroup(Group group)
            {
                var newGroups = new Group[groups.Length + 1];
                for (var i = 0; i < groups.Length; i++) newGroups[i] = groups[i];
                newGroups[groups.Length] = group;
                groups = newGroups;
                return this;
            }
        }

        public struct Group
        {
            public string id;
            public string label;
            public IControl[] controls;

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
                controls = new IControl[controlss.Sum(arr => arr?.Length ?? 0)];
                var offset = 0;
                foreach (var arr in controlss)
                {
                    if (arr == null) continue;
                    Array.Copy(arr, 0, controls, offset, arr.Length);
                    offset += arr.Length;
                }
            }

            public Group(string id, string label,IControl control)
            {
                this.id = id;
                this.label = label;
                this.controls = new [] { control };
            }

            public Group AddControl(IControl control)
            {
                var newControls = new IControl[controls.Length + 1];
                for (var i = 0; i < controls.Length; i++) newControls[i] = controls[i];
                controls[controls.Length] = control;
                controls = newControls;
                return this;
            }

            public Group AddControls(IControl[] controls)
            {
                var newControls = new IControl[this.controls.Length + controls.Length];
                Array.Copy(this.controls, newControls, this.controls.Length);
                Array.Copy(controls, 0, newControls, this.controls.Length, controls.Length);
                this.controls = newControls;
                return this;
            }

            public string ToXml()
            {
                StringBuilder builder = new StringBuilder(512);
                builder.Append($@"<group id=""{id}"" label=""{label}"">");
                foreach (IControl control in controls) builder.Append(control.ToXml());
                builder.Append("</group>");
                return builder.ToString();
            }

            public static Group ButtonsGroupFromArr(string[] strings, string func, string id, string label)
            {
                var buttons = strings.Select(item => (IControl)new Button(item, item, func)).ToArray();
                var group = new Group(id, label, buttons);
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
                    return
                        $"<button id=\"{id}\" label=\"{label}\" onAction=\"{onAction}\" size=\"{size}\" imageMso=\"{imageMso}\"/>";
                return $"<button id=\"{id}\" label=\"{label}\" onAction=\"{onAction}\" />";
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