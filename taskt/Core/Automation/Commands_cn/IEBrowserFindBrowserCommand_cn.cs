using SHDocVw;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(IE Browser Commands)IE浏览器命令")]
    [Attributes.ClassAttributes.Description("此命令允许您查找并附加到现有的IE Web浏览器会话。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从SHDocVw.dll实现“InternetExplorer”应用程序对象以实现自动化。")]
    public class IEBrowserFindBrowserCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        public string v_InstanceName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择或输入浏览器名称")]
        public string v_IEBrowserName { get; set; }

        [XmlIgnore]
        [NonSerialized]
        private ComboBox IEBrowerNameDropdown;

        public IEBrowserFindBrowserCommand_cn()
        {
            this.CommandName = "IEBrowserFindBrowserCommand";
            // this.SelectionName = "Find Browser";
            this.SelectionName = Settings.Default.Find_Browser_cn;
            this.v_InstanceName = "default";
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var instanceName = v_InstanceName.ConvertToUserVariable(sender);

            bool browserFound = false;
            var shellWindows = new ShellWindows();
            foreach (IWebBrowser2 shellWindow in shellWindows)
            {
                if ((shellWindow.Document is MSHTML.HTMLDocument) && (v_IEBrowserName==null || shellWindow.Document.Title == v_IEBrowserName))
                {
                    engine.AddAppInstance(instanceName, shellWindow.Application);
                    browserFound = true;
                    break;
                }
            }

            //try partial match
            if (!browserFound)
            {
                foreach (IWebBrowser2 shellWindow in shellWindows)
                {
                    if ((shellWindow.Document is MSHTML.HTMLDocument) && ((shellWindow.Document.Title.Contains(v_IEBrowserName) || shellWindow.Document.Url.Contains(v_IEBrowserName))))
                    {
                        engine.AddAppInstance(instanceName, shellWindow.Application);
                        browserFound = true;
                        break;
                    }
                }
            }

            if (!browserFound)
            {
                throw new Exception("Browser was not found!");
            }
        }

        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));

            IEBrowerNameDropdown = (ComboBox)CommandControls.CreateDropdownFor("v_IEBrowserName", this);
            var shellWindows = new ShellWindows();
            foreach (IWebBrowser2 shellWindow in shellWindows)
            {
                if (shellWindow.Document is MSHTML.HTMLDocument)
                    IEBrowerNameDropdown.Items.Add(shellWindow.Document.Title);
            }
            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_IEBrowserName", this));
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_IEBrowserName", this, new Control[] { IEBrowerNameDropdown }, editor));
            //IEBrowerNameDropdown.SelectionChangeCommitted += seleniumAction_SelectionChangeCommitted;
            RenderedControls.Add(IEBrowerNameDropdown);

            //ElementParameterControls = new List<Control>();
            //ElementParameterControls.Add(CommandControls.CreateDefaultLabelFor("v_WebActionParameterTable", this));
            //ElementParameterControls.AddRange(CommandControls.CreateUIHelpersFor("v_WebActionParameterTable", this, new Control[] { ElementsGridViewHelper }, editor));
            //ElementParameterControls.Add(ElementsGridViewHelper);

            //RenderedControls.AddRange(ElementParameterControls);

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Browser Name: '" + v_IEBrowserName + "', Instance Name: '" + v_InstanceName + "']";
        }
    }

}