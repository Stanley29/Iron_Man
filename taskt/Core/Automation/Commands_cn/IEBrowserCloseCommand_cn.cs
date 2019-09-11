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
    [Attributes.ClassAttributes.Description("此命令允许您关闭关联的IE Web浏览器")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从SHDocVw.dll实现“InternetExplorer”应用程序对象以实现自动化。")]
    public class IEBrowserCloseCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        public string v_InstanceName { get; set; }

        public IEBrowserCloseCommand_cn()
        {
            this.CommandName = "IEBrowserCloseCommand";
            //  this.SelectionName = "Close Browser";
            this.SelectionName = Settings.Default.Close_Browser_cn;
            this.CommandEnabled = true;
            this.v_InstanceName = "default";
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var vInstance = v_InstanceName.ConvertToUserVariable(engine);

            var browserObject = engine.GetAppInstance(vInstance);


            var browserInstance = (SHDocVw.InternetExplorer)browserObject;
            browserInstance.Quit();

            engine.RemoveAppInstance(vInstance);
        }

        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Instance Name: '" + v_InstanceName + "']";
        }
    }

}