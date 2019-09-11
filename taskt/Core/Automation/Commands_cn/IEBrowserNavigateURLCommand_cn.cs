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
    [Attributes.ClassAttributes.Description("此命令允许您将IE Web浏览器会话导航到给定的URL或资源。")]
    [Attributes.ClassAttributes.UsesDescription("如果要将现有IE实例导航到已知URL或Web资源，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从SHDocVw.dll实现“InternetExplorer”应用程序对象以实现自动化。")]
    public class IEBrowserNavigateURLCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入** Create IE Browser **命令中指定的唯一实例名称")]
        [Attributes.PropertyAttributes.SampleUsage("**myInstance** or **IEInstance**")]
        [Attributes.PropertyAttributes.Remarks("未能输入正确的实例名称或未能首次调用**创建IE浏览器**命令将导致错误")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入要导航到的URL")]
        [Attributes.PropertyAttributes.InputSpecification("输入您希望IE实例导航到的目标URL")]
        [Attributes.PropertyAttributes.SampleUsage("https://mycompany.com/orders")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_URL { get; set; }

        public IEBrowserNavigateURLCommand_cn()
        {
            this.CommandName = "IEBrowserNavigateURLCommand";
            // this.SelectionName = "Navigate to URL";
            this.SelectionName = Settings.Default.Navigate_to_URL_cn;
            this.v_InstanceName = "default";
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            object browserObject = null;

            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var vInstance = v_InstanceName.ConvertToUserVariable(engine);

            browserObject = engine.GetAppInstance(vInstance);

            var browserInstance = (SHDocVw.InternetExplorer)browserObject;

            browserInstance.Navigate(v_URL.ConvertToUserVariable(sender));

            IEBrowserCreateCommand.WaitForReadyState(browserInstance);

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_URL", this, editor));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [URL: '" + v_URL + "', Instance Name: '" + v_InstanceName + "']";
        }
    }
}