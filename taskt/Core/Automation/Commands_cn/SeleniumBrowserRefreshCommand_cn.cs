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
    [Attributes.ClassAttributes.Group("(Web Browser Commands)Web浏览器命令")]
    [Attributes.ClassAttributes.Description("此命令允许您刷新Selenium Web浏览器会话。")]
    [Attributes.ClassAttributes.UsesDescription("如果要在Web浏览器会话中模拟浏览器刷新单击，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现Selenium以实现自动化。")]
    public class SeleniumBrowserRefreshCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入** Create Browser **命令中指定的唯一实例名称")]
        [Attributes.PropertyAttributes.SampleUsage("**myInstance** or **seleniumInstance**")]
        [Attributes.PropertyAttributes.Remarks("未能输入正确的实例名称或未能首先调用** Create Browser **命令将导致错误")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }

        public SeleniumBrowserRefreshCommand_cn()
        {
            this.CommandName = "SeleniumBrowserRefreshCommand";
            // this.SelectionName = "Refresh";
            this.SelectionName = Settings.Default.Refresh_cn;
            this.v_InstanceName = "default";
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var vInstance = v_InstanceName.ConvertToUserVariable(engine);

            var browserObject = engine.GetAppInstance(vInstance);


            var seleniumInstance = (OpenQA.Selenium.IWebDriver)browserObject;
            seleniumInstance.Navigate().Refresh();

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