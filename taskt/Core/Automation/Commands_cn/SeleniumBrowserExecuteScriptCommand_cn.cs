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
    [Attributes.ClassAttributes.Description("此命令允许您在Selenium Web浏览器会话中执行脚本。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现Selenium以实现自动化。")]
    public class SeleniumBrowserExecuteScriptCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入脚本代码")]
        public string v_ScriptCode { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("可选 - 提供参数")]
        public string v_Args { get; set; }
        public SeleniumBrowserExecuteScriptCommand_cn()
        {
            this.CommandName = "SeleniumBrowserExecuteScriptCommand";
            // this.SelectionName = "Execute Script";
            this.SelectionName = Settings.Default.Execute_Script_cn;
            this.v_InstanceName = "default";
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }
        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var vInstance = v_InstanceName.ConvertToUserVariable(engine);

            var browserObject = engine.GetAppInstance(vInstance);


            var script = v_ScriptCode.ConvertToUserVariable(sender);
            var args = v_Args.ConvertToUserVariable(sender);
            var seleniumInstance = (OpenQA.Selenium.IWebDriver)browserObject;


            OpenQA.Selenium.IJavaScriptExecutor js = (OpenQA.Selenium.IJavaScriptExecutor)seleniumInstance;


            if (String.IsNullOrEmpty(args))
            {
                js.ExecuteScript(script);
            }
            else
            {
                js.ExecuteScript(script, args);
            }


        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_ScriptCode", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_Args", this, editor));


            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Instance Name: '" + v_InstanceName + "']";
        }
    }
}