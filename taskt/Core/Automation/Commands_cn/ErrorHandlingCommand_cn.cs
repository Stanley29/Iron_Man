using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;
using taskt.Core.Automation.Commands_cn;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Engine Commands)发动机命令")]
    [Attributes.ClassAttributes.Description("此命令指定遇到错误后要执行的操作。")]
    [Attributes.ClassAttributes.UsesDescription("如果要定义遇到错误时脚本的行为方式，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'Thread.Sleep'以实现自动化。")]
    public class ErrorHandlingCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("选择要在发生错误时执行的操作")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("停止处理")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("继续处理")]
        [Attributes.PropertyAttributes.InputSpecification("遇到错误时，请选择要执行的操作。")]
        [Attributes.PropertyAttributes.SampleUsage("**停止处理**以在遇到错误时结束脚本或**继续处理**以继续运行脚本")]
        [Attributes.PropertyAttributes.Remarks("**如果Command **允许您指定并测试行号是否遇到错误。 要使用该功能，您必须指定**继续处理**")]
        public string v_ErrorHandlingAction { get; set; }

        public ErrorHandlingCommand_cn()
        {
            this.CommandName = "ErrorHandlingCommand";
            //this.SelectionName = "Error Handling";
            this.SelectionName = Settings.Default.Error_handling_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
            engine.ErrorHandler = this;
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);


            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_ErrorHandlingAction", this));
            var dropdown = CommandControls.CreateDropdownFor("v_ErrorHandlingAction", this);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_ErrorHandlingAction", this, new Control[] { dropdown }, editor));
            RenderedControls.Add(dropdown);

            return RenderedControls;
        }


        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Action: " + v_ErrorHandlingAction + "]";
        }
    }
}