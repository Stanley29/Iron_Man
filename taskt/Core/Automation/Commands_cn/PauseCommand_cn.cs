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
    [Attributes.ClassAttributes.Group("(Engine Commands)发动机命令")]
    [Attributes.ClassAttributes.Description("此命令将脚本暂停一段指定的时间（以毫秒为单位）。")]
    [Attributes.ClassAttributes.UsesDescription("如果要在特定时间内暂停脚本，请使用此命令。 在指定的时间结束后，脚本将继续执行。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'Thread.Sleep'以实现自动化。")]
    public class PauseCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("暂停的时间（以毫秒为单位）。")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入特定的时间量（以毫秒为单位）（例如，指定8秒，一个将输入8000）或指定包含值的变量。")]
        [Attributes.PropertyAttributes.SampleUsage("**8000** or **[vVariableWaitTime]**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_PauseLength { get; set; }

        public PauseCommand_cn()
        {
            this.CommandName = "PauseCommand";
            // this.SelectionName = "Pause Script";
            this.SelectionName = Settings.Default.Pause_Script_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var userPauseLength = v_PauseLength.ConvertToUserVariable(sender);
            var pauseLength = int.Parse(userPauseLength);
            System.Threading.Thread.Sleep(pauseLength);
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_PauseLength", this, editor));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Wait for " + v_PauseLength + "ms]";
        }
    }
}