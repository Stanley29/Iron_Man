using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Misc Commands)其他命令")]
    [Attributes.ClassAttributes.Description("此命令允许您向用户显示消息。")]
    [Attributes.ClassAttributes.UsesDescription("如果要在屏幕上向用户显示或显示值，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'MessageBox'并调用VariableCommand以查找变量数据。")]
    public class MessageBoxCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入要显示的消息。")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("指定应在屏幕上显示的任何文本。 您还可以包含用于显示目的的变量。")]
        [Attributes.PropertyAttributes.SampleUsage("**Hello World** or **[vMyText]**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_Message { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("X（秒）后关闭 -  0绕过")]
        [Attributes.PropertyAttributes.InputSpecification("指定在屏幕上显示的秒数。 经过秒数后，消息框将自动关闭，脚本将继续执行。")]
        [Attributes.PropertyAttributes.SampleUsage("** 0 **无限期保持开放或** 5 **保持开放5秒。")]
        [Attributes.PropertyAttributes.Remarks("")]
        public int v_AutoCloseAfter { get; set; }
        public MessageBoxCommand_cn()
        {
            this.CommandName = "MessageBoxCommand";
            // this.SelectionName = "Show Message";
            this.SelectionName = Settings.Default.Show_Message_cn;
            this.CommandEnabled = true;
            this.v_AutoCloseAfter = 0;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
            string variableMessage = v_Message.ConvertToUserVariable(sender);

            variableMessage = variableMessage.Replace("\\n", Environment.NewLine);

            if (engine.tasktEngineUI == null)
            {
                engine.ReportProgress("Complex Messagebox Supported With UI Only");
                System.Windows.Forms.MessageBox.Show(variableMessage, "Message Box Command", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }

            //automatically close messageboxes for server requests
            if (engine.serverExecution && v_AutoCloseAfter <= 0)
            {
                v_AutoCloseAfter = 10;
            }

            var result = engine.tasktEngineUI.Invoke(new Action(() =>
            {
                engine.tasktEngineUI.ShowMessage(variableMessage, "MessageBox Command", UI.Forms.Supplemental.frmDialog.DialogType.OkOnly, v_AutoCloseAfter);
            }

            ));

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create message controls
            var messageControlSet = CommandControls.CreateDefaultInputGroupFor("v_Message", this, editor);
            RenderedControls.AddRange(messageControlSet);


            //create auto close control set
            var autocloseControlSet = CommandControls.CreateDefaultInputGroupFor("v_AutoCloseAfter", this, editor);
            RenderedControls.AddRange(autocloseControlSet);


            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Message: " + v_Message + "]";
        }

        //controls can be overriden and rendered individually
        public System.Windows.Forms.TextBox v_MessageControl()
        {
            var Textbox = new TextBox();
            Textbox.Font = new Font("Segoe UI", 12, FontStyle.Regular);
            Textbox.DataBindings.Add("Text", this, "v_Message", false, DataSourceUpdateMode.OnPropertyChanged);
            Textbox.Height = 30;
            Textbox.Width = 300;

            return Textbox;
        }
    }
}
