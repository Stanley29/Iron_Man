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
    [Attributes.ClassAttributes.Group("(System Commands)系统命令")]
    [Attributes.ClassAttributes.Description("此命令允许您停止程序或进程。")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令按名称“chrome”关闭应用程序。 或者，您可以使用“关闭窗口”或“厚应用程序命令”。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'Process.CloseMainWindow'。")]
    public class RemoteDesktopCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("输入机器的名称")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_MachineName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("输入用户名")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_UserName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("输入密码")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_Password { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("输入RDP窗口的宽度")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_RDPWidth { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("输入RDP窗口的高度")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_RDPHeight { get; set; }

        public RemoteDesktopCommand_cn()
        {
            this.CommandName = "RemoteDesktopCommand";
            // this.SelectionName = "Launch Remote Desktop";
            this.SelectionName = Settings.Default.Launch_Remote_Desktop_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;

            this.v_RDPWidth = SystemInformation.PrimaryMonitorSize.Width.ToString();
            this.v_RDPHeight = SystemInformation.PrimaryMonitorSize.Height.ToString();

        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var machineName = v_MachineName.ConvertToUserVariable(sender);
            var userName = v_UserName.ConvertToUserVariable(sender);
            var password = v_Password.ConvertToUserVariable(sender);
            var width = int.Parse(v_RDPWidth.ConvertToUserVariable(sender));
            var height = int.Parse(v_RDPHeight.ConvertToUserVariable(sender));


            var result = engine.tasktEngineUI.Invoke(new Action(() =>
            {
                engine.tasktEngineUI.LaunchRDPSession(machineName, userName, password, width, height);
            }));


        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_MachineName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_UserName", this, editor));

            //mask passwords
            var passwordGroup = CommandControls.CreateDefaultInputGroupFor("v_Password", this, editor);
            TextBox inputBox = (TextBox)passwordGroup[2];
            inputBox.PasswordChar = '*';

            RenderedControls.AddRange(passwordGroup);


            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_RDPWidth", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_RDPHeight", this, editor));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Logon " + v_UserName + " to " + v_MachineName + "]";
        }
    }
}