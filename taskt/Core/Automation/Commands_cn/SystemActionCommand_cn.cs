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
    [Attributes.ClassAttributes.Description("此命令允许您执行帐户操作")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令执行注销，重新启动，关闭或重新启动等操作。")]
    [Attributes.ClassAttributes.ImplementationDescription("")]
    public class SystemActionCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("选择要执行的系统操作")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从其中一个选项中选择")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_ActionName { get; set; }

        [XmlIgnore]
        [NonSerialized]
        public ComboBox ActionNameComboBox;

        public SystemActionCommand_cn()
        {
            this.CommandName = "SystemActionCommand";
            //this.SelectionName = "System Action";
            this.SelectionName = Settings.Default.System_Action_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var action = (string)v_ActionName.ConvertToUserVariable(sender);
            switch (action)
            {
                case "Shutdown":
                    System.Diagnostics.Process.Start("shutdown", "/s /t 0");
                    break;
                case "Restart":
                    System.Diagnostics.Process.Start("shutdown", "/r /t 0");
                    break;
                case "Logoff":
                    User32.User32Functions.WindowsLogOff();
                    break;
                case "Lock Screen":
                    User32.User32Functions.LockWorkStation();
                    break;
                default:
                    break;
            }
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            var ActionNameComboBoxLabel = CommandControls.CreateDefaultLabelFor("v_ActionName", this);
            ActionNameComboBox = (ComboBox)CommandControls.CreateDropdownFor("v_ActionName", this);
            ActionNameComboBox.DataSource = new List<string> { "Shutdown", "Restart", "Lock Screen", "Logoff" };
            RenderedControls.Add(ActionNameComboBoxLabel);
            RenderedControls.Add(ActionNameComboBox);


            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Action: " + v_ActionName + "]";
        }
    }
}