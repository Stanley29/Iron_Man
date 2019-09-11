using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.Core.Automation.User32;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Misc Commands)其他命令")]
    [Attributes.ClassAttributes.Description("此命令允许您将文本设置到剪贴板。")]
    [Attributes.ClassAttributes.UsesDescription("如果要从剪贴板复制数据并将其应用于变量，请使用此命令。 然后，您可以使用该变量来提取值。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令使用System.Windows.Forms.Clipboard从脚本引擎实现针对VariableList的操作。")]
    public class ClipboardSetTextCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择目标变量或输入值")]
        [Attributes.PropertyAttributes.InputSpecification("选择变量或提供输入值")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InputValue { get; set; }

        [XmlIgnore]
        [NonSerialized]
        public ComboBox VariableNameControl;

        public ClipboardSetTextCommand_cn()
        {
            this.CommandName = "ClipboardSetTextCommand";
            //this.SelectionName = "Set Clipboard Text";
            this.SelectionName = Settings.Default.Set_Clipboard_Text_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var input = v_InputValue.ConvertToUserVariable(sender);
            User32Functions.SetClipboardText(input);
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create window name helper control
            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_InputValue", this));
            VariableNameControl = CommandControls.CreateStandardComboboxFor("v_InputValue", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_InputValue", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Apply '" + v_InputValue + "' to Clipboard]";
        }
    }
}
