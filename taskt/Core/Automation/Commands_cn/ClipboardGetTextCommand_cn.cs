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
    [Attributes.ClassAttributes.Description("此命令允许您从剪贴板获取文本。")]
    [Attributes.ClassAttributes.UsesDescription("如果要从剪贴板复制数据并将其应用于变量，请使用此命令。 然后，您可以使用该变量来提取值。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令使用System.Windows.Forms.Clipboard从脚本引擎实现针对VariableList的操作。")]
    public class ClipboardGetTextCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择一个变量来设置剪贴板内容")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_userVariableName { get; set; }

        [XmlIgnore]
        [NonSerialized]
        public ComboBox VariableNameControl;

        public ClipboardGetTextCommand_cn()
        {
            this.CommandName = "ClipboardCommand";
            // this.SelectionName = "Get Clipboard Text";
            this.SelectionName = Settings.Default.Get_Clipboard_Text_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            User32Functions.GetClipboardText().StoreInUserVariable(sender, v_userVariableName);
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create window name helper control
            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_userVariableName", this));
            VariableNameControl = CommandControls.CreateStandardComboboxFor("v_userVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_userVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Get Text From Clipboard and Apply to Variable: " + v_userVariableName + "]";
        }
    }
}
