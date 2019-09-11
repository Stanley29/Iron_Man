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
    [Attributes.ClassAttributes.Group("(Data Commands)数据命令")]
    [Attributes.ClassAttributes.Description("此命令允许您修剪字符串")]
    [Attributes.ClassAttributes.UsesDescription("如果要选择文本或变量的子集，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令使用String.Substring方法实现自动化。")]
    public class ModifyVariableCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择要修改的变量或文本")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_userVariableName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("选择案例类型")]
        [Attributes.PropertyAttributes.InputSpecification("指示是否应该保留这么多字符")]
        [Attributes.PropertyAttributes.SampleUsage("-1保持余数，1开始索引后1位置等。")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("对大写")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("降低案例")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("到Base64字符串")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("来自Base64字符串")]
        public string v_ConvertType { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择变量以接收更改")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }
        public ModifyVariableCommand_cn()
        {
            this.CommandName = "ModifyVariableCommand";
            //this.SelectionName = "Modify Variable";
            this.SelectionName = Settings.Default.Modify_Variable_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
;
        }
        public override void RunCommand(object sender)
        {


            var stringValue = v_userVariableName.ConvertToUserVariable(sender);

            var caseType = v_ConvertType.ConvertToUserVariable(sender);

            switch (caseType)
            {
                case "To Upper Case":
                    stringValue = stringValue.ToUpper();
                    break;
                case "To Lower Case":
                    stringValue = stringValue.ToLower();
                    break;
                case "To Base64 String":
                    byte[] textAsBytes = System.Text.Encoding.ASCII.GetBytes(stringValue);
                    stringValue = Convert.ToBase64String(textAsBytes);
                    break;
                case "From Base64 String":
                    byte[] encodedDataAsBytes = System.Convert.FromBase64String(stringValue);
                    stringValue = System.Text.ASCIIEncoding.ASCII.GetString(encodedDataAsBytes);
                    break;
                default:
                    throw new NotImplementedException("Conversion Type '" + caseType + "' not implemented!");
            }

            stringValue.StoreInUserVariable(sender, v_applyToVariableName);
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_userVariableName", this));
            var userVariableName = CommandControls.CreateStandardComboboxFor("v_userVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_userVariableName", this, new Control[] { userVariableName }, editor));
            RenderedControls.Add(userVariableName);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_ConvertType", this, editor));


            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Convert '" + v_userVariableName + "' " + v_ConvertType + "']";
        }
    }
}