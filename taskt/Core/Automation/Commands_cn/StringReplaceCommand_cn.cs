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
    [Attributes.ClassAttributes.Description("T此命令允许您替换文本")]
    [Attributes.ClassAttributes.UsesDescription("如果要替换文本中的现有文本或使用新文本替换变量，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令使用String.Substring方法实现自动化。")]
    public class StringReplaceCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择要修改的文本或变量")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_userVariableName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("指示要替换的文本")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入要替换的文本的旧值")]
        [Attributes.PropertyAttributes.SampleUsage("H")]
        [Attributes.PropertyAttributes.Remarks("你好的H将被替换为目标")]
        public string v_replacementText { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("指示替换值")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("更换后输入新值")]
        [Attributes.PropertyAttributes.SampleUsage("J")]
        [Attributes.PropertyAttributes.Remarks("H将被替换为J来创建'Jello")]
        public string v_replacementValue { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择变量以接收更改")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }
        public StringReplaceCommand_cn()
        {
            this.CommandName = "StringReplaceCommand";
            //this.SelectionName = "Replace Text";
            this.SelectionName = Settings.Default.Replace_Text_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }
        public override void RunCommand(object sender)
        {
            //get full text
            string replacementVariable = v_userVariableName.ConvertToUserVariable(sender);

            //get replacement text and value
            string replacementText = v_replacementText.ConvertToUserVariable(sender);
            string replacementValue = v_replacementValue.ConvertToUserVariable(sender);

            //perform replacement
            replacementVariable = replacementVariable.Replace(replacementText, replacementValue);

            //store in variable
            replacementVariable.StoreInUserVariable(sender, v_applyToVariableName);

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_userVariableName", this));
            var userVariableName = CommandControls.CreateStandardComboboxFor("v_userVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_userVariableName", this, new Control[] { userVariableName }, editor));
            RenderedControls.Add(userVariableName);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_replacementText", this, editor));

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_replacementValue", this, editor));

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Replace '" + v_replacementText + "' with '" + v_replacementValue + "', apply to '" + v_userVariableName + "']";
        }
    }
}