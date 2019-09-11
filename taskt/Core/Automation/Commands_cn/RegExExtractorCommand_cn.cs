using System;
using System.Xml.Serialization;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Windows.Forms;
using taskt.UI.Forms;
using taskt.UI.CustomControls;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Data Commands)数据命令")]
    [Attributes.ClassAttributes.Description("此命令允许您使用RegEx执行高级字符串格式。")]
    [Attributes.ClassAttributes.UsesDescription("如果要从文本或变量执行高级RegEx提取，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从脚本引擎实现针对VariableList的操作。")]
    public class RegExExtractorCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请提供值或变量（例如[vSomeVariable]）")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("选择或提供变量或文本值")]
        [Attributes.PropertyAttributes.SampleUsage("**Hello** or **vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_InputValue { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("输入RegEx Extractor模式")]
        [Attributes.PropertyAttributes.InputSpecification("输入应该用于提取文本的RegEx提取器模式")]
        [Attributes.PropertyAttributes.SampleUsage(@"^([\w\-]+)")]
        [Attributes.PropertyAttributes.Remarks("例如，如果提取器拆分句子中的每个单词，则需要指定所需单词的关联索引。")]
        public string v_RegExExtractor { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("选择匹配组索引")]
        [Attributes.PropertyAttributes.InputSpecification("定义结果的索引")]
        [Attributes.PropertyAttributes.SampleUsage("1")]
        [Attributes.PropertyAttributes.Remarks("提取器会将找到的多个模式拆分为多个索引。 测试检索值或创建更好/更多定义提取器所需的索引。")]
        public string v_MatchGroupIndex { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择变量以接收RegEx结果")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }

        public RegExExtractorCommand_cn()
        {
            this.CommandName = "RegExExtractorCommand";
            //this.SelectionName = "RegEx Extraction";
            this.SelectionName = Settings.Default.RegEx_Extraction_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;

            //apply default
            v_MatchGroupIndex = "0";
        }

        public override void RunCommand(object sender)
        {
            //get variablized strings
            var variableInput = v_InputValue.ConvertToUserVariable(sender);
            var variableExtractorPattern = v_RegExExtractor.ConvertToUserVariable(sender);
            var variableMatchGroup = v_MatchGroupIndex.ConvertToUserVariable(sender);

            //create regex matcher
            Regex regex = new Regex(variableExtractorPattern);
            Match match = regex.Match(variableInput);

            int matchGroup = 0;
            if (!int.TryParse(variableMatchGroup, out matchGroup))
            {
                matchGroup = 0;
            }

            if (!match.Success)
            {
                //throw exception if no match found
                throw new Exception("RegEx Match was not found! Input: " + variableInput + ", Pattern: " + variableExtractorPattern);
            }
            else
            {
                //store string in variable
                string matchedValue = match.Groups[matchGroup].Value;
                matchedValue.StoreInUserVariable(sender, v_applyToVariableName);
            }



        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InputValue", this, editor));

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_RegExExtractor", this, editor));

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_MatchGroupIndex", this, editor));


            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Apply Extracted Text To Variable: " + v_applyToVariableName + "]";
        }
    }
}