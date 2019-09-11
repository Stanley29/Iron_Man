using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using System.Data;
using System.Windows.Forms;
using taskt.UI.Forms;
using taskt.UI.CustomControls;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Data Commands)数据命令")]
    [Attributes.ClassAttributes.Description("此命令允许您拆分字符串")]
    [Attributes.ClassAttributes.UsesDescription("如果要将单个文本或变量拆分为多个项目，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令使用String.Split方法实现自动化。")]
    public class StringSplitCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择要拆分的变量或文本")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_userVariableName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("输入分隔符（例如，新行的[crLF]，每个字符的[字符]，'，'）")]
        [Attributes.PropertyAttributes.InputSpecification("声明将用于分隔的字符。 [crLF]可用于换行符，[chars]可用于分割每个数字/字母")]
        [Attributes.PropertyAttributes.SampleUsage("[crLF]，[chars]，'，'（逗号 - 没有单引号包装）")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_splitCharacter { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyDescription("请选择包含结果的列表变量")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyConvertToUserVariableName { get; set; }
        public StringSplitCommand_cn()
        {
            this.CommandName = "StringSplitCommand";
            //this.SelectionName = "Split Text";
            this.SelectionName = Settings.Default.Split_Text_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }
        public override void RunCommand(object sender)
        {
            var stringVariable = v_userVariableName.ConvertToUserVariable(sender);

            List<string> splitString;
            if (v_splitCharacter == "[crLF]")
            {
                splitString = stringVariable.Split(new[] { Environment.NewLine }, StringSplitOptions.None).ToList();
            }
            else if (v_splitCharacter == "[chars]")
            {
                splitString = new List<string>();
                var chars = stringVariable.ToCharArray();
                foreach (var c in chars)
                {
                    splitString.Add(c.ToString());
                }

            }
            else
            {
                splitString = stringVariable.Split(new string[] { v_splitCharacter }, StringSplitOptions.None).ToList();
            }

            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
            
            var v_receivingVariable = v_applyConvertToUserVariableName.Replace(engine.engineSettings.VariableStartMarker, "").Replace(engine.engineSettings.VariableEndMarker, "");
            //get complex variable from engine and assign
            var requiredComplexVariable = engine.VariableList.Where(x => x.VariableName == v_receivingVariable).FirstOrDefault();

            if (requiredComplexVariable == null)
            {
                engine.VariableList.Add(new Script.ScriptVariable() { VariableName = v_receivingVariable, CurrentPosition = 0 });
                requiredComplexVariable = engine.VariableList.Where(x => x.VariableName == v_receivingVariable).FirstOrDefault();
            }

            requiredComplexVariable.VariableValue = splitString;
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_userVariableName", this));
            var userVariableName = CommandControls.CreateStandardComboboxFor("v_userVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_userVariableName", this, new Control[] { userVariableName }, editor));
            RenderedControls.Add(userVariableName);


            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_splitCharacter", this, editor));

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyConvertToUserVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyConvertToUserVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyConvertToUserVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Split '" + v_userVariableName + "' by '" + v_splitCharacter + "' and apply to '" + v_applyConvertToUserVariableName + "']";
        }
    }
}