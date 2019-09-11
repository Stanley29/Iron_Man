using System;
using System.Xml.Serialization;
using System.IO;
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
    [Attributes.ClassAttributes.Description("此命令允许您将格式应用于字符串")]
    [Attributes.ClassAttributes.UsesDescription("如果要将特定格式应用于文本或变量，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从脚本引擎实现针对VariableList的操作。")]
    public class FormatDataCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请提供值或变量（例如[DateTime.Now]")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("指定包含需要格式化的日期或数字的文本或变量")]
        [Attributes.PropertyAttributes.SampleUsage("[DateTime.Now], 1/1/2000, 2500")]
        [Attributes.PropertyAttributes.Remarks("您可以使用已知的文本或变量。")]
        public string v_InputValue { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择数据类型")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("日期")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("数")]
        [Attributes.PropertyAttributes.InputSpecification("指明源类型")]
        [Attributes.PropertyAttributes.SampleUsage("选择**日期**或**数字**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_FormatType { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("指定所需的输出格式")]
        [Attributes.PropertyAttributes.InputSpecification("指定是否需要特定的字符串格式。")]
        [Attributes.PropertyAttributes.SampleUsage("MM/dd/yy, hh:mm, C2, D2, etc.")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_ToStringFormat { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择变量以接收输出")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }

        public FormatDataCommand_cn()
        {
            this.CommandName = "FormatDataCommand";
            //this.SelectionName = "Format Data";
            this.SelectionName = Settings.Default.Format_Data_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;

            this.v_InputValue = "{DateTime.Now}";
            this.v_FormatType = "Date";
            this.v_ToStringFormat = "MM/dd/yyyy";
            
        }

        public override void RunCommand(object sender)
        {
            //get variablized string
            var variableString = v_InputValue.ConvertToUserVariable(sender);

            //get formatting
            var formatting = v_ToStringFormat.ConvertToUserVariable(sender);

            var variableName = v_applyToVariableName.ConvertToUserVariable(sender);


            string formattedString = "";
            switch (v_FormatType)
            {
                case "Date":
                    if (DateTime.TryParse(variableString, out var parsedDate))
                    {
                        formattedString = parsedDate.ToString(formatting);
                    }
                    break;
                case "Number":
                    if (Decimal.TryParse(variableString, out var parsedDecimal))
                    {
                        formattedString = parsedDecimal.ToString(formatting);
                    }
                    break;
                default:
                    throw new Exception("Formatter Type Not Supported: " + v_FormatType);
            }

            if (formattedString == "")
            {
                throw new InvalidDataException("Unable to convert '" + variableString + "' to type '" + v_FormatType + "'");
            }
            else
            {
                formattedString.StoreInUserVariable(sender, v_applyToVariableName);
            }



        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InputValue", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_FormatType", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_ToStringFormat", this, editor));

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Format '" + v_InputValue + "' and Apply Result to Variable '" + v_applyToVariableName + "']";
        }
    }
}