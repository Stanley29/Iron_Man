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
    [Attributes.ClassAttributes.Description("此命令允许您构建日期并将其应用于变量。")]
    [Attributes.ClassAttributes.UsesDescription("如果要执行日期计算，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从脚本引擎实现针对VariableList的操作。")]
    public class DateCalculationCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请提供日期值或变量（例如[DateTime.Now]")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("指定包含开始日期的文本或变量。")]
        [Attributes.PropertyAttributes.SampleUsage("[DateTime.Now] or 1/1/2000")]
        [Attributes.PropertyAttributes.Remarks("您可以使用已知的文本或变量。")]
        public string v_InputValue { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择一种计算方法")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("添加秒")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("添加分钟")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("添加时间")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("添加天数")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("添加年份")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("减去秒数")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("减去分钟数")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("减去小时数")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("减去天数")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("减去年数")]
        [Attributes.PropertyAttributes.InputSpecification("选择必要的操作")]
        [Attributes.PropertyAttributes.SampleUsage("选择从添加秒数，添加分钟数，添加小时数，添加天数，添加年份，减去秒数，减去分钟数，减去小时数，减去天数，减去年数")]
        public string v_CalculationMethod { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请提供增量值")]
        [Attributes.PropertyAttributes.InputSpecification("输入要增加的单位数")]
        [Attributes.PropertyAttributes.SampleUsage("15, [vIncrement]")]
        [Attributes.PropertyAttributes.Remarks("您可以使用相反的负数，例如。 减去天数，增量为-5将添加天数。")]
        public string v_Increment { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("可选 - 指定字符串格式")]
        [Attributes.PropertyAttributes.InputSpecification("指定是否需要特定的字符串格式。")]
        [Attributes.PropertyAttributes.SampleUsage("MM/dd/yy, hh:mm, etc.")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_ToStringFormat { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择变量以接收日期计算")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }

        public DateCalculationCommand_cn()
        {
            this.CommandName = "DateCalculationCommand";
            //this.SelectionName = "Date Calculation";
            this.SelectionName = Settings.Default.Date_Calculation_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;

            this.v_InputValue = "{DateTime.Now}";
            this.v_ToStringFormat = "MM/dd/yyyy hh:mm:ss";

        }

        public override void RunCommand(object sender)
        {
            //get variablized string
            var variableDateTime = v_InputValue.ConvertToUserVariable(sender);

            //convert to date time
            DateTime requiredDateTime;
            if (!DateTime.TryParse(variableDateTime, out requiredDateTime))
            {
                throw new InvalidDataException("Date was unable to be parsed - " + variableDateTime);
            }

            //get increment value
            double requiredInterval;
            var variableIncrement = v_Increment.ConvertToUserVariable(sender);

            //convert to double
            if (!Double.TryParse(variableIncrement, out requiredInterval))
            {
                throw new InvalidDataException("Date was unable to be parsed - " + variableIncrement);
            }

            //perform operation
            switch (v_CalculationMethod)
            {
                case "Add Seconds":
                    requiredDateTime = requiredDateTime.AddSeconds(requiredInterval);
                    break;
                case "Add Minutes":
                    requiredDateTime = requiredDateTime.AddMinutes(requiredInterval);
                    break;
                case "Add Hours":
                    requiredDateTime = requiredDateTime.AddHours(requiredInterval);
                    break;
                case "Add Days":
                    requiredDateTime = requiredDateTime.AddDays(requiredInterval);
                    break;
                case "Add Years":
                    requiredDateTime = requiredDateTime.AddYears((int)requiredInterval);
                    break;
                case "Subtract Seconds":
                    requiredDateTime = requiredDateTime.AddSeconds((requiredInterval * -1));
                    break;
                case "Subtract Minutes":
                    requiredDateTime = requiredDateTime.AddMinutes((requiredInterval * -1));
                    break;
                case "Subtract Hours":
                    requiredDateTime = requiredDateTime.AddHours((requiredInterval * -1));
                    break;
                case "Subtract Days":
                    requiredDateTime = requiredDateTime.AddDays((requiredInterval * -1));
                    break;
                case "Subtract Years":
                    requiredDateTime = requiredDateTime.AddYears(((int)requiredInterval * -1));
                    break;
                default:
                    break;
            }

            //handle if formatter is required     
            var formatting = v_ToStringFormat.ConvertToUserVariable(sender);
            var stringDateFormatted = requiredDateTime.ToString(formatting);


            //store string in variable
            stringDateFormatted.StoreInUserVariable(sender, v_applyToVariableName);

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InputValue", this, editor));

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_CalculationMethod", this));
            RenderedControls.Add(CommandControls.CreateDropdownFor("v_CalculationMethod", this));

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_Increment", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_ToStringFormat", this, editor));

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);


            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            //if calculation method was selected
            if (v_CalculationMethod != null)
            {
                //determine operand and interval
                var operand = v_CalculationMethod.Split(' ')[0];
                var interval = v_CalculationMethod.Split(' ')[1];

                //additional language handling based on selection made
                string operandLanguage;
                if (operand == "Add")
                {
                    operandLanguage = " to ";
                }
                else
                {
                    operandLanguage = " from ";
                }

                //return value
                return base.GetDisplayValue() + " [" + operand + " " + v_Increment + " " + interval + operandLanguage + v_InputValue + ", Apply Result to Variable '" + v_applyToVariableName + "']";
            }
            else
            {
                return base.GetDisplayValue();
            }

        }
    }
}