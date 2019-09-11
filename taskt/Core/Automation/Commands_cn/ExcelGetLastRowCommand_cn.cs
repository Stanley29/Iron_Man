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
    [Attributes.ClassAttributes.Group("(Excel Commands)Excel命令")]
    [Attributes.ClassAttributes.Description("此命令允许您在Excel工作簿中查找已使用范围中的最后一行。")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令可确定Excel工作簿中已使用的行数。 您可以在** Number Of Times **循环中使用此值来获取数据。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现Excel Interop以实现自动化。")]
    public class ExcelGetLastRowCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入** Create Excel **命令中指定的唯一实例名称")]
        [Attributes.PropertyAttributes.SampleUsage("**myInstance** or **seleniumInstance**")]
        [Attributes.PropertyAttributes.Remarks("未能输入正确的实例名称或未能首先调用**创建Excel **命令将导致错误")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入列的字母进行检查（例如A，B，C）")]
        [Attributes.PropertyAttributes.InputSpecification("输入有效的列字母")]
        [Attributes.PropertyAttributes.SampleUsage("A, B, AA, etc.")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_ColumnLetter { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择变量以接收行号")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }
        public ExcelGetLastRowCommand_cn()
        {
            this.CommandName = "ExcelGetLastRowCommand";
            // this.SelectionName = "Get Last Row Index";
            this.SelectionName = Settings.Default.Get_Last_Row_Index_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }
        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
            var vInstance = v_InstanceName.ConvertToUserVariable(engine);

            var excelObject = engine.GetAppInstance(vInstance);

            Microsoft.Office.Interop.Excel.Application excelInstance = (Microsoft.Office.Interop.Excel.Application)excelObject;
                var excelSheet = excelInstance.ActiveSheet;
                var lastRow = (int)excelSheet.Cells(excelSheet.Rows.Count, "A").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row;


                lastRow.ToString().StoreInUserVariable(sender, v_applyToVariableName);


            
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_ColumnLetter", this, editor));

            //create control for variable name
            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);


            return RenderedControls;

        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Column '" + v_ColumnLetter + "', Apply to '" + v_applyToVariableName + "', Instance Name: '" + v_InstanceName + "']";
        }
    }
}