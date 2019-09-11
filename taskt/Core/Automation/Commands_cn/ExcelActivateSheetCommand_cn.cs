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
    [Attributes.ClassAttributes.Description("此命令允许您激活工作簿中的特定工作表")]
    [Attributes.ClassAttributes.UsesDescription("如果要切换到特定工作表，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现Excel Interop以实现自动化。")]
    public class ExcelActivateSheetCommand : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入** Create Excel **命令中指定的唯一实例名称")]
        [Attributes.PropertyAttributes.SampleUsage("**myInstance** or **seleniumInstance**")]
        [Attributes.PropertyAttributes.Remarks("未能输入正确的实例名称或未能首先调用**创建Excel **命令将导致错误")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("指示要激活的工作簿中工作表的名称")]
        [Attributes.PropertyAttributes.InputSpecification("指定实际工作表的名称")]
        [Attributes.PropertyAttributes.SampleUsage("Sheet1, mySheetName, [vSheet]")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_SheetName { get; set; }
        public ExcelActivateSheetCommand()
        {
            this.CommandName = "ExcelActivateSheetCommand";
            // this.SelectionName = "Activate Sheet";
            this.SelectionName = Settings.Default.Activate_Sheet_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }
        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
            var vInstance = v_InstanceName.ConvertToUserVariable(engine);

            var excelObject = engine.GetAppInstance(vInstance);

            Microsoft.Office.Interop.Excel.Application excelInstance = (Microsoft.Office.Interop.Excel.Application)excelObject;
                string sheetToDelete = v_SheetName.ConvertToUserVariable(sender);
                Microsoft.Office.Interop.Excel.Worksheet workSheet = excelInstance.Sheets[sheetToDelete];
                workSheet.Select();



            
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SheetName", this, editor));


            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Sheet Name: " + v_SheetName + ", Instance Name: '" + v_InstanceName + "']";
        }
    }
}