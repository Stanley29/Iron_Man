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
    [Attributes.ClassAttributes.Description("此命令允许您保存Excel工作簿。")]
    [Attributes.ClassAttributes.UsesDescription("如果要将工作簿保存到文件，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现Excel Interop以实现自动化。")]
    public class ExcelSaveAsCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入** Create Excel **命令中指定的唯一实例名称")]
        [Attributes.PropertyAttributes.SampleUsage("**myInstance** or **seleniumInstance**")]
        [Attributes.PropertyAttributes.Remarks("未能输入正确的实例名称或未能首先调用**创建Excel **命令将导致错误")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指出该文件的目录")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowFileSelectionHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或选择文件的路径。")]
        [Attributes.PropertyAttributes.SampleUsage("C:\\temp\\myfile.xlsx or [vExcelFilePath]")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_FileName { get; set; }

        public ExcelSaveAsCommand_cn()
        {
            this.CommandName = "ExcelSaveAsCommand";
            // this.SelectionName = "Save Workbook As";
            this.SelectionName = Settings.Default.Save_Workbook_As_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }
        public override void RunCommand(object sender)
        {
            //get engine context
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            //convert variables
            var vInstance = v_InstanceName.ConvertToUserVariable(engine);
            var fileName = v_FileName.ConvertToUserVariable(engine);

            //get excel app object
            var excelObject = engine.GetAppInstance(vInstance);

            //convert object
            Microsoft.Office.Interop.Excel.Application excelInstance = (Microsoft.Office.Interop.Excel.Application)excelObject;

            //overwrite and save
            excelInstance.DisplayAlerts = false;
            excelInstance.ActiveWorkbook.SaveAs(fileName);
            excelInstance.DisplayAlerts = true;

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_FileName", this, editor));

            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Save To '" + v_FileName + "', Instance Name: '" + v_InstanceName + "']";
        }
    }
}