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
    [Attributes.ClassAttributes.Description("此命令将打开Excel工作簿。")]
    [Attributes.ClassAttributes.UsesDescription("如果要打开现有Excel工作簿，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现Excel Interop以实现自动化。")]
    public class ExcelOpenWorkbookCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入** Create Excel **命令中指定的唯一实例名称")]
        [Attributes.PropertyAttributes.SampleUsage("**myInstance** or **seleniumInstance**")]
        [Attributes.PropertyAttributes.Remarks("未能输入正确的实例名称或未能首先调用**创建Excel **命令将导致错误")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指明工作簿文件路径")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowFileSelectionHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或选择应由Excel打开的适用文件的路径。")]
        [Attributes.PropertyAttributes.SampleUsage(@"C:\temp\myfile.xlsx or [vFilePath]")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_FilePath { get; set; }
        public ExcelOpenWorkbookCommand_cn()
        {
            this.CommandName = "ExcelOpenWorkbookCommand";
            // this.SelectionName = "Open Workbook";
            this.SelectionName = Settings.Default.Open_Workbook_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }
        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
            var vInstance = v_InstanceName.ConvertToUserVariable(engine);
            var vFilePath = v_FilePath.ConvertToUserVariable(sender);

           var excelObject = engine.GetAppInstance(vInstance);
            Microsoft.Office.Interop.Excel.Application excelInstance = (Microsoft.Office.Interop.Excel.Application)excelObject;
            excelInstance.Workbooks.Open(vFilePath);

           
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_FilePath", this, editor));

            return RenderedControls;

        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Open from '" + v_FilePath + "', Instance Name: '" + v_InstanceName + "']";
        }
    }
}