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
    [Attributes.ClassAttributes.Description("此命令允许您关闭Excel。")]
    [Attributes.ClassAttributes.UsesDescription("如果要关闭打开的Excel实例，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现Excel Interop以实现自动化。")]
    public class ExcelCloseApplicationCommand : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入** Create Excel **命令中指定的唯一实例名称")]
        [Attributes.PropertyAttributes.SampleUsage("**myInstance** or **seleniumInstance**")]
        [Attributes.PropertyAttributes.Remarks("未能输入正确的实例名称或未能首先调用**创建Excel **命令将导致错误")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("指示是否应保存工作簿")]
        [Attributes.PropertyAttributes.InputSpecification("输入TRUE或FALSE值")]
        [Attributes.PropertyAttributes.SampleUsage("'TRUE' or 'FALSE'")]
        [Attributes.PropertyAttributes.Remarks("")]
        public bool v_ExcelSaveOnExit { get; set; }
        public ExcelCloseApplicationCommand()
        {
            this.CommandName = "ExcelCloseApplicationCommand";
            // this.SelectionName = "Close Excel Application";
            this.SelectionName = Settings.Default.Close_Excel_Application_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }
        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var vInstance = v_InstanceName.ConvertToUserVariable(engine);
            var excelObject = engine.GetAppInstance(vInstance);


            Microsoft.Office.Interop.Excel.Application excelInstance = (Microsoft.Office.Interop.Excel.Application)excelObject;


            //check if workbook exists and save
            if (excelInstance.ActiveWorkbook != null)
            {
                excelInstance.ActiveWorkbook.Close(v_ExcelSaveOnExit);
            }

            //close excel
            excelInstance.Quit();

            //remove instance
            engine.RemoveAppInstance(vInstance);

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_ExcelSaveOnExit", this, editor));

            return RenderedControls;

        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Save On Close: " + v_ExcelSaveOnExit + ", Instance Name: '" + v_InstanceName + "']";
        }
    }
}