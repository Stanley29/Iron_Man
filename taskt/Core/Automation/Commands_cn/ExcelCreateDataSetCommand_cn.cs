using System;
using System.Xml.Serialization;
using System.Data;
using System.Collections.Generic;
using System.Windows.Forms;
using taskt.UI.Forms;
using taskt.UI.CustomControls;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Excel Commands)Excel命令")]
    [Attributes.ClassAttributes.Description("此命令获取一系列单元格并将其应用于数据集")]
    [Attributes.ClassAttributes.UsesDescription("如果要快速迭代Excel作为数据集，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'OLEDB'以实现自动化。")]
    public class ExcelCreateDataSetCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请创建一个DataSet名称")]
        [Attributes.PropertyAttributes.InputSpecification("指明唯一的参考名称供以后使用")]
        [Attributes.PropertyAttributes.SampleUsage("vMyDataset")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_DataSetName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指明工作簿文件路径")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowFileSelectionHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或选择工作簿文件的路径")]
        [Attributes.PropertyAttributes.SampleUsage(@"C:\temp\myfile.xlsx")]
        [Attributes.PropertyAttributes.Remarks("此命令不需要打开Excel。 将运行此命令时存在的工作簿快照。")]
        public string v_FilePath { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请注明表格名称")]
        [Attributes.PropertyAttributes.InputSpecification("指示应检索的特定工作表。")]
        [Attributes.PropertyAttributes.SampleUsage("Sheet1, mySheet, [vSheet]")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_SheetName { get; set; }

        public ExcelCreateDataSetCommand_cn()
        {
            this.CommandName = "ExcelCreateDatasetCommand";
            //  this.SelectionName = "Create Dataset";
            this.SelectionName = Settings.Default.Create_Dataset_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {

            DatasetCommands dataSetCommand = new DatasetCommands();
            DataTable requiredData = dataSetCommand.CreateDataTable(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + v_FilePath + @";Extended Properties=""Excel 12.0;HDR=No;IMEX=1""", "Select * From [" + v_SheetName + "$]");

            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            Script.ScriptVariable newDataset = new Script.ScriptVariable
            {
                VariableName = v_DataSetName,
                VariableValue = requiredData
            };

            engine.VariableList.Add(newDataset);

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_DataSetName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_FilePath", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SheetName", this, editor));

            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Get '" + v_SheetName + "' from '" + v_FilePath + "' and apply to '" + v_DataSetName + "']";
        }
    }
}