using System;
using System.Linq;
using System.Xml.Serialization;
using System.Data;
using System.Windows.Forms;
using System.Collections.Generic;
using taskt.UI.Forms;
using taskt.UI.CustomControls;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Loop Commands)循环命令")]
    [Attributes.ClassAttributes.Description("此命令允许您循环Excel数据集")]
    [Attributes.ClassAttributes.UsesDescription("如果要迭代一系列Excel单元格，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令尝试遍历已知的Excel DataSet")]
    public class BeginExcelDatasetLoopCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指明Excel DataSet名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入稍后将用于遍历数据的唯一数据集名称")]
        [Attributes.PropertyAttributes.SampleUsage("**myData**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_DataSetName { get; set; }

        public BeginExcelDatasetLoopCommand_cn()
        {
            this.CommandName = "BeginExcelDataSetLoopCommand";
            // this.SelectionName = "Loop Excel Dataset";
            this.SelectionName = Settings.Default.Loop_Dataset_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender, Core.Script.ScriptAction parentCommand)
        {

            Core.Automation.Commands.BeginExcelDatasetLoopCommand loopCommand = (Core.Automation.Commands.BeginExcelDatasetLoopCommand)parentCommand.ScriptCommand;
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var dataSetVariable = engine.VariableList.Where(f => f.VariableName == v_DataSetName).FirstOrDefault();

            if (dataSetVariable == null)
                throw new Exception("DataSet Name Not Found - " + v_DataSetName);

            if (!engine.VariableList.Any(f => f.VariableName == "Loop.CurrentIndex"))
            {
                engine.VariableList.Add(new Script.ScriptVariable() { VariableName = "Loop.CurrentIndex", VariableValue = "0" });
            }


            DataTable excelTable = (DataTable)dataSetVariable.VariableValue;


            var loopTimes = excelTable.Rows.Count;

            for (int i = 0; i < excelTable.Rows.Count; i++)
            {
                dataSetVariable.CurrentPosition = i;

                (i + 1).ToString().StoreInUserVariable(engine, "Loop.CurrentIndex");

                foreach (var cmd in parentCommand.AdditionalScriptCommands)
                {
                    //bgw.ReportProgress(0, new object[] { loopCommand.LineNumber, "Starting Loop Number " + (i + 1) + "/" + loopTimes + " From Line " + loopCommand.LineNumber });

                    if (engine.IsCancellationPending)
                        return;
                    engine.ExecuteCommand(cmd);
                    // bgw.ReportProgress(0, new object[] { loopCommand.LineNumber, "Finished Loop From Line " + loopCommand.LineNumber });
                }
            }

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_DataSetName", this, editor));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue();
        }
    }
}