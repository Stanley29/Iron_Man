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
    [Attributes.ClassAttributes.Group("(Loop Commands)循环命令")]
    [Attributes.ClassAttributes.Description("此命令允许您多次重复操作（循环）。 任何“开始循环”命令都必须具有以下“结束循环”命令。")]
    [Attributes.ClassAttributes.UsesDescription("如果要迭代列表中的每个项目或一系列项目，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令递归调用底层的“BeginLoop”命令以实现自动化。")]
    public class BeginListLoopCommand : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyDescription("请输入要循环的列表变量")]
        [Attributes.PropertyAttributes.InputSpecification("输入包含项目列表的变量")]
        [Attributes.PropertyAttributes.SampleUsage("[vMyList]")]
        [Attributes.PropertyAttributes.Remarks("使用此命令迭代Split命令的结果。")]
        public string v_LoopParameter { get; set; }

        public BeginListLoopCommand()
        {
            this.CommandName = "BeginListLoopCommand";
            //  this.SelectionName = "Loop List";
            this.SelectionName = Settings.Default.Loop_List_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender, Core.Script.ScriptAction parentCommand)
        {
            Core.Automation.Commands.BeginListLoopCommand loopCommand = (Core.Automation.Commands.BeginListLoopCommand)parentCommand.ScriptCommand;
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            int loopTimes;
            Script.ScriptVariable complexVariable = null;


            //get variable by regular name
            complexVariable = engine.VariableList.Where(x => x.VariableName == v_LoopParameter).FirstOrDefault();


            if (!engine.VariableList.Any(f => f.VariableName == "Loop.CurrentIndex"))
            {
                engine.VariableList.Add(new Script.ScriptVariable() { VariableName = "Loop.CurrentIndex", VariableValue = "0" });
            }

            //user may potentially include brackets []
            if (complexVariable == null)
            {
                complexVariable = engine.VariableList.Where(x => x.VariableName.ApplyVariableFormatting() == v_LoopParameter).FirstOrDefault();
            }

            //if still null then throw exception
            if (complexVariable == null)
            {
                throw new Exception("Complex Variable '" + v_LoopParameter + "' or '" + v_LoopParameter.ApplyVariableFormatting() + "' not found. Ensure the variable exists before attempting to modify it.");
            }

            dynamic listToLoop;
            if (complexVariable.VariableValue is List<string>)
            {
             listToLoop = (List<string>)complexVariable.VariableValue;
            }
            else if (complexVariable.VariableValue is List<OpenQA.Selenium.IWebElement>)
            {
              listToLoop = (List<OpenQA.Selenium.IWebElement>)complexVariable.VariableValue;
            }
            else
            {
                throw new Exception("Complex Variable List Type<T> Not Supported");
            }


            loopTimes = listToLoop.Count;


            for (int i = 0; i < loopTimes; i++)
            {
                if (complexVariable != null)
                    complexVariable.CurrentPosition = i;

             

                engine.ReportProgress("Starting Loop Number " + (i + 1) + "/" + loopTimes + " From Line " + loopCommand.LineNumber);

                foreach (var cmd in parentCommand.AdditionalScriptCommands)
                {
                    if (engine.IsCancellationPending)
                        return;

                    (i + 1).ToString().StoreInUserVariable(engine, "Loop.CurrentIndex");

                    engine.ExecuteCommand(cmd);


                    if (engine.CurrentLoopCancelled)
                    {
                        engine.ReportProgress("Exiting Loop From Line " + loopCommand.LineNumber);
                        engine.CurrentLoopCancelled = false;
                        return;
                    }
                }

                engine.ReportProgress("Finished Loop From Line " + loopCommand.LineNumber);
            }
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_LoopParameter", this, editor));
           
            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return "Loop List Variable '" + v_LoopParameter + "'";
        }
    }
}