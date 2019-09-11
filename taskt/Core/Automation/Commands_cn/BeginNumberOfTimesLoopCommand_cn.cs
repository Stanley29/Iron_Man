using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.Core;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{

    [Serializable]
    [Attributes.ClassAttributes.Group("(Loop Commands)循环命令")]
    [Attributes.ClassAttributes.Description("此命令允许您多次重复操作（循环）。 任何“开始循环”命令都必须具有以下“结束循环”命令。")]
    [Attributes.ClassAttributes.UsesDescription("如果要执行指定次数的一系列命令，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令递归调用底层的“BeginLoop”命令以实现自动化。")]
    public class BeginNumberOfTimesLoopCommand : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyDescription("输入执行循环的次数")]
        [Attributes.PropertyAttributes.InputSpecification("输入您要执行封装命令的次数。")]
        [Attributes.PropertyAttributes.SampleUsage("**5** or **10**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_LoopParameter { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyDescription("可选 - 定义开始索引（默认值：0）")]
        [Attributes.PropertyAttributes.InputSpecification("输入循环的起始索引。")]
        [Attributes.PropertyAttributes.SampleUsage("**5** or **10**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_LoopStart { get; set; }

        public BeginNumberOfTimesLoopCommand()
        {
            this.CommandName = "BeginNumberOfTimesLoopCommand";
            // this.SelectionName = "Loop Number Of Times";
            this.SelectionName = Settings.Default.Loop_Number_Of_Times_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
            this.v_LoopStart = "0";
        }

        public override void RunCommand(object sender, Core.Script.ScriptAction parentCommand)
        {
            Core.Automation.Commands.BeginNumberOfTimesLoopCommand loopCommand = (Core.Automation.Commands.BeginNumberOfTimesLoopCommand)parentCommand.ScriptCommand;

            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            if (!engine.VariableList.Any(f => f.VariableName == "Loop.CurrentIndex"))
            {
                engine.VariableList.Add(new Script.ScriptVariable() { VariableName = "Loop.CurrentIndex", VariableValue = "0" });
            }


            int loopTimes;
            Script.ScriptVariable complexVarible = null;

            var loopParameter = loopCommand.v_LoopParameter.ConvertToUserVariable(sender);

            loopTimes = int.Parse(loopParameter);


            int startIndex = 0;
            int.TryParse(v_LoopStart.ConvertToUserVariable(sender), out startIndex);


            for (int i = startIndex; i < loopTimes; i++)
            {
                if (complexVarible != null)
                    complexVarible.CurrentPosition = i;

              //  (i + 1).ToString().StoreInUserVariable(engine, "Loop.CurrentIndex");

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
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_LoopStart", this, editor));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            if (v_LoopStart != "0")
            {
                return "Loop From (" + v_LoopStart + "+1) to " + v_LoopParameter;

            }
            else
            {
                return "Loop " +  v_LoopParameter + " Times";
            }
         
        }
    }
}