using System;
using System.Collections.Generic;
using System.Windows.Forms;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{




    [Serializable]
    [Attributes.ClassAttributes.Group("(Loop Commands)循环命令")]
    [Attributes.ClassAttributes.Description("此命令允许您连续重复操作。 任何“开始循环”命令都必须具有以下“结束循环”命令。")]
    [Attributes.ClassAttributes.UsesDescription("如果要无数次执行一系列命令，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令递归调用底层的“BeginLoop”命令以实现自动化。")]
    public class BeginContinousLoopCommand_cn : ScriptCommand
    {

        public BeginContinousLoopCommand_cn()
        {
            this.CommandName = "BeginContinousLoopCommand";
            // this.SelectionName = "Loop Continuously";
            this.SelectionName = Settings.Default.Loop_Continuously_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender, Core.Script.ScriptAction parentCommand)
        {
            Core.Automation.Commands.BeginContinousLoopCommand loopCommand = (Core.Automation.Commands.BeginContinousLoopCommand)parentCommand.ScriptCommand;

            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;


            engine.ReportProgress("Starting Continous Loop From Line " + loopCommand.LineNumber);

            while (true)
            {


                foreach (var cmd in parentCommand.AdditionalScriptCommands)
                {
                    if (engine.IsCancellationPending)
                        return;

                    engine.ExecuteCommand(cmd);

                    if (engine.CurrentLoopCancelled)
                    {
                        engine.ReportProgress("Exiting Loop From Line " + loopCommand.LineNumber);
                        engine.CurrentLoopCancelled = false;
                        return;
                    }
                }
            }
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_Comment", this));
            RenderedControls.Add(CommandControls.CreateDefaultInputFor("v_Comment", this, 100, 300));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue();
        }
    }
}