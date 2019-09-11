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
    [Attributes.ClassAttributes.Group("(Task Commands)任务命令")]
    [Attributes.ClassAttributes.Description("此命令将停止当前任务。")]
    [Attributes.ClassAttributes.UsesDescription("如果要停止当前运行的任务，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("")]
    public class StopTaskCommand_cn : ScriptCommand
    {


        public StopTaskCommand_cn()
        {
            this.CommandName = "StopTaskCommand";
            // this.SelectionName = "Stop Current Task";
            this.SelectionName = Settings.Default.Stop_Current_Task_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
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