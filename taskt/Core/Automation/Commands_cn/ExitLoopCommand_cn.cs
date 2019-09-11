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
    [Attributes.ClassAttributes.Description("此命令表示当前循环应该退出并恢复当前循环点之后的工作。")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令表示循环应该结束，循环外的命令应该继续执行。")]
    [Attributes.ClassAttributes.ImplementationDescription("引擎使用此命令退出循环")]
    public class ExitLoopCommand_cn : ScriptCommand
    {
        public ExitLoopCommand_cn()
        {
            this.DefaultPause = 0;
            this.CommandName = "ExitLoopCommand";
            // this.SelectionName = "Exit Loop";
            this.SelectionName = Settings.Default.Exit_Loop_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
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
            return "Exit Loop";
        }
    }
}