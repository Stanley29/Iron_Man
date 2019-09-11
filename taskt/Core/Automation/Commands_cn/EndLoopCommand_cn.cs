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
    [Attributes.ClassAttributes.Description("此命令表示循环（重复）操作的退出点。 所有循环都需要。")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令表示循环命令的结束点。")]
    [Attributes.ClassAttributes.ImplementationDescription("序列化程序使用此命令表示循环的结束点。")]
    public class EndLoopCommand_cn : ScriptCommand
    {
        public EndLoopCommand_cn()
        {
            this.DefaultPause = 0;
            this.CommandName = "EndLoopCommand";
            // this.SelectionName = "End Loop";
            this.SelectionName = Settings.Default.End_Loop_cn;
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
            return "End Loop";
        }
    }
}