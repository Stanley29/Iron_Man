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
    [Attributes.ClassAttributes.Group("(If Commands)如果命令")]
    [Attributes.ClassAttributes.Description("此命令表示If操作的退出点。 所有Begin Ifs都需要。")]
    [Attributes.ClassAttributes.UsesDescription("如果要表示if方案的出口点，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("序列化程序使用此命令表示if的结束点。")]
    public class EndIfCommand_cn : ScriptCommand
    {
        public EndIfCommand_cn()
        {
            this.DefaultPause = 0;
            this.CommandName = "EndIfCommand";
            // this.SelectionName = "End If";
            this.SelectionName = Settings.Default.End_If_cn;
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
            return "End If";
        }
    }
}