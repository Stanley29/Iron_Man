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
    [Attributes.ClassAttributes.Description("此命令根据“true”或“false”条件声明操作之间的分隔。")]
    [Attributes.ClassAttributes.UsesDescription("如果要表示if方案的出口点，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("TBD")]
    public class ElseCommand_cn : ScriptCommand
    {
        public ElseCommand_cn()
        {
            this.DefaultPause = 0;
            this.CommandName = "ElseCommand";
            //this.SelectionName = "Else";
            this.SelectionName = Settings.Default.Else_cn;
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
            return "Else";
        }
    }
}