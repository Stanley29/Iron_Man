using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{

    [Serializable]
    [Attributes.ClassAttributes.Group("(Misc Commands)其他命令")]
    [Attributes.ClassAttributes.Description("将多个操作分组的命令")]
    [Attributes.ClassAttributes.UsesDescription("如果要将多个命令组合在一起，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令在列表中实现许多命令。")]
    public class SequenceCommand_cn : ScriptCommand
    {
        public List<ScriptCommand> v_scriptActions = new List<ScriptCommand>();


        public SequenceCommand_cn()
        {
            this.CommandName = "SequenceCommand";
            // this.SelectionName = "Sequence Command";
            this.SelectionName = Settings.Default.Sequence_Command_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender, Core.Script.ScriptAction parentCommand)
        {

            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            foreach (var item in v_scriptActions)
            {

                //exit if cancellation pending
                if (engine.IsCancellationPending)
                {
                    return;
                }

                //only run if not commented
                if (!item.IsCommented)
                    item.RunCommand(sender);



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
            return base.GetDisplayValue() + " [" + v_scriptActions.Count() + " embedded commands]";
        }
    }
}