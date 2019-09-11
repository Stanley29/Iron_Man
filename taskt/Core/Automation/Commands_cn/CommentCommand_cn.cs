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
    [Attributes.ClassAttributes.Group("(Misc Commands)其他命令")]
    [Attributes.ClassAttributes.Description("此命令允许您向脚本添加内嵌注释。")]
    [Attributes.ClassAttributes.UsesDescription("如果要添加代码注释或文档代码，请使用此命令。 运行脚本时，将解析并显示注释块中变量（例如[vVar]）的使用。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令仅用于视觉目的")]
    public class CommentCommand_cn : ScriptCommand
    {
        public CommentCommand_cn()
        {
            this.CommandName = "CommentCommand";
            //  this.SelectionName = "Add Code Comment";
            this.SelectionName = Settings.Default.Add_Code_Comment_cn;
            this.DisplayForeColor = System.Drawing.Color.ForestGreen;
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
            return "// Comment: " + this.v_Comment;
        }
    }
}