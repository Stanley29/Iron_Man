using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Engine Commands)发动机命令")]
    [Attributes.ClassAttributes.Description("此命令允许您设置正在运行的实例中的命令执行之间的延迟。")]
    [Attributes.ClassAttributes.UsesDescription("如果要更改命令之间的执行速度，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("")]
    public class SetEngineDelayCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("设置命令之间的延迟（以毫秒为单位）。")]
        [Attributes.PropertyAttributes.InputSpecification("输入特定的时间量（以毫秒为单位）（例如，指定8秒，一个将输入8000）或指定包含值的变量。")]
        [Attributes.PropertyAttributes.SampleUsage("**250** or **[vVariableSpeed]**")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_EngineSpeed { get; set; }

        public SetEngineDelayCommand_cn()
        {
            this.CommandName = "SetEngineDelayCommand";
            // this.SelectionName = "Set Engine Delay";
            this.SelectionName = Settings.Default.Set_Engine_Delay_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
            this.v_EngineSpeed = "250";
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_EngineSpeed", this, editor));

            return RenderedControls;
        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Set Delay to " + v_EngineSpeed + "ms between commands]";
        }
    }
}