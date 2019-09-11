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
    [Attributes.ClassAttributes.Group("(Program/Process Commands)程序/进程命令")]
    [Attributes.ClassAttributes.Description("此命令允许您停止程序或进程。")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令按名称“chrome”关闭应用程序。 或者，您可以使用“关闭窗口”或“厚应用程序命令”。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'Process.CloseMainWindow'。")]
    public class StopProcessCommand : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("输入要停止的进程名称（calc，notepad）")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("提供程序进程名称，因为它在Windows任务管理器中显示为进程")]
        [Attributes.PropertyAttributes.SampleUsage("**notepad**, **calc**")]
        [Attributes.PropertyAttributes.Remarks("程序名称可能与实际进程名称不同。 您可以使用Thick App命令关闭应用程序窗口。")]
        public string v_ProgramShortName { get; set; }

        public StopProcessCommand()
        {
            this.CommandName = "StopProgramCommand";
            // this.SelectionName = "Stop Process";
            this.SelectionName = Settings.Default.Stop_Process_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            string shortName = v_ProgramShortName.ConvertToUserVariable(sender);
            var processes = System.Diagnostics.Process.GetProcessesByName(shortName);

            foreach (var prc in processes)
                prc.CloseMainWindow();
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_ProgramShortName", this, editor));


            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Process: " + v_ProgramShortName + "]";
        }
    }
}