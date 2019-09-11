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
    [Attributes.ClassAttributes.Description("此命令允许您运行脚本或程序并等待它继续之前退出。")]
    [Attributes.ClassAttributes.UsesDescription("如果要运行脚本（例如vbScript，javascript或可执行文件）但在taskt继续执行之前等待它关闭，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'Process.Start'并等待脚本/程序退出，然后继续。")]
    public class RunScriptCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("输入脚本的路径")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入脚本的完全限定路径，包括脚本扩展。")]
        [Attributes.PropertyAttributes.SampleUsage("**C:\\temp\\myscript.vbs**")]
        [Attributes.PropertyAttributes.Remarks("此命令与** Start Process **不同，因为此命令会阻止执行，直到脚本完成。 如果您不想在脚本执行时停止，请考虑使用** Start Process **。")]
        public string v_ScriptPath { get; set; }

        public RunScriptCommand_cn()
        {
            this.CommandName = "RunScriptCommand";
            // this.SelectionName = "Run Script";
            this.SelectionName = Settings.Default.Run_Script_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            {
                System.Diagnostics.Process scriptProc = new System.Diagnostics.Process();

                var scriptPath = v_ScriptPath.ConvertToUserVariable(sender);
                scriptProc.StartInfo.FileName = scriptPath;
                scriptProc.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                scriptProc.Start();
                scriptProc.WaitForExit();

                scriptProc.Close();
            }
        }

        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_ScriptPath", this, editor));

        

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Script Path: " + v_ScriptPath + "]";
        }
    }
}