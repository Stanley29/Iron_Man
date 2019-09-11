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
    [Attributes.ClassAttributes.Description("此命令允许您启动程序或进程。")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令通过输入名称（例如'chrome.exe'）或文件'c：/some.exe的完全限定路径来启动应用程序")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'Process.Start'并等待脚本/程序退出，然后继续。")]
    public class StartProcessCommand : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入程序的名称或路径（例如记事本，calc）")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowFileSelectionHelper)]
        [Attributes.PropertyAttributes.InputSpecification("提供有效的程序名称或输入脚本/可执行文件的完整路径，包括扩展名")]
        [Attributes.PropertyAttributes.SampleUsage("**notepad**, **calc**, **c:\\temp\\myapp.exe**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_ProgramName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入任何参数（如果适用）")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入任何参数或标志（如果适用）。")]
        [Attributes.PropertyAttributes.SampleUsage(" **-a** or **-version**")]
        [Attributes.PropertyAttributes.Remarks("您需要查阅文档以确定您的可执行文件是否支持启动时的参数或标志。")]
        public string v_ProgramArgs { get; set; }

        public StartProcessCommand()
        {
            this.CommandName = "StartProcessCommand";
            // this.SelectionName = "Start Process";
            this.SelectionName = Settings.Default.Start_Process_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {

            string vProgramName = v_ProgramName.ConvertToUserVariable(sender);
            string vProgramArgs = v_ProgramArgs.ConvertToUserVariable(sender);

            if (v_ProgramArgs == "")
            {
                System.Diagnostics.Process.Start(vProgramName);
            }
            else
            {
                System.Diagnostics.Process.Start(vProgramName, vProgramArgs);
            }

            System.Threading.Thread.Sleep(2000);
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_ProgramName", this, editor));

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_ProgramArgs", this, editor));

         
            return RenderedControls;
        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Process: " + v_ProgramName + "]";
        }
    }
}