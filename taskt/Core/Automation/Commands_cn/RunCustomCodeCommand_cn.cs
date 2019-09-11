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
    [Attributes.ClassAttributes.Description("此命令允许您从输入中运行C＃代码")]
    [Attributes.ClassAttributes.UsesDescription("如果要运行自定义C＃代码命令，请使用此命令。 调用此命令时，将编译此命令中的代码并在运行时运行。 此命令仅支持标准框架类。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'Process.Start'并等待脚本/程序退出，然后继续。")]
    public class RunCustomCodeCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("粘贴C＃代码以执行")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowCodeBuilder)]
        [Attributes.PropertyAttributes.InputSpecification("输入要执行的代码或使用构建器创建自定义C＃代码。 构建器包含可用于构建的Hello World模板。")]
        [Attributes.PropertyAttributes.SampleUsage("n/a")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_Code { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("供应参数（可选）")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入自定义代码在执行期间将接收的参数")]
        [Attributes.PropertyAttributes.SampleUsage("n/a")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_Args { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("可选 - 选择要接收输出的变量")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }

        public RunCustomCodeCommand_cn()
        {
            this.CommandName = "RunCustomCodeCommand";
            // this.SelectionName = "Run Custom Code";
            this.SelectionName = Settings.Default.Run_Custom_Code_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            //create compiler service
            var compilerSvc = new Core.CompilerServices();
            var customCode = v_Code.ConvertToUserVariable(sender);

            //compile custom code
            var result = compilerSvc.CompileInput(customCode);

            //check for errors
            if (result.Errors.HasErrors)
            {
                //throw exception
                var errors = string.Join(", ", result.Errors);
                throw new Exception("Errors Occured: " + errors);
            }
            else
            {

                var arguments = v_Args.ConvertToUserVariable(sender);
            
                //run code, taskt will wait for the app to exit before resuming
                System.Diagnostics.Process scriptProc = new System.Diagnostics.Process();
                scriptProc.StartInfo.FileName = result.PathToAssembly;
                scriptProc.StartInfo.Arguments = arguments;

                if (v_applyToVariableName != "")
                {
                    //redirect output
                    scriptProc.StartInfo.RedirectStandardOutput = true;
                    scriptProc.StartInfo.UseShellExecute = false;
                }
             

                scriptProc.Start();

                scriptProc.WaitForExit();

                if (v_applyToVariableName != "")
                {
                    var output = scriptProc.StandardOutput.ReadToEnd();
                    output.StoreInUserVariable(sender, v_applyToVariableName);
                }
    

                scriptProc.Close();


            }


        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_Code", this, editor));

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_Args", this, editor));

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;
        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue();
        }
    }
}