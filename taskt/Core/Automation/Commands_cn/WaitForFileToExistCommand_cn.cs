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
    [Attributes.ClassAttributes.Group("(File Operation Commands)文件操作命令")]
    [Attributes.ClassAttributes.Description("此命令等待文件存在于指定目标")]
    [Attributes.ClassAttributes.UsesDescription("在继续之前，使用此命令等待文件存在。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现''以实现自动化。")]
    public class WaitForFileToExistCommand_cn : ScriptCommand
    {

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指出该文件的目录")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowFileSelectionHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或选择文件的路径。")]
        [Attributes.PropertyAttributes.SampleUsage("C:\\temp\\myfile.txt or [vTextFilePath]")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_FileName { get; set; }


        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("指示等待文件存在的秒数")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("指定在发生错误之前等待的时间，因为找不到该文件。")]
        [Attributes.PropertyAttributes.SampleUsage("**10** or **20**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_WaitTime { get; set; }

        public WaitForFileToExistCommand_cn()
        {
            this.CommandName = "WaitForFileToExistCommand";
            // this.SelectionName = "Wait For File";
            this.SelectionName = Settings.Default.Wait_For_File_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {

            //convert items to variables
            var fileName = v_FileName.ConvertToUserVariable(sender);
            var pauseTime = int.Parse(v_WaitTime.ConvertToUserVariable(sender));

            //determine when to stop waiting based on user config
            var stopWaiting = DateTime.Now.AddSeconds(pauseTime);

            //initialize flag for file found
            var fileFound = false;


            //while file has not been found
            while (!fileFound)
            {

                //if file exists at the file path
                if (System.IO.File.Exists(fileName))
                {
                    fileFound = true;
                }

                //test if we should exit and throw exception
                if (DateTime.Now > stopWaiting)
                {
                    throw new Exception("File was not found in time!");
                }

                //put thread to sleep before iterating
                System.Threading.Thread.Sleep(100);
            }




        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_FileName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_WaitTime", this, editor));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [file: " + v_FileName + ", wait " + v_WaitTime + "s]";
        }
    }
}