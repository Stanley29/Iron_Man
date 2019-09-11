using System;
using System.Xml.Serialization;
using System.IO;
using taskt.UI.CustomControls;
using System.Collections.Generic;
using System.Windows.Forms;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{

    [Serializable]
    [Attributes.ClassAttributes.Group("(File Operation Commands)文件操作命令")]
    [Attributes.ClassAttributes.Description("此命令重命名指定目标的文件")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令重命名现有文件。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现''以实现自动化。")]
    public class RenameFileCommand_cn : ScriptCommand
    {

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指明源文件的路径")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowFileSelectionHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或选择文件的路径。")]
        [Attributes.PropertyAttributes.SampleUsage("C:\\temp\\myfile.txt or [vTextFilePath]")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_SourceFilePath { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请注明新文件名（带扩展名）")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("指定包含扩展名的新文件名。")]
        [Attributes.PropertyAttributes.SampleUsage("newfile.txt or newfile.png")]
        [Attributes.PropertyAttributes.Remarks("更改文件扩展名不会自动转换文件。")]
        public string v_NewName { get; set; }


        public RenameFileCommand_cn()
        {
            this.CommandName = "RenameFileCommand";
            //this.SelectionName = "Rename File";
            this.SelectionName = Settings.Default.Rename_File_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {

            //apply variable logic
            var sourceFile = v_SourceFilePath.ConvertToUserVariable(sender);
            var newFileName = v_NewName.ConvertToUserVariable(sender);

            //get source file name and info
            System.IO.FileInfo sourceFileInfo = new FileInfo(sourceFile);

            //create destination
            var destinationPath = System.IO.Path.Combine(sourceFileInfo.DirectoryName, newFileName);

            //rename file
            System.IO.File.Move(sourceFile, destinationPath);

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SourceFilePath", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_NewName", this, editor));

            return RenderedControls;
        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [rename " + v_SourceFilePath + " to '" + v_NewName + "']";
        }
    }
}