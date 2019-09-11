using System;
using System.Xml.Serialization;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using taskt.UI.Forms;
using taskt.UI.CustomControls;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{


 
    [Serializable]
    [Attributes.ClassAttributes.Group("(File Operation Commands)文件操作命令")]
    [Attributes.ClassAttributes.Description("此命令将文件移动到指定目标")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令将文件移动到新目标。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现''以实现自动化。")]
    public class MoveFileCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("指示是移动还是复制文件")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("移动文件")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("拷贝文件")]
        [Attributes.PropertyAttributes.InputSpecification("指定是要移动文件还是复制文件。 移动将从原始路径中删除文件，而复制则不会。")]
        [Attributes.PropertyAttributes.SampleUsage("选择**移动文件**或**复制文件**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_OperationType { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指明源文件的路径")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowFileSelectionHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或选择文件的路径。")]
        [Attributes.PropertyAttributes.SampleUsage("C:\\temp\\myfile.txt or [vTextFilePath]")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_SourceFilePath { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指出要复制到的目录")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或选择文件的新路径。")]
        [Attributes.PropertyAttributes.SampleUsage("C:\\temp\\new path\\ or [vTextFolderPath]")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_DestinationDirectory { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("如果目标不存在，则创建文件夹")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("是")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("没有")]
        [Attributes.PropertyAttributes.InputSpecification("指定是否应该创建目录（如果该目录尚不存在）。")]
        [Attributes.PropertyAttributes.SampleUsage("选择**是**或**否**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_CreateDirectory { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("删除文件（如果已存在）")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("是")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("没有")]
        [Attributes.PropertyAttributes.InputSpecification("指定是否应该首先删除该文件（如果已找到该文件存在）。")]
        [Attributes.PropertyAttributes.SampleUsage("选择**是**或**否**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_DeleteExisting { get; set; }


        public MoveFileCommand_cn()
        {
            this.CommandName = "MoveFileCommand";
            // this.SelectionName = "Move/Copy File";
            this.SelectionName = Settings.Default.Move_Copy_File_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {

            //apply variable logic
            var sourceFile = v_SourceFilePath.ConvertToUserVariable(sender);
            var destinationFolder = v_DestinationDirectory.ConvertToUserVariable(sender);

            if ((v_CreateDirectory == "Yes") && (!System.IO.Directory.Exists(destinationFolder)))
            {
                System.IO.Directory.CreateDirectory(destinationFolder);
            }

            //get source file name and info
            System.IO.FileInfo sourceFileInfo = new FileInfo(sourceFile);

            //create destination
            var destinationPath = System.IO.Path.Combine(destinationFolder, sourceFileInfo.Name);

            //delete if it already exists per user
            if (v_DeleteExisting == "Yes")
            {
                System.IO.File.Delete(destinationPath);
            }

            if (v_OperationType == "Move File")
            {
                //move file
                System.IO.File.Move(sourceFile, destinationPath);
            }
            else
            {
                //copy file
                System.IO.File.Copy(sourceFile, destinationPath);
            }


        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_OperationType", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SourceFilePath", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_DestinationDirectory", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_CreateDirectory", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_DeleteExisting", this, editor));
            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [" + v_OperationType + " from '" + v_SourceFilePath + "' to '" + v_DestinationDirectory + "']";
        }
    }
}