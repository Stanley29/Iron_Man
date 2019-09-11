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
    [Attributes.ClassAttributes.Group("(Text File Commands)文本文件命令")]
    [Attributes.ClassAttributes.Description("此命令将指定的数据写入文本文件")]
    [Attributes.ClassAttributes.UsesDescription("如果要将数据写入文本文件，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现''以实现自动化。")]
    public class WriteTextFileCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指出文件的路径")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或选择文本文件的路径。")]
        [Attributes.PropertyAttributes.SampleUsage("C:\\temp\\myfile.txt or [vTextFilePath]")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_FilePath { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请说明要写的文字。 [crLF]插入换行符。")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("指示应将文本写入文件。")]
        [Attributes.PropertyAttributes.SampleUsage("**[vText]** or **Hello World!**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_TextToWrite { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择覆盖选项")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("附加")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("覆盖")]
        [Attributes.PropertyAttributes.InputSpecification("指示此命令是应将文本附加到文件还是覆盖文件中的所有现有文本")]
        [Attributes.PropertyAttributes.SampleUsage("选择**附加**或**覆盖**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_Overwrite { get; set; }
        public WriteTextFileCommand_cn()
        {
            this.CommandName = "WriteTextFileCommand";
            // this.SelectionName = "Write Text File";
            this.SelectionName = Settings.Default.Write_Text_File_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            //convert variables
            var filePath = v_FilePath.ConvertToUserVariable(sender);
            var outputText = v_TextToWrite.ConvertToUserVariable(sender).ToString().Replace("[crLF]",Environment.NewLine);

            //append or overwrite as necessary
            if (v_Overwrite == "Append")
            {
                System.IO.File.AppendAllText(filePath, outputText);
            }
            else
            {
                System.IO.File.WriteAllText(filePath, outputText);
            }

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_FilePath", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_TextToWrite", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_Overwrite", this, editor));

            return RenderedControls;
        }


        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [" + v_Overwrite + " to '" + v_FilePath + "']";
        }
    }
}