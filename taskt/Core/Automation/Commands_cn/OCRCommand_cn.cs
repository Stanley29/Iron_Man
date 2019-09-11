using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{

  

    [Serializable]
    [Attributes.ClassAttributes.Group("(Image Commands)图像命令")]
    [Attributes.ClassAttributes.Description("此命令允许您将图像文件转换为文本以进行解析。")]
    [Attributes.ClassAttributes.UsesDescription("如果要将图像转换为文本，请使用此命令。 然后，您可以使用其他命令来解析数据。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令依赖并实现OneNote OCR以实现自动化。")]
    public class OCRCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("选择图像到OCR")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowFileSelectionHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或选择图像文件的路径。")]
        [Attributes.PropertyAttributes.SampleUsage(@"**c:\temp\myimages.png")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_FilePath { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("将OCR结果应用于变量")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_userVariableName { get; set; }
        public OCRCommand_cn()
        {
            this.DefaultPause = 0;
            this.CommandName = "OCRCommand";
            // this.SelectionName = "Perform OCR";
            this.SelectionName = Settings.Default.Perform_OCR_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var ocrEngine = new OneNoteOCRDll.OneNoteOCR();
            var arr = ocrEngine.OcrTexts(v_FilePath.ConvertToUserVariable(engine)).ToArray();

            string endResult = "";
            foreach (var text in arr)
            {
                endResult += text.Text;
            }

            //apply to user variable
            endResult.StoreInUserVariable(sender, v_userVariableName);
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_FilePath", this, editor));

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_userVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_userVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_userVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return "OCR '" + v_FilePath + "' and apply result to '" + v_userVariableName + "'";
        }
    }
}