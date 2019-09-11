using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.Core.IO;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
        [Serializable]
    [Attributes.ClassAttributes.Group("(Data Commands)数据命令")]
    [Attributes.ClassAttributes.Description("")]
    [Attributes.ClassAttributes.UsesDescription("")]
    [Attributes.ClassAttributes.ImplementationDescription("")]
    public class PDFTextExtractionCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指出PDF文件路径")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowFileSelectionHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或选择适用文件的路径。")]
        [Attributes.PropertyAttributes.SampleUsage(@"C:\temp\myfile.pdf or [vFilePath]")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_FilePath { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择变量以接收PDF文本")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }

        public PDFTextExtractionCommand_cn()
        {
            this.CommandName = "PDFTextExtractionCommand";
            //this.SelectionName = "PDF Extraction";
            this.SelectionName = Settings.Default.PDF_Extraction_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {

            //get variable path to source file
            var vSourceFilePath = v_FilePath.ConvertToUserVariable(sender);


            if (!System.IO.File.Exists(vSourceFilePath))
            {
                throw new System.IO.FileNotFoundException("Could not find file: " + vSourceFilePath);
            }

            //create process interface
            JavaInterface javaInterface = new JavaInterface();

            //get output from process
            var result = javaInterface.ExtractPDFText(vSourceFilePath);

            //apply to variable
            result.StoreInUserVariable(sender, v_applyToVariableName);



        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_FilePath", this, editor));


            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Extract From '" + v_FilePath + "' and apply result to '" + v_applyToVariableName + "'" ;
        }
    }
}