using System;
using System.Linq;
using System.Xml.Serialization;
using System.Data;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Windows.Forms;
using taskt.UI.Forms;
using taskt.UI.CustomControls;
using System.Drawing;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Data Commands)数据命令")]
    [Attributes.ClassAttributes.Description("此命令允许您执行高级字符串提取。")]
    [Attributes.ClassAttributes.UsesDescription("如果要从较大的文本或变量中提取一段文本，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从脚本引擎实现针对VariableList的操作。")]
    public class TextExtractorCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("提供需要提取的值或变量（例如[vSomeVariable]）")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("选择或提供变量或文本值")]
        [Attributes.PropertyAttributes.SampleUsage("**Hello** or **vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_InputValue { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择文字提取类型")]
        [Attributes.PropertyAttributes.InputSpecification("选择所需的提取类型。")]
        [Attributes.PropertyAttributes.SampleUsage("从文本之前，文本之后，文本之间选择")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("提取所有文本后")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("提取所有文本之前")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("在文本之间提取所有内容")]
        public string v_TextExtractionType { get; set; }

        [XmlElement]
        [Attributes.PropertyAttributes.PropertyDescription("提取参数")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("定义所需的提取参数，这取决于提取的类型。")]
        [Attributes.PropertyAttributes.SampleUsage("n/a")]
        [Attributes.PropertyAttributes.Remarks("")]
        public DataTable v_TextExtractionTable { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择变量以接收提取的文本")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }

        [XmlIgnore]
        [NonSerialized]
        private DataGridView ParametersGridViewHelper;

        public TextExtractorCommand_cn()
        {

            this.CommandName = "TextExtractorCommand";
            // this.SelectionName = "Text Extraction";
            this.SelectionName = Settings.Default.Text_Extraction_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;

            //define parameter table
            this.v_TextExtractionTable = new System.Data.DataTable
            {
                TableName = DateTime.Now.ToString("TextExtractorParamTable" + DateTime.Now.ToString("MMddyy.hhmmss"))
            };

            this.v_TextExtractionTable.Columns.Add("Parameter Name");
            this.v_TextExtractionTable.Columns.Add("Parameter Value");

        

        }

        public override void RunCommand(object sender)
        {
            //get variablized input
            var variableInput = v_InputValue.ConvertToUserVariable(sender);


            string variableLeading, variableTrailing, skipOccurences, extractedText;

            //handle extraction cases
            switch (v_TextExtractionType)
            {
                case "Extract All After Text":
                    //extract trailing texts            
                    variableLeading = GetParameterValue("Leading Text").ConvertToUserVariable(sender);
                    skipOccurences = GetParameterValue("Skip Past Occurences").ConvertToUserVariable(sender);
                    extractedText = ExtractLeadingText(variableInput, variableLeading, skipOccurences);
                    break;
                case "Extract All Before Text":
                    //extract leading text
                    variableTrailing = GetParameterValue("Trailing Text").ConvertToUserVariable(sender);
                    skipOccurences = GetParameterValue("Skip Past Occurences").ConvertToUserVariable(sender);
                    extractedText = ExtractTrailingText(variableInput, variableTrailing, skipOccurences);
                    break;
                case "Extract All Between Text":
                    //extract leading and then trailing which gives the items between
                    variableLeading = GetParameterValue("Leading Text").ConvertToUserVariable(sender);
                    variableTrailing = GetParameterValue("Trailing Text").ConvertToUserVariable(sender);
                    skipOccurences = GetParameterValue("Skip Past Occurences").ConvertToUserVariable(sender);

                    //extract leading
                    extractedText = ExtractLeadingText(variableInput, variableLeading, skipOccurences);

                    //extract trailing -- assume we will take to the first item
                    extractedText = ExtractTrailingText(extractedText, variableTrailing, "0");

                    break;
                default:
                    throw new NotImplementedException("Extraction Type Not Implemented: " + v_TextExtractionType);
            }

            //store variable
            extractedText.StoreInUserVariable(sender, v_applyToVariableName);

        }

        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            ParametersGridViewHelper = new DataGridView();
            ParametersGridViewHelper.AllowUserToAddRows = true;
            ParametersGridViewHelper.AllowUserToDeleteRows = true;
            ParametersGridViewHelper.Size = new Size(350, 125);
            ParametersGridViewHelper.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            ParametersGridViewHelper.DataBindings.Add("DataSource", this, "v_TextExtractionTable", false, DataSourceUpdateMode.OnPropertyChanged);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InputValue", this, editor));


            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_TextExtractionType", this));
            var selectionControl = (ComboBox)CommandControls.CreateDropdownFor("v_TextExtractionType", this);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_TextExtractionType", this, new Control[] { selectionControl }, editor));
            selectionControl.SelectionChangeCommitted += textExtraction_SelectionChangeCommitted;
            RenderedControls.Add(selectionControl);


            //create control for variable name
            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_TextExtractionTable", this));
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_TextExtractionTable", this, new Control[] { ParametersGridViewHelper }, editor));
            RenderedControls.Add(ParametersGridViewHelper);



            return RenderedControls;
        }
        private void textExtraction_SelectionChangeCommitted(object sender, EventArgs e)
        {
            ComboBox extractionAction = (ComboBox)sender;

            if ((ParametersGridViewHelper == null) || (extractionAction == null) || (ParametersGridViewHelper.DataSource == null))
                return;


            var textParameters = (DataTable)ParametersGridViewHelper.DataSource;

            textParameters.Rows.Clear();


            switch (extractionAction.SelectedItem)
            {
                case "Extract All After Text":
                    textParameters.Rows.Add("Leading Text", "");
                    textParameters.Rows.Add("Skip Past Occurences", "0");
                    break;
                case "Extract All Before Text":
                    textParameters.Rows.Add("Trailing Text", "");
                    textParameters.Rows.Add("Skip Past Occurences", "0");
                    break;
                case "Extract All Between Text":
                    textParameters.Rows.Add("Leading Text", "");
                    textParameters.Rows.Add("Trailing Text", "");
                    textParameters.Rows.Add("Skip Past Occurences", "0");
                    break;
                default:
                    break;
            }
        }

        private string GetParameterValue(string parameterName)
        {
            return ((from rw in v_TextExtractionTable.AsEnumerable()
                     where rw.Field<string>("Parameter Name") == parameterName
                     select rw.Field<string>("Parameter Value")).FirstOrDefault());

        }
        private string ExtractLeadingText(string input, string substring, string occurences)
        {

            //verify the occurence index
            int leadingOccurenceIndex = 0;

            if (!int.TryParse(occurences, out leadingOccurenceIndex))
            {
                throw new Exception("Invalid Index For Extraction - " + occurences);
            }

            //find index matches
            var leadingOccurencesFound = Regex.Matches(input, substring).Cast<Match>().Select(m => m.Index).ToList();

            //handle if we are searching beyond what was found
            if (leadingOccurenceIndex >= leadingOccurencesFound.Count)
            {
                throw new Exception("No value was found after skipping " + leadingOccurenceIndex + " instance(s).  Only " + leadingOccurencesFound.Count + " instances exist.");
            }

            //declare start position
            var startPosition = leadingOccurencesFound[leadingOccurenceIndex] + substring.Length;

            //substring and apply to variable
            return input.Substring(startPosition);


        }
        private string ExtractTrailingText(string input, string substring, string occurences)
        {
            //verify the occurence index
            int leadingOccurenceIndex = 0;
            if (!int.TryParse(occurences, out leadingOccurenceIndex))
            {
                throw new Exception("Invalid Index For Extraction - " + occurences);
            }

            //find index matches
            var trailingOccurencesFound = Regex.Matches(input, substring).Cast<Match>().Select(m => m.Index).ToList();

            //handle if we are searching beyond what was found
            if (leadingOccurenceIndex >= trailingOccurencesFound.Count)
            {
                throw new Exception("No value was found after skipping " + leadingOccurenceIndex + " instance(s).  Only " + trailingOccurencesFound.Count + " instances exist.");
            }

            //declare start position
            var endPosition = trailingOccurencesFound[leadingOccurenceIndex];

            //substring
            return input.Substring(0, endPosition);
        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Apply Extracted Text To Variable: " + v_applyToVariableName + "]";
        }
    }
}