using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using System.Data;
using System.Windows.Forms;
using taskt.UI.Forms;
using taskt.UI.CustomControls;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Data Commands)数据命令")]
    [Attributes.ClassAttributes.Description("此命令允许您将JSON对象解析为列表。")]
    [Attributes.ClassAttributes.UsesDescription("如果要从JSON对象中提取数据，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("")]
    public class ParseJsonCommand : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("提供需要提取的值或变量（例如[vSomeVariable]）")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("选择或提供变量或文本值")]
        [Attributes.PropertyAttributes.SampleUsage("**Hello** or **vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_InputValue { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("指定JSON提取器")]
        [Attributes.PropertyAttributes.InputSpecification("输入JSON令牌提取器")]
        [Attributes.PropertyAttributes.SampleUsage("从文本之前，文本之后，文本之间选择")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_JsonExtractor { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择变量以接收提取的json")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }

        public ParseJsonCommand()
        {
            this.CommandName = "ParseJsonCommand";
            //this.SelectionName = "Parse JSON Object";
            this.SelectionName = Settings.Default.Parse_JSON_Object_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
          
        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;


            var forbiddenMarkers = new List<string> { "[", "]" };

            if (forbiddenMarkers.Any(f => f == engine.engineSettings.VariableStartMarker) || (forbiddenMarkers.Any(f => f == engine.engineSettings.VariableEndMarker)))
            {
                throw new Exception("Cannot use Parse JSON command with square bracket variable markers [ ]");
            }

            //get variablized input
            var variableInput = v_InputValue.ConvertToUserVariable(sender);

            //get variablized token
            var jsonSearchToken = v_JsonExtractor.ConvertToUserVariable(sender);

            //create objects
            Newtonsoft.Json.Linq.JObject o;
            IEnumerable<Newtonsoft.Json.Linq.JToken> searchResults;
            List<string> resultList = new List<string>();

            //parse json
            try
            {
                 o = Newtonsoft.Json.Linq.JObject.Parse(variableInput);
            }
            catch (Exception ex)
            {
                throw new Exception("Error Occured Selecting Tokens: " + ex.ToString());
            }
 

            //select results
            try
            {
                searchResults = o.SelectTokens(jsonSearchToken);
            }
            catch (Exception ex)
            {
                throw new Exception("Error Occured Selecting Tokens: " + ex.ToString());
            }
        

            //add results to result list since list<string> is supported
            foreach (var result in searchResults)
            {
                resultList.Add(result.ToString());
            }

            //get variable
            var requiredComplexVariable = engine.VariableList.Where(x => x.VariableName == v_applyToVariableName).FirstOrDefault();

            //create if var does not exist
            if (requiredComplexVariable == null)
            {
                engine.VariableList.Add(new Script.ScriptVariable() { VariableName = v_applyToVariableName, CurrentPosition = 0 });
                requiredComplexVariable = engine.VariableList.Where(x => x.VariableName == v_applyToVariableName).FirstOrDefault();
            }

            //assign value to variable
            requiredComplexVariable.VariableValue = resultList;

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InputValue", this, editor));

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_JsonExtractor", this, editor));

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }



        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Selector: " + v_JsonExtractor + ", Apply Result(s) To Variable: " + v_applyToVariableName + "]";
        }
    }
}