using System;
using System.Linq;
using System.Xml.Serialization;
using System.Data;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using taskt.UI.CustomControls;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Variable Commands)变量命令")]
    [Attributes.ClassAttributes.Description("此命令允许您修改变量。")]
    [Attributes.ClassAttributes.UsesDescription("如果要修改变量的值，请使用此命令。 您甚至可以使用变量来修改其他变量。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从脚本引擎实现针对VariableList的操作。")]
    public class AddToVariableCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择要修改的变量")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_userVariableName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请定义要添加到变量的输入")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入变量值应设置为的输入。")]
        [Attributes.PropertyAttributes.SampleUsage("Hello or [vNum]+1")]
        [Attributes.PropertyAttributes.Remarks("如果将变量包含在括号[vName]中，则可以在输入中使用变量。 您还可以执行基本的数学运算。")]
        public string v_Input { get; set; }
        public AddToVariableCommand_cn()
        {
            this.CommandName = "AddToVariableCommand";
            // this.SelectionName = "Add To Variable";
            this.SelectionName = Settings.Default.Add_To_Variable_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            //get sending instance
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var requiredVariable = LookupVariable(engine);

            //if still not found and user has elected option, create variable at runtime
            if ((requiredVariable == null) && (engine.engineSettings.CreateMissingVariablesDuringExecution))
            {
                engine.VariableList.Add(new Script.ScriptVariable() { VariableName = v_userVariableName });
                requiredVariable = LookupVariable(engine);
            }

            var variableInput = v_Input.ConvertToUserVariable(sender);


            if (requiredVariable != null)
            {

                if (requiredVariable.VariableValue is List<string>)
                {
                    var existingList = (List<string>)requiredVariable.VariableValue;
                    existingList.Add(variableInput);
                }
                else
                {
                    var itemList = new List<string>();

                    if (requiredVariable.GetDisplayValue() != "")
                        itemList.Add(requiredVariable.GetDisplayValue());

                    itemList.Add(variableInput);

                    requiredVariable.VariableValue = itemList;
                  
                }
           
    
            }
            else
            {
                throw new Exception("Attempted to add data to a variable, but the variable was not found. Enclose variables within brackets, ex. {vVariable}");
            }
        }

        private Script.ScriptVariable LookupVariable(Core.Automation.Engine.AutomationEngineInstance sendingInstance)
        {
            //search for the variable
            var requiredVariable = sendingInstance.VariableList.Where(var => var.VariableName == v_userVariableName).FirstOrDefault();

            //if variable was not found but it starts with variable naming pattern
            if ((requiredVariable == null) && (v_userVariableName.StartsWith(sendingInstance.engineSettings.VariableStartMarker)) && (v_userVariableName.EndsWith(sendingInstance.engineSettings.VariableEndMarker)))
            {
                //reformat and attempt
                var reformattedVariable = v_userVariableName.Replace(sendingInstance.engineSettings.VariableStartMarker, "").Replace(sendingInstance.engineSettings.VariableEndMarker, "");
                requiredVariable = sendingInstance.VariableList.Where(var => var.VariableName == reformattedVariable).FirstOrDefault();
            }

            return requiredVariable;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Apply '" + v_Input + "' to Variable '" + v_userVariableName + "']";
        }

        public override List<Control> Render(UI.Forms.frmCommandEditor editor)
        {
            //custom rendering
            base.Render(editor);


            //create control for variable name
            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_userVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_userVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_userVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_Input", this, editor));

        

            return RenderedControls;


        }
    }
}