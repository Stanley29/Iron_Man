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
    [Attributes.ClassAttributes.Group("(Variable Commands)变量命令")]
    [Attributes.ClassAttributes.Description("如果您未在运行时使用**设置变量*和设置**创建缺失变量**，则此命令允许您显式添加变量。")]
    [Attributes.ClassAttributes.UsesDescription("如果要修改变量的值，请使用此命令。 您甚至可以使用变量来修改其他变量。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从脚本引擎实现针对VariableList的操作。")]
    public class AddVariableCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请定义新变量的名称")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果变量存在，则旧变量的值将替换为新变量")]
        public string v_userVariableName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请定义要设置为上述变量的输入")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入变量值应设置为的输入。")]
        [Attributes.PropertyAttributes.SampleUsage("Hello or [vNum]+1")]
        [Attributes.PropertyAttributes.Remarks("如果将变量包含在括号[vName]中，则可以在输入中使用变量。 您还可以执行基本的数学运算。")]
        public string v_Input { get; set; }
        [Attributes.PropertyAttributes.PropertyDescription("定义变量已存在时要采取的操作")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("如果变量存在则什么都不做")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("如果变量存在则出错")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("替换如果存在变量")]
        [Attributes.PropertyAttributes.InputSpecification("从列表中选择适当的处理程序")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_IfExists { get; set; }
        public AddVariableCommand_cn()
        {
            this.CommandName = "AddVariableCommand";
            // this.SelectionName = "New Variable";
            this.SelectionName = Settings.Default.New_Variable_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            //get sending instance
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;


            if (!engine.VariableList.Any(f => f.VariableName == v_userVariableName))
            {
                //variable does not exist so add to the list
                try
                {

                    var variableValue = v_Input.ConvertToUserVariable(engine);

                    engine.VariableList.Add(new Script.ScriptVariable
                    {
                        VariableName = v_userVariableName,
                        VariableValue = variableValue
                    });
                }
                catch (Exception ex)
                {
                    throw new Exception("Encountered an error when adding variable '" + v_userVariableName + "': " + ex.ToString());
                }
            }
            else
            {
                //variable exists so decide what to do
                switch (v_IfExists)
                {
                    case "Replace If Variable Exists":
                        v_Input.ConvertToUserVariable(sender).StoreInUserVariable(engine, v_userVariableName);
                        break;
                    case "Error If Variable Exists":
                        throw new Exception("Attempted to create a variable that already exists! Use 'Set Variable' instead or change the Exception Setting in the 'Add Variable' Command.");
                    default:
                        break;
                }
               
            }

         
         

        }

        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_userVariableName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_Input", this, editor));


            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_IfExists", this));
            var dropdown = CommandControls.CreateDropdownFor("v_IfExists", this);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_IfExists", this, new Control[] { dropdown }, editor));
            RenderedControls.Add(dropdown);

            return RenderedControls;
        }



        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Assign '" + v_Input + "' to New Variable '" + v_userVariableName + "']";
        }
    }
}