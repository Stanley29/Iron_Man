using System;
using System.Collections.Generic;
using System.Management;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(System Commands)系统命令")]
    [Attributes.ClassAttributes.Description("此命令允许您专门选择系统/环境变量")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令专门检索系统变量")]
    [Attributes.ClassAttributes.ImplementationDescription("")]
    public class EnvironmentVariableCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("选择所需的环境变量")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从其中一个选项中选择")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_EnvVariableName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择变量以接收输出")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }


        [XmlIgnore]
        [NonSerialized]
        public ComboBox VariableNameComboBox;

        [XmlIgnore]
        [NonSerialized]
        public Label VariableValue;
        public EnvironmentVariableCommand_cn()
        {
            this.CommandName = "EnvironmentVariableCommand";
            // this.SelectionName = "Environment Variable";
            this.SelectionName = Settings.Default.Environment_Variable_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var environmentVariable = (string)v_EnvVariableName.ConvertToUserVariable(sender);

            var variables = Environment.GetEnvironmentVariables();
            var envValue = (string)variables[environmentVariable];

            envValue.StoreInUserVariable(sender, v_applyToVariableName);


        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            var ActionNameComboBoxLabel = CommandControls.CreateDefaultLabelFor("v_EnvVariableName", this);
            VariableNameComboBox = (ComboBox)CommandControls.CreateDropdownFor("v_EnvVariableName", this);


            foreach (System.Collections.DictionaryEntry env in Environment.GetEnvironmentVariables())
            {
                var envVariableKey = env.Key.ToString();
                var envVariableValue = env.Value.ToString();
                VariableNameComboBox.Items.Add(envVariableKey);
            }


            VariableNameComboBox.SelectedValueChanged += VariableNameComboBox_SelectedValueChanged;



            RenderedControls.Add(ActionNameComboBoxLabel);
            RenderedControls.Add(VariableNameComboBox);

            VariableValue = new Label();
            VariableValue.Font = new System.Drawing.Font("Segoe UI", 12);
            VariableValue.ForeColor = System.Drawing.Color.White;

            RenderedControls.Add(VariableValue);


            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_applyToVariableName", this, editor));
            

            return RenderedControls;
        }

        private void VariableNameComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            var selectedValue = VariableNameComboBox.SelectedItem;

            if (selectedValue == null)
                return;


            var variable = Environment.GetEnvironmentVariables();
            var value = variable[selectedValue];

            VariableValue.Text = "[ex. " + value + "]";


        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Apply '" + v_EnvVariableName + "' to Variable '" + v_applyToVariableName + "']";
        }
    }

}