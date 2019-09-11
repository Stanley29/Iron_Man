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
    public class OSVariableCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("选择所需的系统变量")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从其中一个选项中选择")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_OSVariableName { get; set; }

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
        public OSVariableCommand_cn()
        {
            this.CommandName = "OSVariableCommand";
            // this.SelectionName = "OS Variable";
            this.SelectionName = Settings.Default.OS_Variable_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var systemVariable = (string)v_OSVariableName.ConvertToUserVariable(sender);

            ObjectQuery wql = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(wql);
            ManagementObjectCollection results = searcher.Get();

            foreach (ManagementObject result in results)
            {
                foreach (PropertyData prop in result.Properties)
                {
                    if (prop.Name == systemVariable.ToString())
                    {
                        var sysValue = prop.Value.ToString();
                        sysValue.StoreInUserVariable(sender, v_applyToVariableName);
                        return;
                    }
                }

            }

            throw new Exception("System Property '" + systemVariable + "' not found!");


        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            var ActionNameComboBoxLabel = CommandControls.CreateDefaultLabelFor("v_OSVariableName", this);
            VariableNameComboBox = (ComboBox)CommandControls.CreateDropdownFor("v_OSVariableName", this);


            ObjectQuery wql = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(wql);
            ManagementObjectCollection results = searcher.Get();

            foreach (ManagementObject result in results)
            {
                foreach (PropertyData prop in result.Properties)
                {
                    VariableNameComboBox.Items.Add(prop.Name);
                }

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


            ObjectQuery wql = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(wql);
            ManagementObjectCollection results = searcher.Get();

            foreach (ManagementObject result in results)
            {
                foreach (PropertyData prop in result.Properties)
                {
                    if (prop.Name == selectedValue.ToString())
                    {
                        VariableValue.Text = "[ex. " + prop.Value + "]";
                        return;
                    }
                }

            }

            VariableValue.Text = "[ex. **Item not found**]";

          
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Apply '" + v_OSVariableName + "' to Variable '" + v_applyToVariableName + "']";
        }
    }


}