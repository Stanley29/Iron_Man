﻿using System;
using System.Xml.Serialization;
using System.Data;
using taskt.UI.Forms;
using System.Collections.Generic;
using System.Windows.Forms;
using taskt.UI.CustomControls;
using System.Drawing;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Input Commands)输入命令")]
    [Attributes.ClassAttributes.Description("将键击发送到目标窗口")]
    [Attributes.ClassAttributes.UsesDescription("如果要将击键输入发送到窗口，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'Windows.Forms.SendKeys'方法以实现自动化。")]
    public class UserInputCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指定标题名称")]
        [Attributes.PropertyAttributes.InputSpecification("定义要在输入表单上显示的标题。")]
        [Attributes.PropertyAttributes.SampleUsage("n/a")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_InputHeader { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指定输入方向")]
        [Attributes.PropertyAttributes.InputSpecification("定义要为用户提供的路线。")]
        [Attributes.PropertyAttributes.SampleUsage("n/a")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_InputDirections { get; set; }

        [XmlElement]
        [Attributes.PropertyAttributes.PropertyDescription("用户输入参数")]
        [Attributes.PropertyAttributes.InputSpecification("定义所需的输入参数。")]
        [Attributes.PropertyAttributes.SampleUsage("n/a")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public DataTable v_UserInputConfig { get; set; }

        [XmlIgnore]
        [NonSerialized]
        private DataGridView UserInputGridViewHelper;

        [XmlIgnore]
        [NonSerialized]
        private CommandItemControl AddRowControl;

        public UserInputCommand_cn()
        {
            this.CommandName = "UserInputCommand";
            // this.SelectionName = "Prompt for Input";
            this.SelectionName = Settings.Default.Prompt_for_Input_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;

            v_UserInputConfig = new DataTable();
            v_UserInputConfig.TableName = DateTime.Now.ToString("UserInputParamTable" + DateTime.Now.ToString("MMddyy.hhmmss"));
            v_UserInputConfig.Columns.Add("Type");
            v_UserInputConfig.Columns.Add("Label");
            v_UserInputConfig.Columns.Add("Size");
            v_UserInputConfig.Columns.Add("DefaultValue");
            v_UserInputConfig.Columns.Add("UserInput");
            v_UserInputConfig.Columns.Add("ApplyToVariable");

            v_InputHeader = "Please Provide Input";
            v_InputDirections = "Directions: Please fill in the following fields";

        }

        public override void RunCommand(object sender)
        {


            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

                        
            if (engine.tasktEngineUI == null)
            {
                engine.ReportProgress("UserInput Supported With UI Only");
                System.Windows.Forms.MessageBox.Show("UserInput Supported With UI Only", "UserInput Command", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }


            //create clone of original
            var clonedCommand = taskt.Core.Common.Clone(this);

            //translate variable
            clonedCommand.v_InputHeader = clonedCommand.v_InputHeader.ConvertToUserVariable(sender);
            clonedCommand.v_InputDirections = clonedCommand.v_InputDirections.ConvertToUserVariable(sender);

            //translate variables for each label
            foreach (DataRow rw in clonedCommand.v_UserInputConfig.Rows)
            {
                rw["Label"] = rw["Label"].ToString().ConvertToUserVariable(sender);

                var targetVariable = rw["ApplyToVariable"] as string;

                if (string.IsNullOrEmpty(targetVariable))
                {
                    var newMessage = new MessageBoxCommand();
                    newMessage.v_Message = "User Input question '" + rw["Label"] + "' is missing variables to apply results to! Results for the item will not be tracked.  To fix this, assign a variable in the designer!";
                    newMessage.v_AutoCloseAfter = 10;
                    newMessage.RunCommand(sender);
                }
            }


            //invoke ui for data collection
            var result = engine.tasktEngineUI.Invoke(new Action(() =>
            {

                //get input from user
              var userInputs =  engine.tasktEngineUI.ShowInput(clonedCommand);

                //check if user provided input
                if (userInputs != null)
                {

                    //loop through each input and assign
                    for (int i = 0; i < userInputs.Count; i++)
                    {
                        //get target variable
                        var targetVariable = v_UserInputConfig.Rows[i]["ApplyToVariable"] as string;


                        //if engine is expected to create variables, the user will not expect them to contain start/end markers
                        //ex. {vAge} should not be created, vAge should be created and then called by doing {vAge}
                        if ((!string.IsNullOrEmpty(targetVariable)) && (engine.engineSettings.CreateMissingVariablesDuringExecution))
                        {
                            //remove start markers
                            if (targetVariable.StartsWith(engine.engineSettings.VariableStartMarker))
                            {
                                targetVariable = targetVariable.TrimStart(engine.engineSettings.VariableStartMarker.ToCharArray());
                            }

                            //remove end markers
                            if (targetVariable.EndsWith(engine.engineSettings.VariableEndMarker))
                            {
                                targetVariable = targetVariable.TrimEnd(engine.engineSettings.VariableEndMarker.ToCharArray());
                            }
                        }

                       
                        //store user data in variable
                        if (!string.IsNullOrEmpty(targetVariable))
                        {
                            userInputs[i].StoreInUserVariable(sender, targetVariable);
                        }

                                  
                    }


                }

            }

            ));


        }

        public override List<Control> Render(frmCommandEditor editor)
        {


            base.Render(editor);

            UserInputGridViewHelper = new DataGridView();
            UserInputGridViewHelper.KeyDown += UserInputDataGridView_KeyDown;
            UserInputGridViewHelper.DataBindings.Add("DataSource", this, "v_UserInputConfig", false, DataSourceUpdateMode.OnPropertyChanged);

            var typefield = new DataGridViewComboBoxColumn();
            typefield.Items.Add("TextBox");
            typefield.Items.Add("CheckBox");
            typefield.Items.Add("ComboBox");
            typefield.HeaderText = "Input Type";
            typefield.DataPropertyName = "Type";
            UserInputGridViewHelper.Columns.Add(typefield);

            var field = new DataGridViewTextBoxColumn();
            field.HeaderText = "Input Label";
            field.DataPropertyName = "Label";
            UserInputGridViewHelper.Columns.Add(field);


            field = new DataGridViewTextBoxColumn();
            field.HeaderText = "Input Size (X,Y)";
            field.DataPropertyName = "Size";
            UserInputGridViewHelper.Columns.Add(field);

            field = new DataGridViewTextBoxColumn();
            field.HeaderText = "Default Value";
            field.DataPropertyName = "DefaultValue";
            UserInputGridViewHelper.Columns.Add(field);

            field = new DataGridViewTextBoxColumn();
            field.HeaderText = "Assigned Variable";
            field.DataPropertyName = "ApplyToVariable";
            UserInputGridViewHelper.Columns.Add(field);


            UserInputGridViewHelper.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            UserInputGridViewHelper.AllowUserToAddRows = false;
            UserInputGridViewHelper.AllowUserToDeleteRows = false;


            AddRowControl = new CommandItemControl();
            AddRowControl.Padding = new System.Windows.Forms.Padding(10, 0, 0, 0);
            AddRowControl.ForeColor = Color.AliceBlue;
            AddRowControl.Font = new Font("Segoe UI Semilight", 10);
            AddRowControl.CommandImage = UI.Images.GetUIImage("ExecuteDLLCommand");
            AddRowControl.CommandDisplay = "Add Input Parameter";
            AddRowControl.Click += (sender, e) => AddInputParameter(sender, e, editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InputHeader", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InputDirections", this, editor));
            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_UserInputConfig", this));
            RenderedControls.Add(AddRowControl);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_UserInputConfig", this, new Control[] { UserInputGridViewHelper }, editor));
            RenderedControls.Add(UserInputGridViewHelper);



            return RenderedControls;

        }

        private void AddInputParameter(object sender, EventArgs e, frmCommandEditor editor)
        {
            var newRow = v_UserInputConfig.NewRow();
            newRow["Size"] = "500,100";
            v_UserInputConfig.Rows.Add(newRow);

        }

        private void UserInputDataGridView_KeyDown(object sender, KeyEventArgs e)
        {


            if (UserInputGridViewHelper.SelectedRows.Count > 0)
            {
                UserInputGridViewHelper.Rows.RemoveAt(UserInputGridViewHelper.SelectedCells[0].RowIndex);
            }

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [" + v_InputHeader + "]";
        }
    }
}