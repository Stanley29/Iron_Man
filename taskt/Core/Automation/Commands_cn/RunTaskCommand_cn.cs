﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{

    [Serializable]
    [Attributes.ClassAttributes.Group("(Task Commands)任务命令")]
    [Attributes.ClassAttributes.Description("此命令运行任务。")]
    [Attributes.ClassAttributes.UsesDescription("如果要运行其他任务，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("")]
    public class RunTaskCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("选择要运行的任务")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowFileSelectionHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或选择文件的有效路径。")]
        [Attributes.PropertyAttributes.SampleUsage("c:\\temp\\mytask.xml or [vScriptPath]")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_taskPath { get; set; }

        [XmlElement]
        [Attributes.PropertyAttributes.PropertyDescription("分配变量")]
        [Attributes.PropertyAttributes.InputSpecification("输入所需的作业。")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("")]
        public DataTable v_VariableAssignments { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("分配变量")]
        [Attributes.PropertyAttributes.InputSpecification("用于分配变量的用户首选项")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("")]
        public bool v_AssignVariables { get; set; }


        [XmlIgnore]
        [NonSerialized]
        private DataGridView AssignmentsGridViewHelper;

        [XmlIgnore]
        [NonSerialized]
        private CheckBox PassParameters;

        public RunTaskCommand_cn()
        {
            this.CommandName = "RunTaskCommand";
            //this.SelectionName = "Run Task";
            this.SelectionName = Settings.Default.Run_Task_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;

            v_VariableAssignments = new DataTable();
            v_VariableAssignments.Columns.Add("VariableName");
            v_VariableAssignments.Columns.Add("VariableValue");
            v_VariableAssignments.TableName = "RunTaskCommandInputParameters" + DateTime.Now.ToString("MMddyyhhmmss");

            AssignmentsGridViewHelper = new DataGridView();
            AssignmentsGridViewHelper.AllowUserToAddRows = true;
            AssignmentsGridViewHelper.AllowUserToDeleteRows = true;
            AssignmentsGridViewHelper.Size = new Size(400, 250);
            AssignmentsGridViewHelper.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            AssignmentsGridViewHelper.DataSource = v_VariableAssignments;
            AssignmentsGridViewHelper.Hide();
        }

        public override void RunCommand(object sender)
        {
            var startFile = v_taskPath.ConvertToUserVariable(sender);

            //create variable list
            var variableList = new List<Script.ScriptVariable>();
            foreach (DataRow rw in v_VariableAssignments.Rows)
            {
                var variableName = ((string)rw.ItemArray[0]).ConvertToUserVariable(sender);
                var variableValue = ((string)rw.ItemArray[1]).ConvertToUserVariable(sender);
                variableList.Add(new Script.ScriptVariable
                {
                    VariableName = variableName,
                    VariableValue = variableValue
                });
            }

            Application.Run(new UI.Forms.frmScriptEngine(startFile, null, variableList));

        }

        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create file path and helpers
            RenderedControls.Add(UI.CustomControls.CommandControls.CreateDefaultLabelFor("v_taskPath", this));
            var taskPathControl = UI.CustomControls.CommandControls.CreateDefaultInputFor("v_taskPath", this);
            RenderedControls.AddRange(UI.CustomControls.CommandControls.CreateUIHelpersFor("v_taskPath", this, new Control[] { taskPathControl }, editor));
            RenderedControls.Add(taskPathControl);

            taskPathControl.TextChanged += TaskPathControl_TextChanged;

            //
            PassParameters = new CheckBox();
            PassParameters.AutoSize = true;
            PassParameters.Text = "I want to assign variables on startup";
            PassParameters.Font = new Font("Segoe UI Light", 12);
            PassParameters.ForeColor = Color.White;
   
            PassParameters.DataBindings.Add("Checked", this, "v_AssignVariables", false, DataSourceUpdateMode.OnPropertyChanged);
            PassParameters.CheckedChanged += PassParametersCheckbox_CheckedChanged;
            RenderedControls.Add(PassParameters);

            RenderedControls.Add(AssignmentsGridViewHelper);




            return RenderedControls;
        }

        private void TaskPathControl_TextChanged(object sender, EventArgs e)
        {
            PassParameters.Checked = false;
        }

        private void PassParametersCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            var Sender = (CheckBox)sender;
            AssignmentsGridViewHelper.Visible = Sender.Checked;

            //load variables if selected and file exists
            if ((Sender.Checked) && (System.IO.File.Exists(v_taskPath)))
            {
              
                Script.Script deserializedScript = Core.Script.Script.DeserializeFile(v_taskPath);

                foreach (var variable in deserializedScript.Variables)
                {
                    DataRow[] foundVariables  = v_VariableAssignments.Select("VariableName = '" + variable.VariableName + "'");
                    if (foundVariables.Length == 0)
                    {
                        v_VariableAssignments.Rows.Add(variable.VariableName, variable.VariableValue);
                    }                  
                }


                AssignmentsGridViewHelper.DataSource = v_VariableAssignments;
            }


        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [" + v_taskPath + "]";
        }
    }
}