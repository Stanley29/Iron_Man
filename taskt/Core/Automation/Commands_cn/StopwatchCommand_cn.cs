using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Engine Commands)发动机命令")]
    [Attributes.ClassAttributes.Description("此命令允许您停止程序或进程。")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令可以按名称“chrome”关闭应用程序。 或者，您可以使用“关闭窗口”或“厚应用程序命令”。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'Process.CloseMainWindow'。")]
    public class StopwatchCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("输入秒表的实例名称")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("提供一个独特的实例或方式来参考秒表")]
        [Attributes.PropertyAttributes.SampleUsage("**myStopwatch**, **{vStopWatch}**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_StopwatchName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("输入秒表操作")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("提供一个独特的实例或方式来参考秒表")]
        [Attributes.PropertyAttributes.SampleUsage("**myStopwatch**, **{vStopWatch}**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_StopwatchAction { get; set; }


        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("将结果应用于变量")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_userVariableName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("可选 - 指定字符串格式")]
        [Attributes.PropertyAttributes.InputSpecification("指定是否需要特定的字符串格式。")]
        [Attributes.PropertyAttributes.SampleUsage("MM/dd/yy, hh:mm, etc.")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_ToStringFormat { get; set; }

        [XmlIgnore]
        [NonSerialized]
        public ComboBox StopWatchComboBox;

        [XmlIgnore]
        [NonSerialized]
        public List<Control> MeasureControls;

        public StopwatchCommand_cn()
        {
            this.CommandName = "StopwatchCommand";
            // this.SelectionName = "Stopwatch";
            this.SelectionName = Settings.Default.Stopwatch_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
            this.v_StopwatchName = "default";
            this.v_StopwatchAction = "Start Stopwatch";
        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;   
            var instanceName = v_StopwatchName.ConvertToUserVariable(sender);
            System.Diagnostics.Stopwatch stopwatch;

            var action = v_StopwatchAction.ConvertToUserVariable(sender);

            switch (action)
            {
                case "Start Stopwatch":
                    //start a new stopwatch
                    stopwatch = new System.Diagnostics.Stopwatch();
                    engine.AddAppInstance(instanceName, stopwatch);
                    stopwatch.Start();
                    break;
                case "Stop Stopwatch":
                    //stop existing stopwatch
                    stopwatch = (System.Diagnostics.Stopwatch)engine.AppInstances[instanceName];
                    stopwatch.Stop();
                    break;
                case "Restart Stopwatch":
                    //restart which sets to 0 and automatically starts
                    stopwatch = (System.Diagnostics.Stopwatch)engine.AppInstances[instanceName];
                    stopwatch.Restart();
                    break;
                case "Reset Stopwatch":
                    //reset which sets to 0
                    stopwatch = (System.Diagnostics.Stopwatch)engine.AppInstances[instanceName];
                    stopwatch.Reset();
                    break;
                case "Measure Stopwatch":
                    //check elapsed which gives measure
                    stopwatch = (System.Diagnostics.Stopwatch)engine.AppInstances[instanceName];
                    string elapsedTime;
                    if (string.IsNullOrEmpty(v_ToStringFormat))
                    {
                        elapsedTime = stopwatch.Elapsed.ToString();
                    }
                    else
                    {
                        var format = v_ToStringFormat.ConvertToUserVariable(sender);
                        elapsedTime = stopwatch.Elapsed.ToString(format);
                    }

                    elapsedTime.StoreInUserVariable(sender, v_userVariableName);

                    break;
                default:
                    throw new NotImplementedException("Stopwatch Action '" + action + "' not implemented");
            }



        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_StopwatchName", this, editor));

            var StopWatchComboBoxLabel = CommandControls.CreateDefaultLabelFor("v_StopwatchAction", this);
            StopWatchComboBox = (ComboBox)CommandControls.CreateDropdownFor("v_StopwatchAction", this);
            StopWatchComboBox.DataSource = new List<string> { "Start Stopwatch", "Stop Stopwatch", "Restart Stopwatch", "Reset Stopwatch", "Measure Stopwatch" };
            StopWatchComboBox.SelectedIndexChanged += StopWatchComboBox_SelectedValueChanged;
            RenderedControls.Add(StopWatchComboBoxLabel);
            RenderedControls.Add(StopWatchComboBox);

            //create controls for user variable
            MeasureControls = CommandControls.CreateDefaultDropdownGroupFor("v_userVariableName", this, editor);

            //load variables for selection
            var comboBox = (ComboBox)MeasureControls[1];
            comboBox.AddVariableNames(editor);

            MeasureControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_ToStringFormat", this, editor));

            foreach (var ctrl in MeasureControls)
            {
                ctrl.Visible = false;
            }
            RenderedControls.AddRange(MeasureControls);

            return RenderedControls;
        }

        private void StopWatchComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
           
            if (StopWatchComboBox.SelectedValue.ToString() == "Measure Stopwatch")
            {
                foreach (var ctrl in MeasureControls)
                                 ctrl.Visible = true;
               
            }
            else {
                foreach (var ctrl in MeasureControls)
                    ctrl.Visible = false;
            }
            
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Action: " + v_StopwatchAction + ", Name: " + v_StopwatchName + "]";
        }
    }
}