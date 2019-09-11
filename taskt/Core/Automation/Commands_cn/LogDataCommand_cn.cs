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
    [Attributes.ClassAttributes.Group("(Data Commands)数据命令")]
    [Attributes.ClassAttributes.Description("此命令将数据记录到文件中。")]
    [Attributes.ClassAttributes.UsesDescription("如果要将自定义数据记录到文件以进行调试或分析，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'Thread.Sleep'以实现自动化。")]
    public class LogDataCommand_cn : ScriptCommand
    {


        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("选择现有日志文件或输入自定义名称。")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("发动机日志")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("指示应附加日志的文件名")]
        [Attributes.PropertyAttributes.SampleUsage("选择“引擎日志”或指定您自己的文件")]
        [Attributes.PropertyAttributes.Remarks("日期和时间将自动附加到文件名。 日志全部保存在taskt Root \\ Logs文件夹中")]
        public string v_LogFile { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入要记录的文本。")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("指示要保存的文本的值。")]
        [Attributes.PropertyAttributes.SampleUsage("第三步完成，[变量]等")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_LogText { get; set; }

        public LogDataCommand_cn()
        {
            this.CommandName = "LogDataCommand";
            //this.SelectionName = "Log Data";
            this.SelectionName = Settings.Default.Log_Data_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            //get text to log and log file name       
            var textToLog = v_LogText.ConvertToUserVariable(sender);
            var logFile = v_LogFile.ConvertToUserVariable(sender);

            //determine log file
            if (v_LogFile == "Engine Logs")
            {
                //log to the standard engine logs
                var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
                engine.engineLogger.Information(textToLog);
            }
            else
            {
                //create new logger and log to custom file
                using (var logger = new Core.Logging().CreateLogger(logFile, Serilog.RollingInterval.Infinite))
                {
                    logger.Information(textToLog);
                }
            }


        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_LogFile", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_LogText", this, editor));

            return RenderedControls;

        }


        public override string GetDisplayValue()
        {
            string logFileName;
            if (v_LogFile == "Engine Logs")
            {
                logFileName = "taskt Engine Logs.txt";
            }
            else
            {
                logFileName = "taskt " + v_LogFile + " Logs.txt";
            }


            return base.GetDisplayValue() + " [Write Log to 'taskt\\Logs\\" + logFileName + "']";
        }
    }
}