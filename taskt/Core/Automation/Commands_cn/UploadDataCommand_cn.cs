using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.Core.Server;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Engine Commands)发动机命令")]
    [Attributes.ClassAttributes.Description("此命令允许您将数据上载到本地tasktServer bot商店")]
    [Attributes.ClassAttributes.UsesDescription("如果要跨机器人上传或共享数据，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("")]
    public class UploadDataCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指出要创建的密钥的名称")]
        [Attributes.PropertyAttributes.InputSpecification("选择变量或提供输入值")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_KeyName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择要上传的目标变量或输入值")]
        [Attributes.PropertyAttributes.InputSpecification("选择变量或提供输入值")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InputValue { get; set; }

        public UploadDataCommand_cn()
        {
            this.CommandName = "UploadDataCommand";
            // this.SelectionName = "Upload BotStore Data";
            this.SelectionName = Settings.Default.Upload_BotStore_Data_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }
        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var keyName = v_KeyName.ConvertToUserVariable(sender);
            var keyValue = v_InputValue.ConvertToUserVariable(sender);

            try
            {
                var result = HttpServerClient.UploadData(keyName, keyValue);
            }
            catch (Exception ex)
            {
                throw ex;
            }



        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_KeyName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InputValue", this, editor));

            return RenderedControls;
        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Upload Data to Key '" + v_KeyName + "' in tasktServer BotStore]";
        }
    }
}