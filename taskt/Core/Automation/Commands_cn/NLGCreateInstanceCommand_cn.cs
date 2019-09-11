using SimpleNLG;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(NLG Commands)NLG命令")]
    [Attributes.ClassAttributes.Description("此命令将脚本暂停一段指定的时间（以毫秒为单位）。")]
    [Attributes.ClassAttributes.UsesDescription("如果要在特定时间内暂停脚本，请使用此命令。 在指定的时间结束后，脚本将继续执行。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'Thread.Sleep'以实现自动化。")]
    public class NLGCreateInstanceCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入** Create NLG Instance **命令中指定的唯一实例名称")]
        [Attributes.PropertyAttributes.SampleUsage("**nlgDefaultInstance** or **myInstance**")]
        [Attributes.PropertyAttributes.Remarks("未能输入正确的实例名称或未能首次调用**创建NLG实例**命令将导致错误")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }

        public NLGCreateInstanceCommand_cn()
        {
            this.CommandName = "NLGCreateInstanceCommand";
            // this.SelectionName = "Create NLG Instance";
            this.SelectionName = Settings.Default.Create_NLG_Instance_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
            this.v_InstanceName = "nlgDefaultInstance";
        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
  
            Lexicon lexicon = Lexicon.getDefaultLexicon();
            NLGFactory nlgFactory = new NLGFactory(lexicon);
            SPhraseSpec p = nlgFactory.createClause();

            var vInstance = v_InstanceName.ConvertToUserVariable(sender);

            engine.AddAppInstance(vInstance, p);

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            
            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Instance Name: '" + v_InstanceName + "']";
        }
    }
}