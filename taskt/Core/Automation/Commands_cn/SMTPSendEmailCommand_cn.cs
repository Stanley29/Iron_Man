using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Misc Commands)其他命令")]
    [Attributes.ClassAttributes.Description("此命令允许您使用SMTP协议发送电子邮件。")]
    [Attributes.ClassAttributes.UsesDescription("如果要发送电子邮件并有权访问SMTP服务器凭据以生成电子邮件，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现System.Net命名空间以实现自动化")]
    public class SMTPSendEmailCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("主机名")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("定义脚本应使用的主机/服务名称")]
        [Attributes.PropertyAttributes.SampleUsage("**smtp.gmail.com**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_SMTPHost { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("港口")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("定义联系SMTP服务时应使用的端口号")]
        [Attributes.PropertyAttributes.SampleUsage("**587**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public int v_SMTPPort { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("用户名")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("定义联系SMTP服务时要使用的用户名")]
        [Attributes.PropertyAttributes.SampleUsage("**username**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_SMTPUserName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("密码")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("定义联系SMTP服务时要使用的密码")]
        [Attributes.PropertyAttributes.SampleUsage("**password**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_SMTPPassword { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("来自电邮")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("指定“发件人”字段的显示方式。")]
        [Attributes.PropertyAttributes.SampleUsage("myRobot@company.com")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_SMTPFromEmail { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("发邮件")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("指定应解决的目标电子邮件。")]
        [Attributes.PropertyAttributes.SampleUsage("jason@company.com")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_SMTPToEmail { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("学科")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("定义电子邮件应具有的文本主题（或变量）。")]
        [Attributes.PropertyAttributes.SampleUsage("**Alert!** or **[vStatus]**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_SMTPSubject { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("身体")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("指定应发送的消息。")]
        [Attributes.PropertyAttributes.SampleUsage("**Everything ran ok at [DateTime.Now]**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_SMTPBody { get; set; }
        public SMTPSendEmailCommand_cn()
        {
            this.CommandName = "SMTPCommand";
            // this.SelectionName = "Send SMTP Email";
            this.SelectionName = Settings.Default.Send_SMTP_Email_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            string varSMTPHost = v_SMTPHost.ConvertToUserVariable(sender);
            string varSMTPPort = v_SMTPPort.ToString().ConvertToUserVariable(sender);
            string varSMTPUserName = v_SMTPUserName.ConvertToUserVariable(sender);
            string varSMTPPassword = v_SMTPPassword.ConvertToUserVariable(sender);

            string varSMTPFromEmail = v_SMTPFromEmail.ConvertToUserVariable(sender);
            string varSMTPToEmail = v_SMTPToEmail.ConvertToUserVariable(sender);
            string varSMTPSubject = v_SMTPSubject.ConvertToUserVariable(sender);
            string varSMTPBody = v_SMTPBody.ConvertToUserVariable(sender);

            var client = new SmtpClient(varSMTPHost, int.Parse(varSMTPPort))
            {
                Credentials = new System.Net.NetworkCredential(varSMTPUserName, varSMTPPassword),
                EnableSsl = true
            };

            client.Send(varSMTPFromEmail, varSMTPToEmail, varSMTPSubject, varSMTPBody);
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SMTPHost", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SMTPPort", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SMTPUserName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SMTPPassword", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SMTPFromEmail", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SMTPToEmail", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SMTPSubject", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SMTPBody", this, editor));

            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [To Address: '" + v_SMTPToEmail + "']";
        }
    }
}
