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
    [Attributes.ClassAttributes.Group("(IE Browser Commands)IE浏览器命令")]
    [Attributes.ClassAttributes.Description("此命令允许您创建新的IE Web浏览器会话。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从SHDocVw.dll实现“InternetExplorer”应用程序对象以实现自动化。")]
    public class IEBrowserCreateCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        public string v_InstanceName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("实例跟踪（任务结束后）")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("忘记实例")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("保持实例活着")]
        [Attributes.PropertyAttributes.InputSpecification("指定在脚本执行完毕后taskt是否应记住此实例名称。")]
        [Attributes.PropertyAttributes.SampleUsage("选择**忘记实例**以忘记实例或**保持实例活动**以允许后续任务按名称调用实例。")]
        [Attributes.PropertyAttributes.Remarks("调用** Close Browser **命令或结束浏览器会话将结束实例。 此命令仅在应用程序的生命周期内有效。 如果应用程序关闭，将自动忘记引用。")]
        public string v_InstanceTracking { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入要导航到的URL")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_URL { get; set; }

        public IEBrowserCreateCommand_cn()
        {
            this.CommandName = "IEBrowserCreateCommand";
            //this.SelectionName = "Create Browser";
            this.SelectionName = Settings.Default.Create_Browser_cn;
            this.v_InstanceName = "default";
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var instanceName = v_InstanceName.ConvertToUserVariable(sender);

            SHDocVw.InternetExplorer newBrowserSession = new SHDocVw.InternetExplorer();
            try
            {
                newBrowserSession.Navigate(v_URL.ConvertToUserVariable(sender));
                WaitForReadyState(newBrowserSession);
                newBrowserSession.Visible = true;
            }
            catch (Exception ex) { }

            //add app instance
            engine.AddAppInstance(instanceName, newBrowserSession);

            //handle app instance tracking
            if (v_InstanceTracking == "Keep Instance Alive")
            {
                GlobalAppInstances.AddInstance(instanceName, newBrowserSession);
            }

        }

        public static void WaitForReadyState(SHDocVw.InternetExplorer ieInstance)
        {
            try
            {
                DateTime waitExpires = DateTime.Now.AddSeconds(15);

                do
                {
                    System.Threading.Thread.Sleep(500);
                }

                while ((ieInstance.ReadyState != SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE) && (waitExpires > DateTime.Now));
            }
            catch (Exception ex) { }
        }

        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_URL", this, editor));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Instance Name: '" + v_InstanceName + "']";
        }
    }

}