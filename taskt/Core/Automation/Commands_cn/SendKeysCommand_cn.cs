using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.Core.Automation.User32;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Input Commands)输入命令")]
    [Attributes.ClassAttributes.Description("将键击发送到目标窗口")]
    [Attributes.ClassAttributes.UsesDescription("如果要将击键输入发送到窗口，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'Windows.Forms.SendKeys'方法以实现自动化。")]
    public class SendKeysCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入窗口名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入或键入要激活或提前的窗口的名称。")]
        [Attributes.PropertyAttributes.SampleUsage("**Untitled - Notepad**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_WindowName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入要发送的文本")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入应发送到指定窗口的文本。")]
        [Attributes.PropertyAttributes.SampleUsage("**Hello, World!** or **[vEntryText]**")]
        [Attributes.PropertyAttributes.Remarks("此命令支持在括号[variable]内发送变量")]
        public string v_TextToSend { get; set; }

        public SendKeysCommand_cn()
        {
            this.CommandName = "SendKeysCommand";
            //this.SelectionName = "Send Keystrokes";
            this.SelectionName = Settings.Default.Send_Keystrokes_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            if (v_WindowName != "Current Window")
            {
                ActivateWindowCommand_cn activateWindow = new ActivateWindowCommand_cn
                {
                    v_WindowName = v_WindowName
                };
                activateWindow.RunCommand(sender);
            }

            string textToSend = v_TextToSend.ConvertToUserVariable(sender);


            if (textToSend == "{WIN_KEY}")
            {
                User32Functions.KeyDown(System.Windows.Forms.Keys.LWin);
                User32Functions.KeyUp(System.Windows.Forms.Keys.LWin);

            }
            else if (textToSend.Contains("{WIN_KEY+"))
            {
                User32Functions.KeyDown(System.Windows.Forms.Keys.LWin);
                var remainingText = textToSend.Replace("{WIN_KEY+", "").Replace("}","");

                foreach (var c in remainingText)
                {
                    System.Windows.Forms.Keys key = (System.Windows.Forms.Keys)Enum.Parse(typeof(System.Windows.Forms.Keys), c.ToString());
                    User32Functions.KeyDown(key);
                }

                User32Functions.KeyUp(System.Windows.Forms.Keys.LWin);

                foreach (var c in remainingText)
                {
                    System.Windows.Forms.Keys key = (System.Windows.Forms.Keys)Enum.Parse(typeof(System.Windows.Forms.Keys), c.ToString());
                    User32Functions.KeyUp(key);
                }
            }
            else
            {
                System.Windows.Forms.SendKeys.SendWait(textToSend);
            }


          

            System.Threading.Thread.Sleep(500);
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);


            RenderedControls.Add(UI.CustomControls.CommandControls.CreateDefaultLabelFor("v_WindowName", this));
            var WindowNameControl = UI.CustomControls.CommandControls.CreateStandardComboboxFor("v_WindowName", this).AddWindowNames();
            RenderedControls.AddRange(UI.CustomControls.CommandControls.CreateUIHelpersFor("v_WindowName", this, new Control[] { WindowNameControl }, editor));
            RenderedControls.Add(WindowNameControl);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_TextToSend", this, editor));


            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Send '" + v_TextToSend + "' to '" + v_WindowName + "']";
        }
    }
}