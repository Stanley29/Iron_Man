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
    [Attributes.ClassAttributes.Group("(Window Commands)窗口命令")]
    [Attributes.ClassAttributes.Description("此命令等待窗口存在。")]
    [Attributes.ClassAttributes.UsesDescription("如果要在继续执行脚本之前显式等待窗口存在，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从user32.dll实现'FindWindowNative'，'ShowWindow'以实现自动化。")]
    public class WaitForWindowCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入或选择您要等待存在的窗口名称。")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或键入要等待存在的窗口的名称。")]
        [Attributes.PropertyAttributes.SampleUsage("**Untitled - Notepad**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_WindowName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("指示应在引发错误之前等待的秒数。")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("指定应在调用错误之前等待的秒数")]
        [Attributes.PropertyAttributes.SampleUsage("**5**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_LengthToWait { get; set; }

        [XmlIgnore]
        [NonSerialized]
        public ComboBox WindowNameControl;

        public WaitForWindowCommand_cn()
        {
            this.CommandName = "WaitForWindowCommand";
            //this.SelectionName = "Wait For Window To Exist";
            this.SelectionName = Settings.Default.Wait_For_Window_To_Exist_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var lengthToWait = v_LengthToWait.ConvertToUserVariable(sender);
            var waitUntil = int.Parse(lengthToWait);
            var endDateTime = DateTime.Now.AddSeconds(waitUntil);

            IntPtr hWnd = IntPtr.Zero;

            while (DateTime.Now < endDateTime)
            {
                string windowName = v_WindowName.ConvertToUserVariable(sender);
                hWnd = User32Functions.FindWindow(windowName);

                if (hWnd != IntPtr.Zero) //If found
                    break;

                System.Threading.Thread.Sleep(1000);

            }

            if (hWnd == IntPtr.Zero)
            {
                throw new Exception("Window was not found in the allowed time!");
            }
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create window name helper control
            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_WindowName", this));
            WindowNameControl = CommandControls.CreateStandardComboboxFor("v_WindowName", this).AddWindowNames();
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_WindowName", this, new Control[] { WindowNameControl }, editor));
            RenderedControls.Add(WindowNameControl);

            //create standard group controls
            var lengthToWaitControlSet = CommandControls.CreateDefaultInputGroupFor("v_LengthToWait", this, editor);
            RenderedControls.AddRange(lengthToWaitControlSet);


            return RenderedControls;

        }
        public override void Refresh(frmCommandEditor editor)
        {
            base.Refresh();
            WindowNameControl.AddWindowNames();
        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Target Window: '" + v_WindowName + "', Wait Up To " + v_LengthToWait + " seconds]";
        }

    }
}