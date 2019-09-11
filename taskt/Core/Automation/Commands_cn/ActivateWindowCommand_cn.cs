using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.Core.Automation.User32;
using taskt.UI.Forms;
using taskt.UI.CustomControls;
using taskt.Properties;
using taskt.Core.Automation.Commands;
namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Window Commands)窗口命令")]
    [Attributes.ClassAttributes.Description("此命令激活窗口并将其带到前面。")]
    [Attributes.ClassAttributes.UsesDescription("如果要按名称激活窗口或引起注意，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从user32.dll实现'FindWindowNative'，'SetForegroundWindow'，'ShowWindow'以实现自动化。")]
    public class ActivateWindowCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入或选择要激活的窗口。")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或键入要激活或提前的窗口的名称。")]
        [Attributes.PropertyAttributes.SampleUsage("**Untitled - Notepad**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_WindowName { get; set; }

        [XmlIgnore]
        [NonSerialized]
        public ComboBox WindowNameControl;

        public ActivateWindowCommand_cn()
        {
            this.CommandName = "ActivateWindowCommand";
            // this.SelectionName = "Activate Window";
            this.SelectionName = Settings.Default.Activate_Window_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }
        public override void RunCommand(object sender)
        {
            string windowName = v_WindowName.ConvertToUserVariable(sender);

            var targetWindows = User32Functions.FindTargetWindows(windowName);

            //loop each window
            foreach (var targetedWindow in targetWindows)
            {
                User32Functions.SetWindowState(targetedWindow, User32Functions.WindowState.SW_SHOWNORMAL);
                User32Functions.SetForegroundWindow(targetedWindow);
            }

        }

        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create window name helper control
            RenderedControls.Add(UI.CustomControls.CommandControls.CreateDefaultLabelFor("v_WindowName", this));
            WindowNameControl = UI.CustomControls.CommandControls.CreateStandardComboboxFor("v_WindowName", this).AddWindowNames();
            RenderedControls.AddRange(UI.CustomControls.CommandControls.CreateUIHelpersFor("v_WindowName", this, new Control[] { WindowNameControl }, editor));
            RenderedControls.Add(WindowNameControl);

            return RenderedControls;

        }
        public override void Refresh(frmCommandEditor editor)
        {
            base.Refresh();
            WindowNameControl.AddWindowNames();
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Target Window: " + v_WindowName + "]";
        }
    }
}