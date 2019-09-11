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
    [Attributes.ClassAttributes.Description("此命令设置目标窗口的状态。")]
    [Attributes.ClassAttributes.UsesDescription("如果要将窗口的状态更改为最小化，最大化或恢复状态，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从user32.dll实现'FindWindowNative'，'ShowWindow'以实现自动化。")]
    public class SetWindowStateCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入或选择要更改的目标窗口。")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入或键入要更改的窗口的名称。")]
        [Attributes.PropertyAttributes.SampleUsage("**Untitled - Notepad**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_WindowName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择窗口的新所需状态。")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("最大化")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("最小化")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("恢复")]
        [Attributes.PropertyAttributes.InputSpecification("选择所需的适当窗口状态")]
        [Attributes.PropertyAttributes.SampleUsage("选择**最小化**，**最大化**和**恢复**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_WindowState { get; set; }

        [XmlIgnore]
        [NonSerialized]
        public ComboBox WindowNameControl;

        public SetWindowStateCommand_cn()
        {
            this.CommandName = "SetWindowStateCommand";
            // this.SelectionName = "Set Window State";
            this.SelectionName = Settings.Default.Set_Window_State_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            //convert window name
            string windowName = v_WindowName.ConvertToUserVariable(sender);

            var targetWindows = User32Functions.FindTargetWindows(windowName);

            //loop each window and set the window state
            foreach (var targetedWindow in targetWindows)
            {
                User32Functions.WindowState WINDOW_STATE = User32Functions.WindowState.SW_SHOWNORMAL;
                switch (v_WindowState)
                {
                    case "Maximize":
                        WINDOW_STATE = User32Functions.WindowState.SW_MAXIMIZE;
                        break;

                    case "Minimize":
                        WINDOW_STATE = User32Functions.WindowState.SW_MINIMIZE;
                        break;

                    case "Restore":
                        WINDOW_STATE = User32Functions.WindowState.SW_RESTORE;
                        break;

                    default:
                        break;
                }

                User32Functions.SetWindowState(targetedWindow, WINDOW_STATE);
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

            var windowStateLabel = CommandControls.CreateDefaultLabelFor("v_WindowState", this);
            RenderedControls.Add(windowStateLabel);

            var windowStateControl = CommandControls.CreateDropdownFor("v_WindowState", this);
            RenderedControls.Add(windowStateControl);

            return RenderedControls;

        }
        public override void Refresh(frmCommandEditor editor)
        {
            base.Refresh();
            WindowNameControl.AddWindowNames();
        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Target Window: " + v_WindowName + ", Window State: " + v_WindowState + "]";
        }
    }
}