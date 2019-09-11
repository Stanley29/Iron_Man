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
    [Attributes.ClassAttributes.Description("此命令将窗口大小调整为指定大小。")]
    [Attributes.ClassAttributes.UsesDescription("如果要在屏幕上按名称将窗口大小调整为特定大小，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从user32.dll实现'FindWindowNative'，'SetWindowPos'以实现自动化。")]
    public class ResizeWindowCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyDescription("请输入或选择要调整大小的窗口。")]
        [Attributes.PropertyAttributes.InputSpecification("输入或键入要调整大小的窗口的名称。")]
        [Attributes.PropertyAttributes.SampleUsage("**Untitled - Notepad**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_WindowName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyDescription("请指出窗口所需的新宽度（像素）。")]
        [Attributes.PropertyAttributes.InputSpecification("输入窗口的新宽度")]
        [Attributes.PropertyAttributes.SampleUsage("0")]
        [Attributes.PropertyAttributes.Remarks("此数字受您的决议限制。 最大值应该是您的分辨率允许的最大值。 对于1920x1080，有效宽度范围可以是0-1920")]
        public string v_XWindowSize { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyDescription("请指出窗口所需的新高度（像素）")]
        [Attributes.PropertyAttributes.InputSpecification("输入窗口的新高度")]
        [Attributes.PropertyAttributes.SampleUsage("0")]
        [Attributes.PropertyAttributes.Remarks("此数字受您的决议限制。 最大值应该是您的分辨率允许的最大值。 对于1920x1080，有效高度范围可以是0-1080")]
        public string v_YWindowSize { get; set; }

        [XmlIgnore]
        [NonSerialized]
        public ComboBox WindowNameControl;
        public ResizeWindowCommand_cn()
        {
            this.CommandName = "ResizeWindowCommand";
            // this.SelectionName = "Resize Window";
            this.SelectionName = Settings.Default.Resize_Window_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;    
        }

        public override void RunCommand(object sender)
        {
            string windowName = v_WindowName.ConvertToUserVariable(sender);

            var targetWindows = User32Functions.FindTargetWindows(windowName);

            //loop each window and set the window state
            foreach (var targetedWindow in targetWindows)
            {
                var variableXSize = v_XWindowSize.ConvertToUserVariable(sender);
                var variableYSize = v_YWindowSize.ConvertToUserVariable(sender);

                if (!int.TryParse(variableXSize, out int xPos))
                {
                    throw new Exception("X Position Invalid - " + v_XWindowSize);
                }
                if (!int.TryParse(variableYSize, out int yPos))
                {
                    throw new Exception("X Position Invalid - " + v_YWindowSize);
                }

                User32Functions.SetWindowSize(targetedWindow, xPos, yPos);
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
            var xGroup = CommandControls.CreateDefaultInputGroupFor("v_XWindowSize", this, editor);
            var yGroup = CommandControls.CreateDefaultInputGroupFor("v_YWindowSize", this, editor);
            RenderedControls.AddRange(xGroup);
            RenderedControls.AddRange(yGroup);
      
            return RenderedControls;

        }
        public override void Refresh(frmCommandEditor editor)
        {
            base.Refresh();
            WindowNameControl.AddWindowNames();
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Target Window: " + v_WindowName + ", Target Size (" + v_XWindowSize + "," + v_YWindowSize + ")]";
        }
    }
}