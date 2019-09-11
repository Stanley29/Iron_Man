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
    [Attributes.ClassAttributes.Description("此命令将窗口移动到屏幕上的指定位置。")]
    [Attributes.ClassAttributes.UsesDescription("如果要将现有窗口按名称移动到屏幕上的某个点，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从user32.dll实现'FindWindowNative'，'SetWindowPos'以实现自动化。")]
    public class MoveWindowCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyDescription("请输入或选择要移动的窗口。")]
        [Attributes.PropertyAttributes.InputSpecification("输入或键入要移动的窗口的名称。")]
        [Attributes.PropertyAttributes.SampleUsage("**Untitled - Notepad**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_WindowName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyDescription("请指出窗口位置的新X水平坐标（像素）。 0从屏幕左侧开始。")]
        [Attributes.PropertyAttributes.InputSpecification("输入窗口的新水平坐标，0从左侧开始，向右侧移动")]
        [Attributes.PropertyAttributes.SampleUsage("0")]
        [Attributes.PropertyAttributes.Remarks("此数字是屏幕上的像素位置。 最大值应该是您的分辨率允许的最大值。 对于1920x1080，有效范围可以是0-1920")]
        public string v_XWindowPosition { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyDescription("请指出窗口位置的新Y垂直坐标（像素）。 0从屏幕顶部开始。")]
        [Attributes.PropertyAttributes.InputSpecification("输入窗口的新垂直坐标，0从顶部开始向下")]
        [Attributes.PropertyAttributes.SampleUsage("0")]
        [Attributes.PropertyAttributes.Remarks("此数字是屏幕上的像素位置。 最大值应该是您的分辨率允许的最大值。 对于1920x1080，有效范围可以是0-1080")]
        public string v_YWindowPosition { get; set; }

        [XmlIgnore]
        [NonSerialized]
        public ComboBox WindowNameControl;


        public MoveWindowCommand_cn()
        {
            this.CommandName = "MoveWindowCommand";
            // this.SelectionName = "Move Window";
            this.SelectionName = Settings.Default.Move_Window_cn;
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
                var variableXPosition = v_XWindowPosition.ConvertToUserVariable(sender);
                var variableYPosition = v_YWindowPosition.ConvertToUserVariable(sender);

                if (!int.TryParse(variableXPosition, out int xPos))
                {
                    throw new Exception("X Position Invalid - " + v_XWindowPosition);
                }
                if (!int.TryParse(variableYPosition, out int yPos))
                {
                    throw new Exception("X Position Invalid - " + v_XWindowPosition);
                }


                User32Functions.SetWindowPosition(targetedWindow, xPos, yPos);
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

            var xGroup = CommandControls.CreateDefaultInputGroupFor("v_XWindowPosition", this, editor);
            var yGroup = CommandControls.CreateDefaultInputGroupFor("v_YWindowPosition", this, editor);
            RenderedControls.AddRange(xGroup);
            RenderedControls.AddRange(yGroup);

            //RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_XWindowPosition", this));
            //var xPositionControl = CommandControls.CreateDefaultInputFor("v_XWindowPosition", this);
            //RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_XWindowPosition", this, new Control[] { xPositionControl }, editor));
            //RenderedControls.Add(xPositionControl);

            //RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_YWindowPosition", this));
            //var yPositionControl = CommandControls.CreateDefaultInputFor("v_YWindowPosition", this);
            //RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_YWindowPosition", this, new Control[] { yPositionControl }, editor));
            //RenderedControls.Add(yPositionControl);


            return RenderedControls;

        }
        public override void Refresh(frmCommandEditor editor)
        {
            base.Refresh();
            WindowNameControl.AddWindowNames();
        }


        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Target Window: " + v_WindowName + ", Target Coordinates (" + v_XWindowPosition + "," + v_YWindowPosition + ")]";
        }
    }
}