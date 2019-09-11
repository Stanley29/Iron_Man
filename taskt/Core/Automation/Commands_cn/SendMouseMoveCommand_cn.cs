using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.Core.Automation.User32;
using taskt.Core.Script;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{

    [Serializable]
    [Attributes.ClassAttributes.Group("(Input Commands)输入命令")]
    [Attributes.ClassAttributes.Description("模拟鼠标移动")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令可以模拟鼠标的移动，此外，此命令还允许您在移动完成后执行单击。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从user32.dll实现'SetCursorPos'功能以实现自动化。")]
    public class SendMouseMoveCommand : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入X位置以将鼠标移动到")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowMouseCaptureHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入鼠标的新水平坐标，0从左侧开始，向右侧移动")]
        [Attributes.PropertyAttributes.SampleUsage("0")]
        [Attributes.PropertyAttributes.Remarks("此数字是屏幕上的像素位置。 最大值应该是您的分辨率允许的最大值。 对于1920x1080，有效范围可以是0-1920")]
        public string v_XMousePosition { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入Y位置以将鼠标移动到")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowMouseCaptureHelper)]
        [Attributes.PropertyAttributes.InputSpecification("输入窗口的新水平坐标，0从左侧开始向下")]
        [Attributes.PropertyAttributes.SampleUsage("0")]
        [Attributes.PropertyAttributes.Remarks("此数字是屏幕上的像素位置。 最大值应该是您的分辨率允许的最大值。 对于1920x1080，有效范围可以是0-1080")]
        public string v_YMousePosition { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("如果需要，请指明鼠标点击类型")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("没有")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("左键单击")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("中间点击")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("右键点击")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("双击左键")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("左下")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("中下来")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("右下")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("左上")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("中间")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("马上")]
        [Attributes.PropertyAttributes.InputSpecification("指出所需的点击类型")]
        [Attributes.PropertyAttributes.SampleUsage("选择**左键单击**，**中键点击**，**右键单击**，**左键单击**，**左下**，**中下**，**右下* *，**左上**，**中上**，**右上**")]
        [Attributes.PropertyAttributes.Remarks("您可以通过连续使用多个鼠标单击命令来模拟自定义单击，在需要的位置之间添加**暂停命令**。")]
        public string v_MouseClick { get; set; }

        public SendMouseMoveCommand()
        {
            this.CommandName = "SendMouseMoveCommand";
            // this.SelectionName = "Send Mouse Move";
            this.SelectionName = Settings.Default.Send_Mouse_Move_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
           
        }

        public override void RunCommand(object sender)
        {
  

            var mouseX = v_XMousePosition.ConvertToUserVariable(sender);
            var mouseY = v_YMousePosition.ConvertToUserVariable(sender);



            try
            {
                var xLocation = Convert.ToInt32(Math.Floor(Convert.ToDouble(mouseX)));
                var yLocation = Convert.ToInt32(Math.Floor(Convert.ToDouble(mouseY)));

                User32Functions.SetCursorPosition(xLocation, yLocation);
                User32Functions.SendMouseClick(v_MouseClick, xLocation, yLocation);


            }
            catch (Exception ex)
            {
                throw new Exception("Error parsing input to int type (X: " + v_XMousePosition + ", Y:" + v_YMousePosition + ") " + ex.ToString());
            }

          



        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_XMousePosition", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_YMousePosition", this, editor));

            //create window name helper control
            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_MouseClick", this, editor));


            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Target Coordinates (" + v_XMousePosition + "," + v_YMousePosition + ") Click: " + v_MouseClick + "]";
        }

     
    }
}