using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.Core.Automation.User32;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{

    [Serializable]
    [Attributes.ClassAttributes.Group("(Input Commands)输入命令")]
    [Attributes.ClassAttributes.Description("模拟鼠标点击。")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令可以模拟多种类型的鼠标单击。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令从user32.dll实现'SetCursorPos'功能以实现自动化。")]
    public class SendMouseClickCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指明鼠标点击类型")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("左键单击")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("中间点击")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("右键点击")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("左下")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("中下来")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("右下")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("左上")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("中间")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("马上")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("双击左键")]
        [Attributes.PropertyAttributes.InputSpecification("指出所需的点击类型")]
        [Attributes.PropertyAttributes.SampleUsage("选择**左键单击**，**中键点击**，**右键单击**，**左键单击**，**左下**，**中下**，**右下* *，**左上**，**中上**，**右上**")]
        [Attributes.PropertyAttributes.Remarks("您可以通过连续使用多个鼠标单击命令来模拟自定义单击，在需要的位置之间添加**暂停命令**。")]
        public string v_MouseClick { get; set; }

        public SendMouseClickCommand_cn()
        {
            this.CommandName = "SendMouseClickCommand";
            //this.SelectionName = "Send Mouse Click";
            this.SelectionName = Settings.Default.Send_Mouse_Click_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var mousePosition = System.Windows.Forms.Cursor.Position;
            User32Functions.SendMouseClick(v_MouseClick, mousePosition.X, mousePosition.Y);
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create window name helper control
            RenderedControls.AddRange(UI.CustomControls.CommandControls.CreateDefaultDropdownGroupFor("v_MouseClick", this, editor));
       

            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Click Type: " + v_MouseClick + "]";
        }
    }
}