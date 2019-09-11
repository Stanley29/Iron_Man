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
    [Attributes.ClassAttributes.Group("(Image Commands)图像命令")]
    [Attributes.ClassAttributes.Description("此命令采用屏幕截图并将其保存到某个位置")]
    [Attributes.ClassAttributes.UsesDescription("如果要拍摄并保存屏幕截图，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现User32 CaptureWindow以实现自动化")]
    public class ScreenshotCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入窗口名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入或键入要截取屏幕截图的窗口的名称。")]
        [Attributes.PropertyAttributes.SampleUsage("**Untitled - Notepad**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_ScreenshotWindowName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指出保存图像的路径")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowFileSelectionHelper)]
        public string v_FilePath { get; set; }
        public ScreenshotCommand_cn()
        {
            this.CommandName = "ScreenshotCommand";
            // this.SelectionName = "Take Screenshot";
            this.SelectionName = Settings.Default.Take_Screenshot_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            var image = User32Functions.CaptureWindow(v_ScreenshotWindowName);
            string ConvertToUserVariabledString = v_FilePath.ConvertToUserVariable(sender);
            image.Save(ConvertToUserVariabledString);
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create window name helper control
            RenderedControls.Add(UI.CustomControls.CommandControls.CreateDefaultLabelFor("v_ScreenshotWindowName", this));
            var WindowNameControl = UI.CustomControls.CommandControls.CreateStandardComboboxFor("v_ScreenshotWindowName", this).AddWindowNames();
            RenderedControls.AddRange(UI.CustomControls.CommandControls.CreateUIHelpersFor("v_ScreenshotWindowName", this, new Control[] { WindowNameControl }, editor));
            RenderedControls.Add(WindowNameControl);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_FilePath", this, editor));


            return RenderedControls;
        }


        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Target Window: '" + v_ScreenshotWindowName + "', File Path: '" + v_FilePath + "]";
        }
    }
}