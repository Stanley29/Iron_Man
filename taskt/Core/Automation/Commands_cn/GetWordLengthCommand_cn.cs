﻿using System;
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
    [Attributes.ClassAttributes.Group("(Data Commands)数据命令")]
    [Attributes.ClassAttributes.Description("此命令允许您检索字符串或变量的长度。")]
    [Attributes.ClassAttributes.UsesDescription("如果要从JSON对象中提取数据，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("")]
    public class GetWordLengthCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("提供值或变量来查找长度（例如{vSomeVariable}）")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("选择或提供变量或文本值")]
        [Attributes.PropertyAttributes.SampleUsage("**Hello** or **vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_InputValue { get; set; }


        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择变量以接收长度")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }

        public GetWordLengthCommand_cn()
        {
            this.CommandName = "GetLengthCommand";
            //this.SelectionName = "Get Word Length";
            this.SelectionName = Settings.Default.Get_Word_Length_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;

        }

        public override void RunCommand(object sender)
        {
            //get engine
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            //get input value
            var stringRequiringLength = v_InputValue.ConvertToUserVariable(sender);

            //count number of words
            var stringLength = stringRequiringLength.Length;

            //store word count into variable
            stringLength.ToString().StoreInUserVariable(sender, v_applyToVariableName);

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InputValue", this, editor));


            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Apply Result to: '" + v_applyToVariableName + "']";
        }
    }
}