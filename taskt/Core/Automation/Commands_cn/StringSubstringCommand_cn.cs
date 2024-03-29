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
    [Attributes.ClassAttributes.Description("此命令允许您修剪字符串")]
    [Attributes.ClassAttributes.UsesDescription("如果要选择文本或变量的子集，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令使用String.Substring方法实现自动化。")]
    public class StringSubstringCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择要修改的变量或文本")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_userVariableName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("从职位开始")]
        [Attributes.PropertyAttributes.InputSpecification("指示字符串中的起始位置")]
        [Attributes.PropertyAttributes.SampleUsage("指示字符串中的起始位置")]
        [Attributes.PropertyAttributes.Remarks("")]
        public int v_startIndex { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("可选 - 长度（-1保留余数）")]
        [Attributes.PropertyAttributes.InputSpecification("指示是否应该保留这么多字符")]
        [Attributes.PropertyAttributes.SampleUsage("-1保持余数，1开始索引后1位置等。")]
        [Attributes.PropertyAttributes.Remarks("")]
        public int v_stringLength { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择变量以接收更改")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }
        public StringSubstringCommand_cn()
        {
            this.CommandName = "StringSubstringCommand";
            // this.SelectionName = "Substring";
            this.SelectionName = Settings.Default.Substring_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
            v_stringLength = -1;
        }
        public override void RunCommand(object sender)
        {


            var variableName = v_userVariableName.ConvertToUserVariable(sender);

            //apply substring
            if (v_stringLength >= 0)
            {
                variableName = variableName.Substring(v_startIndex, v_stringLength);
            }
            else
            {
                variableName = variableName.Substring(v_startIndex);
            }

            variableName.StoreInUserVariable(sender, v_applyToVariableName);
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_userVariableName", this));
            var userVariableName = CommandControls.CreateStandardComboboxFor("v_userVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_userVariableName", this, new Control[] { userVariableName }, editor));
            RenderedControls.Add(userVariableName);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_startIndex", this, editor));

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_stringLength", this, editor));

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Apply Substring to '" + v_userVariableName + "']";
        }
    }
}