using System;
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
    [Attributes.ClassAttributes.Group("(API Commands)API命令")]
    [Attributes.ClassAttributes.Description("此命令处理HTML源对象")]
    [Attributes.ClassAttributes.UsesDescription("使用此命令从成功的** HTTP请求命令中提取数据**")]
    [Attributes.ClassAttributes.ImplementationDescription("TBD")]
    public class HTTPQueryResultCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("使用HTML选择变量")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_userVariableName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("XPath查询")]
        [Attributes.PropertyAttributes.InputSpecification("输入XPath查询，将提取项目。")]
        [Attributes.PropertyAttributes.SampleUsage("@//*[@id=\"aso_search_form_anchor\"]/div/input")]
        [Attributes.PropertyAttributes.Remarks("您可以使用Chrome开发者工具单击元素并复制XPath。")]
        public string v_xPathQuery { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("将查询结果应用于变量")]
        public string v_applyToVariableName { get; set; }


        public HTTPQueryResultCommand_cn()
        {
            this.CommandName = "HTTPRequestQueryCommand";
            //this.SelectionName = "HTTP Result Query";
            this.SelectionName = Settings.Default.HTTP_Result_Query_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(v_userVariableName.ConvertToUserVariable(sender));

            var div = doc.DocumentNode.SelectSingleNode(v_xPathQuery);
            string divString = div.InnerText;

            divString.StoreInUserVariable(sender, v_applyToVariableName);


        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //helper for user variable name
            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_userVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_userVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_userVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            //create xpath group
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_xPathQuery", this, editor));

            //apply to variable name
            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var applyToVariableControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { applyToVariableControl }, editor));
            RenderedControls.Add(applyToVariableControl);


            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Query Variable '" + v_userVariableName + "' and apply result to '" + v_applyToVariableName + "']";
        }
    }
}