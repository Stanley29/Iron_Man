using System;
using System.Xml.Serialization;
using System.Net;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using taskt.UI.Forms;
using taskt.UI.CustomControls;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(API Commands)API命令")]
    [Attributes.ClassAttributes.Description("此命令下载用于解析的网页的HTML源")]
    [Attributes.ClassAttributes.UsesDescription("如果要从网页检索HTML，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现System.Web以实现自动化")]
    public class HTTPRequestCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入网址")]
        [Attributes.PropertyAttributes.InputSpecification("输入要从中收集数据的有效URL。")]
        [Attributes.PropertyAttributes.SampleUsage("http://mycompany.com/news or [vCompany]")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_WebRequestURL { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("将结果应用于变量")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_userVariableName { get; set; }


        [XmlIgnore]
        [NonSerialized]
        public ComboBox VariableNameControl;
        public HTTPRequestCommand_cn()
        {
            this.CommandName = "HTTPRequestCommand";
            // this.SelectionName = "HTTP Request";
            this.SelectionName = Settings.Default.HTTP_Request_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }

        public override void RunCommand(object sender)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(v_WebRequestURL.ConvertToUserVariable(sender));
            request.Method = "GET";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.2 (KHTML, like Gecko) Chrome/15.0.874.121 Safari/535.2";
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            Stream dataStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream);
            string strResponse = reader.ReadToEnd();

            strResponse.StoreInUserVariable(sender, v_userVariableName);

        }

        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create inputs for request url
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_WebRequestURL", this, editor));


            //create window name helper control
            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_userVariableName", this));
            VariableNameControl = CommandControls.CreateStandardComboboxFor("v_userVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_userVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Target URL: '" + v_WebRequestURL + "' and apply result to '" + v_userVariableName + "']";
        }

    }
}