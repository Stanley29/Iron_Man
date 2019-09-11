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
    [Attributes.ClassAttributes.Group("(Web Browser Commands)Web浏览器命令")]
    [Attributes.ClassAttributes.Description("此命令允许您创建新的Selenium Web浏览器会话，以实现网站的自动化。")]
    [Attributes.ClassAttributes.UsesDescription("如果要创建最终将执行Web自动化的浏览器（如检查内部公司Intranet站点以检索数据），请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现Selenium以实现自动化。")]
    public class SeleniumBrowserCreateCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.InputSpecification("表示将代表应用程序实例的唯一名称。 此唯一名称允许您在将来的命令中按名称引用实例，确保您指定的命令针对正确的应用程序运行。")]
        [Attributes.PropertyAttributes.SampleUsage("**myInstance** or **seleniumInstance**")]
        [Attributes.PropertyAttributes.Remarks("**myInstance** or **seleniumInstance**")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("实例跟踪（任务结束后）")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("忘记实例")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("保持实例活着")]
        [Attributes.PropertyAttributes.InputSpecification("指定在脚本执行完毕后taskt是否应该记住此实例名称。")]
        [Attributes.PropertyAttributes.SampleUsage("选择**忘记实例**以忘记实例或**保持实例活动**以允许后续任务按名称调用实例。")]
        [Attributes.PropertyAttributes.Remarks("调用** Close Browser **命令或结束浏览器会话将结束实例。 此命令仅在应用程序的生命周期内有效。 如果应用程序关闭，将自动忘记引用。")]
        public string v_InstanceTracking { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择一个窗口状态")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("正常")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("最大化")]
        [Attributes.PropertyAttributes.InputSpecification("选择浏览器应启动的窗口状态。")]
        [Attributes.PropertyAttributes.SampleUsage("选择** Normal **以正常模式启动浏览器或** Maximize **以最大化模式启动浏览器。")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_BrowserWindowOption { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指定Selenium命令行选项（可选）")]
        [Attributes.PropertyAttributes.InputSpecification("选择要传递给Selenium命令的可选选项。")]
        [Attributes.PropertyAttributes.SampleUsage("user-data-dir=/user/public/SeleniumTasktProfile")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_SeleniumOptions { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择浏览器引擎类型")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("Chrome")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("IE")]
        [Attributes.PropertyAttributes.InputSpecification("选择浏览器应启动的窗口状态。")]
        [Attributes.PropertyAttributes.SampleUsage("选择** Normal **以正常模式启动浏览器或** Maximize **以最大化模式启动浏览器。")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_EngineType { get; set; }

        public SeleniumBrowserCreateCommand_cn()
        {
            this.CommandName = "SeleniumBrowserCreateCommand";
            // this.SelectionName = "Create Browser";
            this.SelectionName = Settings.Default.Create_Web_Browser_cn;
            this.v_InstanceName = "default";
            this.CommandEnabled = true;
            this.CustomRendering = true;
            this.v_EngineType = "Chrome";
        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
            var driverPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath), "Resources");
            var seleniumEngine = v_EngineType.ConvertToUserVariable(sender);
            var instanceName = v_InstanceName.ConvertToUserVariable(sender);

            OpenQA.Selenium.DriverService driverService;
            OpenQA.Selenium.IWebDriver webDriver;

            if (seleniumEngine == "Chrome")
            {
                OpenQA.Selenium.Chrome.ChromeOptions options = new OpenQA.Selenium.Chrome.ChromeOptions();

                if (!string.IsNullOrEmpty(v_SeleniumOptions))
                {
                    var convertedOptions = v_SeleniumOptions.ConvertToUserVariable(sender);
                    options.AddArguments(convertedOptions);
                }

                driverService = OpenQA.Selenium.Chrome.ChromeDriverService.CreateDefaultService(driverPath);
                webDriver = new OpenQA.Selenium.Chrome.ChromeDriver((OpenQA.Selenium.Chrome.ChromeDriverService)driverService, options);
            }
            else
            {
                driverService = OpenQA.Selenium.IE.InternetExplorerDriverService.CreateDefaultService(driverPath);
                webDriver = new OpenQA.Selenium.IE.InternetExplorerDriver((OpenQA.Selenium.IE.InternetExplorerDriverService)driverService, new OpenQA.Selenium.IE.InternetExplorerOptions());
            }


            //add app instance
            engine.AddAppInstance(instanceName, webDriver);


            //handle app instance tracking
            if (v_InstanceTracking == "Keep Instance Alive")
            {
                GlobalAppInstances.AddInstance(instanceName, webDriver);
            }

            //handle window type on startup - https://github.com/saucepleez/taskt/issues/22
            switch (v_BrowserWindowOption)
            {
                case "Maximize":
                    webDriver.Manage().Window.Maximize();
                    break;
                case "Normal":
                case "":
                default:
                    break;
            }





        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);


            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_EngineType", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_InstanceTracking", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_BrowserWindowOption", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SeleniumOptions", this, editor));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return "Create " + v_EngineType + " Browser - [Instance Name: '" + v_InstanceName + "', Instance Tracking: " + v_InstanceTracking + "]";
        }
    }
}