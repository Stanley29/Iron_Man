using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using Newtonsoft.Json;
using OpenQA.Selenium;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Web Browser Commands)Web浏览器命令")]
    [Attributes.ClassAttributes.Description("此命令允许您关闭Selenium Web浏览器会话。")]
    [Attributes.ClassAttributes.UsesDescription("如果要在Web浏览器中操作，设置或获取网页上的数据，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现Selenium以实现自动化。")]
    public class SeleniumBrowserElementActionCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入** Create Browser **命令中指定的唯一实例名称")]
        [Attributes.PropertyAttributes.SampleUsage("**myInstance** or **seleniumInstance**")]
        [Attributes.PropertyAttributes.Remarks("未能输入正确的实例名称或未能首先调用** Create Browser **命令将导致错误")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("元素搜索方法")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("通过XPath查找元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("按ID查找元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("按名称查找元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("按标签名称查找元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("按类名查找元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("通过CSS选择器查找元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("通过XPath查找元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("按ID查找元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("按名称查找元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("按标签名称查找元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("按类名查找元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("通过CSS选择器查找元素")]
        [Attributes.PropertyAttributes.InputSpecification("选择要用于隔离网页中元素的特定搜索类型。")]
        [Attributes.PropertyAttributes.SampleUsage("选择**按XPath查找元素**，**按ID ID查找元素**，按名称查找元素**，**按标记名称查找元素**，**按类名查找元素**，** 通过CSS选择器查找元素**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_SeleniumSearchType { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("Element Search Parameter")]
        [Attributes.PropertyAttributes.InputSpecification("Specifies the parameter text that matches to the element based on the previously selected search type.")]
        [Attributes.PropertyAttributes.SampleUsage("If search type **Find Element By ID** was specified, for example, given <div id='name'></div>, the value of this field would be **name**")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_SeleniumSearchParameter { get; set; }
        [XmlElement]
        [Attributes.PropertyAttributes.PropertyDescription("元素搜索参数")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("调用Click")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("左键单击")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("右键点击")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("中间点击")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("双击左键")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("清除元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("设置文字")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("获取文字")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("获取属性")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("获取匹配元素")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("等待元素存在")]
        [Attributes.PropertyAttributes.InputSpecification("选择元素后要执行的相应操作")]
        [Attributes.PropertyAttributes.SampleUsage("选择**调用点击**，**左键单击**，**右键单击**，**中键单击**，**双左键单击**，**清除元素**，**设置文字* *，**获取文本**，**获取属性**，**等待元素存在**")]
        [Attributes.PropertyAttributes.Remarks("选择此字段将更改下一步中所需的参数")]
        public string v_SeleniumElementAction { get; set; }
        [XmlElement]
        [Attributes.PropertyAttributes.PropertyDescription("附加参数")]
        [Attributes.PropertyAttributes.InputSpecification("根据所选的操作设置，将需要其他参数。")]
        [Attributes.PropertyAttributes.SampleUsage("附加参数的范围从添加偏移坐标到指定要应用元素文本的变量。")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public System.Data.DataTable v_WebActionParameterTable { get; set; }


        [XmlIgnore]
        [NonSerialized]
        private DataGridView ElementsGridViewHelper;

        [XmlIgnore]
        [NonSerialized]
        private ComboBox ElementActionDropdown;

        [XmlIgnore]
        [NonSerialized]
        private List<Control> ElementParameterControls;

        public SeleniumBrowserElementActionCommand_cn()
        {
            this.CommandName = "SeleniumBrowserElementActionCommand";
            // this.SelectionName = "Element Action";
            this.SelectionName = Settings.Default.Element_Web_Action_cn;
            this.v_InstanceName = "default";
            this.CommandEnabled = true;
            this.CustomRendering = true;

            this.v_WebActionParameterTable = new System.Data.DataTable
            {
                TableName = "WebActionParamTable" + DateTime.Now.ToString("MMddyy.hhmmss")
            };
            this.v_WebActionParameterTable.Columns.Add("Parameter Name");
            this.v_WebActionParameterTable.Columns.Add("Parameter Value");

            ElementsGridViewHelper = new DataGridView();
            ElementsGridViewHelper.AllowUserToAddRows = true;
            ElementsGridViewHelper.AllowUserToDeleteRows = true;
            ElementsGridViewHelper.Size = new Size(400, 250);
            ElementsGridViewHelper.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            ElementsGridViewHelper.DataBindings.Add("DataSource", this, "v_WebActionParameterTable", false, DataSourceUpdateMode.OnPropertyChanged);
            ElementsGridViewHelper.AllowUserToAddRows = false;
            ElementsGridViewHelper.AllowUserToDeleteRows = false;
            ElementsGridViewHelper.AllowUserToResizeRows = false;
            //ElementsGridViewHelper.MouseEnter += ElementsGridViewHelper_MouseEnter;

        }

        //private void ElementsGridViewHelper_MouseEnter(object sender, EventArgs e)
        //{
        //    seleniumAction_SelectionChangeCommitted(null, null);
        //}

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
            //convert to user variable -- https://github.com/saucepleez/taskt/issues/66
            var seleniumSearchParam = v_SeleniumSearchParameter.ConvertToUserVariable(sender);


            var vInstance = v_InstanceName.ConvertToUserVariable(engine);

            var browserObject = engine.GetAppInstance(vInstance);

            var seleniumInstance = (OpenQA.Selenium.IWebDriver)browserObject;

            dynamic element = null;

            if (v_SeleniumElementAction == "Wait For Element To Exist")
            {

                var timeoutText = (from rw in v_WebActionParameterTable.AsEnumerable()
                                   where rw.Field<string>("Parameter Name") == "Timeout (Seconds)"
                                   select rw.Field<string>("Parameter Value")).FirstOrDefault();

                timeoutText = timeoutText.ConvertToUserVariable(sender);

                int timeOut = Convert.ToInt32(timeoutText);

                var timeToEnd = DateTime.Now.AddSeconds(timeOut);

                while (timeToEnd >= DateTime.Now)
                {
                    try
                    {
                        element = FindElement(seleniumInstance, seleniumSearchParam);
                        break;
                    }
                    catch (Exception)
                    {
                        engine.ReportProgress("Element Not Yet Found... " + (timeToEnd - DateTime.Now).Seconds + "s remain");
                        System.Threading.Thread.Sleep(1000);
                    }
                }

                if (element == null)
                {
                    throw new Exception("Element Not Found");
                }

                return;
            }
            else
            {

                element = FindElement(seleniumInstance, seleniumSearchParam);
            }




            switch (v_SeleniumElementAction)
            {
                case "Invoke Click":
                    element.Click();
                    break;

                case "Left Click":
                case "Right Click":
                case "Middle Click":
                case "Double Left Click":


                    int userXAdjust = Convert.ToInt32((from rw in v_WebActionParameterTable.AsEnumerable()
                                                       where rw.Field<string>("Parameter Name") == "X Adjustment"
                                                       select rw.Field<string>("Parameter Value")).FirstOrDefault().ConvertToUserVariable(sender));

                    int userYAdjust = Convert.ToInt32((from rw in v_WebActionParameterTable.AsEnumerable()
                                                       where rw.Field<string>("Parameter Name") == "Y Adjustment"
                                                       select rw.Field<string>("Parameter Value")).FirstOrDefault().ConvertToUserVariable(sender));

                    var elementLocation = element.Location;
                    SendMouseMoveCommand newMouseMove = new SendMouseMoveCommand();
                    var seleniumWindowPosition = seleniumInstance.Manage().Window.Position;
                    newMouseMove.v_XMousePosition = (seleniumWindowPosition.X + elementLocation.X + 30 + userXAdjust).ToString(); // added 30 for offset
                    newMouseMove.v_YMousePosition = (seleniumWindowPosition.Y + elementLocation.Y + 130 + userYAdjust).ToString(); //added 130 for offset
                    newMouseMove.v_MouseClick = v_SeleniumElementAction;
                    newMouseMove.RunCommand(sender);
                    break;

                case "Set Text":

                    string textToSet = (from rw in v_WebActionParameterTable.AsEnumerable()
                                        where rw.Field<string>("Parameter Name") == "Text To Set"
                                        select rw.Field<string>("Parameter Value")).FirstOrDefault();


                    string clearElement = (from rw in v_WebActionParameterTable.AsEnumerable()
                                           where rw.Field<string>("Parameter Name") == "Clear Element Before Setting Text"
                                           select rw.Field<string>("Parameter Value")).FirstOrDefault();

                    if (clearElement == null)
                    {
                        clearElement = "No";
                    }

                    if (clearElement.ToLower() == "yes") 
                    {
                        element.Clear();
                    }


                    string[] potentialKeyPresses = textToSet.Split('{', '}');

                    Type seleniumKeys = typeof(OpenQA.Selenium.Keys);
                    System.Reflection.FieldInfo[] fields = seleniumKeys.GetFields(System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public);

                    //check if chunked string contains a key press command like {ENTER}
                    foreach (string chunkedString in potentialKeyPresses)
                    {
                        if (chunkedString == "")
                            continue;

                        if (fields.Any(f => f.Name == chunkedString))
                        {
                            string keyPress = (string)fields.Where(f => f.Name == chunkedString).FirstOrDefault().GetValue(null);
                            element.SendKeys(keyPress);


                        }
                        else
                        {
                            //convert to user variable - https://github.com/saucepleez/taskt/issues/22
                            var convertedChunk = chunkedString.ConvertToUserVariable(sender);
                            element.SendKeys(convertedChunk);
                        }
                    }

                    break;

                case "Get Text":
                case "Get Attribute":

                    string VariableName = (from rw in v_WebActionParameterTable.AsEnumerable()
                                           where rw.Field<string>("Parameter Name") == "Variable Name"
                                           select rw.Field<string>("Parameter Value")).FirstOrDefault();

                    string attributeName = (from rw in v_WebActionParameterTable.AsEnumerable()
                                            where rw.Field<string>("Parameter Name") == "Attribute Name"
                                            select rw.Field<string>("Parameter Value")).FirstOrDefault().ConvertToUserVariable(sender);

                    string elementValue;
                    if (v_SeleniumElementAction == "Get Text")
                    {
                        elementValue = element.Text;
                    }
                    else
                    {
                        elementValue = element.GetAttribute(attributeName);
                    }

                    elementValue.StoreInUserVariable(sender, VariableName);

                    break;
                case "Get Matching Elements":
                    var variableName = (from rw in v_WebActionParameterTable.AsEnumerable()
                                        where rw.Field<string>("Parameter Name") == "Variable Name"
                                        select rw.Field<string>("Parameter Value")).FirstOrDefault();

                    var requiredComplexVariable = engine.VariableList.Where(x => x.VariableName == variableName).FirstOrDefault();

                    if (requiredComplexVariable == null)
                    {
                        engine.VariableList.Add(new Script.ScriptVariable() { VariableName = variableName, CurrentPosition = 0 });
                        requiredComplexVariable = engine.VariableList.Where(x => x.VariableName == variableName).FirstOrDefault();
                    }


                    //set json settings
                    JsonSerializerSettings settings = new JsonSerializerSettings();
                    settings.Error = (serializer, err) => {
                        err.ErrorContext.Handled = true;
                    };

                    settings.Formatting = Formatting.Indented;

                    //create json list
                    List<string> jsonList = new List<string>();
                    foreach (OpenQA.Selenium.IWebElement item in element)
                    {
                        var json = Newtonsoft.Json.JsonConvert.SerializeObject(item, settings);
                        jsonList.Add(json);
                    }

                    requiredComplexVariable.VariableValue = jsonList;
                    requiredComplexVariable.CurrentPosition = 0;

                    break;
                case "Clear Element":
                    element.Clear();
                    break;
                default:
                    throw new Exception("Element Action was not found");
            }


        }

        private object FindElement(OpenQA.Selenium.IWebDriver seleniumInstance, string searchParameter)
        {
            object element = null;



            switch (v_SeleniumSearchType)
            {
                case "Find Element By XPath":
                    element = seleniumInstance.FindElement(By.XPath(searchParameter));
                    break;

                case "Find Element By ID":
                    element = seleniumInstance.FindElement(By.Id(searchParameter));
                    break;

                case "Find Element By Name":
                    element = seleniumInstance.FindElement(By.Name(searchParameter));
                    break;

                case "Find Element By Tag Name":
                    element = seleniumInstance.FindElement(By.TagName(searchParameter));
                    break;

                case "Find Element By Class Name":
                    element = seleniumInstance.FindElement(By.ClassName(searchParameter));
                    break;
                case "Find Element By CSS Selector":
                    element = seleniumInstance.FindElement(By.CssSelector(searchParameter));
                    break;
                case "Find Elements By XPath":
                    element = seleniumInstance.FindElements(By.XPath(searchParameter));
                    break;

                case "Find Elements By ID":
                    element = seleniumInstance.FindElements(By.Id(searchParameter));
                    break;

                case "Find Elements By Name":
                    element = seleniumInstance.FindElements(By.Name(searchParameter));
                    break;

                case "Find Elements By Tag Name":
                    element = seleniumInstance.FindElements(By.TagName(searchParameter));
                    break;

                case "Find Elements By Class Name":
                    element = seleniumInstance.FindElements(By.ClassName(searchParameter));
                    break;

                case "Find Elements By CSS Selector":
                    element = seleniumInstance.FindElements(By.CssSelector(searchParameter));
                    break;

                default:
                    throw new Exception("Element Search Type was not found: " + v_SeleniumSearchType);
            }

            return element;
        }

        public bool ElementExists(object sender, string searchType, string elementName)
        {
            //get engine reference
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
            var seleniumSearchParam = elementName.ConvertToUserVariable(sender);

            //get instance name
            var vInstance = v_InstanceName.ConvertToUserVariable(engine);

            //get stored app object
            var browserObject = engine.GetAppInstance(vInstance);

            //get selenium instance driver
            var seleniumInstance = (OpenQA.Selenium.Chrome.ChromeDriver)browserObject;

            try
            {
                //search for element
                var element = FindElement(seleniumInstance, seleniumSearchParam);

                //element exists
                return true;
            }
            catch (Exception)
            {
                //element does not exist
                return false;
            }

        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_SeleniumSearchType", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_SeleniumSearchParameter", this, editor));


            ElementActionDropdown = (ComboBox)CommandControls.CreateDropdownFor("v_SeleniumElementAction", this);
            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_SeleniumElementAction", this));
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_SeleniumElementAction", this, new Control[] { ElementActionDropdown }, editor));
            ElementActionDropdown.SelectionChangeCommitted += seleniumAction_SelectionChangeCommitted;
            
            RenderedControls.Add(ElementActionDropdown);

            ElementParameterControls = new List<Control>();
            ElementParameterControls.Add(CommandControls.CreateDefaultLabelFor("v_WebActionParameterTable", this));
            ElementParameterControls.AddRange(CommandControls.CreateUIHelpersFor("v_WebActionParameterTable", this, new Control[] { ElementsGridViewHelper }, editor));
            ElementParameterControls.Add(ElementsGridViewHelper);

            RenderedControls.AddRange(ElementParameterControls);

            return RenderedControls;
        }

        //public override void Refresh(UI.Forms.frmCommandEditor editor)
        //{
        //    //seleniumAction_SelectionChangeCommitted(null, null);
        //}

        public void seleniumAction_SelectionChangeCommitted(object sender, EventArgs e)
        {

            Core.Automation.Commands_cn.SeleniumBrowserElementActionCommand_cn cmd = (Core.Automation.Commands_cn.SeleniumBrowserElementActionCommand_cn)this;
            DataTable actionParameters = cmd.v_WebActionParameterTable;

            if (sender != null)
            {
                actionParameters.Rows.Clear();
            }


            switch (ElementActionDropdown.SelectedItem)
            {
                case "Invoke Click":
                case "Clear Element":

                    foreach (var ctrl in ElementParameterControls)
                    {
                        ctrl.Hide();
                    }

                    break;

                case "Left Click":
                case "Middle Click":
                case "Right Click":
                case "Double Left Click":
                    foreach (var ctrl in ElementParameterControls)
                    {
                        ctrl.Show();
                    }
                    if (sender != null)
                    {
                        actionParameters.Rows.Add("X Adjustment", 0);
                        actionParameters.Rows.Add("Y Adjustment", 0);
                    }
                    break;

                case "Set Text":
                    foreach (var ctrl in ElementParameterControls)
                    {
                        ctrl.Show();
                    }
                    if (sender != null)
                    {
                        actionParameters.Rows.Add("Text To Set");
                        actionParameters.Rows.Add("Clear Element Before Setting Text");
                    }

                    DataGridViewComboBoxCell comparisonComboBox = new DataGridViewComboBoxCell();
                    comparisonComboBox.Items.Add("Yes");
                    comparisonComboBox.Items.Add("No");

                    //assign cell as a combobox
                    if (sender != null)
                    {
                        ElementsGridViewHelper.Rows[1].Cells[1].Value = "No";
                    }
                    ElementsGridViewHelper.Rows[1].Cells[1] = comparisonComboBox;


                    break;

                case "Get Text":
                case "Get Matching Elements":
                    foreach (var ctrl in ElementParameterControls)
                    {
                        ctrl.Show();
                    }
                    if (sender != null)
                    {
                        actionParameters.Rows.Add("Variable Name");
                    }
                    break;

                case "Get Attribute":
                    foreach (var ctrl in ElementParameterControls)
                    {
                        ctrl.Show();
                    }
                    if (sender != null)
                    {
                        actionParameters.Rows.Add("Attribute Name");
                        actionParameters.Rows.Add("Variable Name");
                    }
                    break;

                case "Wait For Element To Exist":
                    foreach (var ctrl in ElementParameterControls)
                    {
                        ctrl.Show();
                    }
                    if (sender != null)
                    {
                        actionParameters.Rows.Add("Timeout (Seconds)");
                    }
                    break;
                default:
                    break;
            }

            ElementsGridViewHelper.DataSource = v_WebActionParameterTable;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [" + v_SeleniumSearchType + " and " + v_SeleniumElementAction + ", Instance Name: '" + v_InstanceName + "']";
        }
    }
}