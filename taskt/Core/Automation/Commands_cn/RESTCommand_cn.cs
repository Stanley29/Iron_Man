using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using RestSharp;
using System.Data;
using System.Drawing;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{

    [Serializable]
    [Attributes.ClassAttributes.Group("(API Commands)API命令")]
    [Attributes.ClassAttributes.Description("此命令允许您向用户显示消息。")]
    [Attributes.ClassAttributes.UsesDescription("如果要在屏幕上向用户显示或显示值，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现'MessageBox'并调用VariableCommand以查找变量数据。")]
    public class RESTCommand_cn : ScriptCommand
    {
        //[XmlAttribute]
        //[Attributes.PropertyAttributes.PropertyDescription("Please enter an instance name")]
        //[Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        //[Attributes.PropertyAttributes.InputSpecification("")]
        //[Attributes.PropertyAttributes.SampleUsage("")]
        //[Attributes.PropertyAttributes.Remarks("")]
        //public string v_InstanceName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入基本URL（例如http://mysite.com)")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("定义包含完整URL的任何API端点。")]
        [Attributes.PropertyAttributes.SampleUsage("**https://example.com** or **{vMyUrl}**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_BaseURL { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入端点（例如/v2/endpoint")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.InputSpecification("定义包含完整URL的任何API端点")]
        [Attributes.PropertyAttributes.SampleUsage("**/v2/getUser/1** or **{vMyUrl}**")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_APIEndPoint { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择方法类型")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("GET")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("POST")]
        [Attributes.PropertyAttributes.InputSpecification("选择必要的方法类型。")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_APIMethodType { get; set; }


        [Attributes.PropertyAttributes.PropertyDescription("设置搜索参数")]
        [Attributes.PropertyAttributes.InputSpecification("使用元素记录器生成潜在搜索参数的列表。")]
        [Attributes.PropertyAttributes.SampleUsage("n/a")]
        [Attributes.PropertyAttributes.Remarks("单击有效窗口后，将填充搜索参数。 仅启用运行时所需的匹配项。")]
        public DataTable v_RESTParameters { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("将结果应用于变量")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_userVariableName { get; set; }


        [XmlIgnore]
        [NonSerialized]
        private DataGridView RESTParametersGridViewHelper;

        public RESTCommand_cn()
        {
            this.CommandName = "RESTCommand";
            //this.SelectionName = "Execute REST API";
            this.SelectionName = Settings.Default.Execute_REST_API_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;

            this.v_RESTParameters = new DataTable();
            this.v_RESTParameters.Columns.Add("Parameter Type");
            this.v_RESTParameters.Columns.Add("Parameter Name");
            this.v_RESTParameters.Columns.Add("Parameter Value");
            this.v_RESTParameters.TableName = DateTime.Now.ToString("RESTParamTable" + DateTime.Now.ToString("MMddyy.hhmmss"));

        }

        public override void RunCommand(object sender)
        {
           // var engine = (Engine.AutomationEngineInstance)sender;

            try
            {
                //make REST Request and store into variable
                var restContent = ExecuteRESTRequest(sender);
                restContent.StoreInUserVariable(sender, v_userVariableName);
            }
            catch (Exception ex)
            {
                throw ex;
            }
         
            

        }
        public string ExecuteRESTRequest(object sender)
        {
            //get parameters
            var targetURL = v_BaseURL.ConvertToUserVariable(sender);
            var targetEndpoint = v_APIEndPoint.ConvertToUserVariable(sender);
            var targetMethod = v_APIMethodType.ConvertToUserVariable(sender);

            //client
            var client = new RestClient(targetURL);

            //methods
            Method method = (Method)Enum.Parse(typeof(Method), targetMethod);
            
            //rest request
            var request = new RestRequest(targetEndpoint, method);

            //get parameters
            var apiParameters = (from rw in v_RESTParameters.AsEnumerable()
                                 where rw.Field<string>("Parameter Type") == "PARAMETER"
                                 select rw);

            //get headers
            var apiHeaders = (from rw in v_RESTParameters.AsEnumerable()
                              where rw.Field<string>("Parameter Type") == "HEADER"
                              select rw);

            //for each api parameter
            foreach (var param in apiParameters)
            {
                var paramName = ((string)param["Parameter Name"]).ConvertToUserVariable(sender);
                var paramValue = ((string)param["Parameter Value"]).ConvertToUserVariable(sender);

                request.AddParameter(paramName, paramValue);
            }

            //for each header
            foreach (var header in apiHeaders)
            {
                var paramName = ((string)header["Parameter Name"]).ConvertToUserVariable(sender);
                var paramValue = ((string)header["Parameter Value"]).ConvertToUserVariable(sender);

                request.AddHeader(paramName, paramValue);
            }

            //get json body
            var jsonBody = (from rw in v_RESTParameters.AsEnumerable()
                            where rw.Field<string>("Parameter Type") == "JSON BODY"
                            select rw.Field<string>("Parameter Value")).FirstOrDefault();
            //add json body
            if (jsonBody != null)
            {
                var json = jsonBody.ConvertToUserVariable(sender);
                request.AddJsonBody(jsonBody);
            }

            //get json body
            var file = (from rw in v_RESTParameters.AsEnumerable()
                        where rw.Field<string>("Parameter Type") == "FILE"
                        select rw).FirstOrDefault();

            //get file
            if (file != null)
            {
                var paramName = ((string)file["Parameter Name"]).ConvertToUserVariable(sender);
                var paramValue = ((string)file["Parameter Value"]).ConvertToUserVariable(sender);
                var fileData = paramValue.ConvertToUserVariable(sender);
                request.AddFile(paramName, fileData);

            }


            //execute client request
            IRestResponse response = client.Execute(request);
            var content = response.Content;

            // return response.Content;
            try
            {
                //try to parse and return json content
                var result = Newtonsoft.Json.JsonConvert.DeserializeObject(content);
                return result.ToString();
            }
            catch (Exception)
            {
                //content failed to parse simply return it
                return content;
            }
           

        }

        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

           var baseURLControlSet = CommandControls.CreateDefaultInputGroupFor("v_BaseURL", this, editor);
            RenderedControls.AddRange(baseURLControlSet);

            var apiEndPointControlSet = CommandControls.CreateDefaultInputGroupFor("v_APIEndPoint", this, editor);
            RenderedControls.AddRange(apiEndPointControlSet);
      

            var apiMethodLabel = CommandControls.CreateDefaultLabelFor("v_APIMethodType", this);
            var apiMethodDropdown = (ComboBox)CommandControls.CreateDropdownFor("v_APIMethodType", this);

            foreach (Method method in (Method[])Enum.GetValues(typeof(Method)))
            {
                apiMethodDropdown.Items.Add(method.ToString());
            }

     
            RenderedControls.Add(apiMethodLabel);
            RenderedControls.Add(apiMethodDropdown);

            RESTParametersGridViewHelper = new DataGridView();
            RESTParametersGridViewHelper.Width = 500;
            RESTParametersGridViewHelper.Height = 140;

            RESTParametersGridViewHelper.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            RESTParametersGridViewHelper.AutoGenerateColumns = false;

            var selectColumn = new DataGridViewComboBoxColumn();
            selectColumn.HeaderText = "Type";
            selectColumn.DataPropertyName = "Parameter Type";
            selectColumn.DataSource = new string[] { "HEADER", "PARAMETER", "JSON BODY", "FILE" };
            RESTParametersGridViewHelper.Columns.Add(selectColumn);

            var paramNameColumn = new DataGridViewTextBoxColumn();
            paramNameColumn.HeaderText = "Name";
            paramNameColumn.DataPropertyName = "Parameter Name";
            RESTParametersGridViewHelper.Columns.Add(paramNameColumn);

            var paramValueColumn = new DataGridViewTextBoxColumn();
            paramValueColumn.HeaderText = "Value";
            paramValueColumn.DataPropertyName = "Parameter Value";
            RESTParametersGridViewHelper.Columns.Add(paramValueColumn);


            RESTParametersGridViewHelper.DataBindings.Add("DataSource", this, "v_RESTParameters", false, DataSourceUpdateMode.OnPropertyChanged);
            RenderedControls.Add(RESTParametersGridViewHelper);

            //create local helper
            //taskt.UI.CustomControls.CommandItemControl helperControl = new taskt.UI.CustomControls.CommandItemControl();
            //helperControl.Padding = new System.Windows.Forms.Padding(10, 0, 0, 0);
            //helperControl.ForeColor = Color.AliceBlue;
            //helperControl.Font = new Font("Segoe UI Semilight", 10);
            //helperControl.Name = "addRow_helper";
            //helperControl.Tag = RESTParametersGridViewHelper;
            //helperControl.CommandDisplay = "Test API";
            //helperControl.Click += HelperControl_Click;
            //RenderedControls.Add(helperControl);

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_userVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_userVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_userVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;

        }


        //private void HelperControl_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        var engine = new Engine.AutomationEngineInstance();
        //        var result = this.ExecuteRESTRequest(engine);
        //        MessageBox.Show("Result Received: " + result);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("An Error Occured: " + ex.ToString());
        //    }
        //}

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [URL: " + v_BaseURL +  v_APIEndPoint + "]";
        }

    }

}

