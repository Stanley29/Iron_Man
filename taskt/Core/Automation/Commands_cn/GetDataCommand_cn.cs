//Copyright (c) 2019 Jason Bayldon
//
//Licensed under the Apache License, Version 2.0 (the "License");
//you may not use this file except in compliance with the License.
//You may obtain a copy of the License at
//
//   http://www.apache.org/licenses/LICENSE-2.0
//
//Unless required by applicable law or agreed to in writing, software
//distributed under the License is distributed on an "AS IS" BASIS,
//WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//See the License for the specific language governing permissions and
//limitations under the License.
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Serialization;
using Newtonsoft.Json;
using taskt.Core.Server;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(Engine Commands)发动机命令")]
    [Attributes.ClassAttributes.Description("此命令允许您从tasktServer获取数据。")]
    [Attributes.ClassAttributes.UsesDescription("如果要从tasktServer检索数据，请使用此命令")]
    [Attributes.ClassAttributes.ImplementationDescription("")]
    public class GetDataCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请指出要检索的密钥的名称")]
        [Attributes.PropertyAttributes.InputSpecification("选择变量或提供输入值")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_KeyName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("指示是检索整个记录还是仅检索值")]
        [Attributes.PropertyAttributes.InputSpecification("根据所选的选项，可能需要包含元数据的整个记录。")]
        [Attributes.PropertyAttributes.SampleUsage("选择一个相关选项")]
        [Attributes.PropertyAttributes.Remarks("")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("检索价值")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("检索整个记录")]
        public string v_DataOption { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("选择要接收输出的变量")]
        [Attributes.PropertyAttributes.InputSpecification("从变量列表中选择或提供变量")]
        [Attributes.PropertyAttributes.SampleUsage("**vSomeVariable**")]
        [Attributes.PropertyAttributes.Remarks("如果您在运行时启用了设置**创建缺失变量**，则无需预先定义变量，但强烈建议您这样做。")]
        public string v_applyToVariableName { get; set; }

        public GetDataCommand_cn()
        {
            this.CommandName = "GetDataCommand";
            // this.SelectionName = "Get BotStore Data";
            this.SelectionName = Settings.Default.Get_BotStore_Data_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }
        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;

            var keyName = v_KeyName.ConvertToUserVariable(sender);
            var dataOption = v_DataOption.ConvertToUserVariable(sender);

            BotStoreRequest.RequestType requestType;
            if (dataOption == "Retrieve Entire Record")
            {
                requestType = BotStoreRequest.RequestType.BotStoreModel;
            }
            else
            {
                requestType = BotStoreRequest.RequestType.BotStoreValue;
            }


            try
            {
                var result = HttpServerClient.GetData(keyName, requestType);

                if (requestType == BotStoreRequest.RequestType.BotStoreValue)
                {
                    result = JsonConvert.DeserializeObject<string>(result);
                }


                result.StoreInUserVariable(sender, v_applyToVariableName);
            }
            catch (Exception ex)
            {
                throw ex;
            }



        }

        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_KeyName", this, editor));

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_DataOption", this));
            var dropdown = CommandControls.CreateDropdownFor("v_DataOption", this);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_DataOption", this, new Control[] { dropdown }, editor));
            RenderedControls.Add(dropdown);


            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_applyToVariableName", this));
            var VariableNameControl = CommandControls.CreateStandardComboboxFor("v_applyToVariableName", this).AddVariableNames(editor);
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_applyToVariableName", this, new Control[] { VariableNameControl }, editor));
            RenderedControls.Add(VariableNameControl);

            return RenderedControls;
        }


        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Get Data from Key '" + v_KeyName + "' in tasktServer BotStore]";
        }
    }




}







