using SimpleNLG;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Xml.Serialization;
using taskt.UI.CustomControls;
using taskt.UI.Forms;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{
    [Serializable]
    [Attributes.ClassAttributes.Group("(NLG Commands)NLG命令")]
    [Attributes.ClassAttributes.Description("此命令允许您定义NLG参数")]
    [Attributes.ClassAttributes.UsesDescription("如果要定义NLG参数，请使用此命令")]
    public class NLGSetParameterCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入** Create NLG Instance **命令中指定的唯一实例名称")]
        [Attributes.PropertyAttributes.SampleUsage("**nlgDefaultInstance** or **myInstance**")]
        [Attributes.PropertyAttributes.Remarks("未能输入正确的实例名称或未能首次调用**创建NLG实例**命令将导致错误")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请选择NLG参数类型")]
        [Attributes.PropertyAttributes.InputSpecification("输入** Create NLG Instance **命令中指定的唯一实例名称")]
        [Attributes.PropertyAttributes.SampleUsage("**nlgDefaultInstance** or **myInstance**")]
        [Attributes.PropertyAttributes.Remarks("未能输入正确的实例名称或未能首次调用**创建NLG实例**命令将导致错误")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("设置主题")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("设置动词")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("设置对象")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("添加补语")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("添加修饰符")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("添加预修改器")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("添加前置修改器")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("添加帖子修饰符")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_ParameterType { get; set; }

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请定义输入")]
        [Attributes.PropertyAttributes.InputSpecification("输入应与参数关联的值")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_Parameter { get; set; }

        public NLGSetParameterCommand_cn()
        {
            this.CommandName = "NLGSetParameterCommand";
            //  this.SelectionName = "Set NLG Parameter";
            this.SelectionName = Settings.Default.Set_NLG_Parameter_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
            this.v_InstanceName = "nlgDefaultInstance";
        }

        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
            var vInstance = v_InstanceName.ConvertToUserVariable(engine);
            var p = (SPhraseSpec)engine.GetAppInstance(vInstance);

            var userInput = v_Parameter.ConvertToUserVariable(sender);


            switch (v_ParameterType)
            {
                case "Set Subject":
                    p.setSubject(userInput);
                    break;
                case "Set Object":
                    p.setObject(userInput);
                    break;
                case "Set Verb":
                    p.setVerb(userInput);
                    break;
                case "Add Complement":
                    p.addComplement(userInput);
                    break;
                case "Add Modifier":
                    p.addModifier(userInput);             
                    break;
                case "Add Front Modifier":
                    p.addFrontModifier(userInput);
                    break;
                case "Add Post Modifier":
                    p.addPostModifier(userInput);
                    break;
                case "Add Pre-Modifier":
                    p.addPreModifier(userInput);
                    break;
                default:
                    break;
            }

            //remove existing associations if override app instances is not enabled
            engine.AppInstances.Remove(vInstance);

            //add to app instance to track
            engine.AddAppInstance(vInstance, p);


        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_ParameterType", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_Parameter", this, editor));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [" + v_ParameterType + ": '" + v_Parameter + "', Instance Name: '" + v_InstanceName + "']";
        }
    }
}