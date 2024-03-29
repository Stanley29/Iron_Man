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
    [Attributes.ClassAttributes.Group("(Excel Commands)Excel命令")]
    [Attributes.ClassAttributes.Description("此命令允许您删除Excel中的特定行")]
    [Attributes.ClassAttributes.UsesDescription("如果要从当前工作表中删除整行，请使用此命令。")]
    [Attributes.ClassAttributes.ImplementationDescription("此命令实现Excel Interop以实现自动化。")]
    public class ExcelDeleteRowCommand_cn : ScriptCommand
    {
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("请输入实例名称")]
        [Attributes.PropertyAttributes.InputSpecification("输入** Create Excel **命令中指定的唯一实例名称")]
        [Attributes.PropertyAttributes.SampleUsage("**myInstance** or **seleniumInstance**")]
        [Attributes.PropertyAttributes.Remarks("未能输入正确的实例名称或未能首先调用**创建Excel **命令将导致错误")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        public string v_InstanceName { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowVariableHelper)]
        [Attributes.PropertyAttributes.PropertyDescription("指示要删除的行号")]
        [Attributes.PropertyAttributes.InputSpecification("输入应删除的行号。")]
        [Attributes.PropertyAttributes.SampleUsage("1, 5, [vNumber]")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_RowNumber { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("删除后，下面的细胞是否应该向上移动？")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("Yes")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("No")]
        [Attributes.PropertyAttributes.InputSpecification("指示下面的行是否会向上移动以替换旧行。")]
        [Attributes.PropertyAttributes.SampleUsage(" 选择“是”或“否” ")]
        [Attributes.PropertyAttributes.Remarks("")]
        public string v_ShiftUp { get; set; }
        public ExcelDeleteRowCommand_cn()
        {
            this.CommandName = "ExcelDeleteRowCommand";
            //this.SelectionName = "Delete Row";
            this.SelectionName = Settings.Default.Delete_Row_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;
        }
        public override void RunCommand(object sender)
        {
            var engine = (Core.Automation.Engine.AutomationEngineInstance)sender;
            var vInstance = v_InstanceName.ConvertToUserVariable(engine);
            var excelObject = engine.GetAppInstance(vInstance);

          


                Microsoft.Office.Interop.Excel.Application excelInstance = (Microsoft.Office.Interop.Excel.Application)excelObject;
                Microsoft.Office.Interop.Excel.Worksheet workSheet = excelInstance.ActiveSheet;

                string rowToDelete = v_RowNumber.ConvertToUserVariable(sender);

                var cells = workSheet.Range["A" + rowToDelete, Type.Missing];
                var entireRow = cells.EntireRow;
                if (v_ShiftUp == "Yes")
                {
                    entireRow.Delete();
                }
                else
                {
                    entireRow.Clear();
                }


            
        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);

            //create standard group controls
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_RowNumber", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_ShiftUp", this, editor));

            return RenderedControls;

        }
        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Row Number: " + v_RowNumber + ", Instance Name: '" + v_InstanceName + "']";
        }
    }
}