using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;
using RemoveOffset.Properties;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace RemoveOffset
{
    public partial class ThisAddIn
    {
        private Excel.Application xl;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            xl = Application;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            var ribbon = new Ribbon();
            ribbon.RibbonCommandClicked += ribbon_RibbonCommandClicked;
            return ribbon;
        }

        void ribbon_RibbonCommandClicked(Ribbon ribbon, RibbonCommands cmd)
        {
            switch (cmd)
            {
                case RibbonCommands.RemoveOffset:
                    Remove();
                    break;
            }
        }

        private void Remove()
        {
            var selection = xl.Selection as Excel.Range;
            if (selection == null)
            {
                MessageBox.Show(null, Resources.No_range_selected, Resources.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show(null, Resources.There_is_no_undo, Resources.Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
                return;

            var errorCells = new List<string>();

            var calculation = xl.Calculation;
            xl.Calculation = XlCalculation.xlCalculationManual;
            try
            {


                foreach (Excel.Range cell in selection)
                {

                    var parser = new Parser(cell.Formula);
                    if (parser.Parse(r =>
                                         {
                                             try
                                             {
                                                 var offset = r.Value;
                                                 offset = offset.Replace("COLUMN()", cell.Column.ToString());
                                                 offset = offset.Replace("ROW()", cell.Row.ToString());
                                                 var result = cell.Worksheet.Evaluate(offset) as Excel.Range;
                                                 if (result != null)
                                                 {
                                                     if (result.Parent.Name == cell.Parent.Name)
                                                         return result.get_Address(External: false, RowAbsolute: false,
                                                                                   ColumnAbsolute: false);
                                                     return string.Format("'{0}'!{1}", result.Parent.Name,
                                                                          result.get_Address(External: false,
                                                                                             RowAbsolute: false,
                                                                                             ColumnAbsolute: false));
                                                 }
                                             }
                                             catch
                                             {
                                                 errorCells.Add(cell.Address);
                                             }
                                             return r.Value;


                                         }))
                        cell.Formula = parser.Output;
                }

                if (errorCells.Count > 0)
                    MessageBox.Show(null,
                                    string.Format(Resources.The_following_cells_were_skipped,
                                                  String.Join(", ", errorCells)), Resources.Title, MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
            }
            finally
            {
                xl.Calculation = calculation;
            }

        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
