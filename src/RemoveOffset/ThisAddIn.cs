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
            ribbon.RibbonCommandClicked += RibbonCommandClicked;
            return ribbon;
        }

        void RibbonCommandClicked(Ribbon ribbon, RibbonCommands cmd)
        {
            switch (cmd)
            {
                case RibbonCommands.RemoveOffset:
                    Remove();
                    break;
            }
        }

        private Range GetRange()
        {
            Range formulaSelection = null;
            try
            {

                var selection = xl.Selection as Range;

                if (selection != null)
                    formulaSelection = xl.Intersect(selection, selection.SpecialCells(XlCellType.xlCellTypeFormulas).Cells);

                if (formulaSelection != null)
                {
                    var foundOffset = (from Range cell in formulaSelection select cell.Formula).Any(f => ((string) f).IndexOf("OFFSET(") != -1);
                    if (!foundOffset)
                        formulaSelection = null;
                }

                if (formulaSelection == null || formulaSelection.Cells.Count == 0)
                    throw new Exception();

            }
            catch
            {
                MessageBox.Show(null, Resources.No_range_selected, Resources.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            }
            return formulaSelection;
        }

        private void Remove()
        {
            var selection = GetRange();
            if (selection == null)
                return;

            if (MessageBox.Show(null, Resources.There_is_no_undo, Resources.Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
                return;

            var errorCells = new List<string>();

            var calculation = xl.Calculation;
            xl.Calculation = XlCalculation.xlCalculationManual;
            try
            {
                foreach (Range cell in selection)
                {
                    var substitutionFunction = new Func<ParserResult, string>(r =>
                    {
                        try
                        {
                            var offset = r.Value;
                            offset = offset.Replace("COLUMN()", cell.Column.ToString());
                            offset = offset.Replace("ROW()", cell.Row.ToString());
                            var result = cell.Worksheet.Evaluate(offset) as Range;
                            if (result != null)
                            {
                                if (result.Parent.Name == cell.Parent.Name)
                                    return result.get_Address(External: false, RowAbsolute: false, ColumnAbsolute: false);
                                return string.Format("'{0}'!{1}", result.Parent.Name, result.get_Address(External: false, RowAbsolute: false, ColumnAbsolute: false));
                            }
                        }
                        catch
                        {
                            errorCells.Add(cell.Address);
                        }
                        return r.Value;
                    });

                    var parser = new FunctionParser("OFFSET", cell.Formula, substitutionFunction);
                    if (parser.Parse() == ParserState.Success)
                        cell.Formula = parser.Output;
                }

                if (errorCells.Count > 0)
                    MessageBox.Show(null, string.Format(Resources.The_following_cells_were_skipped, String.Join(", ", errorCells)), Resources.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
