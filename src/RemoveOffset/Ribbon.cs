using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Core;
using Office = Microsoft.Office.Core;

namespace RemoveOffset
{
    public enum RibbonCommands
    {
        RemoveOffset
    }

    public delegate void RibbonCommandHandler(Ribbon ribbon, RibbonCommands cmd);

    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        public event RibbonCommandHandler RibbonCommandClicked;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("RemoveOffset.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void RemoveOffsetButton_Click(IRibbonControl control)
        {
            OnRibbonCommandClicked(RibbonCommands.RemoveOffset);
        }

        #endregion

        #region Helpers

        protected void OnRibbonCommandClicked(RibbonCommands cmd)
        {
            if (RibbonCommandClicked != null)
                RibbonCommandClicked(this, cmd);
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
