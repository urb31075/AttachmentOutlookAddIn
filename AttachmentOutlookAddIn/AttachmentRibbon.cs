using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Application = System.Windows.Forms.Application;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;
using Exception = System.Exception;

namespace AttachmentOutlookAddIn
{
    using System.Windows.Forms;

    public partial class AttachmentRibbon
    {
        /// <summary>
        /// The this ribbon.
        /// </summary>
        private IRibbonUI thisRibbon;
        
        /// <summary>
        /// The email status pane.
        /// </summary>
        private CustomTaskPane monitorPane;        
        
        
        private void TestRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }


        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var currentCursor = Cursor.Current;
                Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();
                if (this.monitorPane == null)
                {
                    this.monitorPane = ThisAddIn.thisAddIn.CustomTaskPanes.Add(new AttachmentUserControl(), "Вложения");
                    this.monitorPane.VisibleChanged += (s, ea) =>
                    {
                        if (this.thisRibbon != null)
                        {
                            this.thisRibbon.Invalidate();
                        }
                    };

                    this.monitorPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                    this.monitorPane.Width = 370;
                }

                this.monitorPane.Visible = !this.monitorPane.Visible; // Visiblethis.MonitorToggleButton.Checked;
                Cursor.Current = currentCursor;
            }
            catch (Exception ex)
            {
                var message = string.Format("Ошибка запуска: {0}", ex.Message);
                MessageBox.Show(message, @"Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
