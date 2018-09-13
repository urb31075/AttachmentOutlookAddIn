// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ThisAddIn.cs" company="urb31075">
//  All Right Reserved 
// </copyright>
// <summary>
//   Defines the ThisAddIn type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------
namespace AttachmentOutlookAddIn
{
    using Office = Microsoft.Office.Core;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// The this add in.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>
        /// Gets the this add in.
        /// </summary>
        public static ThisAddIn thisAddIn { get; private set; }

        /// <summary>
        /// Gets the this application.
        /// </summary>
        public static Outlook.Application thisApplication { get; private set; }
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            thisAddIn = this;
            thisApplication = this.Application;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
