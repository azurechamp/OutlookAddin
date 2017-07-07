using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace InsideSales
{
    public partial class ThisAddIn
    {
        #region vars
        public Outlook.Application OutlookApplication;
        #endregion

        public object OutlookMailItem { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            Outlook.MailItem mail = Item as Outlook.MailItem;
            var flawcount = 0;
            
            if (mail != null)
            {
                foreach (Outlook.Recipient reci in mail.Recipients)
                {

                    if (reci.Address.ToLower().EndsWith("ems.com")|| reci.Address.ToLower().EndsWith("bootbarn.com") || reci.Address.ToLower().EndsWith("skiphop.com") || reci.Address.ToLower().EndsWith("lacrossefootwear.com"))
                    {
                        reci.Delete();
                        flawcount++;
                    }
                }
            }

            if (flawcount > 0)
            {
                MessageBox.Show("Your email contains Recipients which are in NOT ALLOWED list \n NOT ALLOWED RECIPIENTS are removed. \n Please Make sure !" , "Warning!! Recipient!!!");
                Cancel = true;
            }
           
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
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
