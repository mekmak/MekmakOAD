using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace Mekmak.OAD
{
    [ComVisible(true)]
    public class RibbonExtension : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public string GetCustomUI(string ribbonId)
        {
            string ribbonXml = string.Empty;
            switch (ribbonId)
            {
                case "Microsoft.Outlook.Explorer":
                    ribbonXml = GetResourceText("Mekmak.OAD.MailItemContextMenu.xml");
                    break;
            }

            return ribbonXml;
        }

        #region Ribbon Callbacks       

        public void OnDownloadAllAttachmentsContextMenuMailItemClick(Office.IRibbonControl control)
        {
            List<MailItem> emails = GetMailItems(control);
            if (emails == null)
            {
                return;
            }

            string directory = Path.Combine(Path.GetTempPath(), "MekmakOAD", DateTime.Now.Ticks.ToString());
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            foreach (MailItem mailItem in emails)
            {
                DownloadAttachments(mailItem, directory);
            }

            Process.Start(directory);
        }

        private void DownloadAttachments(MailItem mailItem, string directory)
        {
            if(mailItem == null)
            {
                return;
            }

            int attachmentCount = mailItem.Attachments.Count;
            if (attachmentCount == 0)
            {
                return;
            }

            for(int index = 1 /* COM arrays are 1-indexed ... */; index <= attachmentCount; index++)
            {
                Attachment attachment = mailItem.Attachments[index];
                string fileName = Path.Combine(directory, attachment.FileName);
                attachment.SaveAsFile(fileName);
            }
        }

        private List<MailItem> GetMailItems(Office.IRibbonControl control)
        {
            Selection selection = control.Context as Selection;
            if (selection == null)
            {
                return null;
            }

            int count = selection.Count;
            if (count <= 0)
            {
                return null;
            }

            List<MailItem> emails = selection.OfType<MailItem>().ToList();
            return emails;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

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
