using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InvAddIn
{
    internal class GettingPath
    {
        public Inventor.Application inventorApp = null;
        public AssemblyDocument oAsmDoc =  null;
        public PartDocument oPartDoc = null;
        public DocumentTypeEnum oAsmDocType;
        public string oAsmDocPath;
        public string oPartDocPath;
        public GettingPath()
        {
            try
            {
                // Attempt to get a reference to a running instance of Inventor.
                inventorApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
                oAsmDocType = inventorApp.ActiveDocumentType;
                switch (oAsmDocType)
                {
                    case DocumentTypeEnum.kAssemblyDocumentObject:
                        oAsmDoc = (AssemblyDocument)inventorApp.ActiveDocument;
                        oAsmDocPath = oAsmDoc.FullFileName.Replace(oAsmDoc.DisplayName, "");
                        break;
                    case DocumentTypeEnum.kPartDocumentObject:
                        oPartDoc = (PartDocument)inventorApp.ActiveDocument;
                        oPartDocPath = oPartDoc.FullFileName.Replace(oPartDoc.DisplayName, "");
                        break;
                    default:
                        MessageBox.Show("Запустите правило, находясь в сборке.", "iLogic");
                        return;
                }
            }
            catch(Exception e)
            {
                MessageBox.Show("Unable to connect to Inventor."+e.Message);
                return;
            }
        }
    }
}
