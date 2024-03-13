using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;

namespace InvAddIn
{
    public class TransientObjects
    {
       public DataMedium oDataMedium;
        public TransientObjects(Inventor.Application inventorApp)
        {
           oDataMedium = inventorApp.TransientObjects.CreateDataMedium();
           TranslationContext oContext = inventorApp.TransientObjects.CreateTranslationContext();
            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;
            NameValueMap oOptions = inventorApp.TransientObjects.CreateNameValueMap();
        }
    }
}
