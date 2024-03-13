using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace InvAddIn
{
    internal class CreatFolderSave
    {
        public CreatFolderSave(string oFolder) 
        {
            if (System.IO.Directory.Exists(oFolder)==false)
            {
                System.IO.Directory.CreateDirectory(oFolder);
            };
        }
    }
}
