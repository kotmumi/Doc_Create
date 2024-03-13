﻿using Inventor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InvAddIn
{
    class CreatDXFOne
    {

        public void Rule1()
        {
            GettingPath gettingPath = new GettingPath();
            //TransientObjects _transientObjects;
            PartDocument oDrawDoc = gettingPath.oPartDoc;
            ModelParameters oParam;
            string CustomName;
            string _oFolderDXF = gettingPath.oPartDocPath;
            string _oMaterial = "...";
            string oFileName = oDrawDoc.DisplayName;
            oFileName = oFileName.Substring(0, oDrawDoc.DisplayName.Length - 4);
            try
            {
                try
                {
                    oParam = oDrawDoc.ComponentDefinition.Parameters.ModelParameters;
                    SheetMetalComponentDefinition sheetMetal = (SheetMetalComponentDefinition)oDrawDoc.ComponentDefinition;
                    Material_Definition material_definition = new Material_Definition(sheetMetal);
                    double thickness = (double)sheetMetal.Thickness.Value * 10;
                    CustomName = thickness.ToString() + "мм-" + _oMaterial + "-" + "...шт-";
                }
                catch (Exception ex)
                {
                    CustomName = "??мм-??шт";
                }
                DxfSettingSave dxfSettingSave = new DxfSettingSave(oDrawDoc, gettingPath, CustomName, oFileName, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка закрытия документа"+ex.Message);
            }
            Process.Start("explorer.exe", _oFolderDXF);
        }
    }
}
        

