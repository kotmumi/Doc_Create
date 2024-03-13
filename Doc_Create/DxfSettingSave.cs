using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvAddIn
{
    class DxfSettingSave
    {
        TransientObjects _transientObjects;
        private static string oAsmNameDXF = "Чертежи\\Резка DXF";
        private static string oAsmNameDXF_PDF = "Чертежи\\Резка и Гибка DXF-PDF";
       
        public DxfSettingSave(PartDocument oDrawDoc, GettingPath gettingPath, string CustomName,string oFileName,bool onePart) 
        {
            string _oFolderDXF="";
            string _oFolderDXF_PDF="";
            try
            {
                SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDrawDoc.ComponentDefinition;
                if (onePart != true)
                {
                    _oFolderDXF = gettingPath.oAsmDocPath + oAsmNameDXF;
                    _oFolderDXF_PDF = gettingPath.oAsmDocPath + oAsmNameDXF_PDF;
                    CreatFolderSave creatFolderSaveDXF = new CreatFolderSave(_oFolderDXF);
                    CreatFolderSave creatFolderSaveDXF_PDF = new CreatFolderSave(_oFolderDXF_PDF);
                }
                else 
                {
                    _oFolderDXF = gettingPath.oPartDocPath;
                }

                string _oFolderDXForPDF = _oFolderDXF;
                if (onePart!=true)
                {
                    foreach (PartFeature oFeaturel in oCompDef.Features)
                    {
                        switch (oFeaturel.Type)
                        {
                            case Inventor.ObjectTypeEnum.kFlangeFeatureObject:
                                _oFolderDXForPDF = _oFolderDXF_PDF;
                                break;
                            default:
                                break;
                        }
                    }
                }
                if (oCompDef.HasFlatPattern == false)
                {
                    oCompDef.Unfold();
                }
                else
                {
                    oCompDef.FlatPattern.Edit();
                }
                _transientObjects = new TransientObjects(gettingPath.inventorApp);
                string sOut = "FLAT PATTERN DXF?AcadVersion=2010&OuterProfileLayer=IV_OUTER_PR​OFILE&InvisibleLayers=IV_TANGENT;IV_ARC_CENTERS;IV_BEND_DOWN;IV_BEND";
                _transientObjects.oDataMedium.FileName = _oFolderDXForPDF + "\\" + CustomName + oFileName + ".dxf";
                oCompDef.DataIO.WriteDataToFile(sOut, _transientObjects.oDataMedium.FileName);
                oCompDef.FlatPattern.ExitEdit();

                if (onePart != true)
                {
                    oDrawDoc.Close(true);
                }
            }
            catch (Exception e) { }
        }
    }
}
