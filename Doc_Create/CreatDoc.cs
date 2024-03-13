using System;
using Inventor;
using System.Windows.Forms;
using System.Diagnostics;

namespace InvAddIn
{
    class CreatDoc
    {
        public CreatDoc()
        {
        }
        public void Rule()
        {
            GettingPath gettingPath = new GettingPath();
            TransientObjects _transientObjects;
            PartDocument oDrawDoc;
            ModelParameters oParam;
            BOM oBOM = null;
            DocumentsEnumerator oRefDocs = gettingPath.oAsmDoc.AllReferencedDocuments;
            string CustomName;
            string oAsmNameDXF = "Чертежи\\Резка DXF";
            string oAsmNameDXF_PDF = "Чертежи\\Резка и Гибка DXF-PDF";
            string oFolderDrowing = gettingPath.oAsmDocPath+"Чертежи";
            string _oFolderDXF = gettingPath.oAsmDocPath  + oAsmNameDXF;
            string _oFolderDXF_PDF = gettingPath.oAsmDocPath + oAsmNameDXF_PDF;
            string _oMaterial = "...";
            try
            {
                CreatFolderSave creatFolderSaveDXF = new CreatFolderSave(_oFolderDXF);
                CreatFolderSave creatFolderSaveDXF_PDF = new CreatFolderSave(_oFolderDXF_PDF);
                oBOM = gettingPath.oAsmDoc.ComponentDefinition.BOM;
                oBOM.StructuredViewEnabled = true;
                oBOM.PartsOnlyViewEnabled = true;
                oBOM.StructuredViewFirstLevelOnly = false;
            }
            
            catch (Exception e) { MessageBox.Show("Ошибка BOM 2"); }
            
            foreach (Document oRefDoc in oRefDocs)
            {
   
                string iptPathName = oRefDoc.FullDocumentName;
                iptPathName = iptPathName.Substring(0, oRefDoc.FullDocumentName.Length - 3) + "ipt";
                if (System.IO.File.Exists(iptPathName))
                {
                    string oFileName = oRefDoc.DisplayName;
                    oFileName = oFileName.Substring(0, oRefDoc.DisplayName.Length - 4);
                  
                    oDrawDoc = (PartDocument)gettingPath.inventorApp.Documents.Open(iptPathName, true);
                    
                    try {
                        QuantitySearch quantitySearch = new QuantitySearch(oBOM, oDrawDoc);
                        try
                            {
                            oParam = oDrawDoc.ComponentDefinition.Parameters.ModelParameters;
                            SheetMetalComponentDefinition sheetMetal = (SheetMetalComponentDefinition)oDrawDoc.ComponentDefinition;
                            Material_Definition material_definition = new Material_Definition(sheetMetal);
                            _oMaterial = material_definition.Material_Definition_Part();
                            double thickness = (double)sheetMetal.Thickness.Value * 10;
                            CustomName = thickness.ToString() + "мм-"+ _oMaterial+"-" + quantitySearch.BomQuantity(quantitySearch.bomView.BOMRows) + "шт-" ;
                        }
                        catch (Exception ex)
                        {
                            CustomName = "??мм-??шт";
                        }
                        SheetMetalComponentDefinition oCompDef =(SheetMetalComponentDefinition) oDrawDoc.ComponentDefinition;
                        string _oFolderDXForPDF= _oFolderDXF;
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
                    } catch (Exception ex) 
                    {
                      MessageBox.Show("Ошибка закрытия документа"+ex.Message);
                    }
                    oDrawDoc.Close(true);
                }
            }
            Process.Start("explorer.exe", oFolderDrowing);
        }
    }
}
