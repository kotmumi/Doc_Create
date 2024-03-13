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
            PartDocument oDrawDoc;
            ModelParameters oParam;
            BOM oBOM = null;
            DocumentsEnumerator oRefDocs = gettingPath.oAsmDoc.AllReferencedDocuments;
            string CustomName;
            string oFolderDrowing = gettingPath.oAsmDocPath+"Чертежи";
            string _oMaterial = "...";
            try
            {
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
                        DxfSettingSave dxfSettingSave = new DxfSettingSave(oDrawDoc,gettingPath,CustomName,oFileName,false);
                        
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
