using Inventor;
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
            TransientObjects _transientObjects;
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
                    switch (sheetMetal.Material.Name)
                    {
                        case "Нержавеющая сталь":
                        case "Сталь нержавеющая, 440C":
                        case "Сталь, нержавеющая AISI 440C, сварочная":
                        case "Сталь, нержавеющая, аустенитная":
                            _oMaterial = "AISI";
                            break;
                        case "Алюминий 6061":
                        case "Алюминий 6061-АНС":
                        case "Алюминий 6061, сварочный":
                            _oMaterial = "AL";
                            break;
                        case "Пластик АБС":
                        case "Пластик ЖКП":
                        case "Пластик ПАЭК":
                        case "Пластик ПБТ":
                        case "Пластик ПК/АБС":
                        case "Пластик ПММА":
                        case "Пластик ПФС":
                        case "Пластик ПЭТ":
                        case "Пластик САН":
                        case "Поликарбонат, прозрачный":
                        case "Полипропилен":
                        case "Полистирол":
                            _oMaterial = "Пластик";
                            break;
                        default:
                            _oMaterial = "...";
                            break;
                    }
                    double thickness = (double)sheetMetal.Thickness.Value * 10;
                    CustomName = thickness.ToString() + "мм-" + _oMaterial + "-" + "...шт-";
                }
                catch (Exception ex)
                {
                    CustomName = "??мм-??шт";
                }
                SheetMetalComponentDefinition oCompDef = (SheetMetalComponentDefinition)oDrawDoc.ComponentDefinition;

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
                _transientObjects.oDataMedium.FileName = _oFolderDXF + "\\" + CustomName + oFileName + ".dxf";
                oCompDef.DataIO.WriteDataToFile(sOut, _transientObjects.oDataMedium.FileName);
                oCompDef.FlatPattern.ExitEdit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка закрытия документа"+ex.Message);
            }
            Process.Start("explorer.exe", _oFolderDXF);
        }
    }
}
        

