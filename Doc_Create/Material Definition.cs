using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

namespace InvAddIn
{
    class Material_Definition
    {
        public SheetMetalComponentDefinition _sheetMetal;

        public Material_Definition(SheetMetalComponentDefinition sheetMetal)
        {
            _sheetMetal = sheetMetal;
        }
        public string Material_Definition_Part()
        {
            switch (_sheetMetal.Material.Name)
            {
                case "Нержавеющая сталь":
                case "Сталь нержавеющая, 440C":
                case "Сталь, нержавеющая AISI 440C, сварочная":
                case "Сталь, нержавеющая, аустенитная":
                    return "AISI";

                case "Алюминий 6061":
                case "Алюминий 6061-АНС":
                case "Алюминий 6061, сварочный":
                    return "AL";

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
                    return "Пластик";

                default:
                    return "...";
            }
        }
    }
}
