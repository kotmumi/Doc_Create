using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Windows.Forms;

namespace InvAddIn
{
    class QuantitySearch
    {
        public BOMView bomView;
        BOM bom;
        ComponentDefinition oCompDef;
        Document a;
        PartDocument partDoc;
        Property b;
        Property PartNum;

        public QuantitySearch(BOM bomAssembly,PartDocument partDocument)
        {
            partDoc = partDocument;
            bom = bomAssembly;
            bomView = bom.BOMViews["Только детали"];
            PartNum = partDoc.PropertySets["Design Tracking Properties"]["Part Number"];
        }

        public int BomQuantity(BOMRowsEnumerator rowsEnum)
        {
            foreach (BOMRow _oRow in rowsEnum)
            {
                oCompDef = _oRow.ComponentDefinitions[1];
                a = ((Document)oCompDef.Document);
                b = a.PropertySets["Design Tracking Properties"]["Part Number"];
               // if (b.Value.ToString() != PartNum.Value.ToString() & _oRow.ChildRows != null)
                //{
               //     return _oRow.ItemQuantity * BomQuantity(_oRow.ChildRows);
                //}
                //else
                if (b.Value.ToString() == PartNum.Value.ToString())
                {
                    return _oRow.ItemQuantity;
                }
                else { continue; }
            }
            return -1;
            }
        }
    }
















// foreach (BOMRow _oRowChild in _oRow.ChildRows)
//{
//  oCompDefChild = _oRowChild.ComponentDefinitions[1];
//aChild = ((Document)oCompDefChild.Document);
//bChild = aChild.PropertySets["Design Tracking Properties"]["Part Number"];
//if (bChild.Value.ToString() == PartNum.Value.ToString())
//{
//  MessageBox.Show(_oRow.TotalQuantity.ToString() + " = " + (_oRow.ItemQuantity * _oRowChild.ItemQuantity).ToString());
// return _oRow.ItemQuantity * _oRowChild.ItemQuantity;
//}
//}