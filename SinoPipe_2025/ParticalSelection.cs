using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;

namespace SinoPipe_2025
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class ParticalSelection : IExternalEventHandler
    {
        public void Execute(UIApplication app)
        {
            Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Document doc = uidoc.Document;
            Selection sel = uidoc.Selection;

            //pick objects from Revit
            IList<Reference> pick = sel.PickObjects(ObjectType.Element, "請選擇多個元件，程式會自動偵測執行");
            IList<Element> sel_ele = new List<Element>();
            if (pick.Count != 0)
            {
                foreach (Reference obj in pick)
                {
                    Element ele = doc.GetElement(obj);
                    sel_ele.Add(ele);
                }
            }

            //抓取資料
            int ele_count = sel_ele.Count();
            List<string> length = new List<string>();
            List<double> total_length = new List<double>();
            List<string> section = new List<string>();
            foreach (Element edit in sel_ele)
            {
                length.Add(edit.LookupParameter("管線長度").AsString().ToString());
                total_length.Add(double.Parse(edit.LookupParameter("管線長度").AsString()));
                section.Add(edit.LookupParameter("管線總類代碼").AsString().ToString() + "ψ" + edit.LookupParameter("管路規格").AsString().Split('x').First().ToString() + "mmX" + edit.LookupParameter("管路規格").AsString().Split('x').Last().ToString());
            }


            string end = null;
            try
            {
                for (int i = 0; i <= section.Count(); i++)
                {
                    end += section[i] + "  " + total_length[i] + "  " + length[i] + "\n";
                }
            }
            catch { }
            //產生監測結果
            TaskDialog.Show("test", "您一共選擇了" + sel_ele.Count().ToString() + "個元件，其管線規格如下:\n" + end);

        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
