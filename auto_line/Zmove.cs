using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;
using System.Diagnostics;
using System.Windows.Forms;

namespace auto_line
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class Zmove : IExternalEventHandler
    {
        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document; Selection sel = uidoc.Selection;

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
            Transaction t = new Transaction(doc);
            t.Start("修正埋管深度");
            foreach(FamilyInstance pipe in sel_ele)
            {
                if (pipe.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsValueString() != "0" && pipe.Name.Contains("edit"))
                {
                    double shift_z = pipe.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble()*304.8/1000;

                    double new_value = double.Parse(pipe.LookupParameter("埋管深度").AsString()) - shift_z;
                    pipe.LookupParameter("埋管深度").Set(new_value.ToString());
                }
            }
            TaskDialog.Show("修正資訊", "管線埋管深度已修正。");
            t.Commit();
        }


        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
