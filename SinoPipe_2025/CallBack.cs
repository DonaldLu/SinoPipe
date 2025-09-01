using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;


namespace SinoPipe_2025
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class Callback : IExternalEventHandler
    {
        public string callValues
        {
            get;
            set;
        }
        public void Execute(UIApplication app)
        {
            Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Document doc = uidoc.Document;

            try
            {
            //用逗號分開形成list
            ICollection<string> temp = callValues.Split(',').ToList();

            //去掉空格
            foreach (string a in temp)
            {
                a.Trim();
            }

            //利用ID找到元件
            ICollection<ElementId> id_list = new List<ElementId>();
            foreach (string id_str in temp)
            {
                long id_int = Convert.ToInt64(id_str);
                ElementId id = new ElementId(id_int);
                id_list.Add(id);
            }
            View view = doc.ActiveView;
            Transaction t = new Transaction(doc);

            t.Start("隔離元素");

            //隔離元素
            view.IsolateElementsTemporary(id_list);
            //畫面 zoom 到元素位置
            uidoc.ShowElements(id_list);

            t.Commit();

            TaskDialog.Show("Done", "Done");
            }
            catch (Exception e)
            { TaskDialog.Show("Error", e.Message + e.StackTrace); }
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
