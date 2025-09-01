using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.Creation;
using ExReaderConsole;

namespace SinoPipe
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    public class Start : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Form2 form2 = new Form2();
            form2.Show();
            form2.Visible = false;//隱藏form2
            Form1 form1 = new Form1(commandData.Application.ActiveUIDocument, form2);
            form1.Show();

            return Result.Succeeded;
        }
    }
}
