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

namespace SinoPipe_2025
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class 測試元件 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Autodesk.Revit.DB.Document document = commandData.Application.ActiveUIDocument.Document;

            Transaction t = new Transaction(document);
            t.Start("測試元件");
            FamilySymbol familySymbol = new FilteredElementCollector(document).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "變電箱").First();
            familySymbol.Activate();
            FamilyInstance a = document.Create.NewFamilyInstance(XYZ.Zero, familySymbol, StructuralType.NonStructural);
            Color color = new Color(255, 127, 0); // RGB
            OverrideGraphicSettings overrideGraphicSettings = new OverrideGraphicSettings();
            //overrideGraphicSettings.SetProjectionFillColor(color);
            overrideGraphicSettings.SetProjectionLineColor(color);
            document.ActiveView.SetElementOverrides(a.Id, overrideGraphicSettings);
            t.Commit();
            return Result.Succeeded;
        }

    }
}
