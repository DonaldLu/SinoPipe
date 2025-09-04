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

namespace SinoPipe_2020
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    public class SweepDraw : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Autodesk.Revit.DB.Document document = commandData.Application.ActiveUIDocument.Document;

            Family detailArc = new FilteredElementCollector(document).OfClass(typeof(Family)).Cast<Family>().ToList().Where(x => x.Name == "Box_Nx1").First();

            Autodesk.Revit.DB.Document fa_doc = document.EditFamily(detailArc);

            List<string> para_s = new List<string> {"Xn", "nB", "nH", "t" };

            List<string> parValue = new List<string> { 4.ToString(), 1250.ToString(), 1200.ToString(), 200.ToString()};// x, y, dia, real num

            ProduceProfile( ref fa_doc, ref para_s, ref parValue);

            DrawBox(ref fa_doc, ref parValue);

            fa_doc.LoadFamily(document, new FamilyOption());

            fa_doc.Close(false);

            return Result.Succeeded;
        }
        public void ProduceProfile(ref Autodesk.Revit.DB.Document doc, ref List<string> vs, ref List<string> pv)
        {
            using (Transaction t = new Transaction(doc, "setting profile parameters"))
            {
                t.Start();

                FamilyManager familyManager = doc.FamilyManager;

                for (int i = 0; i != vs.Count; i++)
                {
                    FamilyParameter familyParameter = familyManager.get_Parameter(vs[i]);

                    try
                    {
                        familyManager.SetValueString(familyParameter, pv[i]);
                    }
                    catch
                    {
                        familyManager.Set(familyParameter, int.Parse(pv[i]));
                    }
                }
                t.Commit();
            }
        }
        public void DrawBox(ref Autodesk.Revit.DB.Document doc, ref List<string> xy)
        {
            using (Transaction t = new Transaction(doc, "make profile xn * yn"))
            {

                t.Start();

                IList<CurveElement> detailArc = new FilteredElementCollector(doc).OfClass(typeof(CurveElement)).Cast<CurveElement>().ToList().Where(x => x.GeometryCurve.IsBound == true).ToList();
                
                ViewPlan viewplan = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>().ToList().Where(x => x.Name == "參考樓層").First();

                IList<CurveElement> Ndetails = detailArc;

                double td = double.Parse(xy[3])+double.Parse(xy[1]);

                for(int j = 4; j < 8; j++)
                {

                    LocationCurve lc = Ndetails[j].Location as LocationCurve;
                    for (int i = 0; i < int.Parse(xy[0]); i++)
                    {
                        if (i == 0)
                        {
                        }
                        else
                        {

                            Transform trans2 = Transform.CreateTranslation(new XYZ(td * (i) / 304.8, 0, 0));

                            Curve curve_y = lc.Curve.CreateTransformed(trans2);

                            CurveElement c = doc.FamilyCreate.NewDetailCurve(viewplan, curve_y) as CurveElement;

                        }

                    }
                }

                t.Commit();
            }
        }
        public void DrawCurve_Inner(ref Autodesk.Revit.DB.Document doc, ref List<string> xy)
        {
            using (Transaction t = new Transaction(doc, "make inner profile xn * yn"))
            {

                t.Start();

                CurveElement detailArc = new FilteredElementCollector(doc).OfClass(typeof(CurveElement)).Cast<CurveElement>().ToList().Where(x => x.GeometryCurve.IsBound == false).ToList()[1];

                ViewPlan viewplan = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>().ToList().Where(x => x.Name == "參考樓層").First();


                CurveElement Ndetail = detailArc;

                IList<CurveElement> curveElements = new List<CurveElement>();

                curveElements.Add(Ndetail);

                LocationCurve lc = Ndetail.Location as LocationCurve;

                double PB2 = double.Parse(xy[2]) * 2 * 2;

                for (int i = 0; i < int.Parse(xy[1]); i++)
                {

                    for (int j = 0; j < int.Parse(xy[0]); j++)
                    {
                        if (i == 0 && j == 0)
                        {
                        }
                        else
                        {

                            Transform trans2 = Transform.CreateTranslation(new XYZ(PB2 * (j) / 304.8, PB2 * (-i) / 304.8, 0));

                            Curve curve_y = lc.Curve.CreateTransformed(trans2);

                            CurveElement c = doc.FamilyCreate.NewDetailCurve(viewplan, curve_y) as CurveElement;

                            curveElements.Add(c);

                        }
                    }

                }

                int n = int.Parse(xy[0]) * int.Parse(xy[1]) - int.Parse(xy[3]);


                for (int i = 0; i < n; i++)
                {
                    doc.Delete(curveElements[i].Id);
                }

                t.Commit();
            }
        }
        class FamilyOption : IFamilyLoadOptions
        {

            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse,out FamilySource source,out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                source = FamilySource.Family;
                return true;
            }
        }
    }

}
