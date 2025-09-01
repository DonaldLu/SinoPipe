using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.IO;
using ExReaderConsole;
using Autodesk.Revit.DB.Structure;
using System.Runtime.InteropServices;

namespace SinoPipe
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class SweepSideDitch : IExternalEventHandler
    {
        public string file_path
        {
            get;
            set;
        }
        public IList<double> xyz_shift
        {
            get;
            set;
        }
        public void Execute(UIApplication app)
        {
            Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Document doc = uidoc.Document;
            UIDocument edit_uidoc = app.OpenAndActivateDocument(@"D:\SinoPipe-RFA\sideditch.rfa");
            Document edit_doc = edit_uidoc.Document;

            Application app_ = doc.Application;
            UIApplication uiapp = new UIApplication(app_);


            var filename = doc.PathName;
            var rvtname = edit_doc.PathName;


            //excel
            ExReader rain = new ExReader();
            ExReader dan = new ExReader();
            rain.SetData(file_path, 3);
            rain.PassSSD();//為了跳過前面兩個
            rain.CloseEx();
            dan.SetData(file_path, 1);
            dan.PassMH();
            dan.CloseEx();
            //處理座標
            List<List<XYZ>> pos_list = new List<List<XYZ>>();
            //陰井座標
            List<List<XYZ>> Start_InGing_list = new List<List<XYZ>>();
            List<List<XYZ>> End_InGing_list = new List<List<XYZ>>();
            //存XYZ string
            List<string> pos_data = new List<string>();
            //存陰井 XY string
            List<string> Start_InGing_data = new List<string>();
            List<string> End_InGing_data = new List<string>();
            //存陰井 Z string
            List<string> z_string = new List<string>();
            foreach (List<string> rows in rain.MHdata)
            {
                pos_data.Add(rows[1]);
                Start_InGing_data.Add(rows[3]);
                End_InGing_data.Add(rows[6]);
            }
            //處理string

            foreach (string pos_string in pos_data)
            {
                string[] pos_row = pos_string.Replace(";", ",").Split(',');
                string tx = "", ty = "", tz = "";
                List<XYZ> pos_Row = new List<XYZ>();

                //刪除重複
                for (int i = 0; i != pos_row.Length; i = i + 3)
                {
                    string x = pos_row[i], y = pos_row[i + 1], z = pos_row[i + 2];
                    if (x == tx && y == ty && z == tz || x == "")
                        continue;
                    z_string.Add(z);
                    XYZ pos = new XYZ(Double.Parse(x), Double.Parse(y), Double.Parse(z));
                    pos_Row.Add(pos);

                    tx = x; ty = y; tz = z;
                }
                pos_list.Add(pos_Row);

            }


            //計算座標平移量(利用人孔座標)
            /*double[] sumpos = new double[3];

            foreach (List<string> rows in dan.MHdata)
            {

                for (int j = 0; j != 3; j++)
                {
                    sumpos[j] += double.Parse(rows[j + 5]);
                }
            }

            int mh_count = dan.MHdata.Count;

            //座標平移
            double xshift = (int)(sumpos[0] / mh_count);
            double yshift = (int)(sumpos[1] / mh_count);
            double zshift = 0;*/
            double xshift = xyz_shift[0];
            double yshift = xyz_shift[1];
            double zshift = xyz_shift[2];


            int index = 0;
            //設置警告、處理錯誤訊息
            WarningSwallower warningSwallower = new WarningSwallower();
            warningSwallower.Case_name = file_path;
            warningSwallower.Pipe_number_list = new List<string>();

            //開始建置掃略
            try
            {
                foreach (List<string> rows in rain.MHdata)
                {
                    int count = pos_list[index].Count;
                    IList<XYZ> start_point = new List<XYZ>();
                    IList<XYZ> end_point = new List<XYZ>();
                    using (Transaction t = new Transaction(edit_doc, "Create sphere direct shape"))
                    {
                        t.Start();

                        //刪除前一個做過的

                        ICollection<Sweep> sweeps = new FilteredElementCollector(edit_doc).OfClass(typeof(Sweep)).Cast<Sweep>().ToList();

                        foreach (Sweep sweep in sweeps)
                        {
                            edit_doc.Delete(sweep.Id);
                        }

                        //開始掃掠
                        try
                        {
                            
                            //U型管或方形管
                            if (rows[2].Contains("U"))
                            {
                                double side_z = zshift;
                                string profilename = "溝槽U型管";
                                if (double.Parse(rows[8]) != 0)
                                {
                                    side_z = zshift - (double.Parse(rows[8]));
                                }
                                //掃略U型管
                                GoSweep(count, pos_list, xshift, yshift, side_z, edit_doc, index, profilename);


                                //修改屬性
                                ICollection<FamilySymbol> familySymbol_list = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                                FamilySymbol profile_test = familySymbol_list.Where(x => x.Name == profilename).ToList().First();
                                profile_test.LookupParameter("W_寬度").SetValueString((double.Parse(rows[4]) * 1000).ToString());
                                profile_test.LookupParameter("H_深度").SetValueString((double.Parse(rows[5]) * 1000).ToString());
                                profile_test.LookupParameter("d_半徑").SetValueString((double.Parse(rows[6]) * 1000).ToString());
                                profile_test.LookupParameter("T_壁厚").SetValueString((double.Parse(rows[7]) * 1000).ToString());

                                //暗溝的話則設定蓋子屬性
                                if (double.Parse(rows[8]) != 0)
                                {
                                    string GAZname = "溝槽蓋子";
                                    GoSweep(count, pos_list, xshift, yshift, side_z, edit_doc, index, GAZname);

                                    FamilySymbol Gaze = familySymbol_list.Where(x => x.Name == GAZname).ToList().First();
                                    Gaze.LookupParameter("蓋厚").SetValueString((double.Parse(rows[8]) * 1000).ToString());
                                    Gaze.LookupParameter("蓋寬").SetValueString(((double.Parse(rows[4]) + (2 * double.Parse(rows[7]))) * 1000).ToString());
                                    Gaze.LookupParameter("蓋寬*0.5").SetValueString(((double.Parse(rows[4]) + (2 * double.Parse(rows[7]))) * 1000 / 2).ToString());

                                }
                            }
                            else if (rows[2].Contains("L"))
                            {
                                double side_z = zshift;
                                string profilename = "溝槽方型管";
                                if (double.Parse(rows[8]) != 0)
                                {
                                    side_z = zshift - (double.Parse(rows[8]));
                                }
                                //掃略方型管
                                GoSweep(count, pos_list, xshift, yshift, side_z, edit_doc, index, profilename);


                                //修改屬性
                                ICollection<FamilySymbol> familySymbol_list = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                                FamilySymbol profile_test = familySymbol_list.Where(x => x.Name == profilename).ToList().First();
                                profile_test.LookupParameter("W_寬度").SetValueString((double.Parse(rows[4]) * 1000).ToString());
                                profile_test.LookupParameter("H_深度").SetValueString((double.Parse(rows[5]) * 1000).ToString());
                                profile_test.LookupParameter("T_壁厚").SetValueString((double.Parse(rows[7]) * 1000).ToString());

                                //暗溝的話則設定蓋子屬性
                                if (double.Parse(rows[8]) != 0)
                                {
                                    string GAZname = "溝槽蓋子";
                                    GoSweep(count, pos_list, xshift, yshift, side_z, edit_doc, index, GAZname);

                                    FamilySymbol Gaze = familySymbol_list.Where(x => x.Name == GAZname).ToList().First();
                                    Gaze.LookupParameter("蓋厚").SetValueString((double.Parse(rows[8]) * 1000).ToString());
                                    Gaze.LookupParameter("蓋寬").SetValueString(((double.Parse(rows[4]) + (2 * double.Parse(rows[7]))) * 1000).ToString());
                                    Gaze.LookupParameter("蓋寬*0.5").SetValueString(((double.Parse(rows[4]) + (2 * double.Parse(rows[7]))) * 1000 / 2).ToString());

                                }

                            }
                            else { }

                            FailureHandlingOptions failOpt = t.GetFailureHandlingOptions();

                            warningSwallower.Pipe_number = rows[2];
                            failOpt.SetFailuresPreprocessor(warningSwallower);
                            t.SetFailureHandlingOptions(failOpt);
                        }


                        catch (Exception e)
                        {
                            TaskDialog.Show("error", e.Message);
                        }



                        t.Commit();

                        //判斷是否成功
                        sweeps = new FilteredElementCollector(edit_doc).OfClass(typeof(Sweep)).Cast<Sweep>().ToList();
                        if (sweeps.Count == 0)
                        {
                            warningSwallower.Pipe_number_list.Add(warningSwallower.Pipe_number);
                        }

                        //另存成一個rfa
                        SaveAsOptions saveAsOptions = new SaveAsOptions { OverwriteExistingFile = true, MaximumBackups = 1 };

                        edit_doc.SaveAs(@"D:\SinoPipe-RFA\realcase\" + "edit_TU_SideDitch" + index + ".rfa", saveAsOptions);

                    }
                    using (Transaction t = new Transaction(doc, "load family"))
                    {
                        t.Start();

                        //載入剛剛的rfa
                        doc.LoadFamily(@"D:\SinoPipe-RFA\realcase\" + "edit_TU_SideDitch" + index + ".rfa");
                        ICollection<FamilySymbol> familySymbol_list = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                        foreach (FamilySymbol fam in familySymbol_list)
                        {
                            //找到剛剛的rfa檔並建立元件
                            if (fam.Name == "edit_TU_SideDitch" + index)
                            {
                                fam.Activate();
                                FamilyInstance instance = doc.Create.NewFamilyInstance(XYZ.Zero, fam, StructuralType.NonStructural);
                                instance.LookupParameter("代號").Set(rows[2]);
                                instance.LookupParameter("類別碼").Set(rows[0]);
                                instance.LookupParameter("類型").Set(rows[3]);
                            }
                        }


                        t.Commit();

                    }
                    index++;
                }
            }
            catch (Exception e) { TaskDialog.Show("error", e.Message); }


            var rebirthdoc = app.OpenAndActivateDocument(filename);
            edit_doc.Close(false);
            TaskDialog.Show("Done", "建置完畢");
            warningSwallower.Set_failue(warningSwallower.Pipe_number_list);




        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
        //訊息處理
        public class WarningSwallower : IFailuresPreprocessor
        {

            public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
            {
                IList<FailureMessageAccessor> failureMessageAccessors = a.GetFailureMessages();
                foreach (FailureMessageAccessor fma in failureMessageAccessors)
                {
                    if (fma.GetDescriptionText() == "無法建立掃掠")
                    {
                        a.ResolveFailure(fma);
                        return FailureProcessingResult.ProceedWithCommit;
                    }

                }

                a.DeleteAllWarnings();
                return FailureProcessingResult.Continue;
            }
            public string Pipe_number
            {
                get;
                set;
            }
            public IList<string> Pipe_number_list
            {
                get;
                set;
            }
            public string Case_name
            {
                get;
                set;
            }
            public void Set_failue(IList<string> number_list)
            {
                //excel
                ExReader mhx = new ExReader();
                mhx.SetData(Case_name, 2);
                try
                {
                    if (number_list.Count != 0)
                    {
                        foreach (string number in number_list)
                        {
                            mhx.Change_color(mhx.FindAddress(number));
                        }
                    }
                    mhx.Save_excel(Case_name.Replace(Case_name.Split('\\').Last(), "") + "Result_" + Case_name.Split('\\').Last());
                    mhx.CloseEx();
                }
                catch { mhx.CloseEx(); };


            }
        }
        public SketchPlane Sketch_plain(Document doc, XYZ start, XYZ end)
        {
            SketchPlane sk = null;

            XYZ v = end - start;

            double dxy = Math.Abs(v.X) + Math.Abs(v.Y);

            XYZ w = (dxy > 0.00000001)
              ? XYZ.BasisZ
              : XYZ.BasisY;

            XYZ norm = v.CrossProduct(w).Normalize();

            Plane geomPlane = Plane.CreateByNormalAndOrigin(norm, start);

            sk = SketchPlane.Create(doc, geomPlane);

            return sk;
        }
        public void GoSweep(int count, List<List<XYZ>> pos_list, double xshift, double yshift, double zshift, Document doc, int index, string profilename)
        {
            ReferenceArray reff = new ReferenceArray();
            IList<ElementId> model_curveids = new List<ElementId>();
            for (int j = 0; j != count - 1; j++)
            {
                XYZ start = (pos_list[index][j] - new XYZ(xshift, yshift, zshift)) * 1000 / 304.8;
                XYZ end = (pos_list[index][j + 1] - new XYZ(xshift, yshift, zshift)) * 1000 / 304.8;
                //Sweep需要1.布林子 2.ReferenceArray(需要用ModelCurve(1.Curve=>line.GetBound)(2.Sketch_plain(學長寫好的code))) 
                //3.SweepProfile(先抓所有的FamilySymbol，然後用Where去找他) 4.0 5.選擇從哪裡開始
                //2.
                Curve curve = Line.CreateBound(start, end);
                ModelCurve modelCurve = doc.FamilyCreate.NewModelCurve(curve, Sketch_plain(doc, start, end));
                reff.Append(modelCurve.GeometryCurve.Reference);
                model_curveids.Add(modelCurve.Id);


            }
            //3.
            ICollection<FamilySymbol> familySymbol_list = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            FamilySymbol profile_test = familySymbol_list.Where(x => x.Name == profilename).ToList().First();
            SweepProfile sweepProfile = doc.Application.Create.NewFamilySymbolProfile(profile_test);
            //掃掠
            Sweep sweep = doc.FamilyCreate.NewSweep(true, reff, sweepProfile, 0, ProfilePlaneLocation.Start);
            sweep.get_Parameter(BuiltInParameter.PROFILE_ANGLE).SetValueString("-90");

            doc.Delete(model_curveids);
        }
    }
}
