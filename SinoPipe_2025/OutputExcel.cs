using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Windows.Forms;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace SinoPipe_2025
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class OutputExcel : IExternalEventHandler
    {
        //樣本檔路徑
        public string Filepath
        {
            get;
            set;
        }

        public List<double> xyz_shift
        {
            get;
            set;
        }

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ElementCategoryFilter instanceFilter = new ElementCategoryFilter(BuiltInCategory.OST_GenericModel);
            IList<Element> instanceList = collector.WherePasses(instanceFilter).WhereElementIsNotElementType().ToElements();

            //建立所有管線及管線相關元件
            IList<Element> pipeList = new List<Element>();
            IList<Element> manholeList = new List<Element>();
            IList<Element> side_ditch_list = new List<Element>();
            IList<Element> con_list = new List<Element>();
            string[] check_list = { "圓人孔", "圓手孔", "方人孔" , "方手孔" , "陰井",
                                   "消防栓", "號誌設備", "號誌開關", "號誌電桿", "變電箱",
                                   "路燈電桿", "電信箱", "電信電桿", "電力開關", "電力電桿", "電塔"};
            //所有元件依照側溝、管線、人孔分類
            for (int i = 0; i < instanceList.Count(); i++)
            {
                if (instanceList[i].Name[0] == 'e')
                {
                    if (instanceList[i].Name.Contains("SideDitch"))
                    {
                        side_ditch_list.Add(instanceList[i]);
                    }
                    else
                    {
                        pipeList.Add(instanceList[i]);
                    }

                }
                foreach (string a in check_list)
                {
                    if (instanceList[i].Name == a)
                    {
                        manholeList.Add(instanceList[i]);
                    }
                }
                if (instanceList[i].Name == "接頭")
                {
                    con_list.Add(instanceList[i]);
                }
            }

            //應用程序
            Excel.Application excelAPP = new Excel.Application();
            string filepath = Filepath;
            //檔案
            Excel.Workbook excelWorkbook = excelAPP.Workbooks.Open(filepath);
            //工作表
            Excel.Worksheet excelWorksheet1 = new Excel.Worksheet();
            excelWorksheet1 = excelWorkbook.Worksheets["0-1人孔建立"];
            Excel.Worksheet excelWorksheet2 = new Excel.Worksheet();
            excelWorksheet2 = excelWorkbook.Worksheets["0-2管線建立"];
            Excel.Worksheet excelWorksheet3 = new Excel.Worksheet();
            excelWorksheet3 = excelWorkbook.Worksheets["0-3側溝建立"];

            //宣告及初始化管線參數
            int countPipe = 2;
            double B = 0;//寬度
            double H = 0;//高度
            double D = 0;//直徑
            double t = 0;//壁厚
            string buried_depth;//埋管深度
            int Xn = 0;//列數X
            int Yn = 0;//列數Y
            double move_Z = 0;//偏移量

            //寫入管線參數
            foreach (Element ele in pipeList)
            {
                countPipe++;
                B = 0;//寬度
                H = 0;//高度
                D = 0;//直徑
                t = 0;//壁厚
                buried_depth = null;//埋管深度
                move_Z = ele.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble() * 304.8 / 1000;//管線偏移量

                //匯出管線資訊
                string pipe_tyoe = "";
                string specification = "";
                try { excelWorksheet2.Cells[countPipe, 1] = ele.LookupParameter("類別碼").AsString(); } catch { }
                try { excelWorksheet2.Cells[countPipe, 2] = ele.LookupParameter("管線編號").AsString(); } catch { }
                try { excelWorksheet2.Cells[countPipe, 3] = ele.LookupParameter("管線類型").AsString(); } catch { }
                try { excelWorksheet2.Cells[countPipe, 4] = ele.LookupParameter("管線材質").AsString(); } catch { }
                try { excelWorksheet2.Cells[countPipe, 5] = pipe_tyoe = ele.LookupParameter("管線型式").AsString(); } catch { }
                try { excelWorksheet2.Cells[countPipe, 13] = specification = ele.LookupParameter("管路規格").AsString(); } catch { }
                try { excelWorksheet2.Cells[countPipe, 14] = Xn = Int32.Parse(ele.LookupParameter("XY").AsString().Split(',')[0]); } catch { }
                try { excelWorksheet2.Cells[countPipe, 15] = Yn = Int32.Parse(ele.LookupParameter("XY").AsString().Split(',')[1]); } catch { }
                try { excelWorksheet2.Cells[countPipe, 16] = ele.LookupParameter("管線長度").AsString(); } catch { }
                try { excelWorksheet2.Cells[countPipe, 17] = buried_depth = ele.LookupParameter("埋管深度").AsString(); } catch { }
                try { excelWorksheet2.Cells[countPipe, 18] = ele.LookupParameter("管線總類代碼").AsString(); } catch { }
                try { excelWorksheet2.Cells[countPipe, 19] = ele.LookupParameter("附註").AsString(); } catch { }


                //進入元件內部取得元件Location
                try
                {
                    //讀取管線族群
                    FamilyInstance FI = ele as FamilyInstance;
                    FamilySymbol FS = FI.Symbol;

                    Document famDoc = doc.EditFamily(FS.Family);

                    //撈取族群內的元件
                    FilteredElementCollector famCollector = new FilteredElementCollector(famDoc);
                    IList<Element> famList = famCollector.WhereElementIsNotElementType().ToElements();

                    string location = "";
                    bool if_done = false;//若該元件已經讀取一次掃掠就不用執行第二次

                    foreach (Element e in famList)
                    {
                        //名字包含“掃掠”代表量體
                        if (e.Name.ToString() == "掃掠")
                        {
                            //進入掃掠內讀取curve array以得到XYZ
                            Element sweepElement = e;
                            Sweep sweepe = e as Sweep;

                            CurveArray CA = sweepe.Path3d.AllCurveLoops.get_Item(0);
                            int arraySize = CA.Size;
                            int Arrcount = 0;
                            if (if_done == false)
                            {
                                //double rvt_shiftZ = ele.LookupParameter("偏移").AsDouble();
                                double rvt_shiftZ = ele.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble();
                                foreach (Curve c in CA)
                                {
                                    Line line = c as Line;
                                    XYZ p = line.GetEndPoint(0);
                                    location += ((p.X * 304.8 / 1000) + xyz_shift[0]) + "," + ((p.Y * 304.8 / 1000) + xyz_shift[1]) + "," + (p.Z * 304.8 / 1000 + xyz_shift[2] + double.Parse(buried_depth) + rvt_shiftZ * 304.8 / 1000) + ";";
                                    if (Arrcount == arraySize - 1)
                                    {
                                        p = line.GetEndPoint(1);
                                        location += ((p.X * 304.8 / 1000) + xyz_shift[0]) + "," + ((p.Y * 304.8 / 1000) + xyz_shift[1]) + "," + (p.Z * 304.8 / 1000 + xyz_shift[2] + double.Parse(buried_depth) + rvt_shiftZ * 304.8 / 1000);
                                    }
                                    Arrcount++;
                                }
                                if_done = true;
                            }

                            //撈取量體族群，並依照族群行撈取不同參數
                            FamilySymbolProfile FSP = sweepe.ProfileSymbol as FamilySymbolProfile;
                            FamilySymbol P = FSP.Profile as FamilySymbol;

                            if (P.Name == "box")
                            {
                                double.TryParse(P.LookupParameter("B").AsValueString(), out B);
                                double.TryParse(P.LookupParameter("H").AsValueString(), out H);
                                t = P.LookupParameter("t").AsDouble() * 304.8;
                                D = 0;
                            }
                            else if (P.Name == "normal")
                            {
                                double.TryParse(P.LookupParameter("D").AsValueString(), out D);
                                t = P.LookupParameter("t").AsDouble() * 304.8;
                                B = 0;
                                H = 0;
                            }
                            else if (P.Name == "profile_base_origin_Edit")
                            {
                                double.TryParse(P.LookupParameter("內徑").AsValueString(), out D);
                                double.TryParse(P.LookupParameter("B").AsValueString(), out B);
                                double.TryParse(P.LookupParameter("H").AsValueString(), out H);
                                double.TryParse(P.LookupParameter("t").AsValueString(), out t);
                                D = D * 2;
                            }
                            else if (P.Name == "profile_core_origin_Edit")
                            {
                                double.TryParse(P.LookupParameter("內徑").AsValueString(), out D);
                                double.TryParse(P.LookupParameter("B").AsValueString(), out B);
                                double.TryParse(P.LookupParameter("H").AsValueString(), out H);
                                double.TryParse(P.LookupParameter("t").AsValueString(), out t);
                                D = D * 2;
                            }
                            else if (P.Name == "profile_circle_core_origin_Edit")
                            {
                                double.TryParse(P.LookupParameter("內徑").AsValueString(), out D);
                                double.TryParse(P.LookupParameter("壁厚").AsValueString(), out t);
                                D = D * 2;
                            }
                            else if (P.Name == "Box_Nx1")
                            {
                                Document fam_doc = doc.EditFamily(P.Family);
                                FilteredElementCollector fam_collector = new FilteredElementCollector(fam_doc);
                                ElementCategoryFilter dim_filter = new ElementCategoryFilter(BuiltInCategory.OST_Dimensions);
                                IList<Element> fam_list = fam_collector.WherePasses(dim_filter).WhereElementIsNotElementType().ToElements();
                                foreach (Element p_e in fam_list)
                                {
                                    //MessageBox.Show((p_e as Dimension).Name);
                                    if (p_e.Name.Contains("nB"))
                                    {
                                        B = double.Parse((p_e as Dimension).ValueString);
                                    }
                                    else if (p_e.Name.Contains("nH"))
                                    {
                                        H = double.Parse((p_e as Dimension).ValueString);
                                    }
                                    else if (p_e.Name.Contains("t"))
                                    {
                                        t = double.Parse((p_e as Dimension).ValueString);
                                    }
                                    else if (p_e.Name.Contains("D"))
                                    {
                                        D = double.Parse((p_e as Dimension).ValueString);
                                    }
                                }

                            }

                            excelWorksheet2.Cells[countPipe, 10] = B / 1000;
                            excelWorksheet2.Cells[countPipe, 11] = H / 1000;
                            excelWorksheet2.Cells[countPipe, 9] = D / 1000;
                            excelWorksheet2.Cells[countPipe, 12] = t / 1000;
                        }
                    }
                    //寫入管線座標
                    try { excelWorksheet2.Cells[countPipe, 8] = location; } catch { }
                }
                catch { TaskDialog.Show("!!", "fail"); excelWorkbook.Close(); break; }
            }

            //建立及初始化人孔參數
            int countManhole = 2;//列數
            double Bh = 0; //頸部深度
            double NH = 0; //箱體深度
            double NL = 0; //頸部長度
            double NW = 0; //頸部寬度
            H = 0;//深
            double GL = 0;//地表高程
            D = 0;//圓蓋直徑
            double L = 0;//方蓋長度
            double W = 0;//方蓋寬度
            double T = 0;//壁厚
            double BD = 0;//箱體直徑
            double BLL = 0;//箱左長
            double BLR = 0;//箱右長
            double BLW = 0;//箱下寬
            double BRW = 0;//箱上寬

            //寫入人孔參數
            foreach (Element ele in manholeList)
            {
                Bh = 0; NH = 0; NL = 0; NW = 0; H = 0; GL = 0; D = 0;
                L = 0; W = 0; T = 0; BD = 0; BLL = 0; BLR = 0; BLW = 0; BRW = 0;

                countManhole++;
                //匯出及寫入人孔參數
                try { excelWorksheet1.Cells[countManhole, 1] = ele.LookupParameter("類別碼").AsString(); } catch { }
                try { excelWorksheet1.Cells[countManhole, 2] = ele.LookupParameter("編號").AsString(); } catch { }
                try { excelWorksheet1.Cells[countManhole, 3] = ele.LookupParameter("類型").AsString(); } catch { }
                try { excelWorksheet1.Cells[countManhole, 4] = ele.LookupParameter("型式").AsString(); } catch { }
                try { excelWorksheet1.Cells[countManhole, 17] = ele.LookupParameter("頸部長度").AsDouble() * 304.8 / 1000; } catch { excelWorksheet1.Cells[countManhole, 17] = 0; }
                try { excelWorksheet1.Cells[countManhole, 18] = ele.LookupParameter("頸部長度").AsDouble() * 304.8 / 1000; } catch { excelWorksheet1.Cells[countManhole, 18] = 0; }
                try { D = ele.LookupParameter("圓蓋直徑").AsDouble() * 304.8; excelWorksheet1.Cells[countManhole, 13] = D / 1000; } catch { excelWorksheet1.Cells[countManhole, 13] = 0; }
                try { double rTemp = ele.LookupParameter("頸部半徑").AsDouble() * 304.8; excelWorksheet1.Cells[countManhole, 16] = 2 * rTemp / 1000; } catch { excelWorksheet1.Cells[countManhole, 16] = 0; }
                try { BD = ele.LookupParameter("箱體直徑").AsDouble() * 304.8; excelWorksheet1.Cells[countManhole, 21] = BD / 1000; } catch { excelWorksheet1.Cells[countManhole, 21] = 0; }
                try { T = ele.LookupParameter("壁厚").AsDouble() * 304.8; excelWorksheet1.Cells[countManhole, 20] = T / 1000; } catch { excelWorksheet1.Cells[countManhole, 20] = 0; }
                try { NH = ele.LookupParameter("頸部深度").AsDouble() * 304.8; excelWorksheet1.Cells[countManhole, 19] = NH / 1000; } catch { excelWorksheet1.Cells[countManhole, 19] = 0; }
                try { Bh = ele.LookupParameter("箱體深度").AsDouble() * 304.8; excelWorksheet1.Cells[countManhole, 26] = Bh / 1000; } catch { excelWorksheet1.Cells[countManhole, 26] = 0; }
                try { excelWorksheet1.Cells[countManhole, 9] = (ele.LookupParameter("方位角").AsDouble() / Math.PI) * 180; } catch { excelWorksheet1.Cells[countManhole, 9] = 0; }
                try { BLR = ele.LookupParameter("箱右長").AsDouble() * 304.8; excelWorksheet1.Cells[countManhole, 23] = BLR / 1000; } catch { excelWorksheet1.Cells[countManhole, 23] = 0; }
                try { BLL = ele.LookupParameter("箱左長").AsDouble() * 304.8; excelWorksheet1.Cells[countManhole, 22] = BLL / 1000; } catch { excelWorksheet1.Cells[countManhole, 22] = 0; }
                try { L = ele.LookupParameter("長度").AsDouble() * 304.8; excelWorksheet1.Cells[countManhole, 14] = L / 1000; } catch { excelWorksheet1.Cells[countManhole, 14] = 0; }
                try { W = ele.LookupParameter("寬度").AsDouble() * 304.8; excelWorksheet1.Cells[countManhole, 15] = W / 1000; } catch { excelWorksheet1.Cells[countManhole, 15] = 0; }
                try { BLW = ele.LookupParameter("箱下寬").AsDouble() * 304.8; excelWorksheet1.Cells[countManhole, 25] = BLW / 1000; } catch { excelWorksheet1.Cells[countManhole, 24] = 0; }
                try { BRW = ele.LookupParameter("箱上寬").AsDouble() * 304.8; excelWorksheet1.Cells[countManhole, 24] = BRW / 1000; } catch { excelWorksheet1.Cells[countManhole, 25] = 0; }
                try { H = Bh + NH; excelWorksheet1.Cells[countManhole, 12] = H / 1000; } catch { excelWorksheet1.Cells[countManhole, 12] = 0; }
                try { excelWorksheet1.Cells[countManhole, 5] = ele.LookupParameter("形狀").AsString(); } catch { TaskDialog.Show("message", "形狀參數為寫入"); }
                try { excelWorksheet1.Cells[countManhole, 28] = ele.LookupParameter("附註").AsString(); } catch { }

                //讀取人孔座標
                LocationPoint LPMH = ele.Location as LocationPoint;
                XYZ point = LPMH.Point;
                //double MH_Z = 0; //人孔Z座標
                excelWorksheet1.Cells[countManhole, 6] = (point.X * 304.8 / 1000) + xyz_shift[0];
                excelWorksheet1.Cells[countManhole, 7] = (point.Y * 304.8 / 1000) + xyz_shift[1];

                double z_correction = 0;
                string type = ele.LookupParameter("型式").AsString();

                if (type == "電力電桿") { z_correction = -9; }
                else if (type == "電塔") { z_correction = -25; }
                else if (type == "電信電桿") { z_correction = -5; }
                else if (type == "變電箱") { z_correction = -1.2; }
                else if (type == "電信箱") { z_correction = -1.4; }
                else if (type == "地上消防栓") { z_correction = -0.65; }
                else if (type == "路燈電桿") { z_correction = -5.3; }
                else if (type == "號誌電桿") { z_correction = -2.3; }
                else if (type == "地上設備" && ele.LookupParameter("類型").AsString() == "電信") { z_correction = -1.08; }
                else if (type == "路燈開關") { z_correction = -0.2; }
                else if (type == "號誌開關") { z_correction = -0.2; }

                try { excelWorksheet1.Cells[countManhole, 8] = (point.Z * 304.8 / 1000) + xyz_shift[2] + z_correction; } catch { }
                try { excelWorksheet1.Cells[countManhole, 10] = GL = (point.Z * 304.8 / 1000) + xyz_shift[2] + z_correction; } catch { }
                try { excelWorksheet1.Cells[countManhole, 11] = (GL - H / 1000); } catch { }
            }

            //宣告及初始化側溝參數
            int side_ditch_count = 2;
            double Bt = 0;//蓋厚
            T = 0;//壁厚
            double d = 0;//溝底半徑
            H = 0; //溝深
            W = 0;//溝寬
            foreach (Element side_ditch in side_ditch_list)
            {
                Bt = 0;//蓋厚
                T = 0;//壁厚
                d = 0;//溝底半徑
                H = 0; //溝深
                W = 0;//溝寬
                side_ditch_count++;
                //匯出及寫入側溝參數
                try { excelWorksheet3.Cells[side_ditch_count, 1] = side_ditch.LookupParameter("類別碼").AsString(); } catch { }
                try { excelWorksheet3.Cells[side_ditch_count, 3] = side_ditch.LookupParameter("代號").AsString(); } catch { }
                try { excelWorksheet3.Cells[side_ditch_count, 4] = side_ditch.LookupParameter("類型").AsString(); } catch { }

                //進入側溝元件族群內讀取資訊
                try
                {
                    FamilyInstance FI = side_ditch as FamilyInstance;
                    FamilySymbol FS = FI.Symbol;

                    Document famDoc = doc.EditFamily(FS.Family);

                    FilteredElementCollector famCollector = new FilteredElementCollector(famDoc);

                    IList<Element> famList = famCollector.WhereElementIsNotElementType().ToElements();
                    string location = "";
                    double length = 0;
                    foreach (Element e in famList)
                    {
                        if (e.Name.ToString() == "掃掠")
                        {
                            Element sweepElement = e;
                            Sweep sweepe = e as Sweep;

                            CurveArray CA = sweepe.Path3d.AllCurveLoops.get_Item(0);
                            int arraySize = CA.Size;
                            int Arrcount = 0;

                            FamilySymbolProfile FSP = sweepe.ProfileSymbol as FamilySymbolProfile;
                            FamilySymbol P = FSP.Profile as FamilySymbol;

                            //寫入及匯出側溝資訊
                            if (P.Name.Contains("蓋子"))
                            {
                                Bt = P.LookupParameter("蓋厚").AsDouble() * 304.8;
                            }
                            else if (P.Name.Contains("型管"))
                            {
                                T = P.LookupParameter("T_壁厚").AsDouble() * 304.8;
                                H = P.LookupParameter("H_深度").AsDouble() * 304.8;
                                W = P.LookupParameter("W_寬度").AsDouble() * 304.8;
                                if (side_ditch.LookupParameter("代號").AsString().Contains('U'))
                                {
                                    try { d = P.LookupParameter("d_半徑").AsDouble() * 304.8; } catch { d = 0; }
                                }

                                //讀取側溝XYZ
                                length = 0;
                                foreach (Curve c in CA)
                                {
                                    Line line = c as Line;
                                    XYZ p = line.GetEndPoint(0);
                                    location += ((p.X * 304.8 / 1000) + xyz_shift[0]) + "," + ((p.Y * 304.8 / 1000) + xyz_shift[1]) + "," + (p.Z * 304.8 / 1000 + xyz_shift[2] - T / 1000) + ";";
                                    if (Arrcount == arraySize - 1)
                                    {
                                        p = line.GetEndPoint(1);
                                        location += ((p.X * 304.8 / 1000) + xyz_shift[0] - T / 1000) + "," + ((p.Y * 304.8 / 1000) + xyz_shift[1]) + "," + (p.Z * 304.8 / 1000 + xyz_shift[2] - T / 1000);
                                    }
                                    length += c.Length;
                                    Arrcount++;
                                }
                            }


                        }

                    }

                    //寫入及匯出側溝資訊
                    try
                    {
                        excelWorksheet3.Cells[side_ditch_count, 2] = location;
                        excelWorksheet3.Cells[side_ditch_count, 9] = Bt / 1000;
                        excelWorksheet3.Cells[side_ditch_count, 6] = H / 1000;
                        excelWorksheet3.Cells[side_ditch_count, 5] = W / 1000;
                        excelWorksheet3.Cells[side_ditch_count, 7] = d / 1000;
                        excelWorksheet3.Cells[side_ditch_count, 8] = T / 1000;
                        excelWorksheet3.Cells[side_ditch_count, 10] = length * 304.8 / 1000;
                    }
                    catch { TaskDialog.Show("message", "location error!"); }
                }
                catch { TaskDialog.Show("!!", "fail"); excelWorkbook.Close(); break; }

            }

            //接頭參數
            foreach (Element ele in con_list)
            {
                string index = ele.LookupParameter("附註").AsString();

                for (int i = 0; i <= 1; i++)
                {
                    try { excelWorksheet2.Cells[int.Parse(index.Split('/')[i]) + 3, 20] = ele.LookupParameter("接頭編號").AsString().Split('/')[i]; } catch { }
                    try { excelWorksheet2.Cells[int.Parse(index.Split('/')[i]) + 3, 21] = ele.LookupParameter("接頭點位").AsString().Split('/')[i]; } catch { }
                    try { excelWorksheet2.Cells[int.Parse(index.Split('/')[i]) + 3, 22] = ele.LookupParameter("接頭管徑").AsString().Split('/')[i]; } catch { }
                    try { excelWorksheet2.Cells[int.Parse(index.Split('/')[i]) + 3, 23] = ele.LookupParameter("接頭連接").AsString().Split('/')[i]; } catch { }
                }
            }

            //save
            SaveFileDialog save_file_dialog = new SaveFileDialog();
            save_file_dialog.ShowDialog();
            excelWorkbook.SaveAs(save_file_dialog.FileName, Type.Missing, "", "", Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, 1, false, Type.Missing, Type.Missing, Type.Missing);

            //關閉即釋放物件
            excelWorksheet1 = null;
            excelWorksheet2 = null;
            excelWorksheet3 = null;
            excelWorkbook.Close();

            excelWorkbook = null;
            excelAPP.Quit();

            excelAPP = null;

            TaskDialog.Show("通知", "完成圖表匯出");
        }

        //尋找樣板檔內文字位置
        public Tuple<int, int> FindAddress(string name, Excel.Range xlRange)
        {
            Excel.Range address;
            address = xlRange.Find(name, MatchCase: true);
            var pos = Tuple.Create(address.Row, address.Column);
            return pos;
        }

        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
