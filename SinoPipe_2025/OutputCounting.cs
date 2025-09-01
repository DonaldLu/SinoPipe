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
using System.Text.RegularExpressions;

namespace SinoPipe_2025
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class OutputCounting : IExternalEventHandler
    {

        public string file_path
        {
            get;
            set;
        }
        public string path
        {
            get;
            set;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
                UIDocument uidoc = new UIDocument(document);
                Document doc = uidoc.Document;

                ICollection<FamilyInstance> FamIn_list = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList();
                List<FamilyInstance> edit_list = new List<FamilyInstance>();

                foreach (FamilyInstance Fmin in FamIn_list)
                {
                    //找管線而不是側溝
                    if (Fmin.Name.Contains("edit") == true && Fmin.Name.Contains("TU") == false)
                    {
                        edit_list.Add(Fmin);
                    }
                }
                List<string> pipe_type = new List<string>();

                //找規格
                foreach (FamilyInstance f in edit_list)
                {
                    string type = f.LookupParameter("管線總類代碼").AsString();

                    if (pipe_type.Contains(type) == false)
                    {
                        pipe_type.Add(type);
                    }
                    pipe_type = pipe_type.Distinct().ToList();
                }

                //找代碼
                Dictionary<string, List<string>> pipe_dictionary = new Dictionary<string, List<string>>();
                foreach (string type in pipe_type)
                {
                    List<string> list = new List<string>();

                    foreach (FamilyInstance f in edit_list.Where(x => x.LookupParameter("管線總類代碼").AsString() == type).ToList())
                    {
                        string str = f.LookupParameter("管路規格").AsString();
                        list.Add(str);
                        list = list.Distinct().ToList();

                    }
                    pipe_dictionary.Add(type, list);
                }

                List<string> n_length = new List<string>();
                List<string> n_name = new List<string>();
                List<double> n_total_length = new List<double>();
                List<string> n_section = new List<string>();

                foreach (string key in pipe_dictionary.Keys)
                {
                    foreach (string type in pipe_dictionary[key])
                    {
                        double total_l = new double();
                        string l = null;
                        string id = null;
                        foreach (FamilyInstance fami in edit_list.Where(x => x.LookupParameter("管線總類代碼").AsString() == key && x.LookupParameter("管路規格").AsString() == type).ToList())
                        {
                            //累加管線長度
                            total_l += double.Parse(fami.LookupParameter("管線長度").AsString());
                            //分別顯示長度
                            l += "+" + fami.LookupParameter("管線長度").AsString().ToString();
                            //分別顯示管線ID
                            id += ", " + fami.Id.ToString();
                        }
                        l = l.Remove(0, 1);
                        id = id.Remove(0, 2);
                        n_length.Add(l);
                        n_name.Add(id);
                        //形成數量計算書需求之格式
                        if (type.Split('x').Count() == 2)
                        {
                            n_section.Add(key + "ψ" + type.Split('x').First() + "mmX" + type.Split('x').Last());
                        }
                        else
                        {
                            n_section.Add(key + "ψ" + type);
                        }
                        n_total_length.Add(total_l);
                    }
                }

                //資料讀取完畢
                //開始寫進excel
                Excel.Application Eapp = new Excel.Application();
                Excel.Workbook EWb = Eapp.Workbooks.Open(file_path);

                Excel.Worksheet EWs1 = EWb.Worksheets[1];
                Excel.Range ERa1_whole = EWs1.UsedRange;

                Excel.Worksheet EWs3 = EWb.Worksheets[3];
                Excel.Range ERa3_whole = EWs3.Columns[1];

                //管線部份
                foreach (string key in pipe_type)
                {
                    string first_key = key.Substring(0, 1);
                    if (first_key == "H")
                    {
                        continue;
                    }
                    string second_key = "";
                    try
                    {
                        second_key = key.Substring(1, 1);
                    }
                    catch { }

                    Excel.Range currentFind = null;
                    Excel.Range firstFind = null;
                    currentFind = ERa1_whole.Find(key, Type.Missing,Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,Type.Missing, Type.Missing);

                    while (currentFind != null)
                    {
                        // Keep track of the first range you find. 
                        if (firstFind == null)
                        {
                            firstFind = currentFind;
                        }
                        // If you didn't move to a new range, you are done.
                        else if (currentFind.Row <= firstFind.Row)
                        {
                            break;
                        }

                        int cou = currentFind.Row;
                        Excel.Range aa = EWs1.Cells[cou, 1];
                        aa = aa.EntireRow;
                        string replay = EWs1.Cells[aa.Row, 2].Text;
                        string reclass = EWs1.Cells[aa.Row, 6].Text;
                        //TaskDialog.Show("1", EWs1.Cells[aa.Row, 1].Text);
                        for (int i = (n_section.Count - 1); i >= 0; i--)
                        {
                            if (replay.Contains("管線規格") && reclass.Contains(first_key) && reclass.Contains(second_key))
                            {
                                if (n_section[i].Contains(key))
                                {
                                    //determine the detail spec (radius) of the pipe
                                    if (reclass.Contains("mm"))
                                    {
                                        int index = reclass.IndexOf("mm");

                                        int correction = 0;
                                        //consider when the radius of the pipe < 100
                                        if (Char.IsDigit(reclass[index - 3]) == false)
                                        {
                                            correction = 1;
                                        }
                                        Int32.TryParse(reclass.Substring(index - 3 + correction, 3 - correction), out int standard);
                                        correction = 0;
                                        if(Char.IsDigit(n_section[i][5]) == false)
                                        {
                                            correction = 1;
                                        }
                                        Int32.TryParse(n_section[i].Substring(3, 3 - correction), out int specimen);

                                        if (reclass[index + 6].ToString() == "上")
                                        {
                                            if (standard > specimen)
                                            {
                                                continue;
                                            }
                                        }
                                        else if (reclass[index + 6].ToString() == "下")
                                        {
                                            if (standard < specimen)
                                            {
                                                continue;
                                            }
                                        }
                                    }

                                    aa.Insert(Excel.XlInsertShiftDirection.xlShiftDown, aa.Copy(Type.Missing));
                                    EWs1.Cells[aa.Row, 2] = replay.Replace("管線規格", n_section[i]);
                                    EWs1.Cells[aa.Row, 4] = n_total_length[i];
                                    EWs1.Cells[aa.Row, 5] = n_length[i];
                                    EWs1.Cells[aa.Row, 6] = n_name[i];
                                }
                            }
                        }
                        currentFind = ERa1_whole.FindNext(currentFind);
                    }
                }

                //人手孔部分
                string[] search_list = {"人孔", "手孔", "電塔", "電力電桿", "電信電桿", "路燈電桿", "路燈開關", "號誌電桿", "號誌開關", "地上設備", "閥類", "開關", "地上消防栓", "取水器", "陰井", "閘門", "共同管線地上設備"};

                foreach (string target in search_list)
                {
                    List<FamilyInstance> instance_list = new List<FamilyInstance>();
                    List<string> class_list = new List<string>();
                    List<FamilyInstance> temp_list = new List<FamilyInstance>();

                    string id = null;
                    string temp_id = null;

                    string category_code = null;
                    Excel.Range category = null;

                    //category_name = 類別碼,  category_code_name = 代碼
                    string category_name = null;
                    string category_code_name = null;

                    //create class_list
                    foreach (FamilyInstance Fmin in FamIn_list)
                    {
                        try
                        {
                            if (Fmin.ParametersMap.get_Item("型式").AsString() == target)
                            {
                                instance_list.Add(Fmin);
                                id += ", " + Fmin.Id.ToString();
                                category_code = Fmin.ParametersMap.get_Item("類別碼").AsString();
                                category = ERa3_whole.Find(category_code, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                                category_code_name = category.Offset[0, 1].Text;

                                //"first letter of category_code is 'S'" or "category_code which only has one letter" don't need to remove second letter
                                if (category_code_name[0] == 'S' || category_code_name.Length == 1)
                                {
                                    category_name = category.Offset[0, 1].Text;
                                }
                                else
                                {
                                    category_name = category.Offset[0, 1].Text.Remove(1, 1);
                                }

                                //avoid 不明管線 import to the form
                                if (category_name != "A")
                                {
                                    class_list.Add(category_name);
                                    //TaskDialog.Show("1", category_name);
                                }
                            }
                        }
                        catch { }
                    }
                    class_list = class_list.Distinct().ToList();

                    //if model doesn't have that target, skip following program 
                    //how : class_list will be empty >> class_list.Any() will be false >> !class_list.Any() will be true
                    if (!class_list.Any())
                    {
                        continue;
                    }

                    //
                    //TaskDialog.Show("start", "開始寫入" + target );
                    Excel.Range current = null;
                    Excel.Range first = null;
                    Excel.Range last = null;

                    //find current target
                    current = ERa1_whole.Find(target, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                    while (current != null)
                    {
                        // Keep track of the first range you find. 
                        if (first == null)
                        {
                            first = current;
                        }
                        // If you didn't move to a new range, you are done.
                        else if (current.get_Address(Excel.XlReferenceStyle.xlA1) == first.get_Address(Excel.XlReferenceStyle.xlA1))
                        {
                            current.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                            break;
                        }
                        else if (current.Offset[1, 0].get_Address(Excel.XlReferenceStyle.xlA1) == first.get_Address(Excel.XlReferenceStyle.xlA1))
                        {
                            break;
                        }

                        int cou = current.Row;
                        Excel.Range aa = EWs1.Cells[cou, 4];
                        aa = aa.EntireRow;
                        string replay = EWs1.Cells[aa.Row, 6].Text;

                        foreach (string cat in class_list)
                        {
                            if (replay.Contains(cat))  
                            {
                                foreach (FamilyInstance Fmin in instance_list)
                                {
                                    category_code = Fmin.ParametersMap.get_Item("類別碼").AsString();
                                    category = ERa3_whole.Find(category_code, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                                    category_code_name = category.Offset[0, 1].Text;

                                    //"first letter of category_code is 'S'" or "category_code which only has one letter" don't need to remove second letter
                                    if (category_code_name[0] == 'S' || category_code_name.Length == 1)
                                    {
                                        category_name = category.Offset[0, 1].Text;
                                    }
                                    else
                                    {
                                        category_name = category.Offset[0, 1].Text.Remove(1, 1);
                                    }

                                    try
                                    {
                                        if (category_name == cat)
                                        {
                                            temp_id += ", " + Fmin.Id.ToString();
                                            temp_list.Add(Fmin);
                                        }
                                    }
                                    catch { }
                                }
                                
                                aa.Insert(Excel.XlInsertShiftDirection.xlShiftDown, aa.Copy(Type.Missing));
                                EWs1.Cells[aa.Row, 4] = temp_list.Count.ToString();
                                EWs1.Cells[aa.Row, 6] = temp_id.Remove(0, 2);
                                temp_list.Clear();
                                temp_id = null;
                            }
                        }
                        last = current;
                        current = ERa1_whole.FindNext(current);
                    }
                }


                EWb.SaveAs(path, Type.Missing, "", "", Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, 1, false, Type.Missing, Type.Missing, Type.Missing);
                EWs1 = null;
                EWs3 = null;
                EWb.Close();
                EWb = null;
                Eapp.Quit();
                Eapp = null;

                TaskDialog.Show("Finish", "結束數量計算。");
            }
            catch
            (Exception e)
            { TaskDialog.Show("Error", e.Message + e.StackTrace); }
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
