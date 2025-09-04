using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using System.Windows.Media.Imaging;

namespace SinoPipe_2020
{
    class App : IExternalApplication
    {
        static string addinAssmeblyPath = Assembly.GetExecutingAssembly().Location;
        public Result OnStartup(UIControlledApplication a)
        {

            RibbonPanel ribbonPanel = null;
            try { a.CreateRibbonTab("中興工程自動化建模"); } catch { }
            try { ribbonPanel = a.CreateRibbonPanel("中興工程自動化建模", "中興工程自動化建模"); }
            catch
            {
                List<RibbonPanel> panel_list = new List<RibbonPanel>();
                panel_list = a.GetRibbonPanels("中興工程自動化建模");
                foreach (RibbonPanel rp in panel_list)
                {
                    if (rp.Name == "中興工程自動化建模")
                    {
                        ribbonPanel = rp;
                    }
                }
            }

            PushButton pushbutton1 = ribbonPanel.AddItem(
                new PushButtonData("sinopipe", "sinopipe",
                    addinAssmeblyPath, "SinoPipe_2020.Start"))
                        as PushButton;
            pushbutton1.ToolTip = "SinoPipe";
            pushbutton1.LargeImage = convertFromBitmap(Properties.Resources.Red);

            return Result.Succeeded;
        }

        BitmapSource convertFromBitmap(System.Drawing.Bitmap bitmap)
        {
            return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                bitmap.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
