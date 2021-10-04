using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace auto_line
{

    class App : IExternalApplication
    {
        static string addinAssmeblyPath = Assembly.GetExecutingAssembly().Location;
        public Result OnStartup(UIControlledApplication a)
        {
            
            try { a.CreateRibbonTab("中興自動化"); } catch { }
            RibbonPanel ribbonPanel = a.CreateRibbonPanel("中興自動化", "SinoPipe API");

            PushButton pushbutton1 = ribbonPanel.AddItem(
                new PushButtonData("SinoPipe", "Pipeline",
                    addinAssmeblyPath, "auto_line.Start"))
                        as PushButton;
            pushbutton1.ToolTip = "SinoPipe";
            pushbutton1.LargeImage = convertFromBitmap(Properties.Resources.藍灰2);

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
