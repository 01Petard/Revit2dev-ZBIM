using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;



namespace UiLable
{
    [Transaction(TransactionMode.Manual)]
    class UIDemo : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            //【1】第一步：创建一个RibbonTab
            application.CreateRibbonTab("UITab");//new tab
            //【2】第二步:在刚才的RibbonTab中创建UIPanel
            RibbonPanel rp = application.CreateRibbonPanel("UITab", "UIPanel");
            //【3】第三步:指定程序集的名称以及所使用的类名
          //  string assemblyPath = @"E:\b站课程\面向工程人员的revit二次开发课堂\08Revit二开之UIRibbon\RibbonUIDemo\UIDemo\UIButtonDemo\bin\Debug\UIButtonDemo.dll";
            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            //当前程序集路径，当前程序路径

            //string revitExeLoaction = Process.GetCurrentProcess().MainModule.FileName;//可获得当前执行的exe的文件名。
            //TaskDialog.Show("当前程序路径", $"revitExeLoaction是{revitExeLoaction}");


            string classNameFilterWallDemo = "SolidPractise.FilterWallDemo";
            //【4】第四步：创建PushButton
            PushButtonData pbd = new PushButtonData("InnerNameRevit", "计算墙体体积", assemblyPath, classNameFilterWallDemo);
            //【4-1】讲pushButton添加到面板中
          
            PushButton pushButton = rp.AddItem(pbd) as PushButton;

            //【4-2】给按钮设置一个图片（大图标一般是32px，小图标一般是16px，格式可以是ico,png,jpg）
            // string imgPath = @"E:\b站课程\面向工程人员的revit二次开发课堂\08Revit二开之UIRibbon\圣诞节_棒棒糖.png";
            // pushButton.LargeImage = new BitmapImage(new Uri(imgPath));

            pushButton.LargeImage = new BitmapImage(new Uri("pack://application:,,,/UiLable;component/pic/建材.png", UriKind.Absolute));
            //【4-3】给按钮设置一个默认提示信息
            pushButton.ToolTip = "FilterWallDemo";


            return Result.Succeeded;

        }
    }
}
