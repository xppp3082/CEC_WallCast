#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection; // for getting the assembly path
using System.Windows.Media; // for the graphics 需引用prsentationCore
using System.Windows.Media.Imaging;

#endregion

namespace CEC_WallCast
{
    class App : IExternalApplication
    {
        //測試將其他button加到現有TAB
        const string RIBBON_TAB = "【CEC MEP】";
        const string RIBBON_PANEL = "穿牆開口";
        const string RIBBON_PANEL2 = "穿牆CSD&SEM";
        public Result OnStartup(UIControlledApplication a)
        {

            RibbonPanel targetPanel = null;
            // get the ribbon tab
            try
            {
                a.CreateRibbonTab(RIBBON_TAB);
            }
            catch (Exception) { } //tab alreadt exists
            RibbonPanel panel = null;
            //創建「穿樑套管」頁籤
            List<RibbonPanel> panels = a.GetRibbonPanels(RIBBON_TAB); //在此要確保RIBBON_TAB在這行之前已經被創建
            foreach (RibbonPanel pnl in panels)
            {
                if (pnl.Name == RIBBON_PANEL)
                {
                    panel = pnl;
                    break;
                }
            }
            // couldn't find panel, create it
            if (panel == null)
            {
                panel = a.CreateRibbonPanel(RIBBON_TAB, RIBBON_PANEL);
            }

            //創建「SEM&CSD」頁籤
            RibbonPanel panel2 = null;
            foreach (RibbonPanel pnl in panels)
            {
                if (pnl.Name == RIBBON_PANEL2)
                {
                    panel2 = pnl;
                    break;
                }
            }
            // couldn't find panel, create it
            if (panel2 == null)
            {
                panel2 = a.CreateRibbonPanel(RIBBON_TAB, RIBBON_PANEL2);
            }

            // get the image for the button
            System.Drawing.Image image_CreateST = Properties.Resources.穿牆套管ICON合集_更新_svg;
            ImageSource imgSrc0 = GetImageSource(image_CreateST);

            System.Drawing.Image image_Create = Properties.Resources.穿牆套管ICON合集_放置_svg;
            ImageSource imgSrc = GetImageSource(image_Create);


            System.Drawing.Image image_Update = Properties.Resources.穿牆套管ICON合集_複製外參_svg;
            ImageSource imgSrc2 = GetImageSource(image_Update);

            System.Drawing.Image image_SetUp = Properties.Resources.穿牆套管ICON合集_編號_svg;
            ImageSource imgSrc3 = GetImageSource(image_SetUp);


            System.Drawing.Image image_Num = Properties.Resources.穿牆套管ICON合集_重新編號_svg;
            ImageSource imgSrc4 = GetImageSource(image_Num);

            System.Drawing.Image image_Rect = Properties.Resources.穿牆套管ICON合集_方開口_svg;
            ImageSource imgSrc5 = GetImageSource(image_Rect);

            System.Drawing.Image image_MultiRect = Properties.Resources.穿牆套管ICON合集_多管方開口_svg;
            ImageSource imgSrc6 = GetImageSource(image_MultiRect);


            // create the button data
            PushButtonData btnData0 = new PushButtonData(
             "MyButton_WallCastUpdate",
             "更新\n   穿牆資訊   ",
             Assembly.GetExecutingAssembly().Location,
             "CEC_WallCast.WallCastUpdate"//按鈕的全名-->要依照需要參照的command打入
             );
            {
                btnData0.ToolTip = "一鍵更新穿牆開口資訊";
                btnData0.LongDescription = "一鍵更新穿牆開口資訊";
                btnData0.LargeImage = imgSrc0;
            };

            PushButtonData btnData = new PushButtonData(
                "MyButton_WallCastCreate",
                "   穿牆套管   ",
                Assembly.GetExecutingAssembly().Location,
                "CEC_WallCast.CreateWallCastV2"//按鈕的全名-->要依照需要參照的command打入
                );
            {
                btnData.ToolTip = "點選外參牆與管生成穿牆套管";
                btnData.LongDescription = "先點選需要創建的管段，再點選其穿過的外參牆，生成穿牆套管";
                btnData.LargeImage = imgSrc;
            };


            PushButtonData btnData2 = new PushButtonData(
                "MyButton_WallCastCopy",
                "複製外參\n   穿牆套管   ",
                Assembly.GetExecutingAssembly().Location,
                "CEC_WallCast.CopyAllWallCast"
                );
            {
                btnData2.ToolTip = "複製所有連結模型中的套管";
                btnData2.LongDescription = "複製所有連結模型中的套管，以供SEM開口編號用";
                btnData2.LargeImage = imgSrc2;
            }


            PushButtonData btnData3 = new PushButtonData(
    "MyButton_WallCastNum",
    "穿牆套管\n   編號   ",
    Assembly.GetExecutingAssembly().Location,
    "CEC_WallCast.UpdateWallCastNumber"
    );
            {
                btnData3.ToolTip = "穿牆套管自動編號";
                btnData3.LongDescription = "根據每層樓的開口數量與位置，依序自動帶入編號，第二次上入編號時則會略過已經填入編號的套管";
                btnData3.LargeImage = imgSrc3;
            }

            PushButtonData btnData4 = new PushButtonData(
"MyButton_WallCastReNum",
"穿牆套管\n   重新編號   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.ReUpdateWallCastNumber"
);
            {
                btnData4.ToolTip = "穿牆套管重新編號";
                btnData4.LongDescription = "根據每層樓的開口數量，重新帶入編號";
                btnData4.LargeImage = imgSrc4;
            }

            PushButtonData btnData5 = new PushButtonData(
"MyButton_WallCastRect",
"   方型牆開口   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.CreateRectWallCast"
);
            {
                btnData5.ToolTip = "點選外參牆與管生成穿牆方開口";
                btnData5.LongDescription = "先點選需要創建的管段，再點選其穿過的外參牆，生成穿牆方開口";
                btnData5.LargeImage = imgSrc5;
            }

            PushButtonData btnData6 = new PushButtonData(
"MyButton_WallCastRectMulti",
"   多管牆開口   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.MultiWallRectCast"
);
            {
                btnData6.ToolTip = "點選外參牆與多支管生成穿牆方開口";
                btnData6.LongDescription = "先點選需要創建的管段(複數)，再點選其穿過的外參牆，生成穿牆方開口";
                btnData6.LargeImage = imgSrc6;
            }


            //創建穿牆套管(圓&方)
            PushButton button0 = panel.AddItem(btnData0) as PushButton;
            PushButton button = panel.AddItem(btnData) as PushButton;
            SplitButtonData rectCastButtonData = new SplitButtonData("WallCastRect", "方型牆開口");
            SplitButton splitButton = panel.AddItem(rectCastButtonData) as SplitButton;
            PushButton button5 = splitButton.AddPushButton(btnData5);
            button5 = splitButton.AddPushButton(btnData6);

            //複製所有套管
            PushButton button2 = panel2.AddItem(btnData2) as PushButton;

            //穿樑套管編號(編號&重編)
            SplitButtonData setNumButtonData = new SplitButtonData("WallCastSetNumButton", "穿牆套管編號");
            SplitButton splitButton2 = panel2.AddItem(setNumButtonData) as SplitButton;
            PushButton button3 = splitButton2.AddPushButton(btnData3);
            button3 = splitButton2.AddPushButton(btnData4);
;


            //預設Enabled本來就為true，不用特別設定
            //button0.Enabled = true;
            //button.Enabled = true;
            ////splitButton1.Enabled = true;
            ////splitButton2.Enabled = true;
            ////button2.Enabled = true;
            ////button4.Enabled = true;
            ////splitButton1.Enabled = true;
            ////splitButton2.Enabled = true;
            //button2.Enabled = true;
            //button3.Enabled = true;
            //button4.Enabled = true;
            //button5.Enabled = true;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }

        private BitmapSource GetImageSource(Image img)
        {
            //製作一個function專門來處理圖片
            BitmapImage bmp = new BitmapImage();
            using (MemoryStream ms = new MemoryStream())
            {
                img.Save(ms, ImageFormat.Png);
                ms.Position = 0;

                bmp.BeginInit();

                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.UriSource = null;
                bmp.StreamSource = ms;

                bmp.EndInit();
            }
            return bmp;
        }
    }
}
