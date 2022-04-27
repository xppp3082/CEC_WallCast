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
using System.Windows.Media; // for the graphics �ݤޥ�prsentationCore
using System.Windows.Media.Imaging;

#endregion

namespace CEC_WallCast
{
    class App : IExternalApplication
    {
        //���ձN��Lbutton�[��{��TAB
        const string RIBBON_TAB = "�iCEC MEP�j";
        const string RIBBON_PANEL = "����}�f";
        const string RIBBON_PANEL2 = "����CSD&SEM";
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
            //�Ыءu��ٮM�ޡv����
            List<RibbonPanel> panels = a.GetRibbonPanels(RIBBON_TAB); //�b���n�T�ORIBBON_TAB�b�o�椧�e�w�g�Q�Ы�
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

            //�ЫءuSEM&CSD�v����
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
            System.Drawing.Image image_CreateST = Properties.Resources.����M��ICON�X��_��s_svg;
            ImageSource imgSrc0 = GetImageSource(image_CreateST);

            System.Drawing.Image image_Create = Properties.Resources.����M��ICON�X��_��m_svg;
            ImageSource imgSrc = GetImageSource(image_Create);


            System.Drawing.Image image_Update = Properties.Resources.����M��ICON�X��_�ƻs�~��_svg;
            ImageSource imgSrc2 = GetImageSource(image_Update);

            System.Drawing.Image image_SetUp = Properties.Resources.����M��ICON�X��_�s��_svg;
            ImageSource imgSrc3 = GetImageSource(image_SetUp);


            System.Drawing.Image image_Num = Properties.Resources.����M��ICON�X��_���s�s��_svg;
            ImageSource imgSrc4 = GetImageSource(image_Num);

            System.Drawing.Image image_Rect = Properties.Resources.����M��ICON�X��_��}�f_svg;
            ImageSource imgSrc5 = GetImageSource(image_Rect);

            System.Drawing.Image image_MultiRect = Properties.Resources.����M��ICON�X��_�h�ޤ�}�f_svg;
            ImageSource imgSrc6 = GetImageSource(image_MultiRect);


            // create the button data
            PushButtonData btnData0 = new PushButtonData(
             "MyButton_WallCastUpdate",
             "��s\n   �����T   ",
             Assembly.GetExecutingAssembly().Location,
             "CEC_WallCast.WallCastUpdate"//���s�����W-->�n�̷ӻݭn�ѷӪ�command���J
             );
            {
                btnData0.ToolTip = "�@���s����}�f��T";
                btnData0.LongDescription = "�@���s����}�f��T";
                btnData0.LargeImage = imgSrc0;
            };

            PushButtonData btnData = new PushButtonData(
                "MyButton_WallCastCreate",
                "   ����M��   ",
                Assembly.GetExecutingAssembly().Location,
                "CEC_WallCast.CreateWallCastV2"//���s�����W-->�n�̷ӻݭn�ѷӪ�command���J
                );
            {
                btnData.ToolTip = "�I��~����P�ޥͦ�����M��";
                btnData.LongDescription = "���I��ݭn�Ыت��ެq�A�A�I����L���~����A�ͦ�����M��";
                btnData.LargeImage = imgSrc;
            };


            PushButtonData btnData2 = new PushButtonData(
                "MyButton_WallCastCopy",
                "�ƻs�~��\n   ����M��   ",
                Assembly.GetExecutingAssembly().Location,
                "CEC_WallCast.CopyAllWallCast"
                );
            {
                btnData2.ToolTip = "�ƻs�Ҧ��s���ҫ������M��";
                btnData2.LongDescription = "�ƻs�Ҧ��s���ҫ������M�ޡA�H��SEM�}�f�s����";
                btnData2.LargeImage = imgSrc2;
            }


            PushButtonData btnData3 = new PushButtonData(
    "MyButton_WallCastNum",
    "����M��\n   �s��   ",
    Assembly.GetExecutingAssembly().Location,
    "CEC_WallCast.UpdateWallCastNumber"
    );
            {
                btnData3.ToolTip = "����M�ަ۰ʽs��";
                btnData3.LongDescription = "�ھڨC�h�Ӫ��}�f�ƶq�P��m�A�̧Ǧ۰ʱa�J�s���A�ĤG���W�J�s���ɫh�|���L�w�g��J�s�����M��";
                btnData3.LargeImage = imgSrc3;
            }

            PushButtonData btnData4 = new PushButtonData(
"MyButton_WallCastReNum",
"����M��\n   ���s�s��   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.ReUpdateWallCastNumber"
);
            {
                btnData4.ToolTip = "����M�ޭ��s�s��";
                btnData4.LongDescription = "�ھڨC�h�Ӫ��}�f�ƶq�A���s�a�J�s��";
                btnData4.LargeImage = imgSrc4;
            }

            PushButtonData btnData5 = new PushButtonData(
"MyButton_WallCastRect",
"   �諬��}�f   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.CreateRectWallCast"
);
            {
                btnData5.ToolTip = "�I��~����P�ޥͦ������}�f";
                btnData5.LongDescription = "���I��ݭn�Ыت��ެq�A�A�I����L���~����A�ͦ������}�f";
                btnData5.LargeImage = imgSrc5;
            }

            PushButtonData btnData6 = new PushButtonData(
"MyButton_WallCastRectMulti",
"   �h����}�f   ",
Assembly.GetExecutingAssembly().Location,
"CEC_WallCast.MultiWallRectCast"
);
            {
                btnData6.ToolTip = "�I��~����P�h��ޥͦ������}�f";
                btnData6.LongDescription = "���I��ݭn�Ыت��ެq(�Ƽ�)�A�A�I����L���~����A�ͦ������}�f";
                btnData6.LargeImage = imgSrc6;
            }


            //�Ыج���M��(��&��)
            PushButton button0 = panel.AddItem(btnData0) as PushButton;
            PushButton button = panel.AddItem(btnData) as PushButton;
            SplitButtonData rectCastButtonData = new SplitButtonData("WallCastRect", "�諬��}�f");
            SplitButton splitButton = panel.AddItem(rectCastButtonData) as SplitButton;
            PushButton button5 = splitButton.AddPushButton(btnData5);
            button5 = splitButton.AddPushButton(btnData6);

            //�ƻs�Ҧ��M��
            PushButton button2 = panel2.AddItem(btnData2) as PushButton;

            //��ٮM�޽s��(�s��&���s)
            SplitButtonData setNumButtonData = new SplitButtonData("WallCastSetNumButton", "����M�޽s��");
            SplitButton splitButton2 = panel2.AddItem(setNumButtonData) as SplitButton;
            PushButton button3 = splitButton2.AddPushButton(btnData3);
            button3 = splitButton2.AddPushButton(btnData4);
;


            //�w�]Enabled���ӴN��true�A���ίS�O�]�w
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
            //�s�@�@��function�M���ӳB�z�Ϥ�
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
