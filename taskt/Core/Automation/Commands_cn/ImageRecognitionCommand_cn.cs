using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using System.Drawing;
using System.Windows.Forms;
using taskt.UI.Forms;
using taskt.UI.CustomControls;
using taskt.Properties;
using taskt.Core.Automation.Commands;

namespace taskt.Core.Automation.Commands_cn
{

    [Serializable]
    [Attributes.ClassAttributes.Group("(Image Commands)图像命令")]
    [Attributes.ClassAttributes.Description("此命令尝试在屏幕上查找现有图像。")]
    [Attributes.ClassAttributes.UsesDescription("如果要尝试在屏幕上查找图像，请使用此命令。 您随后可以执行诸如将鼠标移动到该位置或执行单击等操作。 此命令从比较图像生成指纹并在桌面上搜索它。")]
    [Attributes.ClassAttributes.ImplementationDescription("TBD")]
    public class ImageRecognitionCommand_cn : ScriptCommand
    {

        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("捕获搜索图像")]
        [Attributes.PropertyAttributes.PropertyUIHelper(Attributes.PropertyAttributes.PropertyUIHelper.UIAdditionalHelperType.ShowImageRecogitionHelper)]
        [Attributes.PropertyAttributes.InputSpecification("使用该工具捕获图像")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("该图像将用作在屏幕上找到的图像。")]
        public string v_ImageCapture { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("偏移X坐标 - 可选")]
        [Attributes.PropertyAttributes.InputSpecification("指定是否需要偏移量。")]
        [Attributes.PropertyAttributes.SampleUsage("0 or 100")]
        [Attributes.PropertyAttributes.Remarks("这会将鼠标X像素移动到图像位置的右侧")]
        public int v_xOffsetAdjustment { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("偏移Y坐标 - 可选")]
        [Attributes.PropertyAttributes.InputSpecification("指定是否需要偏移量。")]
        [Attributes.PropertyAttributes.SampleUsage("0 or 100")]
        [Attributes.PropertyAttributes.Remarks("这将从图像位置的顶部向下移动鼠标X像素")]
        public int v_YOffsetAdjustment { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("如果需要，请指明鼠标点击类型")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("没有")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("左键单击")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("中间点击")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("右键点击")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("左下")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("中下来")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("右下")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("左上")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("中间")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("马上")]
        [Attributes.PropertyAttributes.PropertyUISelectionOption("双击左键")]
        [Attributes.PropertyAttributes.InputSpecification("指出所需的点击类型")]
        [Attributes.PropertyAttributes.SampleUsage("选择**左键单击**，**中键点击**，**右键单击**，**左键单击**，**左下**，**中下**，**右下* *，**左上**，**中上**，**右上**")]
        [Attributes.PropertyAttributes.Remarks("您可以通过连续使用多个鼠标单击命令来模拟自定义单击，在需要的位置之间添加**暂停命令**。")]
        public string v_MouseClick { get; set; }
        [XmlAttribute]
        [Attributes.PropertyAttributes.PropertyDescription("超时（秒数，0表示无限制搜索时间）")]
        [Attributes.PropertyAttributes.InputSpecification("如果需要，输入超时长度。")]
        [Attributes.PropertyAttributes.SampleUsage("")]
        [Attributes.PropertyAttributes.Remarks("对于诸如白色的颜色，搜索时间变得过多。 为获得最佳效果，请在屏幕上捕获较大的颜色差异，而不仅仅是白色块。")]
        public double v_TimeoutSeconds { get; set; }
        public bool TestMode = false;
        public ImageRecognitionCommand_cn()
        {
            this.CommandName = "ImageRecognitionCommand";
            //this.SelectionName = "Image Recognition";
            this.SelectionName = Settings.Default.Image_Recognition_cn;
            this.CommandEnabled = true;
            this.CustomRendering = true;

            v_xOffsetAdjustment = 0;
            v_YOffsetAdjustment = 0;
            v_TimeoutSeconds = 30;
        }
        public override void RunCommand(object sender)
        {
            bool testMode = TestMode;
           
            //user image to bitmap
            Bitmap userImage = new Bitmap(Core.Common.Base64ToImage(v_ImageCapture));

            //take screenshot
            Size shotSize = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Size;
            Point upperScreenPoint = new Point(0, 0);
            Point upperDestinationPoint = new Point(0, 0);
            Bitmap desktopImage = new Bitmap(shotSize.Width, shotSize.Height);
            Graphics graphics = Graphics.FromImage(desktopImage);
            graphics.CopyFromScreen(upperScreenPoint, upperDestinationPoint, shotSize);

            //create desktopOutput file
            Bitmap desktopOutput = new Bitmap(desktopImage);

            //get graphics for drawing on output file
            Graphics screenShotUpdate = Graphics.FromImage(desktopOutput);

            //declare maximum boundaries
            int userImageMaxWidth = userImage.Width - 1;
            int userImageMaxHeight = userImage.Height - 1;
            int desktopImageMaxWidth = desktopImage.Width - 1;
            int desktopImageMaxHeight = desktopImage.Height - 1;

            //newfingerprinttechnique

            //create desktopOutput file
            Bitmap sampleOut = new Bitmap(userImage);

            //get graphics for drawing on output file
            Graphics sampleUpdate = Graphics.FromImage(sampleOut);

            List<ImageRecognitionFingerPrint> uniqueFingerprint = new List<ImageRecognitionFingerPrint>();
            Color lastcolor = Color.Transparent;

            //create fingerprint
            var pixelDensity = (userImage.Width * userImage.Height);

            int iteration = 0;
            Random random = new Random();
            while ((uniqueFingerprint.Count() < 10) && (iteration < pixelDensity))
            {
                int x = random.Next(userImage.Width);
                int y = random.Next(userImage.Height);
                Color color = sampleOut.GetPixel(x, y);

                if ((lastcolor != color) && (!uniqueFingerprint.Any(f => f.xLocation == x && f.yLocation == y)))
                {
                    uniqueFingerprint.Add(new ImageRecognitionFingerPrint() { PixelColor = color, xLocation = x, yLocation = y });
                    sampleUpdate.DrawRectangle(Pens.Yellow, x, y, 1, 1);
                }

                iteration++;
            }




            //begin search
            DateTime timeoutDue = DateTime.Now.AddSeconds(v_TimeoutSeconds);


            bool imageFound = false;
            //for each row on the screen
            for (int rowPixel = 0; rowPixel < desktopImage.Height - 1; rowPixel++)
            {

                if (rowPixel + uniqueFingerprint.First().yLocation >= desktopImage.Height)
                    continue;


                //for each column on screen
                for (int columnPixel = 0; columnPixel < desktopImage.Width - 1; columnPixel++)
                {

                    if ((v_TimeoutSeconds > 0) && (DateTime.Now > timeoutDue))
                    {
                        throw new Exception("Image recognition command ran out of time searching for image");
                    }

                    if (columnPixel + uniqueFingerprint.First().xLocation >= desktopImage.Width)
                        continue;



                    try
                    {




                        //get the current pixel from current row and column
                        // userImageFingerPrint.First() for now will always be from top left (0,0)
                        var currentPixel = desktopImage.GetPixel(columnPixel + uniqueFingerprint.First().xLocation, rowPixel + uniqueFingerprint.First().yLocation);

                        //compare to see if desktop pixel matches top left pixel from user image
                        if (currentPixel == uniqueFingerprint.First().PixelColor)
                        {

                            //look through each item in the fingerprint to see if offset pixel colors match
                            int matchCount = 0;
                            for (int item = 0; item < uniqueFingerprint.Count; item++)
                            {
                                //find pixel color from offset X,Y relative to current position of row and column
                                currentPixel = desktopImage.GetPixel(columnPixel + uniqueFingerprint[item].xLocation, rowPixel + uniqueFingerprint[item].yLocation);

                                //if color matches
                                if (uniqueFingerprint[item].PixelColor == currentPixel)
                                {
                                    matchCount++;

                                    //draw on output to demonstrate finding
                                    if (testMode)
                                        screenShotUpdate.DrawRectangle(Pens.Blue, columnPixel + uniqueFingerprint[item].xLocation, rowPixel + uniqueFingerprint[item].yLocation, 5, 5);
                                }
                                else
                                {
                                    //mismatch in the pixel series, not a series of matching coordinate
                                    //?add threshold %?
                                    imageFound = false;

                                    //draw on output to demonstrate finding
                                    if (testMode)
                                        screenShotUpdate.DrawRectangle(Pens.OrangeRed, columnPixel + uniqueFingerprint[item].xLocation, rowPixel + uniqueFingerprint[item].yLocation, 5, 5);
                                }

                            }

                            if (matchCount == uniqueFingerprint.Count())
                            {
                                imageFound = true;

                                var topLeftX = columnPixel;
                                var topLeftY = rowPixel;

                                if (testMode)
                                {
                                    //draw on output to demonstrate finding
                                    var Rectangle = new Rectangle(topLeftX, topLeftY, userImageMaxWidth, userImageMaxHeight);
                                    Brush brush = new SolidBrush(Color.ForestGreen);
                                    screenShotUpdate.FillRectangle(brush, Rectangle);
                                }

                                //move mouse to position
                                var mouseMove = new SendMouseMoveCommand
                                {
                                    v_XMousePosition = (topLeftX + (v_xOffsetAdjustment)).ToString(),
                                    v_YMousePosition = (topLeftY + (v_xOffsetAdjustment)).ToString(),
                                    v_MouseClick = v_MouseClick
                                };

                                mouseMove.RunCommand(sender);


                            }



                        }


                        if (imageFound)
                            break;

                    }
                    catch (Exception)
                    {
                        //continue
                    }
                }


                if (imageFound)
                    break;
            }













            if (testMode)
            {
                //screenShotUpdate.FillRectangle(Brushes.White, 5, 20, 275, 105);
                //screenShotUpdate.DrawString("Blue = Matching Point", new Font("Arial", 12, FontStyle.Bold), Brushes.SteelBlue, 5, 20);
                // screenShotUpdate.DrawString("OrangeRed = Mismatched Point", new Font("Arial", 12, FontStyle.Bold), Brushes.SteelBlue, 5, 60);
                // screenShotUpdate.DrawString("Green Rectangle = Match Area", new Font("Arial", 12, FontStyle.Bold), Brushes.SteelBlue, 5, 100);

                //screenShotUpdate.DrawImage(sampleOut, desktopOutput.Width - sampleOut.Width, 0);

                UI.Forms.Supplement_Forms.frmImageCapture captureOutput = new UI.Forms.Supplement_Forms.frmImageCapture();
                captureOutput.pbTaggedImage.Image = sampleOut;
                captureOutput.pbSearchResult.Image = desktopOutput;
                captureOutput.Show();
                captureOutput.TopMost = true;
                //captureOutput.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            }

            graphics.Dispose();
            userImage.Dispose();
            desktopImage.Dispose();
            screenShotUpdate.Dispose();

            if (!imageFound)
            {
                throw new Exception("Specified image was not found in window!");
            }


        }
        public override List<Control> Render(frmCommandEditor editor)
        {
            base.Render(editor);


            UIPictureBox imageCapture = new UIPictureBox();
            imageCapture.Width = 200;
            imageCapture.Height = 200;
            imageCapture.DataBindings.Add("EncodedImage", this, "v_ImageCapture", false, DataSourceUpdateMode.OnPropertyChanged);

            RenderedControls.Add(CommandControls.CreateDefaultLabelFor("v_ImageCapture", this));
            RenderedControls.AddRange(CommandControls.CreateUIHelpersFor("v_ImageCapture", this, new Control[] { imageCapture }, editor));
            RenderedControls.Add(imageCapture);

            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_xOffsetAdjustment", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_YOffsetAdjustment", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultDropdownGroupFor("v_MouseClick", this, editor));
            RenderedControls.AddRange(CommandControls.CreateDefaultInputGroupFor("v_TimeoutSeconds", this, editor));


            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + " [Find On Screen]";
        }
    }
}