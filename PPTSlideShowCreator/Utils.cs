using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Data;
using System.Windows.Forms;
using System.Threading;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using Access = Microsoft.Office.Interop.Access;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace PPTSlideShowCreator
{
    class Utils
    {
        public static Bitmap ResizeImage(Image imgToResize, Size size)
        {
            int sourceWidth = imgToResize.Width;
            int sourceHeight = imgToResize.Height;

            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;

            nPercentW = ((float)size.Width / (float)sourceWidth);
            nPercentH = ((float)size.Height / (float)sourceHeight);

            if (nPercentH < nPercentW)
                nPercent = nPercentH;
            else
                nPercent = nPercentW;

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);

            Bitmap b = new Bitmap(destWidth, destHeight);
            Graphics g = Graphics.FromImage((Image)b);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            g.DrawImage(imgToResize, 0, 0, destWidth, destHeight);
            g.Dispose();

            return b;
        }

        public static Image OpenImage(string path)
        {
            Bitmap imageFile;
            byte[] imageFileBytes;
            // name is passed.
            using (FileStream fsImageFile = File.OpenRead(path))
            {
                imageFileBytes = new byte[fsImageFile.Length];
                fsImageFile.Read(imageFileBytes, 0, imageFileBytes.Length);

                imageFile = new Bitmap(fsImageFile);

                // Determine the format of the image file.
                // This sample code supports working with the following types of image files:
                //
                // Bitmap (BMP)
                // Graphics Interchange Format (GIF)
                // Joint Photographic Experts Group (JPG, JPEG)
                // Portable Network Graphics (PNG)
                // Tagged Image File Format (TIFF)

                /*if (imageFile.RawFormat.Guid == ImageFormat.Bmp.Guid)
                    imagePartType = ImagePartType.Bmp;
                else if (imageFile.RawFormat.Guid == ImageFormat.Gif.Guid)
                    imagePartType = ImagePartType.Gif;
                else if (imageFile.RawFormat.Guid == ImageFormat.Jpeg.Guid)
                    imagePartType = ImagePartType.Jpeg;
                else if (imageFile.RawFormat.Guid == ImageFormat.Png.Guid)
                    imagePartType = ImagePartType.Png;
                else if (imageFile.RawFormat.Guid == ImageFormat.Tiff.Guid)
                    imagePartType = ImagePartType.Tiff;
                else
                {
                    throw new ArgumentException("Unsupported image file format: " + imageFilePath);
                }*/
            }
            return imageFile;
        }

        /// <summary>
        /// Executes a shell command synchronously.
        /// </summary>
        /// <param name="command">string command</param>
        /// <returns>string, as output of the command.</returns>
        public static void OpenPowerPoint(string file)
        {
            
            Microsoft.Office.Interop.PowerPoint.Application pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
            Microsoft.Office.Core.MsoTriState ofalse = Microsoft.Office.Core.MsoTriState.msoFalse;
            Microsoft.Office.Core.MsoTriState otrue = Microsoft.Office.Core.MsoTriState.msoTrue;
            pptApp.Visible = otrue;
            pptApp.Activate();
            Microsoft.Office.Interop.PowerPoint.Presentations ps = pptApp.Presentations;
            Microsoft.Office.Interop.PowerPoint.Presentation p = ps.Open(file, ofalse, ofalse, otrue);
            System.Diagnostics.Debug.Print(p.Windows.Count.ToString());
            MessageBox.Show(pptApp.ActiveWindow.Caption);
        }
    }
}
