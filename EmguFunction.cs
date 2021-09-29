using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Emgu.CV;
using Emgu.Util;
using Emgu.CV.Structure;
using Emgu.CV.Util;

namespace WindowsFormsApp1
{
    class EmguFunction
    {
        Size _imageSize;
        Rectangle _roi;

        public EmguFunction(Size size , Rectangle roi)
        {
            _imageSize = size;
            _roi = roi;
        }
        public Bitmap Run(Bitmap bmp,int blockSize , double offset, int ExculdeArea,int minLineLength, int lineGap)
        {
            Bitmap result =null;

            Image<Gray, byte> image = new Image<Gray, byte>(bmp);

            var drawimage = image.CopyBlank();
            CvInvoke.AdaptiveThreshold(image, image, 255, Emgu.CV.CvEnum.AdaptiveThresholdType.MeanC, Emgu.CV.CvEnum.ThresholdType.Binary
                , blockSize, offset);
            VectorOfVectorOfPoint use_vvp = new VectorOfVectorOfPoint();
            VectorOfVectorOfPoint vvp = new VectorOfVectorOfPoint();
            CvInvoke.FindContours(image, vvp, null, Emgu.CV.CvEnum.RetrType.External, Emgu.CV.CvEnum.ChainApproxMethod.ChainApproxSimple);
            int number = vvp.ToArrayOfArray().Length;//取得轮廓的数量.
            for (int i = 0; i < number; i++)
            {
                VectorOfPoint vp = vvp[i];
                double area = CvInvoke.ContourArea(vp);
                if (area > ExculdeArea)//可按实际图片修改
                {
                    use_vvp.Push(vp);
                }
            }
            CvInvoke.DrawContours(drawimage, use_vvp, -1, new MCvScalar(255), -1);

            image = drawimage.Resize(_imageSize.Width, _imageSize.Height, Emgu.CV.CvEnum.Inter.Area);
            image.ROI = _roi;

            image = image.Copy();

            Mat dst = new Mat();
            CvInvoke.CvtColor(image, dst, Emgu.CV.CvEnum.ColorConversion.Gray2Bgr);
            LineSegment2D[] lines = CvInvoke.HoughLinesP(image, 1, Math.PI / 180.0, 165, minLineLength, lineGap);

            for (int i = 0; i < lines.Length; i++)
            {
                CvInvoke.Line(dst, lines[i].P1, lines[i].P2, new MCvScalar(0, 0, 255), 3, Emgu.CV.CvEnum.LineType.AntiAlias);
            }

            dst.Save(@"D:\---AUO\[00]PROJECT\---AreaGrabber\01.Code\[Image]\達智慧\TPK\RemoveHV\TEST.bmp");
            result = new Bitmap(dst.Bitmap);
            dst.Dispose();
            image.Dispose();
            drawimage.Dispose();
            return result;
        }
        

    }
}
