using Emgu.CV;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Emgu.CV;
using Emgu.CV.Structure;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            EmguFunction function = new EmguFunction(new Size(1500, 1000), new Rectangle(30, 30, 1440, 940));
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "Image files (*.bmp, *.jpg, *.jpeg, *.tif , *.tiff, *.png) |" +
                "*.bmp; *.jpg; *.jpeg; *.tif; *.tiff; *.png";
            if (op.ShowDialog() == DialogResult.OK)
            {
                Mat inputImage = CvInvoke.Imread(op.FileName, Emgu.CV.CvEnum.ImreadModes.Unchanged);
                if (inputImage.Depth != Emgu.CV.CvEnum.DepthType.Cv8U)
                {
                    double min = 0;
                    double max = 0;
                    Point minLoc = new Point();
                    Point maxLoc = new Point();
                    CvInvoke.MinMaxLoc(inputImage,ref min,ref max,ref minLoc,ref maxLoc);
                 
                    double scale = 255 / (max - min);
                    CvInvoke.ConvertScaleAbs(inputImage, inputImage, 1, -min);
                    CvInvoke.ConvertScaleAbs(inputImage, inputImage, scale, 0);
                    inputImage.ConvertTo(inputImage, Emgu.CV.CvEnum.DepthType.Cv8U);
                }
                CvInvoke.Blur(inputImage, inputImage, new Size(5, 5), new Point(2, 2));
                pictureBox1.Image = function.Run(inputImage.Bitmap,(int)numericUpDown2.Value,(double)numericUpDown1.Value,
                    (int)numericUpDown3.Value, (int)numericUpDown4.Value, (int)numericUpDown5.Value);
            }
        }
    }
}
