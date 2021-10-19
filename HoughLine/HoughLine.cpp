// HoughLine.cpp: 主要專案檔。

#include "stdafx.h"
#include "Form1.h"

using namespace HoughLine;
using namespace Emgu::CV;
using namespace Emgu::Util;
using namespace Emgu::CV::Structure;
using namespace Emgu::CV::Util;

[STAThreadAttribute]
int main(array<System::String ^> ^args)
{
	// 建立任何控制項之前，先啟用 Windows XP 視覺化效果
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false); 

	// 建立主視窗並執行
	Application::Run(gcnew Form1());
	return 0;
}

void Form1::Form1_Load(System::Object^ sender, System::EventArgs^ e)
{
	//this->_milApp = MappAlloc(M_DEFAULT, M_NULL);
	//this->_milSys = MsysAlloc(M_SYSTEM_HOST, 0, M_DEFAULT, M_NULL);
}

void Form1::Form1_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e)
{
	//if (this->_milSys != M_NULL)
	//{
	//	MsysFree(this->_milSys);
	//}
	//if (this->_milApp != M_NULL)
	//{
	//	MappFree(this->_milApp);
	//}
}

void Form1::button1_Click(System::Object^ sender, System::EventArgs^ e)
{
	OpenFileDialog ^op = gcnew OpenFileDialog();
	op->Filter = "Image files (*.bmp, *.jpg, *.jpeg, *.tif , *.tiff, *.png) |" +
		"*.bmp; *.jpg; *.jpeg; *.tif; *.tiff; *.png";
	if (op->ShowDialog() == ::DialogResult::OK)
	{
		try
		{
			//MIL_ID milImg = M_NULL;
			////int sizeX = 6131;
			////int sizeY = 4042;
			////MIL to Mat
			//pin_ptr<const wchar_t> wch = PtrToStringChars(op->FileName);

			////wchar_t *charFilePath = wch;
			//MIL_INT sizeX = MbufDiskInquire(wch, M_SIZE_X, M_NULL);
			//MIL_INT sizeY = MbufDiskInquire(wch, M_SIZE_Y, M_NULL);
			//MIL_INT imgBit = MbufDiskInquire(wch, M_SIZE_BIT, M_NULL);
			//milImg = MbufAlloc2d(this->_milSys, sizeX, sizeY, imgBit + M_UNSIGNED, M_PROC + M_IMAGE, M_NULL);
			//MbufLoad(wch , milImg);

			//unsigned char *buf = nullptr;
			//buf = new unsigned char[sizeX * sizeY * 2]();
			//MbufGet(milImg, buf);

			//IntPtr dotnet_buf(buf);
			//Emgu::CV::Mat ^inputImage;
			//inputImage = gcnew Emgu::CV::Mat(System::Drawing::Size(sizeX,sizeY), Emgu::CV::CvEnum::DepthType::Cv16U, 1, dotnet_buf, sizeX*2);

			Emgu::CV::Mat ^inputImage;
			inputImage = Emgu::CV::CvInvoke::Imread(op->FileName, Emgu::CV::CvEnum::LoadImageType::Unchanged);
			if (inputImage->Depth != Emgu::CV::CvEnum::DepthType::Cv8U)
			{
				CvInvoke::Normalize(inputImage, inputImage, 255, 0, Emgu::CV::CvEnum::NormType::MinMax, Emgu::CV::CvEnum::DepthType::Cv16U, nullptr);
				inputImage->ConvertTo(inputImage, Emgu::CV::CvEnum::DepthType::Cv8U , 1, 0);
				CvInvoke::Imwrite("D:\\Temp\\TEST_Normalize.bmp", inputImage);
			}
			//System::Drawing::Size kernelSize = gcnew System::Drawing::Size(5,5);
			//System::Drawing::Point kernelCenter = gcnew System::Drawing::Point(-1,-1);
			CvInvoke::Blur(inputImage, inputImage, System::Drawing::Size(5,5), System::Drawing::Point(-1,-1),CvEnum::BorderType::Reflect101);

			Emgu::CV::Image<Emgu::CV::Structure::Gray, System::Byte> ^img_8U;
			img_8U = gcnew Emgu::CV::Image<Emgu::CV::Structure::Gray, System::Byte>(inputImage->Bitmap);
			int blockSize = (int)this->numericUpDown2->Value;
			double offset = (double)this->numericUpDown1->Value;
			int ExculdeArea = (int)this->numericUpDown3->Value;
			Emgu::CV::Image<Emgu::CV::Structure::Gray, System::Byte> ^drawimage;
			drawimage = img_8U->CopyBlank();
			Emgu::CV::CvInvoke::AdaptiveThreshold(img_8U, img_8U, 255, Emgu::CV::CvEnum::AdaptiveThresholdType::MeanC, Emgu::CV::CvEnum::ThresholdType::Binary
				, blockSize, offset);
			CvInvoke::Imwrite("D:\\Temp\\TEST_Binary.bmp", img_8U);

			Emgu::CV::Util::VectorOfVectorOfPoint ^use_vvp = gcnew Emgu::CV::Util::VectorOfVectorOfPoint();
			Emgu::CV::Util::VectorOfVectorOfPoint ^vvp = gcnew Emgu::CV::Util::VectorOfVectorOfPoint();
			CvInvoke::FindContours(img_8U, vvp, nullptr, Emgu::CV::CvEnum::RetrType::External, Emgu::CV::CvEnum::ChainApproxMethod::ChainApproxSimple, System::Drawing::Point(0,0));
			int number = vvp->ToArrayOfArray().Length;//
			for (int i = 0; i < number; i++)
			{
				Emgu::CV::Util::VectorOfPoint ^vp = vvp[i];
				double area = CvInvoke::ContourArea(vp, false);
				if ( area > ExculdeArea)
				{
					use_vvp->Push(vp);
				}
			}
			CvInvoke::DrawContours(drawimage, use_vvp, -1, Emgu::CV::Structure::MCvScalar(255), -1 , Emgu::CV::CvEnum::LineType::AntiAlias, nullptr, 2147483647, System::Drawing::Point(0,0));
			CvInvoke::Imwrite("D:\\Temp\\TEST_DrawContours.bmp", drawimage);
			
			int resize_sizeX = inputImage->Width/2;
			int resize_sizeY = inputImage->Height/2;
			int offsetX = 50 ;
			int offsetY = 50 ;

			img_8U = drawimage->Resize(resize_sizeX, resize_sizeY, Emgu::CV::CvEnum::Inter::Area);
			img_8U->ROI = Rectangle(offsetX, offsetY, resize_sizeX - (offsetX*2), resize_sizeY - (offsetY*2));
			
			img_8U = img_8U->Copy();
			int LineTH = (int)this->numericUpDown6->Value;
			double minLineLength = (double)this->numericUpDown4->Value;
			double lineGap = (double)this->numericUpDown5->Value;
			array<Emgu::CV::Structure::LineSegment2D> ^lines;
			//Emgu::CV::Structure::LineSegment2D[] ^lines;
			lines = CvInvoke::HoughLinesP(img_8U, 1, Math::PI / 180.0, LineTH, minLineLength, lineGap);
			Emgu::CV::Mat ^dst ;
			Emgu::CV::Image<Emgu::CV::Structure::Gray, System::Byte> ^img_HoughLine_Binary;
			img_HoughLine_Binary = img_8U->CopyBlank();
			dst = gcnew Emgu::CV::Mat();

			CvInvoke::CvtColor(img_8U, dst, Emgu::CV::CvEnum::ColorConversion::Gray2Bgr, 0);

			for (int i = 0; i < lines.Length; i++)
			{				

				//CvInvoke::Line(dst, lines[i].P1, lines[i].P2, Emgu::CV::Structure::MCvScalar(0, 0, 255), 3, Emgu::CV::CvEnum::LineType::AntiAlias, 0);
				double deltaX = (lines[i].P1.X - lines[i].P2.X);
				double deltaY = (lines[i].P1.Y - lines[i].P2.Y);
				if (deltaX == 0)
				{
					CvInvoke::Line(img_HoughLine_Binary, lines[i].P1, lines[i].P2, Emgu::CV::Structure::MCvScalar(255), 3, Emgu::CV::CvEnum::LineType::AntiAlias, 0);
				}
				else
				{
					double slop = deltaY / deltaX; //(y2 -y1) / (x2 - x1)
					double angle = Math::Atan(slop) * 180 / Math::PI;
					if (angle > 45 && angle < 135)
					{
						CvInvoke::Line(img_HoughLine_Binary, lines[i].P1, lines[i].P2, Emgu::CV::Structure::MCvScalar(255), 3, Emgu::CV::CvEnum::LineType::AntiAlias, 0);
					}
					else if (angle > 225 && (angle < 315 || angle < -45))
					{
						CvInvoke::Line(img_HoughLine_Binary, lines[i].P1, lines[i].P2, Emgu::CV::Structure::MCvScalar(255), 3, Emgu::CV::CvEnum::LineType::AntiAlias, 0);
					}
				}					
			}
			CvInvoke::Imwrite("D:\\Temp\\TEST_Result.bmp", dst);
			vvp->Clear();
			use_vvp->Clear();
			CvInvoke::FindContours(img_HoughLine_Binary, vvp, nullptr, Emgu::CV::CvEnum::RetrType::External, Emgu::CV::CvEnum::ChainApproxMethod::ChainApproxSimple, System::Drawing::Point(0,0));
			
			number = vvp->ToArrayOfArray().Length;//
			for (int i = 0; i < number; i++)
			{
				Emgu::CV::Util::VectorOfPoint ^vp = vvp[i];
				double area = CvInvoke::ContourArea(vp, false);

				//if ( area < 26000)
				//{
					System::Drawing::Rectangle objRect= CvInvoke::BoundingRectangle(vp);
					use_vvp->Push(vp);
					CvInvoke::Line(dst, System::Drawing::Point(objRect.X,objRect.Y), System::Drawing::Point(objRect.Right,objRect.Bottom), Emgu::CV::Structure::MCvScalar(0, 0, 255), 3, Emgu::CV::CvEnum::LineType::AntiAlias, 0);
				//}
			}
			CvInvoke::Imwrite("D:\\Temp\\TEST_Result.bmp", dst);
			pictureBox1->Image = gcnew System::Drawing::Bitmap(dst->Bitmap);

			//if (milImg != M_NULL)
			//{
			//	MbufFree(milImg);
			//}
			//delete dotnet_buf;

			delete inputImage;
			delete drawimage;
			delete dst;
			delete img_8U;
			delete lines;
			GC::Collect();
		}
		catch(Exception ^ex)
		{
			MessageBox::Show("ERR " + ex->Message);
		}

		return;
	}
}
