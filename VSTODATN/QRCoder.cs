using Microsoft.Office.Interop.Excel;
using ZXing;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace VSTODATN.FunctionsExcel
{
    internal class QRCoder
    {
        public static void GenerateQRCode(Excel.Worksheet worksheet, Excel.Range actCell)
        {

            string data = actCell.Value;
            BarcodeWriter barcodeWriter = new BarcodeWriter();
            barcodeWriter.Format = BarcodeFormat.QR_CODE;
            barcodeWriter.Options = new ZXing.Common.EncodingOptions
            {
                Width = 200,
                Height = 200
            };

            // Tạo bitmap từ mã QR
            if (data != null)
            {
                var bitmap = barcodeWriter.Write(data);
                string tempFile = Path.GetTempFileName();
                bitmap.Save(tempFile, System.Drawing.Imaging.ImageFormat.Png);
                Shapes shapes = worksheet.Shapes;
                Shape pictureShape = shapes.AddPicture(tempFile, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 100, 100, 100, 100);
                //Shape pictureShape = null;
                //pictureShape = worksheet.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, actCell.Left, actCell.Top, actCell.Width, actCell.Height);
                //pictureShape.Fill.UserPicture(tempFile);
                //System.Windows.Forms.Clipboard.SetImage(bitmap);
                //worksheet.Paste(worksheet.Range["A1"]);

                //worksheet.Pictures[worksheet.Pictures.Count].ShapeRange.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                //worksheet.Cells.EntireColumn.ColumnWidth = 15;
            }
            else
            {
                MessageBox.Show("Cell không có nội dung");
            }
        }
    }
}
