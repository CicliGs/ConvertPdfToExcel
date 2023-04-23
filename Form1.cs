using System;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace FormConvertPdfToXls
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Show file open dialog for PDF file
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PDF files (*.pdf)|*.pdf";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string pdfFilePath = openFileDialog.FileName;

                // Show file save dialog for Excel file
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xls)|*.xls";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string excelFilePath = saveFileDialog.FileName;

                    // Convert PDF to Excel
                    PdfReader pdfReader = new PdfReader(pdfFilePath);
                    HSSFWorkbook workbook = new HSSFWorkbook();
                    ISheet sheet = workbook.CreateSheet("Sheet1");

                    int rowNumber = 0;
                    for (int pageNumber = 1; pageNumber <= pdfReader.NumberOfPages; pageNumber++)
                    {
                        IRow row = sheet.CreateRow(rowNumber);

                        string pageText = PdfTextExtractor.GetTextFromPage(pdfReader, pageNumber, new LocationTextExtractionStrategy());
                        string[] lines = pageText.Split(new[] { "\n" }, StringSplitOptions.None);

                        foreach (string line in lines)
                        {
                            string[] values = line.Split('\t');
                            IRow newRow = sheet.CreateRow(rowNumber);
                            for (int j = 0; j < values.Length; j++)
                            {
                                newRow.CreateCell(j).SetCellValue(values[j]);
                            }
                            rowNumber++;
                        }
                    }

                    // Auto-size columns
                    for (int i = 0; i < sheet.GetRow(0).LastCellNum; i++)
                    {
                        sheet.AutoSizeColumn(i);
                    }

                    using (var fileStream = new System.IO.FileStream(excelFilePath, System.IO.FileMode.Create))
                    {
                        workbook.Write(fileStream);
                    }

                    MessageBox.Show("PDF to Excel conversion complete.");
                }
            }
        }
    }
}