using Microsoft.Office.Core;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace VSTODATN.FunctionsExcel
{
    internal class ExportJson
    {
        public static void ExportExcelToJson(Excel.Application excellApp)
        {
            int startRow = 1;
            int startColumn = 1;
            int endRow = 10;
            int endColumn = 100;

            JObject json = new JObject();
            JObject config = new JObject();
            config["visible"] = excellApp.Visible;
            config["activatesheet"] = true;
            config["terminate"] = false;
            json["config"] = config;
            JArray sheets = new JArray();
            JArray cells = new JArray();
            Excel.Workbook workbook = excellApp.ActiveWorkbook;
            foreach (Excel.Worksheet sheet1 in workbook.Sheets)
            {
                JObject sheet = new JObject();
                sheet["name"] = sheet1.Name;
                sheet["visible"] = true;
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int column = startColumn; column <= endColumn; column++)
                    {

                        Excel.Range cell = sheet1.Cells[row, column];
                        if (cell.Value2 != null)
                        {
                            JObject cell1 = new JObject();
                            if (column > 26)
                            {
                                cell1["pos"] = ((char)(column / 26 + 64)).ToString() + ((char)(column % 26 + 64)).ToString() + row;
                            }
                            else
                            {
                                cell1["pos"] = ((char)(column % 26 + 64)).ToString() + row;
                            }
                            cell1["value"] = cell.Value != null ? cell.Value.ToString() : string.Empty;
                            cells.Add(cell1);
                        }

                    }
                }

                sheet["cells"] = cells;
                sheets.Add(sheet);
            }

            json["sheets"] = sheets;

            string jsonContent = json.ToString();

            try
            {

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Json (*.json)|*.json|All files (*.*)|*.*";
                saveFileDialog.Title = "Save File";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    FileStream fileStream = File.Create(filePath);
                    fileStream.Close();
                    File.WriteAllText(filePath, jsonContent);

                }
                else
                {
                    MessageBox.Show("File creation canceled by the user.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating the file: " + ex.Message);
            }
        }
        public static void ExportExcelToJsonFormat(Excel.Application excellApp)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Json (*.json)|*.json|All files (*.*)|*.*";
            openFileDialog.Title = "Open File";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            string filePath = openFileDialog.FileName;
            string jsonContent = File.ReadAllText(filePath);
            JObject json = new JObject();
            dynamic jsonData = JsonConvert.DeserializeObject(jsonContent);
            Excel.Workbook workbook = excellApp.ActiveWorkbook;
            if (jsonData["config"] == null)
            {
                MessageBox.Show("File format error");
                return;
            }
                json["config"] = jsonData["config"];
                JArray sheets = (JArray)jsonData["sheets"];
                JArray sheets1 = new JArray();
                foreach (JObject sheetData in sheets)
                {
                    string sheetName = (string)sheetData["name"];
                    bool checkSheetName = false;
                    Excel.Worksheet worksheet;
                    foreach (Excel.Worksheet sheet in workbook.Sheets)
                    {
                        if (sheet.Name == sheetName)
                        {
                            checkSheetName = true;

                        }
                    }
                    if (checkSheetName)
                    {
                        worksheet = workbook.Worksheets[sheetName];

                        JArray cells = (JArray)sheetData["cells"];
                        foreach (JObject cellData in cells)
                        {
                            string cellPos = (string)cellData["pos"];
                            Excel.Range cell = worksheet.Range[cellPos];
                            if (cell.Value != null)
                            {
                                string cellValue = cell.Value.ToString();
                                cellData["value"] = cell.Value.ToString();
                            }
                        }
                    }

                }
                json["sheets"] = sheets;
                string jsonContentExport = json.ToString();
                try
                {
                    File.WriteAllText(filePath, jsonContentExport);
                    MessageBox.Show("File export successful");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error creating the file: " + ex.Message);
                }
     
        }
        public static void ExportExcelToJson1(Excel.Workbook workbook)
        {
            JArray jsonArray = new JArray();
            DocumentProperties properties = workbook.BuiltinDocumentProperties;
            JObject jsonObject = new JObject();
            foreach (Microsoft.Office.Core.DocumentProperty prop in properties)
            {
                if (prop.Name != null||prop.Value!=null)
                {
                    string propertyName = prop.Name;
                    string propertyValue = prop.Value.ToString();
                    jsonObject.Add(propertyName, propertyValue);
                }
            }
            jsonArray.Add(jsonObject);
            //Get the custom properties of the workbook
            //DocumentProperties customProperties = workbook.CustomDocumentProperties;
            //foreach (DocumentProperty customProp in customProperties)
            //{
            //    string propertyName = customProp.Name;
            //    string propertyValue = customProp.Value;
            //    jsonObject.Add(propertyName, propertyValue);
            //    // Do something with the custom property name and value
            //}
            //ICustomDocumentProperties customProperties = workbook.CustomDocumentProperties;
            //for (int i = 0; i < customProperties.Count; i++)
            //{
            //    string propertyName = customProperties[i].Name;
            //    string propertyValue = customProperties[i].Text;
            //    jsonObject.Add(propertyName, propertyValue);
            //    Do something with the custom property name and value
            //}
            //jsonArray.Add(jsonObject);
            string jsonContent = jsonArray.ToString();

            // Lưu chuỗi JSON vào tệp
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Json (*.json)|*.json|All files (*.*)|*.*";
                saveFileDialog.Title = "Save File";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    FileStream fileStream = File.Create(filePath);
                    fileStream.Close();
                    File.WriteAllText(filePath, jsonContent);

                }
                else
                {
                    MessageBox.Show("File creation canceled by the user.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating the file: " + ex.Message);
            }

        }
    }
}
