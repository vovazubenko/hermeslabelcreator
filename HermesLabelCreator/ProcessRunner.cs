using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HermesLabelCreator.Configurations;
using HermesLabelCreator.Helpers;
using HermesLabelCreator.Interfaces;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HermesLabelCreator
{
    public class ProcessRunner : IProcessRunner
    {
        private readonly string importFolder = "_import";
        private readonly string exportFolder = "_export";
        private readonly string[] supportedExtensions = new string[] { "xls", "xlsx", "csv" };
        private readonly ILogger<ProcessRunner> _logger;
        private readonly Iservice _service;

        public ProcessRunner(
            ILogger<ProcessRunner> logger, 
            Iservice service)
        {
            _logger = logger;
            _service = service;
        }
        public void Run(string directoriesPath, bool singleFile)
        {
            string fullImportPath = Path.Combine(directoriesPath, importFolder);
            string fullExportPath = Path.Combine(directoriesPath, exportFolder);

            if (!Directory.Exists(fullExportPath))
            {
                Directory.CreateDirectory(fullExportPath);
            }

            IEnumerable<string> importFiles = Directory
                .EnumerateFiles(fullImportPath, "*.*", SearchOption.AllDirectories)
                .Where(s => supportedExtensions.Contains(Path.GetExtension(s).TrimStart('.').ToLowerInvariant()));

            foreach (var importFile in importFiles)
            {
                List<string> labelPaths = new List<string>();

                string extension = Path.GetExtension(importFile);
                string newImportFile = importFile;
                if (extension.Equals(".csv"))
                {
                    newImportFile = ExcelManager.ConvertCsvToExcel(importFile);
                }

                string exportFile = $"{fullExportPath}/{(extension.Equals(".csv") ? "csv" : "excel")}_export_{DateTime.Now:ddMMyyyy_hhmmss}{Path.GetExtension(newImportFile)}";
                File.Copy(newImportFile, exportFile, true);
                
                FileStream importFileStream = File.OpenRead(newImportFile);
                var rows = ExcelManager.GetRows(importFileStream);
                var shipments = ImportFileParser.ParseShipments(rows);

                FileStream exportFileStream = File.Open(exportFile, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(exportFileStream, true);
                WorksheetPart worksheetPart = spreadSheet.WorkbookPart.WorksheetParts.First();

                try
                {
                    var options = new ParallelOptions { MaxDegreeOfParallelism = 1 };
                    Parallel.For(0, shipments.Length, options, (i) =>
                    {
                        if (!string.IsNullOrWhiteSpace(shipments[i].ShipmentID))
                        {
                            return;
                        }

                        var response = _service.GenerateShipmentLabel(shipments[i]);
                        if (!response.Success || response.Delivery == null)
                        {
                            _logger.LogInfo(i + 2, string.IsNullOrWhiteSpace(response.ResponseString)
                                ? response.Message
                                : response.ResponseString);

                            return;
                        }

                        string orderNumber = !string.IsNullOrWhiteSpace(shipments[i].OrderNumber)
                            ? $"{shipments[i].OrderNumber}_"
                            : string.Empty;

                        if (response.Delivery != null && !string.IsNullOrWhiteSpace(response.Delivery.LabelBase64) && shipments[i].Returns != "2")
                        {
                            string ending = "H";
                            if (shipments[i].BulkyGoods == "1")
                            {
                                ending = "S";
                            }
                            else if (shipments[i].NextDay == "1")
                            {
                                ending = "N";
                            }
                            else if (shipments[i].Kws == "1")
                            {
                                ending = "K";
                            }

                            string fileName = $"{(i+1):D3}_{orderNumber}{shipments[i].PuNumber}_{ending}_{DateTime.Now:ddMMyyyy-hhmm}.pdf";
                            string fullFilePath = Path.Combine(fullExportPath, fileName);
                            CreateLabelFile(fullFilePath, response.Delivery.LabelBase64);

                            labelPaths.Add(fullFilePath);
                        }

                        if (response.Returns != null && !string.IsNullOrWhiteSpace(response.Returns.LabelBase64))
                        {
                            string fileName = $"{(i+1):D3}_{orderNumber}{shipments[i].PuNumber}_R_{DateTime.Now:ddMMyyyy-hhmm}.pdf";
                            string fullFilePath = Path.Combine(fullExportPath, fileName);
                            CreateLabelFile(fullFilePath, response.Returns.LabelBase64);

                            labelPaths.Add(fullFilePath);
                        }

                        string barcode = string.Empty;
                        if (response.Delivery.LabelContent?.Import?.BarcodeObject != null && response.Delivery.LabelContent.Import.BarcodeObject.Length > 0)
                        {
                            barcode = response.Delivery.LabelContent.Import.BarcodeObject[0].BarcodeFormatted;
                        }

                        string barcodeAtHermes = string.Empty;
                        if (response.Delivery.LabelContent?.Export?.BarcodeObject != null && response.Delivery.LabelContent.Export.BarcodeObject.Length > 0)
                        {
                            barcodeAtHermes = response.Delivery.LabelContent.Export.BarcodeObject[0].BarcodeFormatted;
                        }

                        if (string.IsNullOrWhiteSpace(barcodeAtHermes))
                        {
                            lock (exportFileStream)
                            {
                                ExcelManager.UpdateCell(barcode, "B", (uint)i + 2, worksheetPart);
                            }
                        }
                        else
                        {
                            lock (exportFileStream)
                            {
                                ExcelManager.UpdateCell(barcode, "A", (uint)i + 2, worksheetPart);
                            }

                            lock (exportFileStream)
                            {
                                ExcelManager.UpdateCell(barcodeAtHermes, "B", (uint)i + 2, worksheetPart);
                            }
                        }

                        if (response.Returns != null)
                        {
                            string shipmentReturn = string.Empty;
                            string barcodeAtHermesReturn = string.Empty;

                            if (response.Returns.LabelContent?.Import?.BarcodeObject != null && response.Returns.LabelContent.Import.BarcodeObject.Length > 0)
                            {
                                shipmentReturn = response.Returns.LabelContent.Import.BarcodeObject[0].BarcodeFormatted;
                            }

                            if (response.Returns.LabelContent?.Export?.BarcodeObject != null && response.Returns.LabelContent.Export.BarcodeObject.Length > 0)
                            {
                                barcodeAtHermesReturn = response.Returns.LabelContent.Export.BarcodeObject[0].BarcodeFormatted;
                            }

                            lock (exportFileStream)
                            {
                                ExcelManager.UpdateCell(shipmentReturn, "C", (uint)i + 2, worksheetPart);
                            }

                            lock (exportFileStream)
                            {
                                ExcelManager.UpdateCell(barcodeAtHermesReturn, "D", (uint)i + 2, worksheetPart);
                            }
                        }

                    });

                    importFileStream.Close();
                    spreadSheet.Dispose();
                    exportFileStream.Dispose();

                    if (extension.Equals(".csv"))
                    {
                        ExcelManager.ConvertExcelToCsv(exportFile);
                        File.Delete(exportFile);
                    }

                    if (singleFile && labelPaths.Count > 0)
                    {
                        string[] sortedLabelPaths = labelPaths.OrderBy(l => l).ToArray();
                        string combinedLabelsFilePath = $"{fullExportPath}/Labels_{Path.GetFileNameWithoutExtension(exportFile)}.pdf";

                        CombineMultiplePDFs(sortedLabelPaths, combinedLabelsFilePath);

                        foreach (string labelPath in labelPaths)
                        {
                            File.Delete(labelPath);
                        }
                    }

                    //File.Delete(newImportFile);
                    //if (newImportFile != importFile)
                    //{
                    //    File.Delete(importFile);
                    //}
                }
                catch
                {
                    throw;
                }
                finally
                {
                    
                }
            }
        }

        private void CreateLabelFile(string fullFilePath, string fileBase64)
        {
            if (File.Exists(fullFilePath))
            {
                File.Delete(fullFilePath);
            }

            byte[] bytes = Convert.FromBase64String(fileBase64);
            FileStream stream = new FileStream(fullFilePath, FileMode.CreateNew);
            using (BinaryWriter writer = new BinaryWriter(stream))
            {
                writer.Write(bytes, 0, bytes.Length);
            }
        }

        public static void CombineMultiplePDFs(string[] fileNames, string outFile)
        {
            // step 1: creation of a document-object
            Document document = new Document();
            //create newFileStream object which will be disposed at the end
            using (FileStream newFileStream = new FileStream(outFile, FileMode.Create))
            {
                // step 2: we create a writer that listens to the document
                PdfCopy writer = new PdfCopy(document, newFileStream);

                // step 3: we open the document
                document.Open();

                foreach (string fileName in fileNames)
                {
                    // we create a reader for a certain document
                    PdfReader reader = new PdfReader(fileName);
                    reader.ConsolidateNamedDestinations();

                    // step 4: we add content
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        PdfImportedPage page = writer.GetImportedPage(reader, i);
                        writer.AddPage(page);
                    }

                    PRAcroForm form = reader.AcroForm;
                    if (form != null)
                    {
                        writer.AddDocument(reader);
                    }

                    reader.Close();
                }

                // step 5: we close the document and writer
                writer.Close();
                document.Close();
            }//disposes the newFileStream object
        }
    }
}
