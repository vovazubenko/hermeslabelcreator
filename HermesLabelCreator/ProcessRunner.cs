using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HermesLabelCreator.Configurations;
using HermesLabelCreator.Helpers;
using HermesLabelCreator.Interfaces;
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
        private readonly string[] supportedExtensions = new string[] { "xls", "xlsx" };
        private readonly ILogger<ProcessRunner> _logger;
        private readonly Iservice _service;

        public ProcessRunner(
            ILogger<ProcessRunner> logger, 
            Iservice service)
        {
            _logger = logger;
            _service = service;
        }
        public void Run(string directoriesPath)
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
                Path.GetExtension(importFile);
                string exportFile = $"{fullExportPath}\\csv_export_{DateTime.Now:ddMMyyyy_hhmm}{Path.GetExtension(importFile)}";
                File.Copy(importFile, exportFile, true);
                
                FileStream importFileStream = File.OpenRead(importFile);
                var rows = ExcelManager.GetRows(importFileStream);
                var shipments = ImportFileParser.ParseShipments(rows);

                FileStream exportFileStream = File.Open(exportFile, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(exportFileStream, true);
                WorksheetPart worksheetPart = spreadSheet.WorkbookPart.WorksheetParts.First();

                try
                {
                    var options = new ParallelOptions { MaxDegreeOfParallelism = 4 };
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

                        if (response.Delivery != null && !string.IsNullOrWhiteSpace(response.Delivery.LabelBase64))
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

                            string fileName = $"{orderNumber}{shipments[i].PuNumber}{ending}.pdf";
                            string fullFilePath = Path.Combine(fullExportPath, fileName);
                            CreateLabelFile(fullFilePath, response.Delivery.LabelBase64);
                        }

                        if (response.Returns != null && !string.IsNullOrWhiteSpace(response.Returns.LabelBase64))
                        {
                            string fileName = $"{orderNumber}{shipments[i].PuNumber}R.pdf";
                            string fullFilePath = Path.Combine(fullExportPath, fileName);
                            CreateLabelFile(fullFilePath, response.Returns.LabelBase64);
                        }

                        string barcode = string.Empty;
                        if (response.Delivery.LabelContent?.Import?.BarcodeObject != null && response.Delivery.LabelContent.Import.BarcodeObject.Length > 0)
                        {
                            barcode = response.Delivery.LabelContent.Import.BarcodeObject[0].BarcodeFormatted;
                        }

                        lock(exportFileStream)
                        {
                            ExcelManager.UpdateCell(barcode, "A", (uint)i + 2, worksheetPart);
                        }

                        string barcodeAtHermes = string.Empty;
                        if (response.Delivery.LabelContent?.Export?.BarcodeObject != null && response.Delivery.LabelContent.Export.BarcodeObject.Length > 0)
                        {
                            barcodeAtHermes = response.Delivery.LabelContent.Export.BarcodeObject[0].BarcodeFormatted;
                        }

                        lock (exportFileStream)
                        {
                            ExcelManager.UpdateCell(barcodeAtHermes, "B", (uint)i + 2, worksheetPart);
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
                    File.Delete(importFile);
                }
                catch
                {
                    throw;
                }
                finally
                {
                    spreadSheet.Dispose();
                    exportFileStream.Dispose();
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
    }
}
