using Microsoft.AspNetCore.Mvc;
using Offer_Letter_Automation.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Pdf.Security;
using Syncfusion.SmartDataExtractor;
using Syncfusion.XlsIO;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.IO.Compression;

namespace Offer_Letter_Automation.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }

        /// <summary>
        /// Main entry point for document generation. Handles both single and batch processing.
        /// - Single: Returns document directly (DOCX or PDF)
        /// - Batch: Returns ZIP archive with split documents
        /// </summary>
   
        public IActionResult GenerateOfferLetter(OfferLetterGenerationViewModel model)
        {
            try
            {
                // Get template and Excel streams
                Stream documentStream = GetWordDocument(model.TemplateFile);
                Stream excelStream = GetExcel(model.TemplateFile, model.DataFile);

                if (documentStream == null || excelStream == null)
                {
                    TempData["ErrorMessage"] = "Invalid file format.";
                    return RedirectToAction("Index");
                }

                // Parse offer codes based on selection mode
                List<string> offerCodesList = ParseOfferCodes(model);

                // Create DataSet from Excel (with filtering if single mode with specific codes)
                DataSet dataSet = CreateMailMergeDataSet(excelStream, offerCodesList);

                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    TempData["ErrorMessage"] = "No data found in the Excel file.";
                    return RedirectToAction("Index");
                }

                // For "all" mode with batch processing - extract all offer codes from dataset
                if (model.SelectionMode == "all" && model.ProcessType == "batch")
                {
                    offerCodesList = dataSet.Tables[0].AsEnumerable()
                        .Select(row => row[0]?.ToString()?.Trim())
                        .Where(p => !string.IsNullOrWhiteSpace(p))
                        .Distinct()
                        .ToList();
                }

                // Generate documents using unified method
                return CreateOfferLetters(documentStream, dataSet, offerCodesList, model);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating offer letters");
                TempData["ErrorMessage"] = $"Error: {ex.Message}";
                return RedirectToAction("Index");
            }
        }

        /// <summary>
        /// Unified document generation method supporting both single and batch outputs.
        /// Performs mail merge once, then either returns single file or splits by page breaks into ZIP.
        /// </summary>
        private IActionResult CreateOfferLetters(Stream documentStream, DataSet dataSet, List<string> offerCodesList, OfferLetterGenerationViewModel model)
        {
            try
            {
                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    TempData["ErrorMessage"] = "No data found for the selected criteria.";
                    return RedirectToAction("Index");
                }

                bool isPdf = model.OutputFormat?.ToLower() == "pdf";
                string fileExtension = isPdf ? "pdf" : "docx";
                string mimeType = isPdf ? "application/pdf" : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

                using (WordDocument document = new WordDocument(documentStream, FormatType.Automatic))
                {
                    // Execute mail merge once with all data
                    ExecuteMailMerge(document, dataSet);

                    // Update document fields
                    document.UpdateDocumentFields();

                    // Decide output based on process type
                    if (model.ProcessType == "batch" && offerCodesList.Count > 1)
                    {
                        // Split into multiple documents and return ZIP
                        byte[] zipBytes = SplitByPageBreak(document, isPdf, model.SignatureFile, model.SignatureKeywords,model.EnableDigitalSign);

                        if (zipBytes != null && zipBytes.Length > 0)
                        {
                            int count = dataSet.Tables[0].Rows.Count;
                            string zipFileName = $"OfferLetters_Batch_{count}_Files_{DateTime.Now:yyyyMMdd_HHmmss}.zip";
                            return File(zipBytes, "application/zip", zipFileName);
                        }
                        else
                        {
                            // Fallback: return as single document if split fails
                            TempData["Warning"] = "Unable to split documents. Returning as single file.";
                            using (MemoryStream outputStream = new MemoryStream())
                            {
                                SaveDocumentToStream(document, outputStream, isPdf);
                                return File(outputStream.ToArray(), mimeType, $"OfferLetters_Fallback.{fileExtension}");
                            }
                        }
                    }
                    else
                    {
                        // Return single merged document
                        using (MemoryStream outputStream = new MemoryStream())
                        {
                            SaveDocumentToStream(document, outputStream, isPdf);

                            // Generate filename based on selection
                            string fileName = $"OfferLetters.{fileExtension}";

                            if (isPdf && model.EnableDigitalSign)
                            {
                                outputStream.Position = 0;
                                using (MemoryStream signedStream = ApplyDigitalSignatureIfEnabled(
                                        outputStream,
                                        model.SignatureFile,
                                        model.SignatureKeywords,
                                        true))
                                {
                                    return File(signedStream.ToArray(), mimeType, fileName);
                                }
                            }
                            else
                                return File(outputStream.ToArray(), mimeType, fileName);                          
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating offer letters");
                TempData["ErrorMessage"] = $"Error: {ex.Message}";
                return RedirectToAction("Index");
            }
        }
        /// <summary>
        /// Helper method to save WordDocument to stream in either DOCX or PDF format.
        /// </summary>
        private void SaveDocumentToStream(WordDocument document, MemoryStream outputStream, bool convertToPdf)
        {
            if (convertToPdf)
            {
                using (DocIORenderer renderer = new DocIORenderer())
                using (PdfDocument pdfDocument = renderer.ConvertToPDF(document))
                {
                    pdfDocument.Save(outputStream);
                }
            }
            else
            {
                document.Save(outputStream, FormatType.Docx);
            }
            outputStream.Position = 0;
        }      
        //// <summary>
        /// Parses offer codes from single mode input.
        /// Only called when user selects "single" mode with specific code.
        /// </summary>
        private List<string> ParseOfferCodes(OfferLetterGenerationViewModel model)
        {
            List<string> offerCodes = new List<string>();

            if (string.IsNullOrWhiteSpace(model.OfferCode))
                return offerCodes;

            offerCodes = model.OfferCode
                .Split(new[] { '\n', '\r', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim())
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct()
                .ToList();

            return offerCodes;
        }                   
        /// <summary>
        /// Splits document by page breaks using bookmarks and returns ZIP bytes
        /// Each section between page breaks becomes a separate PDF
        /// </summary>
        private byte[] SplitByPageBreak(WordDocument wordDocument, bool isPDF, IFormFile signatureImage = null, string signatureKeywords = null, bool enableDigitalSign = false)
        {
            // Find all page breaks in the document
            List<Entity> entities = wordDocument.FindAllItemsByProperty(EntityType.Break, "BreakType", "PageBreak");
            if (entities == null || entities.Count == 0)
                return null;

            WSection section = wordDocument.Sections[0];
            WTextBody body = section.Body;
            int bookmarkIndex = 1;
            // Step 1: Insert a NEW paragraph at the very beginning with BookmarkStart
            WParagraph firstBookmarkPara = new WParagraph(wordDocument);
            firstBookmarkPara.AppendBookmarkStart($"Page_Bookmark_{bookmarkIndex}");
            body.ChildEntities.Insert(0, firstBookmarkPara);

            // Step 2: Iterate page break entities → insert bookmark paragraph directly after each
            foreach (Entity entity in entities)
            {
                WParagraph breakParagraph = entity.Owner as WParagraph;

                if (breakParagraph == null) continue;

                // Get the current index of this paragraph in the body
                int paraIndex = body.ChildEntities.IndexOf(breakParagraph);

                if (paraIndex < 0) continue;

                // Insert new paragraph right after the page break paragraph
                // Close current bookmark and open next bookmark in same paragraph
                WParagraph bookmarkPara = new WParagraph(wordDocument);
                bookmarkPara.AppendBookmarkEnd($"Page_Bookmark_{bookmarkIndex}");
                bookmarkIndex++;
                bookmarkPara.AppendBookmarkStart($"Page_Bookmark_{bookmarkIndex}");
                body.ChildEntities.Insert(paraIndex + 1, bookmarkPara);
            }

            // Step 3: Insert a NEW paragraph at the very end with BookmarkEnd
            WParagraph lastBookmarkPara = new WParagraph(wordDocument);
            lastBookmarkPara.AppendBookmarkEnd($"Page_Bookmark_{bookmarkIndex}");
            body.ChildEntities.Add(lastBookmarkPara);
            // Step 4: Create ZIP file and convert each bookmarked section to PDF
            MemoryStream zipStream = new MemoryStream();
            using (ZipArchive zip = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
            {
                for (int i = 1; i <= bookmarkIndex; i++)
                {
                    try
                    {
                        // Navigate to each bookmark section
                        BookmarksNavigator navigator = new BookmarksNavigator(wordDocument);
                        navigator.MoveToBookmark($"Page_Bookmark_{i}", true, true);
                        WordDocumentPart documentPart = navigator.GetContent();

                        if (documentPart == null) continue;

                        using (WordDocument extractedDoc = documentPart.GetAsWordDocument())
                        {
                            if (isPDF)
                            {
                                // Convert to PDF
                                using (Syncfusion.DocIORenderer.DocIORenderer renderer = new Syncfusion.DocIORenderer.DocIORenderer())
                                using (Syncfusion.Pdf.PdfDocument pdfDocument = renderer.ConvertToPDF(extractedDoc))
                                using (MemoryStream pdfStream = new MemoryStream())
                                {
                                    pdfDocument.Save(pdfStream);
                                    pdfDocument.Close();
                                    pdfStream.Position = 0;

                                    // Apply digital signatures to each PDF if enabled
                                    using (MemoryStream signedPdfStream = ApplyDigitalSignatureIfEnabled(
                                        pdfStream, signatureImage, signatureKeywords, enableDigitalSign))
                                    {
                                        // Add signed PDF to ZIP
                                        ZipArchiveEntry entry = zip.CreateEntry($"Document_{i}.pdf", System.IO.Compression.CompressionLevel.Fastest);
                                        using (Stream entryStream = entry.Open())
                                        {
                                            signedPdfStream.Position = 0;
                                            signedPdfStream.CopyTo(entryStream);
                                        }
                                    }
                                }
                            }
                            else if (!isPDF)
                            {
                                // Save as DOCX
                                using (MemoryStream docxStream = new MemoryStream())
                                {
                                    extractedDoc.Save(docxStream, FormatType.Docx);
                                    docxStream.Position = 0;

                                    ZipArchiveEntry entry = zip.CreateEntry($"Document_{i}.docx", CompressionLevel.Fastest);
                                    using (Stream entryStream = entry.Open())
                                    {
                                        docxStream.CopyTo(entryStream);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log and continue to next section if one fails
                        _logger.LogError(ex, $"Error converting bookmark {i} to PDF");
                    }
                }
            }
            return zipStream.ToArray();
        }       
        /// <summary>
        /// Creates DataSet from Excel based on selection mode.
        /// - If selectionMode = "all": Process all candidates
        /// - If selectionMode = "single": Filter by specific offer code
        /// </summary>
        public DataSet CreateMailMergeDataSet(Stream excelStream, List<string> offerCode = null)
        {
            DataSet dataSet = new DataSet();

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Open(excelStream);

                // Read all sheets into DataTables
                foreach (IWorksheet sheet in workbook.Worksheets)
                {
                    if (sheet.UsedRange == null || sheet.UsedRange.LastRow < 2)
                        continue;

                    DataTable dt = ReadExcelSheetToDataTable(
                        workbook,
                        sheet.Name,
                        offerCode);

                    if (dt != null && dt.Rows.Count > 0)
                    {
                        dataSet.Tables.Add(dt);
                    }
                }
            }

            return dataSet;
        }

        /// <summary>
        /// Reads a single Excel sheet into a DataTable with optional filtering by multiple policy numbers
        /// </summary>
        private DataTable ReadExcelSheetToDataTable(IWorkbook workbook, string sheetName, List<string> offerCode)
        {
            IWorksheet sheet = workbook.Worksheets[sheetName];
            if (sheet?.UsedRange == null)
                return null;

            IRange usedRange = sheet.UsedRange;
            int headerRow = usedRange.Row;
            int lastRow = usedRange.LastRow;
            int lastCol = usedRange.LastColumn;

            // Create DataTable
            DataTable dt = new DataTable(sheetName);

            // Add columns from header row
            for (int col = 1; col <= lastCol; col++)
            {
                string columnName = sheet[headerRow, col].Value ?? $"Column{col}";
                dt.Columns.Add(columnName);
            }

            // Find policy column using case-insensitive comparison (handles OfferCode, Offer Code, Offer_Code, etc.)
            DataColumn policyColumn = dt.Columns.Cast<DataColumn>()
                .FirstOrDefault(c => c.ColumnName.Replace(" ", "").Replace("_", "").Equals("OfferCode", StringComparison.OrdinalIgnoreCase));

            bool shouldFilter = policyColumn != null && offerCode != null && offerCode.Count > 0;
            int policyColumnIndex = shouldFilter ? dt.Columns.IndexOf(policyColumn) : -1;

            // Add data rows (skip header)
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                // Apply filter if needed
                if (shouldFilter)
                {
                    string policyValue = sheet[row, policyColumnIndex + 1].Value?.ToString()?.Trim();

                    if (string.IsNullOrWhiteSpace(policyValue) ||
                        !offerCode.Any(p => p.Equals(policyValue, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue; // Skip this row
                    }
                }

                // Add row to DataTable
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= lastCol; col++)
                {
                    dr[col - 1] = sheet[row, col].Value ?? string.Empty;
                }
                dt.Rows.Add(dr);
            }

            return dt;
        }
        /// <summary>
        /// Executes mail merge based on data structure (flat, nested, or multiple tables)
        /// Maintains parent-child relationships for nested groups
        /// </summary>
        public void ExecuteMailMerge(WordDocument document, DataSet dataSet)
        {
            if (dataSet.Tables.Count == 0)
                return;

            try
            {
                document.MailMerge.StartAtNewPage = true;

                string[] groupNames = document.MailMerge.GetMergeGroupNames();

                document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_EmployeeImage);

                // Case 1: No groups in template AND single table → Simple Execute
                if ((groupNames == null || groupNames.Length == 0) && dataSet.Tables.Count == 1)
                {
                    document.MailMerge.Execute(dataSet.Tables[0]);
                }
                // Case 2: Template has groups → ExecuteNestedGroup
                else if (groupNames != null && groupNames.Length > 0)
                {
                    ExecuteNestedGroupWithSmartCommands(document, dataSet, groupNames);
                }
                // Update document fields
                document.UpdateDocumentFields();
            }
            catch (Exception ex)
            {
                throw new Exception($"Mail merge failed: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Executes a nested mail merge operation using the provided document,
        /// dataset, and group names. Builds smart commands dynamically for nested groups.
        /// </summary>
        private void ExecuteNestedGroupWithSmartCommands(WordDocument document, DataSet dataSet, string[] groupNames)
        {
            try
            {
                // List to hold dynamically built nested group command
                ArrayList commands = new ArrayList();

                // Simply build commands based on group names
                // If groupNames is null/empty, commands will be empty and DocIO handles it            
                BuildNestedCommands(document, dataSet, groupNames, commands);

                // Execute the nested mail merge with the generated commands
                document.MailMerge.ExecuteNestedGroup(dataSet, commands);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error executing nested group merge: {ex.Message}", ex);
            }
        }
        /// <summary>
        /// Represents the method that handles MergeImageField event.
        /// </summary>
        private void MergeField_EmployeeImage(object sender, MergeImageFieldEventArgs args)
        {
            //Binds image from file system during mail merge.
            if (args.FieldName == "Photo")
            {
                string photoFileName = args.FieldValue.ToString();
                // Load digital certificate
                string defaultPath = Path.Combine(_hostingEnvironment.WebRootPath, "Data/");
                //Gets the image from file system.
                FileStream imageStream = new FileStream(Path.GetFullPath(defaultPath + photoFileName), FileMode.Open, FileAccess.Read);
                args.ImageStream = imageStream;
                WPicture pic = args.Picture;
                if (pic != null)
                {
                    pic.Height = 100;
                    pic.Width = 110;
                }
               
            }
        }
        /// <summary>
        /// Builds nested mail merge commands dynamically based on the provided group names.
        /// Each group represents a table in the dataset, and relationships are built sequentially.
        /// </summary
        private void BuildNestedCommands(WordDocument document, DataSet dataSet, string[] groupNames, ArrayList commands)
        {
            // Exit early if no group names provided
            if (groupNames == null || groupNames.Length == 0)
                return;
            // Get the first (root) group name
            string firstGroupName = groupNames[0];
            // Find matching table in dataset (case-insensitive
            DataTable firstTable = dataSet.Tables.Cast<DataTable>()
                .FirstOrDefault(t => string.Equals(t.TableName, firstGroupName, StringComparison.OrdinalIgnoreCase));
            // If root table not found, log warning and stop processing
            if (firstTable == null)
            {
                _logger.LogWarning($"First group '{firstGroupName}' not found in dataset.");
                return;
            }

            // Add root parent with empty relation (like "Employees" in your example)
            commands.Add(new DictionaryEntry(firstTable.TableName, string.Empty));

            // Track current parent - changes at each level
            string currentParentTableName = firstTable.TableName;

            // Add child tables - each relates to IMMEDIATE parent (not root parent)
            for (int i = 1; i < groupNames.Length; i++)
            {
                string childGroupName = groupNames[i];
                DataTable childTable = dataSet.Tables.Cast<DataTable>()
                    .FirstOrDefault(t => string.Equals(t.TableName, childGroupName, StringComparison.OrdinalIgnoreCase));

                if (childTable != null)
                {
                    // Build relation to immediate parent (previous table)
                    // Example: Orders relates to Customers (not to Employees)
                    string relationString = BuildRelationString(dataSet, currentParentTableName, childTable.TableName);

                    commands.Add(new DictionaryEntry(childTable.TableName, relationString));

                    if (string.IsNullOrEmpty(relationString))
                    {
                        _logger.LogInformation($"No common column found between '{currentParentTableName}' and '{childTable.TableName}'. Using empty relation.");
                    }

                    // Update parent for next iteration
                    // Next child will relate to THIS table
                    currentParentTableName = childTable.TableName;
                }
                else
                {
                    _logger.LogWarning($"Child group '{childGroupName}' not found in dataset.");
                }
            }
        }
        /// <summary>
        /// Builds a relation string between a parent and child table based on a common column name.
        /// This relation format is used by DocIO for nested mail merge operations.
        /// </summary>
        private string BuildRelationString(DataSet dataSet, string parentTableName, string childTableName)
        {
            // Retrieve parent and child tables
            DataTable parentTable = dataSet.Tables[parentTableName];
            DataTable childTable = dataSet.Tables[childTableName];

            // Return empty if either table is missing
            if (parentTable == null || childTable == null)
                return string.Empty;

            // Find common column
            HashSet<string> parentColumns = parentTable.Columns
                .Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            HashSet<string> childColumns = childTable.Columns
                .Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            // Find the first common column between parent and child
            string commonKey = parentColumns
                .Intersect(childColumns)
                .FirstOrDefault();
            // If no common column found, return empty relation
            if (commonKey == null)
                return string.Empty;

            // DocIO relation format
            return $"{commonKey} = %{parentTableName}.{commonKey}%";
        }
        /// <summary>
        /// Retrieves a Word document stream from the uploaded file or a default template.
        /// </summary>
        private Stream GetWordDocument(IFormFile file)
        {
            // Case 1: Uploaded file exists and has content
            if (file != null && file.Length > 0)
            {
                string extension = Path.GetExtension(file.FileName).ToLower();
                string[] supportedExtensions = { ".doc", ".docx", ".dot", ".dotx", ".dotm", ".md", ".xml", ".rtf", ".html" };
                // Validate the file extension
                if (supportedExtensions.Contains(extension))
                {
                    // Copy the uploaded file into an in-memory stream
                    MemoryStream stream = new MemoryStream();
                    file.CopyTo(stream);
                    // Reset stream position to the beginning for downstream reading
                    stream.Position = 0;
                    return stream;
                }
                else
                {
                    ViewBag.Message = "Please choose a Word format document to convert to PDF.";
                    return null;
                }
            }
            else
            {
                // Load default file from wwwroot\Data\
                string defaultFilePath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", "Template.docx");
                return new FileStream(defaultFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
        }
        /// <summary>
        /// Retrieves an Json stream based on the uploaded files.
        /// </summary>
        private Stream GetExcel(IFormFile file, IFormFile jsonFile)
        {
            // If Word file was uploaded, JSON file must also be uploaded
            if (file != null && file.Length > 0)
            {
                // Ensure an Json file is also uploaded
                if (jsonFile != null && jsonFile.Length > 0)
                {
                    // Copy uploaded Json file into an in-memory stream.
                    MemoryStream stream = new MemoryStream();
                    jsonFile.CopyTo(stream);
                    // Reset stream position so it can be read from the beginning
                    stream.Position = 0;
                    return stream;
                }
                else
                {
                    ViewBag.Message = "Please upload a JSON data file along with the Word document.";
                    return null;
                }
            }
            else
            {
                // Both Word and JSON are defaults (no file uploaded)
                string defaultJsonPath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", "Data.xlsx");
                return new FileStream(defaultJsonPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
        }
        /// <summary>
        /// Conditionally applies digital signatures to PDF based on enableDigitalSign flag
        /// </summary>
        private MemoryStream ApplyDigitalSignatureIfEnabled(MemoryStream inputStream, IFormFile signatureImage, string signatureKeywordsInput, bool enableDigitalSign)
        {
            Stream signatureStream = null;
            try
            {
                // Early exit if digital signature is not enabled
                if (!enableDigitalSign)
                {
                    _logger.LogInformation("Digital signature is not enabled. Returning PDF stream directly.");
                    inputStream.Position = 0;
                    return inputStream;
                }

                // Check if signature image is available
                signatureStream = GetSignatureImageStream(signatureImage);
                if (signatureStream == null)
                {
                    _logger.LogWarning("Digital signature enabled but no signature image found. Returning PDF stream directly.");
                    inputStream.Position = 0;
                    return inputStream;
                }

                // Apply digital signatures
                _logger.LogInformation("Applying digital signatures to PDF document.");

                // Initialize the extractor with required detection settings
                DataExtractor extractor = new DataExtractor { EnableFormDetection = false, EnableTableDetection = true, ConfidenceThreshold = 0.6 };

                // Extract PDF document from the input stream
                inputStream.Position = 0;
                PdfLoadedDocument pdfDocument = extractor.ExtractDataAsPdfDocument(inputStream);

                // Add signatures
                AddSignaturesToPDFDocument(pdfDocument, signatureStream, signatureKeywordsInput);

                // Save PDF with signatures
                MemoryStream outputMs = new MemoryStream();
                pdfDocument.Save(outputMs);
                pdfDocument.Close(true);
                outputMs.Position = 0;

                _logger.LogInformation("Digital signatures applied successfully.");
                return outputMs;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing PDF document with digital signatures");
                throw;
            }
            finally
            {
                signatureStream?.Dispose();
            }
        }

        /// <summary>
        /// Adds digital signatures to PDF document at locations matching specified keywords
        /// </summary>
        private void AddSignaturesToPDFDocument(PdfLoadedDocument pdfDocument, Stream signatureImageStream, string keywords)
        {
            // Use default keywords if none are provided
            string[] signatureKeywords = string.IsNullOrWhiteSpace(keywords)
                ? new[] { "Sign", "Principal", "Signature", "Director" }
                : keywords.Split(',').Select(k => k.Trim()).ToArray();

            _logger.LogInformation($"Applying digital signatures for keywords: {string.Join(", ", signatureKeywords)}");

            // Iterate through each page in the document
            for (int pageIndex = 0; pageIndex < pdfDocument.Pages.Count; pageIndex++)
            {
                PdfPageBase page = pdfDocument.Pages[pageIndex];
                TextLineCollection textLines;

                // Extract text lines from the page
                page.ExtractText(out textLines);

                // Iterate through each word on the page
                foreach (TextLine line in textLines.TextLine)
                {
                    foreach (TextWord word in line.WordCollection)
                    {
                        // Skip words that do not match any signature keyword
                        if (!signatureKeywords.Any(k => word.Text.Contains(k, StringComparison.Ordinal)))
                            continue;

                        // Calculate signature position above the keyword
                        RectangleF bounds = word.Bounds;
                        float signatureX = bounds.X;
                        float signatureY = bounds.Y - bounds.Height - 10;
                        float signatureWidth = 80;
                        float signatureHeight = 20;

                        try
                        {
                            // Load digital certificate
                            string certPath = Path.Combine(_hostingEnvironment.ContentRootPath, "PDF.pfx");

                            if (!System.IO.File.Exists(certPath))
                            {
                                _logger.LogWarning($"Certificate file not found at: {certPath}");
                                continue;
                            }

                            using System.IO.FileStream cert = new System.IO.FileStream(certPath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                            PdfCertificate pdfCert = new PdfCertificate(cert, "syncfusion");

                            // Create and configure the PDF signature
                            PdfSignature signature = new PdfSignature(pdfDocument, page, pdfCert, "Signature");
                            signature.Bounds = new RectangleF(signatureX, signatureY, signatureWidth, signatureHeight);

                            // Load signature image directly from stream
                            signatureImageStream.Position = 0;
                            PdfBitmap signatureImageBitmap = new PdfBitmap(signatureImageStream);
                            signature.Appearance.Normal.Graphics.DrawImage(signatureImageBitmap, 0, 0, signatureWidth, signatureHeight);

                            _logger.LogDebug($"Signature added at page {pageIndex + 1}, position ({signatureX}, {signatureY})");
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, $"Error adding signature at page {pageIndex + 1}");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets signature image as a stream from user upload or default signature
        /// </summary>
        private Stream GetSignatureImageStream(IFormFile signatureImage)
        {
            // If user provided an image, return its stream
            if (signatureImage != null && signatureImage.Length > 0)
            {
                _logger.LogInformation("Using user-provided signature image stream.");
                MemoryStream memoryStream = new MemoryStream();
                signatureImage.OpenReadStream().CopyTo(memoryStream);
                memoryStream.Position = 0;
                return memoryStream;
            }

            // No user image - use default signature from project
            string defaultImagePath = Path.Combine(_hostingEnvironment.ContentRootPath, "Signature.png");

            if (System.IO.File.Exists(defaultImagePath))
            {
                _logger.LogInformation($"Using default signature image from: {defaultImagePath}");
                try
                {
                    FileStream fileStream = new FileStream(defaultImagePath, FileMode.Open, FileAccess.Read);
                    MemoryStream memoryStream = new MemoryStream();
                    fileStream.CopyTo(memoryStream);
                    fileStream.Dispose();
                    memoryStream.Position = 0;
                    return memoryStream;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"Error loading default signature image: {defaultImagePath}");
                    return null;
                }
            }

            _logger.LogWarning($"Default signature image not found at: {defaultImagePath}");
            return null;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
