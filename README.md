# Syncfusion ASP.NET Core – Offer Letter Generation Demo

This repository contains a complete showcase sample demonstrating how to build an **Automated Offer Letter Generation System** using **Syncfusion DocIO** and **Syncfusion Excel** , **Syncfusion PDF** and **Syncfusion Smart Data Extractor** library in an ASP.NET Core MVC application. The sample illustrates how HR professionals can streamline offer letter creation by merging Excel data with Word templates, supporting both single and bulk document generation with mail merge capabilities.

---

## 📁 Project Structure

```
├── Controllers/
│   └── HomeController.cs
├── Models/
│   └── OfferLetterGenerationViewModel.cs
├── Views/
│   ├── Home/
│   │   └── Index.cshtml
│   └── Shared/
├── wwwroot/
│   └── Data/
│       ├── Template.docx
│       ├── Data.xlsx
└── README.md
```

---

## 🚀 Getting Started

### Prerequisites

- [.NET 7.0 SDK](https://dotnet.microsoft.com/download) or later
- Visual Studio 2022 or VS Code
- A valid **Syncfusion License Key** (or use the free Community License)

### 1. Clone the Repository

```bash
git clone https://github.com/SyncfusionExamples/DocIO-Offer-Letter-Automation
cd Offer-Letter-Automation
```

### 2. Install Dependencies

Restore all NuGet packages:

```bash
dotnet restore
```

### 3. Add Syncfusion License Key

In your `Program.cs`, register your Syncfusion license:

```csharp
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");
```

### 4. Run the Application

```bash
dotnet run
```

Open your browser and navigate to:
```
https://localhost:5001
```

---

## 📋 How to Use

### Basic Workflow

1. **Upload Template Document**
   - Upload your Word template with merge fields
   - Drag & drop supported
   - Or use default template provided

2. **Upload Data File**
   - Upload Excel file containing candidate data
   - First row should contain headers matching your template merge fields
   - Or use default candidate data file

3. **Select Candidates**
   - All Candidates – Generate offer letters for all records in Excel
   - Specific Offers – Enter offer codes to generate letters for specific candidates
   - Note: Data file must contain an Offer Code or identifier column for specific candidate filtering

4. **Choose Output Format**
   - **DOCX** – Editable Word document
   - **PDF** – Distribution-ready format

5. **Configure Extra Features**
   - **Separate Files** – Generate individual files per policy (ZIP archive)
   - **Digital Signature** *(PDF only)* 
     - Upload your signature image (PNG, JPG, etc.)
     - Enter keywords where you want signatures placed (e.g., "Signature, WITNESS")
     - Signatures will automatically appear above those keywords in the PDF

6. **Generate Offer Letters**
   - Click "Generate Offer Letters" button
   - Download automatically starts
   - Single file or ZIP archive based on selection

---

## 🎯 Use Cases

### Insurance Document Scenarios

- **New Hire Offers** – Create personalized offer letters from template and candidate data
- **Bulk Offer Generation** – Generate multiple offer letters in a single operation
- **Position-Specific Templates** – Merge data with role-specific offer letter templates
- **Counter Offers** – Produce counter offer documents with updated compensation terms

---

## 🔗 Resources

- [Syncfusion DocIO Getting Started](https://help.syncfusion.com/document-processing/word/word-library/net/getting-started)
- [Mail Merge](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-mailmerge)
- [Syncfusion Excel Getting Started](https://help.syncfusion.com/document-processing/excel/excel-library/net/overview)
- [Word to PDF Conversion](https://help.syncfusion.com/document-processing/word/conversions/word-to-pdf/overview)
- [Smart Data Extractor](https://help.syncfusion.com/document-processing/data-extraction/smart-data-extractor/overview)

---

## 📣 Try It Out

Clone the repository, run the sample, and discover how **Syncfusion DocIO** can revolutionize policy document generation in your insurance operations.

### Customization

This sample application is provided as a reference implementation and can be freely customized to suit your specific business requirements.

You can modify the templates, data sources, merge logic, and output formats based on your use case. If you have any questions, need clarification, or require assistance while customizing this sample, please feel free to contact our [Syncfusion Support Team](https://support.syncfusion.com/support/tickets/create) for guidance.

---

## 📄 License and Copyright

> This is a commercial product and requires a paid license for possession or use. Syncfusion® licensed software, including this component, is subject to the terms and conditions of Syncfusion®. To acquire a license, visit https://www.syncfusion.com/account/downloads.

Are you already a Syncfusion user? You can download the product setup [here](https://www.syncfusion.com/account/downloads). If you're not yet a Syncfusion user, you can download a [30-day free trial](https://www.syncfusion.com/downloads).

---

## 📞 Support

For technical support and questions:
- [Syncfusion Support Portal](https://support.syncfusion.com/support/tickets/create)
- [Documentation](https://help.syncfusion.com/)
- [Community Forums](https://www.syncfusion.com/forums)

---