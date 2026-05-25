using System.ComponentModel.DataAnnotations;

namespace Offer_Letter_Automation.Models
{
    public class OfferLetterGenerationViewModel
    {
        /// <summary>
        /// Word template file (.docx, .doc, .dotx)
        /// Contains merge fields like «CandidateName», «PositionTitle», «Salary», etc.
        /// </summary>
        [Required(ErrorMessage = "Template file is required")]
        public IFormFile? TemplateFile { get; set; }

        /// <summary>
        /// Excel data file (.xlsx, .xls, .csv)
        /// Contains candidate records with columns matching template merge fields
        /// </summary>
        [Required(ErrorMessage = "Candidate data file is required")]
        public IFormFile? DataFile { get; set; }

        /// <summary>
        /// Offer letter selection mode: "all" or "single"
        /// all - Generate offer letters for all candidates in Excel
        /// single - Generate offer letter for specific candidate by offer code
        /// </summary>
        [Required(ErrorMessage = "Selection mode is required")]
        public string SelectionMode { get; set; } = "all";

        /// <summary>
        /// Unique offer code for single offer letter generation
        /// Only used when SelectionMode = "single"
        /// Example: OFF-2026-001, POL-HR-123
        /// </summary>
        public string? OfferCode { get; set; }

        /// <summary>
        /// Processing mode: "single" or "batch"
        /// single - Generate one document with all candidates
        /// batch - Generate individual document per candidate (returns ZIP)
        /// </summary>
        [Required(ErrorMessage = "Processing mode is required")]
        public string ProcessType { get; set; } = "batch";

        /// <summary>
        /// Output document format: "docx" or "pdf"
        /// docx - Editable Word document
        /// pdf - Professional PDF document
        /// </summary>
        [Required(ErrorMessage = "Output format is required")]
        public string OutputFormat { get; set; } = "docx";

        /// <summary>
        /// Flag to include candidate signature in offer letter
        /// If true, signature will be added to generated documents
        /// Default: false
        /// </summary>
        public bool EnableDigitalSign { get; set; } = false;

        /// <summary>
        /// Signature image file (.png, .jpg, .jpeg)
        /// Optional - If not provided, default signature from project directory will be used
        /// Recommended size: 200x80px transparent PNG for best results
        /// Only processed when EnableDigitalSign = true
        /// </summary>
        public IFormFile? SignatureFile { get; set; }

        /// <summary>
        /// Comma-separated keywords to identify signature placement locations in the document
        /// Example: "Sign, WITNESS, AuthorizedSign, Signature"
        /// Signatures will be placed above these keywords in the PDF document
        /// Default: "Sign, WITNESS, AuthorizedSign"
        /// </summary>
        [Display(Name = "Signature Keywords")]
        public string SignatureKeywords { get; set; } = "Signature, CandidateSign";



    }
}
