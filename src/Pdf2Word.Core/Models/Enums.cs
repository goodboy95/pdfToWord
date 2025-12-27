namespace Pdf2Word.Core.Models;

public enum HeaderFooterRemoveMode
{
    None,
    RemoveHeader,
    RemoveFooter,
    RemoveBoth
}

public enum PageSizeMode
{
    A4,
    FollowPdf
}

public enum ParagraphRole
{
    Title,
    Body
}

public enum JobStatus
{
    Pending,
    Running,
    Succeeded,
    Failed,
    Canceled
}

public enum ErrorSeverity
{
    Fatal,
    Recoverable,
    Warning
}

public enum JobStage
{
    Init,
    PdfOpen,
    PdfRender,
    Crop,
    Preprocess,
    TableDetect,
    TableGrid,
    GeminiTableOcr,
    GeminiPageOcr,
    AssembleIr,
    DocxWrite,
    Finalize
}
