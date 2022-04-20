using Google.Apis.Docs.v1.Data;

namespace plot_document_sync;

public static class HandleDocument
{
    public const string PlotDocumentId = "1wqQgIPYHDs9DO_pvTghzFsQ16DSAqq5kOciIB21mHQI";
    
    public static IEnumerable<string> GetDocumentContent(Document document)
    {
        IList<StructuralElement> content = document.Body.Content;

        List<string> documentContent = new();

        foreach (StructuralElement value in content)
        {
            if (value.Paragraph is null) continue;

            string paragraphContent = string.Empty;
            foreach (ParagraphElement element in value.Paragraph.Elements)
            {
                if (element.TextRun is null) continue;

                paragraphContent += element.TextRun.Content.Trim();
            }

            if (paragraphContent == string.Empty)
            {
                continue;
            }

            documentContent.Add(paragraphContent);
        }
        
        return documentContent;
    }
}