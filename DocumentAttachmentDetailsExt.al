pageextension 50112 DocumentAttachmentDetailsExt extends "Document Attachment Details"
{
    actions
    {
        addafter(Preview_Promoted)
        {
            actionref(ConvertWordToHTML_Promoted; ConvertWordToHTML)
            {
            }
        }
        addfirst(processing)
        {
            action(ConvertWordToHTML)
            {
                Caption = 'Convert Word To HTML';
                ApplicationArea = All;
                Image = View;
                trigger OnAction()
                var
                    DocAttach: Record "Document Attachment";
                    DocAttachHTML: Record "Document Attachment";
                    InStr: InStream;
                    OutStr: OutStream;
                    TempBlob: Codeunit "Temp Blob";
                    DocumentReportMgt: Codeunit "Document Report Mgt.";
                    FileName: Text;
                begin
                    DocAttach.Reset();
                    CurrPage.SetSelectionFilter(DocAttach);
                    if DocAttach.FindFirst() then
                        if DocAttach."Document Reference ID".HasValue then begin
                            TempBlob.CreateOutStream(OutStr);
                            DocAttach."Document Reference ID".ExportStream(OutStr);
                            DocumentReportMgt.ConvertWordToHtml(TempBlob);
                            TempBlob.CreateInStream(InStr);
                            FileName := DocAttach."File Name" + ' (1)';
                            DocAttachHTML.Init();
                            DocAttachHTML.Validate("Table ID", DocAttach."Table ID");
                            DocAttachHTML.Validate("No.", DocAttach."No.");
                            DocAttachHTML.Validate("File Name", FileName);
                            DocAttachHTML.Validate("File Extension", 'html');
                            DocAttachHTML."Document Reference ID".ImportStream(InStr, FileName);
                            DocAttachHTML.Insert(true);
                        end;
                end;
            }
        }
    }
}
