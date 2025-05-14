pageextension 50100 CustomerListExt extends "Customer List"
{
    actions
    {
        addfirst(processing)
        {
            action(ConvertWordToHTML)
            {
                ApplicationArea = All;
                Caption = 'Convert Word To HTML';
                Image = Change;
                Promoted = true;
                PromotedCategory = Process;
                trigger OnAction()
                var
                    WordFileName: Text[100];
                    HTMLFileName: Text[100];
                    InStr: InStream;
                    OutStr: OutStream;
                    FileMgt: Codeunit "File Management";
                    TempBlob: Codeunit "Temp Blob";
                    DocumentReportMgt: Codeunit "Document Report Mgt.";
                begin
                    TempBlob.CreateOutStream(OutStr);
                    FileMgt.BLOBImport(TempBlob, WordFileName);
                    DocumentReportMgt.ConvertWordToHtml(TempBlob);
                    TempBlob.CreateInStream(InStr);
                    HTMLFileName := 'Converted' + '.html';
                    DownloadFromStream(InStr, '', '', '', HTMLFileName);
                end;
            }
        }
    }
}
