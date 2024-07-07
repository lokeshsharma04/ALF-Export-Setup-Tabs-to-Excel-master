report 50126 "ALF Export Setup Tabs to Excel"
{
    // version 1.0

    // //Please add these properties for Cloud app
    // //Caption = 'ALF Export Setup Tabs to Excel';
    // //UsageCategory = ReportsAndAnalysis;
    // //ApplicationArea = All,Basic,Suite;

    ProcessingOnly = true;
    UseRequestPage = false;

    dataset
    {
    }

    requestpage
    {

        layout
        {
        }

        actions
        {
        }
    }

    labels
    {
    }

    trigger OnPostReport();
    begin
        GlobalRecRef.OPEN(DATABASE::"General Ledger Setup");
        GlobalRecRef.FINDFIRST;
        FillHeader(StartRowNo,GlobalRecRef);
        StartRowNo := FillLine(StartRowNo,GlobalRecRef) + 2;
        GlobalRecRef.CLOSE;

        GlobalRecRef.OPEN(DATABASE::"Sales & Receivables Setup");
        GlobalRecRef.FINDFIRST;
        FillHeader(StartRowNo,GlobalRecRef);
        StartRowNo := FillLine(StartRowNo,GlobalRecRef) + 2;
        GlobalRecRef.CLOSE;

        GlobalRecRef.OPEN(DATABASE::"Purchases & Payables Setup");
        GlobalRecRef.FINDFIRST;
        FillHeader(StartRowNo,GlobalRecRef);
        StartRowNo := FillLine(StartRowNo,GlobalRecRef) + 2;
        GlobalRecRef.CLOSE;

        GlobalRecRef.OPEN(DATABASE::"Inventory Setup");
        GlobalRecRef.FINDFIRST;
        FillHeader(StartRowNo,GlobalRecRef);
        StartRowNo := FillLine(StartRowNo,GlobalRecRef) + 2;
        GlobalRecRef.CLOSE;

        TempExcelBuffer.CreateNewBook(SheetName);
        TempExcelBuffer.WriteSheet(SheetName,'','');
        TempExcelBuffer.CloseBook;
        TempExcelBuffer.OpenExcel;
        TempExcelBuffer.GiveUserControl; // please comment this line for Cloud app
    end;

    trigger OnPreReport();
    begin
        TempExcelBuffer.DELETEALL;
        RowNo := 1;
        ColNo := 1;
        EnterCell(RowNo,ColNo,FORMAT(COMPANYPROPERTY.URLNAME+';'+TENANTID+';'+USERID+';'+FORMAT(TIME)),TempExcelBuffer."Cell Type"::Text,true,false,true);
        StartRowNo := 3;
    end;

    var
        TempExcelBuffer : Record "Excel Buffer" temporary;
        RowNo : Integer;
        ColNo : Integer;
        GlobalRecRef : RecordRef;
        StartRowNo : Integer;
        SheetName : TextConst ENU='ENU=ALF Export Setup Tabs to Excel';

    local procedure EnterCell(RowNo : Integer;ColumnNo : Integer;CellValue : Text[250];CellType : Option Number,Text,Date,Time;Bold : Boolean;Italic : Boolean;Underline : Boolean);
    begin
        TempExcelBuffer.INIT;
        TempExcelBuffer.VALIDATE("Row No.",RowNo);
        TempExcelBuffer.VALIDATE("Column No.",ColumnNo);
        TempExcelBuffer."Cell Value as Text" := CellValue;
        TempExcelBuffer."Cell Type" := CellType;
        TempExcelBuffer.Bold := Bold;
        TempExcelBuffer.Italic := Italic;
        TempExcelBuffer.Underline := Underline;
        TempExcelBuffer.INSERT;
    end;

    local procedure FillHeader(StartRowNo : Integer;LocalRecRef : RecordRef);
    var
        LocalFieldRef : FieldRef;
    begin
        RowNo := StartRowNo;
        ColNo := 1;
        EnterCell(RowNo,ColNo,LocalRecRef.CAPTION + ' ('+FORMAT(LocalRecRef.NUMBER)+')',TempExcelBuffer."Cell Type"::Text,true,true,true);
        ColNo := 2;
        EnterCell(RowNo,ColNo,'VALUE',TempExcelBuffer."Cell Type"::Text,true,true,true);
        ColNo := 3;
        EnterCell(RowNo,ColNo,'CLASS',TempExcelBuffer."Cell Type"::Text,true,true,true);
        ColNo := 4;
        EnterCell(RowNo,ColNo,'OPTION',TempExcelBuffer."Cell Type"::Text,true,true,true);
        ColNo := 1;
        for RowNo := (1 + StartRowNo) to (LocalRecRef.FIELDCOUNT + StartRowNo) do begin
          LocalFieldRef := LocalRecRef.FIELDINDEX(RowNo - StartRowNo);
          EnterCell(RowNo,ColNo,LocalFieldRef.CAPTION + ' ('+FORMAT(LocalFieldRef.NUMBER)+')',TempExcelBuffer."Cell Type"::Text,true,false,false);
        end;
    end;

    local procedure FillLine(StartRowNo : Integer;LocalRecRef : RecordRef) : Integer;
    var
        LocalFieldRef : FieldRef;
    begin
        for RowNo := (1 + StartRowNo) to (LocalRecRef.FIELDCOUNT + StartRowNo) do begin
          LocalFieldRef := LocalRecRef.FIELDINDEX(RowNo - StartRowNo);
          ColNo := 2;
          EnterCell(RowNo,ColNo,FORMAT(LocalFieldRef.VALUE),TempExcelBuffer."Cell Type"::Text,false,false,false);
          ColNo := 3;
          EnterCell(RowNo,ColNo,FORMAT(LocalFieldRef.CLASS),TempExcelBuffer."Cell Type"::Text,false,false,false);
          ColNo := 4;
          EnterCell(RowNo,ColNo,FORMAT(LocalFieldRef.OPTIONCAPTION),TempExcelBuffer."Cell Type"::Text,false,false,false);
        end;
        exit(RowNo);
    end;
}

