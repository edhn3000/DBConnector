unit U_ExportUtil;

interface

uses
  Windows, Messages, Dialogs, Forms, ComCtrls, Controls, Classes, Types,
  ExtCtrls, Graphics, ADODB, StdCtrls, DBGrids, DB, Printers, OleServer,
  U_FieldInfo,
  U_fStrUtil;

type
  TItemLengths = array of Integer;

  { TExportUtil }
  TExportUtil = class
  public

    class function ExportToCSV(AQ: TDataSet; sFileName: string; sSheetName: string = ''): Boolean;
    class function ExportToExcel(AQ: TDataSet; sFileName: string; sSheetName: string = ''): Boolean;

    class procedure FillExportLvwList(ALvw: TListView; AStrs: TStrings;
        aiLengths: TItemLengths; JoinSeparator: string; cLineChar: Char = '-');
    class procedure FillExportQueryList(AQry: TDataSet; AStrs: TStrings;
        JoinSeparator: string; cLineChar: Char = '-');
  private
    class function BuildStringByChar(c: Char; nLength: Integer): string; static;
  public

  end;

implementation

uses
  SysUtils, jpeg, Variants, U_PlanarList, U_ExcelObject;

{ TExportUtil }

class function TExportUtil.BuildStringByChar(c: Char; nLength: Integer): string;
var
  s: string;
begin
  s := c;
  while Length(s) * 2 <= nLength do
    s := s + s;
  s := s + Copy(s, 1, nLength - Length(s));
  Result := s;
end;

class function TExportUtil.ExportToCSV(AQ: TDataSet; sFileName: string;
  sSheetName: string): Boolean;
var
  i: Integer;
  sLine, sCell: string;
  fileContent: TStrings;
begin
  try
    fileContent := TStringList.Create;
    AQ.DisableControls;
    AQ.First;
    // 标题
    sLine := '';
    for i := 0 to AQ.FieldCount -1 do
    begin
      if i = 0 then
        sLine := sLine + AQ.Fields[i].DisplayName
      else
        sLine := sLine + ',' + AQ.Fields[i].DisplayName;
    end;
    fileContent.Add(sLine);

    while not AQ.Eof do
    begin
      Application.ProcessMessages;
      sLine := '';
      for i := 0 to AQ.FieldCount - 1 do
      begin
        sCell := AQ.Fields[i].AsString;
        if Pos(',', sCell) > 0 then
          sCell := '"' + sCell + '"';

        if i = 0 then
          sLine := sLine + sCell
        else
          sLine := sLine + ',' + sCell;
      end;
      fileContent.Add(sLine);
      AQ.Next;
    end;

    AQ.First;
    AQ.EnableControls;
    fileContent.SaveToFile(sFileName);
    fileContent.Free;

    Result := True;
  except
    Result := False;
  end;
end;

class function TExportUtil.ExportToExcel(AQ: TDataSet; sFileName: string;
  sSheetName: string): Boolean;
var
  i, nRow: Integer;
  sDataType: string;
  slstPlanar: TPlanarStringList;
  slstItem: TStringList;
  fieldItem: TFieldItem;
  ExcelExp: TExcelApplication;
begin
  try
    ExcelExp := TExcelApplication.Create;
    ExcelExp.AddWorkSheet;
    if sSheetName <> '' then
      ExcelExp.ActiveSheet.Name := sSheetName;
    slstPlanar := TPlanarStringList.Create;
    try
      AQ.DisableControls;
      AQ.First;
      // 标题
      for i := 0 to AQ.FieldCount -1 do
      begin
        try
          fieldItem := TFieldItem.Create(AQ.Fields[i]);
          try
            sDataType := fieldItem.DataTypeStr;
//            ExcelExp.ActiveSheet.Sheet.Columns.Item(i+1).NumberFormatLocal := '@';
            ExcelExp.ActiveSheet.CellValue[1,i+1] := AQ.Fields[i].DisplayName
              +':'+sDataType;
          finally
            fieldItem.Free;
          end;
        except
          // 不能设置Range的ColumnWidth属性，会出异常
        end;
      end;

      nRow := 2;
      while not AQ.Eof do
      begin
        Application.ProcessMessages;
        slstItem := TStringList.Create;
        for i := 0 to AQ.FieldCount - 1 do
        begin
//          ExcelExp.ActiveSheet.Cell[nRow,i+1].Cell.NumberFormatLocal := '@';  // 文本形式
          ExcelExp.ActiveSheet.CellValue[nRow,i+1] := AQ.Fields[i].AsString;
          slstItem.Add(AQ.Fields[i].AsString);
        end;
        slstPlanar.AddItem(slstItem);
        AQ.Next;
        Inc(nRow);
      end;
      slstPlanar.FormatItemLengths;

      for i := 0 to AQ.FieldCount -1 do
      begin
        try
          ExcelExp.ActiveSheet.SetColumnWidth(i+1,
            slstPlanar.ItemLength[i]+1+Length(sDataType));
        except
          // 不能设置Range的ColumnWidth属性，会出异常
        end;
      end;

    finally
      slstPlanar.Free;
    end;

    AQ.First;
    AQ.EnableControls;
    ExcelExp.SaveAs(sFileName);
    ExcelExp.Free;

    Result := True;
  except
    Result := False;
  end;
end;

class procedure TExportUtil.FillExportLvwList(ALvw: TListView; AStrs: TStrings;
  aiLengths: TItemLengths; JoinSeparator: string; cLineChar: Char);
var
  sFormat, sALine, sItemText: string;
  i, j: Integer;
begin
  with ALvw do
  begin
    sALine := '';
    for i := 0 to Columns.Count - 1 do
    begin
      sFormat := '%-' + IntToStr(aiLengths[i]) + 's';    // 字段是字符串 左对齐
//      if i <> Columns.Count - 1 then
      if sALine <> '' then
        sALine := sALine + JoinSeparator + Format(sFormat,[Columns[i].Caption])
      else
        sALine := Format(sFormat,[Columns[i].Caption]);
    end;
    AStrs.Add(sALine);

    // 分隔线
    sALine := '';
    for i := Low(aiLengths) to High(aiLengths) do
    begin
      if sALine = '' then
        sALine := BuildStringByChar(cLineChar, aiLengths[i])
      else
        sALine := sALine + JoinSeparator + BuildStringByChar(cLineChar, aiLengths[i]);
    end;
    AStrs.Add(sALine);

    //字段内容
    for i := 0 to Items.Count - 1 do
    begin
      sALine := '';
      for j := Low(aiLengths) to High(aiLengths) do
      begin
        if j = 0 then
          sItemText := Items[i].Caption
        else
          sItemText := Items[i].SubItems[j-1];

        if fStrUtil.IsNumber( sItemText ) then
          sFormat := '%' + IntToStr(aiLengths[j]) + 's'       // 数字右对齐
        else
          sFormat := '%-' + IntToStr(aiLengths[j]) + 's';     // 字符串左对齐

//        if j <> High(aiLengths) then
        if sALine <> '' then
          sALine := sALine + JoinSeparator + Format(sFormat, [sItemText])
        else
          sALine := Format(sFormat, [sItemText]);
      end;
      AStrs.Add( sALine );
    end;
  end;
end;

class procedure TExportUtil.FillExportQueryList(AQry: TDataSet; AStrs: TStrings;
   JoinSeparator: string; cLineChar: Char);
var
  i: Integer;
  slstPlanar: TPlanarStringList;
  slstItem: TStringList;
begin
  with AQry do
  begin
    slstPlanar := TPlanarStringList.Create;
    try
      DisableControls;
      slstItem := TStringList.Create;
      for i := 0 to FieldCount - 1 do
      begin
        if i = FieldCount - 1 then
          slstItem.Add(Fields[i].FieldName)
        else
          slstItem.Add(Fields[i].FieldName + JoinSeparator);
      end;
      slstPlanar.AddItem(slstItem);

       //字段内容
      First;
      while not Eof do
      begin
        Application.ProcessMessages;
        slstItem := TStringList.Create;
        for i := 0 to FieldCount - 1 do
        begin
          if i = FieldCount - 1 then
          begin
            // blob字段 无法显示内容
            if Fields[i].DataType in [ftOraBlob, ftBlob] then
              slstItem.Add('Blob')
            else
              slstItem.Add(Fields[i].AsString)
          end
          else
          begin
            if Fields[i].DataType in [ftOraBlob, ftBlob] then
              slstItem.Add('Blob' + JoinSeparator)
            else
              slstItem.Add(Fields[i].AsString + JoinSeparator);
          end;
        end;
        slstPlanar.AddItem(slstItem);
        Next;
      end;
      for i := 0 to Fields.Count - 1 do
      begin
        if Fields[i].DataType in [ftOraBlob, ftOraClob, ftBlob, ftMemo] then
          slstPlanar.AddExcludeCol(i);
      end;
      slstPlanar.FormatItemLengths;
      AStrs.Add(slstPlanar.ItemStr[0]);

      // 分隔线
      slstItem := TStringList.Create;
      for i := 0 to Fields.Count - 1 do
      begin
        slstItem.Add(BuildStringByChar(cLineChar,
            slstPlanar.ItemLength[i]));
      end;
      slstPlanar.InsertPlanarItem(1, slstItem);

      for i := 1 to slstPlanar.Count - 1 do
      begin
        AStrs.Add(slstPlanar.ItemStr[i]);
      end;

      First;
      EnableControls;
    finally
      slstPlanar.Free;
    end;
  end;
end;

end.
