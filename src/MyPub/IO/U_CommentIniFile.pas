unit U_CommentIniFile;

interface

uses
  Classes, SysUtils, StrUtils, IniFiles, Forms, Windows;

const
  COMMENT_HEAD = 'head';

type
  TCommentIniFile = class(TMemIniFile)
  private
    FSectionComments: TStrings;

  public
    constructor Create(const FileName: string); overload;
    constructor Create(const FileName: string; Encoding: TEncoding); overload;
    destructor Destroy; override;

    procedure UpdateFile; override;

    procedure AddComment(section: string; comment: String);

  end;

implementation

{ TCommentIniFile }

constructor TCommentIniFile.Create(const FileName: string);
begin
  inherited Create(FileName);
  FSectionComments := THashedStringList.Create;
end;

constructor TCommentIniFile.Create(const FileName: string; Encoding: TEncoding);
begin
  inherited Create(FileName, Encoding);
  FSectionComments := THashedStringList.Create;
end;

destructor TCommentIniFile.Destroy;
begin
  FSectionComments.Free;
  inherited;
end;

procedure TCommentIniFile.UpdateFile;
var
  List: TStringList;
  i, index: Integer;
  sSectionName, sComment: String;
begin
  List := TStringList.Create;
  try
    GetStrings(List);
    // head comment
    if List.Count > 0 then begin
      index := FSectionComments.IndexOfName(COMMENT_HEAD);
      if index >= 0 then begin
        sComment := '# ' + FSectionComments.ValueFromIndex[index];
        if List[0] <> sComment then
          List.Insert(0, sComment);
      end;
    end;

    // section comment
    for i := 0 to list.Count - 1 do begin
      if AnsiStartsStr('[', List[i]) and AnsiEndsStr(']', List[i]) then begin
        sSectionName := Copy(List[i], 2, Length(List[i]) - 2);
        index := FSectionComments.IndexOfName(sSectionName);
        if index >= 0 then begin
          sComment := '# ' + FSectionComments.ValueFromIndex[index];
          if (List.Count>i+1) and (List[i+1] <> sComment) then
            List.Insert(i+1, sComment);
        end;
      end;
    end;

    List.SaveToFile(FileName, Encoding);
  finally
    List.Free;
  end;
end;

procedure TCommentIniFile.AddComment(section, comment: String);
begin
  FSectionComments.Add(section + FSectionComments.NameValueSeparator + comment);
end;

end.
