{**
 * <p>Title: U_PubFunc.pas  </p>
 * <p>Copyright: Copyright (c) 2009-3-28</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 本工程内的方法集合单元：使用原则如下：</p>
 * <p>Description: 1：将本工程相关可提取方法放到这里，非本工程相关放到commomfunc或其他util单元 </p>
 * <p>Description: 2：每个方法注明在哪个单元使用，逗号分隔多个单元 </p>
 * <p>Description: 3：不引用其他任何窗体单元，可引用global单元，commonfunc或其他util单元 </p>
 *}
unit U_PubFunc;

interface

uses
  U_fStrUtil;

type
  TPubFunc = class
  public
    {**
     * 一般情况 参数是FieldNode的Text
     * 使用单元 U_MAIN
     *}
    function GetFieldShortName(sFieldFullName: string): string;
  
  end;


var
  pubFunc: TPubFunc;
implementation

{ TPubFunc }

function TPubFunc.GetFieldShortName(sFieldFullName: string): string;
begin
  if Pos(' ', sFieldFullName) > 0 then
    Result := fStrUtil.SubStringBetween(sFieldFullName, '', ' ')
  else
    Result := sFieldFullName;
end;    

initialization
  pubFunc := TPubFunc.Create;
finalization
  pubFunc.Free;

end.
