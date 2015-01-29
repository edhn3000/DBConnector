{**
 * <p>Title: U_PubFunc.pas  </p>
 * <p>Copyright: Copyright (c) 2009-3-28</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: �������ڵķ������ϵ�Ԫ��ʹ��ԭ�����£�</p>
 * <p>Description: 1������������ؿ���ȡ�����ŵ�����Ǳ�������طŵ�commomfunc������util��Ԫ </p>
 * <p>Description: 2��ÿ������ע�����ĸ���Ԫʹ�ã����ŷָ������Ԫ </p>
 * <p>Description: 3�������������κδ��嵥Ԫ��������global��Ԫ��commonfunc������util��Ԫ </p>
 *}
unit U_PubFunc;

interface

uses
  U_fStrUtil;

type
  TPubFunc = class
  public
    {**
     * һ����� ������FieldNode��Text
     * ʹ�õ�Ԫ U_MAIN
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
