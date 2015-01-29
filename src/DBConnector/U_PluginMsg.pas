unit U_PluginMsg;

interface

uses
  Messages;

const
  WMUSER_DBCONNECTOR_PLUGIN_RESULTDIR = WM_USER+1021;

type
  TPluginMsgProtocol = class
  public
  class function BuildResultDir(s: string):PString;
  class function BuildResultDirMsg(sDir: string): Integer;
  class function GetResultDirByMsg(msg: TMessage): string;
  end;

implementation

class function TPluginMsgProtocol.BuildResultDir(s: string):PString;
var
  ps: PString;
begin
  New(ps);
  ps^ := s;
  Result := ps;
end;   

class function TPluginMsgProtocol.BuildResultDirMsg(sDir: string): Integer;
begin
  Result := Integer(BuildResultDir(sDir));
end;

class function TPluginMsgProtocol.GetResultDirByMsg(msg: TMessage): string;
var
  ps: PString;
begin
  ps := PString(msg.WParam);
  Result := ps^;
end;

end.
