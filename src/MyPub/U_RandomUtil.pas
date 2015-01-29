unit U_RandomUtil;

interface

type
  TRandomUtil = class

  public
    constructor Create;

    function GetRandomDouble(minValue, maxValue: Double): Double;
    function GetRandomInteger(minValue, maxValue: Integer): Integer;
  end;

implementation

uses
  Math;

{ TRandomUtils }

function TRandomUtil.GetRandomDouble(minValue, maxValue: Double): Double;
var
  r: Double;
begin
  r := Random;
  Result := (maxValue - minValue) * r + minValue;
end;

constructor TRandomUtil.Create;
begin
  Randomize;
end;

function TRandomUtil.GetRandomInteger(minValue, maxValue: Integer): Integer;
var
  r: Double;
begin
  r := Random;
  Result := Floor((maxValue - minValue + 1) * r) + minValue;
end;

end.
