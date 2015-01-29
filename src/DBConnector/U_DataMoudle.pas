unit U_DataMoudle;

interface

uses
  SysUtils, Classes, ImgList, Controls, acAlphaImageList;

type
  TDataModule1 = class(TDataModule)
    ilsTree: TsAlphaImageList;
    ilTree: TImageList;
    ilBtnIcons: TImageList;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DataModule1: TDataModule1;

implementation

{$R *.dfm}

end.
