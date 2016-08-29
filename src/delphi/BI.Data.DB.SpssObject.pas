unit BI.Data.DB.SpssObject;

interface

uses
  System.Classes,SysUtils,Windows,DB,System.IniFiles,Spss.SpssIO,System.Variants,
  Contnrs, Common, CodeSiteLogging;

const
  SPSS_STRING=1024;
  Lang_SIMPLE_CHINESE= 'zh_CN.65001';
type
  TRevType=(revVName_Label,revConverString,revStringToEnum,RevItemValue,
  RevNewVar,RevSingleToMulti,RevFixValue);
  TSaveAsProgress=procedure(MaxValue,Current:Integer) of object;
  TSpssOpenState=(sOPenRead,sOPenWrite,SOpenAppend,sClose);
  TspssVarType=(sNum=0,sString=1,sEnum=2,sDateTime=3);
  TspssVarHandles=array of Double;


  TSpssFilter = class;
  TRevSpssValue = class;
  TSpssVariable = class;
  TRevObject = class;
  TspssEnumItem =class;

 TRevSpssVarItem = class(TCollectionItem)
  private
    FPrintType: Integer;
    FRevColumnWidth: Integer;
    FRevDecimal: Integer;
    FRevLabel: string;
    FRevType: TRevType;
    FRevVarName: string;
    FRevVarType: TspssVarType;
    FSourceVar: TSpssVariable;
    FTargetVar: TSpssVariable;
    procedure SetRevType(const Value: TRevType);
    procedure SetRevVarType(const Value: TspssVarType);
  public
    constructor Create(Collection: TCollection); override;
    property PrintType: Integer read FPrintType write FPrintType;
    property RevColumnWidth: Integer read FRevColumnWidth write FRevColumnWidth;
    property RevDecimal: Integer read FRevDecimal write FRevDecimal;
    property RevLabel: string read FRevLabel write FRevLabel;
    property RevType: TRevType read FRevType write SetRevType;
    property RevVarName: string read FRevVarName write FRevVarName;
    property RevVarType: TspssVarType read FRevVarType write SetRevVarType;
    property SourceVar: TSpssVariable read FSourceVar write FSourceVar;
    property TargetVar: TSpssVariable read FTargetVar write FTargetVar;
  end;
  TRevSpssVarItemX = class(TRevSpssVarItem)
  private
    FHash: TObjectHash;
    FItems: TCollection;
    function GetItemCount: Integer;
    function GetValueItem(Index: Integer): TRevSpssValue;
    function GetValueItemByKey(Key:string): TRevSpssValue;
  public
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;
    function AddRevValue(const Key: string; const Value: Variant): TRevSpssValue;
        overload;
    function AddRevValue(const Key, RLabel: string; const Value: Variant):
        TRevSpssValue; overload;
    function ConvertValue(const Value: Variant): Variant;
    function GetRevValueByKey(const Key: string): Variant;
    procedure ChangeValue(const Key: string; const NewValue: Integer);
    function LabelExists(const TextLabel: string): Boolean;
    function ValueExits(const Key: string): Boolean;
    property ItemCount: Integer read GetItemCount;
    property ValueItem[Index: Integer]: TRevSpssValue read GetValueItem;
    property ValueItemByKey[Key:string]: TRevSpssValue read GetValueItemByKey;
  published
  end;
  SpssException = class(Exception)
  private
    FErrNum: Integer;
  public
    property ErrNum: Integer read FErrNum write FErrNum;
  end;
  TSpssEnumItem = class(TCollectionItem)
  private
    FEnumLabel: string;
    FParentVar: TSpssVariable;
    FValue: Integer;
    procedure SetEnumLabel(const Value: string);
  public
    constructor Create(Collection: TCollection); override;
    property EnumLabel: string read FEnumLabel write SetEnumLabel;
    property ParentVar: TSpssVariable read FParentVar;
    property Value: Integer read FValue write FValue;
  end;

  TSpssEnums = class(TCollection)
  private
    FHash: TStringHash;
    function GetEnumItem(Index: Integer): TSpssEnumItem;
  public
    constructor Create(ItemClass: TCollectionItemClass);
    destructor Destroy; override;
    procedure AddHash(Key: string; IndexValue: Integer);
    function EnumValueByName(EnumName: string): Integer;
    function HashExists(Key: string): Boolean;
    property EnumItem[Index: Integer]: TSpssEnumItem read GetEnumItem;
  end;
 TSpssVariable = class(TPersistent)
  private
    FColumnWidth: Integer;
    FDecimal: Integer;
    FFileHandle: Integer;
    FFilter: TSpssFilter;
    FList: TList;
    FPrintType: Integer;
    FRevItemList: TList;
    FVarLabel: string;
    FspssType: TspssvarType;
    FVarH: Double;
    FVarName: AnsiString;
    FVarType: TspssVarType;
    procedure AddRev(const RevItem: TRevSpssVarItemX);
    function GetEnumByValue(Value: Integer): TSpssEnumItem;
    function GetEnumCount: Integer;
    function GetEnumItem(Index: Integer): TSpssEnumItem;
    function GetFirstRevItem: TRevSpssVarItemX;
    function GetRevItem(Index: Integer): TRevSpssVarItemX;
    function GetRevItemCount: Integer;
    function GetRevVarLabel: string;
    function GetstrValue: string;
    function GetValue: Variant;
    function GetRevVarName: AnsiString;
    function GetVarRevised: Boolean;
    procedure SetColumnWidth(const Value: Integer);
    procedure SetFileHandle(const Value: Integer);
    procedure SetRevVarLabel(const Value: string);
    procedure SetVarLabel(const Value: string);
    procedure SetspssType(const Value: TspssvarType);
    procedure SetValue(const Value: Variant);
    procedure SetVarH(const Value: Double);
  protected
    procedure AssignTo(Dest: TPersistent); override;
  public
    constructor Create;
    destructor Destroy; override;
    function AddEnumItem(Lab: string; VAlue: Variant): TSpssEnumItem;
    procedure AddHash(Key: string; Enum: TSpssEnumItem);
    procedure ConverStringType(Size: Integer);
    procedure ConvertBooleanType;
    procedure ConvertFloatType;
    procedure ConvertsmallintType;
    procedure ConvertIntegerType;
    procedure ConvertLongIntType;
    procedure ConverEnumType;
    function IsBoolean: Boolean;
    function KeyExists(Key: string): boolean;
    procedure ConverttDateTimeType;
    procedure DeleteRevItem(RevItem: TRevSpssVarItemX);
    procedure EnumStringVar;
    function GetMaxEnumValue: int64;
    function ValueIsNull: Boolean;
    property Decimal: Integer read FDecimal write FDecimal;
    property EnumByValue[Value: Integer]: TSpssEnumItem read GetEnumByValue;
    property EnumCount: Integer read GetEnumCount;
    property EnumItem[Index: Integer]: TSpssEnumItem read GetEnumItem;
    property FileHandle: Integer read FFileHandle write SetFileHandle;
    property Filter: TSpssFilter read FFilter write FFilter;
    property FirstRevItem: TRevSpssVarItemX read GetFirstRevItem;
    property PrintType: Integer read FPrintType write FPrintType;
    property RevItem[Index: Integer]: TRevSpssVarItemX read GetRevItem;
    property RevItemCount: Integer read GetRevItemCount;
    property RevVarLabel: string read GetRevVarLabel write SetRevVarLabel;
    property RevVarName: AnsiString read GetRevVarName;
    property VarRevised: Boolean read GetVarRevised;
    property strValue: string read GetstrValue;
    property Value: Variant read GetValue write SetValue;
    property VarLabel: string read FVarLabel write SetVarLabel;
    property VarH: Double read FVarH write SetVarH;
  published
    property ColumnWidth: Integer read FColumnWidth write SetColumnWidth;
    property spssType: TspssvarType read FspssType write SetspssType;
    property VarName: AnsiString read FVarName write FVarName;
    property VarType: TspssVarType read FVarType write FVarType;
  end;

  TSpssFilter = class(TComponent)
  private
    FLang: ansistring;
    FRev: TRevObject;
    class function GetEncoding: ShortInt; static;
    function GetSpssVarItem(Index: Integer): TSpssVariable;
    function GetSpssVarByName(Name:string): TSpssVariable;
    function GetVarCount: Integer;
    class procedure SetEncoding(const Value: ShortInt); static;
  protected
    FEnumItems: TSpssEnums;
    FHandle: Integer;
    FHash: TObjectHash;
    FSpssState: TSpssOpenState;
    FValueHash: TStringHash;
    FVars: TObjectList;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function AddVariable(VarName: AnsiString; varType: TspssVarType; vHandle:
        Double): TSpssVariable;
    procedure SetLocal;
    function VarExists(const VarName: string): Boolean;
    class property Encoding: ShortInt read GetEncoding write SetEncoding;
    property Lang: ansistring read FLang write FLang;
    property Rev: TRevObject read FRev;
    property SpssVarItem[Index: Integer]: TSpssVariable read GetSpssVarItem;
    property SpssVarByName[Name:string]: TSpssVariable read GetSpssVarByName;
    property VarCount: Integer read GetVarCount;
  published
    procedure Close; virtual; abstract;
    procedure OpenSpssFile(FileName: AnsiString); virtual; abstract;
  end;

  TSpssReader = class(TSpssFilter)
  private
    FActive: Boolean;
    FFileName: string;
    FOnSaveAsProgress: TSaveAsProgress;
    FSpssCaseCount: Integer;
    function AddVariableHeader(VarName: AnsiString; varType: TspssVarType; vHandle:
        Double): TSpssVariable;
    function GetSpssCaseCount: Integer;
  protected
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ClearHeaders;
    procedure CloseNOClearHeader;
    procedure EnumStringVar;
    procedure GetVariable(VariableName: AnsiString);
    procedure GetVList;
    procedure InitHeader;
    function ReadRecord: Boolean;
    procedure SaveAs(const NewSpssFileName: string);
    procedure RevSpss(const NewSpssFileName: string);
    procedure SpssRecordSeek(const CaseNo: Integer);
    function VariableByName(VName: string): TSpssVariable;
    property Active: Boolean read FActive;
    property FileName: string read FFileName;
    property SpssCaseCount: Integer read GetSpssCaseCount;
    property OnSaveAsProgress: TSaveAsProgress read FOnSaveAsProgress write
        FOnSaveAsProgress;
  published
    procedure Close; override;
    procedure GetValueLabels;
    procedure OpenSpssFile(FileName: AnsiString); override;
  end;

  TSVarSelector = class(TComponent)
  private
    FList: TList;
    function GetCount: Integer;
    function GetVarItem(Index: Integer): TSpssVariable;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function AddVar(SpssVar: TSpssVariable): Integer;
    procedure Clear;
    function Exists(SVar: TSpssVariable): Boolean;
    procedure RemoveVar(Spssvar: TSpssVariable);
    property Count: Integer read GetCount;
    property VarItem[Index: Integer]: TSpssVariable read GetVarItem;
  published
  end;

  TSpssWriter = class(TSpssFilter)
  private
    FOpend: Boolean;
    procedure CreateVarHeader(const SpssVar: TSpssVariable);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ChangeVarName(const V: TSpssVariable; const NewVarName: string);
    procedure Close; override;
    procedure CommitCase;
    procedure CreateHeader;
    procedure NewCase;
    function NewVariable(const VarName, VarLabel: string; const VarType:
        TspssVarType): TSpssVariable;
    procedure OpenSpssFile(FileName: AnsiString); override;
  published
  end;

  TRevObject = class(TComponent)
  private
    FParentFilter: TSpssFilter;
    FValueHash: TObjectHash;
    FVarRevHash: TObjectHash;
    FVarRevs: TCollection;
    function GetRevVar(ID: Integer): TRevSpssVarItemX;
    function GetRevVarByIndex(Index: Integer): TRevSpssVarItemX;
    function GetVarCount: Integer;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure ConvertEnumToString(const SVar: TSpssVariable);
    procedure ConvertStringToEnum(const SVar: TSpssVariable);
    procedure ReNameAndLabel(const SpssVar: TSpssVariable; const NewVarName,
        NewLabel: string); overload;
    procedure RevokeRev(SpssVar: TSpssVariable);
    procedure AddOpenVarToBoolanVar(const SourceVar: TSpssVariable; const
        TextLabel, VarPrefix: string; const ValueNum: Integer);
    procedure AddStringValueFix(const SVar: TSpssVariable; const NewValue,
        OldValue: string);
    procedure ClearAll;
    procedure ConvertSingleVarToMultiVar(const SVar: TSpssVariable; const
        Var_Prefix: string);
    function GetNewNodeName(NodeName: string): string;
    function RevVarExits(const VarName: string): Boolean; overload;
    function RevVarExits(const SpssVar: TSpssVariable): Boolean; overload;
    property RevVar[ID: Integer]: TRevSpssVarItemX read GetRevVar;
    property RevVarByIndex[Index: Integer]: TRevSpssVarItemX read GetRevVarByIndex;
    property VarCount: Integer read GetVarCount;
  published
    function GetRevSpssVar(SpssVar: TSpssVariable): TRevSpssVarItemX;
    function RevVaByName(const VarName: string): TRevSpssVarItemX;
    function RevVarByValue(const TextLabel: string): TRevSpssVarItemX;
    property ParentFilter: TSpssFilter read FParentFilter;
  end;
  TRevSpssValue = class(TCollectionItem)
  private
    FObj: TSpssEnumItem;
    FRevLabel: string;
    FRevName: string;
    FRevValue: Variant;
    FTargetVar: TSpssVariable;
  public
    property Obj: TSpssEnumItem read FObj write FObj;
    property RevLabel: string read FRevLabel write FRevLabel;
    property RevName: string read FRevName write FRevName;
    property RevValue: Variant read FRevValue write FRevValue;
    property TargetVar: TSpssVariable read FTargetVar write FTargetVar;
  end;
















procedure CheckError(Err: Integer);

implementation




procedure CheckError(Err: Integer);
var
  SpssE:SpssException;
begin
  if Err<>SPSS_OK then
  begin
    case Err of
      SPSS_FILE_OERROR:
      begin
        SpssE  := SpssException.Create('Spss Open File Error!,File already open or file is not exits!');
        SpssE.ErrNum :=SPSS_FILE_OERROR;
        raise SpssE;
      end;
      SPSS_INVALID_FILE:
      begin
        SpssE := SpssException.Create('Spss Open File Error!,File already open or file is not exits!');
        SpssE.ErrNum :=SPSS_INVALID_HANDLE;
        raise SpssE;
      end;
      SPSS_BUFFER_SHORT:
      begin
        SpssE := SpssException.Create('Spss Buffer is too short!');
        spsse.ErrNum :=SPSS_BUFFER_SHORT;
        raise SpssE;
      end;
      SPSS_INVALID_HANDLE:
      begin
        SpssE := SpssException.Create('IBM SPSS Statistics files are open');
        spsse.ErrNum :=SPSS_INVALID_HANDLE;
        raise SpssE;
      end;
      SPSS_FILES_OPEN:
      begin
        SpssE := SpssException.Create('The file handle is not valid!');
        spsse.ErrNum :=SPSS_INVALID_HANDLE;
        raise SpssE;
      end ;
      SPSS_INVALID_VARNAME:
      begin
        SpssE :=SpssException.Create('One or more variable names in list are invalid');
        SpssE.ErrNum :=SPSS_INVALID_VARNAME;
        raise SpssE;
      end;
      SPSS_VAR_NOTFOUND :
      begin
        SpssE :=SpssException.Create('One or more variables in list were not found in dictionary');
        SpssE.ErrNum :=SPSS_VAR_NOTFOUND;
        raise SpssE;
      end;
      SPSS_INVALID_PRFOR:
      begin
        SpssE :=SpssException.Create('The print format specification is invalid or is incompatible with the variable type');
        SpssE.ErrNum :=SPSS_INVALID_PRFOR;
        raise SpssE;
      end;
      SPSS_INVALID_CASE:
      begin
        SpssE :=SpssException.Create('The given case weight variable is invalid.');
        SpssE.ErrNum :=SPSS_INVALID_CASE;
        raise SpssE;
      end
    else
      SpssE := SpssException.Create('Spss Error_'+inttostr(Err));
      SpssE.ErrNum :=Err;
      raise SpssE;

    end;
  end;
end;

constructor TSpssReader.Create(AOwner: TComponent);
begin
  inherited ;
  FHandle :=-1;
  FSpssState :=sOPenRead;
  FFileName :='';
end;

destructor TSpssReader.Destroy;
begin
  Close;
  inherited;

end;

function TSpssReader.AddVariableHeader(VarName: AnsiString; varType:
    TspssVarType; vHandle: Double): TSpssVariable;
var
  v:TSpssVariable;
 Index :integer;
begin
  if Assigned(FHash.ValueOf(string(VarName))) then
    raise Exception.Create('Variable already Exists!');
  v :=TSpssVariable.Create;

  v.VarName := AnsiString(UpperCase(string(VarName))) ;
  v.VarType :=varType;
  FVars.Add(V);
  FHash.Add(UpperCase(v.VarName),V);
  result :=v;

end;

procedure TSpssReader.ClearHeaders;
begin
  FHash.Clear;
  FVars.Clear;


end;

procedure TSpssReader.Close;
var
  ErrH:Integer;
begin

  if FActive then
  begin
    ErrH :=spssCloseRead(FHandle);
    CheckError(ErrH);
    FSpssState :=sClose;
    ClearHeaders;
    FActive :=false;
  end;
end;

procedure TSpssReader.CloseNOClearHeader;
begin
  if FActive then
  begin
    spssCloseRead(FHandle);
    FActive :=false;
  end;
end;

procedure TSpssReader.EnumStringVar;
var
  I: Integer;
  V:TSpssVariable;
  str:string;
  enum:TSpssEnumItem;
begin

  while self.ReadRecord do
  begin
    for I := 0 to self.VarCount-1 do
    begin
      V :=SpssVarItem[I];
      if v.VarType =sString then
      begin
        str :=v.strValue;
        if not V.KeyExists(str) and (str<>'') then
        begin
          enum :=v.AddEnumItem(str,V.EnumCount+1);
          v.AddHash(str,enum);
        end;
      end;

    end;
  end;


  Self.SpssRecordSeek(0);
end;

function TSpssReader.GetSpssCaseCount: Integer;
var
  ErrH:integer;
begin
  if FSpssCaseCount<0 then
  begin
    ErrH :=spssGetNumberofCases(FHandle,FSpssCaseCount);
    CheckError(ErrH);
    result :=FSpssCaseCount;
  end else
  begin
    Result :=FSpssCaseCount;
  end;
end;

procedure TSpssReader.GetValueLabels;
var
  I,J:integer;
  V:TSpssVariable;
  pValues:PTValued;
  PLabels:PTValuec;
  numLabels:integer;
  ErrH:integer;
  Lab:string;
  Value:double;
begin
  for I := 0 to varcount-1 do
  begin
    V :=SpssVarItem[I];
    ErrH :=spssGetVarNValueLabels(FHandle,PAnsiChar(v.VarName),PValues,PLabels,numLabels);
    CheckError(ErrH);
    if numLabels<>0 then
    begin
      for J := 0 to numLabels-1 do
      begin
        value :=PValues^[J];
        Lab :=string(PLabels^[J]);
        if Value>=0 then
        begin
          v.AddEnumItem(Lab,Value);
        end;
      end;
      spssFreeVarNValueLabels(PValues,PLabels,numLabels);
    end;
  end;

end;

procedure TSpssReader.GetVariable(VariableName: AnsiString);
var
  Lab:array[0..512] of AnsiChar;
  PLabel:PAnsiChar;
  LenLabel:Integer;
  ErrH:integer;
begin
  PLabel :=@lab[0];
  ErrH :=spssGetVarLabelLong(FHandle,PAnsiChar(VariableName),PLabel,512,LenLabel);
  CheckError(ErrH);
end;

procedure TSpssReader.GetVList;
begin
  // TODO -cMM: TSpssReader.GetVList default body inserted
end;

procedure TSpssReader.InitHeader;
var
  PvName:PTValuec;
  PvType:PTValuei;
  I: Integer;
  pName:PAnsichar;
  PLabel:PAnsiChar;
  Lab:array[0..512] of AnsiChar;
  strName:string;
  VType:integer;
  vHandle:double;
  ErrH:integer;
  V:TSpssVariable;
  len:integer;
  varName:AnsiString;
  numLabels:integer;
  VCount:integer;
  varHandle:Double;
  ColumnWidth:integer;
  PrintType:Integer;
  PrintDecimal:integer;

  ItemValues:TValued;
  ItemLabels:TValuec;
  PitemValues:PTValued;
  PItemLabels:PTValuec;
  EnumCount:integer;
  J: Integer;
  ItemValue:integer;
  ItemLabel:string;
  NumVars:Integer;
  VarType :TspssVarType;
  VarLabel:string;
begin



  PLabel :=@Lab[0];
  ErrH:=spssGetNumberofVariables(FHandle,NumVars);
  CheckError(ErrH);

  ErrH :=spssGetVarNames(FHandle,VCount,PvName,PvType);
  CheckError(ErrH);

  for I := 0 to VCount-1 do
  begin
    PName :=Pvname^[I];
    Vtype :=PVtype^[I];

    strName :=string(AnsiString(pName));
    if Vtype=0 then
    begin
      VarType:=snum;
    end else
    begin
      VarType :=sstring;
    end;

    VType :=PVType^[I];
    errH:=spssGetVarHandle(FHandle,PVName^[I],vHandle);
    CheckError(ErrH);

    ErrH :=spssGetVarLabelLong(FHandle,PName,PLabel,Length(Lab),len);
    if ErrH<>SPSS_NO_LABEL then
    begin
      CheckError(ErrH);
      varLabel:=AnsiString(PLabel);
    end else
    begin
      VarLabel :='';
    end;
    ErrH:=spssGetVarHandle(FHandle,PName,varHandle);
    CheckError(ErrH);

    ErrH :=spssGetVarPrintFormat(FHandle,pName,PrintType,PrintDecimal,ColumnWidth);
    //ErrH :=spssGetVarColumnWidth(FHandle,PName,ColumnWidth);
    CheckError(ErrH);
    if PrintType  in [SPSS_FMT_DATE,SPSS_FMT_TIME,SPSS_FMT_DATE_TIME] then
    begin
      Continue;
    end;
    //
    V :=AddVariableHeader(strName,VarType,vHandle);
    v.VarLabel :=VarLabel;
    v.Filter :=self;
    v.VarH :=varHandle;
    v.PrintType :=PrintType;
    v.FileHandle :=self.FHandle;
    v.Decimal :=PrintDecimal;
    v.ColumnWidth :=ColumnWidth;

    PitemValues :=@ItemValues[0];
    PItemLabels :=@ItemLabels[0];
    ErrH:=spssGetVarNValueLabels(FHandle,Pname,PitemValues,PItemLabels,EnumCount);
    if ErrH=0 then
    begin
      //Enum Type
      if EnumCount >0 then
      begin
        v.VarType :=sEnum;
      //重新赋值类型
        for J := 0 to EnumCount-1 do
        begin
          ItemLabel :=PItemLabels^[J];
          ItemValue :=Trunc(PitemValues^[J]);
          v.AddEnumItem(ItemLabel,ItemValue);
        end;
      end;




    end else
    if ErrH <>-8 then
    begin
     // CheckError(ErrH);
    end;
  end;
  spssfreevarnames(PvName,PvType,VCount);


end;

procedure TSpssReader.OpenSpssFile(FileName: AnsiString);
var
  ErrH:integer;
begin
  self.FRev.ClearAll;
  if FSpssState<>sClose then
  begin
    close;
  end;
  FHandle :=-1;
  FSpssCaseCount :=-1;
  FFileName :=FileName;
  ErrH:=spssOpenRead(PAnsiChar(FileName),FHandle);
  CheckError(ErrH);
  FSpssState :=sOPenRead;
  Self.ClearHeaders;
  FActive :=true;

end;

function TSpssReader.ReadRecord: Boolean;
var
  ErrH:integer;
begin

  ErrH:=spssReadCaseRecord(FHandle);
  if ErrH<>SPSS_OK then
    result :=False
  else
    result :=true;
end;

procedure TSpssReader.SaveAs(const NewSpssFileName: string);
var
  SW:TSpssWriter;
  I:integer;
  varS,VarD:TSpssVariable;
  cnt:integer;
begin
  SW :=TSpssWriter.Create(nil);
  try
    SW.OpenSpssFile(NewSpssFileName);
    for I := 0 to self.VarCount-1 do
    begin
      varS :=self.SpssVarItem[I];
      if vars.EnumCount>0 then
      begin
        VarD :=SW.NewVariable(vars.VarName,varS.VarLabel,sEnum);
      end else
      begin
        VarD :=SW.NewVariable(vars.VarName,varS.VarLabel,varS.VarType);
      end;
      vard.Assign(varS);

    end;
    SW.CreateHeader;
    //Self.SpssRecordSeek(0);
    cnt :=0;
    if Assigned(FOnSaveAsProgress) then
    begin
      FOnSaveAsProgress( self.SpssCaseCount ,cnt);
    end;

    while self.ReadRecord do
    begin
      SW.NewCase;
      for I := 0 to self.VarCount-1 do
      begin
        varS :=self.SpssVarItem[I];
        VarD :=sw.SpssVarItem[I];
        if not vars.ValueIsNull then
        begin

          VarD.Value :=vars.Value;
        end;
      end;
      SW.CommitCase;
      Inc(cnt);
      if cnt mod 500 =0 then
      begin
        if Assigned(FOnSaveAsProgress) then
        begin
          FOnSaveAsProgress(self.SpssCaseCount ,cnt);
        end;
      end;
    end;
    SW.Close;
  finally
    FreeAndNIl(SW);
  end;

end;

procedure TSpssReader.RevSpss(const NewSpssFileName: string);
var
  SW:TSpssWriter;
  I,J:integer;
  varS,VarD:TSpssVariable;
  cnt:integer;
  Item:TRevSpssVarItemX;
  Val:Variant;
  ValItem:TRevSpssValue;

begin
  SW :=TSpssWriter.Create(nil);
  try
    SW.OpenSpssFile(NewSpssFileName);
    for I := 0 to self.VarCount-1 do
    begin
      varS :=self.SpssVarItem[I];
      Item :=  Rev.RevVaByName(varS.VarName);
      if not Assigned(Item) then
      begin
        VarD :=SW.NewVariable(vars.VarName,varS.VarLabel,varS.VarType);
        varD.ColumnWidth :=VarS.ColumnWidth;
        VarD.Decimal :=VarS.Decimal;
        vard.PrintType :=vars.PrintType;
        if vars.EnumCount>0 then
        begin
          for J := 0 to vars.EnumCount-1 do
          begin
            VarD.AddEnumItem(varS.EnumItem[J].EnumLabel,vars.EnumItem[J].Value);
          end;
        end;
      end else
      if item.RevType <>RevSingleToMulti then
      begin
        VarD :=SW.NewVariable(Item.RevVarName,Item.RevLabel,Item.RevVarType);
        VarD.ColumnWidth :=item.RevColumnWidth;
        VarD.Decimal :=Item.RevDecimal;
        VarD.PrintType :=Item.PrintType;
        Item.TargetVar :=VarD;
        if Item.RevType in [revStringToEnum,revVName_Label]   then
        begin
          for J := 0 to Item.ItemCount-1 do
          begin

            VarD.AddEnumItem(Item.ValueItem[J].RevLabel,Item.ValueItem[J].RevValue);

          end;

        end;
      end;


    end;
    for I := 0 to self.Rev.VarCount-1 do
    begin
      Item :=self.Rev.RevVarByIndex[I];
      if item.RevType =RevSingleToMulti then
      begin
        for J := 0 to item.ItemCount-1 do
        begin
          ValItem:=item.ValueItem[J];
          if not SW.VarExists(valitem.RevName) then
          begin
            VarD :=SW.NewVariable(valitem.RevName,ValItem.RevLabel,sEnum);
            VarD.AddEnumItem('Yes',1);
            vard.AddEnumItem('No',0);
            ValItem.TargetVar :=VarD;
          end else
          begin
            varD :=SW.SpssVarByName[ValItem.RevName];
            ValItem.TargetVar :=VarD;
          end;
        end;
      end;
    end;


    SW.CreateHeader;
    Self.SpssRecordSeek(0);
    cnt :=0;
    if Assigned(FOnSaveAsProgress) then
    begin
      FOnSaveAsProgress( self.SpssCaseCount ,cnt);
    end;
    Self.SpssRecordSeek(0);
    while self.ReadRecord do
    begin
      SW.NewCase;
      for I := 0 to self.VarCount-1 do
      begin
        varS :=self.SpssVarItem[I];

        Val :=varS.Value;
        if not  varS.VarRevised then
        begin
          Val :=varS.Value;
          if not (VarType(val) in [varEmpty,varNull]) then
          begin
            VarD :=SW.SpssVarByName[varS.VarName];
            VarD.Value :=Val;
          end;

        end;

      end;
      for I := 0 to Self.Rev.VarCount-1 do
      begin
        Item :=rev.RevVarByIndex[I];
        if item.RevType =RevSingleToMulti then
        begin

          ValItem :=item.ValueItemByKey[IntToStr(Item.SourceVar.Value )];
          if Assigned(ValItem) then
          begin
            ValItem.TargetVar.Value :=1;
          end;
        end else
        if not item.SourceVar.ValueIsNull then
        begin
          item.TargetVar.Value :=item.ConvertValue(item.SourceVar.Value);

        end;
      end;

      SW.CommitCase;
      Inc(cnt);
      if cnt mod 500 =0 then
      begin
        if Assigned(FOnSaveAsProgress) then
        begin
          FOnSaveAsProgress(self.SpssCaseCount ,cnt);
        end;
      end;
    end;
    SW.Close;
  finally
    FreeAndNIl(SW);
  end;

end;

procedure TSpssReader.SpssRecordSeek(const CaseNo: Integer);
var
  ErrH:Integer;
begin
  ErrH:=spssSeekNextCase(FHandle,CaseNo);
  CheckError(ErrH);
end;

function TSpssReader.VariableByName(VName: string): TSpssVariable;

begin
  result  :=FHash.ValueOf(UpperCase(VName)) as TSpssVariable


end;

constructor TSpssEnumItem.Create(Collection: TCollection);
begin
  inherited Create(Collection);

end;

procedure TSpssEnumItem.SetEnumLabel(const Value: string);
begin
  FEnumLabel := Value;
end;

constructor TSpssVariable.Create;
begin
  inherited ;

  FList :=TList.Create;
  FRevItemList :=TList.Create;

end;

destructor TSpssVariable.Destroy;
begin
  FreeandNil(FRevItemList);
  FreeAndNil(FList);
  inherited Destroy;
end;

function TSpssVariable.AddEnumItem(Lab: string; VAlue: Variant): TSpssEnumItem;
begin
  Result :=FFilter.FEnumitems.Add as TSpssEnumItem;
  result.EnumLabel:=lab;
  result.Value :=Value;
  result.FParentVar :=Self;
  FFilter.FValueHash.Add(self.VarName+'|'+IntToStr(Result.Value),Result.Index);
  FLIst.Add(Result);
end;

procedure TSpssVariable.AddHash(Key: string; Enum: TSpssEnumItem);
begin
  FFilter.FEnumItems.AddHash(FVarName +'|'+ Key,Enum.Value);
end;

procedure TSpssVariable.AddRev(const RevItem: TRevSpssVarItemX);
begin
  if FList.IndexOf(RevItem)<0 then
  begin
    FRevItemList.Add(RevItem);
  end;
end;

procedure TSpssVariable.AssignTo(Dest: TPersistent);
var
  DES:TSpssVariable;
  I: Integer;
  item:TSpssEnumItem;
begin
  if not (Dest is TSpssVariable) then  raise  Exception.Create('  Dest is not spsss variable');
  DES :=dest as TSpssVariable;
  DES.VarName :=self.VarName;
  DES.VarType :=self.VarType;
  des.VarLabel :=self.VarLabel;
  des.PrintType :=Self.PrintType;
  DES.ColumnWidth :=self.ColumnWidth;
  des.Decimal :=self.Decimal;
  for I := 0 to self.EnumCount-1 do
  begin
    item :=self.EnumItem[I];
    DES.AddEnumItem(item.EnumLabel,item.Value);
  end;
end;

procedure TSpssVariable.ConverStringType(Size: Integer);
begin
  FVarType :=sString;
  FColumnWidth :=Size;
  FDecimal :=0;
end;

procedure TSpssVariable.ConvertBooleanType;
begin
  FVarType :=sEnum;
  FColumnWidth :=1;
  FDecimal :=0;
  FPrintType :=SPSS_FMT_F;
end;

procedure TSpssVariable.ConvertFloatType;
begin
  FVarType :=sNum;
  FColumnWidth :=10;
  FDecimal :=4;
  FPrintType :=SPSS_FMT_F;
end;

procedure TSpssVariable.ConvertsmallintType;
begin
  FVarType :=sNum;
  FColumnWidth :=5;
  FDecimal :=0;
  FPrintType :=SPSS_FMT_F;
end;

procedure TSpssVariable.ConvertIntegerType;
begin
  FVarType :=sNum;
  FColumnWidth :=10;
  FDecimal :=0;
  FPrintType :=SPSS_FMT_F;
end;

procedure TSpssVariable.ConvertLongIntType;
begin
  FVarType :=sNum;
  FColumnWidth :=20;
  FDecimal :=0;
  FPrintType :=SPSS_FMT_F;
end;

procedure TSpssVariable.ConverEnumType;
begin
  FVarType :=sEnum;
  FColumnWidth :=5;
  FDecimal :=0;
  FPrintType :=SPSS_FMT_F;
end;

function TSpssVariable.GetEnumByValue(Value: Integer): TSpssEnumItem;
var
  Index:integer;
begin
  Index :=FFilter.FValueHash.ValueOf(FVarName +'|'+IntToStr(Value));
  if Index<0 then
    result :=nil
  else
    Result :=FFilter.FEnumItems.EnumItem[Index];
end;

function TSpssVariable.GetEnumCount: Integer;
begin
  Result :=FList.Count
end;

function TSpssVariable.GetEnumItem(Index: Integer): TSpssEnumItem;
begin

  Result :=Tobject( FList.Items[Index] ) as TSpssEnumItem ;
end;

function TSpssVariable.GetstrValue: string;
var
  ErrH:integer;
  P:PAnsiChar;
  ch:array[0..2048] of AnsiChar;
  Value:Double;
  PI64:PInt64;
  m1,m2,m3:double;
  str:string;
begin

  if VarType in [Snum,Senum] then
  begin
    ErrH :=spssGetValueNumeric(FileHandle,FVarH,Value);
    CheckError(ErrH);
    if Value<0 then
    begin
      result :='';
    end else
    begin
      result :=FloatToStr(Value);
    end;
  end else
  begin
    //SetLength(ch,FColumnWidth);
    //FillChar(ch[0],FColumnWidth,0);
    P :=@ch[0];
    ErrH :=spssGetValueChar(FileHandle,FVarH,P,2047);
    CheckError(ErrH);
    str := Trim(AnsiString(P));
    if str<>'' then
    begin
      Result :=str;
    end else
    begin
      result :='';
    end;
  end;
end;

function TSpssVariable.GetValue: Variant;
const
  BUFFER_SIZE =8192;
var
  P:PAnsiChar;
  Value:double;
  ErrH:integer;
  strValue:RawByteString;
  Buffer:array[0..BUFFER_SIZE-1] of AnsiChar;
  Enum:TSpssEnumItem;
  IntValue:integer;
  Index:integer;
  int64Val:int64 absolute Value;
begin
  if VarType in [sNum,sEnum] then
  begin
    ErrH :=spssGetValueNumeric(FileHandle,VarH,Value);
    CheckError(ErrH);
    if EnumCount >0 then
    begin

      if Value>0 then
      begin
        IntValue :=Trunc(Value);
        Enum :=getEnumByValue(intValue);
        if Assigned(Enum) then
          result :=enum.Value
        else
          Result :=Null;
           //TVarData(result).VType :=varNull;
      end else
      begin
        //TVarData(Result).VType :=varNull;
        result :=Null;
      end;
    end else
    begin
      if FDecimal >0 then
      begin
        if Value<-999999999 then
        begin
          Result :=Null;
        end else
        begin
        result :=Value;
        end;
      end
      else
      begin
        if Value>=0 then
        begin
          result :=trunc(Value);
        end else
        begin
           result :=Null;
        end;
      end;
    end;
  end else
  begin
      FillChar(Buffer[0],BUFFER_SIZE,$0);
      P:=@Buffer[0];
      if FColumnWidth < Buffer_Size then
      begin
        ErrH :=spssGetValueChar(FileHandle,VarH,P,self.ColumnWidth+1);
      end else
      begin
        ErrH :=spssGetValueChar(FileHandle,VarH,P,BUFFER_SIZE);
      end;
      CheckError(ErrH);
      SetString(strValue,P,buffer_size);
      strValue := Trim(strValue);

    if EnumCount >0 then
    begin
      if FFilter.FEnumItems.HashExists(FVarName+'|'+ strValue) then
      begin
        Result :=FFilter.FEnumItems.EnumValueByName(FVarName+'|'+ strValue)
      end else
         Result :=Null;
      begin
      end;
    end else
    begin
      result  :=strValue;
    end;
  end;
end;

function TSpssVariable.IsBoolean: Boolean;
var
  I: Integer;
  Value1,Value2:integer;
begin
  if (EnumCount <=2) and (enumCount>0) then
  begin
    if FList.Count =2 then
    begin
      Value1 :=(TObject(FList.Items[0]) as TSpssEnumItem).Value ;
      Value2 :=(TObject(FList.Items[1]) as TSpssEnumItem).Value ;
      if ((Value1 =1) and  (Value2 =0)) or ((Value1=0) and (Value2=1)) then
      begin
        result :=True;
      end else
      begin
        result :=false;
      end;
    end else
    begin
      if (TObject(FList.Items[0]) as TSpssEnumItem).Value =1 then
      begin
        result :=true;
      end else
      begin
        result :=false;
      end;
    end;
  end else
  begin
    result :=false;
  end;

end;

function TSpssVariable.KeyExists(Key: string): boolean;
begin
  Result :=FFilter.FEnumItems.HashExists(FVarName +'|'+ Key) ;

end;

procedure TSpssVariable.SetColumnWidth(const Value: Integer);
begin
  FColumnWidth := Value;
end;

procedure TSpssVariable.ConverttDateTimeType;
begin
  FVarType :=sNum;
  FColumnWidth :=20;
  FDecimal :=0;
  self.FPrintType :=SPSS_FMT_DATE;

end;

procedure TSpssVariable.DeleteRevItem(RevItem: TRevSpssVarItemX);
begin
  FRevItemList.Remove(RevItem);
end;

procedure TSpssVariable.EnumStringVar;
var
  I: Integer;
  V:TSpssVariable;
  str:string;
  enum:TSpssEnumItem;
  Reader:TSpssReader;
  FHash:TStringHash;
begin
  if self.VarType =sString  then
  begin
    if self.Filter is TSpssReader then
    begin
      Reader:=Self.Filter as TSpssReader;
      FHash :=TStringHash.Create(4096);
      try
        while Reader.ReadRecord do
        begin

          if (not v.ValueIsNull) and (V.Value <>'') then
          begin
            str :=V.Value;
            if Fhash.ValueOf(UpperCase(str))<0 then
            begin
              enum :=v.AddEnumItem(str,V.EnumCount+1);
              v.AddHash(str,enum);
              FHash.Add(UpperCase(str),enum.Index);
            end;
          end;
        end;
      finally
        FreeAndNil(FHash);
      end;
      Reader.SpssRecordSeek(0);
    end;




  end;
end;

function TSpssVariable.GetFirstRevItem: TRevSpssVarItemX;
begin
  if FRevItemList.Count >0 then
  begin
    result :=TObject(FRevItemList.Items[0]) as TRevSpssVarItemX;
  end else
  begin
    result :=nil;
  end;

end;

function TSpssVariable.GetMaxEnumValue: int64;
var
  I:integer;
begin
  result :=0;
  for I := 0 to self.EnumCount-1  do
  begin
    if self.EnumItem[I].Value > Result then
    begin
      result :=self.EnumItem[I].Value;
    end;

  end;

end;

function TSpssVariable.GetRevItem(Index: Integer): TRevSpssVarItemX;
begin

  Result :=TObject(FRevItemList.Items[Index]) as TRevSpssVarItemX;
end;

function TSpssVariable.GetRevItemCount: Integer;
begin
  Result :=FRevItemList.Count ;
end;

function TSpssVariable.GetRevVarLabel: string;
var
  REv:TRevObject;
  Item:TRevSpssVarItemX;
begin

  if FRevItemList.Count =1 then
  begin
    Item := TObject(FRevItemList.Items[0]) as TRevSpssVarItemX;
    result :=item.RevLabel;
  end else
  begin
    result :=FVarLabel;
  end;
 { if  (FVarRevised)  and (FFilter is TSpssReader) then
  begin
    REv :=(FFilter as TSpssReader).Rev;
    Item :=REv.GetRevSpssVar(Self);
    if Assigned(Item) then
      result :=item.RevLabel
    else
      result :=FVarLabel;

  end else
  begin
    Result := FVarLabel;
  end; }
end;

function TSpssVariable.GetRevVarName: AnsiString;
var
  Item:TRevSpssVarItemX;
begin
  if FRevItemList.Count=1 then
  begin
    Item :=TObject(FRevItemList.Items[0]) as TRevSpssVarItemX;
    result :=item.RevVarName;
  end else
  begin
    result :=FVarName;
  end;

end;

function TSpssVariable.GetVarRevised: Boolean;
begin
  Result := FRevItemList.Count>0;
end;

procedure TSpssVariable.SetFileHandle(const Value: Integer);
begin
  FFileHandle := Value;
end;

procedure TSpssVariable.SetRevVarLabel(const Value: string);
var
  Item:TRevSpssVarItemX;
begin
  if FRevItemList.Count=1 then
  begin
    item :=TObject(FRevItemList.Items[0]) as TRevSpssVarItemX;
    item.RevLabel :=Value;
  end;
end;

procedure TSpssVariable.SetVarLabel(const Value: string);
begin
  FVarLabel := Value;
end;

procedure TSpssVariable.SetspssType(const Value: TspssvarType);
begin
  FspssType := Value;
end;

procedure TSpssVariable.SetValue(const Value: Variant);
var
  strVal:RawByteString;
  ErrH :integer;
begin
  if self.Filter is TSpssReader then
  begin
    raise Exception.Create('Spss file is open read only mode ,you can''t write it!');
  end;
  if self.VarType =sString then
  begin
    strVal :=Value;
    ErrH :=spssSetValueChar(filter.FHandle,VarH,PAnsiChar(strVal));
    CheckError(ErrH);


  end else
  begin
    // Spss numeric




    case System.Variants.VarType(Value) of
      varBoolean :
      begin
        if Value =true then
        begin
          ErrH :=spssSetValueNumeric(filter.FHandle,VarH,1)
        end else
        if Value =false then
        begin
          ErrH :=spssSetValueNumeric(Filter.FHandle,VarH,0);
        end;
      end;
      varNull:
      begin
        //ErrH :=spssSetValueNumeric(Filter.FHandle,VarH,-1);
      end

    else

     ErrH :=spssSetValueNumeric(Filter.FHandle,VarH,Value);

    end;


   end;
  CheckError(ErrH);
end;

procedure TSpssVariable.SetVarH(const Value: Double);
begin
  FVarH := Value;
end;

function TSpssVariable.ValueIsNull: Boolean;
begin
  result :=  System.Variants.VarType(self.Value) =VarNull;
end;

constructor TSpssEnums.Create(ItemClass: TCollectionItemClass);
begin
  inherited;
  FHash := TStringHash.Create();
end;

destructor TSpssEnums.Destroy;
begin
  FreeAndNil(FHash);
  inherited Destroy;
end;

procedure TSpssEnums.AddHash(Key: string; IndexValue: Integer);

begin
  if FHash.ValueOf(Key)<0 then
  begin
    FHash.Add(Key,IndexValue);
  end;

end;

function TSpssEnums.EnumValueByName(EnumName: string): Integer;
var
  Index:integer;
begin
  Index :=FHash.ValueOf(EnumName);
  result :=Index;

end;

function TSpssEnums.GetEnumItem(Index: Integer): TSpssEnumItem;
begin
  Result :=self.Items[Index] as TSpssEnumItem ;
end;

function TSpssEnums.HashExists(Key: string): Boolean;
begin
  Result := FHash.ValueOf(Key)>=0;

end;

constructor TSVarSelector.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FList := TList.Create();
end;

destructor TSVarSelector.Destroy;
begin
  FreeAndNil(FList);
  inherited Destroy;
end;

function TSVarSelector.AddVar(SpssVar: TSpssVariable): Integer;
begin

  result := FList.IndexOf(SpssVar);
  if Result <0 then
    result :=FList.Add(SpssVar);
end;

procedure TSVarSelector.Clear;
begin
  FList.Clear;
end;

function TSVarSelector.Exists(SVar: TSpssVariable): Boolean;
begin
  result :=flist.IndexOf(Svar)>=0;
end;

function TSVarSelector.GetCount: Integer;
begin

  Result :=FList.Count ;
end;

function TSVarSelector.GetVarItem(Index: Integer): TSpssVariable;
begin
  Result := TObject(FList.Items[Index]) as TSpssVariable ;
end;

procedure TSVarSelector.RemoveVar(Spssvar: TSpssVariable);
begin
  flist.Remove(Spssvar);
end;

constructor TSpssFilter.Create(AOwner: TComponent);
begin
  inherited ;
  FVars := TObjectList.Create();
  FHash :=TObjectHash.Create(10000);
  FLang :=Lang_SIMPLE_CHINESE;
  FValueHash := TStringHash.Create(65535);
  FEnumItems := TSpssEnums.Create(TspssEnumItem);
  FRev :=TRevObject.Create(Self);
  Frev.FParentFilter :=self;
end;

destructor TSpssFilter.Destroy;
begin
  FreeAndNil(FRev);
  FreeAndNil(FEnumItems);
  FreeAndNil(FValueHash);
  FreeAndNil(FHash);
  FreeAndNil(FVars);
  inherited Destroy;
end;

function TSpssFilter.AddVariable(VarName: AnsiString; varType: TspssVarType;
    vHandle: Double): TSpssVariable;
var
  v:TSpssVariable;
begin
  if Assigned(FHash.ValueOf(UpperCase(VarName))) then
    raise Exception.Create('Variable already Exists!'+VarName);
  v :=TSpssVariable.Create;
  v.VarName :=VarName ;
  v.VarType :=varType;
  FVars.Add(V);
  FHash.Add(UpperCase(v.VarName),V);
  result :=v;


end;

class function TSpssFilter.GetEncoding: ShortInt;

begin
  Result :=spssGetInterfaceEncoding;
end;

function TSpssFilter.GetSpssVarItem(Index: Integer): TSpssVariable;
begin
  Result :=FVars.Items[Index] as TSpssVariable ;
end;

function TSpssFilter.GetSpssVarByName(Name:string): TSpssVariable;

begin
  result:=FHash.ValueOf(UpperCase(Name)) as TSpssVariable;

end;

function TSpssFilter.GetVarCount: Integer;
begin
  Result :=FVars.Count ;
end;

class procedure TSpssFilter.SetEncoding(const Value: ShortInt);
var
  ErrH:integer;
begin
  ErrH :=spssSetInterfaceEncoding(Value);
  CheckError(ErrH);
end;

procedure TSpssFilter.SetLocal;
var
  ErrH:integer;
  Lang :PAnsiChar;
begin
  Lang :=spssSetLocale(LC_ALL,PAnsiChar(FLang));
  //CheckError(ErrH);
end;

function TSpssFilter.VarExists(const VarName: string): Boolean;
begin
  Result :=Assigned(FHash.ValueOf(UpperCase(VarName)));
end;

constructor TSpssWriter.Create(AOwner: TComponent);
begin
  inherited;
  FSpssState :=sOPenWrite;
  FOpend :=false;
end;

destructor TSpssWriter.Destroy;
begin
  if FOpend then
    Close;
  inherited;

end;

procedure TSpssWriter.ChangeVarName(const V: TSpssVariable; const NewVarName:
    string);
var
  Index:integer;
begin
  if Assigned(FHash.ValueOf(UpperCase(NewVarName))) then
    raise Exception.Create('Var exists!'+NewvarName);

  FHash.Remove(UpperCase( v.VarName));
  V.VarName :=NewVarName;
  FHash.Add(V.VarName,V);
end;

procedure TSpssWriter.Close;
var
  ErrH:integer;
begin
  ErrH :=spssCloseWrite(FHandle);
  CheckError(ErrH);
  FOpend :=false;
end;

procedure TSpssWriter.CommitCase;
var
  ErrH:integer;
begin
  ErrH:=spssCommitCaseRecord(FHandle);
  CheckError(ErrH);
end;

procedure TSpssWriter.CreateHeader;
var
  I: Integer;
  V:TSpssVariable;
  ErrH:Integer;
begin
  for I := 0 to self.VarCount-1 do
  begin
    V :=self.SpssVarItem[I];
    {$IFDEF CODESITE}
    CodeSite.Send('VariableName:',v.VarName);
    {$ENDIF}
    CreateVarHeader(V);
  end;
  ErrH:=spssCommitHeader(FHandle);
  CheckError(ErrH);


end;

procedure TSpssWriter.CreateVarHeader(const SpssVar: TSpssVariable);
var
  VName,Lab:RawByteString;
  ErrH:integer;
  I:Integer;
  Item:TSpssEnumItem;
  EnumLabel:RawByteString;
  VarH:double;
begin
  VName :=SpssVar.VarName;
  Lab :=SpssVar.VarLabel;
  if SpssVar.VarType =sString then
  begin
    ErrH:=spssSetVarName(FHandle,PAnsiChar(VName),SpssVar.ColumnWidth);
    CheckError(ErrH);
  end else
  if SpssVar.VarType=sNum then
  begin
    ErrH :=spssSetVarName(FHandle,PAnsiChar(VName),SPSS_NUMERIC);
    CheckError(ErrH);
    ErrH:=spssSetVarPrintFormat(FHandle,PAnsiChar(VName),SpssVar.PrintType,SpssVar.Decimal,spssvar.ColumnWidth);
    CheckError(ErrH);
  end else
  if SpssVar.VarType=sEnum then
  begin
    ErrH :=spssSetVarName(FHandle,PAnsiChar(VName),SPSS_NUMERIC);
    CheckError(ErrH);
    ErrH:=spssSetVarPrintFormat(FHandle,PAnsiChar(VName),SpssVar.PrintType,spssvar.Decimal,SpssVar.ColumnWidth);
    CheckError(ErrH);
  end;
  if Length(Lab)>255 then
  begin
    lab :=Copy(Lab,255);
  end;
  ErrH :=spssSetVarLabel(FHandle,PAnsiChar(VName),PAnsiChar(Lab));
  CheckError(ErrH);
  if SpssVar.VarType=sEnum then
  begin
    for I := 0 to SpssVar.EnumCount-1 do
    begin
      Item :=SpssVar.EnumItem[I];
      EnumLabel :=item.EnumLabel;
      if Length(EnumLabel)>120then
      begin
          EnumLabel :=Copy(EnumLabel,120);
      end;
      ErrH :=spssSetVarNValueLabel(FHandle,PAnsiChar(VName),Item.Value,PAnsiChar(EnumLabel));
      CheckError(ErrH);
    end;
  end;
  ErrH := spssGetVarHandle(FHandle,PAnsiChar(VName),VarH);
  CheckError(ErrH);
  SpssVar.VarH :=VarH;
end;

procedure TSpssWriter.NewCase;
begin
  // TODO -cMM: TSpssWriter.NewCase default body inserted

end;

function TSpssWriter.NewVariable(const VarName, VarLabel: string; const
    VarType: TspssVarType): TSpssVariable;
begin

  result :=self.AddVariable(VarName,VarType,-1);
  result.Filter :=self;
  result.VarLabel :=VarLabel;
  if VarType in [sNum,sEnum] then
  begin
    result.PrintType :=SPSS_FMT_F;
    result.Decimal :=0;
    result.ColumnWidth :=5;
  end else
  if vartype in[sString] then
  begin
    Result.ColumnWidth :=SPSS_STRING;
  end;




end;

procedure TSpssWriter.OpenSpssFile(FileName: AnsiString);
var
  ErrH:Integer;
  Lang:string;
begin



  ErrH:=spssOpenWrite(PAnsiChar(FileName),FHandle);
  CheckError(ErrH);
  FOpend :=true;
  //ErrH :=spssSetInterfaceEncoding(SPSS_ENCODING_UTF8);

  Lang :='zh_CN.65001';
  lang :=spssSetLocale(LC_ALL,PAnsiChar(FLang));
    //create spss file
  // compress File
  spssSetCompression(FHandle,1);

end;

constructor TRevObject.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FVarRevs := TCollection.Create(TRevSpssVarItemX);
  FVarRevHash :=TObjectHash.Create(4096);
  FValueHash :=TObjectHash.Create(4006);
end;

destructor TRevObject.Destroy;
begin
  FreeAndNil(FValueHash);
  FreeAndNil(FVarRevHash);
  FreeAndNil(FVarRevs);
  inherited Destroy;
end;

procedure TRevObject.ConvertEnumToString(const SVar: TSpssVariable);
var
  RevItem:TRevSpssVarItemX;
  I:integer;
begin
  if Svar.VarType =sEnum then
  begin
    RevItem := FVarRevHash.ValueOf(UpperCase(svar.VarName)) as TRevSpssVarItemX;
    if not Assigned(RevItem) then
    begin
      RevItem :=TRevSpssVarItemX.Create(FVarRevs);
    end;
    revitem.SourceVar :=SVar;
    RevItem.RevVarName :=Svar.VarName;
    RevItem.RevLabel :=SVar.VarLabel;
    RevItem.RevVarType :=sString;
    RevItem.RevColumnWidth :=255;
    RevItem.RevDecimal :=0;
    //RevItem.PrintType :=-1;
    RevItem.RevType :=revConverString;
    SVar.AddRev(RevItem);
    FVarRevHash.Add(UpperCase(svar.VarName),revItem);
    // Add Rev Enum Item;
    for I := 0 to SVar.EnumCount-1 do
    begin
      RevItem.AddRevValue(IntToStr(SVar.EnumItem[I].Value),svar.EnumItem[I].EnumLabel);
    end;

  end;
end;

procedure TRevObject.ConvertStringToEnum(const SVar: TSpssVariable);
var
  SR:TSpssReader;
  RevItem:TRevSpssVarItemX;
  Index:integer;
  FHash:TStringHash;
begin
  if Assigned(SVar) then
  begin
    FHash :=TStringHash.Create(65535);
    try
      SR :=Svar.Filter as TSpssReader;
      RevItem := FVarRevHash.ValueOf(UpperCase(svar.VarName)) as TRevSpssVarItemX;
      if not Assigned(RevItem) then
      begin
        RevItem :=TRevSpssVarItemX.Create(FVarRevs);
        RevItem.SourceVar :=SVar;
      end;
      RevItem.RevVarName :=Svar.VarName;
      RevItem.RevLabel :=SVar.VarLabel;
      RevItem.RevVarType :=sEnum;
      RevItem.RevColumnWidth :=8;
      RevItem.RevDecimal :=0;
      RevItem.PrintType :=SPSS_FMT_F;
      RevItem.RevType :=revStringToEnum;
      SVar.AddRev(RevItem);
      FVarRevHash.Add(UpperCase(svar.VarName),RevItem);

      sr.SpssRecordSeek(0);
      Index :=0;
      while sr.ReadRecord do
      begin
        //
        if (not svar.ValueIsNull) and (Svar.Value<>'')  then
        begin
          if FHash.ValueOf(UpperCase(SVar.Value))<0 then
          begin
            Inc(Index);
            RevItem.AddRevValue(SVar.value,Svar.Value,Index);
            FHash.Add(UpperCase(svar.Value),1);
          end;
        end;
      end;
    finally
      FreeAndNil(Fhash);
    end;
  end;
end;

function TRevObject.GetRevVar(ID: Integer): TRevSpssVarItemX;
begin
  Result :=FVarRevs.FindItemID(ID) as TRevSpssVarItemX ;
end;

function TRevObject.GetRevSpssVar(SpssVar: TSpssVariable): TRevSpssVarItemX;
var
  ID :Integer;
begin
  result :=FVarRevHash.ValueOf(UpperCase(SpssVar.VarName)) as TRevSpssVarItemX;
  {ID :=FVarRevHash.ValueOf(UpperCase(SpssVar.VarName));
  if ID>=0 then

    Result :=self.RevVar[ID]
  else
    result :=nil; }
end;

procedure TRevObject.ReNameAndLabel(const SpssVar: TSpssVariable; const
    NewVarName, NewLabel: string);
var
  Index:Integer;
  item :TRevSpssVarItemX;
  I:Integer;
  ValItem:TRevSpssValue;
begin
  Item :=FVarRevHash.ValueOf(UpperCase(SpssVar.VarName)) as TRevSpssVarItemX;

  if Assigned(item) then   raise Exception.Create('VarName is revised:'+spssvar.VarName);
  if spssvar.Filter.VarExists(UpperCase(NewVarName)) then raise Exception.Create('Variable Exits!'+newvarname);
  item :=TRevSpssVarItemX.Create(FVarRevs);
  item.RevType :=revVName_Label;
  item.RevVarType :=spssvar.VarType;
  item.RevVarName :=NewVarName;
  item.RevLabel :=NewLabel;
  item.RevColumnWidth :=spssvar.ColumnWidth;
  item.PrintType :=SpssVar.PrintType;
  item.TargetVar :=SpssVar;
  item.SourceVar :=SpssVar;



  for I := 0 to spssvar.EnumCount-1 do
  begin
    ValItem :=item.AddRevValue(spssvar.EnumItem[I].EnumLabel,SpssVar.EnumItem[I].Value);
    ValItem.RevLabel :=SpssVar.EnumItem[I].EnumLabel;
  end;



  spssvar.AddRev(item);

  FVarRevHash.Add(UpperCase(SpssVar.VarName),item);


end;

function TRevObject.RevVarExits(const SpssVar: TSpssVariable): Boolean;
begin
  Result :=Assigned(FVarRevHash.ValueOf(UpperCase(spssvar.VarName)));
end;

function TRevObject.GetVarCount: Integer;
begin
  Result :=FVarRevs.Count ;
end;

procedure TRevObject.RevokeRev(SpssVar: TSpssVariable);
var
  Item:TRevSpssVarItemX;
  I: Integer;
begin
  for I:=SpssVar.RevItemCount-1 downto 0 do
  begin
    Item :=SpssVar.RevItem[I];
    if Assigned(Item) then
    begin
      FVarRevHash.Remove(UpperCase(SpssVar.VarName));
      FVarRevs.Delete(Item.Index);
      SpssVar.DeleteRevItem(Item);
    end;
  end;


end;

procedure TRevObject.AddOpenVarToBoolanVar(const SourceVar: TSpssVariable;
    const TextLabel, VarPrefix: string; const ValueNum: Integer);
var
  NewVarName:string;
  Item:TRevSpssVarItemX;
  Svar:TSpssVariable;
begin
  NewVarName :=VarPrefix +'_'+IntTostr(ValueNum);
  if not FParentFilter.VarExists(NewVarName) then
  begin

    Svar :=FParentFilter.AddVariable(NewVarName,sEnum,-1);
    svar.AddEnumItem('yes',1);
    svar.AddEnumItem('no',0);
    Item :=TRevSpssVarItemX.Create(FVarRevs);
    Item.RevVarName :=NewVarName;
    item.RevLabel :=TextLabel;
    item.RevType :=RevNewVar;
    FVarRevHash.Add(UpperCase(NewVarName),Item);
    FValueHash.Add(UpperCase(TextLabel),Item);
    svar.AddRev(Item);



  end else
  begin
    //raise Exception.Create('Variable Exists!'+ NewVarName);
  end;

 {NewVarName :=VarPrefix+'_'+inttostr(ValueNum);
  SourceVar.VarRevised :=true;



  if not RevVarExits(NewVarName) then
  begin
    Item :=TSpssVariable.Create(FVarRevs);
    Item.RevLabel :=TextLabel;
    Item.RevVarName :=NewVarName;
    Item.RevVarType :=sEnum;
    Item.AddRevValue('yes',1);
    Item.AddRevValue('no',0);
    Item.RevColumnWidth :=1;
    FVarRevHash.Add(UpperCase(NewVarName),Item);
    item.RevType :=RevNewVar;
  end else
  begin
    item :=RevVaByName(NewVarName);
  end;
  Item.Obj :=nil;

  FValueHash.Add(UpperCase(TextLabel),Item);

  {Svar :=SourceVar.Filter.SpssVarByName[NewVarName];

  if Assigned(Svar) then
  begin
    raise Exception.Create('Var Exists!'+NewVarName);

  end else
  begin

  end;}

end;

procedure TRevObject.AddStringValueFix(const SVar: TSpssVariable; const
    NewValue, OldValue: string);
var
  Item:TRevSpssVarItemX;
begin

  if SVar.VarType=sString then
  begin
    Item := FVarRevHash.ValueOf(UpperCase(svar.VarName)) as TRevSpssVarItemX;
    if not Assigned( Item) then
    begin
      Item :=TRevSpssVarItemX.Create(FVarRevs);
      Item.RevType :=RevFixValue;
      item.RevLabel :=svar.VarLabel;
      item.RevVarName :=svar.VarName;
      item.RevColumnWidth :=svar.ColumnWidth;
      item.RevDecimal :=svar.Decimal;
      Item.RevVarType :=svar.VarType;
      Item.SourceVar :=SVar;
      FVarRevHash.Add(UpperCase(svar.VarName),Item);
      svar.AddRev(item);
    end;
    if not  item.ValueExits(OldValue) then
    begin
      item.AddRevValue(OldValue,NewValue) ;
    end else
    begin
      //
    end;



  end else
  begin
    raise Exception.Create('Error Variable Type,Please select a string variable ');
  end;
end;

procedure TRevObject.ClearAll;
begin
  FValueHash.Clear;
  FVarRevHash.Clear;
  FVarRevs.Clear;
end;

procedure TRevObject.ConvertSingleVarToMultiVar(const SVar: TSpssVariable;
    const Var_Prefix: string);
var
  I:Integer;
  revItem:TRevSpssVarItemX;
  Item:TspssEnumItem;
  RevVal:TRevSpssValue;
begin
  if Assigned(FVarRevHash.ValueOf(UpperCase(SVar.VarName))) then
  begin
    raise Exception.Create('Rev Exits!'+Svar.VarName);
  end else
  begin
    revItem :=TRevSpssVarItemX.Create(FvarRevs);
    revItem.SourceVar :=SVar;
    RevItem.RevVarName :=Svar.VarName;
    RevItem.RevLabel :=SVar.VarLabel;
    revItem.RevVarType :=svar.VarType;
    revItem.RevColumnWidth :=svar.ColumnWidth;
    revitem.RevDecimal :=svar.Decimal;
    revitem.RevType := RevSingleToMulti;
    for I := 0 to SVar.EnumCount-1 do
    begin
      item :=Svar.EnumItem[I];
      RevVal :=revitem.AddRevValue(IntToStr(Item.Value),item.EnumLabel,item.Value);
      RevVal.RevName :=Var_Prefix+'_'+IntTostr(Item.Value);



    end;
    FVarRevHash.Add(UpperCase(SVar.VarName),REvItem);
    svar.AddRev(REvItem);
  end;

end;

function TRevObject.GetNewNodeName(NodeName: string): string;
var
  Filter:TSpssFilter;
  Index:integer;
begin
  Filter :=self.ParentFilter;
  Index :=1;
  if filter.VarExists(NodeName) then
  begin
    while Filter.VarExists(NodeName+'_'+IntTostr(Index)) do
    begin

      Inc(Index);
    end;
    result :=NodeName+'_'+Inttostr(index);
  end else
  begin
    result :=NodeName;
  end;
end;

function TRevObject.GetRevVarByIndex(Index: Integer): TRevSpssVarItemX;
begin
  Result :=FVarRevs.Items[Index] as TRevSpssVarItemX ;
end;

function TRevObject.RevVaByName(const VarName: string): TRevSpssVarItemX;
begin
  Result :=FVarRevHash.ValueOf(UpperCase(VarName)) as TRevSpssVarItemX;
end;

function TRevObject.RevVarByValue(const TextLabel: string): TRevSpssVarItemX;
begin
  Result :=FVarRevHash.ValueOf(UpperCase(TextLabel)) as TRevSpssVarItemX;
end;

function TRevObject.RevVarExits(const VarName: string): Boolean;
begin
  Result :=Assigned(FVarRevHash.ValueOf(UpperCase(VarName)));
end;

constructor TRevSpssVarItemX.Create(Collection: TCollection);
begin
  inherited ;
  FItems :=TCollection.Create(TRevSpssValue);
  FHash := TObjectHash.Create(4096);

end;

destructor TRevSpssVarItemX.Destroy;
begin
  FreeAndNil(FItems);
  FreeAndNil(FHash);
  inherited ;
end;

function TRevSpssVarItemX.AddRevValue(const Key: string; const Value: Variant):
    TRevSpssValue;
var
  Index:integer;
  Item:TRevSpssValue;
begin
  Item :=FHash.ValueOf(uppercase(Key)) as TRevSpssValue;
  if not Assigned(Item) then
  begin
    Item :=TRevSpssValue.Create(FItems);
    FHash.Add(UpperCase(Key),Item);
  end ;
 { if Index<0 then
  begin
    Item :=TRevSpssValue.Create(FItems);
    FHash.Add(IntToStr(item.Value),Item.Index);

  end else
  begin
    Item :=fitems.items[Index] as TRevSpssValue;
  end; }
  Item.RevLabel :='';
  Item.RevValue :=Value;
  Result :=Item;
end;

function TRevSpssVarItemX.AddRevValue(const Key, RLabel: string; const Value:
    Variant): TRevSpssValue;

begin
  result :=AddRevValue(Key,Value);
  result.RevLabel :=RLabel;
end;

function TRevSpssVarItemX.ConvertValue(const Value: Variant): Variant;
var
  ValKey:string;
begin
  case RevType of
    revConverString,revStringToEnum:
    begin
      ValKey :=VarToStr(Value);
      result :=GetRevValueByKey(ValKey);
    end;
    RevItemValue:
    begin
      raise Exception.Create('Error Message');
    end;
    revVName_Label:
    begin
      result :=Value;
    end;
    RevFixValue:
    begin
      ValKey :=VarToStr(Value);
      result :=GetRevValueByKey(ValKey);
      if VarType(result)=varEmpty then
      begin
        Result :=Value;
      end;

    end;
  end;


end;

function TRevSpssVarItemX.GetItemCount: Integer;
begin
  Result :=FItems.Count ;
end;

function TRevSpssVarItemX.GetRevValueByKey(const Key: string): Variant;
Var
  ValItem:TRevSpssValue;
  Val:Variant;
begin
  ValItem :=fhash.ValueOf(UpperCase(key)) as TRevSpssValue;
  if Assigned(ValItem) then
  begin
    Val :=ValItem.RevValue;
  end else
  begin
    VarClear(Val);
  end;
  result :=Val;

end;

procedure TRevSpssVarItemX.ChangeValue(const Key: string; const NewValue:
    Integer);
Var
  ValItem:TRevSpssValue;
begin
  ValItem :=fhash.ValueOf(UpperCase(key)) as TRevSpssValue;
  if Assigned(ValItem) then
  begin
    ValItem.RevValue :=NewValue;

  end else
  begin
    raise Exception.Create('not found "'+Key+'"');
  end;
end;

function TRevSpssVarItemX.GetValueItem(Index: Integer): TRevSpssValue;
begin
  Result := FItems.Items[Index] as TRevSpssValue ;
end;

function TRevSpssVarItemX.GetValueItemByKey(Key:string): TRevSpssValue;
begin

  Result :=FHash.ValueOf(uppercase(Key)) as Trevspssvalue ;
end;

function TRevSpssVarItemX.LabelExists(const TextLabel: string): Boolean;
begin
  Result := Assigned(FHash.ValueOf(UpperCase(TextLabel)));
end;

function TRevSpssVarItemX.ValueExits(const Key: string): Boolean;
begin
  Result := Assigned(FHash.ValueOf(Key));
end;

constructor TRevSpssVarItem.Create(Collection: TCollection);
begin
  inherited ;
  FPrintType := 0;
end;

procedure TRevSpssVarItem.SetRevType(const Value: TRevType);
begin
  FRevType := Value;
end;

procedure TRevSpssVarItem.SetRevVarType(const Value: TspssVarType);
begin
  FRevVarType := Value;
end;

end.
