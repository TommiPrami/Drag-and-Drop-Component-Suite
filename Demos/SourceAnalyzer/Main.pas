unit Main;

interface

{$include dragdrop.inc} // Disables .NET warnings

uses
  System.SysUtils, System.ImageList, System.Actions, System.Classes,
  WinApi.ActiveX, WinApi.Windows, WinApi.Messages,
  Vcl.ActnList, Vcl.ComCtrls, Vcl.ToolWin, Vcl.Graphics, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.Menus, Vcl.Dialogs, Vcl.ImgList, Vcl.Controls, Vcl.Forms, Vcl.Imaging.pngimage,
  Vcl.Buttons,
  DragDrop, DropTarget;

const
  MSG_ASYNC_DROP = WM_USER;

const
  MAX_DATA = 32768; // Max bytes to render in preview

type
  TOmnipotentDropTarget = class(TCustomDropMultiTarget)
  protected
    function DoGetData: boolean; override;
  public
    function HasValidFormats(const ADataObject: IDataObject): boolean; override;
  end;

  TDropData = record
    FormatEtc: TFormatEtc;
    ActualTymed: longInt;
    Data: AnsiString;
    HasFetched: boolean;
    HasData: boolean;
    IsPending: boolean;
  end;
  PDropData = ^TDropData;

  TFormMain = class(TForm)
    Panel2: TPanel;
    Splitter1: TSplitter;
    StatusBar1: TStatusBar;
    EditHexView: TRichEdit;
    ListViewDataFormats: TListView;
    ActionList1: TActionList;
    ActionClear: TAction;
    ActionPaste: TAction;
    ImageListMain: TImageList;
    ActionSave: TAction;
    SaveDialog1: TSaveDialog;
    ActionDirTarget: TAction;
    ActionDirSource: TAction;
    IntroView: TRichEdit;
    ActionPrefetch: TAction;
    PanelError: TPanel;
    PanelErrorInner: TPanel;
    Panel3: TPanel;
    Label1: TLabel;
    LabelError: TLabel;
    Panel4: TPanel;
    Image1: TImage;
    ImageListStatus: TImageList;
    PanelButtons: TPanel;
    ButtonClear: TButton;
    ButtonPaste: TButton;
    ButtonSave: TButton;
    CheckBoxPrefetch: TCheckBox;
    GroupBoxDirection: TGroupBox;
    RadioButtonDirectionTarget: TRadioButton;
    RadioButtonDirectionSource: TRadioButton;
    CheckBoxAsync: TCheckBox;
    ActionAsync: TAction;
    ButtonAsyncAvailable: TSpeedButton;
    ActionAsyncAvailable: TAction;
    ActionAsyncNotAvailable: TAction;
    ActionAsyncActive: TAction;
    ActionAsyncInfo: TAction;
    procedure FormCreate(Sender: TObject);
    procedure ListViewDataFormatsDeletion(Sender: TObject;
      Item: TListItem);
    procedure ListViewDataFormatsSelectItem(Sender: TObject;
      Item: TListItem; Selected: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure ActionClearExecute(Sender: TObject);
    procedure ActionPasteUpdate(Sender: TObject);
    procedure ActionPasteExecute(Sender: TObject);
    procedure ActionSaveExecute(Sender: TObject);
    procedure ActionSaveUpdate(Sender: TObject);
    procedure ListViewDataFormatsAdvancedCustomDrawSubItem(
      Sender: TCustomListView; Item: TListItem; SubItem: Integer;
      State: TCustomDrawState; Stage: TCustomDrawStage;
      var DefaultDraw: Boolean);
    procedure ActionDummyExecute(Sender: TObject);
  private
    FDataObject: IDataObject;
    FDropTarget: TCustomDropTarget;
    FAsyncPending: integer;
    procedure OnDrop(Sender: TObject; ShiftState: TShiftState;
      APoint: TPoint; var Effect: Longint);
    procedure OnDragOver(Sender: TObject; ShiftState: TShiftState;
      APoint: TPoint; var Effect: Integer);
    procedure OnDragEnter(Sender: TObject; ShiftState: TShiftState;
      APoint: TPoint; var Effect: Integer);
    procedure OnDragLeave(Sender: TObject);
    procedure OnAsyncThreadTerminate(Sender: TObject);
    procedure OnAsyncDrop(Sender: TObject);
    procedure OnAsyncDropFailed(Sender: TObject);
  private
    procedure MsgAsyncDrop(var Msg: TMessage); message MSG_ASYNC_DROP;
  private
    function DataToHexDump(const Data: AnsiString): string;
    function GetDataSize(const AFormatEtc: TFormatEtc): integer;
    function VerifyMedia(var AFormatEtc: TFormatEtc): longInt;
    procedure CheckAsyncPending;
    procedure LoadRTF(const s: string);
    procedure Error(const Msg: string);
    procedure Init;
    procedure Clear;
    procedure Reset;
    function PrefetchDropData(var DropData: TDropData): boolean; // Returns False if data is pending
    function GetDropDataAsString(var DropData: TDropData): AnsiString;
    procedure UpdateAsyncAbility(const DataObject: IDataObject);
    property DataObject: IDataObject read FDataObject;
  public
  end;

var
  FormMain: TFormMain;

implementation

{$R *.dfm}

uses
  WinApi.CommCtrl,
  Win.ComObj,
  DragDropFormats;

resourcestring
  sIntro = '{\rtf1\ansi\ansicpg1252\deff0\deflang1030{\fonttbl{\f0\fswiss\fcharset0 Arial;}{\f1\fnil\fcharset2 Symbol;}}'+#13+
    '\viewkind4\uc1\pard{\pntext\f1\''B7\tab}{\*\pn\pnlvlblt\pnf1\pnindent0{\pntxtb\''B7}}\fi-284\li284\f0\fs20Drag data from any application and drop it on this window.\line'+#13+
    'The clipboard formats supported by the drop source will be displayed in the list above.\par '+#13+
    '{\pntext\f1\''B7\tab}Select a clipboard format from the list to display its content here.\par }';

const
  ImageIndexWarning = 0;
  ImageIndexPending = 1;
  ImageIndexError = 2;

function GetDropData(const ADataObject: IDataObject; const AFormatEtc: TFormatEtc; var ADropData: TDropData): boolean;
var
  ClipFormat: TRawClipboardFormat;
begin
  // Create a temporary clipboard format object to retrieve the raw data
  // from the drop source.
  ClipFormat := TRawClipboardFormat.CreateFormatEtc(AFormatEtc);
  try
    // Note: It would probably be better (more efficient & safer) if we used a
    // custom clipboard format class which only copied a limited amount of data
    // from the source data object. However, I'm lazy and this solution works
    // fine most of the time.
    if (ClipFormat.GetData(ADataObject)) then
    begin
      ADropData.Data := Copy(ClipFormat.AsString, 1, MAX_DATA);
      ADropData.HasData := True;

      Result := True;
    end else
      Result := False;
  finally
    ClipFormat.Free;
  end;
end;

{ TAsyncOperationThread }

type
  TAsyncOperationThread = class(TThread)
  private
    FDataObject: IDataObject;
    FDropData: PDropData;
    FFormatEtc: TFormatEtc;

    FDataObjectStream: pointer;
    FAsyncOperationStream: pointer;

    FOnDrop: TNotifyEvent;
    FOnDropFailed: TNotifyEvent;
  protected
    procedure Execute; override;

    property DataObjectStream: pointer read FDataObjectStream;
    property AsyncOperationStream: pointer read FAsyncOperationStream;
  public
    constructor Create(const ADataObject: IDataObject; ADropData: PDropData; const AFormatEtc: TFormatEtc);
    destructor Destroy; override;

    property DropData: PDropData read FDropData;
    property OnDrop: TNotifyEvent read FOnDrop write FOnDrop;
    property OnDropFailed: TNotifyEvent read FOnDropFailed write FOnDropFailed;
  end;

constructor TAsyncOperationThread.Create(const ADataObject: IDataObject; ADropData: PDropData; const AFormatEtc: TFormatEtc);
begin
  inherited Create(True);
  FreeOnTerminate := True;

  FDataObject := ADataObject; // Keeps data object alive until we're done with it

  FDropData := ADropData;
  FFormatEtc := AFormatEtc;

  FDropData.IsPending := True;

  OleCheck(CoMarshalInterThreadInterfaceInStream(IDataObject, FDataObject, IStream(FDataObjectStream)));
  OleCheck(CoMarshalInterThreadInterfaceInStream(IAsyncOperation2, FDataObject, IStream(FAsyncOperationStream)));
end;

destructor TAsyncOperationThread.Destroy;
begin
  FDataObjectStream := nil;
  FAsyncOperationStream := nil;
  FDataObject := nil;
  inherited Destroy;
end;

procedure TAsyncOperationThread.Execute;
var
  DataObject: IDataObject;

  function DoGetDropData: boolean;
  begin
    // Get data from the drop source using the marshalled data object
    if (GetDropData(DataObject, FFormatEtc, FDropData^)) then
    begin
      // Generate an OnDrop event
      // Note that this event is executed in the context of this thread and thus
      // must adhere to the rules of thread safe use of the VCL (e.g. don't
      // update visual stuff directly).
      if (Assigned(FOnDrop)) then
        FOnDrop(Self);

      Result := True;
    end else
      Result := False;
  end;

  procedure DoFailed;
  begin
    if (Assigned(FOnDropFailed)) then
      FOnDropFailed(Self);
  end;

var
  Res: HResult;
  AsyncOperation: IAsyncOperation2;
begin
  CoInitialize(nil);
  try
    try
      OleCheck(CoGetInterfaceAndReleaseStream(IStream(DataObjectStream), IDataObject, DataObject));
      OleCheck(CoGetInterfaceAndReleaseStream(IStream(AsyncOperationStream), IAsyncOperation, AsyncOperation));

      Res := S_OK;

      // Retry with lindex=0 if lindex=-1 failed
      if (not DoGetDropData) and (FFormatEtc.lindex = -1) then
      begin
        FFormatEtc.lindex := 0;
        if (not DoGetDropData) then
          Res := E_FAIL;
      end;

    except
      on E: EOleSysError do
        Res := E.ErrorCode
      else
        Res := E_UNEXPECTED;
    end;

    // Notify drop source that we are done
    AsyncOperation.EndOperation(Res, nil, DROPEFFECT_NONE);

    FDropData.IsPending := False;

  finally
    DataObject := nil;
    AsyncOperation := nil;
    CoUninitialize;
  end;

  if (Failed(Res)) then
    DoFailed;
end;

{ TOmnipotentDropTarget }

function TOmnipotentDropTarget.DoGetData: boolean;
begin
  Result := True;
end;

function TOmnipotentDropTarget.HasValidFormats(const ADataObject: IDataObject): boolean;
begin
  Result := True;
end;

{ TFormMain }

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  Clear;
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
  FDropTarget := TOmnipotentDropTarget.Create(Self);
  FDropTarget.DragTypes := [dtCopy, dtLink, dtMove];
  FDropTarget.OptimizedMove := True;
  FDropTarget.Target := Self;
  FDropTarget.OnEnter := Self.OnDragEnter;
  FDropTarget.OnDrop := Self.OnDrop;
  FDropTarget.OnDragOver := Self.OnDragOver;
  FDropTarget.OnLeave := Self.OnDragLeave;
  FDropTarget.OnGetDropEffect := Self.OnDragOver;
  LoadRTF(sIntro);
  Init;
end;

function TFormMain.GetDataSize(const AFormatEtc: TFormatEtc): integer;
var
  Medium: TStgMedium;
  RetryFormatEtc: TFormatEtc;
  Res: HRESULT;
begin
  FillChar(Medium, SizeOf(Medium), 0);
  Res := DataObject.GetData(AFormatEtc, Medium);

  // Retry with lindex=0 if lindex=-1 failed
  // Windows Explorer returns DV_E_LINDEX when lindex=-1 with CF_FILECONTENTS
  // Windows Explorer CompressedFolder namespace (e.g. Zip) returns E_INVALIDARG when lindex=-1 with CF_FILECONTENTS
  if ((Res = DV_E_LINDEX) or (Res = E_INVALIDARG)) and (AFormatEtc.lindex = -1) then
  begin
    RetryFormatEtc := AFormatEtc;
    RetryFormatEtc.lindex := 0;
    Res := DataObject.GetData(RetryFormatEtc, Medium);
  end;

  if (Succeeded(Res)) then
  begin
    try
      Result := GetMediumDataSize(Medium);
    finally
      ReleaseStgMedium(Medium);
    end;
  end else
    Result := -1;
end;

procedure TFormMain.OnAsyncDrop(Sender: TObject);
begin
  // Post message to have event handled in main thread
  PostMessage(Handle, MSG_ASYNC_DROP, WPARAM(TAsyncOperationThread(Sender).DropData), Ord(True));
end;

procedure TFormMain.OnAsyncDropFailed(Sender: TObject);
begin
  PostMessage(Handle, MSG_ASYNC_DROP, WPARAM(TAsyncOperationThread(Sender).DropData), Ord(False));
end;

procedure TFormMain.OnAsyncThreadTerminate(Sender: TObject);
begin
  if (InterlockedDecrement(FAsyncPending) = 0) then
    ButtonAsyncAvailable.Action := ActionAsyncAvailable;
end;

procedure TFormMain.OnDragEnter(Sender: TObject; ShiftState: TShiftState;
  APoint: TPoint; var Effect: Integer);
begin
  Clear;
  StatusBar1.SimpleText := 'Drag detected - Drop to analyze the drop source';

  UpdateAsyncAbility(TCustomDropTarget(Sender).DataObject);
end;

procedure TFormMain.OnDragLeave(Sender: TObject);
begin
  StatusBar1.SimpleText := 'Drop cancelled';
  Init;
end;

procedure TFormMain.OnDragOver(Sender: TObject; ShiftState: TShiftState;
  APoint: TPoint; var Effect: Integer);
begin
  // Prefer the copy drop effect
  if (Effect and DROPEFFECT_COPY <> 0) then
    Effect := DROPEFFECT_COPY;
end;

function AspectsToString(Aspects: DWORD): string;
const
  AspectNames: array[0..3] of string =
    ('Content', 'Thumbnail', 'Icon', 'Print');
var
  Aspect: DWORD;
  AspectNum: integer;
begin
  AspectNum := 0;
  Aspect := $0001;
  Result := '';
  while (Aspects >= Aspect) do
  begin
    if (Aspects and Aspect <> 0) then
    begin
      if (Result <> '') then
        Result := Result+'+';
      Result := Result+AspectNames[AspectNum];
    end;
    inc(AspectNum);
    Aspect := Aspect shl 1;
  end;
end;

function MediaToString(Media: DWORD): string;
const
  MediaNames: array[0..7] of string =
    ('GlobalMem', 'File', 'IStream', 'IStorage', 'GDI', 'MetaFile', 'EnhMetaFile', 'Unknown');
var
  Medium: DWORD;
  MediumNum: integer;
begin
  Result := '';
  MediumNum := 0;
  Medium := $0001;
  while (Media >= Medium) and (MediumNum <= High(MediaNames)) do
  begin
    if (Media and Medium <> 0) then
    begin
      if (Result <> '') then
        Result := Result+', ';
      Result := Result+MediaNames[MediumNum];
    end;
    inc(MediumNum);
    Medium := Medium shl 1;
  end;
end;

procedure TFormMain.OnDrop(Sender: TObject; ShiftState: TShiftState;
  APoint: TPoint; var Effect: Integer);
var
  GetNum, GotNum: longInt;
  FormatEnumerator: IEnumFormatEtc;
  SourceFormatEtc: TFormatEtc;
  DropData: PDropData;
  Item: TListItem;
  s: string;
  Size: integer;
  Direction: Longint;
begin
  StatusBar1.SimpleText := 'Data dropped';
  ListViewDataFormats.Items.BeginUpdate;
  try
    ListViewDataFormats.Items.Clear;

    // Save a reference to the data object for later use. We need it to fetch
    // data from the drop source when the user selects an item from the list.
    FDataObject := TCustomDropTarget(Sender).DataObject;

    UpdateAsyncAbility(FDataObject);

    if (ActionDirTarget.Checked) then
      Direction := DATADIR_GET
    else
      Direction := DATADIR_SET;

    // Bail out if the drop source won't allow us to enumerate the formats.
    if (FDataObject.EnumFormatEtc(Direction, FormatEnumerator) <> S_OK) or
      (FormatEnumerator.Reset <> S_OK) then
    begin
      Clear;
      exit;
    end;

    GetNum := 1; // Get one format at a time.

    // Enumerate all data formats offered by the drop source.
    while (FormatEnumerator.Next(GetNum, SourceFormatEtc, @GotNum) = S_OK) and
      (GetNum = GotNum) do
    begin
      Item := ListViewDataFormats.Items.Add;
      Item.ImageIndex := -1;

      // Format ID
      Item.Caption := IntToStr(SourceFormatEtc.cfFormat);

      // Format name
      Item.SubItems.Add(GetClipboardFormatNameStr(SourceFormatEtc.cfFormat));

      // Aspect names
      Item.SubItems.Add(AspectsToString(SourceFormatEtc.dwAspect));

      // Media names
      Item.SubItems.Add(MediaToString(SourceFormatEtc.tymed));

      // Data size
      Size := GetDataSize(SourceFormatEtc);
      if (Size > 0) then
        s := Format('%.0n', [Int(Size)])
      else
        s := '-';
      Item.SubItems.Add(s);

      // Save a copy of the format descriptor in the listview
      New(DropData);
      Item.Data := DropData;

      DropData^ := Default(TDropData);
      DropData.FormatEtc := SourceFormatEtc;

      // Verify media
      DropData.ActualTymed := VerifyMedia(DropData.FormatEtc);
      if (DropData.ActualTymed <> DropData.FormatEtc.tymed) then
        Item.ImageIndex := ImageIndexWarning;

      if (ActionPrefetch.Checked) then
      begin
        if PrefetchDropData(DropData^) then
          Item.ImageIndex := ImageIndexPending;
      end;
    end;
  finally
    ListViewDataFormats.Items.EndUpdate;
  end;

  // Force resize of listview to get rid of horizontal scrollbar
  ListViewDataFormats.Width := ListViewDataFormats.Width - 1;

  // Reject the drop so the drop source doesn't think we actually did something
  // usefull with it.
  // This is important when moving data or when dropping from the recycle bin;
  // If we do not reject the drop, the source will assume that it is safe to
  // delete the source data. See also "Optimized move".

  // Note: The code below has been disabled as we now handle the above scenario
  // as an optimized move (OptimizedMove has been set to True) and break our
  // (that is, the drop target's) part of the optmized move contract by not
  // actually deleting the dropped data.
  (*
  if ((Effect and not(DROPEFFECT_SCROLL)) = DROPEFFECT_MOVE) then
    Effect := DROPEFFECT_NONE
  else
    Effect := DROPEFFECT_COPY;
  *)
end;

function TFormMain.VerifyMedia(var AFormatEtc: TFormatEtc): longInt;
var
  Mask: longInt;
  FormatEtc: TFormatEtc;
  RetryFormatEtc: TFormatEtc;
  Medium: TStgMedium;
  Res: HRESULT;
begin
  // Some drop sources lie about which media they support (e.g. Mozilla Thunderbird).
  // Here we try them all through IDataObject.GetData.
  Mask := $0000001;
  FormatEtc := AFormatEtc;
  Result := 0;
  while (Mask <> 0) and (Mask <= TYMED_ENHMF) do
  begin
    FormatEtc.tymed := Mask;
    FillChar(Medium, SizeOf(Medium), 0);

    Res := FDataObject.GetData(FormatEtc, Medium);

    // Retry with lindex=0 if lindex=-1 failed
    if ((Res = DV_E_LINDEX) or (Res = E_INVALIDARG)) and (FormatEtc.lindex = -1) then
    begin
      RetryFormatEtc := AFormatEtc;
      RetryFormatEtc.lindex := 0;
      Res := DataObject.GetData(RetryFormatEtc, Medium);
    end;

    //if (Succeeded(FDataObject.QueryGetData(FormatEtc))) then
    if (Succeeded(Res)) then
    begin
      // Mozilla Thunderbird speciality:
      // If we ask Thunderbird for FileContents on a TYMED_HGLOBAL then it
      // returns it on a TYMED_ISTREAM. If we ask for FileContents on a
      // TYMED_ISTREAM then it fails.
      // Wow!
      Result := Result or Medium.tymed;
      ReleaseStgMedium(Medium);
    end;
    Mask := Mask shl 1;
  end;
end;

procedure TFormMain.ActionDummyExecute(Sender: TObject);
begin
  // Empty on purpose
end;

procedure TFormMain.ActionClearExecute(Sender: TObject);
begin
  Reset;
end;

procedure TFormMain.ActionPasteExecute(Sender: TObject);
begin
  CheckAsyncPending;
  FDropTarget.PasteFromClipboard;
end;

procedure TFormMain.ActionPasteUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := (FDropTarget.CanPasteFromClipboard);
end;

procedure TFormMain.ActionSaveExecute(Sender: TObject);
var
  Strings: TStrings;
  i: integer;
  ClipFormat: TRawClipboardFormat;
  DropData: PDropData;
  Name, Media: string;
const
  sFormat =     'Name  : %s'+#13#10+
                'ID    : %d'+#13#10+
                'Medium: %s'+#13#10+
                'Aspect: %s'+#13#10+
                'Index : %d'+#13#10+
                'Size  : %s'+#13#10+
                '============================================================================';
  sDataHeader = 'Offset    Bytes                                             ASCII'+#13#10+
                '--------  ------------------------------------------------  ----------------';
  sSeparator =  '============================================================================';
  sActualMedia = ' (actual: %s)';
begin
  CheckAsyncPending;

  if (not SaveDialog1.Execute) then
    exit;

  Strings := TStringList.Create;
  try
    for i := 0 to ListViewDataFormats.Items.Count-1 do
      if (ListViewDataFormats.Items[i].Data <> nil) then
      begin
        DropData := PDropData(ListViewDataFormats.Items[i].Data);
        ClipFormat := TRawClipboardFormat.CreateFormatEtc(DropData.FormatEtc);
        try
          ClipFormat.GetData(DataObject);

          Name := ClipFormat.ClipboardFormatName;
          if (Name = '') then
            Name := GetClipboardFormatNameStr(ClipFormat.ClipboardFormat);

          Media := MediaToString(DropData.FormatEtc.tymed);
          if (DropData.FormatEtc.tymed <> DropData.ActualTymed) then
            Media := Media + Format(sActualMedia, [MediaToString(DropData.ActualTymed)]);

          Strings.Add(Format(sFormat,
            [Name, ClipFormat.ClipboardFormat,
            Media,
            AspectsToString(DropData.FormatEtc.dwAspect),
            DropData.FormatEtc.lindex,
            ListViewDataFormats.Items[i].SubItems[3]]));

          Strings.Add(sDataHeader);
          Strings.Add(DataToHexDump(GetDropDataAsString(DropData^)));
          Strings.Add(sSeparator);
        finally
          ClipFormat.Free;
        end;
    end;
    Strings.SaveToFile(SaveDialog1.FileName);
  finally
    Strings.Free;
  end;
end;

procedure TFormMain.ActionSaveUpdate(Sender: TObject);
begin
  TAction(Sender).Enabled := (ListViewDataFormats.Items.Count > 0);
end;

procedure TFormMain.CheckAsyncPending;
begin
  if (FAsyncPending > 0) then
  begin
    ShowMessage('Async transfer in progress.'#13'Please wait for transfer to complete');
    Abort;
  end;
end;

procedure TFormMain.Clear;
begin
  EditHexView.Text := '';
  ListViewDataFormats.Items.Clear;
  FDataObject := nil;
end;

procedure TFormMain.Reset;
begin
  CheckAsyncPending;
  Clear;
  Init;
end;

procedure TFormMain.UpdateAsyncAbility(const DataObject: IDataObject);
begin
  if (DataObject = nil) then
    ButtonAsyncAvailable.Action := ActionAsyncInfo
  else
  if (Supports(DataObject, IAsyncOperation2)) then
    ButtonAsyncAvailable.Action := ActionAsyncAvailable
  else
    ButtonAsyncAvailable.Action := ActionAsyncNotAvailable;
end;

function TFormMain.DataToHexDump(const Data: AnsiString): string;
var
  i: integer;
  Offset: integer;
  Hex: string;
  ASCII: string;
  LineLength: integer;
  Size: integer;
begin
  Result := '';
  LineLength := 0;
  Hex := '';
  ASCII := '';
  Offset := 0;

  Size := Length(Data);
  if (Size > MAX_DATA) then
    Size := MAX_DATA;

  for i := 0 to Size-1 do
  begin
    Hex := Hex+IntToHex(ord(Data[i+1]), 2)+' ';
    if (Data[i+1] in [' '..#$7F]) then
      ASCII := ASCII+Char(Data[i+1])
    else
      ASCII := ASCII+'.';
    inc(LineLength);
    if (LineLength = 16) or (i = Length(Data)-1) then
    begin
      Result := Result+Format('%.8x  %-48.48s  %-16.16s'+#13+#10, [Offset, Hex, ASCII]);
      inc(Offset, LineLength);
      LineLength := 0;
      Hex := '';
      ASCII := '';
    end;
  end;
end;

procedure TFormMain.Error(const Msg: string);
begin
  IntroView.Hide;
  EditHexView.Hide;
  LabelError.Caption := Msg;
  PanelError.Show;
end;

procedure TFormMain.ListViewDataFormatsAdvancedCustomDrawSubItem(
  Sender: TCustomListView; Item: TListItem; SubItem: Integer;
  State: TCustomDrawState; Stage: TCustomDrawStage; var DefaultDraw: Boolean);
var
  Mask: longInt;
  Tymed: longInt;
  r, TextRect: TRect;
  s: string;
  Canvas: TControlCanvas;
begin
  case SubItem of
    3:
      if (Stage = cdPrePaint) and
        (PDropData(Item.Data).FormatEtc.tymed <> PDropData(Item.Data).ActualTymed) then
      begin
        // Draw media names:
        // - Black: Reported by EnumFormatEtc and supported by GetData.
        // - Red: Reported by EnumFormatEtc but not supported by GetData.
        // - Green: Not reported by EnumFormatEtc but supported by GetData.
        ListView_GetSubItemRect(Sender.Handle, Item.Index, SubItem, LVIR_BOUNDS, @TextRect);

        // Work around for long standing bug in TListView ownerdraw:
        // Because ListView.Canvas.Font.OnChange is rerouted by the listview,
        // changes to the font does not update the GDI object. Same goes for the
        // brush.
        Canvas := TControlCanvas.Create;
        try
          Canvas.Control := Sender;

          Canvas.Font.Assign(Sender.Canvas.Font);
          Canvas.Brush.Assign(Sender.Canvas.Brush);

          if Item.Selected then
          begin
            if Sender.Focused then
              Canvas.Brush.Color := clHighlight
            else
              Canvas.Brush.Color := clBtnFace;
          end else
            Canvas.Brush.Color := clWindow;
          Canvas.Brush.Style := bsSolid;

          Canvas.FillRect(TextRect);

          InflateRect(TextRect, -1, -1);
          inc(TextRect.Left, 5);

          Mask := $0001;
          Tymed := PDropData(Item.Data).FormatEtc.tymed or PDropData(Item.Data).ActualTymed;
          while (Mask <= Tymed) do
          begin
            if ((Tymed and Mask) <> 0) then
            begin
              if ((PDropData(Item.Data).ActualTymed and Mask) = 0) then
              begin
                Canvas.Font.Color := clRed;
                Canvas.Font.Style := [fsStrikeOut];
              end else
              if ((PDropData(Item.Data).FormatEtc.tymed and Mask) = 0) then
              begin
                Canvas.Font.Color := clGreen;
                Canvas.Font.Style := [];
              end else
              begin
                if Item.Selected and Sender.Focused then
                  Canvas.Brush.Color := clHighlightText
                else
                  Canvas.Font.Color := clWindowText;
                Canvas.Font.Style := [];
              end;

              s := MediaToString(Mask) + ' ';

              r := TextRect;

              DrawText(Canvas.Handle, PChar(s), Length(s), r, DT_CALCRECT or DT_LEFT or DT_NOPREFIX or DT_SINGLELINE or DT_VCENTER);
              IntersectRect(r, r, TextRect);
              TextRect.Left := r.Right;
              DrawText(Canvas.Handle, PChar(s), Length(s)-1, r, DT_LEFT or DT_NOPREFIX or DT_SINGLELINE or DT_VCENTER);
            end;

            Mask := Mask shl 1;
          end;
        finally
          Canvas.Free;
        end;

        DefaultDraw := False;
      end;
  end;
end;

procedure TFormMain.ListViewDataFormatsDeletion(Sender: TObject;
  Item: TListItem);
begin
  if (Item.Data <> nil) then
  begin
    Finalize(PDropData(Item.Data)^);
    Dispose(Item.Data);
    Item.Data := nil;
  end;
end;

procedure TFormMain.ListViewDataFormatsSelectItem(Sender: TObject;
  Item: TListItem; Selected: Boolean);
var
  DropData: PDropData;
begin
  // Work around for RichEdit failing to change font
  PanelError.Hide;
  IntroView.Hide;
  (* This works on newer version of Delphi:
  EditHexView.PlainText := True;
  // Force RichEdit to recreate window handle in PlainText mode.
  EditHexView.Perform(CM_RECREATEWND, 0, 0);
  *)
  if (Selected) and (Item.Data <> nil) then
  begin
    DropData := PDropData(Item.Data);
    try
      if (not DropData.IsPending) and (PrefetchDropData(DropData^)) then
      begin
        // We just initiated an async transfer; Let user know
        Item.ImageIndex := ImageIndexPending;
        ListViewDataFormats.Update;
      end;

      if (not DropData.IsPending) then
      begin
        EditHexView.Text := DataToHexDump(GetDropDataAsString(DropData^));

        if (DropData.HasData) then
          EditHexView.Show
        else
          Error('Failed to retrieve data from Drop Source');
      end else
      begin
        EditHexView.Text := 'Async transfer in progress - Please wait';
        EditHexView.Show;
        EditHexView.Update;
      end;
    except
      on E: Exception do
        Error(E.Message);
    end;
  end else
    EditHexView.Hide;
end;

function TFormMain.PrefetchDropData(var DropData: TDropData): boolean;
var
  IsAsync: boolean;

  function DoGetDropData(const FormatEtc: TFormatEtc): boolean;
  var
    AsyncOperation: IAsyncOperation2;
    DoAsync: Bool;
    Thread: TAsyncOperationThread;
  begin
    // Query source for async mode if requested
    if (not ActionAsync.Checked) or
      (not Supports(DataObject, IAsyncOperation2, AsyncOperation)) or
      (Failed(AsyncOperation.GetAsyncMode(DoAsync))) then
      DoAsync := False;

    // Start an async data transfer.
    if (DoAsync) and
      // Notify drop source that an async data transfer is starting.
      Succeeded(AsyncOperation.StartOperation(nil)) then
    begin
      try

        // Create the data transfer thread and launch it.
        Thread := TAsyncOperationThread.Create(DataObject, @DropData, FormatEtc);
        try

          Thread.OnTerminate := OnAsyncThreadTerminate;
          Thread.OnDrop := OnAsyncDrop;
          Thread.OnDropFailed := OnAsyncDropFailed;

          if (InterlockedIncrement(FAsyncPending) = 1) then
            ButtonAsyncAvailable.Action := ActionAsyncActive;

          Thread.Start;

          IsAsync := True;

        except
          Thread.Free;
        end;

        Result := True;

      except
        // Notify drop source that async data transfer failed
        AsyncOperation.EndOperation(E_UNEXPECTED, nil, DROPEFFECT_NONE);
        Result := False;
      end;

    end else
      Result := GetDropData(DataObject, FormatEtc, DropData);
  end;

var
  FormatEtc: TFormatEtc;
begin
  IsAsync := False;

  if (not DropData.HasFetched) then
  begin
    DropData.HasFetched := True;
    // We have to ask for both the reported media and the actual media in
    // order to work with Mozilla Thunderbird 3rc2.
    FormatEtc := DropData.FormatEtc;
    FormatEtc.tymed := FormatEtc.tymed or DropData.ActualTymed;

    if (not DoGetDropData(FormatEtc)) and (FormatEtc.lindex = -1) then
    begin
      // Retry with FormatEtc.lindex=0 if FormatEtc.lindex=-1 failed
      FormatEtc.lindex := 0;
      DoGetDropData(FormatEtc);
    end;
  end;

  Result := IsAsync;
end;

function TFormMain.GetDropDataAsString(var DropData: TDropData): AnsiString;
begin
  PrefetchDropData(DropData);
  Result := DropData.Data;
end;

procedure TFormMain.LoadRTF(const s: string);
var
  Stream: TStream;
begin
  Stream := TStringStream.Create(s);
  try
    IntroView.Lines.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TFormMain.MsgAsyncDrop(var Msg: TMessage);
var
  DropData: PDropData;
  Success: boolean;
  i: integer;
begin
  DropData := PDropData(Msg.WParam);
  Success := boolean(Msg.LParam);

  // Find the listview node that corresponds to the DropData
  for i := 0 to ListViewDataFormats.Items.Count-1 do
    if (ListViewDataFormats.Items[i].Data = DropData) then
    begin
      if (Success) then
      begin
        if (ListViewDataFormats.Items[i].Selected) then
          ListViewDataFormatsSelectItem(ListViewDataFormats, ListViewDataFormats.Items[i], True);
        ListViewDataFormats.Items[i].ImageIndex := -1;
      end else
        ListViewDataFormats.Items[i].ImageIndex := ImageIndexError;

      break;
    end;
end;

procedure TFormMain.Init;
begin
  PanelError.Hide;
  EditHexView.Hide;
  IntroView.Show;
  UpdateAsyncAbility(nil);
end;

end.

