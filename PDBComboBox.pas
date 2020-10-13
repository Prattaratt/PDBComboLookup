unit PDBComboBox;

interface

uses
  System.SysUtils, System.Classes,
  Vcl.Controls, Vcl.StdCtrls,
  Data.DB,DBCtrls,
  WinAPI.Messages;

type
  TNotInListEvent = procedure (Sender:TComponent; Data:String) of object;
  TPDBComboBox = class(TCustomComboBox)
  private
    { Private declarations }
    fListSource:TDataSource;
    fListField:TField;
    fDataLink:TFieldDataLink;
    fNewData:Boolean;
    fNotInList:TNotInListEvent;
//    fOnKeyPress:TKeyPressEvent;
    function InList(S:String):Boolean;
    function getNewData:Boolean;
    procedure setNewData(Value:Boolean);
    procedure LoadList;
    procedure DataLinkActiveChange(Sender:Tobject);
//    procedure Refreshlist;
  protected
    { Protected declarations }
    procedure KeyPress(var Key:Char);override;
    procedure Select;override;
    procedure DropDown; override;
    procedure DoExit;override;
    procedure DoNotInList;virtual;
    property NewData:boolean read getNewData write setNewData;

  public
    { Public declarations }
    procedure Loaded;override;
    Constructor Create(AOwner:TComponent);
    Destructor Destroy;
  published
    { Published declarations }
    property Align;
    property AutoComplete default True;
    property AutoCompleteDelay default 500;
    property AutoDropDown default False;
    property AutoCloseUp default False;
    property BevelEdges;
    property BevelInner;
    property BevelKind default bkNone;
    property BevelOuter;
    property Style; {Must be published before Items}
    property Anchors;
    property BiDiMode;
    property CharCase;
    property Color;
    property Constraints;
    property Ctl3D;
    property DoubleBuffered;
    property DragCursor;
    property DragKind;
    property DragMode;
    property DropDownCount;
    property Enabled;
    property Font;
    property ImeMode;
    property ImeName;
    property ItemHeight;
    property ItemIndex default -1;
    property MaxLength;
    property ParentBiDiMode;
    property ParentColor;
    property ParentCtl3D;
    property ParentDoubleBuffered;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
    property ShowHint;
    property Sorted;
    property TabOrder;
    property TabStop;
    property Text;
    property TextHint;
    property Touch;
    property Visible;
    property OnChange;
    property OnClick;
    property OnCloseUp;
    property OnContextPopup;
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnDrawItem;
    property OnDropDown;
    property OnEndDock;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnGesture;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMeasureItem;
    property OnMouseEnter;
    property OnMouseLeave;
    property OnSelect;
    property OnStartDock;
    property OnStartDrag;
    property Items; { Must be published after OnMeasureItem }

    property ListSource:TDataSource  read fListSource write fListSource;
    property ListField:TField read fListField write fListField;
    property OnNotInList:TNotInListEvent read fNotInList write fNotInList;

  end;

  procedure Register;

implementation

uses StrUtils, Character;
procedure Register;
begin
  RegisterComponents('Pratt Software', [TPDBComboBox]);
end;

{ TPDBComboBox }

procedure TPDBComboBox.KeyPress(var Key:Char);
var S:String;
    K:Char;
    Found:Boolean;
    I:Integer;
begin
  K:=Key;
  inherited KeyPress(Key);
  if (TCharacter.IsControl(K)) and (K <> #13)then
    exit;
  if Not DroppedDown then begin //Not dropped branch
    if K = #13 then begin // Handle end of entry
      if NewData then
        DoNotInList
    end
    else begin  // regular char entry
      S:= LeftStr(Text,Length(Text)-Self.SelLength);
      if (Length(S) = 0) OR (K <> S[Length(s)]) then
         S:=S+K;
      if Inlist(S) then begin
        NewData:=false;
        Select;
      end
      else begin
        NewData:=True;
        Text:=S;
        SelStart := Length(Text);
      end;
    end;
  end
  else begin  // Dropped down branch
    if NewData then
      Exit;
    S := self.Text;
    S :=  LeftStr(S, self.SelStart) + Char(K);
    if not InList(S) then begin
      NewData := true;
     end
     else begin
       Found:=False;
       I:=0;
       repeat
         if StartsText(LeftStr(Text, SelStart), Text) then begin;
         end;
       until Found or (I= Items.Count-1);
     end;
  end;
  Key:=#0
end;

function TPDBComboBox.InList(S: String): Boolean;
var I:Integer;
begin
  result := false;
  for I := 0 to Items.Count-1 do
     if StartsText(S, Items.Strings[I]) then begin
       Result := true;
       exit;
     end;
end;

procedure TPDBComboBox.Loaded;
begin
  inherited;
  Text:='';
  LoadList;
end;

procedure TPDBComboBox.LoadList;
begin
  if Items.Count<>0 then
    Items.Clear;
  if (fListSource = Nil) OR (NOT Assigned(ListSource.DataSet))then
    exit;
  if NOT fListSource.DataSet.Active then
    fListSource.DataSet.Active:=True;
  fListSource.DataSet.First;
  while NOT fListSource.DataSet.Eof do Begin
    self.Items.Add(fListField.AsString);
    fListSource.DataSet.Next;
  End;
end;

procedure TPDBComboBox.Select;
var opts:TLocateOptions;
begin
  inherited;
  if InList(self.Text) then begin
    NewData:=false;
    Opts:=[loCaseInsensitive];
    ListSource.DataSet.Locate(ListField.FieldName, Text, opts);
  end
  else begin
    NewData := true;
  end;
end;

constructor TPDBComboBox.Create(AOwner: TComponent);
begin
  inherited;
  fDataLink.Create;
  fDataLink.Control:=self;
  fDataLink.OnActiveChange := DataLinkActiveChange;
  fDataLink.DataSource:=fListSource;
  fDataLink.FieldName:=fListField.FieldName;
end;

procedure TPDBComboBox.DataLinkActiveChange(Sender:Tobject);
begin
  LoadList;
end;

destructor TPDBComboBox.Destroy;
begin
  fListSource.DataSet.Close;
  fDataLink.Free;
  fDataLink:=nil;
  inherited;
end;

procedure TPDBComboBox.DoExit;
begin
  inherited;
  if NewData then
    DoNotInList;
end;

procedure TPDBComboBox.DoNotInList;
begin
    if Assigned(fNotInList) then
      fNotInList(self,self.Text);
end;

procedure TPDBComboBox.DropDown;
begin
  ListSource.DataSet.Filtered:=false;
  inherited;
  NewData:=false;
  if Items.Count <> ListSource.DataSet.RecordCount then
    LoadList;

end;

function TPDBComboBox.getNewData: Boolean;
begin
   result:=fNewData;
end;

procedure TPDBComboBox.setNewData(Value: Boolean);
begin
  if Value = fNewData then
    Exit;
  fNewData:=Value;
  if fNewData then
    ListSource.Dataset.Append;
end;

end.
