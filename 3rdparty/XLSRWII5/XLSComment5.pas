unit XLSComment5;

{-
********************************************************************************
******* XLSReadWriteII V6.00                                             *******
*******                                                                  *******
******* Copyright(C) 1999,2017 Lars Arvidsson, Axolot Data               *******
*******                                                                  *******
******* email: components@axolot.com                                     *******
******* URL:   http://www.axolot.com                                     *******
********************************************************************************
** Users of the XLSReadWriteII component must accept the following            **
** disclaimer of warranty:                                                    **
**                                                                            **
** XLSReadWriteII is supplied as is. The author disclaims all warranties,     **
** expressedor implied, including, without limitation, the warranties of      **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSReadWriteII.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}
{$I XLSRWII.inc}

interface

uses Classes, SysUtils, Contnrs,
     Xc12Utils5, Xc12Manager5, Xc12DataComments5, Xc12DataStyleSheet5,
     XLSUtils5, XLSColumn5, XLSRow5, XLSTools5;

type TXLSComments = class;

     TXLSComment = class(TObject)
private
     function  GetAuthor: AxUCString;
     function  GetCol: integer;
     function  GetPlainText: AxUCString;
     function  GetRow: integer;
     procedure SetAuthor(const Value: AxUCString);
     procedure SetCol(const Value: integer);
     procedure SetPlainText(const Value: AxUCString);
     procedure SetRow(const Value: integer);
     function  GetVisible: boolean;
     function  GetPlainTextNoAuthor: AxUCString;
protected
     // For XSS
     FAutoVisible: boolean;

     FOwner: TXLSComments;
     FXc12Comment: TXc12Comment;
public
     constructor Create(AXc12Comment: TXc12Comment; AOwner: TXLSComments);
     destructor Destroy; override;

     property Col: integer read GetCol write SetCol;
     property Row: integer read GetRow write SetRow;

     property Author: AxUCString read GetAuthor write SetAuthor;
     property PlainText: AxUCString read GetPlainText write SetPlainText;
     property PlainTextNoAuthor: AxUCString read GetPlainTextNoAuthor;
     property Visible: boolean read GetVisible;
     property AutoVisible: boolean read FAutoVisible write FAutoVisible;

     property Xc12Comment: TXc12Comment read FXc12Comment;
     end;

     TXLSComments = class(TObject)
private
     function  GetItems(Index: integer): TXLSComment;
     function  GetAsPlainText(const ACol, ARow: integer): AxUCString;
     function  GetAsSimpleTags(const ACol, ARow: integer): AxUCString;
     procedure SetAsPlainText(const ACol, ARow: integer; const Value: AxUCString);
     procedure SetAsSimpleTags(const ACol, ARow: integer; const Value: AxUCString);
protected
     FManager     : TXc12Manager;
     FStyleSheet  : TXc12DataStyleSheet;
     FAuthor      : AxUCString;
     FItems       : TObjectList;
     FXc12Comments: TXc12Comments;

     function  Add(AXc12Comment: TXc12Comment): TXLSComment; overload;
     function  AddSimpleTags(const ACol,ARow: integer; const AText: AxUCString): TXLSComment;
public
     constructor Create(AManager: TXc12Manager; AXc12Comments: TXc12Comments);
     destructor Destroy; override;

     procedure Assign(AComments: TXLSComments);

     procedure Clear;

     procedure ColWidthChanged(const ACol1,ACol2: integer);

     function  Find(const ACol,ARow: integer): TXLSComment;
     function  FindIndex(const ACol,ARow: integer): integer;
     function  Count: integer;
     function  Add: TXLSComment; overload;
     function  Add(const ACol,ARow: integer; const AText: AxUCString): TXLSComment; overload;
     function  Add(const ACol,ARow: integer; const AAuthor,AText: AxUCString): TXLSComment; overload;

     procedure Delete(const AIndex: integer); overload;
     procedure Delete(const ACol,ARow: integer); overload;
     procedure Delete(const ACol1,ARow1,ACol2,ARow2: integer); overload;

     property Items[Index: integer]: TXLSComment read GetItems; default;
     property DefaultAuthor: AxUCString read FAuthor write FAuthor;
     property AsPlainText[const ACol,ARow: integer]: AxUCString read GetAsPlainText write SetAsPlainText;
     property AsSimpleTags[const ACol,ARow: integer]: AxUCString read GetAsSimpleTags write SetAsSimpleTags;
     end;

implementation

{ TXLSComments }

function TXLSComments.Add(AXc12Comment: TXc12Comment): TXLSComment;
begin
  Result := TXLSComment.Create(AXc12Comment,Self);
  FItems.Add(Result);
end;

function TXLSComments.Add: TXLSComment;
begin
  Result := Add(FXc12Comments.Add);
end;

function TXLSComments.Add(const ACol, ARow: integer; const AText: AxUCString): TXLSComment;
var
  i: integer;
  FR: PXc12FontRun;
begin
  i := FindIndex(ACol,ARow);
  if i >= 0 then
    Delete(i);

  Result := Add;
  Result.FXc12Comment.SetDefault;
  Result.Col := ACol;
  Result.Row := ARow;
  Result.Author := FAuthor;
  if FAuthor <> '' then
    Result.PlainText := FAuthor + ':' + #13 + AText
  else
    Result.PlainText := AText;

  Result.FXc12Comment.Col1 := 1;
  Result.FXc12Comment.Col2 := 4;
  Result.FXc12Comment.Row1 := 1;
  Result.FXc12Comment.Row2 := 6;

  i := 0;
  if FAuthor <> '' then begin
    FR := Result.FXc12Comment.AddFontRun;
    FR.Index := 0;
    FR.Font := FStyleSheet.Fonts.MakeCommentFont(True);
    Inc(i,Length(FAuthor) + 1);  // 1 for ':'
  end;
  FR := Result.FXc12Comment.AddFontRun;
  FR.Index := i;
  FR.Font := FStyleSheet.Fonts.MakeCommentFont(False);

  FStyleSheet.XFEditor.UseFonts(Result.FXc12Comment.FontRuns);
end;

function TXLSComments.Add(const ACol, ARow: integer; const AAuthor, AText: AxUCString): TXLSComment;
begin
  FAuthor := AAuthor;
  Result := Add(ACol,ARow,AText);
end;

function TXLSComments.AddSimpleTags(const ACol, ARow: integer; const AText: AxUCString): TXLSComment;
var
  i: integer;
  S: AxUCString;
  Builder: TXLSFontRunBuilder;
begin
  i := FindIndex(ACol,ARow);
  if i >= 0 then
    Delete(i);

  Result := Add;
  Result.FXc12Comment.SetDefault;
  Result.Col := ACol;
  Result.Row := ARow;
  Result.Author := FAuthor;

  Builder := TXLSFontRunBuilder.Create(FManager);
  try
    if FAuthor <> '' then
      S := '<b>' + FAuthor + ':' + #13 + '</b>' + AText
    else
      S := AText;

    Builder.FromSimpleTags(S);
    Result.PlainText := FAuthor + ':' + #13 + AText;

    Result.FXc12Comment.AssignFontRuns(Builder.FontRuns);
  finally
    Builder.Free;
  end;

end;

procedure TXLSComments.Assign(AComments: TXLSComments);
var
  i: integer;
begin
  for i := 0 to AComments.Count - 1 do
    Add.FXc12Comment.Assign(AComments[i].FXc12Comment);
end;

procedure TXLSComments.Clear;
begin
  FItems.Clear;
  FXc12Comments.Clear;
end;

procedure TXLSComments.ColWidthChanged(const ACol1,ACol2: integer);
begin

end;

function TXLSComments.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TXLSComments.Create(AManager: TXc12Manager; AXc12Comments: TXc12Comments);
var
  i: integer;
begin
  FManager := AManager;
  FStyleSheet := FManager.StyleSheet;
  FXc12Comments := AXc12Comments;

  FItems := TObjectList.Create;

  for i := 0 to FXc12Comments.Count - 1 do
    Add(FXc12Comments[i]);
end;

procedure TXLSComments.Delete(const ACol, ARow: integer);
var
  i: integer;
begin
  i := FindIndex(ACol,ARow);
  if i >= 0 then
    Delete(i);
end;

procedure TXLSComments.Delete(const AIndex: integer);
begin
  FItems.Delete(AIndex);
  FXc12Comments.Delete(AIndex);
end;

procedure TXLSComments.Delete(const ACol1, ARow1, ACol2, ARow2: integer);
var
  i: integer;
begin
  i := 0;
  while i < FItems.Count do begin
    if (Items[i].Col >= ACol1) and (Items[i].Col <= ACol2) and (Items[i].Row >= ARow1) and (Items[i].Row <= ARow2) then begin
      FItems.Delete(i);
      FXc12Comments.Delete(i);
    end
    else
      Inc(i);
  end;
end;

destructor TXLSComments.Destroy;
begin
  FItems.Free;
  inherited;
end;

function TXLSComments.Find(const ACol, ARow: integer): TXLSComment;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if (Items[i].Col = ACol) and (Items[i].Row = ARow) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXLSComments.FindIndex(const ACol, ARow: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if (Items[Result].Col = ACol) and (Items[Result].Row = ARow) then
      Exit;
  end;
  Result := -1;
end;

function TXLSComments.GetAsPlainText(const ACol, ARow: integer): AxUCString;
var
  C: TXLSComment;
begin
  C := Find(ACol,ARow);
  if C <> Nil then
    Result := C.PlainText
  else
    Result := '';
end;

function TXLSComments.GetAsSimpleTags(const ACol, ARow: integer): AxUCString;
begin
  Result := '';
end;

function TXLSComments.GetItems(Index: integer): TXLSComment;
begin
  Result := TXLSComment(FItems.Items[Index]);
end;

procedure TXLSComments.SetAsPlainText(const ACol, ARow: integer; const Value: AxUCString);
begin
  Add(ACol,ARow,Value);
end;

procedure TXLSComments.SetAsSimpleTags(const ACol, ARow: integer; const Value: AxUCString);
begin
  AddSimpleTags(ACol,ARow,Value);
end;

{ TXLSComment }

constructor TXLSComment.Create(AXc12Comment: TXc12Comment; AOwner: TXLSComments);
begin
  FXc12Comment := AXc12Comment;
  FOwner := AOwner;
end;

destructor TXLSComment.Destroy;
begin
  inherited;
end;

function TXLSComment.GetAuthor: AxUCString;
begin
  if FXc12Comment.Author <> Nil then
    Result := FXc12Comment.Author.Name
  else
    Result := '';
end;

function TXLSComment.GetCol: integer;
begin
  Result := FXc12Comment.Col;
end;

function TXLSComment.GetPlainText: AxUCString;
begin
  Result := FXc12Comment.Text;
end;

function TXLSComment.GetPlainTextNoAuthor: AxUCString;
var
  S: AxUCString;
begin
  S := Author + ':';
  Result := PlainText;
  if S = Copy(Result,1,Length(S)) then
    Result := Copy(Result,Length(S) + 1,MAXINT);
end;

function TXLSComment.GetRow: integer;
begin
  Result := FXc12Comment.Row;
end;

function TXLSComment.GetVisible: boolean;
begin
  Result := True;
end;

procedure TXLSComment.SetAuthor(const Value: AxUCString);
var
  A: TXc12Author;
begin
  A := FOwner.FXc12Comments.Authors.Find(Value);
  if A = Nil then
    A := FOwner.FXc12Comments.Authors.Add(Value);
  FXc12Comment.Author := A;
end;

procedure TXLSComment.SetCol(const Value: integer);
begin
  FXc12Comment.Col := Value;
end;

procedure TXLSComment.SetPlainText(const Value: AxUCString);
begin
  FXc12Comment.ClearFontRuns;
  FXc12Comment.Text := Value;
end;

procedure TXLSComment.SetRow(const Value: integer);
begin
  FXc12Comment.Row := Value;
end;

end.
