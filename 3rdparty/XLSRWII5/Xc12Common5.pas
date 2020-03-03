unit Xc12Common5;

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
     XLSUtils5;

type TXLSStyleObjectList = class;

     TXLSIndexObject = class(TObject)
protected
     FIndex: integer;
public
     property Index: integer read FIndex;
     end;

     TXLSStyleObject = class(TIndexObject)
{$ifdef DELPHI_2009_OR_LATER}
strict private
{$else}
private
{$endif}
     FRefCount: integer;
protected
     FHashvalid: boolean;
     FHash: longword;
     // (Default) object that not can be deleted.
     FLocked: boolean;

     procedure CalcHash; virtual; abstract;
protected
     FOwner: TXLSStyleObjectList;

     procedure IncRef;
     procedure DecRef;
     function  Equal(AItem: TXLSStyleObject): boolean; virtual; abstract;
     function  HashKey: longword; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
public
     constructor Create(AOwner: TXLSStyleObjectList);

     function  AsString(const AIndent: integer = 0): AxUCString; virtual;

//     procedure Use;

     property Owner: TXLSStyleObjectList read FOwner;
     property Locked: boolean read FLocked write FLocked;
     property RefCount: integer read FRefCount;
     end;

     TXLSIndexObjectList = class(TObjectList)
protected
     FIsDestroying: boolean;

     procedure Enumerate(AStartIndex: integer = 0);
     procedure Notify(Ptr: Pointer; Action: TListNotification); override;
public
     destructor Destroy; override;
     end;

     PPXLSHashItem = ^PXLSHashItem;
     PXLSHashItem = ^TXLSHashItem;
     TXLSHashItem = record
       Next: PXLSHashItem;
       Key: TXLSStyleObject;
     end;

     TXLSStyleObjectHash = class
private
     FBuckets: array of PXLSHashItem;
protected
     function  Find(AKey: TXLSStyleObject): PPXLSHashItem;
public
     constructor Create(const ASize: Cardinal = 256);
     destructor Destroy; override;

     procedure Add(AKey: TXLSStyleObject);
     procedure Clear;
     procedure Remove(AKey: TXLSStyleObject);
     end;

     TXLSStyleObjectList = class(TXLSIndexObjectList)
private
     FItemHash: TXLSStyleObjectHash;
     FItemHashValid: Boolean;

     procedure UpdateItemHash;
     function GetItems(Index: integer): TXLSStyleObject;
protected
     procedure Notify(Ptr: Pointer; Action: TListNotification); override;
public
     constructor Create(const AHashSize: integer = 256);
     destructor  Destroy; override;

     function  Find(AItem: TXLSStyleObject): TXLSStyleObject;

     procedure SetIsDefault; virtual; abstract;
     function  DefaultCount: integer;

     procedure Add(AObject: TXLSStyleObject);

     property StyleItems[Index: integer]: TXLSStyleObject read GetItems;
     end;

type TXc12Data = class(TObject)
protected
public
     end;

implementation

{ TXLSStyleObject }

function TXLSStyleObject.AsString(const AIndent: integer = 0): AxUCString;
begin
  Result := '';
end;

constructor TXLSStyleObject.Create(AOwner: TXLSStyleObjectList);
begin
  FOwner := AOwner;
end;

procedure TXLSStyleObject.DecRef;
begin
  Dec(FRefCount);
end;

function TXLSStyleObject.HashKey: longword;
begin
  if not FHashvalid then
    CalcHash;
  Result := FHash
end;

procedure TXLSStyleObject.IncRef;
begin
  Inc(FRefCount);
end;

//procedure TXLSStyleObject.Use;
//begin
//  IncRef;
//end;

{ TXLSStyleObjectList }

procedure TXLSStyleObjectList.Add(AObject: TXLSStyleObject);
begin
  inherited Add(AObject);
end;

constructor TXLSStyleObjectList.Create(const AHashSize: integer = 256);
begin
  inherited Create;
  FItemHash := TXLSStyleObjectHash.Create(AHashSize);
end;

function TXLSStyleObjectList.DefaultCount: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do begin
    if TXLSStyleObject(Items[i]).Locked then
      Inc(Result)
    else
      Exit;
  end;
end;

destructor TXLSStyleObjectList.Destroy;
begin
  FItemHash.Free;
  inherited;
end;

function TXLSStyleObjectList.Find(AItem: TXLSStyleObject): TXLSStyleObject;
var
  Hash: PXLSHashItem;
begin
  if not FItemHashValid then
    UpdateItemHash;

  Hash := FItemHash.Find(AItem)^;
  if Hash <> Nil then
    Result := Hash^.Key
  else
    Result := Nil;
end;

function TXLSStyleObjectList.GetItems(Index: integer): TXLSStyleObject;
begin
  Result := TXLSStyleObject(inherited Items[Index]);
end;

procedure TXLSStyleObjectList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  inherited;
  FItemHashValid := False;
end;

procedure TXLSStyleObjectList.UpdateItemHash;
var
  i: Integer;
  Key: TXLSStyleObject;
begin
  if FItemHashValid then
    Exit;

  FItemHash.Clear;
  for i := 0 to Count - 1 do begin
    Key := Get(i);
    if Key <> Nil then
      FItemHash.Add(Key);
  end;
  FItemHashValid := True;
end;

{ TXLSIndexObjectList }

destructor TXLSIndexObjectList.Destroy;
begin
  FIsDestroying := True;

  inherited;
end;

procedure TXLSIndexObjectList.Enumerate(AStartIndex: integer);
var
  i: integer;
begin
  if not FIsDestroying then begin
    for i := AStartIndex to Count - 1 do begin
      if Items[i] <> Nil then
        TXLSStyleObject(Items[i]).FIndex := i;
    end;
  end;
end;

procedure TXLSIndexObjectList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  inherited;

  case Action of
    lnAdded    : TXLSIndexObject(Ptr).FIndex := Count - 1;
    lnExtracted,
    lnDeleted  : Enumerate;
  end;
end;

{ TXLSStyleObjectHash }

procedure TXLSStyleObjectHash.Add(AKey: TXLSStyleObject);
var
  Hash: Integer;
  Bucket: PXLSHashItem;
begin
  Hash := AKey.HashKey mod Cardinal(Length(FBuckets));
  New(Bucket);
  Bucket^.Key := AKey;
  Bucket^.Next := FBuckets[Hash];
  FBuckets[Hash] := Bucket;
end;

procedure TXLSStyleObjectHash.Clear;
var
  i: Integer;
  P,N: PXLSHashItem;
begin
  for i := 0 to High(FBuckets) do begin
    P := FBuckets[i];
    while P <> nil do begin
      N := P^.Next;
      Dispose(P);
      P := N;
    end;
    FBuckets[i] := nil;
  end;
end;

constructor TXLSStyleObjectHash.Create(const ASize: Cardinal);
begin
  SetLength(FBuckets,ASize);
end;

destructor TXLSStyleObjectHash.Destroy;
begin
  Clear;
  inherited Destroy;
end;

function TXLSStyleObjectHash.Find(AKey: TXLSStyleObject): PPXLSHashItem;
var
  Hash: Integer;
begin
  Hash := AKey.HashKey mod Cardinal(Length(FBuckets));
  Result := @FBuckets[Hash];
  while Result^ <> nil do begin
    if Result^.Key.Equal(AKey) then
      Exit
    else
      Result := @Result^.Next;
  end;
end;

procedure TXLSStyleObjectHash.Remove(AKey: TXLSStyleObject);
var
  P: PXLSHashItem;
  Prev: PPXLSHashItem;
begin
  Prev := Find(AKey);
  P := Prev^;
  if P <> nil then begin
    Prev^ := P^.Next;
    Dispose(P);
  end;
end;

end.
