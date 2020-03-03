unit XLSMMU5;

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

uses Classes, SysUtils,
     XLSUtils5;

{$ifdef DELPHI_6}
type NativeInt = integer;
{$endif}

//#! Vector
//#!
//#! Vector header  : word = Number of items in the vector.
//#!                  dword = Bytes to last item of vector (byte 0).
//#!
//#! Items header   : Header relevant to the items. Size defined by ItemsHeaderSize.
//#!
//#! 0 = First byte : Number of equal size items, bits 0-6.
//#!                  If there is more than one equal item, item 2..n have no
//#!                  byte 0 and 1.
//#!                  If bit 7 of the byte is set, this indicates empty items and
//#!                  second byte and data is not present. The remaining bits (0-6)
//#!                  is the number of empty items.
//#!
//#! 1 = Second byte: The size of the item. Coded according to the translation list,
//#!                  L_SizeXLate.
//#! [X = word]     : If bit 7 is set, remaining bits are the size code but the
//#!                  the data is followed by a second extended data with a word
//#!                  prefix size. Extended items has always equal count = 1.
//#!
//#! 2..n           : The item data.                                   + Item 2, empty.
//#!                                                                   |
//#!                                 Item 0          Item 1            |   Item 3 & 4, two equal  Item 5, extended
//#! +---------------+--------------++---+---+------++---+---+------++---++---+---+------+------++---+---+----+-------++----
//#! | Vector header | Items header || 0 | 1 | 2..n || 0 | 1 | 2..n || 0 || 0 | 1 | 2..n | 2..n || 0 | 1 | X  | 2..n  ||
//#! +---------------+--------------++---+---+------++---+---+------++---++---+---+------+------++---+---+----+-------++----

//#! Block
//#!
//#! A block has a list of vector offsets (VectorsOffs) where the offset to
//#! a vectors memory is. Each block can hold XBLOCK_VECTOR_COUNT vectors.
//#! Vectors are assigned memory in the block in the order they are added,
//#! so vectors memory is not sequential.
//#!
//#!  Block memory
//#! +---------+----------------------+-------+ ... +----------+---
//#! | Vect 1  | Vect 0               |Vect 2 |     | Vect n   |
//#! +----^----+-----------^----------+---^---+ ... +----^-----+---
//#!      |                |              |              |
//#!      +---------+      |              |              |
//#!  Vect offs     |      |              |              |
//#! +----------+   |      |              |              |
//#! | Vector 0 |----------+              |              |
//#! +----------+   |                     |              |
//#! | Vector 1 |---+                     |              |
//#! +----------+                         |              |
//#! | Vector 2 |-------------------------+              |
//#! +----------+                                        |
//#! | .......  |                                        |
//#! +----------+                                        |
//#! | Vector n |----------------------------------------+
//#! +----------+
//#!


{.$define MMU_DEBUG}

const XVECT_MAX_EQCOUNT   = $7F;
//#! Flag bit for empty items.
const XVECT_EQCOUNT_EMPTY = $80;
//#! XLate Extended size flag. XLate byte is followed by a word byte with the actually size.
//#! Equal count for these items are always 1. Remaining bits are the XLate code.
const XVECT_XLATE_EX      = $80;
//#! Bits 0-4 available for XLate value
const XVECT_XLATE_MASK    = $1F;

const XVECT_MAX_ITEMS     = 16385;

//#! Size of the equal byte
const SZ_HDR_EQ           = 1;
//#! Size of the XLate byte
const SZ_HDR_XL           = 1;
//#! Size if extended size
const SZ_HDR_EX           = 2;

//#! Number of vectors in a block.
const XBLOCK_VECTOR_COUNT_BITS = 8;
const XBLOCK_VECTOR_MASK       = (1 shl XBLOCK_VECTOR_COUNT_BITS) - 1;
const XBLOCK_VECTOR_COUNT      = 1 shl XBLOCK_VECTOR_COUNT_BITS;

const XBLOCK_VECTOR_UNUSED     = $FFFFFFFF;

const XBLOCK_PREALLOCSIZE = $FFFF;

type TMMUPtr = ^Byte;

type TMMUPtrEvent = function(APtr: TMMUPtr): boolean of object;

type PXVectHeader = ^TXVectHeader;
     TXVectHeader = packed record
     ItemCount: word;
     //#! Offset to the equal byte of the last item. This may not be the actuall
     //#! last item if there is more than one equal item.
     OffsetToLastEq: longword;
     end;

type PXLSMMUBlock = ^TXLSMMUBlock;
     TXLSMMUBlock = record
     MemSize: integer;
     Memory: TMMUPtr;
     //#! Offset to allocation in Memory. Vector memory is not sequential, it's in
     //#! the order the vectors are added.
     //#! Unused vectors are XBLOCK_VECTOR_UNUSED.
     VectorsOffs: array[0..XBLOCK_VECTOR_COUNT - 1] of integer;
     end;

type TXLSMMURebuildVector = class(TObject)
protected
     FMem      : TMMUPtr;
     FMemSize  : integer;
     FAllocSize: integer;

     FWrittenMem : integer;

     FItems    : array[0..XVECT_MAX_ITEMS] of integer;
     FItemCount: integer;

     procedure AddItem(const AOffset: integer); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  AddMem(const ASize: integer): TMMUPtr; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  Add(const APtr: TMMUPtr): TMMUPtr;
     function  Address(const AIndex: integer): TMMUPtr; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
public
     constructor Create(const AVector: TMMUPtr);
     destructor Destroy; override;

     //#! Insert an extended size item.
     procedure InsertExItem(const AItem: integer; const AXLate: byte; const ASize: integer);
     procedure InsertXLateItem(const AItem: integer; const AXLate: byte);
     procedure InsertEmpty(const AItem, ACount: integer);
     function  DeleteItem(const AItem: integer): boolean;
     function  MoveItems(const ASrcItem, ADestItem: integer): integer;
     //#! Returns the address for AItem, if AItem is written, else Nil.
     function  WriteVector(const AVector: TMMUPtr; const  AItem: integer): TMMUPtr;

     function  RequiredMem: integer;

     property ItemCount: integer read FItemCount;

     property WrittenMem: integer read FWrittenMem;
     end;

type TXLSMMUVectorIterator = class(TObject)
{$ifdef DELPHI_2006_OR_LATER}
strict private
{$else}
private
{$endif}
     FVector   : TMMUPtr;

     //#! Number of slots (equal items).
     FSlotCount: integer;
     //#! True if items are empty
     FIsEmpty  : boolean;
     //#! Index to the first item of the current slots, FSlotCount
     FEQIndex  : integer;
     //#! XLate byte of the FEQIndex item.
     FItemXL   : integer;

     FResItem  : TMMUPtr;
     FItem     : TMMUPtr;

     procedure BeginSlot;
     function  HitSlot(const AItem: integer): boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
public
     procedure Clear(const AVector: TMMUPtr);

     function Find(var AItem: integer): boolean;
     function FindNext(var AItem: integer): boolean;

     property ResItem: TMMUPtr read FResItem;
     property ItemXL: integer read FItemXL;
     end;

type TXLSMMUBlockManager = class;

     TXLSMMUVectorManager = class(TObject)
{$ifdef DELPHI_2006_OR_LATER}
strict private
{$else}
private
{$endif}
     FVector      : TMMUPtr;
private
     procedure SetVector(const Value: TMMUPtr);
protected
     FBlockManager: TXLSMMUBlockManager;
     FIterator    : TXLSMMUVectorIterator;

     function  AddEmpty(const APtr: TMMUPtr; const ACount: integer): TMMUPtr;
     //#! Returns number of free'd bytes at APtr. Zero if no bytes where free'e.
     function  DeleteEmpty(const APtr: TMMUPtr): integer;
     function  InitVector(const AItem,AItemSize: integer): TMMUPtr;
     //#! The size of an block of equal items, with equal and size bytes.
     function  GetItemsBlockSize(const APtr: TMMUPtr): integer;
     //#! Returns the end of the vector memory, that is, the insertion point for
     //#! a new item.
     function  GetEndOfMem: TMMUPtr;
     //#! Returns a pointer to the last equal byte.
     function  GetLastEq: TMMUPtr;
     function  SetEqualAndXLate(const APtr: TMMUPtr; const AEqual,AXLate: byte): TMMUPtr;
     function  GetEmptySlotCount(const AEmptyCount: integer): integer; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     //#! AIndexEqual = the item index of the first equal items that AItem is within.
     //#! Only of interrest if there is more than one equal.
     //#! Example: If AItem is 12 and #12 is in a block of equal items from
     //#! #10 to #13, AIndexEqual will be 10
     function  GetItem(const AItem: integer; out APtrEqual: TMMUPtr; out AIndexEqual: integer): TMMUPtr;

     class function  GetXLate(const AIndex: integer): integer; // {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
public
     class procedure SetXLate(const AIndex, AValue: integer);
     class procedure SetItemsHeaderSize(const ASize: integer; const ADefaultData: TMMUPtr; const ADefaultDataSize: integer);

     constructor Create(ABlockManager: TXLSMMUBlockManager);
     destructor Destroy; override;

     function  WalkVector(AList: TStrings): boolean;

     function  ItemCount: integer;

     //#! Memory used by the vector.
     function  MemorySize: integer;

     //#! Creates a new, empty vector (only header and items header).
     function  New: TMMUPtr;
     function  Alloc(const AItem, AXLate: integer): TMMUPtr;
     function  AllocEx(const AItem, AXLate, ASize: integer): TMMUPtr;
     procedure FreeMem(const AItem1,AItem2: integer);
     function  GetMem(const AItem: integer; out AXLate: integer): TMMUPtr;
     //#! Returns a pointer to the first non-empty item, starting at AItem.
     function  GetNextMem(var AItem: integer; out AXLate: integer): TMMUPtr;
     procedure MoveMem(const ASrcItem, ADestItem: integer);
     procedure InsertEmpty(const AItem, ACount: integer);

     function  FirstItem: integer; // {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  LastItem: integer; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     property Vector: TMMUPtr read FVector write SetVector;
     end;

     TXLSMMUBlockManager = class(TObject)
private
     procedure SetBlock(const Value: PXLSMMUBlock); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  GetVector(const AIndex: integer): TMMUPtr;
protected
     FBlock          : PXLSMMUBlock;
     FVectManager    : TXLSMMUVectorManager;
     FCurrVectorIndex: integer;

     //#! Adds memory to the end of the block memory.
     //#! That is, Vector must be Nil as otherwise the vector may be at any
     //#! location in the block.
     //#! Current vector is also updated,
     function  _AddMem(const ASize: integer): TMMUPtr;
     //#! Insert ASize bytes in memory at the location APtr points to.
     function  _InsertMem(const APtr: TMMUPtr; const ASize: integer): TMMUPtr;
     //#! Reallocates the current vector.
     procedure _Realloc(const AOldSize,ANewSize: integer);
     //#! Deletes the current vector.
     procedure _Delete(const ASize: integer);

     procedure SetVector(AVector: integer); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
public
     constructor Create;
     destructor Destroy; override;

     function  WalkVectors(AList: TStrings; ABlockOffs: integer = 0): boolean;

     function  GetMem(const AVector,AItem: integer; out AXLate: integer): TMMUPtr;
     function  AllocMem(const AVector,AItem,AXLate: integer): TMMUPtr;
     function  AllocMemEx(const AVector,AItem,AXLate,ASize: integer): TMMUPtr;
     procedure MoveMem(const AVector,ASrcItem,ADestItem: integer);
     procedure MoveMemAll(const ASrcItem,ADestItem: integer);
     procedure FreeMem(const AVector,AItem: integer);
     procedure FreeMemAll(const AItem1,AItem2: integer);
     procedure InsertEmpty(const AVector,AItem,ACount: integer);
     procedure InsertEmptyAll(const AItem,ACount: integer);

     function  GetItemsHeader(const AVector: integer): TMMUPtr;
     //#! Adds a new vector and returns a pointer to the Items header.
     function  NewVector(const AVector: integer): TMMUPtr;
     procedure FreeVector(const AVector: integer);
     procedure FreeVectors(const AVector1, AVector2: integer);
     procedure MoveVector(const ASrcVector, ADestVector: integer);
     function  CopyVector(const AVector: integer): TMMUPtr;
     procedure InsertVector(const AVector: integer; const APtr: TMMUPtr);
     function  _VectorItemsCount: integer;
     //#! Returns a pointer to the Items header if the vector after AVector (AVector + 1).
     function  GetNextVector(var AVector: integer): TMMUPtr;
     //#! Vector set by GetNextVector.
     function  GetNextItem(var AItem: integer; out AXLate: integer): TMMUPtr; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     procedure CalcDimensions(out AMinVect,AMaxVect,AMinItem,AMaxItem: integer);

     property Block: PXLSMMUBlock read FBlock write SetBlock;
     property Vector[const AIndex: integer]: TMMUPtr read GetVector;
     end;

var  G_XLSCellMMU5_KeepItemsHeaderEvent: TMMUPtrEvent;

implementation

var
     L_SizeXLate: array[0..$FF] of integer;
     L_ItemsHeaderSize: integer;
     L_ItemsHeaderDefaultData: TMMUPtr;

{ TXLSMMUBlockManager }

function TXLSMMUBlockManager.AllocMemEx(const AVector, AItem, AXLate, ASize: integer): TMMUPtr;
begin
  SetVector(AVector);
  Result := FVectManager.AllocEx(AItem,AXLate,ASize);
{$ifdef MMU_DEBUG}
  WalkVectors(Nil);
{$endif}
end;

function TXLSMMUBlockManager.AllocMem(const AVector, AItem, AXLate: integer): TMMUPtr;
begin
  SetVector(AVector);
  Result := FVectManager.Alloc(AItem,AXLate);
{$ifdef MMU_DEBUG}
  WalkVectors(Nil);
{$endif}
end;

procedure TXLSMMUBlockManager.CalcDimensions(out AMinVect, AMaxVect, AMinItem, AMaxItem: integer);
var
  i: integer;
  I1,I2: integer;
begin
  AMinVect := MAXINT;
  AMaxVect := -(MAXINT - 1);
  AMinItem := MAXINT;
  AMaxItem := -(MAXINT - 1);
  for i := 0 to XBLOCK_VECTOR_COUNT - 1 do begin
    if FBlock.VectorsOffs[i] <> Integer(XBLOCK_VECTOR_UNUSED) then begin
      if i < AMinVect then
        AMinVect := i;
      if i > AMaxVect then
        AMaxVect := i;

      SetVector(i);
      I1 := FVectManager.FirstItem;
      I2 := FVectManager.LastItem;
      if (I1 >= 0) and (I1 < AMinItem) then
        AMinItem := I1;
      if (I2 >= 0) and (I2 > AMaxItem) then
        AMaxItem := I2;
    end;
  end;
end;

function TXLSMMUBlockManager.CopyVector(const AVector: integer): TMMUPtr;
var
  P: TMMUPtr;
  Sz: integer;
begin
  if FBlock.VectorsOffs[AVector] = Integer(XBLOCK_VECTOR_UNUSED) then
    Result := Nil
  else begin
    SetVector(AVector);
    P := FVectManager.Vector;
    Sz := FVectManager.MemorySize;
    System.GetMem(Result,Sz);
    System.Move(P^,Result^,Sz);
  end;
end;

constructor TXLSMMUBlockManager.Create;
begin
  FCurrVectorIndex := -1;
  FVectManager := TXLSMMUVectorManager.Create(Self);
end;

destructor TXLSMMUBlockManager.Destroy;
begin
  FVectManager.Free;
  // Memory is free'd by CellMMU
//  System.FreeMem(FBlock.Memory);
  inherited;
end;

procedure TXLSMMUBlockManager.FreeMem(const AVector, AItem: integer);
begin
  if FBlock.VectorsOffs[AVector] <> Integer(XBLOCK_VECTOR_UNUSED) then begin
    FCurrVectorIndex := AVector;
    FVectManager.Vector := Vector[FCurrVectorIndex];
    if FVectManager.Vector <> Nil then
      FVectManager.FreeMem(AItem,AItem);
  end;
end;

procedure TXLSMMUBlockManager.FreeMemAll(const AItem1, AItem2: integer);
var
  i: integer;
begin
  for i := 0 to XBLOCK_VECTOR_COUNT - 1 do begin
    FCurrVectorIndex := i;
    if FBlock.VectorsOffs[i] <> Integer(XBLOCK_VECTOR_UNUSED) then begin
      FVectManager.Vector := Vector[i];
      FVectManager.FreeMem(AItem1,AItem2);
    end;
  end;
end;

procedure TXLSMMUBlockManager.FreeVector(const AVector: integer);
begin
  FCurrVectorIndex := AVector;
  if FBlock.VectorsOffs[AVector] <> Integer(XBLOCK_VECTOR_UNUSED) then begin
    FVectManager.Vector := Vector[FCurrVectorIndex];
    _Delete(FVectManager.MemorySize);
  end;
end;

procedure TXLSMMUBlockManager.FreeVectors(const AVector1, AVector2: integer);
var
  i: integer;
begin
  for i := AVector1 to AVector2 do
    FreeVector(i);
end;

function TXLSMMUBlockManager.GetItemsHeader(const AVector: integer): TMMUPtr;
begin
  if FBlock.VectorsOffs[AVector] <> Integer(XBLOCK_VECTOR_UNUSED) then begin
    Result := GetVector(AVector);
    if Result <> Nil then
      Inc(Result,SizeOf(TXVectHeader));
  end
  else
    Result := Nil;
end;

function TXLSMMUBlockManager.GetMem(const AVector, AItem: integer; out AXLate: integer): TMMUPtr;
begin
  FCurrVectorIndex := AVector;
  if FBlock.VectorsOffs[FCurrVectorIndex] <> Integer(XBLOCK_VECTOR_UNUSED) then begin
    FVectManager.Vector := Vector[FCurrVectorIndex];
    Result := FVectManager.GetMem(AItem,AXLate);
  end
  else
    Result := Nil;
end;

function TXLSMMUBlockManager.GetNextItem(var AItem: integer; out AXLate: integer): TMMUPtr;
begin
  if FCurrVectorIndex >= 0 then
    Result := FVectManager.GetNextMem(AItem,AXLate)
  else
    Result := Nil;
end;

function TXLSMMUBlockManager.GetNextVector(var AVector: integer): TMMUPtr;
begin
  Result := Nil;
  FCurrVectorIndex := AVector;
  while FCurrVectorIndex < XBLOCK_VECTOR_COUNT do begin
    if FBlock.VectorsOffs[FCurrVectorIndex] <> Integer(XBLOCK_VECTOR_UNUSED) then begin
      Result := FBlock.Memory;
      Inc(Result,FBlock.VectorsOffs[FCurrVectorIndex]);
      Break;
    end;
    Inc(FCurrVectorIndex);
  end;
  if Result <> Nil then begin
    FVectManager.Vector := Vector[FCurrVectorIndex];
    AVector := FCurrVectorIndex;
  end
  else begin
    FCurrVectorIndex := -1;
    AVector := -1;
  end;
end;

function TXLSMMUBlockManager.GetVector(const AIndex: integer): TMMUPtr;
begin
{$ifdef MMU_DEBUG}
  if (AIndex < 0) or (AIndex > High(FBlock.VectorsOffs)) then
    raise XLSRWException.Create('Index out of range');
{$endif}
  SetVector(AIndex);
  Result := FBlock.Memory;
  if Result <> Nil then
    Inc(Result,FBlock.VectorsOffs[AIndex]);
end;

procedure TXLSMMUBlockManager.InsertEmpty(const AVector, AItem, ACount: integer);
begin
  SetVector(AVector);
  FVectManager.InsertEmpty(AItem,ACount);
end;

procedure TXLSMMUBlockManager.InsertEmptyAll(const AItem, ACount: integer);
var
  i: integer;
begin
  for i := 0 to XBLOCK_VECTOR_COUNT - 1 do begin
    if FBlock.VectorsOffs[i] <> Integer(XBLOCK_VECTOR_UNUSED) then
      InsertEmpty(i,AItem,ACount);
  end;
end;

procedure TXLSMMUBlockManager.InsertVector(const AVector: integer; const APtr: TMMUPtr);
var
  P: TMMUPtr;
  Sz: integer;
begin
  if FBlock.VectorsOffs[AVector] <> Integer(XBLOCK_VECTOR_UNUSED) then
    FreeVector(AVector);

  if APtr = Nil then
    Exit;

  FBlock.VectorsOffs[AVector] := FBlock.MemSize;

  FCurrVectorIndex := AVector;

  FVectManager.Vector := APtr;
  Sz := FVectManager.MemorySize;
  P := _AddMem(Sz);
  System.Move(Aptr^,P^,Sz);
end;

procedure TXLSMMUBlockManager.MoveMemAll(const ASrcItem, ADestItem: integer);
var
  i: integer;
begin
  for i := 0 to XBLOCK_VECTOR_COUNT - 1 do begin
    FCurrVectorIndex := i;
    if FBlock.VectorsOffs[i] <> Integer(XBLOCK_VECTOR_UNUSED) then begin
      FVectManager.Vector := Vector[i];
      FVectManager.MoveMem(ASrcItem,ADestItem);
    end;
  end;
end;

procedure TXLSMMUBlockManager.MoveMem(const AVector, ASrcItem, ADestItem: integer);
begin
  SetVector(AVector);
  FVectManager.MoveMem(ASrcItem,ADestItem);
{$ifdef MMU_DEBUG}
  WalkVectors(Nil);
{$endif}
end;

procedure TXLSMMUBlockManager.MoveVector(const ASrcVector, ADestVector: integer);
begin
  // Free destination vector if it is allocated.
  if FBlock.VectorsOffs[ADestVector] <> Integer(XBLOCK_VECTOR_UNUSED) then
    FreeVector(ADestVector);
  // Assign if source is allocated.
  if FBlock.VectorsOffs[ASrcVector] <> Integer(XBLOCK_VECTOR_UNUSED) then begin
    FBlock.VectorsOffs[ADestVector] := FBlock.VectorsOffs[ASrcVector];
    FBlock.VectorsOffs[ASrcVector] := Integer(XBLOCK_VECTOR_UNUSED);
  end;
end;

function TXLSMMUBlockManager.NewVector(const AVector: integer): TMMUPtr;
begin
  FCurrVectorIndex := AVector;
  if FBlock.VectorsOffs[AVector] = Integer(XBLOCK_VECTOR_UNUSED) then begin
    FBlock.VectorsOffs[AVector] := FBlock.MemSize;
    FVectManager.Vector := Nil;
    Result := FVectManager.New;
  end
  else begin
    Result := FBlock.Memory;
    Inc(Result,FBlock.VectorsOffs[AVector]);
  end;
  Inc(Result,SizeOf(TXVectHeader));
end;

procedure TXLSMMUBlockManager.SetVector(AVector: integer);
begin
  if AVector <> FCurrVectorIndex then begin
    FCurrVectorIndex := AVector;
    if FBlock.VectorsOffs[AVector] = Integer(XBLOCK_VECTOR_UNUSED) then begin
      FBlock.VectorsOffs[AVector] := FBlock.MemSize;
      FVectManager.Vector := Nil;
    end
    else
      FVectManager.Vector := Vector[FCurrVectorIndex];
  end;
end;

function TXLSMMUBlockManager._VectorItemsCount: integer;
begin
  if FVectManager.Vector = Nil then
    Result := 0
  else
    Result := FVectManager.ItemCount;
end;

procedure TXLSMMUBlockManager.SetBlock(const Value: PXLSMMUBlock);
begin
//  if FIsUpdating then
//    raise XLSRWException.Create('call to SetBlock while IsUpdating');
  FBlock := Value;
  FCurrVectorIndex := -1;
  FVectManager.Vector := Nil;
end;

function TXLSMMUBlockManager.WalkVectors(AList: TStrings; ABlockOffs: integer): boolean;
var
  i: integer;
begin
  Result := (FBlock <> Nil);
  if Result then begin
    for i := 0 to High(FBlock.VectorsOffs) do begin
      if FBlock.VectorsOffs[i] <> Integer(XBLOCK_VECTOR_UNUSED) then begin
        FVectManager.Vector := Vector[i];
        if AList <> Nil then
          AList.Add(Format('Vector #%d =========================',[i + ABlockOffs]));
        Result := FVectManager.WalkVector(AList);
        if not Result then
          Exit;
      end;
    end;
  end;
end;

function TXLSMMUBlockManager._AddMem(const ASize: integer): TMMUPtr;
begin
  System.ReAllocMem(FBlock.Memory,FBlock.MemSize + ASize);

  Result := FBlock.Memory;
  Inc(Result,FBlock.MemSize);
  Inc(FBlock.MemSize,ASize);

  FVectManager.Vector := Vector[FCurrVectorIndex];
end;

procedure TXLSMMUBlockManager._Delete(const ASize: integer);
var
  i: integer;
  Sz: integer;
  pSrc,pDest: TMMUPtr;
  Offs: integer;
begin
  Offs := FBlock.VectorsOffs[FCurrVectorIndex];
  FBlock.VectorsOffs[FCurrVectorIndex] := Integer(XBLOCK_VECTOR_UNUSED);
  for i := 0 to XBLOCK_VECTOR_COUNT - 1 do begin
    if (FBlock.VectorsOffs[i] <> Integer(XBLOCK_VECTOR_UNUSED)) and (FBlock.VectorsOffs[i] > Offs) then
      Dec(FBlock.VectorsOffs[i],ASize);
  end;

  Sz := FBlock.MemSize - Offs - ASize;

  pSrc := FBlock.Memory;
  Inc(pSrc,Offs);
  pDest := pSrc;
  Inc(pSrc,ASize);
  System.Move(pSrc^,pDest^,Sz);

  Dec(FBlock.MemSize,ASize);

{$ifdef MMU_DEBUG}
  if FBlock.MemSize < 0 then
    raise XLSRWException.Create('Free vector: Memory < 0');
{$endif}
  System.ReallocMem(FBlock.Memory,FBlock.MemSize);

  FCurrVectorIndex := -1;
  FVectManager.Vector := Nil;
end;

function TXLSMMUBlockManager._InsertMem(const APtr: TMMUPtr; const ASize: integer): TMMUPtr;
var
  i: integer;
  P: TMMUPtr;
  Offs: integer;
  Sz: integer;
begin
  Offs := Longword(NativeInt(APtr) - NativeInt(FBlock.Memory));
{$ifdef MMU_DEBUG}
  if Offs <> FBlock.MemSize then
    raise XLSRWException.Create('_InsertMem: Offs <> MemSize');
{$endif}

  System.ReAllocMem(FBlock.Memory,FBlock.MemSize + ASize);

  Inc(FBlock.MemSize,ASize);
  Result := FBlock.Memory;
  Inc(Result,Offs);

  Sz := FBlock.MemSize - Offs - ASize;
  if Sz > 0 then begin
    P := Result;
    Inc(P,ASize);
    System.Move(Result^,P^,Sz);

    Offs := FBlock.VectorsOffs[FCurrVectorIndex];
    for i := 0 to XBLOCK_VECTOR_COUNT - 1 do begin
      if (FBlock.VectorsOffs[i] <> Integer(XBLOCK_VECTOR_UNUSED)) and (FBlock.VectorsOffs[i] > Offs) then
        Inc(FBlock.VectorsOffs[i],ASize);
    end;
  end;

{$ifdef MMU_DEBUG}
  if Sz < 0 then
    raise XLSRWException.Create('_InsertMem: Sz < 0');

  if FBlock.MemSize < 0 then
    raise XLSRWException.Create('_InsertMem: Memory < 0');
{$endif}
  FVectManager.Vector := Vector[FCurrVectorIndex];
end;

procedure TXLSMMUBlockManager._Realloc(const AOldSize,ANewSize: integer);
var
  i: integer;
  Sz: integer;
  Offs: integer;
  Delta: integer;
  pSrc,pDest: TMMUPtr;
begin
  Delta := ANewSize - AOldSize;
  if Delta = 0 then
    Exit;

  Offs := FBlock.VectorsOffs[FCurrVectorIndex];

  Sz := FBlock.MemSize - Offs - AOldSize;

  if Delta < 0 then begin
    pSrc := Vector[FCurrVectorIndex];
    pDest := pSrc;
    Inc(pSrc,AOldSize);
    Inc(pDest,ANewSize);
    System.Move(pSrc^,pDest^,Sz);
  end;

  System.ReAllocMem(FBlock.Memory,FBlock.MemSize + Delta);

  for i := 0 to XBLOCK_VECTOR_COUNT - 1 do begin
    if (FBlock.VectorsOffs[i] <> Integer(XBLOCK_VECTOR_UNUSED)) and (FBlock.VectorsOffs[i] > Offs) then
      Inc(FBlock.VectorsOffs[i],Delta);
  end;

  if Delta > 0 then begin
    pSrc := Vector[FCurrVectorIndex];
    pDest := pSrc;
    Inc(pSrc,AOldSize);
    Inc(pDest,ANewSize);
    System.Move(pSrc^,pDest^,Sz);
  end;

  Inc(FBlock.MemSize,Delta);

  FVectManager.Vector := Vector[FCurrVectorIndex];
end;

{ TXLSMMUVectorManager }

function TXLSMMUVectorManager.AllocEx(const AItem, AXLate, ASize: integer): TMMUPtr;
var
  n: integer;
  AllocSz: integer;
  pItem: TMMUPtr;
  pItemEq: TMMUPtr;
  Rebuild: TXLSMMURebuildVector;
begin
  AllocSz := SZ_HDR_EQ + SZ_HDR_XL + SZ_HDR_EX + ASize;
  if (FVector = Nil) or (PXVectHeader(FVector).ItemCount <= 0) then begin
    Result := InitVector(AItem,AllocSz);
    Result := SetEqualAndXLate(Result,1,XVECT_XLATE_EX or Byte(AXLate));
    PWord(Result)^ := ASize;
  end
  //#! Directly after the last item.
  else if AItem = PXVectHeader(FVector).ItemCount then begin
    pItem := GetEndOfMem;
    Result := FBlockManager._InsertMem(pItem,AllocSz);
    pItemEq := GetLastEq;
    Result := SetEqualAndXLate(Result,1,XVECT_XLATE_EX or Byte(AXLate));
    PWord(Result)^ := ASize;
    PXVectHeader(FVector).ItemCount := AItem + 1;
{$ifdef MMU_DEBUG}
    if PXVectHeader(FVector).ItemCount > XVECT_MAX_ITEMS then
      raise XLSRWException.Create('Vector overlow');
{$endif}
    Inc(PXVectHeader(FVector).OffsetToLastEq,GetItemsBlockSize(pItemEq));
  end
  //#! One or more empty between the item and the last item. Insert empty.
  else if AItem > PXVectHeader(FVector).ItemCount then begin
    pItem := GetEndOfMem;

    n := GetEmptySlotCount(AItem - PXVectHeader(FVector).ItemCount);
    Result := FBlockManager._InsertMem(pItem,n + AllocSz);

    pItemEq := GetLastEq;
    Result := AddEmpty(Result,AItem - PXVectHeader(FVector).ItemCount);
    Result := SetEqualAndXLate(Result,1,XVECT_XLATE_EX or Byte(AXLate));
    PWord(Result)^ := ASize;
    PXVectHeader(FVector).ItemCount := AItem + 1;
{$ifdef MMU_DEBUG}
    if PXVectHeader(FVector).ItemCount > XVECT_MAX_ITEMS then
      raise XLSRWException.Create('Vector overlow');
{$endif}
    Inc(PXVectHeader(FVector).OffsetToLastEq,GetItemsBlockSize(pItemEq) + n);
  end
  //#! Somewhere between the first and the last.
  //#! Rebuild the vector. To compilicated right now to sort out all combinations
  //#! when inserting new items. Hughe time savings can be done here...
  else begin
    Rebuild := TXLSMMURebuildVector.Create(FVector);
    try
      Rebuild.InsertExItem(AItem,AXLate,ASize);
      AllocSz := Rebuild.RequiredMem;

      FBlockManager._Realloc(MemorySize,AllocSz);
      Result := Rebuild.WriteVector(FVector,AItem);

      if Rebuild.WrittenMem <> AllocSz then
        raise XLSRWException.Create('Rebuild ex mem error');
    finally
      Rebuild.Free;
    end;
  end;
{$ifdef MMU_DEBUG}
  MemorySize;
  if Result = Nil then
    raise XLSRWException.Create('Result is Nil');
  WalkVector(Nil);
{$endif}
end;

function TXLSMMUVectorManager.Alloc(const AItem, AXLate: integer): TMMUPtr;
var
  n: integer;
  ItemSz: integer;
  AllocSz: integer;
  pItem: TMMUPtr;
  pItemEq: TMMUPtr;
  Rebuild: TXLSMMURebuildVector;
begin
  ItemSz := L_SizeXLate[AXLate];
  AllocSz := SZ_HDR_EQ + SZ_HDR_XL + ItemSz;
  if (FVector = Nil) or (PXVectHeader(FVector).ItemCount <= 0) then begin
    Result := InitVector(AItem,AllocSz);
    Result := SetEqualAndXLate(Result,1,AXLate);
  end
  //#! Directly after the last item.
  else if AItem = PXVectHeader(FVector).ItemCount then begin
    pItemEq := GetLastEq;
    pItem := GetEndOfMem;
    //#! Equal count full, or not the same item.
    if (PByteArray(pItemEq)[0] = XVECT_MAX_EQCOUNT) or (PByteArray(pItemEq)[1] <> AXLate) then begin
      Result := FBlockManager._InsertMem(pItem,AllocSz);
      //#! pItemEq is not valid after mem reallocation.
      pItemEq := GetLastEq;
      Result := SetEqualAndXLate(Result,1,AXLate);
      PXVectHeader(FVector).ItemCount := AItem + 1;
{$ifdef MMU_DEBUG}
      if PXVectHeader(FVector).ItemCount > XVECT_MAX_ITEMS then
        raise XLSRWException.Create('Vector overlow');
{$endif}
      Inc(PXVectHeader(FVector).OffsetToLastEq,GetItemsBlockSize(pItemEq));
    end
    else begin
      Result := FBlockManager._InsertMem(pItem,ItemSz);
      pItemEq := GetLastEq;
      pItemEq^ := pItemEq^ + 1;
      PXVectHeader(FVector).ItemCount := AItem + 1;
    end;
  end

  //#! One or more empty between the item and the last item. Insert empty.
  else if AItem > PXVectHeader(FVector).ItemCount then begin
    pItem := GetEndOfMem;

    n := GetEmptySlotCount(AItem - PXVectHeader(FVector).ItemCount);
    Result := FBlockManager._InsertMem(pItem,n + AllocSz);

    pItemEq := GetLastEq;
    Result := AddEmpty(Result,AItem - PXVectHeader(FVector).ItemCount);
    Result := SetEqualAndXLate(Result,1,AXLate);
    PXVectHeader(FVector).ItemCount := AItem + 1;
{$ifdef MMU_DEBUG}
    if PXVectHeader(FVector).ItemCount > XVECT_MAX_ITEMS then
      raise XLSRWException.Create('Vector overlow');
{$endif}
    Inc(PXVectHeader(FVector).OffsetToLastEq,GetItemsBlockSize(pItemEq) + n);
  end
  //#! Somewhere between the first and the last.
  //#! Rebuild the vector. To compilicated right now to sort out all combinations
  //#! when inserting new items. Hughe time savings can be done here...
  else begin
    Rebuild := TXLSMMURebuildVector.Create(FVector);
    try
      Rebuild.InsertXLateItem(AItem,AXLate);
      AllocSz := Rebuild.RequiredMem;
      FBlockManager._Realloc(MemorySize,AllocSz);
      Result := Rebuild.WriteVector(FVector,AItem);
    finally
      Rebuild.Free;
    end;
  end;
{$ifdef MMU_DEBUG}
  MemorySize;
  if Result = Nil then
    raise XLSRWException.Create('Result is Nil');
{$endif}
end;

constructor TXLSMMUVectorManager.Create(ABlockManager: TXLSMMUBlockManager);
begin
  FBlockManager := ABlockManager;
  FIterator := TXLSMMUVectorIterator.Create;
end;

function TXLSMMUVectorManager.DeleteEmpty(const APtr: TMMUPtr): integer;
var
  n: integer;
begin
  Result := 0;
{$ifdef MMU_DEBUG}
  if (APtr^ and XVECT_EQCOUNT_EMPTY) <> 0 then
    raise XLSRWException.Create('Empty item excpectd');
{$endif}
  n := APtr^ and XVECT_MAX_EQCOUNT;
  if n = 1 then
    Result := 1
  else if n = XVECT_MAX_EQCOUNT then begin
    APtr^ := (n - 1) or XVECT_EQCOUNT_EMPTY;
    Result := 0;
  end;
end;

destructor TXLSMMUVectorManager.Destroy;
begin
  FIterator.Free;

  inherited;
end;

function TXLSMMUVectorManager.FirstItem: integer;
var
  n: integer;
  Cnt: integer;
  P: TMMUPtr;
begin
  Result := 1;
  Cnt := PXVectHeader(FVector).ItemCount;
  P := FVector;
  Inc(P,SizeOf(TXVectHeader) + L_ItemsHeaderSize);
  while Cnt > 0 do begin
    if (P^ and XVECT_EQCOUNT_EMPTY) <> 0 then begin
      n := P^ and XVECT_MAX_EQCOUNT;
      Inc(Result,n);
      Dec(Cnt,n);
      Inc(P);
    end
    else begin
      Dec(Result);
      Exit;
    end;
  end;
  Result := -1;
end;

procedure TXLSMMUVectorManager.FreeMem(const AItem1, AItem2: integer);
var
  i: integer;
  n: integer;
  AllocSz: integer;
  Rebuild: TXLSMMURebuildVector;
begin
  if (AItem1 >= PXVectHeader(FVector).ItemCount) and (AItem2 >= PXVectHeader(FVector).ItemCount) then
    Exit;

  Rebuild := TXLSMMURebuildVector.Create(FVector);
  try
    n := 0;
    for i := AItem1 to AItem2 do begin
      if Rebuild.DeleteItem(i) then
        Inc(n);
    end;

    if n > 0 then begin
      if Rebuild.ItemCount <= 0 then begin
        if G_XLSCellMMU5_KeepItemsHeaderEvent(FVector) then begin
          PXVectHeader(FVector).ItemCount := 0;
          PXVectHeader(FVector).OffsetToLastEq := SizeOf(TXVectHeader) + L_ItemsHeaderSize;
        end
        else
          FBlockManager._Delete(MemorySize);
      end
      else begin
        AllocSz := Rebuild.RequiredMem;
        FBlockManager._Realloc(MemorySize,AllocSz);
        Rebuild.WriteVector(FVector,MAXINT);
{$ifdef MMU_DEBUG}
        WalkVector(Nil);
{$endif}
      end;
    end;
  finally
    Rebuild.Free;
  end;
end;

function TXLSMMUVectorManager.GetEmptySlotCount(const AEmptyCount: integer): integer;
begin
  //#! Number of slots of empty items.
  Result := AEmptyCount div XVECT_MAX_EQCOUNT;
  //#! Will there be an extra slot?
  if (AEmptyCount mod XVECT_MAX_EQCOUNT) <> 0 then
    Inc(Result);
end;

function TXLSMMUVectorManager.GetEndOfMem: TMMUPtr;
var
  EQ,XL: integer;
begin
  Result := FVector;
  Inc(Result,PXVectHeader(FVector).OffsetToLastEq);
  EQ := Result^;
  Inc(Result);

{$ifdef MMU_DEBUG}
  if (EQ = 0) or (EQ > XVECT_MAX_EQCOUNT) then
    raise XLSRWException.Create('Invalid equal byte');
{$endif}
  XL := Result^;
  Inc(Result);
  if (XL and XVECT_XLATE_EX) <> 0 then begin
    Inc(Result,PWord(Result)^);
    Inc(Result,SZ_HDR_EX);
  end
  else begin
{$ifdef MMU_DEBUG}
    if (XL = 0) or (L_SizeXLate[XL] = 0) then
      raise XLSRWException.Create('Invalid XLate byte');
{$endif}
    Inc(Result,EQ * L_SizeXLate[XL]);
  end;
end;

function TXLSMMUVectorManager.GetItem(const AItem: integer; out APtrEqual: TMMUPtr; out AIndexEqual: integer): TMMUPtr;
var
  Cnt: integer;
  IsEmpty: boolean;
  EqCnt: integer;
begin
  Result := FVector;
  Cnt := PXVectHeader(Result).ItemCount;
  Inc(Result,SizeOf(TXVectHeader) + L_ItemsHeaderSize);
  AIndexEqual := 0;
  while AIndexEqual <= Cnt do begin
    IsEmpty := (Result^ and XVECT_EQCOUNT_EMPTY) <> 0;
    EqCnt := Result^ and XVECT_MAX_EQCOUNT;
    if (AItem >= AIndexEqual) and (AItem < (AIndexEqual + EqCnt)) then begin
      APtrEqual := Result;
      if IsEmpty then
        Result := Nil
      else
        Inc(Result,SizeOf(byte) + (AItem - 1) * L_SizeXLate[Result^]);
      Exit;
    end;
    Inc(Result);
    if not IsEmpty then
      Inc(Result,SizeOf(byte) + EqCnt * L_SizeXLate[Result^]);
    Inc(AIndexEqual,EqCnt);
  end;
end;

function TXLSMMUVectorManager.GetItemsBlockSize(const APtr: TMMUPtr): integer;
var
  n: integer;
  P: TMMUPtr;
begin
  P := APtr;
  n := P^;
{$ifdef MMU_DEBUG}
  if (n and XVECT_EQCOUNT_EMPTY) <> 0 then
    raise XLSRWException.Create('Last item is empty item');
  if (n = 0) or (n > XVECT_MAX_EQCOUNT) then
    raise XLSRWException.Create('Invalid equal byte');
{$endif}
  Inc(P);
  if (P^ and XVECT_XLATE_EX) <> 0 then begin
    Inc(P);
    Result := SZ_HDR_EQ + SZ_HDR_XL + SZ_HDR_EX + PWord(P)^;
  end
  else begin
{$ifdef MMU_DEBUG}
    if (P^ = 0) or (L_SizeXLate[P^] = 0) then
      raise XLSRWException.Create('Invalid XLate byte');
{$endif}
    Result := SZ_HDR_EQ + SZ_HDR_XL + L_SizeXLate[P^] * n;
  end;
end;

function TXLSMMUVectorManager.GetLastEq: TMMUPtr;
begin
  Result := FVector;
  Inc(Result,PXVectHeader(FVector).OffsetToLastEq);
end;

function TXLSMMUVectorManager.GetMem(const AItem: integer; out AXLate: integer): TMMUPtr;
var
  I: integer;
begin
  I := AItem;
  if FIterator.Find(I) then begin
    AXLate := FIterator.ItemXL;
    Result := FIterator.ResItem;
  end
  else
    Result := Nil;
end;

function TXLSMMUVectorManager.GetNextMem(var AItem: integer; out AXLate: integer): TMMUPtr;
begin
  if FIterator.FindNext(AItem) then begin
    AXLate := FIterator.ItemXL;
    Result := FIterator.ResItem;
  end
  else
    Result := Nil;
end;

class function TXLSMMUVectorManager.GetXLate(const AIndex: integer): integer;
begin
  Result := L_SizeXLate[AIndex];
end;

function TXLSMMUVectorManager.InitVector(const AItem, AItemSize: integer): TMMUPtr;
var
  Sz: integer;
  P: TMMUPtr;
  HasItemsHeader: boolean;
begin
  Sz := SizeOf(TXVectHeader) + L_ItemsHeaderSize + AItemSize;

  //#! Empty items before the item?
  if AItem > 0 then
    Inc(Sz,GetEmptySlotCount(AItem));

  //#! FVector <> Nil when all items are deleted but the items header contains
  //#! interresting data.
  HasItemsHeader := FVector <> Nil;
  if HasItemsHeader then
    FBlockManager._Realloc(PXVectHeader(FVector).OffsetToLastEq,Sz)
  else
    FVector := FBlockManager._AddMem(Sz);

  Result := FVector;
  Inc(Result,SizeOf(TXVectHeader) + L_ItemsHeaderSize);
  if AItem > 0 then
    Result := AddEmpty(Result,AItem);

  PXVectHeader(FVector).ItemCount := AItem + 1;
  PXVectHeader(FVector).OffsetToLastEq := Integer(NativeInt(Result) - NativeInt(FVector));

  if not HasItemsHeader then begin
    P := FVector;
    Inc(P,SizeOf(TXVectHeader));
    Move(L_ItemsHeaderDefaultData^,P^,L_ItemsHeaderSize);
  end;

  FIterator.Clear(FVector);
end;

procedure TXLSMMUVectorManager.InsertEmpty(const AItem, ACount: integer);
var
  AllocSz: integer;
  Rebuild: TXLSMMURebuildVector;
begin
  if (AItem >= PXVectHeader(FVector).ItemCount) then
    Exit;

  Rebuild := TXLSMMURebuildVector.Create(FVector);
  try
    Rebuild.InsertEmpty(AItem,ACount);
    if Rebuild.ItemCount <= 0 then
      FBlockManager._Delete(MemorySize)
    else begin
      AllocSz := Rebuild.RequiredMem;
      FBlockManager._Realloc(MemorySize,AllocSz);
      Rebuild.WriteVector(FVector,MAXINT);
    end;
  finally
    Rebuild.Free;
  end;
end;

function TXLSMMUVectorManager.ItemCount: integer;
begin
  Result := PXVectHeader(FVector).ItemCount;
end;

function TXLSMMUVectorManager.LastItem: integer;
begin
  Result := PXVectHeader(FVector).ItemCount - 1;
end;

function TXLSMMUVectorManager.MemorySize: integer;
begin
  if PXVectHeader(FVector).ItemCount <= 0 then
    Result := SizeOf(TXVectHeader) + L_ItemsHeaderSize
  else
    Result := Integer(PXVectHeader(FVector).OffsetToLastEq) + GetItemsBlockSize(TMMUPtr(Int64(FVector) + PXVectHeader(FVector).OffsetToLastEq));
end;

procedure TXLSMMUVectorManager.MoveMem(const ASrcItem, ADestItem: integer);
var
  n      : integer;
  AllocSz: integer;
  Rebuild: TXLSMMURebuildVector;
begin
  if (ASrcItem >= PXVectHeader(FVector).ItemCount) and (ADestItem >= PXVectHeader(FVector).ItemCount) then
    Exit;

  Rebuild := TXLSMMURebuildVector.Create(FVector);
  try
    n := Rebuild.MoveItems(ASrcItem,ADestItem);
    if Rebuild.ItemCount <= 0 then
      FBlockManager._Delete(MemorySize)
    else if n > 0 then begin
      AllocSz := Rebuild.RequiredMem;
      FBlockManager._Realloc(MemorySize,AllocSz);
      Rebuild.WriteVector(FVector,MAXINT);
    end
    else
      FreeMem(ADestItem,LastItem);
  finally
    Rebuild.Free;
  end;
end;

function TXLSMMUVectorManager.New: TMMUPtr;
var
  Sz: integer;
begin
  Sz := SizeOf(TXVectHeader) + L_ItemsHeaderSize;
  FVector := FBlockManager._AddMem(Sz);

  Result := FVector;

  PXVectHeader(FVector).ItemCount := 0;
  PXVectHeader(FVector).OffsetToLastEq := Sz;
end;

function TXLSMMUVectorManager.SetEqualAndXLate(const APtr: TMMUPtr; const AEqual, AXLate: byte): TMMUPtr;
begin
  Result := APtr;
  Result^ := AEqual;
  Inc(Result);
  Result^ := AXLate;
  Inc(Result);
end;

class procedure TXLSMMUVectorManager.SetItemsHeaderSize(const ASize: integer; const ADefaultData: TMMUPtr; const ADefaultDataSize: integer);
begin
  L_ItemsHeaderSize := ASize;

  if L_ItemsHeaderDefaultData = Nil then
    System.GetMem(L_ItemsHeaderDefaultData,ADefaultDataSize);

  System.Move(ADefaultData^,L_ItemsHeaderDefaultData^,ADefaultDataSize);
//  L_ItemsHeaderDefaultData := ADefaultData;
end;

procedure TXLSMMUVectorManager.SetVector(const Value: TMMUPtr);
begin
  if (Value <> FVector) or (Value = Nil) then begin
    FVector := Value;
    FIterator.Clear(FVector);
  end;
end;

class procedure TXLSMMUVectorManager.SetXLate(const AIndex, AValue: integer);
begin
  if AIndex > High(L_SizeXLate) then
    raise XLSRWException.Create('Index out of range');
  L_SizeXLate[AIndex] := AValue;
end;

function TXLSMMUVectorManager.WalkVector(AList: TStrings): boolean;
var
  i: integer;
  S: string;
  P: TMMUPtr;
  Offs: integer;
  Cnt: integer;
  MemSz: integer;
  ItemSz: integer;
  EqItemCnt: integer;
  XLate: integer;

function DumpBytes(PBytes: PByteArray; Size: integer): string;
var
  i: integer;
begin
  Result := '';
  if AList <> Nil then begin
    for i := 0 to Size - 1 do
      Result := Result + Format('%.2X',[PBytes[i]]);
  end;
end;

procedure VectError(const AText: string);
begin
  if AList <> Nil then
    AList.Add(AText)
  else
    raise XLSRWException.Create(AText);
end;

begin
  Result := True;

  P := FVector;

  if AList <> Nil then
    AList.Add(Format('Vect hdr: Cont=%d, EqOffs=%.4X',[PXVectHeader(P).ItemCount,PXVectHeader(P).OffsetToLastEq]));

  Cnt := PXVectHeader(P).ItemCount;
//  if Cnt <= 0 then
//    raise XLSRWException.Create('Vector is empty');

  Offs := SizeOf(TXVectHeader) + L_ItemsHeaderSize;

  if Cnt > XVECT_MAX_ITEMS then begin
    VectError('Invalid item count in vector');
    Exit;
  end;

  MemSz := MemorySize;

  Inc(P,SizeOf(TXVectHeader));

  Inc(P,L_ItemsHeaderSize);

  Dec(MemSz,SizeOf(TXVectHeader) + L_ItemsHeaderSize);

  i := 0;
  while Cnt > 0 do begin
    EqItemCnt := PByte(P)^;
    if (EqItemCnt and XVECT_EQCOUNT_EMPTY) <> 0 then begin
      EqItemCnt := EqItemCnt and not XVECT_EQCOUNT_EMPTY;
      Dec(Cnt,EqItemCnt);
      Inc(P);
      Dec(MemSz);

      if AList <> Nil then begin
        if EqItemCnt > 1 then
          AList.Add(Format('[%.4X] Item #%d..%d: Empty',[Offs,i,i + EqItemCnt - 1]))
        else
          AList.Add(Format('[%.4X] Item #%d: Empty',[Offs,i]));
      end;

      Inc(i,EqItemCnt);
      Inc(Offs);
    end
    else begin
      Dec(Cnt,EqItemCnt);
      Inc(P);
      XLate := P^;
      if (XLate and XVECT_XLATE_EX) <> 0 then begin
        Inc(P);
        S := DumpBytes(PByteArray(P),SZ_HDR_EX + PWord(P)^);
        ItemSz := PWord(P)^;
        Inc(P,SZ_HDR_EX);
        Inc(P,ItemSz);
        Dec(MemSz,SZ_HDR_EQ + SZ_HDR_XL + SZ_HDR_EX + ItemSz);

        if AList <> Nil then begin
          AList.Add(Format('[%.4X] Item #%d: [Extended] Size=%d',[Offs,i,ItemSz]));
          AList.Add(S);
        end;

        Inc(i);
        Inc(Offs,SZ_HDR_EQ + SZ_HDR_XL + SZ_HDR_EX + ItemSz);
      end
      else begin
        ItemSz := L_SizeXLate[XLate];
        if ItemSz = 0 then begin
          VectError('Invalid XLate byte');
          Exit;
        end;
        Inc(P,SZ_HDR_XL);
        S := DumpBytes(PByteArray(P),EqItemCnt * ItemSz);
        Inc(P,EqItemCnt * ItemSz);
        Dec(MemSz,SZ_HDR_EQ + SZ_HDR_XL + EqItemCnt * ItemSz);

        if AList <> Nil then begin
          if EqItemCnt > 1 then
            AList.Add(Format('[%.4X] Item #%d..%d: XLate=%d, Size=%d',[Offs,i,i + EqItemCnt - 1,XLate,ItemSz]))
          else
            AList.Add(Format('[%.4X] Item #%d: XLate=%d, Size=%d',[Offs,i,XLate,ItemSz]));
          AList.Add(S);
        end;

        Inc(i,EqItemCnt);
        Inc(Offs,SZ_HDR_EQ + SZ_HDR_XL + EqItemCnt * ItemSz);
      end;
    end;

    if (AList <> Nil) and (Cnt < 0) then begin
      VectError('Item count error');
      Exit;
    end;
    if (AList <> Nil) and (MemSz < 0) then begin
      VectError('Memory size error');
      Exit;
    end;
  end;
  if (AList <> Nil) and (Cnt <> 0) then begin
    raise XLSRWException.Create('Item count error');
  end;
  if (AList <> Nil) and (MemSz <> 0) then begin
    VectError('Memory size error');
  end;
end;

function TXLSMMUVectorManager.AddEmpty(const APtr: TMMUPtr; const ACount: integer): TMMUPtr;
var
  i: integer;
  ExByteCount: integer;
begin
  Result := APtr;
  ExByteCount := (ACount div XVECT_MAX_EQCOUNT);
  if ExByteCount > 0 then begin
    for i := 1 to ExByteCount do begin
      Result^ := XVECT_EQCOUNT_EMPTY or XVECT_MAX_EQCOUNT;
      Inc(Result);
    end;
    if ACount - (ExByteCount * XVECT_MAX_EQCOUNT) > 0 then begin
      Result^ := XVECT_EQCOUNT_EMPTY or (ACount - (ExByteCount * XVECT_MAX_EQCOUNT));
      Inc(Result);
    end;
  end
  else begin
    Result^ := XVECT_EQCOUNT_EMPTY or Byte(ACount);
    Inc(Result);
  end;
end;

{ TXLSMMURebuildVector }

function TXLSMMURebuildVector.Add(const APtr: TMMUPtr): TMMUPtr;
var
  i: integer;
  n: integer;
  Sz: integer;
  XL: byte;
  P: TMMUPtr;
  pMem: TMMUPtr;
begin
  Result := APtr;
  n := Result^ and XVECT_MAX_EQCOUNT;
  if (Result^ and XVECT_EQCOUNT_EMPTY) <> 0 then begin
    for i := 0 to n - 1 do
      AddItem(MAXINT);
    Inc(Result);
  end
  else begin
    Inc(Result);
    XL := Result^;
    if (XL and XVECT_XLATE_EX) <> 0 then begin
      AddItem(FMemSize);
      P := Result;
      Inc(P);
      Sz := PWord(P)^;
      pMem := AddMem(SZ_HDR_XL + SZ_HDR_EX + Sz);

      System.Move(Result^,pMem^,SZ_HDR_XL + SZ_HDR_EX + Sz);
      Inc(Result,SZ_HDR_XL + SZ_HDR_EX + Sz);
    end
    else begin
      Sz := L_SizeXLate[XL];
      Inc(Result);
      for i := 0 to n - 1 do begin
        AddItem(FMemSize);
        pMem := AddMem(SZ_HDR_XL + Sz);
        pMem^ := XL;
        Inc(pMem);
        System.Move(Result^,pMem^,Sz);
        Inc(Result,Sz);
      end;
    end;
  end;
end;

procedure TXLSMMURebuildVector.AddItem(const AOffset: integer);
begin
  FItems[FItemCount] := AOffset;
  Inc(FItemCount);
end;

function TXLSMMURebuildVector.AddMem(const ASize: integer): TMMUPtr;
begin
  Inc(FMemSize,ASize);
  if FMemSize > FAllocSize then begin
    Inc(FAllocSize,XBLOCK_PREALLOCSIZE);
    ReAllocMem(FMem,FAllocSize);
  end;
  Result := FMem;
  Inc(Result,FMemSize - ASize);
end;

function TXLSMMURebuildVector.Address(const AIndex: integer): TMMUPtr;
begin
  Result := FMem;
  Inc(Result,FItems[AIndex]);
end;

constructor TXLSMMURebuildVector.Create(const AVector: TMMUPtr);
var
  Cnt: integer;
  P: TMMUPtr;
begin
  FItemCount := 0;
  FAllocSize := XBLOCK_PREALLOCSIZE;
  GetMem(FMem,FAllocSize);

  if AVector <> Nil then begin
    P := AVector;
    Cnt := PXVectHeader(P).ItemCount;
    Inc(P,SizeOf(TXVectHeader) + L_ItemsHeaderSize);
    while Cnt > 0 do begin
      Dec(Cnt,P^ and XVECT_MAX_EQCOUNT);
      P := Add(P);
    end;
  end;
end;

function TXLSMMURebuildVector.DeleteItem(const AItem: integer): boolean;
begin
  Result := AItem < FItemCount;
  if Result then begin
    Result := FItems[AItem] <> MAXINT;
    if Result then begin
      FItems[AItem] := MAXINT;
      if AItem = (FItemCount - 1) then begin
        Dec(FItemCount);
        while (FItemCount > 0) and (FItems[FItemCount - 1] = MAXINT) do
          Dec(FItemCount);
      end;
    end;
  end;
end;

destructor TXLSMMURebuildVector.Destroy;
begin
  FreeMem(FMem);
  inherited;
end;

procedure TXLSMMURebuildVector.InsertEmpty(const AItem, ACount: integer);
var
  i1,i2: integer;
begin
  i1 := FItemCount - 1;
  i2 := FItemCount - 1 + ACount;
  while i1 >= AItem do begin
    if i2 < XVECT_MAX_ITEMS then begin
      FItems[i2] := FItems[i1];
      Dec(i1);
      Dec(i2);
    end;
  end;
  i2 := AItem + ACount - 1;
  if i2 >= XVECT_MAX_ITEMS then
    i2 := XVECT_MAX_ITEMS - 1;
  for i1 := AItem to i2 do
    FItems[i1] := MAXINT;
  Inc(FItemCount,ACount);
end;

procedure TXLSMMURebuildVector.InsertExItem(const AItem: integer; const AXLate: byte; const ASize: integer);
var
  P: TMMUPtr;
begin
  P := FMem;
  Inc(P,FMemSize);
  if AItem < FItemCount then
    FItems[AItem] := FMemSize
  else
    AddItem(FMemSize);
  AddMem(SZ_HDR_XL + SZ_HDR_EX + ASize);
  P^ := XVECT_XLATE_EX or AXLate;
  Inc(P);
  PWord(P)^ := ASize;
end;

procedure TXLSMMURebuildVector.InsertXLateItem(const AItem: integer; const AXLate: byte);
var
  P: TMMUPtr;
begin
  P := FMem;
  Inc(P,FMemSize);
  if AItem < FItemCount then
    FItems[AItem] := FMemSize
  else
    AddItem(FMemSize);
  AddMem(SZ_HDR_XL + L_SizeXLate[AXLate]);
  P^ := AXLate;
end;

function TXLSMMURebuildVector.MoveItems(const ASrcItem, ADestItem: integer): integer;
var
  i: integer;
  iMax: integer;
  Delta,D: integer;
begin
  Result := 0;

  Delta := ADestItem - ASrcItem;
  if Delta > 0 then begin
    iMax := FItemCount - 1 + Delta;
    if iMax > High(FItems) then
      iMax := High(FItems);
    for i := iMax downto ADestItem do begin
      FItems[i] := FItems[i + Delta];
      Inc(Result);
    end;
    for i := ASrcItem to ASrcItem + Delta do begin
      FItems[i] := MAXINT;
      Inc(Result);
    end;
  end
  else if Delta < 0 then begin
    D := Abs(Delta);
    for i := ADestItem to FItemCount - D - 1 do begin
      FItems[i] := FItems[i + D];
      Inc(Result);
    end;
  end;
  Inc(FItemCount,Delta);
  while (FItemCount >= 0) and (FItems[FItemCount] = MAXINT) do
    Dec(FItemCount);
end;

function TXLSMMURebuildVector.RequiredMem: integer;
var
  i: integer;
  n: integer;
  Sz: integer;
  Ex: integer;
  XL: byte;
  P: TMMUPtr;
begin
  Result := SizeOf(TXVectHeader) + L_ItemsHeaderSize;
  i := 0;
  while i < FItemCount do begin

    //#! Empty items
    n := 0;
    while FItems[i] = MAXINT do begin
      Inc(n);
      Inc(i);
    end;
    if n > 0 then
      Inc(Result,SZ_HDR_EQ + (n div XVECT_MAX_EQCOUNT));

    if i < FItemCount then begin
      P := TMMUPtr(FItems[i]);
      Inc(P,NativeInt(FMem));
      Inc(i);
      XL := P^;
      if (P^ and XVECT_XLATE_EX) <> 0 then begin
        Inc(P);
        Inc(Result,SZ_HDR_EQ + SZ_HDR_XL + SZ_HDR_EX + PWord(P)^);
      end
      else begin
        Sz := L_SizeXLate[XL];
        n := 1;
        while (i < FItemCount) and (FItems[i] <> MAXINT) do begin
          P := TMMUPtr(FItems[i]);
          Inc(P,NativeInt(FMem));
          if P^ <> XL then
            Break;
          Inc(n);
          Inc(i);
        end;
        Ex := n div XVECT_MAX_EQCOUNT;
        if Ex > 0 then begin
          Inc(Result,(SZ_HDR_EQ + SZ_HDR_XL) * Ex + Ex * Sz * XVECT_MAX_EQCOUNT);
          Dec(n,Ex * XVECT_MAX_EQCOUNT);
          if n > 0 then
            Inc(Result,SZ_HDR_EQ + SZ_HDR_XL + n * Sz);
        end
        else
          Inc(Result,SZ_HDR_EQ + SZ_HDR_XL + Sz * n);
      end;
    end;
  end;
end;

function TXLSMMURebuildVector.WriteVector(const AVector: TMMUPtr; const AItem: integer): TMMUPtr;
var
  i: integer;
  n: integer;
  Ex: integer;
  Sz: integer;
  XL: byte;
  P: TMMUPtr;
  pEQ: TMMUPtr;
  pVect: TMMUPtr;
begin
  Result := Nil;
  pEQ := AVector;
  pVect := AVector;
  Inc(pVect,SizeOf(TXVectHeader) + L_ItemsHeaderSize);
  i := 0;
  while i < FItemCount do begin
    n := 0;
    while FItems[i] = MAXINT do begin
      Inc(n);
      Inc(i);
    end;
    if n > 0 then begin
      Ex := (n div XVECT_MAX_EQCOUNT);
      while Ex > 0 do begin
        pVect^ := XVECT_EQCOUNT_EMPTY or XVECT_MAX_EQCOUNT;
        Inc(pVect);
        Dec(Ex);
        Dec(n,XVECT_MAX_EQCOUNT);
      end;
      if n > 0 then begin
        pVect^ := XVECT_EQCOUNT_EMPTY or n;
        Inc(pVect)
      end;
    end;
    if i < FItemCount then begin
      P := Address(i);
      XL := P^;
      Inc(P);
      pEQ := pVect;
      if (XL and XVECT_XLATE_EX) <> 0 then begin
        pVect^ := 1;
        Inc(pVect);
        pVect^ := XL;
        Inc(pVect);
        Sz := PWord(P)^;
        Inc(P,SZ_HDR_EX);
        PWord(pVect)^ := Sz;
        Inc(pVect,SZ_HDR_EX);
        Move(P^,pVect^,Sz);

        //#! Result shall point to the extended size word.
        if AItem = i then begin
          Result := pVect;
          Dec(Result,SZ_HDR_EX);
        end;

        Inc(i);
        Inc(pVect,Sz);
      end
      else begin
{$ifdef MMU_DEBUG}
        if (XL = 0) or (L_SizeXLate[XL] = 0) then
          raise XLSRWException.Create('Invalid XLate byte');
{$endif}
        Inc(pVect);
        pVect^ := XL;
        Inc(pVect);

        if AItem = i then
          Result := pVect;

        Sz := L_SizeXLate[XL];
        case Sz of
          2: begin
            PWord(pVect)^ := PWord(P)^;
            Inc(pVect,2);
          end;
          4: begin
            PLongword(pVect)^ := PLongword(P)^;
            Inc(pVect,4);
          end;
          6: begin
             PWord(pVect)^ := PWord(P)^;
             Inc(P,2);
             Inc(pVect,2);
             PLongword(pVect)^ := PLongword(P)^;
             Inc(pVect,4);
          end;
          8: begin
            PDouble(pVect)^ := PDouble(P)^;
            Inc(pVect,8);
          end;
          10: begin
             PWord(pVect)^ := PWord(P)^;
             Inc(P,2);
             Inc(pVect,2);
             PDouble(pVect)^ := PDouble(P)^;
             Inc(pVect,8);
          end;
          else begin
            System.Move(P^,pVect^,Sz);
            Inc(pVect,Sz);
          end;
        end;

        Inc(i);
        n := 1;
        while (i < FItemCount) and (FItems[i] <> MAXINT) and (Address(i)^ = XL) and (n < XVECT_MAX_EQCOUNT) do begin
          P := Address(i);
          Inc(P);
          if AItem = i then
            Result := pVect;
          case Sz of
            2: begin
              PWord(pVect)^ := PWord(P)^;
              Inc(pVect,2);
            end;
            4: begin
              PLongword(pVect)^ := PLongword(P)^;
              Inc(pVect,4);
            end;
            6: begin
               PWord(pVect)^ := PWord(P)^;
               Inc(P,2);
               Inc(pVect,2);
               PLongword(pVect)^ := PLongword(P)^;
               Inc(pVect,4);
            end;
            8: begin
              PDouble(pVect)^ := PDouble(P)^;
              Inc(pVect,8);
            end;
            10: begin
               PWord(pVect)^ := PWord(P)^;
               Inc(P,2);
               Inc(pVect,2);
               PDouble(pVect)^ := PDouble(P)^;
               Inc(pVect,8);
            end;
            else begin
              System.Move(P^,pVect^,Sz);
              Inc(pVect,Sz);
            end;
          end;
          Inc(n);
          Inc(i);
        end;
        pEQ^ := n;
      end;
    end;
  end;
  PXVectHeader(AVector).ItemCount := FItemCount;
  PXVectHeader(AVector).OffsetToLastEq := Integer(pEQ) - Integer(AVector);
  FWrittenMem := Integer(NativeInt(pVect) - NativeInt(AVector));
end;

{ TXLSMMUVectorIterator }

procedure TXLSMMUVectorIterator.BeginSlot;
begin
  FSlotCount := PByte(FItem)^ and XVECT_MAX_EQCOUNT;
  FIsEmpty := (PByte(FItem)^ and XVECT_EQCOUNT_EMPTY) <> 0;
  if not FIsEmpty then
    FItemXL := PByteArray(FItem)[1];
end;

procedure TXLSMMUVectorIterator.Clear(const AVector: TMMUPtr);
begin
  FVector := AVector;
  if (AVector = Nil) or (PXVectHeader(FVector).ItemCount <= 0) then begin
    FEQIndex := -1;
    FSlotCount := 0;
  end
  else begin
    FItem := FVector;
    Inc(FItem,SizeOf(TXVectHeader) + L_ItemsHeaderSize);
    FEQIndex := 0;
    BeginSlot;
  end;
end;

function TXLSMMUVectorIterator.Find(var AItem: integer): boolean;
begin
  if (AItem < 0) or (AItem >= PXVectHeader(FVector).ItemCount) then
    Result := False
  else if AItem < FEQIndex then begin
    Clear(FVector);
    Result := Find(AItem);
  end
  else begin
    Result := True;
    while True do begin
      if HitSlot(AItem) then begin
        if FIsEmpty then begin
          AItem := FEQIndex + FSlotCount - 1;
          Result := False;
        end
        else if (FItemXL and XVECT_XLATE_EX) <> 0 then begin
          FResItem := FItem;
          Inc(FResItem,SZ_HDR_EQ + SZ_HDR_XL);
        end
        else begin
          FResItem := FItem;
          Inc(FResItem,SZ_HDR_EQ + SZ_HDR_XL + (AItem - FEQIndex) * L_SizeXLate[FItemXL]);
        end;
        Exit;
      end
      else begin
        Inc(FEQIndex,FSlotCount);
        if FIsEmpty then
          Inc(FItem,SZ_HDR_EQ)
        else begin
          if (FItemXL and XVECT_XLATE_EX) <> 0 then begin
            Inc(FItem,SZ_HDR_EQ + SZ_HDR_XL);
            Inc(FItem,SZ_HDR_EX + PWord(FItem)^);
          end
          else
            Inc(FItem,SZ_HDR_EQ + SZ_HDR_XL + FSlotCount * L_SizeXLate[FItemXL]);
        end;
        BeginSlot;
      end;
    end;
  end;
end;

function TXLSMMUVectorIterator.FindNext(var AItem: integer): boolean;
begin
  Result := False;
  while (AItem <= PXVectHeader(FVector).ItemCount) and not Result do begin
    Result := Find(AItem);
    if not Result then
      Inc(AItem);
  end;
end;

function TXLSMMUVectorIterator.HitSlot(const AItem: integer): boolean;
begin
  Result := (AItem >= FEQIndex) and (AItem < (FEQIndex + FSlotCount));
end;

initialization
  L_ItemsHeaderDefaultData := Nil;

finalization
  FreeMem(L_ItemsHeaderDefaultData);

end.
