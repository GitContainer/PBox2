unit XLSTSort;

{$I AxCompilers.inc}
{$I XLSRWII.inc}

interface

uses Classes, SysUtils, Contnrs, Math,
     Xc12Utils5,
     XLSUtils5, XLSCellMMU5;

const TOPOSORT_TEXT_LOOP = '** Loop **';

type TDepRefArray = array[0..10] of TXLS3dCompactRef;
     PDepRefArray = ^TDepRefArray;

type TDepRefsItem = class(TObject)
private
     function GetItems(Index: integer): PXLS3dCompactRef;
protected
     FData: PDepRefArray;
     FCount: integer;
     FAllocBlocks: word;
public
     constructor Create(ADep: PXLS3dCompactRef);
     destructor Destroy; override;

     procedure AddRef(ADep: PXLS3dCompactRef); {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     procedure EndAddRefs; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}

     function  AsText: AxUCString;

     property Count: integer read FCount;
     property Items[Index: integer]: PXLS3dCompactRef read GetItems; default;
     end;

type TDepRefsItems = class(TObjectList)
private
     function GetItems(Index: integer): TDepRefsItem;
protected
public
     constructor Create;

     function  Add(ADep: PXLS3dCompactRef): TDepRefsItem; overload;
     procedure Add(ADep: TDepRefsItem); overload;
     function  Last: TDepRefsItem;
     function  AsText: AxUCString;

     property Items[Index: integer]: TDepRefsItem read GetItems; default;
     end;

type TTopoItem = class;

     TSuccessor = class(TObject)
protected
     FSuc: TTopoItem;
     FNext: TSuccessor;
public
     property Suc: TTopoItem read FSuc write FSuc;
     property Next: TSuccessor read FNext write FNext;
     end;

     TTopoItem = class(TObject)
protected
     FDep: PXLS3dCompactRef;
     FLeft,FRight: TTopoItem;
     FCount: integer;
     FBalance: integer;
     FQLink: TTopoItem;
     FTop: TSuccessor;
public
     property Dep: PXLS3dCompactRef read FDep write FDep;
     property Left: TTopoItem read FLeft write FLeft;
     property Right: TTopoItem read FRight write FRight;
     property Balance: integer read FBalance write FBalance;
     property Count: integer read FCount write FCount;
     property QLink: TTopoItem read FQLink write FQLink;
     property Top: TSuccessor read FTop write FTop;
     end;

type TTSortCallback = function(AItem: TTopoItem): boolean of object;

type TTopoSort = class(TObject)
protected
     FRoot: TTopoItem;
     FHead: TTopoItem;
     FZeros: TTopoItem;
     FLoop: TTopoItem;
     FNStrings: integer;

     FGarbages: TObjectList;

     FOutput: TList;
     FCircular: TDepRefsItems;

     function  NewItem(ADep: PXLS3dCompactRef): TTopoItem;
     function  SearchItem(ARoot: TTopoItem; ADep: PXLS3dCompactRef): TTopoItem;
     procedure RecordRelation(j,k: TTopoItem);
     function  CountItems(AItem: TTopoItem): boolean;
     function  ScanZeros(k: TTopoItem): boolean;
     function  DetectLoop(k: TTopoItem): boolean;
     function  RecurseTree(ARoot: TTopoItem; ACallback: TTSortCallback): boolean;
     procedure WalkTree(ARoot: TTopoItem; ACallback: TTSortCallback);
     function  DoSort(AList: TList): boolean;
public
     constructor Create;
     destructor Destroy; override;

     function AsText: AxUCString;

     procedure Clear;

{  D5 pukes on these comments...
     // Format A B C D ...
     // A is dependent on  B C D
}
     function  Sort(AList: TDepRefsItems): boolean;

     property Output: TList read FOutput;
     property Circular: TDepRefsItems read FCircular;
     end;

implementation

function DepEqual(ADep1,ADep2: PXLS3dCompactRef): boolean; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
begin
  Result := (ADep1.SheetId = ADep2.SheetId) and (ADep1.Row = ADep2.Row) and (ADep1.Col = ADep2.Col);
end;

function DepCompare(ADep1,ADep2: PXLS3dCompactRef): integer; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
begin
  Result := ADep1.SheetId - ADep2.SheetId;
  if Result = 0 then
    Result := ADep1.Row - ADep2.Row;
    if Result = 0 then
      Result := ADep1.Col - ADep2.Col;
end;

{ TTopoSort }

function TTopoSort.AsText: AxUCString;
var
  i: integer;
begin
  Result := '';

  for i := FOutput.Count - 1 downto 0 do begin
//    if (PDepRef(FOutput[i]).Col1 = PDepRef(FOutput[i]).Col2) and (PDepRef(FOutput[i]).Row1 = PDepRef(FOutput[i]).Row2) then
      Result := Result + ColRowToRefStr(PXLS3dCompactRef(FOutput[i]).Col,PXLS3dCompactRef(FOutput[i]).Row) + '->'
//    else
//      Result := Result + AreaToRefStr(PDepRef(FOutput[i]).Col1,PDepRef(FOutput[i]).Row1,PDepRef(FOutput[i]).Col2,PDepRef(FOutput[i]).Row2) + '->';
  end;
  Result := Copy(Result,1,Length(Result) - 2);
end;

procedure TTopoSort.Clear;
begin
  FGarbages.Clear;

  FOutput.Clear;
  FCircular.Clear;

  FRoot := Nil;
  FHead := Nil;
  FZeros := Nil;
  FLoop := Nil;
  FNStrings := 0;
end;

function TTopoSort.CountItems(AItem: TTopoItem): boolean;
begin
  Inc(FNStrings);
  Result := False;
end;

constructor TTopoSort.Create;
begin
  FOutput := TList.Create;
  FCircular := TDepRefsItems.Create;

  FGarbages := TObjectList.Create;
end;

destructor TTopoSort.Destroy;
begin
  FOutput.Free;
  FCircular.Free;

  FGarbages.Free;

  inherited;
end;

function TTopoSort.DetectLoop(k: TTopoItem): boolean;
var
  C: TDepRefsItem;
  p,last: TSuccessor;
  tmp: TTopoItem;
begin
  if k.Count > 0 then begin
    if FLoop = Nil then
      FLoop := k
    else begin
	    p := k.Top;
      last := Nil;
      while p <> Nil do begin
        if p.Suc = FLoop then begin
          if k.Qlink <> Nil then begin
            C := FCircular.Add(PXLS3dCompactRef(Nil));
            while FLoop <> Nil do begin
              tmp := FLoop.qlink;

              C.AddRef(FLoop.Dep);
              if FLoop = k then begin
                p.Suc.Count := p.Suc.Count - 1;

                // TODO Is this correct?
                // Original code: *p = (*p)->next;
                if last <> Nil then
                  last.Next := p.Next;

			          break;
              end;
              FLoop.QLink := Nil;
              FLoop := tmp;
            end;
            while FLoop <> Nil do begin
              tmp := FLoop.QLink;

              FLoop.QLink := Nil;
              FLoop := tmp;
            end;
            C.EndAddRefs;
            Result := True;
            Exit;
          end
          else begin
            k.QLink := FLoop;
            FLoop := k;
            break;
          end;
        end;
        last := p;
        p := p.Next;
      end;
    end;
  end;
  Result := False;
end;

function TTopoSort.NewItem(ADep: PXLS3dCompactRef): TTopoItem;
begin
  Result := TTopoItem.Create;
  Result.Dep := ADep;
  Result.Left := Nil;
  Result.Right := Nil;
  Result.Balance := 0;
  Result.Count := 0;
  Result.QLink := Nil;
  Result.Top := Nil;

  FGarbages.Add(Result);
end;

procedure TTopoSort.RecordRelation(j, k: TTopoItem);
var
  P: TSuccessor;
begin
  if not DepEqual(j.Dep, k.Dep) then begin
    k.Count := k.Count + 1;
    p := TSuccessor.Create;
    FGarbages.Add(p);
    p.Suc := k;
    p.Next := j.Top;
    j.Top := p;
  end;
end;

function TTopoSort.RecurseTree(ARoot: TTopoItem; ACallback: TTSortCallback): boolean;
begin
  if (ARoot.Left = Nil) and (ARoot.Right = Nil) then begin
    Result := ACallback(ARoot);
    Exit;
  end
  else begin
    if ARoot.Left <> Nil then begin
      if RecurseTree(ARoot.Left, ACallback) then begin
        Result := True;
        Exit
      end;
    end;
    if ACallback(ARoot) then begin
      Result := True;
      Exit;
    end;
    if ARoot.Right <> Nil then begin
      if RecurseTree(ARoot.right, ACallback) then begin
        Result := True;
        Exit;
      end;
    end;
  end;
  Result := False;
end;

function TTopoSort.ScanZeros(k: TTopoItem): boolean;
begin
  if (k.Count = 0) and (k.Dep <> Nil) then begin
    if FHead = Nil then
      FHead := k
    else
      FZeros.QLink := k;
    FZeros := k;
  end;
  Result := False;
end;

function TTopoSort.SearchItem(ARoot: TTopoItem; ADep: PXLS3dCompactRef): TTopoItem;
var
  p,q,r,s,t: TTopoItem;
  a: integer;
begin
  if ARoot.Right = Nil then begin
    ARoot.Right := NewItem(ADep);
    Result := ARoot.Right;
    Exit;
  end;

  t := ARoot;
  p := ARoot.Right;
  s := p;

  while True  do begin
    a := DepCompare(ADep,p.Dep);
    if a = 0 then begin
      Result := p;
      Exit;
    end;
    if a < 0 then
	    q := p.Left
    else
	    q := p.Right;

    if q = Nil then begin

      q := NewItem(ADep);

	    if a < 0 then
  	    p.Left := q
	    else
	      p.Right := q;

      if DepCompare(ADep, s.Dep) < 0 then begin
        r := s.Left;
        p := r;
        a := -1;
      end
      else begin
        r := s.Right;
        p := r;
        a := 1;
      end;

	    while p <> q do begin
        if DepCompare(ADep, p.Dep) < 0 then begin
          p.Balance := -1;
          p := p.Left;
        end
	      else begin
          p.Balance := 1;
          p := p.Right;
        end;
	    end;

	    if (s.Balance = 0) or (s.Balance = -a) then begin
	      s.Balance := s.Balance + a;
	      Result := q;
        Exit;
	    end;
      if r.Balance = a then begin
        p := r;
        if a < 0 then begin
          s.Left := r.Right;
          r.Right := s;
        end
	      else begin
          s.Right := r.Left;
          r.Left := s;
		    end;
        r.Balance := 0;
	      s.Balance := 0;
	    end
	    else begin
	      if a < 0 then begin
          p := r.Right;
          r.Right := p.Left;
          p.Left := r;
          s.Left := p.Right;
          p.Right := s;
        end
	      else begin
          p := r.Left;
          r.Left := p.Right;
          p.Right := r;
          s.Right := p.Left;
          p.Left := s;
        end;

	      s.Balance := 0;
	      r.Balance := 0;
        if p.Balance = a then
          s.Balance := -a
        else if p.Balance = -a then
		      r.Balance := a;
	      p.Balance := 0;
	    end;

      if s = t.Right then
        t.Right := p
      else
        t.Left := p;

      Result := q;
      Exit;
    end;

    if q.Balance <> 0 then begin
	    t := p;
	    s := q;
    end;
    p := q;
  end;
end;

function TTopoSort.Sort(AList: TDepRefsItems): boolean;
var
  i,j: integer;
  Item: TDepRefsItem;
  D2,D3: PXLS3dCompactRef;
  L: TList;
begin
  L := TList.Create;
  try
    for i := 0 to AList.Count - 1 do begin
      Item := AList[i];
      D2 := Item[0];
      for j := 1 to Item.Count - 1 do begin
        D3 := Item[j];
//        // This is not considered to be a loop by the sorter...
//        if S2 = S3 then begin
//          FCircular.Add(TOPOSORT_TEXT_LOOP);
//          FCircular.Add(S2);
//          FCircular.Add(S3);
//        end;
        L.Add(D2);
        L.Add(D3);
      end;
    end;
    Result := DoSort(L);
  finally
    L.Free;
  end;
end;

function TTopoSort.DoSort(AList: TList): boolean;
var
  i: integer;
  D: PXLS3dCompactRef;
  ok: boolean;
  j,k: TTopoItem;
  p: TSuccessor;
begin
  i := 0;
  ok := true;
  j := Nil;
  k := Nil;

  FRoot := NewItem(Nil);

  while True do begin
    if i >= AList.Count then
      Break;
    D := AList[i];
    Inc(i);
    k := SearchItem (FRoot,D);
    if j <> Nil then begin
	    RecordRelation (j, k);
	    k := Nil;
    end;
    j := k;
  end;

  if k <> Nil then
    raise XLSRWException.Create('Odd number items in input');

  WalkTree(FRoot,CountItems);

  while FNstrings > 0 do begin
    WalkTree(FRoot,ScanZeros);
    while FHead <> Nil do begin
	    p := FHead.Top;

      FOutput.Add(FHead.Dep);
	    FHead.Dep := Nil;
	    Dec(FNstrings);

      while p <> Nil do begin
        p.Suc.Count := p.Suc.Count - 1;

	      if p.Suc.Count = 0 then begin
          FZeros.QLink := p.Suc;
          FZeros := p.Suc;
        end;
        p := p.Next;
      end;

    FHead := FHead.QLink;
    end;

    if FNStrings > 0 then begin
      ok := false;

      repeat
        WalkTree (FRoot,DetectLoop);
      until FLoop = Nil;
    end;
  end;
  Result := ok;
end;

procedure TTopoSort.WalkTree(ARoot: TTopoItem; ACallback: TTSortCallback);
begin
  if ARoot.Right <> Nil then
    RecurseTree(ARoot.Right,ACallback);
end;

{ TDepRefsItem }

procedure TDepRefsItem.AddRef(ADep: PXLS3dCompactRef);
begin
  if FCount >= (FAllocBlocks * $FF) then begin
    Inc(FAllocBlocks);
    ReAllocMem(FData,FAllocBlocks * $FF * SizeOf(TXLS3dCompactRef));
  end;
  FData[FCount] := ADep^;
  Inc(FCount);
end;

function TDepRefsItem.AsText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := FCount - 1 downto 0 do begin
//    if (Items[i].Col1 = Items[i].Col2) and (Items[i].Row1 = Items[i].Row2) then
      Result := IntToStr(Items[i].SheetId) + '!' + ColRowToRefStr(Items[i].Col,Items[i].Row) + '->' + Result
//    else
//      Result := IntToStr(Items[i].SheetId) + '!' + AreaToRefStr(Items[i].Col1,Items[i].Row1,Items[i].Col2,Items[i].Row2) + '->' + Result;
  end;
  Result := Copy(Result,1,Length(Result) - 2);
end;

constructor TDepRefsItem.Create(ADep: PXLS3dCompactRef);
begin
  if ADep <> Nil then
    AddRef(ADep);
end;

destructor TDepRefsItem.Destroy;
begin
  FreeMem(FData);
  inherited;
end;

procedure TDepRefsItem.EndAddRefs;
begin
  ReAllocMem(FData,FCount * SizeOf(TXLS3dCompactRef));
end;

function TDepRefsItem.GetItems(Index: integer): PXLS3dCompactRef;
begin
  Result := @FData[Index];
end;

{ TDepRefsItems }

function TDepRefsItems.Add(ADep: PXLS3dCompactRef): TDepRefsItem;
begin
  Result := TDepRefsItem.Create(ADep);
  inherited Add(Result);
end;

procedure TDepRefsItems.Add(ADep: TDepRefsItem);
begin
  inherited Add(ADep);
end;

function TDepRefsItems.AsText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to Count - 1 do
    Result := Result + Items[i].AsText + #13#10;
end;

constructor TDepRefsItems.Create;
begin
  inherited Create;
end;

function TDepRefsItems.GetItems(Index: integer): TDepRefsItem;
begin
  Result := TDepRefsItem(inherited Items[Index]);
end;

function TDepRefsItems.Last: TDepRefsItem;
begin
  Result := Items[Count - 1];
end;

end.

