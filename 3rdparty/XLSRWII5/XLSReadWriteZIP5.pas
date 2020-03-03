unit XLSReadWriteZIP5;

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

uses SysUtils, Classes, Contnrs,
{$ifdef DELPHI_XE2_OR_LATER}
     Zip,
{$endif}
     XLSUtils5, XLSZlibPas5, XLSZip5;

type TXLSOwnerStream = class(TStream)
protected
     FZipStream: TXLSZipFile;
     FId: string;

{$ifdef DELPHI_7_OR_LATER}
     function GetSize: Int64; override;
{$endif}
public
     constructor Create(const AZipStream: TXLSZipFile; AId: string);
     destructor Destroy; override;

     function Read(var Buffer; Count: Longint): Longint; overload; override;
     function Write(const Buffer; Count: Longint): Longint; overload; override;
{$ifdef DELPHI_5}
     function Seek(Offset: Longint; Origin: Word): Longint; overload; override;
{$else}
     function Seek(const Offset: Int64; Origin: TSeekOrigin): Int64; overload; override;
{$endif}
     end;

type TXLSZipArchiveIntf = class(TObject)
private
     FZipStream: TFileStream;
public
     destructor Destroy; override;

     function  Count: integer; virtual; abstract;
     procedure LoadFromStream(Stream: TStream); virtual; abstract;
     procedure LoadFromFile(const AFilename: AxUCString);
     function  Find(Filename: AxUCString): integer; virtual; abstract;

     function  OpenStream(Filename: AxUCString): TStream; virtual; abstract;

     function  CreateStream(Filename: AxUCString): TStream; overload; virtual; abstract;
     function  CreateStream(FolderName, Filename: AxUCString): TStream; overload; virtual; abstract;
     procedure CloseStream(AStream: TStream); virtual; abstract;

     procedure Write(Filename: AxUCString; Source: TStream); overload; virtual; abstract;
     procedure Write(FolderName,Filename: AxUCString; Source: TStream); overload; virtual; abstract;
     procedure SaveToStream(Stream: TStream); virtual; abstract;
     procedure Close; virtual; abstract;
     end;

type TXLSZipArchiveIntfXLS = class(TXLSZipArchiveIntf)
private
     FZip: TXLSZipArchive;
public
     constructor Create;
     destructor Destroy; override;

     function  Count: integer; override;
     procedure LoadFromStream(Stream: TStream); override;
     function  Find(Filename: AxUCString): integer; override;

     function  OpenStream(Filename: AxUCString): TStream; override;

     function  CreateStream(Filename: AxUCString): TStream; overload; override;
     function  CreateStream(FolderName, Filename: AxUCString): TStream; overload; override;
     procedure CloseStream(AStream: TStream); override;

     procedure Write(Filename: AxUCString; Source: TStream); overload; override;
     procedure Write(FolderName,Filename: AxUCString; Source: TStream); overload; override;
     procedure SaveToStream(Stream: TStream); override;
     procedure Close; override;
     end;

{$ifdef DELPHI_XE2_OR_LATER}
type TXLSZipArchiveIntfDelphi = class(TXLSZipArchiveIntf)
private
     FZip: TZipFile;
     FCurrFilename: AxUCString;
     FWriteStream: TStream;
public
     constructor Create;
     destructor Destroy; override;

     function  Count: integer; override;
     procedure LoadFromStream(Stream: TStream); override;
     function  Find(Filename: AxUCString): integer; override;

     function  OpenStream(Filename: AxUCString): TStream; override;

     function  CreateStream(Filename: AxUCString): TStream; overload; override;
     function  CreateStream(FolderName, Filename: AxUCString): TStream; overload; override;
     procedure CloseStream(AStream: TStream); override;

     procedure Write(Filename: AxUCString; Source: TStream); overload; override;
     procedure Write(FolderName,Filename: AxUCString; Source: TStream); overload; override;
     procedure SaveToStream(Stream: TStream); override;
     procedure Close; override;
     end;
{$endif}

implementation

{ TXLSZipArchive }

procedure TXLSZipArchiveIntfXLS.Close;
begin
  FZip.Close;
end;

procedure TXLSZipArchiveIntfXLS.CloseStream(AStream: TStream);
begin
  if not (AStream is TCompressionStream) then
    raise XLSRWException.Create('Stream is not TCompressionStream');
  FZip.CloseStream(TCompressionStream(AStream));
end;

function TXLSZipArchiveIntfXLS.Count: integer;
begin
  Result := FZip.Count;
end;

constructor TXLSZipArchiveIntfXLS.Create;
begin
  FZip := TXLSZipArchive.Create;
end;

function TXLSZipArchiveIntfXLS.CreateStream(FolderName, Filename: AxUCString): TStream;
begin
  if (FolderName <> '') and (FolderName[Length(FolderName)] = '/') then
    FolderName := Copy(FolderName,1,Length(FolderName) - 1);
  if (Filename <> '') and (Filename[1] = '/') then
    Filename := Copy(Filename,2,MAXINT);

  if FolderName <> '' then
    Result := FZip.CreateStream(FolderName + '/' + Filename)
  else
    Result := FZip.CreateStream(Filename);
end;

function TXLSZipArchiveIntfXLS.CreateStream(Filename: AxUCString): TStream;
var
  p: integer;
  FolderName: AxUCString;
begin
  p := RCPos('/',Filename);
  if p > 1 then begin
    FolderName := Copy(Filename,1,p - 1);
    Filename := Copy(Filename,p + 1,MAXINT);
  end
  else
    FolderName := '';

  Result := CreateStream(FolderName,Filename);
end;

destructor TXLSZipArchiveIntfXLS.Destroy;
begin
  FZip.Free;
  inherited;
end;

function TXLSZipArchiveIntfXLS.Find(Filename: AxUCString): integer;
begin
  Result := FZip.Find(Filename);
end;

function TXLSZipArchiveIntfXLS.OpenStream(Filename: AxUCString): TStream;
var
  i: integer;
begin
  i := Find(Filename);
  if i < 0 then
    raise XLSRWException.CreateFmt('Can not find file %s in ZIP',[Filename]);
  Result := TXLSOwnerStream.Create(FZip[i],Filename);
end;

procedure TXLSZipArchiveIntfXLS.LoadFromStream(Stream: TStream);
begin
  FZip.OpenRead(Stream);
end;

procedure TXLSZipArchiveIntfXLS.SaveToStream(Stream: TStream);
begin
  FZip.OpenWrite(Stream);
end;

procedure TXLSZipArchiveIntfXLS.Write(Filename: AxUCString; Source: TStream);
var
  p: integer;
  FolderName: AxUCString;
begin
  p := RCPos('/',Filename);
  if p > 1 then begin
    FolderName := Copy(Filename,1,p - 1);
    Filename := Copy(Filename,p + 1,MAXINT);
  end
  else
    FolderName := '';
  Write(FolderName,Filename,Source);
end;

procedure TXLSZipArchiveIntfXLS.Write(FolderName,Filename: AxUCString; Source: TStream);
begin
  if (FolderName <> '') and (FolderName[Length(FolderName)] = '/') then
    FolderName := Copy(FolderName,1,Length(FolderName) - 1);
  if (Filename <> '') and (Filename[1] = '/') then
    Filename := Copy(Filename,2,MAXINT);

  Source.Seek(0,soFromBeginning);

  if FolderName <> '' then
    FZip.AddStream(FolderName + '/' + Filename,Source)
  else
    FZip.AddStream(Filename,Source);
end;

{ TXLSOwnerStream }

constructor TXLSOwnerStream.Create(const AZipStream: TXLSZipFile; AId: string);
begin
  FId := AId;
  FZipStream := AZipStream;
  FZipStream.OpenRead;
end;

destructor TXLSOwnerStream.Destroy;
begin
  FZipStream.Close;
  inherited;
end;

{$ifdef DELPHI_7_OR_LATER}
function TXLSOwnerStream.GetSize: Int64;
begin
  Result := FZipStream.UncompressedSize;
end;
{$endif}

function TXLSOwnerStream.Read(var Buffer; Count: Integer): Longint;
begin
  Result := FZipStream.Read(Buffer,Count);
end;

{$ifdef DELPHI_5}
function TXLSOwnerStream.Seek(Offset: Integer; Origin: Word): Longint;
{$else}
function TXLSOwnerStream.Seek(const Offset: Int64; Origin: TSeekOrigin): Int64;
{$endif}
begin
  Result := FZipStream.Seek(Offset,Origin);
end;

function TXLSOwnerStream.Write(const Buffer; Count: Integer): Longint;
begin
  Result := FZipStream.Write(Buffer,Count);
end;

{$ifdef DELPHI_XE2_OR_LATER}

{ TXLSZipArchiveIntfDelphi }

procedure TXLSZipArchiveIntfDelphi.Close;
var
  Magic: array[0..3] of byte;
begin
  FZip.Close;

  // Zip signature must be written at the beginning of the file, otherwise
  // Excel will not open it. Even with signature Excel will give an error
  // mesage but the file will open.
  if FWriteStream <> Nil then begin
    Magic[0] := $50;
    Magic[1] := $4B;
    Magic[2] := $03;
    Magic[3] := $04;
    FWriteStream.Seek(0,soFromBeginning);
    FWriteStream.Write(Magic,Length(Magic));
  end;
end;

procedure TXLSZipArchiveIntfDelphi.CloseStream(AStream: TStream);
begin
  if FCurrFilename <> '' then begin
    AStream.Seek(0,soFromBeginning);
    FZip.Add(AStream,FCurrFilename);
    FCurrFilename := '';
  end;
  AStream.Free;
end;

function TXLSZipArchiveIntfDelphi.Count: integer;
begin
  Result := FZip.FileCount;
end;

constructor TXLSZipArchiveIntfDelphi.Create;
begin
  FZip := TZipFile.Create;
end;

function TXLSZipArchiveIntfDelphi.CreateStream(Filename: AxUCString): TStream;
begin
  FCurrFilename := Filename;
  Result := TMemoryStream.Create;
end;

function TXLSZipArchiveIntfDelphi.CreateStream(FolderName, Filename: AxUCString): TStream;
begin
  if FolderName <> '' then
    FCurrFilename := FolderName + '/' + Filename
  else
    FCurrFilename := Filename;
  Result := TMemoryStream.Create;
end;

destructor TXLSZipArchiveIntfDelphi.Destroy;
begin
  FZip.Free;
  inherited;
end;

function TXLSZipArchiveIntfDelphi.Find(Filename: AxUCString): integer;
begin
  Result := FZip.IndexOf(Filename);
end;

procedure TXLSZipArchiveIntfDelphi.LoadFromStream(Stream: TStream);
begin
  FZip.Open(Stream,zmRead);
end;

function TXLSZipArchiveIntfDelphi.OpenStream(Filename: AxUCString): TStream;
var
  i: integer;
  Hdr: TZipHeader;
begin
  i := Find(Filename);
  if i < 0 then
    raise XLSRWException.CreateFmt('Can not find file %s in ZIP',[Filename]);

  FZip.Read(i,Result,Hdr);
end;

procedure TXLSZipArchiveIntfDelphi.SaveToStream(Stream: TStream);
begin
  FZip.Open(Stream,zmWrite);
  FWriteStream := Stream;
end;

procedure TXLSZipArchiveIntfDelphi.Write(Filename: AxUCString; Source: TStream);
begin
  Source.Seek(0,soFromBeginning);
  FZip.Add(Source,Filename);
end;

procedure TXLSZipArchiveIntfDelphi.Write(FolderName, Filename: AxUCString; Source: TStream);
begin
  Source.Seek(0,soFromBeginning);
  if FolderName <> '' then
    FZip.Add(Source,FolderName + '/' + Filename)
  else
    FZip.Add(Source,Filename);
end;

{$endif}

{ TXLSZipArchiveIntf }

destructor TXLSZipArchiveIntf.Destroy;
begin
  if FZipStream <> Nil then
    FZipStream.Free;

  inherited;
end;

procedure TXLSZipArchiveIntf.LoadFromFile(const AFilename: AxUCString);
begin
  if FZipStream <> Nil then
    FZipStream.Free;

  FZipStream := TFileStream.Create(AFilename,fmOpenRead + fmShareDenyNone);

  LoadFromStream(FZipStream);
end;

end.

