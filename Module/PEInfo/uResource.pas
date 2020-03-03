unit uResource;

interface

uses windows, Classes, ComCtrls, SysUtils, Forms, Dialogs;

function LoadPEResource(const PEFileName: string; TreeNode: TTreeNode; DispLevel: Boolean; DispAddress: Boolean): Boolean;

implementation

const
  IMAGE_RESOURCE_NAME_IS_STRING    = $80000000;
  IMAGE_OFFSET_STRIP_HIGH          = $7FFFFFFF;
  IMAGE_RESOURCE_DATA_IS_DIRECTORY = $80000000;

type
  PIMAGE_RESOURCE_DIRECTORY = ^TImageResourceDirectory;

  _IMAGE_RESOURCE_DIRECTORY = packed record
    Characteristics: DWORD;
    TimeDateStamp: DWORD;
    MajorVersion: WORD;
    MinorVersion: WORD;
    NumberOfNamedEntries: WORD;
    NumberOfIdEntries: WORD;
  end;

  TImageResourceDirectory = _IMAGE_RESOURCE_DIRECTORY;

  PIMAGE_RESOURCE_DIRECTORY_ENTRY = ^TImageResourceDirectoryEntry;

  _IMAGE_RESOURCE_DIRECTORY_ENTRY = packed record
    Name: DWORD;         { NameOffset:31,NameIsString:1 }
    OffsetToData: DWORD; { OffsetToDirectory:31,DataIsDirectory:1 }
  end;

  TImageResourceDirectoryEntry = _IMAGE_RESOURCE_DIRECTORY_ENTRY;

  PIMAGE_RESOURCE_DIR_STRING_U = ^TImageResourceDirStringU;

  _IMAGE_RESOURCE_DIR_STRING_U = packed record
    Length: WORD;
    NameString: WCHAR;
  end;

  TImageResourceDirStringU = _IMAGE_RESOURCE_DIR_STRING_U;

  PIMAGE_RESOURCE_DATA_ENTRY = ^TImageResourceDataEntry;

  _IMAGE_RESOURCE_DATA_ENTRY = packed record
    OffsetToData: DWORD;
    Size: DWORD;
    CodePage: DWORD;
    Reserved: DWORD;
  end;

  TImageResourceDataEntry = _IMAGE_RESOURCE_DATA_ENTRY;

function CheckResourceType(W: Byte): string;
begin
  case W of
    0:
      Result := 'Unknown0';
    1:
      Result := 'Cursor_entry';
    2:
      Result := 'λͼ';
    3:
      Result := 'Icon_entry';
    4:
      Result := '�˵�';
    5:
      Result := '�Ի���';
    6:
      Result := '�ַ���';
    7:
      Result := 'Font_dir';
    8:
      Result := '����';
    9:
      Result := '���ټ�';
    10:
      Result := 'RCDATA';
    11:
      Result := 'Message_table';
    13:
      Result := 'Unknow13';
    12:
      Result := '���';
    14:
      Result := 'ͼ��';
    15:
      Result := 'Unknow15';
    16:
      Result := '�汾';
  else
    Result := 'Unknow';
  end;
end;

function LoadPEResource(const PEFileName: string; TreeNode: TTreeNode; DispLevel: Boolean; DispAddress: Boolean): Boolean;
var
  iFile, iNum, jNum, I, J            : Integer;
  TmpPos                             : LongInt;  { �ļ���ʱλ�� }
  TmpStr                             : string;   { ��ʱ�ַ��� }
  PEResourceDirectoryPointerToRawData: Cardinal; { ��ԴĿ¼�����ַ }
  PEResourceDirectoryVirtualAddress  : Cardinal; { ��Դ���ʵ�ʵ�ַ }
  tn_ResDir1, tn_ResDir2, tn_ResDir3 : TTreeNode;
  PEResourceDirectory                : TImageResourceDirectory;
  PEResourceDirStringU               : TImageResourceDirStringU;
  ResWName                           : array [0 .. MAX_PATH] of WCHAR;
  ResSName                           : WideString; { PEResourceDirStringU.Name for String }
  iName                              : Integer;
  PEResDataEnt                       : TImageResourceDataEntry;
  ResDir_1                           : array of TImageResourceDirectory;
  ResDir_2                           : TImageResourceDirectory;
  ResdirEnt_1, ResdirEnt_2           : array of TImageResourceDirectoryEntry;
  ResdirEnt_3                        : TImageResourceDirectoryEntry;
  AAA                                : TImageDosHeader;
  BBB                                : TImageOptionalHeader;
  CCC                                : TImageFileHeader;
  DDD                                : array of TImageSectionHeader;
  P                                  : Integer;
  Temp                               : string;
begin
  Result := False;
  iFile  := FileOpen(PEFileName, fmOpenRead or fmShareDenyNone);
  if iFile <= 0 then
    Exit;

  FileSeek(iFile, 0, 0);
  FileRead(iFile, AAA, Sizeof(AAA));
  P := AAA._lfanew;

  FileSeek(iFile, P + 4, 0);
  FileRead(iFile, CCC, 20);
  J := CCC.NumberOfSections;
  if J = 0 then
    Exit;

  FileSeek(iFile, P + 24, 0);
  FileRead(iFile, BBB, 220);
  if (BBB.DataDirectory[2].VirtualAddress = 0) and (BBB.DataDirectory[2].Size = 0) then
    Exit;
  PEResourceDirectoryVirtualAddress := BBB.DataDirectory[2].VirtualAddress; // $55000

  PEResourceDirectoryPointerToRawData := 0;
  SetLength(DDD, J);
  for I := 0 to J - 1 do
  begin
    FileSeek(iFile, P + 220 + 28 + I * 40, 0);
    FileRead(iFile, DDD[I], Sizeof(DDD[I]));
    if DDD[I].VirtualAddress = PEResourceDirectoryVirtualAddress then
    begin
      PEResourceDirectoryPointerToRawData := DDD[I].PointerToRawData;
      Break;
    end;
  end;
  if PEResourceDirectoryPointerToRawData = 0 then
    Exit;

  FileSeek(iFile, PEResourceDirectoryPointerToRawData, 0);
  FileRead(iFile, PEResourceDirectory, Sizeof(PEResourceDirectory));
  iNum := PEResourceDirectory.NumberOfNamedEntries + PEResourceDirectory.NumberOfIdEntries;
  if iNum = 0 then
    Exit;

  SetLength(ResDir_1, iNum);
  SetLength(ResdirEnt_1, iNum);
  for I := 0 to iNum - 1 do
  begin
    FileRead(iFile, ResdirEnt_1[I], Sizeof(ResdirEnt_1[I]));
  end;

  // TmpStep := 90 div iNum; { �����ʱ���� }

  if DispLevel = True then
  begin
    for I := 0 to iNum - 1 do
    begin
      Application.ProcessMessages;
      FileRead(iFile, ResDir_1[I], Sizeof(ResDir_1[I]));
      if (ResdirEnt_1[I].Name and IMAGE_RESOURCE_NAME_IS_STRING) = 0 then
      begin { ���λΪ 0 ��ʣ��� 31 λΪһ������ID }
        TmpStr     := Format('%s', [CheckResourceType(ResdirEnt_1[I].Name)]);
        tn_ResDir1 := TTreeNodes(TreeNode.Owner).AddChild(TreeNode, TmpStr);
      end
      else { ���λΪ 1 ��ʣ�� 31 λΪ IMAGE_RESOURCE_DIR_STRING_U �ṹ��ƫ��λ�� �� Resource Section ��ʼ }
      begin
        TmpPos := FileSeek(iFile, 0, 1); { �����ļ�λ�� }
        FileSeek(iFile, ResdirEnt_1[I].Name and IMAGE_OFFSET_STRIP_HIGH + PEResourceDirectoryPointerToRawData, soFromBeginning);
        FileRead(iFile, PEResourceDirStringU, Sizeof(PEResourceDirStringU));
        FileRead(iFile, ResWName, (PEResourceDirStringU.Length - 1) * 2);
        ResSName   := PEResourceDirStringU.NameString;
        for iName  := 0 to PEResourceDirStringU.Length - 2 do
          ResSName := ResSName + ResWName[iName];
        TmpStr     := Format('%s', [ResSName]);
        tn_ResDir1 := TTreeNodes(TreeNode.Owner).AddChild(TreeNode, TmpStr);
        FileSeek(iFile, TmpPos, 0); { �ָ��ļ�λ�� i λ�� }
      end;
      jNum := ResDir_1[I].NumberOfNamedEntries + ResDir_1[I].NumberOfIdEntries;
      SetLength(ResdirEnt_2, jNum);
      for J := 0 to jNum - 1 do
      begin
        Application.ProcessMessages;
        FileRead(iFile, ResdirEnt_2[J], Sizeof(ResdirEnt_2[J]));
        TmpPos := FileSeek(iFile, 0, 1); { �����ַ }

        FileSeek(iFile, PEResourceDirectoryPointerToRawData + ResdirEnt_2[J].OffsetToData - IMAGE_RESOURCE_DATA_IS_DIRECTORY, 0);
        FileRead(iFile, ResDir_2, Sizeof(ResDir_2));
        FileRead(iFile, ResdirEnt_3, Sizeof(ResdirEnt_3));

        if (ResdirEnt_2[J].Name and IMAGE_RESOURCE_NAME_IS_STRING) = 0 then
        begin { ���λΪ 0 ��ʣ��� 31 λΪһ������ID }
          TmpStr     := Format('%s  ', [Format('%.2d', [ResdirEnt_2[J].Name])]);
          tn_ResDir2 := TTreeNodes(TreeNode.Owner).AddChild(tn_ResDir1, TmpStr);
        end
        else
        begin
          FileSeek(iFile, ResdirEnt_2[J].Name and IMAGE_OFFSET_STRIP_HIGH + PEResourceDirectoryPointerToRawData, soFromBeginning);
          FileRead(iFile, PEResourceDirStringU, Sizeof(PEResourceDirStringU));
          FileRead(iFile, ResWName, (PEResourceDirStringU.Length - 1) * 2);
          ResSName   := PEResourceDirStringU.NameString;
          for iName  := 0 to PEResourceDirStringU.Length - 2 do
            ResSName := ResSName + ResWName[iName];
          TmpStr     := Format('%s ByName[%s] ByID[%s]', [ResSName, Format('%.2d', [ResDir_2.NumberOfNamedEntries]), Format('%.2d', [ResDir_2.NumberOfIdEntries])]);
          tn_ResDir2 := TTreeNodes(TreeNode.Owner).AddChild(tn_ResDir1, TmpStr);
        end;

        { ���յ���Դ��ʾ }
        TmpStr     := Format('ID[%u] OffSetToData[$%s]', [ResdirEnt_3.Name, IntToHex(ResdirEnt_3.OffsetToData, 8)]);
        tn_ResDir3 := TTreeNodes(TreeNode.Owner).AddChild(tn_ResDir2, TmpStr);
        FileSeek(iFile, ResdirEnt_3.OffsetToData + PEResourceDirectoryPointerToRawData, 0);
        FileRead(iFile, PEResDataEnt, Sizeof(PEResDataEnt));
        TmpStr := Format('OffsetToData[$%s] Size[$%s] AddressInFile[$%s]', [IntToHex(PEResDataEnt.OffsetToData, 8), IntToHex(PEResDataEnt.Size, 8), IntToHex(PEResourceDirectoryPointerToRawData + PEResDataEnt.OffsetToData - PEResourceDirectoryVirtualAddress, 8)]);
        TTreeNodes(TreeNode.Owner).AddChild(tn_ResDir3, TmpStr); { ��������ļ��е�ʵ�ʵ�ַ }
        FileSeek(iFile, TmpPos, 0);                              { �ָ���ַ j ѭ�� }
      end;
    end;
  end;

  if DispLevel = False then
  begin
    for I := 0 to iNum - 1 do
    begin
      Application.ProcessMessages;
      FileRead(iFile, ResDir_1[I], Sizeof(ResDir_1[I]));
      if (ResdirEnt_1[I].Name and IMAGE_RESOURCE_NAME_IS_STRING) = 0 then
      begin { ���λΪ 0 ��ʣ��� 31 λΪһ������ID }
        TmpStr     := Format('%s', [CheckResourceType(ResdirEnt_1[I].Name)]);
        tn_ResDir1 := TTreeNodes(TreeNode.Owner).AddChild(TreeNode, TmpStr);
      end
      else { ���λΪ 1 ��ʣ�� 31 λΪ IMAGE_RESOURCE_DIR_STRING_U �ṹ��ƫ��λ�� �� Resource Section ��ʼ }
      begin
        TmpPos := FileSeek(iFile, 0, 1); { �����ļ�λ�� }
        FileSeek(iFile, ResdirEnt_1[I].Name and IMAGE_OFFSET_STRIP_HIGH + PEResourceDirectoryPointerToRawData, soFromBeginning);
        FileRead(iFile, PEResourceDirStringU, Sizeof(PEResourceDirStringU));
        FileRead(iFile, ResWName, (PEResourceDirStringU.Length - 1) * 2);
        ResSName   := PEResourceDirStringU.NameString;
        for iName  := 0 to PEResourceDirStringU.Length - 2 do
          ResSName := ResSName + ResWName[iName];
        TmpStr     := Format('%s', [ResSName]);
        tn_ResDir1 := TTreeNodes(TreeNode.Owner).AddChild(TreeNode, TmpStr);
        FileSeek(iFile, TmpPos, 0); { �ָ��ļ�λ�� i λ�� }
      end;
      jNum := ResDir_1[I].NumberOfNamedEntries + ResDir_1[I].NumberOfIdEntries;
      SetLength(ResdirEnt_2, jNum);
      for J := 0 to jNum - 1 do
      begin
        Application.ProcessMessages;
        FileRead(iFile, ResdirEnt_2[J], Sizeof(ResdirEnt_2[J]));
        TmpPos := FileSeek(iFile, 0, 1); { �����ַ }

        FileSeek(iFile, PEResourceDirectoryPointerToRawData + ResdirEnt_2[J].OffsetToData - IMAGE_RESOURCE_DATA_IS_DIRECTORY, 0);
        FileRead(iFile, ResDir_2, Sizeof(ResDir_2));
        FileRead(iFile, ResdirEnt_3, Sizeof(ResdirEnt_3));

        if (ResdirEnt_2[J].Name and IMAGE_RESOURCE_NAME_IS_STRING) = 0 then
        begin { ���λΪ 0 ��ʣ��� 31 λΪһ������ID }
          TmpStr := Format('%s  ', [Format('%.2d', [ResdirEnt_2[J].Name])]);
          Temp   := TmpStr;
        end
        else { ���λΪ 1 ��ʣ�� 31 λΪ IMAGE_RESOURCE_DIR_STRING_U �ṹ��ƫ��λ�� �� Resource Section ��ʼ }
        begin
          FileSeek(iFile, ResdirEnt_2[J].Name and IMAGE_OFFSET_STRIP_HIGH + PEResourceDirectoryPointerToRawData, soFromBeginning);
          FileRead(iFile, PEResourceDirStringU, Sizeof(PEResourceDirStringU));
          FileRead(iFile, ResWName, (PEResourceDirStringU.Length - 1) * 2);
          ResSName   := PEResourceDirStringU.NameString;
          for iName  := 0 to PEResourceDirStringU.Length - 2 do
            ResSName := ResSName + ResWName[iName];
          TmpStr     := Format('%s ', [ResSName]);
          Temp       := TmpStr;
        end;

        { ���յ���Դ��ʾ }
        if DispAddress then
          TmpStr := Format('OffsetToData[$%s] Size[$%s] AddressInFile[$%s]', [IntToHex(PEResDataEnt.OffsetToData, 8), IntToHex(PEResDataEnt.Size, 8), IntToHex(PEResourceDirectoryPointerToRawData + PEResDataEnt.OffsetToData - PEResourceDirectoryVirtualAddress, 8)])
        else
          TmpStr := '';
        FileSeek(iFile, ResdirEnt_3.OffsetToData + PEResourceDirectoryPointerToRawData, 0);
        FileRead(iFile, PEResDataEnt, Sizeof(PEResDataEnt));
        TmpStr := Temp + TmpStr;
        TTreeNodes(TreeNode.Owner).AddChild(tn_ResDir1, TmpStr);
        FileSeek(iFile, TmpPos, 0); { �ָ���ַ j ѭ�� }
      end;
    end;
  end;
  Result := True;
  FileClose(iFile);
end;

end.
