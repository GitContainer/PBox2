unit xpgParseDrawing;

// Copyright (c) 2010,2012 Axolot Data
// Web : http://www.axolot.com/xpg
// Mail: xpg@axolot.com
//
// X   X  PPP    GGG
//  X X   P  P  G
//   X    PPP   G  GG
//  X X   P     G   G
// X   X  P      GGG
//
// File generated with Axolot XPG, Xml Parser Generator.
// Version 0.00.90.
// File created on 2012-10-30 16:02:45

{$MINENUMSIZE 2}
{$BOOLEVAL OFF}
{$RANGECHECKS OFF}

interface

uses Classes, SysUtils, Contnrs, IniFiles, Math,
     xpgPUtils, xpgPLists, xpgPXMLUtils, xpgPXML, xpgParseDrawingCommon, xpgParseChart,
     Xc12Graphics,
     XLSReadWriteOPC5;

type TST_EditAs =  (steaTwoCell,steaOneCell,steaAbsolute);
const StrTST_EditAs: array[0..2] of AxUCString = ('twoCell','oneCell','absolute');

type TCT_GroupShape = class;

     TCT_GroupShapeXpgList = class;

     TCT_Marker_Pos = class(TXPGBase)
protected
     FCol: longword;
     FColOff: longword;
     FRow: longword;
     FRowOff: longword;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Marker_Pos);
     procedure CopyTo(AItem: TCT_Marker_Pos);

     property Col: longword read FCol write FCol;
     property ColOff: longword read FColOff write FColOff;
     property Row: longword read FRow write FRow;
     property RowOff: longword read FRowOff write FRowOff;
     end;

     TCT_GraphicalObjectFrameLocking = class(TXPGBase)
protected
     FNoGrp: boolean;
     FNoDrilldown: boolean;
     FNoSelect: boolean;
     FNoChangeAspect: boolean;
     FNoMove: boolean;
     FNoResize: boolean;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GraphicalObjectFrameLocking);
     procedure CopyTo(AItem: TCT_GraphicalObjectFrameLocking);
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property NoGrp: boolean read FNoGrp write FNoGrp;
     property NoDrilldown: boolean read FNoDrilldown write FNoDrilldown;
     property NoSelect: boolean read FNoSelect write FNoSelect;
     property NoChangeAspect: boolean read FNoChangeAspect write FNoChangeAspect;
     property NoMove: boolean read FNoMove write FNoMove;
     property NoResize: boolean read FNoResize write FNoResize;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_ConnectorLocking = class(TXPGBase)
protected
     FNoGrp: boolean;
     FNoSelect: boolean;
     FNoRot: boolean;
     FNoChangeAspect: boolean;
     FNoMove: boolean;
     FNoResize: boolean;
     FNoEditPoints: boolean;
     FNoAdjustHandles: boolean;
     FNoChangeArrowheads: boolean;
     FNoChangeShapeType: boolean;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ConnectorLocking);
     procedure CopyTo(AItem: TCT_ConnectorLocking);
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property NoGrp: boolean read FNoGrp write FNoGrp;
     property NoSelect: boolean read FNoSelect write FNoSelect;
     property NoRot: boolean read FNoRot write FNoRot;
     property NoChangeAspect: boolean read FNoChangeAspect write FNoChangeAspect;
     property NoMove: boolean read FNoMove write FNoMove;
     property NoResize: boolean read FNoResize write FNoResize;
     property NoEditPoints: boolean read FNoEditPoints write FNoEditPoints;
     property NoAdjustHandles: boolean read FNoAdjustHandles write FNoAdjustHandles;
     property NoChangeArrowheads: boolean read FNoChangeArrowheads write FNoChangeArrowheads;
     property NoChangeShapeType: boolean read FNoChangeShapeType write FNoChangeShapeType;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_PictureLocking = class(TXPGBase)
protected
     FNoGrp: boolean;
     FNoSelect: boolean;
     FNoRot: boolean;
     FNoChangeAspect: boolean;
     FNoMove: boolean;
     FNoResize: boolean;
     FNoEditPoints: boolean;
     FNoAdjustHandles: boolean;
     FNoChangeArrowheads: boolean;
     FNoChangeShapeType: boolean;
     FNoCrop: boolean;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PictureLocking);
     procedure CopyTo(AItem: TCT_PictureLocking);
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property NoGrp: boolean read FNoGrp write FNoGrp;
     property NoSelect: boolean read FNoSelect write FNoSelect;
     property NoRot: boolean read FNoRot write FNoRot;
     property NoChangeAspect: boolean read FNoChangeAspect write FNoChangeAspect;
     property NoMove: boolean read FNoMove write FNoMove;
     property NoResize: boolean read FNoResize write FNoResize;
     property NoEditPoints: boolean read FNoEditPoints write FNoEditPoints;
     property NoAdjustHandles: boolean read FNoAdjustHandles write FNoAdjustHandles;
     property NoChangeArrowheads: boolean read FNoChangeArrowheads write FNoChangeArrowheads;
     property NoChangeShapeType: boolean read FNoChangeShapeType write FNoChangeShapeType;
     property NoCrop: boolean read FNoCrop write FNoCrop;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_GroupLocking = class(TXPGBase)
protected
     FNoGrp: boolean;
     FNoUngrp: boolean;
     FNoSelect: boolean;
     FNoRot: boolean;
     FNoChangeAspect: boolean;
     FNoMove: boolean;
     FNoResize: boolean;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GroupLocking);
     procedure CopyTo(AItem: TCT_GroupLocking);
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property NoGrp: boolean read FNoGrp write FNoGrp;
     property NoUngrp: boolean read FNoUngrp write FNoUngrp;
     property NoSelect: boolean read FNoSelect write FNoSelect;
     property NoRot: boolean read FNoRot write FNoRot;
     property NoChangeAspect: boolean read FNoChangeAspect write FNoChangeAspect;
     property NoMove: boolean read FNoMove write FNoMove;
     property NoResize: boolean read FNoResize write FNoResize;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_NonVisualGraphicFrameProperties = class(TXPGBase)
protected
     FA_GraphicFrameLocks: TCT_GraphicalObjectFrameLocking;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_NonVisualGraphicFrameProperties);
     procedure CopyTo(AItem: TCT_NonVisualGraphicFrameProperties);
     function  Create_A_GraphicFrameLocks: TCT_GraphicalObjectFrameLocking;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property A_GraphicFrameLocks: TCT_GraphicalObjectFrameLocking read FA_GraphicFrameLocks;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_GraphicalObjectData = class(TXPGBase)
protected
     FUri: AxUCString;
     FChart: TXPGDocXLSXChart;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     function  Create_Chart: TXPGDocXLSXChart;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Assign(AItem: TCT_GraphicalObjectData);
     procedure CopyTo(AItem: TCT_GraphicalObjectData);

     property Uri: AxUCString read FUri write FUri;
     property Chart: TXPGDocXLSXChart read FChart;
     end;

     TCT_NonVisualConnectorProperties = class(TXPGBase)
protected
     FA_CxnSpLocks: TCT_ConnectorLocking;
     FA_StCxn: TCT_Connection;
     FA_EndCxn: TCT_Connection;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_NonVisualConnectorProperties);
     procedure CopyTo(AItem: TCT_NonVisualConnectorProperties);
     function  Create_A_CxnSpLocks: TCT_ConnectorLocking;
     function  Create_A_StCxn: TCT_Connection;
     function  Create_A_EndCxn: TCT_Connection;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property A_CxnSpLocks: TCT_ConnectorLocking read FA_CxnSpLocks;
     property A_StCxn: TCT_Connection read FA_StCxn;
     property A_EndCxn: TCT_Connection read FA_EndCxn;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_NonVisualPictureProperties = class(TXPGBase)
protected
     FPreferRelativeResize: boolean;
     FA_PicLocks: TCT_PictureLocking;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_NonVisualPictureProperties);
     procedure CopyTo(AItem: TCT_NonVisualPictureProperties);
     function  Create_A_PicLocks: TCT_PictureLocking;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property PreferRelativeResize: boolean read FPreferRelativeResize write FPreferRelativeResize;
     property A_PicLocks: TCT_PictureLocking read FA_PicLocks;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_NonVisualGroupDrawingShapeProps = class(TXPGBase)
protected
     FA_GrpSpLocks: TCT_GroupLocking;
     FA_ExtLst: TCT_OfficeArtExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_NonVisualGroupDrawingShapeProps);
     procedure CopyTo(AItem: TCT_NonVisualGroupDrawingShapeProps);
     function  Create_A_GrpSpLocks: TCT_GroupLocking;
     function  Create_A_ExtLst: TCT_OfficeArtExtensionList;

     property A_GrpSpLocks: TCT_GroupLocking read FA_GrpSpLocks;
     property A_ExtLst: TCT_OfficeArtExtensionList read FA_ExtLst;
     end;

     TCT_GraphicalObjectFrameNonVisual = class(TXPGBase)
protected
     FCNvPr: TCT_NonVisualDrawingProps;
     FCNvGraphicFramePr: TCT_NonVisualGraphicFrameProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GraphicalObjectFrameNonVisual);
     procedure CopyTo(AItem: TCT_GraphicalObjectFrameNonVisual);
     function  Create_CNvPr: TCT_NonVisualDrawingProps;
     function  Create_CNvGraphicFramePr: TCT_NonVisualGraphicFrameProperties;

     property CNvPr: TCT_NonVisualDrawingProps read FCNvPr;
     property CNvGraphicFramePr: TCT_NonVisualGraphicFrameProperties read FCNvGraphicFramePr;
     end;

     TCT_GraphicalObject = class(TXPGBase)
protected
     FA_GraphicData: TCT_GraphicalObjectData;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GraphicalObject);
     procedure CopyTo(AItem: TCT_GraphicalObject);
     function  Create_A_GraphicData: TCT_GraphicalObjectData;

     property GraphicData: TCT_GraphicalObjectData read FA_GraphicData;
     end;

     TCT_ConnectorNonVisual = class(TXPGBase)
protected
     FCNvPr: TCT_NonVisualDrawingProps;
     FCNvCxnSpPr: TCT_NonVisualConnectorProperties;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ConnectorNonVisual);
     procedure CopyTo(AItem: TCT_ConnectorNonVisual);
     function  Create_CNvPr: TCT_NonVisualDrawingProps;
     function  Create_CNvCxnSpPr: TCT_NonVisualConnectorProperties;

     property CNvPr: TCT_NonVisualDrawingProps read FCNvPr;
     property CNvCxnSpPr: TCT_NonVisualConnectorProperties read FCNvCxnSpPr;
     end;

     TCT_PictureNonVisual = class(TXPGBase)
protected
     FCNvPr: TCT_NonVisualDrawingProps;
     FCNvPicPr: TCT_NonVisualPictureProperties;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PictureNonVisual);
     procedure CopyTo(AItem: TCT_PictureNonVisual);
     function  Create_CNvPr: TCT_NonVisualDrawingProps;
     function  Create_CNvPicPr: TCT_NonVisualPictureProperties;

     property CNvPr: TCT_NonVisualDrawingProps read FCNvPr;
     property CNvPicPr: TCT_NonVisualPictureProperties read FCNvPicPr;
     end;

     TCT_GroupShapeNonVisual = class(TXPGBase)
protected
     FCNvPr: TCT_NonVisualDrawingProps;
     FCNvGrpSpPr: TCT_NonVisualGroupDrawingShapeProps;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GroupShapeNonVisual);
     procedure CopyTo(AItem: TCT_GroupShapeNonVisual);
     function  Create_CNvPr: TCT_NonVisualDrawingProps;
     function  Create_CNvGrpSpPr: TCT_NonVisualGroupDrawingShapeProps;

     property CNvPr: TCT_NonVisualDrawingProps read FCNvPr;
     property CNvGrpSpPr: TCT_NonVisualGroupDrawingShapeProps read FCNvGrpSpPr;
     end;

     TCT_ShapeXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Shape;
public
     function  Add: TCT_Shape;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ShapeXpgList);
     procedure CopyTo(AItem: TCT_ShapeXpgList);
     property Items[Index: integer]: TCT_Shape read GetItems; default;
     end;

     TCT_GraphicalObjectFrame = class(TXPGBase)
protected
     FMacro: AxUCString;
     FFPublished: boolean;
     FNvGraphicFramePr: TCT_GraphicalObjectFrameNonVisual;
     FXfrm: TCT_Transform2D;
     FA_Graphic: TCT_GraphicalObject;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;

     procedure Clear;

     procedure CreateDefault(AId: integer; AName: AxUCString);

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Assign(AItem: TCT_GraphicalObjectFrame);
     procedure CopyTo(AItem: TCT_GraphicalObjectFrame);
     function  Create_NvGraphicFramePr: TCT_GraphicalObjectFrameNonVisual;
     function  Create_Xfrm: TCT_Transform2D;
     function  Create_A_Graphic: TCT_GraphicalObject;

     function  IsChart: boolean;

     property Macro: AxUCString read FMacro write FMacro;
     property FPublished: boolean read FFPublished write FFPublished;
     property NvGraphicFramePr: TCT_GraphicalObjectFrameNonVisual read FNvGraphicFramePr;
     property Xfrm: TCT_Transform2D read FXfrm;
     property Graphic: TCT_GraphicalObject read FA_Graphic;
     end;

     TCT_GraphicalObjectFrameXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_GraphicalObjectFrame;
public
     function  Add: TCT_GraphicalObjectFrame;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_GraphicalObjectFrameXpgList);
     procedure CopyTo(AItem: TCT_GraphicalObjectFrameXpgList);
     property Items[Index: integer]: TCT_GraphicalObjectFrame read GetItems; default;
     end;

     TCT_Connector = class(TXPGBase)
protected
     FMacro: AxUCString;
     FFPublished: boolean;
     FNvCxnSpPr: TCT_ConnectorNonVisual;
     FSpPr: TCT_ShapeProperties;
     FStyle: TCT_ShapeStyle;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Connector);
     procedure CopyTo(AItem: TCT_Connector);
     function  Create_NvCxnSpPr: TCT_ConnectorNonVisual;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_Style: TCT_ShapeStyle;

     property Macro: AxUCString read FMacro write FMacro;
     property FPublished: boolean read FFPublished write FFPublished;
     property NvCxnSpPr: TCT_ConnectorNonVisual read FNvCxnSpPr;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property Style: TCT_ShapeStyle read FStyle;
     end;

     TCT_ConnectorXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Connector;
public
     function  Add: TCT_Connector;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ConnectorXpgList);
     procedure CopyTo(AItem: TCT_ConnectorXpgList);
     property Items[Index: integer]: TCT_Connector read GetItems; default;
     end;

     TCT_Picture = class(TXPGBase)
protected
     FMacro: AxUCString;
     FFPublished: boolean;
     FNvPicPr: TCT_PictureNonVisual;
     FBlipFill: TCT_BlipFillProperties;
     FSpPr: TCT_ShapeProperties;
     FStyle: TCT_ShapeStyle;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Picture);
     procedure CopyTo(AItem: TCT_Picture);
     function  Create_NvPicPr: TCT_PictureNonVisual;
     function  Create_BlipFill: TCT_BlipFillProperties;
     function  Create_SpPr: TCT_ShapeProperties;
     function  Create_Style: TCT_ShapeStyle;

     function  GetImage: TXc12GraphicImage;

     property Macro: AxUCString read FMacro write FMacro;
     property FPublished: boolean read FFPublished write FFPublished;
     property NvPicPr: TCT_PictureNonVisual read FNvPicPr;
     property BlipFill: TCT_BlipFillProperties read FBlipFill;
     property SpPr: TCT_ShapeProperties read FSpPr;
     property Style: TCT_ShapeStyle read FStyle;
     end;

     TCT_PictureXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Picture;
public
     function  Add: TCT_Picture;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PictureXpgList);
     procedure CopyTo(AItem: TCT_PictureXpgList);
     property Items[Index: integer]: TCT_Picture read GetItems; default;
     end;

     TCT_GroupShape = class(TXPGBase)
protected
     FNvGrpSpPr: TCT_GroupShapeNonVisual;
     FGrpSpPr: TCT_GroupShapeProperties;
     FSpXpgList: TCT_ShapeXpgList;
     FGrpSpXpgList: TCT_GroupShapeXpgList;
     FGraphicFrameXpgList: TCT_GraphicalObjectFrameXpgList;
     FCxnSpXpgList: TCT_ConnectorXpgList;
     FPicXpgList: TCT_PictureXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GroupShape);
     procedure CopyTo(AItem: TCT_GroupShape);
     function  Create_NvGrpSpPr: TCT_GroupShapeNonVisual;
     function  Create_GrpSpPr: TCT_GroupShapeProperties;
     function  Create_SpXpgList: TCT_ShapeXpgList;
     function  Create_GrpSpXpgList: TCT_GroupShapeXpgList;
     function  Create_GraphicFrameXpgList: TCT_GraphicalObjectFrameXpgList;
     function  Create_CxnSpXpgList: TCT_ConnectorXpgList;
     function  Create_PicXpgList: TCT_PictureXpgList;

     property NvGrpSpPr: TCT_GroupShapeNonVisual read FNvGrpSpPr;
     property GrpSpPr: TCT_GroupShapeProperties read FGrpSpPr;
     property SpXpgList: TCT_ShapeXpgList read FSpXpgList;
     property GrpSpXpgList: TCT_GroupShapeXpgList read FGrpSpXpgList write FGrpSpXpgList;
     property GraphicFrameXpgList: TCT_GraphicalObjectFrameXpgList read FGraphicFrameXpgList;
     property CxnSpXpgList: TCT_ConnectorXpgList read FCxnSpXpgList;
     property PicXpgList: TCT_PictureXpgList read FPicXpgList;
     end;

     TCT_GroupShapeXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_GroupShape;
public
     function  Add: TCT_GroupShape;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_GroupShapeXpgList);
     procedure CopyTo(AItem: TCT_GroupShapeXpgList);
     property Items[Index: integer]: TCT_GroupShape read GetItems; default;
     end;

     TEG_ObjectChoices = class(TXPGBase)
protected
     FSp: TCT_Shape;
     FGrpSp: TCT_GroupShape;
     FGraphicFrame: TCT_GraphicalObjectFrame;
     FCxnSp: TCT_Connector;
     FPic: TCT_Picture;

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure Clear;
     procedure Assign(AItem: TEG_ObjectChoices);
     procedure CopyTo(AItem: TEG_ObjectChoices);
     function  Create_Sp: TCT_Shape;
     function  Create_GrpSp: TCT_GroupShape;
     function  Create_GraphicFrame: TCT_GraphicalObjectFrame;
     function  Create_CxnSp: TCT_Connector;
     function  Create_Pic: TCT_Picture;

     function  IsPicture: boolean;
     function  IsChart: boolean;

     property Sp: TCT_Shape read FSp;
     property GrpSp: TCT_GroupShape read FGrpSp;
     property GraphicFrame: TCT_GraphicalObjectFrame read FGraphicFrame;
     property CxnSp: TCT_Connector read FCxnSp;
     property Pic: TCT_Picture read FPic;
     end;

     TCT_AnchorClientData = class(TXPGBase)
protected
     FFLocksWithSheet: boolean;
     FFPrintsWithSheet: boolean;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AnchorClientData);
     procedure CopyTo(AItem: TCT_AnchorClientData);

     property FLocksWithSheet: boolean read FFLocksWithSheet write FFLocksWithSheet;
     property FPrintsWithSheet: boolean read FFPrintsWithSheet write FFPrintsWithSheet;
     end;

     TCT_TwoCellAnchor = class(TXPGBase)
protected
     FEditAs: TST_EditAs;
     FFrom: TCT_Marker_Pos;
     FTo: TCT_Marker_Pos;
     FEG_ObjectChoices: TEG_ObjectChoices;
     FClientData: TCT_AnchorClientData;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;

     function  MakeChart: TXPGDocXLSXChart;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Clear;
     procedure Assign(AItem: TCT_TwoCellAnchor);
     procedure CopyTo(AItem: TCT_TwoCellAnchor);

     property EditAs: TST_EditAs read FEditAs write FEditAs;
     property From: TCT_Marker_Pos read FFrom;
     property To_: TCT_Marker_Pos read FTo;
     property Objects: TEG_ObjectChoices read FEG_ObjectChoices;
     property ClientData: TCT_AnchorClientData read FClientData;
     end;

     TCT_TwoCellAnchorXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_TwoCellAnchor;
public
     function  Add: TCT_TwoCellAnchor; overload;
     function  Add(ACol1,ARow1,ACol2,ARow2: integer): TCT_TwoCellAnchor; overload;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_TwoCellAnchorXpgList);
     procedure CopyTo(AItem: TCT_TwoCellAnchorXpgList);
     property Items[Index: integer]: TCT_TwoCellAnchor read GetItems; default;
     end;

     TCT_XYMarker = class(TXPGBase)
protected
     FX: double;
     FY: double;
public
     constructor Create(AOwner: TXPGDocBase);

     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     property X: double read FX write FX;
     property Y: double read FY write FY;
     end;

     TCT_RelSizeAnchor = class(TXPGBase)
protected
     FFrom: TCT_XYMarker;
     FTo  : TCT_XYMarker;
     FSp  : TCT_Shape;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;

     function  CheckAssigned: integer; override;

     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure AddTextBox(AText: AxUCString); overload;
     function  AddTextBox: TCT_TextParagraph; overload;

     property From: TCT_XYMarker read FFrom;
     property To_ : TCT_XYMarker read FTo;
     property Sp  : TCT_Shape read FSp;
     end;

     TCT_RelSizeAnchorXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_RelSizeAnchor;
public
     function  Add: TCT_RelSizeAnchor; overload;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_RelSizeAnchor read GetItems; default;
     end;

     TCT_OneCellAnchor = class(TXPGBase)
protected
     FFrom: TCT_Marker_Pos;
     FExt: TCT_PositiveSize2D;
     FEG_ObjectChoices: TEG_ObjectChoices;
     FClientData: TCT_AnchorClientData;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_OneCellAnchor);
     procedure CopyTo(AItem: TCT_OneCellAnchor);
     function  Create_From: TCT_Marker_Pos;
     function  Create_Ext: TCT_PositiveSize2D;
     function  Create_ClientData: TCT_AnchorClientData;

     property From: TCT_Marker_Pos read FFrom;
     property Ext: TCT_PositiveSize2D read FExt;
     property Objects: TEG_ObjectChoices read FEG_ObjectChoices;
     property ClientData: TCT_AnchorClientData read FClientData;
     end;

     TCT_OneCellAnchorXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_OneCellAnchor;
public
     function  Add: TCT_OneCellAnchor;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_OneCellAnchorXpgList);
     procedure CopyTo(AItem: TCT_OneCellAnchorXpgList);
     property Items[Index: integer]: TCT_OneCellAnchor read GetItems; default;
     end;

     TCT_AbsoluteAnchor = class(TXPGBase)
protected
     FPos: TCT_Point2D;
     FExt: TCT_PositiveSize2D;
     FEG_ObjectChoices: TEG_ObjectChoices;
     FClientData: TCT_AnchorClientData;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AbsoluteAnchor);
     procedure CopyTo(AItem: TCT_AbsoluteAnchor);
     function  Create_Pos: TCT_Point2D;
     function  Create_Ext: TCT_PositiveSize2D;
     function  Create_ClientData: TCT_AnchorClientData;

     property Pos: TCT_Point2D read FPos;
     property Ext: TCT_PositiveSize2D read FExt;
     property Objects: TEG_ObjectChoices read FEG_ObjectChoices;
     property ClientData: TCT_AnchorClientData read FClientData;
     end;

     TCT_AbsoluteAnchorXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_AbsoluteAnchor;
public
     function  Add: TCT_AbsoluteAnchor;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_AbsoluteAnchorXpgList);
     procedure CopyTo(AItem: TCT_AbsoluteAnchorXpgList);
     property Items[Index: integer]: TCT_AbsoluteAnchor read GetItems; default;
     end;

     TEG_Anchor = class(TXPGBase)
protected
     FTwoCellAnchorXpgList: TCT_TwoCellAnchorXpgList;
     FOneCellAnchorXpgList: TCT_OneCellAnchorXpgList;
     FAbsoluteAnchorXpgList: TCT_AbsoluteAnchorXpgList;
     FRelSizeAnchorXpgList: TCT_RelSizeAnchorXpgList;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     function  HasData: boolean;

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_Anchor);
     procedure CopyTo(AItem: TEG_Anchor);

     property TwoCellAnchors: TCT_TwoCellAnchorXpgList read FTwoCellAnchorXpgList;
     property OneCellAnchors: TCT_OneCellAnchorXpgList read FOneCellAnchorXpgList;
     property AbsoluteAnchors: TCT_AbsoluteAnchorXpgList read FAbsoluteAnchorXpgList;
     property RelSizeAnchors: TCT_RelSizeAnchorXpgList read FRelSizeAnchorXpgList;
     end;

     TCT_Drawing = class(TXPGBase)
protected
     FEG_Anchor: TEG_Anchor;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     function  HasData: boolean;

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Drawing);
     procedure CopyTo(AItem: TCT_Drawing);

     property Anchors: TEG_Anchor read FEG_Anchor;
     end;

     T__ROOT__ = class(TXPGBase)
protected
     FRootAttributes: TStringXpgList;
     FCurrWriteClass: TClass;
     FWsDr: TCT_Drawing;
public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);
     procedure WriteUserShapes(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: T__ROOT__);
     procedure CopyTo(AItem: T__ROOT__);

     property RootAttributes: TStringXpgList read FRootAttributes;
     property WsDr: TCT_Drawing read FWsDr;
     end;

     TXPGDocXLSXDrawing = class(TXPGDocBase)
private
     function GetErrors: TXpgPErrors;
protected
     FRoot: T__ROOT__;
     FReader: TXPGReader;
     FWriter: TXpgWriteXML;

     function  GetWsDr: TCT_Drawing;
public
     constructor Create(AGrManager: TXc12GraphicManager);
     destructor Destroy; override;

     procedure LoadFromFile(AFilename: AxUCString);
     procedure _LoadFromStream(AStream: TStream; AUserShapes: boolean = False);
     procedure SaveToFile(AFilename: AxUCString; AClassToWrite: TClass);

     procedure SaveToStream(AStream: TStream; AUserShapes: boolean = False);

     property Root: T__ROOT__ read FRoot;
     property WsDr: TCT_Drawing read GetWsDr;
     property Errors: TXpgPErrors read GetErrors;
     end;

implementation

var
  L_DrwCounter: integer;

procedure ReadUnionTST_AdjAngle(AValue: AxUCString; APtr: PST_AdjAngle);
begin
  if XmlTryStrToInt(AValue,APtr.Val1) then
  begin
    APtr.nVal := 0;
    Exit;
  end;
  if AValue <> '' then
  begin
    APtr.Val2 := AValue;
    APtr.nVal := 1;
    Exit;
  end;
end;

function  WriteUnionTST_AdjAngle(APtr: PST_AdjAngle): AxUCString;
begin
  case APtr.nVal of
    0: Result := XmlIntToStr(APtr.Val1);
    1: Result := APtr.Val2;
  end;
end;

procedure ReadUnionTST_AdjCoordinate(AValue: AxUCString; APtr: PST_AdjCoordinate);
begin
  if XmlTryStrToInt(AValue,APtr.Val1) then
  begin
    APtr.nVal := 0;
    Exit;
  end;
  if AValue <> '' then
  begin
    APtr.Val2 := AValue;
    APtr.nVal := 1;
    Exit;
  end;
end;

function  WriteUnionTST_AdjCoordinate(APtr: PST_AdjCoordinate): AxUCString;
begin
  case APtr.nVal of
    0: Result := XmlIntToStr(APtr.Val1);
    1: Result := APtr.Val2;
  end;
end;

{ TCT_Marker }

function  TCT_Marker_Pos.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FCol <> 2147483632 then
    Inc(ElemsAssigned);
  if FColOff <> 2147483632 then
    Inc(ElemsAssigned);
  if FRow <> 2147483632 then
    Inc(ElemsAssigned);
  if FRowOff <> 2147483632 then
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Marker_Pos.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002C6: FCol := XmlStrToIntDef(AReader.Text,0);
    $000003E1: FColOff := XmlStrToIntDef(AReader.Text,0);
    $000002E0: FRow := XmlStrToIntDef(AReader.Text,0);
    $000003FB: FRowOff := XmlStrToIntDef(AReader.Text,0);
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_Marker_Pos.Write(AWriter: TXpgWriteXML);
begin
  if FCol <> 2147483632 then begin
    AWriter.SimpleTextTag('xdr:col',XmlIntToStr(FCol));
    if FColOff <> 2147483632 then
      AWriter.SimpleTextTag('xdr:colOff',XmlIntToStr(FColOff))
    else
      AWriter.SimpleTextTag('xdr:colOff',XmlIntToStr(0));
  end;
  if FRow <> 2147483632 then begin
    AWriter.SimpleTextTag('xdr:row',XmlIntToStr(FRow));
    if FRowOff <> 2147483632 then
      AWriter.SimpleTextTag('xdr:rowOff',XmlIntToStr(FRowOff))
    else
      AWriter.SimpleTextTag('xdr:rowOff',XmlIntToStr(0));
  end;
end;

constructor TCT_Marker_Pos.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FCol := 2147483632;
  FColOff := 2147483632;
  FRow := 2147483632;
  FRowOff := 2147483632;
end;

destructor TCT_Marker_Pos.Destroy;
begin
end;

procedure TCT_Marker_Pos.Clear;
begin
  FAssigneds := [];
  FCol := 2147483632;
  FColOff := 2147483632;
  FRow := 2147483632;
  FRowOff := 2147483632;
end;

procedure TCT_Marker_Pos.Assign(AItem: TCT_Marker_Pos);
begin
end;

procedure TCT_Marker_Pos.CopyTo(AItem: TCT_Marker_Pos);
begin
end;

{ TCT_GraphicalObjectFrameLocking }

function  TCT_GraphicalObjectFrameLocking.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FNoGrp <> False then 
    Inc(AttrsAssigned);
  if FNoDrilldown <> False then 
    Inc(AttrsAssigned);
  if FNoSelect <> False then 
    Inc(AttrsAssigned);
  if FNoChangeAspect <> False then 
    Inc(AttrsAssigned);
  if FNoMove <> False then 
    Inc(AttrsAssigned);
  if FNoResize <> False then 
    Inc(AttrsAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_GraphicalObjectFrameLocking.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:extLst' then 
  begin
    if FA_ExtLst = Nil then 
      FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
    Result := FA_ExtLst;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GraphicalObjectFrameLocking.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.Assigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_GraphicalObjectFrameLocking.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FNoGrp <> False then 
    AWriter.AddAttribute('noGrp',XmlBoolToStr(FNoGrp));
  if FNoDrilldown <> False then 
    AWriter.AddAttribute('noDrilldown',XmlBoolToStr(FNoDrilldown));
  if FNoSelect <> False then
    AWriter.AddAttribute('noSelect',XmlBoolToStr(FNoSelect));
  if FNoChangeAspect <> False then 
    AWriter.AddAttribute('noChangeAspect',XmlBoolToStr(FNoChangeAspect));
  if FNoMove <> False then 
    AWriter.AddAttribute('noMove',XmlBoolToStr(FNoMove));
  if FNoResize <> False then 
    AWriter.AddAttribute('noResize',XmlBoolToStr(FNoResize));
end;

procedure TCT_GraphicalObjectFrameLocking.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000206: FNoGrp := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000048C: FNoDrilldown := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000033D: FNoSelect := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000583: FNoChangeAspect := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000274: FNoMove := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000034F: FNoResize := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_GraphicalObjectFrameLocking.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 6;
  FNoGrp := False;
  FNoDrilldown := False;
  FNoSelect := False;
  FNoChangeAspect := False;
  FNoMove := False;
  FNoResize := False;
end;

destructor TCT_GraphicalObjectFrameLocking.Destroy;
begin
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_GraphicalObjectFrameLocking.Clear;
begin
  FAssigneds := [];
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
  FNoGrp := False;
  FNoDrilldown := False;
  FNoSelect := False;
  FNoChangeAspect := False;
  FNoMove := False;
  FNoResize := False;
end;

procedure TCT_GraphicalObjectFrameLocking.Assign(AItem: TCT_GraphicalObjectFrameLocking);
begin
end;

procedure TCT_GraphicalObjectFrameLocking.CopyTo(AItem: TCT_GraphicalObjectFrameLocking);
begin
end;

function  TCT_GraphicalObjectFrameLocking.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_ConnectorLocking }

function  TCT_ConnectorLocking.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FNoGrp <> False then 
    Inc(AttrsAssigned);
  if FNoSelect <> False then 
    Inc(AttrsAssigned);
  if FNoRot <> False then 
    Inc(AttrsAssigned);
  if FNoChangeAspect <> False then 
    Inc(AttrsAssigned);
  if FNoMove <> False then 
    Inc(AttrsAssigned);
  if FNoResize <> False then 
    Inc(AttrsAssigned);
  if FNoEditPoints <> False then 
    Inc(AttrsAssigned);
  if FNoAdjustHandles <> False then 
    Inc(AttrsAssigned);
  if FNoChangeArrowheads <> False then 
    Inc(AttrsAssigned);
  if FNoChangeShapeType <> False then 
    Inc(AttrsAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ConnectorLocking.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:extLst' then 
  begin
    if FA_ExtLst = Nil then
      FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
    Result := FA_ExtLst;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ConnectorLocking.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.Assigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:extLst');
end;

procedure TCT_ConnectorLocking.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FNoGrp <> False then 
    AWriter.AddAttribute('noGrp',XmlBoolToStr(FNoGrp));
  if FNoSelect <> False then 
    AWriter.AddAttribute('noSelect',XmlBoolToStr(FNoSelect));
  if FNoRot <> False then 
    AWriter.AddAttribute('noRot',XmlBoolToStr(FNoRot));
  if FNoChangeAspect <> False then 
    AWriter.AddAttribute('noChangeAspect',XmlBoolToStr(FNoChangeAspect));
  if FNoMove <> False then 
    AWriter.AddAttribute('noMove',XmlBoolToStr(FNoMove));
  if FNoResize <> False then 
    AWriter.AddAttribute('noResize',XmlBoolToStr(FNoResize));
  if FNoEditPoints <> False then 
    AWriter.AddAttribute('noEditPoints',XmlBoolToStr(FNoEditPoints));
  if FNoAdjustHandles <> False then 
    AWriter.AddAttribute('noAdjustHandles',XmlBoolToStr(FNoAdjustHandles));
  if FNoChangeArrowheads <> False then 
    AWriter.AddAttribute('noChangeArrowheads',XmlBoolToStr(FNoChangeArrowheads));
  if FNoChangeShapeType <> False then 
    AWriter.AddAttribute('noChangeShapeType',XmlBoolToStr(FNoChangeShapeType));
end;

procedure TCT_ConnectorLocking.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000206: FNoGrp := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000033D: FNoSelect := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000212: FNoRot := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000583: FNoChangeAspect := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000274: FNoMove := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000034F: FNoResize := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004E0: FNoEditPoints := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000607: FNoAdjustHandles := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000733: FNoChangeArrowheads := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000006B6: FNoChangeShapeType := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_ConnectorLocking.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 10;
  FNoGrp := False;
  FNoSelect := False;
  FNoRot := False;
  FNoChangeAspect := False;
  FNoMove := False;
  FNoResize := False;
  FNoEditPoints := False;
  FNoAdjustHandles := False;
  FNoChangeArrowheads := False;
  FNoChangeShapeType := False;
end;

destructor TCT_ConnectorLocking.Destroy;
begin
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_ConnectorLocking.Clear;
begin
  FAssigneds := [];
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
  FNoGrp := False;
  FNoSelect := False;
  FNoRot := False;
  FNoChangeAspect := False;
  FNoMove := False;
  FNoResize := False;
  FNoEditPoints := False;
  FNoAdjustHandles := False;
  FNoChangeArrowheads := False;
  FNoChangeShapeType := False;
end;

procedure TCT_ConnectorLocking.Assign(AItem: TCT_ConnectorLocking);
begin
end;

procedure TCT_ConnectorLocking.CopyTo(AItem: TCT_ConnectorLocking);
begin
end;

function  TCT_ConnectorLocking.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_PictureLocking }

function  TCT_PictureLocking.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FNoGrp <> False then 
    Inc(AttrsAssigned);
  if FNoSelect <> False then 
    Inc(AttrsAssigned);
  if FNoRot <> False then 
    Inc(AttrsAssigned);
  if FNoChangeAspect <> False then 
    Inc(AttrsAssigned);
  if FNoMove <> False then 
    Inc(AttrsAssigned);
  if FNoResize <> False then 
    Inc(AttrsAssigned);
  if FNoEditPoints <> False then 
    Inc(AttrsAssigned);
  if FNoAdjustHandles <> False then 
    Inc(AttrsAssigned);
  if FNoChangeArrowheads <> False then 
    Inc(AttrsAssigned);
  if FNoChangeShapeType <> False then 
    Inc(AttrsAssigned);
  if FNoCrop <> False then 
    Inc(AttrsAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PictureLocking.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:extLst' then
  begin
    if FA_ExtLst = Nil then
      FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
    Result := FA_ExtLst;
  end
  else
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_PictureLocking.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then
    if xaElements in FA_ExtLst.Assigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_PictureLocking.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FNoGrp <> False then 
    AWriter.AddAttribute('noGrp',XmlBoolToStr(FNoGrp));
  if FNoSelect <> False then 
    AWriter.AddAttribute('noSelect',XmlBoolToStr(FNoSelect));
  if FNoRot <> False then 
    AWriter.AddAttribute('noRot',XmlBoolToStr(FNoRot));
  if FNoChangeAspect <> False then 
    AWriter.AddAttribute('noChangeAspect',XmlBoolToStr(FNoChangeAspect));
  if FNoMove <> False then 
    AWriter.AddAttribute('noMove',XmlBoolToStr(FNoMove));
  if FNoResize <> False then 
    AWriter.AddAttribute('noResize',XmlBoolToStr(FNoResize));
  if FNoEditPoints <> False then 
    AWriter.AddAttribute('noEditPoints',XmlBoolToStr(FNoEditPoints));
  if FNoAdjustHandles <> False then 
    AWriter.AddAttribute('noAdjustHandles',XmlBoolToStr(FNoAdjustHandles));
  if FNoChangeArrowheads <> False then 
    AWriter.AddAttribute('noChangeArrowheads',XmlBoolToStr(FNoChangeArrowheads));
  if FNoChangeShapeType <> False then 
    AWriter.AddAttribute('noChangeShapeType',XmlBoolToStr(FNoChangeShapeType));
  if FNoCrop <> False then 
    AWriter.AddAttribute('noCrop',XmlBoolToStr(FNoCrop));
end;

procedure TCT_PictureLocking.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000206: FNoGrp := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000033D: FNoSelect := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000212: FNoRot := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000583: FNoChangeAspect := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000274: FNoMove := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000034F: FNoResize := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004E0: FNoEditPoints := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000607: FNoAdjustHandles := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000733: FNoChangeArrowheads := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000006B6: FNoChangeShapeType := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000271: FNoCrop := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PictureLocking.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 11;
  FNoGrp := False;
  FNoSelect := False;
  FNoRot := False;
  FNoChangeAspect := False;
  FNoMove := False;
  FNoResize := False;
  FNoEditPoints := False;
  FNoAdjustHandles := False;
  FNoChangeArrowheads := False;
  FNoChangeShapeType := False;
  FNoCrop := False;
end;

destructor TCT_PictureLocking.Destroy;
begin
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_PictureLocking.Clear;
begin
  FAssigneds := [];
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
  FNoGrp := False;
  FNoSelect := False;
  FNoRot := False;
  FNoChangeAspect := False;
  FNoMove := False;
  FNoResize := False;
  FNoEditPoints := False;
  FNoAdjustHandles := False;
  FNoChangeArrowheads := False;
  FNoChangeShapeType := False;
  FNoCrop := False;
end;

procedure TCT_PictureLocking.Assign(AItem: TCT_PictureLocking);
begin
end;

procedure TCT_PictureLocking.CopyTo(AItem: TCT_PictureLocking);
begin
end;

function  TCT_PictureLocking.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_GroupLocking }

function  TCT_GroupLocking.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FNoGrp <> False then 
    Inc(AttrsAssigned);
  if FNoUngrp <> False then 
    Inc(AttrsAssigned);
  if FNoSelect <> False then 
    Inc(AttrsAssigned);
  if FNoRot <> False then 
    Inc(AttrsAssigned);
  if FNoChangeAspect <> False then 
    Inc(AttrsAssigned);
  if FNoMove <> False then 
    Inc(AttrsAssigned);
  if FNoResize <> False then 
    Inc(AttrsAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_GroupLocking.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:extLst' then 
  begin
    if FA_ExtLst = Nil then 
      FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
    Result := FA_ExtLst;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GroupLocking.Write(AWriter: TXpgWriteXML);
begin
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.Assigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_GroupLocking.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FNoGrp <> False then 
    AWriter.AddAttribute('noGrp',XmlBoolToStr(FNoGrp));
  if FNoUngrp <> False then 
    AWriter.AddAttribute('noUngrp',XmlBoolToStr(FNoUngrp));
  if FNoSelect <> False then 
    AWriter.AddAttribute('noSelect',XmlBoolToStr(FNoSelect));
  if FNoRot <> False then 
    AWriter.AddAttribute('noRot',XmlBoolToStr(FNoRot));
  if FNoChangeAspect <> False then 
    AWriter.AddAttribute('noChangeAspect',XmlBoolToStr(FNoChangeAspect));
  if FNoMove <> False then 
    AWriter.AddAttribute('noMove',XmlBoolToStr(FNoMove));
  if FNoResize <> False then 
    AWriter.AddAttribute('noResize',XmlBoolToStr(FNoResize));
end;

procedure TCT_GroupLocking.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000206: FNoGrp := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002E9: FNoUngrp := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000033D: FNoSelect := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000212: FNoRot := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000583: FNoChangeAspect := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000274: FNoMove := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000034F: FNoResize := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_GroupLocking.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 7;
  FNoGrp := False;
  FNoUngrp := False;
  FNoSelect := False;
  FNoRot := False;
  FNoChangeAspect := False;
  FNoMove := False;
  FNoResize := False;
end;

destructor TCT_GroupLocking.Destroy;
begin
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_GroupLocking.Clear;
begin
  FAssigneds := [];
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
  FNoGrp := False;
  FNoUngrp := False;
  FNoSelect := False;
  FNoRot := False;
  FNoChangeAspect := False;
  FNoMove := False;
  FNoResize := False;
end;

procedure TCT_GroupLocking.Assign(AItem: TCT_GroupLocking);
begin
end;

procedure TCT_GroupLocking.CopyTo(AItem: TCT_GroupLocking);
begin
end;

function  TCT_GroupLocking.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_NonVisualGraphicFrameProperties }

function  TCT_NonVisualGraphicFrameProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_GraphicFrameLocks <> Nil then 
    Inc(ElemsAssigned,FA_GraphicFrameLocks.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_NonVisualGraphicFrameProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000760: begin
      if FA_GraphicFrameLocks = Nil then 
        FA_GraphicFrameLocks := TCT_GraphicalObjectFrameLocking.Create(FOwner);
      Result := FA_GraphicFrameLocks;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_NonVisualGraphicFrameProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_GraphicFrameLocks <> Nil) and FA_GraphicFrameLocks.Assigned then 
  begin
    FA_GraphicFrameLocks.WriteAttributes(AWriter);
    if xaElements in FA_GraphicFrameLocks.Assigneds then
    begin
      AWriter.BeginTag('a:graphicFrameLocks');
      FA_GraphicFrameLocks.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:graphicFrameLocks');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.Assigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

constructor TCT_NonVisualGraphicFrameProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_NonVisualGraphicFrameProperties.Destroy;
begin
  if FA_GraphicFrameLocks <> Nil then 
    FA_GraphicFrameLocks.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_NonVisualGraphicFrameProperties.Clear;
begin
  FAssigneds := [];
  if FA_GraphicFrameLocks <> Nil then 
    FreeAndNil(FA_GraphicFrameLocks);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
end;

procedure TCT_NonVisualGraphicFrameProperties.Assign(AItem: TCT_NonVisualGraphicFrameProperties);
begin
end;

procedure TCT_NonVisualGraphicFrameProperties.CopyTo(AItem: TCT_NonVisualGraphicFrameProperties);
begin
end;

function  TCT_NonVisualGraphicFrameProperties.Create_A_GraphicFrameLocks: TCT_GraphicalObjectFrameLocking;
begin
  Result := TCT_GraphicalObjectFrameLocking.Create(FOwner);
  FA_GraphicFrameLocks := Result;
end;

function  TCT_NonVisualGraphicFrameProperties.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_GraphicalObjectData }

function  TCT_GraphicalObjectData.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FUri <> '' then
    Inc(AttrsAssigned);
  if FChart <> Nil then
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function TCT_GraphicalObjectData.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  i,j    : integer;
  Stream : TStream;
  Item   : TOPCItem;
  Item2  : TOPCItem;
  Drawing: TXPGDocXLSXDrawing;
begin
  if AReader.QName = 'c:chart' then begin
    for i := 0 to AReader.Attributes.Count - 1 do begin
      if AReader.Attributes[i] = 'r:id' then begin
        FChart := TXPGDocXLSXChart.Create(FOwner.GrManager);

        Item := FOwner.GrManager.XLSOPC.FindAndOpenChart(AReader.Attributes.Values[i]);
        for j := 0 to Item.Count - 1 do begin
          if Item[j].Type_ = OPC_XLSX_CHARTSTYLE then begin
            FChart.AddStyle;
            Stream := FOwner.GrManager.XLSOPC.ReadChartStyle(Item,Item[j].Id);
            try
              FChart.Style.LoadFromStream(Stream);
            finally
              Stream.Free;
            end;
          end
          else if Item[j].Type_ = OPC_XLSX_CHARTCOLORS then begin
            FChart.AddColors;
            Stream := FOwner.GrManager.XLSOPC.ReadChartColors(Item,Item[j].Id);
            try
              FChart.Colors.LoadFromStream(Stream);
            finally
              Stream.Free;
            end;
          end;
        end;
        Item.Close;

        Item := FOwner.GrManager.XLSOPC.ItemOpenRead(FOwner.GrManager.XLSOPC.CurrDrawing,OPC_XLSX_CHART,AReader.Attributes.Values[i]);
        Stream := FOwner.GrManager.XLSOPC.ItemOpenStream(Item);
        try
          FChart.LoadFromStream(Stream);

          FOwner.GrManager.XLSOPC.ItemClose(Item);
        finally
          Stream.Free;
        end;

        if FChart.ChartSpace.UserShapes <> Nil then begin
          Item2 := FOwner.GrManager.XLSOPC.ItemOpenRead(Item,OPC_XLSX_CHARTUSERSHAPES,FChart.ChartSpace.UserShapes.R_Id);
          Stream := FOwner.GrManager.XLSOPC.ItemOpenStream(Item2);
          try
            Drawing := TXPGDocXLSXDrawing.Create(FOwner.GrManager);

            Drawing._LoadFromStream(Stream,True);

            // FChart.ChartSpace owns Drawing.
            FChart.ChartSpace.AddDrawing(Drawing);

            FOwner.GrManager.XLSOPC.ItemClose(Item2);
          finally
            Stream.Free;
          end;
        end;
      end;
    end;
    AReader.Attributes.Clear;
  end;
  Result := Self;
end;

procedure TCT_GraphicalObjectData.Write(AWriter: TXpgWriteXML);
var
  opcChart: TOPCItem;
  opcDrw  : TOPCItem;
  OPC     : TOPCItem;
  Stream  : TMemoryStream;
  Drw     : TXPGDocXLSXDrawing;
begin
  if FChart <> Nil then begin
    opcDrw := Nil;

    Stream := TMemoryStream.Create;
    try
      opcChart := FOwner.GrManager.XLSOPC.CreateChart;

      if FChart.ChartSpace.Drawing <> Nil then begin
        opcDrw := FOwner.GrManager.XLSOPC.CreateChartUserShapes(opcChart,FOwner.GrManager.XLSOPC.CurrSheet.Parent.Count * 1000);
//        opcDrw := FOwner.GrManager.XLSOPC.CreateChartUserShapes(opcChart,L_DrwCounter);
        Inc(L_DrwCounter);


        FChart.ChartSpace.Create_UserShapes;
        FChart.ChartSpace.UserShapes.R_Id := opcDrw.Id;
      end;

      FChart.SaveToStream(Stream);
      FOwner.GrManager.XLSOPC.ItemWrite(opcChart,Stream);
      FOwner.GrManager.XLSOPC.CloseChart(opcChart);

      AWriter.AddAttribute('xmlns:c',OOXML_URI_OFFICEDOC_CHART);
      AWriter.AddAttribute('xmlns:r',OOXML_URI_OFFICEDOC_RELATIONSHIPS);
      AWriter.AddAttribute('r:id',opcChart.Id);
      AWriter.SimpleTag('c:chart');

      if FChart.Style <> Nil then begin
        Stream.Clear;
        FChart.Style.SaveToStream(Stream);
        OPC := FOwner.GrManager.XLSOPC.CreateChartStyle(opcChart);
        FOwner.GrManager.XLSOPC.ItemWrite(OPC,Stream);
        FOwner.GrManager.XLSOPC.ItemClose(OPC);
      end;
      if FChart.Colors <> Nil then begin
        Stream.Clear;
        FChart.Colors.SaveToStream(Stream);
        OPC := FOwner.GrManager.XLSOPC.CreateChartColors(opcChart);
        FOwner.GrManager.XLSOPC.ItemWrite(OPC,Stream);
        FOwner.GrManager.XLSOPC.ItemClose(OPC);
      end;

      if FChart.ChartSpace.Drawing <> Nil then begin
        Stream.Clear;

        Drw := TXPGDocXLSXDrawing(FChart.ChartSpace.Drawing);

        Drw.SaveToStream(Stream,True);

        FOwner.GrManager.XLSOPC.ItemWrite(opcDrw,Stream);

        FOwner.GrManager.XLSOPC.ItemClose(opcDrw);
      end;
    finally
      Stream.Free;
    end;
  end;

//  FAnyElements.Write(AWriter);
end;

procedure TCT_GraphicalObjectData.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FUri <> '' then
    AWriter.AddAttribute('uri',FUri);
end;

procedure TCT_GraphicalObjectData.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'uri' then
    FUri := AAttributes.Values[0]
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_GraphicalObjectData.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FUri := '';
end;

function TCT_GraphicalObjectData.Create_Chart: TXPGDocXLSXChart;
begin
  Result := TXPGDocXLSXChart.Create(FOwner.GrManager);
  FChart := Result;
end;

destructor TCT_GraphicalObjectData.Destroy;
begin
  if FChart <> Nil then
    FChart.Free;
end;

procedure TCT_GraphicalObjectData.Clear;
begin
  FAssigneds := [];
  if FChart <> Nil then begin
    FChart.Free;
    FChart := Nil;
  end;
  FUri := '';
end;

procedure TCT_GraphicalObjectData.Assign(AItem: TCT_GraphicalObjectData);
begin
end;

procedure TCT_GraphicalObjectData.CopyTo(AItem: TCT_GraphicalObjectData);
begin
end;

{ TCT_NonVisualConnectorProperties }

function  TCT_NonVisualConnectorProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_CxnSpLocks <> Nil then 
    Inc(ElemsAssigned,FA_CxnSpLocks.CheckAssigned);
  if FA_StCxn <> Nil then 
    Inc(ElemsAssigned,FA_StCxn.CheckAssigned);
  if FA_EndCxn <> Nil then 
    Inc(ElemsAssigned,FA_EndCxn.CheckAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_NonVisualConnectorProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000004A3: begin
      if FA_CxnSpLocks = Nil then 
        FA_CxnSpLocks := TCT_ConnectorLocking.Create(FOwner);
      Result := FA_CxnSpLocks;
    end;
    $000002AB: begin
      if FA_StCxn = Nil then 
        FA_StCxn := TCT_Connection.Create(FOwner);
      Result := FA_StCxn;
    end;
    $000002FB: begin
      if FA_EndCxn = Nil then 
        FA_EndCxn := TCT_Connection.Create(FOwner);
      Result := FA_EndCxn;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_NonVisualConnectorProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_CxnSpLocks <> Nil) and FA_CxnSpLocks.Assigned then 
  begin
    FA_CxnSpLocks.WriteAttributes(AWriter);
    if xaElements in FA_CxnSpLocks.Assigneds then
    begin
      AWriter.BeginTag('a:cxnSpLocks');
      FA_CxnSpLocks.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:cxnSpLocks');
  end;
  if (FA_StCxn <> Nil) and FA_StCxn.Assigned then 
  begin
    FA_StCxn.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:stCxn');
  end;
  if (FA_EndCxn <> Nil) and FA_EndCxn.Assigned then 
  begin
    FA_EndCxn.WriteAttributes(AWriter);
    AWriter.SimpleTag('a:endCxn');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.Assigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

constructor TCT_NonVisualConnectorProperties.Create(AOwner: TXPGDocBase);
begin
  inherited Create;
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
end;

destructor TCT_NonVisualConnectorProperties.Destroy;
begin
  if FA_CxnSpLocks <> Nil then 
    FA_CxnSpLocks.Free;
  if FA_StCxn <> Nil then 
    FA_StCxn.Free;
  if FA_EndCxn <> Nil then 
    FA_EndCxn.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_NonVisualConnectorProperties.Clear;
begin
  FAssigneds := [];
  if FA_CxnSpLocks <> Nil then 
    FreeAndNil(FA_CxnSpLocks);
  if FA_StCxn <> Nil then 
    FreeAndNil(FA_StCxn);
  if FA_EndCxn <> Nil then 
    FreeAndNil(FA_EndCxn);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
end;

procedure TCT_NonVisualConnectorProperties.Assign(AItem: TCT_NonVisualConnectorProperties);
begin
end;

procedure TCT_NonVisualConnectorProperties.CopyTo(AItem: TCT_NonVisualConnectorProperties);
begin
end;

function  TCT_NonVisualConnectorProperties.Create_A_CxnSpLocks: TCT_ConnectorLocking;
begin
  Result := TCT_ConnectorLocking.Create(FOwner);
  FA_CxnSpLocks := Result;
end;

function  TCT_NonVisualConnectorProperties.Create_A_StCxn: TCT_Connection;
begin
  Result := TCT_Connection.Create(FOwner);
  FA_StCxn := Result;
end;

function  TCT_NonVisualConnectorProperties.Create_A_EndCxn: TCT_Connection;
begin
  Result := TCT_Connection.Create(FOwner);
  FA_EndCxn := Result;
end;

function  TCT_NonVisualConnectorProperties.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_NonVisualPictureProperties }

function  TCT_NonVisualPictureProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 1;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPreferRelativeResize <> True then
    Inc(AttrsAssigned);
  if FA_PicLocks <> Nil then
    Inc(ElemsAssigned,FA_PicLocks.CheckAssigned);
  if FA_ExtLst <> Nil then
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_NonVisualPictureProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000003D3: begin
      if FA_PicLocks = Nil then 
        FA_PicLocks := TCT_PictureLocking.Create(FOwner);
      Result := FA_PicLocks;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_NonVisualPictureProperties.Write(AWriter: TXpgWriteXML);
begin
  if (FA_PicLocks <> Nil) and FA_PicLocks.Assigned then 
  begin
    FA_PicLocks.WriteAttributes(AWriter);
    if xaElements in FA_PicLocks.Assigneds then
    begin
      AWriter.BeginTag('a:picLocks');
      FA_PicLocks.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:picLocks');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.Assigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

procedure TCT_NonVisualPictureProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPreferRelativeResize <> True then 
    AWriter.AddAttribute('preferRelativeResize',XmlBoolToStr(FPreferRelativeResize));
end;

procedure TCT_NonVisualPictureProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'preferRelativeResize' then 
    FPreferRelativeResize := XmlStrToBoolDef(AAttributes.Values[0],True)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_NonVisualPictureProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 1;
  FPreferRelativeResize := True;
end;

destructor TCT_NonVisualPictureProperties.Destroy;
begin
  if FA_PicLocks <> Nil then 
    FA_PicLocks.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_NonVisualPictureProperties.Clear;
begin
  FAssigneds := [];
  if FA_PicLocks <> Nil then 
    FreeAndNil(FA_PicLocks);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
  FPreferRelativeResize := True;
end;

procedure TCT_NonVisualPictureProperties.Assign(AItem: TCT_NonVisualPictureProperties);
begin
end;

procedure TCT_NonVisualPictureProperties.CopyTo(AItem: TCT_NonVisualPictureProperties);
begin
end;

function  TCT_NonVisualPictureProperties.Create_A_PicLocks: TCT_PictureLocking;
begin
  Result := TCT_PictureLocking.Create(FOwner);
  FA_PicLocks := Result;
end;

function  TCT_NonVisualPictureProperties.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_NonVisualGroupDrawingShapeProps }

function  TCT_NonVisualGroupDrawingShapeProps.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_GrpSpLocks <> Nil then 
    Inc(ElemsAssigned);
  if FA_ExtLst <> Nil then 
    Inc(ElemsAssigned,FA_ExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_NonVisualGroupDrawingShapeProps.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000004A3: begin
      if FA_GrpSpLocks = Nil then 
        FA_GrpSpLocks := TCT_GroupLocking.Create(FOwner);
      Result := FA_GrpSpLocks;
    end;
    $0000031F: begin
      if FA_ExtLst = Nil then 
        FA_ExtLst := TCT_OfficeArtExtensionList.Create(FOwner);
      Result := FA_ExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_NonVisualGroupDrawingShapeProps.Write(AWriter: TXpgWriteXML);
begin
//  if (FA_GrpSpLocks <> Nil) and FA_GrpSpLocks.Assigned then
  if FA_GrpSpLocks <> Nil then
  begin
    FA_GrpSpLocks.WriteAttributes(AWriter);
    if xaElements in FA_GrpSpLocks.Assigneds then
    begin
      AWriter.BeginTag('a:grpSpLocks');
      FA_GrpSpLocks.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:grpSpLocks');
  end;
  if (FA_ExtLst <> Nil) and FA_ExtLst.Assigned then 
    if xaElements in FA_ExtLst.Assigneds then
    begin
      AWriter.BeginTag('a:extLst');
      FA_ExtLst.Write(AWriter);
      AWriter.EndTag;
    end;
end;

constructor TCT_NonVisualGroupDrawingShapeProps.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_NonVisualGroupDrawingShapeProps.Destroy;
begin
  if FA_GrpSpLocks <> Nil then 
    FA_GrpSpLocks.Free;
  if FA_ExtLst <> Nil then 
    FA_ExtLst.Free;
end;

procedure TCT_NonVisualGroupDrawingShapeProps.Clear;
begin
  FAssigneds := [];
  if FA_GrpSpLocks <> Nil then 
    FreeAndNil(FA_GrpSpLocks);
  if FA_ExtLst <> Nil then 
    FreeAndNil(FA_ExtLst);
end;

procedure TCT_NonVisualGroupDrawingShapeProps.Assign(AItem: TCT_NonVisualGroupDrawingShapeProps);
begin
end;

procedure TCT_NonVisualGroupDrawingShapeProps.CopyTo(AItem: TCT_NonVisualGroupDrawingShapeProps);
begin
end;

function  TCT_NonVisualGroupDrawingShapeProps.Create_A_GrpSpLocks: TCT_GroupLocking;
begin
  Result := TCT_GroupLocking.Create(FOwner);
  FA_GrpSpLocks := Result;
end;

function  TCT_NonVisualGroupDrawingShapeProps.Create_A_ExtLst: TCT_OfficeArtExtensionList;
begin
  Result := TCT_OfficeArtExtensionList.Create(FOwner);
  FA_ExtLst := Result;
end;

{ TCT_GraphicalObjectFrameNonVisual }

function  TCT_GraphicalObjectFrameNonVisual.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FCNvPr <> Nil then 
    Inc(ElemsAssigned,FCNvPr.CheckAssigned);
  if FCNvGraphicFramePr <> Nil then 
    Inc(ElemsAssigned,FCNvGraphicFramePr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_GraphicalObjectFrameNonVisual.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000371: begin
      if FCNvPr = Nil then 
        FCNvPr := TCT_NonVisualDrawingProps.Create(FOwner);
      Result := FCNvPr;
    end;
    $0000081A: begin
      if FCNvGraphicFramePr = Nil then 
        FCNvGraphicFramePr := TCT_NonVisualGraphicFrameProperties.Create(FOwner);
      Result := FCNvGraphicFramePr;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GraphicalObjectFrameNonVisual.Write(AWriter: TXpgWriteXML);
begin
  if (FCNvPr <> Nil) and FCNvPr.Assigned then
  begin
    FCNvPr.WriteAttributes(AWriter);
    if xaElements in FCNvPr.Assigneds then
    begin
      AWriter.BeginTag('xdr:cNvPr');
      FCNvPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('xdr:cNvPr');
  end;
  if (FCNvGraphicFramePr <> Nil) and FCNvGraphicFramePr.Assigned then begin
      AWriter.BeginTag('xdr:cNvGraphicFramePr');
      FCNvGraphicFramePr.Write(AWriter);
      AWriter.EndTag;
  end
  else
    AWriter.SimpleTag('xdr:cNvGraphicFramePr');
end;

constructor TCT_GraphicalObjectFrameNonVisual.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_GraphicalObjectFrameNonVisual.Destroy;
begin
  if FCNvPr <> Nil then 
    FCNvPr.Free;
  if FCNvGraphicFramePr <> Nil then 
    FCNvGraphicFramePr.Free;
end;

procedure TCT_GraphicalObjectFrameNonVisual.Clear;
begin
  FAssigneds := [];
  if FCNvPr <> Nil then 
    FreeAndNil(FCNvPr);
  if FCNvGraphicFramePr <> Nil then 
    FreeAndNil(FCNvGraphicFramePr);
end;

procedure TCT_GraphicalObjectFrameNonVisual.Assign(AItem: TCT_GraphicalObjectFrameNonVisual);
begin
end;

procedure TCT_GraphicalObjectFrameNonVisual.CopyTo(AItem: TCT_GraphicalObjectFrameNonVisual);
begin
end;

function  TCT_GraphicalObjectFrameNonVisual.Create_CNvPr: TCT_NonVisualDrawingProps;
begin
  Result := TCT_NonVisualDrawingProps.Create(FOwner);
  FCNvPr := Result;
end;

function  TCT_GraphicalObjectFrameNonVisual.Create_CNvGraphicFramePr: TCT_NonVisualGraphicFrameProperties;
begin
  Result := TCT_NonVisualGraphicFrameProperties.Create(FOwner);
  FCNvGraphicFramePr := Result;
end;

{ TCT_GraphicalObject }

function  TCT_GraphicalObject.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FA_GraphicData <> Nil then 
    Inc(ElemsAssigned,FA_GraphicData.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_GraphicalObject.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'a:graphicData' then 
  begin
    if FA_GraphicData = Nil then 
      FA_GraphicData := TCT_GraphicalObjectData.Create(FOwner);
    Result := FA_GraphicData;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GraphicalObject.Write(AWriter: TXpgWriteXML);
begin
  if (FA_GraphicData <> Nil) and FA_GraphicData.Assigned then 
  begin
    FA_GraphicData.WriteAttributes(AWriter);
    if xaElements in FA_GraphicData.Assigneds then
    begin
      AWriter.BeginTag('a:graphicData');
      FA_GraphicData.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:graphicData');
  end;
end;

constructor TCT_GraphicalObject.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_GraphicalObject.Destroy;
begin
  if FA_GraphicData <> Nil then 
    FA_GraphicData.Free;
end;

procedure TCT_GraphicalObject.Clear;
begin
  FAssigneds := [];
  if FA_GraphicData <> Nil then 
    FreeAndNil(FA_GraphicData);
end;

procedure TCT_GraphicalObject.Assign(AItem: TCT_GraphicalObject);
begin
end;

procedure TCT_GraphicalObject.CopyTo(AItem: TCT_GraphicalObject);
begin
end;

function  TCT_GraphicalObject.Create_A_GraphicData: TCT_GraphicalObjectData;
begin
  Result := TCT_GraphicalObjectData.Create(FOwner);
  FA_GraphicData := Result;
end;

{ TCT_ConnectorNonVisual }

function  TCT_ConnectorNonVisual.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FCNvPr <> Nil then 
    Inc(ElemsAssigned,FCNvPr.CheckAssigned);
  if FCNvCxnSpPr <> Nil then 
    Inc(ElemsAssigned,FCNvCxnSpPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ConnectorNonVisual.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000371: begin
      if FCNvPr = Nil then
        FCNvPr := TCT_NonVisualDrawingProps.Create(FOwner);
      Result := FCNvPr;
    end;
    $0000055D: begin
      if FCNvCxnSpPr = Nil then 
        FCNvCxnSpPr := TCT_NonVisualConnectorProperties.Create(FOwner);
      Result := FCNvCxnSpPr;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ConnectorNonVisual.Write(AWriter: TXpgWriteXML);
begin
  if (FCNvPr <> Nil) and FCNvPr.Assigned then 
  begin
    FCNvPr.WriteAttributes(AWriter);
    if xaElements in FCNvPr.Assigneds then
    begin
      AWriter.BeginTag('xdr:cNvPr');
      FCNvPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:cNvPr');
  end;
  if FCNvCxnSpPr <> Nil then
    if xaElements in FCNvCxnSpPr.Assigneds then
    begin
      AWriter.BeginTag('xdr:cNvCxnSpPr');
      FCNvCxnSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:cNvCxnSpPr');
end;

constructor TCT_ConnectorNonVisual.Create(AOwner: TXPGDocBase);
begin
  inherited Create;

  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_ConnectorNonVisual.Destroy;
begin
  if FCNvPr <> Nil then 
    FCNvPr.Free;
  if FCNvCxnSpPr <> Nil then 
    FCNvCxnSpPr.Free;
end;

procedure TCT_ConnectorNonVisual.Clear;
begin
  FAssigneds := [];
  if FCNvPr <> Nil then 
    FreeAndNil(FCNvPr);
  if FCNvCxnSpPr <> Nil then 
    FreeAndNil(FCNvCxnSpPr);
end;

procedure TCT_ConnectorNonVisual.Assign(AItem: TCT_ConnectorNonVisual);
begin
end;

procedure TCT_ConnectorNonVisual.CopyTo(AItem: TCT_ConnectorNonVisual);
begin
end;

function  TCT_ConnectorNonVisual.Create_CNvPr: TCT_NonVisualDrawingProps;
begin
  Result := TCT_NonVisualDrawingProps.Create(FOwner);
  FCNvPr := Result;
end;

function  TCT_ConnectorNonVisual.Create_CNvCxnSpPr: TCT_NonVisualConnectorProperties;
begin
  Result := TCT_NonVisualConnectorProperties.Create(FOwner);
  FCNvCxnSpPr := Result;
end;

{ TCT_PictureNonVisual }

function  TCT_PictureNonVisual.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FCNvPr <> Nil then
    Inc(ElemsAssigned,FCNvPr.CheckAssigned);
  if FCNvPicPr <> Nil then
    Inc(ElemsAssigned,FCNvPicPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PictureNonVisual.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000371: begin
      if FCNvPr = Nil then
        FCNvPr := TCT_NonVisualDrawingProps.Create(FOwner);
      Result := FCNvPr;
    end;
    $0000048D: begin
      if FCNvPicPr = Nil then
        FCNvPicPr := TCT_NonVisualPictureProperties.Create(FOwner);
      Result := FCNvPicPr;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PictureNonVisual.Write(AWriter: TXpgWriteXML);
begin
  if (FCNvPr <> Nil) and FCNvPr.Assigned then
  begin
    FCNvPr.WriteAttributes(AWriter);
    if xaElements in FCNvPr.Assigneds then
    begin
      AWriter.BeginTag('xdr:cNvPr');
      FCNvPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('xdr:cNvPr');
  end;
//  TODO Must be written if this is written.
//  if (FCNvPicPr <> Nil) and FCNvPicPr.Assigned then
//  begin
    FCNvPicPr.WriteAttributes(AWriter);
    if xaElements in FCNvPicPr.Assigneds then
    begin
      AWriter.BeginTag('xdr:cNvPicPr');
      FCNvPicPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('xdr:cNvPicPr');
//  end;
end;

constructor TCT_PictureNonVisual.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_PictureNonVisual.Destroy;
begin
  if FCNvPr <> Nil then 
    FCNvPr.Free;
  if FCNvPicPr <> Nil then 
    FCNvPicPr.Free;
end;

procedure TCT_PictureNonVisual.Clear;
begin
  FAssigneds := [];
  if FCNvPr <> Nil then 
    FreeAndNil(FCNvPr);
  if FCNvPicPr <> Nil then 
    FreeAndNil(FCNvPicPr);
end;

procedure TCT_PictureNonVisual.Assign(AItem: TCT_PictureNonVisual);
begin
end;

procedure TCT_PictureNonVisual.CopyTo(AItem: TCT_PictureNonVisual);
begin
end;

function  TCT_PictureNonVisual.Create_CNvPr: TCT_NonVisualDrawingProps;
begin
  Result := TCT_NonVisualDrawingProps.Create(FOwner);
  FCNvPr := Result;
end;

function  TCT_PictureNonVisual.Create_CNvPicPr: TCT_NonVisualPictureProperties;
begin
  Result := TCT_NonVisualPictureProperties.Create(FOwner);
  FCNvPicPr := Result;
end;

{ TCT_GroupShapeNonVisual }

function  TCT_GroupShapeNonVisual.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FCNvPr <> Nil then 
    Inc(ElemsAssigned,FCNvPr.CheckAssigned);
  if FCNvGrpSpPr <> Nil then 
    Inc(ElemsAssigned,FCNvGrpSpPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_GroupShapeNonVisual.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000371: begin
      if FCNvPr = Nil then 
        FCNvPr := TCT_NonVisualDrawingProps.Create(FOwner);
      Result := FCNvPr;
    end;
    $0000055D: begin
      if FCNvGrpSpPr = Nil then 
        FCNvGrpSpPr := TCT_NonVisualGroupDrawingShapeProps.Create(FOwner);
      Result := FCNvGrpSpPr;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GroupShapeNonVisual.Write(AWriter: TXpgWriteXML);
begin
  if (FCNvPr <> Nil) and FCNvPr.Assigned then
  begin
    FCNvPr.WriteAttributes(AWriter);
    AWriter.BeginTag('xdr:cNvPr');
    FCNvPr.Write(AWriter);
    AWriter.EndTag;
  end
  else
    AWriter.SimpleTag('xdr:cNvPr');

  if (FCNvGrpSpPr <> Nil) and FCNvGrpSpPr.Assigned then 
    if xaElements in FCNvGrpSpPr.Assigneds then
    begin
      AWriter.BeginTag('xdr:cNvGrpSpPr');
      FCNvGrpSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:cNvGrpSpPr');
end;

constructor TCT_GroupShapeNonVisual.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
end;

destructor TCT_GroupShapeNonVisual.Destroy;
begin
  if FCNvPr <> Nil then 
    FCNvPr.Free;
  if FCNvGrpSpPr <> Nil then 
    FCNvGrpSpPr.Free;
end;

procedure TCT_GroupShapeNonVisual.Clear;
begin
  FAssigneds := [];
  if FCNvPr <> Nil then 
    FreeAndNil(FCNvPr);
  if FCNvGrpSpPr <> Nil then 
    FreeAndNil(FCNvGrpSpPr);
end;

procedure TCT_GroupShapeNonVisual.Assign(AItem: TCT_GroupShapeNonVisual);
begin
end;

procedure TCT_GroupShapeNonVisual.CopyTo(AItem: TCT_GroupShapeNonVisual);
begin
end;

function  TCT_GroupShapeNonVisual.Create_CNvPr: TCT_NonVisualDrawingProps;
begin
  Result := TCT_NonVisualDrawingProps.Create(FOwner);
  FCNvPr := Result;
end;

function  TCT_GroupShapeNonVisual.Create_CNvGrpSpPr: TCT_NonVisualGroupDrawingShapeProps;
begin
  Result := TCT_NonVisualGroupDrawingShapeProps.Create(FOwner);
  FCNvGrpSpPr := Result;
end;

{ TCT_ShapeXpgList }

function  TCT_ShapeXpgList.GetItems(Index: integer): TCT_Shape;
begin
  Result := TCT_Shape(inherited Items[Index]);
end;

function  TCT_ShapeXpgList.Add: TCT_Shape;
begin
  Result := TCT_Shape.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ShapeXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ShapeXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].Assigneds then
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_ShapeXpgList.Assign(AItem: TCT_ShapeXpgList);
begin
end;

procedure TCT_ShapeXpgList.CopyTo(AItem: TCT_ShapeXpgList);
begin
end;

{ TCT_GraphicalObjectFrame }

function  TCT_GraphicalObjectFrame.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FMacro <> '' then 
    Inc(AttrsAssigned);
  if FFPublished <> False then 
    Inc(AttrsAssigned);
  if FNvGraphicFramePr <> Nil then 
    Inc(ElemsAssigned,FNvGraphicFramePr.CheckAssigned);
  if FXfrm <> Nil then 
    Inc(ElemsAssigned,FXfrm.CheckAssigned);
  if FA_Graphic <> Nil then 
    Inc(ElemsAssigned,FA_Graphic.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_GraphicalObjectFrame.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000007D7: begin
      if FNvGraphicFramePr = Nil then 
        FNvGraphicFramePr := TCT_GraphicalObjectFrameNonVisual.Create(FOwner);
      Result := FNvGraphicFramePr;
    end;
    $00000345: begin
      if FXfrm = Nil then 
        FXfrm := TCT_Transform2D.Create(FOwner);
      Result := FXfrm;
    end;
    $00000379: begin
      if FA_Graphic = Nil then 
        FA_Graphic := TCT_GraphicalObject.Create(FOwner);
      Result := FA_Graphic;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

function TCT_GraphicalObjectFrame.IsChart: boolean;
begin
  Result := (FA_Graphic <> Nil) and (FA_Graphic.GraphicData <> Nil) and (FA_Graphic.GraphicData.Chart <> Nil);
end;

procedure TCT_GraphicalObjectFrame.Write(AWriter: TXpgWriteXML);
begin
  if (FNvGraphicFramePr <> Nil) and FNvGraphicFramePr.Assigned then 
    if xaElements in FNvGraphicFramePr.Assigneds then
    begin
      AWriter.BeginTag('xdr:nvGraphicFramePr');
      FNvGraphicFramePr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:nvGraphicFramePr')
  else 
    AWriter.SimpleTag('xdr:nvGraphicFramePr');
  if (FXfrm <> Nil) and FXfrm.Assigned then 
  begin
    FXfrm.WriteAttributes(AWriter);
    if xaElements in FXfrm.Assigneds then
    begin
      AWriter.BeginTag('xdr:xfrm');
      FXfrm.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:xfrm');
  end
  else 
    AWriter.SimpleTag('xdr:xfrm');
  if (FA_Graphic <> Nil) and FA_Graphic.Assigned then 
    if xaElements in FA_Graphic.Assigneds then
    begin
      AWriter.BeginTag('a:graphic');
      FA_Graphic.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('a:graphic')
  else 
    AWriter.SimpleTag('a:graphic');
end;

procedure TCT_GraphicalObjectFrame.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FMacro <> '' then 
    AWriter.AddAttribute('macro',FMacro);
  if FFPublished <> False then 
    AWriter.AddAttribute('fPublished',XmlBoolToStr(FFPublished));
end;

procedure TCT_GraphicalObjectFrame.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000212: FMacro := AAttributes.Values[i];
      $00000406: FFPublished := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_GraphicalObjectFrame.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 2;
  FFPublished := False;
end;

procedure TCT_GraphicalObjectFrame.CreateDefault(AId: integer; AName: AxUCString);
var
  CNvPr : TCT_NonVisualDrawingProps;
begin
  FNvGraphicFramePr.Create_CNvGraphicFramePr;
  CNvPr := FNvGraphicFramePr.Create_CNvPr;
  CNvPr.Id := AId;
  CNvPr.Name := AName;

  FXfrm.Create_A_Off;
  FXfrm.A_Off.X := 0;
  FXfrm.A_Off.Y := 0;
  FXfrm.Create_A_Ext;
  FXfrm.A_Ext.Cx := 0;
  FXfrm.A_Ext.Cy := 0;
end;

destructor TCT_GraphicalObjectFrame.Destroy;
begin
  if FNvGraphicFramePr <> Nil then 
    FNvGraphicFramePr.Free;
  if FXfrm <> Nil then 
    FXfrm.Free;
  if FA_Graphic <> Nil then 
    FA_Graphic.Free;
end;

procedure TCT_GraphicalObjectFrame.Clear;
begin
  FAssigneds := [];
  if FNvGraphicFramePr <> Nil then 
    FreeAndNil(FNvGraphicFramePr);
  if FXfrm <> Nil then 
    FreeAndNil(FXfrm);
  if FA_Graphic <> Nil then 
    FreeAndNil(FA_Graphic);
  FMacro := '';
  FFPublished := False;
end;

procedure TCT_GraphicalObjectFrame.Assign(AItem: TCT_GraphicalObjectFrame);
begin
end;

procedure TCT_GraphicalObjectFrame.CopyTo(AItem: TCT_GraphicalObjectFrame);
begin
end;

function  TCT_GraphicalObjectFrame.Create_NvGraphicFramePr: TCT_GraphicalObjectFrameNonVisual;
begin
  Result := TCT_GraphicalObjectFrameNonVisual.Create(FOwner);
  FNvGraphicFramePr := Result;
end;

function  TCT_GraphicalObjectFrame.Create_Xfrm: TCT_Transform2D;
begin
  Result := TCT_Transform2D.Create(FOwner);
  FXfrm := Result;
end;

function  TCT_GraphicalObjectFrame.Create_A_Graphic: TCT_GraphicalObject;
begin
  Result := TCT_GraphicalObject.Create(FOwner);
  FA_Graphic := Result;
end;

{ TCT_GraphicalObjectFrameXpgList }

function  TCT_GraphicalObjectFrameXpgList.GetItems(Index: integer): TCT_GraphicalObjectFrame;
begin
  Result := TCT_GraphicalObjectFrame(inherited Items[Index]);
end;

function  TCT_GraphicalObjectFrameXpgList.Add: TCT_GraphicalObjectFrame;
begin
  Result := TCT_GraphicalObjectFrame.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_GraphicalObjectFrameXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_GraphicalObjectFrameXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].Assigneds then
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_GraphicalObjectFrameXpgList.Assign(AItem: TCT_GraphicalObjectFrameXpgList);
begin
end;

procedure TCT_GraphicalObjectFrameXpgList.CopyTo(AItem: TCT_GraphicalObjectFrameXpgList);
begin
end;

{ TCT_Connector }

function  TCT_Connector.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FMacro <> '' then 
    Inc(AttrsAssigned);
  if FFPublished <> False then 
    Inc(AttrsAssigned);
  if FNvCxnSpPr <> Nil then 
    Inc(ElemsAssigned,FNvCxnSpPr.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FStyle <> Nil then 
    Inc(ElemsAssigned,FStyle.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Connector.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000051A: begin
      if FNvCxnSpPr = Nil then 
        FNvCxnSpPr := TCT_ConnectorNonVisual.Create(FOwner);
      Result := FNvCxnSpPr;
    end;
    $0000032D: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $000003B9: begin
      if FStyle = Nil then 
        FStyle := TCT_ShapeStyle.Create(FOwner);
      Result := FStyle;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Connector.Write(AWriter: TXpgWriteXML);
begin
  if (FNvCxnSpPr <> Nil) and FNvCxnSpPr.Assigned then 
    if xaElements in FNvCxnSpPr.Assigneds then
    begin
      AWriter.BeginTag('xdr:nvCxnSpPr');
      FNvCxnSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:nvCxnSpPr')
  else 
    AWriter.SimpleTag('xdr:nvCxnSpPr');
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('xdr:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:spPr');
  end
  else 
    AWriter.SimpleTag('xdr:spPr');
  if (FStyle <> Nil) and FStyle.Assigned then 
    if xaElements in FStyle.Assigneds then
    begin
      AWriter.BeginTag('xdr:style');
      FStyle.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:style');
end;

procedure TCT_Connector.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FMacro <> '' then 
    AWriter.AddAttribute('macro',FMacro);
  if FFPublished <> False then 
    AWriter.AddAttribute('fPublished',XmlBoolToStr(FFPublished));
end;

procedure TCT_Connector.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000212: FMacro := AAttributes.Values[i];
      $00000406: FFPublished := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Connector.Create(AOwner: TXPGDocBase);
begin
  inherited Create;

  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 2;
  FFPublished := False;
end;

destructor TCT_Connector.Destroy;
begin
  if FNvCxnSpPr <> Nil then 
    FNvCxnSpPr.Free;
  if FSpPr <> Nil then 
    FSpPr.Free;
  if FStyle <> Nil then 
    FStyle.Free;
end;

procedure TCT_Connector.Clear;
begin
  FAssigneds := [];
  if FNvCxnSpPr <> Nil then 
    FreeAndNil(FNvCxnSpPr);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FStyle <> Nil then 
    FreeAndNil(FStyle);
  FMacro := '';
  FFPublished := False;
end;

procedure TCT_Connector.Assign(AItem: TCT_Connector);
begin
end;

procedure TCT_Connector.CopyTo(AItem: TCT_Connector);
begin
end;

function  TCT_Connector.Create_NvCxnSpPr: TCT_ConnectorNonVisual;
begin
  Result := TCT_ConnectorNonVisual.Create(FOwner);
  FNvCxnSpPr := Result;
end;

function  TCT_Connector.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_Connector.Create_Style: TCT_ShapeStyle;
begin
  Result := TCT_ShapeStyle.Create(FOwner);
  FStyle := Result;
end;

{ TCT_ConnectorXpgList }

function  TCT_ConnectorXpgList.GetItems(Index: integer): TCT_Connector;
begin
  Result := TCT_Connector(inherited Items[Index]);
end;

function  TCT_ConnectorXpgList.Add: TCT_Connector;
begin
  Result := TCT_Connector.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ConnectorXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ConnectorXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].Assigneds then
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_ConnectorXpgList.Assign(AItem: TCT_ConnectorXpgList);
begin
end;

procedure TCT_ConnectorXpgList.CopyTo(AItem: TCT_ConnectorXpgList);
begin
end;

{ TCT_Picture }

function  TCT_Picture.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FMacro <> '' then 
    Inc(AttrsAssigned);
  if FFPublished <> False then 
    Inc(AttrsAssigned);
  if FNvPicPr <> Nil then 
    Inc(ElemsAssigned,FNvPicPr.CheckAssigned);
  if FBlipFill <> Nil then 
    Inc(ElemsAssigned,FBlipFill.CheckAssigned);
  if FSpPr <> Nil then 
    Inc(ElemsAssigned,FSpPr.CheckAssigned);
  if FStyle <> Nil then 
    Inc(ElemsAssigned,FStyle.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Picture.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000044A: begin
      if FNvPicPr = Nil then 
        FNvPicPr := TCT_PictureNonVisual.Create(FOwner);
      Result := FNvPicPr;
    end;
    $000004B6: begin
      if FBlipFill = Nil then 
        FBlipFill := TCT_BlipFillProperties.Create(FOwner);
      Result := FBlipFill;
    end;
    $0000032D: begin
      if FSpPr = Nil then 
        FSpPr := TCT_ShapeProperties.Create(FOwner);
      Result := FSpPr;
    end;
    $000003B9: begin
      if FStyle = Nil then 
        FStyle := TCT_ShapeStyle.Create(FOwner);
      Result := FStyle;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Picture.Write(AWriter: TXpgWriteXML);
begin
  if (FNvPicPr <> Nil) and FNvPicPr.Assigned then 
    if xaElements in FNvPicPr.Assigneds then
    begin
      AWriter.BeginTag('xdr:nvPicPr');
      FNvPicPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:nvPicPr')
  else 
    AWriter.SimpleTag('xdr:nvPicPr');
  if (FBlipFill <> Nil) and FBlipFill.Assigned then 
  begin
    FBlipFill.WriteAttributes(AWriter);
    if xaElements in FBlipFill.Assigneds then
    begin
      AWriter.BeginTag('xdr:blipFill');
      FBlipFill.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:blipFill');
  end
  else 
    AWriter.SimpleTag('xdr:blipFill');
  if (FSpPr <> Nil) and FSpPr.Assigned then 
  begin
    FSpPr.WriteAttributes(AWriter);
    if xaElements in FSpPr.Assigneds then
    begin
      AWriter.BeginTag('xdr:spPr');
      FSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:spPr');
  end
  else 
    AWriter.SimpleTag('xdr:spPr');
  if (FStyle <> Nil) and FStyle.Assigned then 
    if xaElements in FStyle.Assigneds then
    begin
      AWriter.BeginTag('xdr:style');
      FStyle.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:style');
end;

procedure TCT_Picture.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FMacro <> '' then 
    AWriter.AddAttribute('macro',FMacro);
  if FFPublished <> False then 
    AWriter.AddAttribute('fPublished',XmlBoolToStr(FFPublished));
end;

procedure TCT_Picture.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000212: FMacro := AAttributes.Values[i];
      $00000406: FFPublished := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Picture.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 2;
  FFPublished := False;
end;

destructor TCT_Picture.Destroy;
begin
  if FNvPicPr <> Nil then 
    FNvPicPr.Free;
  if FBlipFill <> Nil then
    FBlipFill.Free;
  if FSpPr <> Nil then
    FSpPr.Free;
  if FStyle <> Nil then
    FStyle.Free;
end;

function TCT_Picture.GetImage: TXc12GraphicImage;
begin
  Result := FBlipFill.Blip.Image;
end;

procedure TCT_Picture.Clear;
begin
  FAssigneds := [];
  if FNvPicPr <> Nil then 
    FreeAndNil(FNvPicPr);
  if FBlipFill <> Nil then 
    FreeAndNil(FBlipFill);
  if FSpPr <> Nil then 
    FreeAndNil(FSpPr);
  if FStyle <> Nil then 
    FreeAndNil(FStyle);
  FMacro := '';
  FFPublished := False;
end;

procedure TCT_Picture.Assign(AItem: TCT_Picture);
begin
end;

procedure TCT_Picture.CopyTo(AItem: TCT_Picture);
begin
end;

function  TCT_Picture.Create_NvPicPr: TCT_PictureNonVisual;
begin
  Result := TCT_PictureNonVisual.Create(FOwner);
  FNvPicPr := Result;
end;

function  TCT_Picture.Create_BlipFill: TCT_BlipFillProperties;
begin
  Result := TCT_BlipFillProperties.Create(FOwner);
  FBlipFill := Result;
end;

function  TCT_Picture.Create_SpPr: TCT_ShapeProperties;
begin
  Result := TCT_ShapeProperties.Create(FOwner);
  FSpPr := Result;
end;

function  TCT_Picture.Create_Style: TCT_ShapeStyle;
begin
  Result := TCT_ShapeStyle.Create(FOwner);
  FStyle := Result;
end;

{ TCT_PictureXpgList }

function  TCT_PictureXpgList.GetItems(Index: integer): TCT_Picture;
begin
  Result := TCT_Picture(inherited Items[Index]);
end;

function  TCT_PictureXpgList.Add: TCT_Picture;
begin
  Result := TCT_Picture.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PictureXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PictureXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].Assigneds then
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PictureXpgList.Assign(AItem: TCT_PictureXpgList);
begin
end;

procedure TCT_PictureXpgList.CopyTo(AItem: TCT_PictureXpgList);
begin
end;

{ TCT_GroupShape }

function  TCT_GroupShape.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FNvGrpSpPr <> Nil then 
    Inc(ElemsAssigned,FNvGrpSpPr.CheckAssigned);
  if FGrpSpPr <> Nil then 
    Inc(ElemsAssigned,FGrpSpPr.CheckAssigned);
  if FSpXpgList <> Nil then 
    Inc(ElemsAssigned,FSpXpgList.CheckAssigned);
  if FGrpSpXpgList <> Nil then 
    Inc(ElemsAssigned,FGrpSpXpgList.CheckAssigned);
  if FGraphicFrameXpgList <> Nil then 
    Inc(ElemsAssigned,FGraphicFrameXpgList.CheckAssigned);
  if FCxnSpXpgList <> Nil then 
    Inc(ElemsAssigned,FCxnSpXpgList.CheckAssigned);
  if FPicXpgList <> Nil then 
    Inc(ElemsAssigned,FPicXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_GroupShape.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashB of
    $74A4FF56: begin
      if FNvGrpSpPr = Nil then 
        FNvGrpSpPr := TCT_GroupShapeNonVisual.Create(FOwner);
      Result := FNvGrpSpPr;
    end;
    $DC890916: begin
      if FGrpSpPr = Nil then 
        FGrpSpPr := TCT_GroupShapeProperties.Create(FOwner);
      Result := FGrpSpPr;
    end;
    $FF83FE11: begin
      if FSpXpgList = Nil then 
        FSpXpgList := TCT_ShapeXpgList.Create(FOwner);
      Result := FSpXpgList.Add;
    end;
    $A53E4C4C: begin
      if FGrpSpXpgList = Nil then 
        FGrpSpXpgList := TCT_GroupShapeXpgList.Create(FOwner);
      Result := FGrpSpXpgList.Add;
    end;
    $AF637773: begin
      if FGraphicFrameXpgList = Nil then 
        FGraphicFrameXpgList := TCT_GraphicalObjectFrameXpgList.Create(FOwner);
      Result := FGraphicFrameXpgList.Add;
    end;
    $F3841F94: begin
      if FCxnSpXpgList = Nil then 
        FCxnSpXpgList := TCT_ConnectorXpgList.Create(FOwner);
      Result := FCxnSpXpgList.Add;
    end;
    $9C4F1918: begin
      if FPicXpgList = Nil then 
        FPicXpgList := TCT_PictureXpgList.Create(FOwner);
      Result := FPicXpgList.Add;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GroupShape.Write(AWriter: TXpgWriteXML);
begin
  if (FNvGrpSpPr <> Nil) and FNvGrpSpPr.Assigned then 
    if xaElements in FNvGrpSpPr.Assigneds then
    begin
      AWriter.BeginTag('xdr:nvGrpSpPr');
      FNvGrpSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:nvGrpSpPr')
  else 
    AWriter.SimpleTag('xdr:nvGrpSpPr');

  if (FGrpSpPr <> Nil) and FGrpSpPr.Assigned then
  begin
    FGrpSpPr.WriteAttributes(AWriter);
    if xaElements in FGrpSpPr.Assigneds then
    begin
      AWriter.BeginTag('xdr:grpSpPr');
      FGrpSpPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('xdr:grpSpPr');
  end
  else
    AWriter.SimpleTag('xdr:grpSpPr');

  if FGrpSpXpgList <> Nil then
    FGrpSpXpgList.Write(AWriter,'xdr:grpSp');
  if FSpXpgList <> Nil then
    FSpXpgList.Write(AWriter,'xdr:sp');
  if FGraphicFrameXpgList <> Nil then
    FGraphicFrameXpgList.Write(AWriter,'xdr:graphicFrame');
  if FCxnSpXpgList <> Nil then 
    FCxnSpXpgList.Write(AWriter,'xdr:cxnSp');
  if FPicXpgList <> Nil then 
    FPicXpgList.Write(AWriter,'xdr:pic');
end;

constructor TCT_GroupShape.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 7;
  FAttributeCount := 0;
end;

destructor TCT_GroupShape.Destroy;
begin
  if FNvGrpSpPr <> Nil then 
    FNvGrpSpPr.Free;
  if FGrpSpPr <> Nil then 
    FGrpSpPr.Free;
  if FSpXpgList <> Nil then 
    FSpXpgList.Free;
  if FGrpSpXpgList <> Nil then 
    FGrpSpXpgList.Free;
  if FGraphicFrameXpgList <> Nil then 
    FGraphicFrameXpgList.Free;
  if FCxnSpXpgList <> Nil then 
    FCxnSpXpgList.Free;
  if FPicXpgList <> Nil then 
    FPicXpgList.Free;
end;

procedure TCT_GroupShape.Clear;
begin
  FAssigneds := [];
  if FNvGrpSpPr <> Nil then 
    FreeAndNil(FNvGrpSpPr);
  if FGrpSpPr <> Nil then 
    FreeAndNil(FGrpSpPr);
  if FSpXpgList <> Nil then 
    FreeAndNil(FSpXpgList);
  if FGrpSpXpgList <> Nil then 
    FreeAndNil(FGrpSpXpgList);
  if FGraphicFrameXpgList <> Nil then 
    FreeAndNil(FGraphicFrameXpgList);
  if FCxnSpXpgList <> Nil then 
    FreeAndNil(FCxnSpXpgList);
  if FPicXpgList <> Nil then 
    FreeAndNil(FPicXpgList);
end;

procedure TCT_GroupShape.Assign(AItem: TCT_GroupShape);
begin
end;

procedure TCT_GroupShape.CopyTo(AItem: TCT_GroupShape);
begin
end;

function  TCT_GroupShape.Create_NvGrpSpPr: TCT_GroupShapeNonVisual;
begin
  Result := TCT_GroupShapeNonVisual.Create(FOwner);
  FNvGrpSpPr := Result;
end;

function  TCT_GroupShape.Create_GrpSpPr: TCT_GroupShapeProperties;
begin
  Result := TCT_GroupShapeProperties.Create(FOwner);
  FGrpSpPr := Result;
end;

function  TCT_GroupShape.Create_SpXpgList: TCT_ShapeXpgList;
begin
  Result := TCT_ShapeXpgList.Create(FOwner);
  FSpXpgList := Result;
end;

function  TCT_GroupShape.Create_GrpSpXpgList: TCT_GroupShapeXpgList;
begin
  Result := TCT_GroupShapeXpgList.Create(FOwner);
  FGrpSpXpgList := Result;
end;

function  TCT_GroupShape.Create_GraphicFrameXpgList: TCT_GraphicalObjectFrameXpgList;
begin
  Result := TCT_GraphicalObjectFrameXpgList.Create(FOwner);
  FGraphicFrameXpgList := Result;
end;

function  TCT_GroupShape.Create_CxnSpXpgList: TCT_ConnectorXpgList;
begin
  Result := TCT_ConnectorXpgList.Create(FOwner);
  FCxnSpXpgList := Result;
end;

function  TCT_GroupShape.Create_PicXpgList: TCT_PictureXpgList;
begin
  Result := TCT_PictureXpgList.Create(FOwner);
  FPicXpgList := Result;
end;

{ TCT_GroupShapeXpgList }

function  TCT_GroupShapeXpgList.GetItems(Index: integer): TCT_GroupShape;
begin
  Result := TCT_GroupShape(inherited Items[Index]);
end;

function  TCT_GroupShapeXpgList.Add: TCT_GroupShape;
begin
  Result := TCT_GroupShape.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_GroupShapeXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_GroupShapeXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_GroupShapeXpgList.Assign(AItem: TCT_GroupShapeXpgList);
begin
end;

procedure TCT_GroupShapeXpgList.CopyTo(AItem: TCT_GroupShapeXpgList);
begin
end;

{ TEG_ObjectChoices }

function  TEG_ObjectChoices.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FSp <> Nil then 
    Inc(ElemsAssigned,FSp.CheckAssigned);
  if FGrpSp <> Nil then 
    Inc(ElemsAssigned,FGrpSp.CheckAssigned);
  if FGraphicFrame <> Nil then 
    Inc(ElemsAssigned,FGraphicFrame.CheckAssigned);
  if FCxnSp <> Nil then 
    Inc(ElemsAssigned,FCxnSp.CheckAssigned);
  if FPic <> Nil then 
    Inc(ElemsAssigned,FPic.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_ObjectChoices.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashB of
    $FF83FE11: begin
      if FSp = Nil then 
        FSp := TCT_Shape.Create(FOwner);
      Result := FSp;
    end;
    $A53E4C4C: begin
      if FGrpSp = Nil then 
        FGrpSp := TCT_GroupShape.Create(FOwner);
      Result := FGrpSp;
    end;
    $AF637773: begin
      if FGraphicFrame = Nil then 
        FGraphicFrame := TCT_GraphicalObjectFrame.Create(FOwner);
      Result := FGraphicFrame;
    end;
    $F3841F94: begin
      if FCxnSp = Nil then 
        FCxnSp := TCT_Connector.Create(FOwner);
      Result := FCxnSp;
    end;
    $9C4F1918: begin
      if FPic = Nil then
        FPic := TCT_Picture.Create(FOwner);
      Result := FPic;
    end;
    else
    begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

function TEG_ObjectChoices.IsChart: boolean;
begin
  Result := (FGraphicFrame <> Nil) and (FGraphicFrame.Graphic <> Nil) and (FGraphicFrame.Graphic.GraphicData <> Nil) and (FGraphicFrame.Graphic.GraphicData.Chart <> Nil);
end;

function TEG_ObjectChoices.IsPicture: boolean;
begin
  Result := FPic <> Nil;
end;

procedure TEG_ObjectChoices.Write(AWriter: TXpgWriteXML);
begin
  if (FSp <> Nil) and FSp.Assigned then 
  begin
    FSp.WriteAttributes(AWriter);
    if xaElements in FSp.Assigneds then
    begin
      AWriter.BeginTag('xdr:sp');
      FSp.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:sp');
  end;
  if (FGrpSp <> Nil) and FGrpSp.Assigned then 
    if xaElements in FGrpSp.Assigneds then
    begin
      AWriter.BeginTag('xdr:grpSp');
      FGrpSp.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:grpSp');
  if (FGraphicFrame <> Nil) and FGraphicFrame.Assigned then 
  begin
    FGraphicFrame.WriteAttributes(AWriter);
    if xaElements in FGraphicFrame.Assigneds then
    begin
      AWriter.BeginTag('xdr:graphicFrame');
      FGraphicFrame.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:graphicFrame');
  end;
  if (FCxnSp <> Nil) and FCxnSp.Assigned then 
  begin
    FCxnSp.WriteAttributes(AWriter);
    if xaElements in FCxnSp.Assigneds then
    begin
      AWriter.BeginTag('xdr:cxnSp');
      FCxnSp.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:cxnSp');
  end;
  if (FPic <> Nil) and FPic.Assigned then 
  begin
    FPic.WriteAttributes(AWriter);
    if xaElements in FPic.Assigneds then
    begin
      AWriter.BeginTag('xdr:pic');
      FPic.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:pic');
  end;
end;

constructor TEG_ObjectChoices.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 0;
end;

destructor TEG_ObjectChoices.Destroy;
begin
  if FSp <> Nil then 
    FSp.Free;
  if FGrpSp <> Nil then 
    FGrpSp.Free;
  if FGraphicFrame <> Nil then 
    FGraphicFrame.Free;
  if FCxnSp <> Nil then 
    FCxnSp.Free;
  if FPic <> Nil then 
    FPic.Free;
end;

procedure TEG_ObjectChoices.Clear;
begin
  FAssigneds := [];
  if FSp <> Nil then 
    FreeAndNil(FSp);
  if FGrpSp <> Nil then 
    FreeAndNil(FGrpSp);
  if FGraphicFrame <> Nil then 
    FreeAndNil(FGraphicFrame);
  if FCxnSp <> Nil then 
    FreeAndNil(FCxnSp);
  if FPic <> Nil then 
    FreeAndNil(FPic);
end;

procedure TEG_ObjectChoices.Assign(AItem: TEG_ObjectChoices);
begin
end;

procedure TEG_ObjectChoices.CopyTo(AItem: TEG_ObjectChoices);
begin
end;

function  TEG_ObjectChoices.Create_Sp: TCT_Shape;
begin
  Result := TCT_Shape.Create(FOwner);
  FSp := Result;
end;

function  TEG_ObjectChoices.Create_GrpSp: TCT_GroupShape;
begin
  Result := TCT_GroupShape.Create(FOwner);
  FGrpSp := Result;
end;

function  TEG_ObjectChoices.Create_GraphicFrame: TCT_GraphicalObjectFrame;
begin
  Result := TCT_GraphicalObjectFrame.Create(FOwner);
  FGraphicFrame := Result;
end;

function  TEG_ObjectChoices.Create_CxnSp: TCT_Connector;
begin
  Result := TCT_Connector.Create(FOwner);
  FCxnSp := Result;
end;

function  TEG_ObjectChoices.Create_Pic: TCT_Picture;
begin
  Result := TCT_Picture.Create(FOwner);
  FPic := Result;
end;

{ TCT_AnchorClientData }

function  TCT_AnchorClientData.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FFLocksWithSheet <> True then 
    Inc(AttrsAssigned);
  if FFPrintsWithSheet <> True then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_AnchorClientData.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_AnchorClientData.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FFLocksWithSheet <> True then 
    AWriter.AddAttribute('fLocksWithSheet',XmlBoolToStr(FFLocksWithSheet));
  if FFPrintsWithSheet <> True then 
    AWriter.AddAttribute('fPrintsWithSheet',XmlBoolToStr(FFPrintsWithSheet));
end;

procedure TCT_AnchorClientData.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000005F7: FFLocksWithSheet := XmlStrToBoolDef(AAttributes.Values[i],True);
      $0000067B: FFPrintsWithSheet := XmlStrToBoolDef(AAttributes.Values[i],True);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_AnchorClientData.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FFLocksWithSheet := True;
  FFPrintsWithSheet := True;
end;

destructor TCT_AnchorClientData.Destroy;
begin
end;

procedure TCT_AnchorClientData.Clear;
begin
  FAssigneds := [];
  FFLocksWithSheet := True;
  FFPrintsWithSheet := True;
end;

procedure TCT_AnchorClientData.Assign(AItem: TCT_AnchorClientData);
begin
end;

procedure TCT_AnchorClientData.CopyTo(AItem: TCT_AnchorClientData);
begin
end;

{ TCT_TwoCellAnchor }

function  TCT_TwoCellAnchor.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FEditAs <> steaTwoCell then
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FFrom.CheckAssigned);
  Inc(ElemsAssigned,FTo.CheckAssigned);
  if FEG_ObjectChoices <> Nil then
    Inc(ElemsAssigned,FEG_ObjectChoices.CheckAssigned);
  if FClientData <> Nil then
    Inc(ElemsAssigned,FClientData.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_TwoCellAnchor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashB of
    $13D59EEE: Result := FFrom;
    $D5D0631F: Result := FTo;
    $515DD6BF: Result := FClientData;
    else
    begin
      Result := FEG_ObjectChoices.HandleElement(AReader);
      if Result = Nil then
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

function TCT_TwoCellAnchor.MakeChart: TXPGDocXLSXChart;
var
  GOF: TCT_GraphicalObjectFrame;
  GO : TCT_GraphicalObject;
  GOD: TCT_GraphicalObjectData;
begin
  FClientData.Free;
  FClientData := Nil;

  GOF := FEG_ObjectChoices.Create_GraphicFrame;
  GOF.Create_NvGraphicFramePr;
  GOF.Create_Xfrm;
  GO :=GOF.Create_A_Graphic;
  GOD := GO.Create_A_GraphicData;
  GOD.Uri := 'http://schemas.openxmlformats.org/drawingml/2006/chart';

  Result := GOD.Create_Chart;
end;

procedure TCT_TwoCellAnchor.Write(AWriter: TXpgWriteXML);
begin
  if FFrom.Assigned then
    if xaElements in FFrom.Assigneds then
    begin
      AWriter.BeginTag('xdr:from');
      FFrom.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('xdr:from')
  else
    AWriter.SimpleTag('xdr:from');
  if FTo.Assigned then
    if xaElements in FTo.Assigneds then
    begin
      AWriter.BeginTag('xdr:to');
      FTo.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('xdr:to')
  else
    AWriter.SimpleTag('xdr:to');

  if FEG_ObjectChoices <> Nil then
    FEG_ObjectChoices.Write(AWriter);

  if (FClientData <> Nil) and FClientData.Assigned then
    FClientData.WriteAttributes(AWriter);
  AWriter.SimpleTag('xdr:clientData');
end;

procedure TCT_TwoCellAnchor.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FEditAs <> steaTwoCell then
    AWriter.AddAttribute('editAs',StrTST_EditAs[Integer(FEditAs)]);
end;

procedure TCT_TwoCellAnchor.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'editAs' then
    FEditAs := TST_EditAs(StrToEnum('stea' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TwoCellAnchor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 1;
  FEG_ObjectChoices := TEG_ObjectChoices.Create(FOwner);
  FEditAs := steaTwoCell;

  FFrom := TCT_Marker_Pos.Create(FOwner);
  FTo := TCT_Marker_Pos.Create(FOwner);
  FClientData := TCT_AnchorClientData.Create(FOwner);
end;

destructor TCT_TwoCellAnchor.Destroy;
begin
  FFrom.Free;
  FTo.Free;
  FEG_ObjectChoices.Free;
  FClientData.Free;
end;

procedure TCT_TwoCellAnchor.Clear;
begin
  FAssigneds := [];
  FFrom.Clear;
  FTo.Clear;
  FEG_ObjectChoices.Clear;
  FClientData.Clear;
  FEditAs := steaTwoCell;
end;

procedure TCT_TwoCellAnchor.Assign(AItem: TCT_TwoCellAnchor);
begin
end;

procedure TCT_TwoCellAnchor.CopyTo(AItem: TCT_TwoCellAnchor);
begin
end;

{ TCT_TwoCellAnchorXpgList }

function  TCT_TwoCellAnchorXpgList.GetItems(Index: integer): TCT_TwoCellAnchor;
begin
  Result := TCT_TwoCellAnchor(inherited Items[Index]);
end;

function  TCT_TwoCellAnchorXpgList.Add: TCT_TwoCellAnchor;
begin
  Result := TCT_TwoCellAnchor.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_TwoCellAnchorXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_TwoCellAnchorXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].Assigneds then
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

function TCT_TwoCellAnchorXpgList.Add(ACol1, ARow1, ACol2, ARow2: integer): TCT_TwoCellAnchor;
begin
  Result := Add;

  Result.From.Col := ACol1;
  Result.From.ColOff := 0;
  Result.From.Row := ARow1;
  Result.From.RowOff := 0;

  Result.To_.Col := ACol2;
  Result.To_.ColOff := 0;
  Result.To_.Row := ARow2;
  Result.To_.RowOff := 0;
end;

procedure TCT_TwoCellAnchorXpgList.Assign(AItem: TCT_TwoCellAnchorXpgList);
begin
end;

procedure TCT_TwoCellAnchorXpgList.CopyTo(AItem: TCT_TwoCellAnchorXpgList);
begin
end;

{ TCT_OneCellAnchor }

function  TCT_OneCellAnchor.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FFrom <> Nil then 
    Inc(ElemsAssigned,FFrom.CheckAssigned);
  if FExt <> Nil then 
    Inc(ElemsAssigned,FExt.CheckAssigned);
  Inc(ElemsAssigned,FEG_ObjectChoices.CheckAssigned);
  if FClientData <> Nil then 
    Inc(ElemsAssigned,FClientData.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_OneCellAnchor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashB of
    $13D59EEE: begin
      if FFrom = Nil then 
        FFrom := TCT_Marker_Pos.Create(FOwner);
      Result := FFrom;
    end;
    $6018D3C3: begin
      if FExt = Nil then 
        FExt := TCT_PositiveSize2D.Create(FOwner);
      Result := FExt;
    end;
    $515DD6BF: begin
      if FClientData = Nil then 
        FClientData := TCT_AnchorClientData.Create(FOwner);
      Result := FClientData;
    end;
    else 
    begin
      Result := FEG_ObjectChoices.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_OneCellAnchor.Write(AWriter: TXpgWriteXML);
begin
  if (FFrom <> Nil) and FFrom.Assigned then 
    if xaElements in FFrom.Assigneds then
    begin
      AWriter.BeginTag('xdr:from');
      FFrom.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('xdr:from')
  else 
    AWriter.SimpleTag('xdr:from');
  if (FExt <> Nil) and FExt.Assigned then 
  begin
    FExt.WriteAttributes(AWriter);
    AWriter.SimpleTag('xdr:ext');
  end
  else 
    AWriter.SimpleTag('xdr:ext');
  FEG_ObjectChoices.Write(AWriter);
  if (FClientData <> Nil) and FClientData.Assigned then 
  begin
    FClientData.WriteAttributes(AWriter);
    AWriter.SimpleTag('xdr:clientData');
  end
  else 
    AWriter.SimpleTag('xdr:clientData');
end;

constructor TCT_OneCellAnchor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FEG_ObjectChoices := TEG_ObjectChoices.Create(FOwner);
end;

destructor TCT_OneCellAnchor.Destroy;
begin
  if FFrom <> Nil then 
    FFrom.Free;
  if FExt <> Nil then 
    FExt.Free;
  FEG_ObjectChoices.Free;
  if FClientData <> Nil then 
    FClientData.Free;
end;

procedure TCT_OneCellAnchor.Clear;
begin
  FAssigneds := [];
  if FFrom <> Nil then 
    FreeAndNil(FFrom);
  if FExt <> Nil then 
    FreeAndNil(FExt);
  FEG_ObjectChoices.Clear;
  if FClientData <> Nil then 
    FreeAndNil(FClientData);
end;

procedure TCT_OneCellAnchor.Assign(AItem: TCT_OneCellAnchor);
begin
end;

procedure TCT_OneCellAnchor.CopyTo(AItem: TCT_OneCellAnchor);
begin
end;

function  TCT_OneCellAnchor.Create_From: TCT_Marker_Pos;
begin
  Result := TCT_Marker_Pos.Create(FOwner);
  FFrom := Result;
end;

function  TCT_OneCellAnchor.Create_Ext: TCT_PositiveSize2D;
begin
  Result := TCT_PositiveSize2D.Create(FOwner);
  FExt := Result;
end;

function  TCT_OneCellAnchor.Create_ClientData: TCT_AnchorClientData;
begin
  Result := TCT_AnchorClientData.Create(FOwner);
  FClientData := Result;
end;

{ TCT_OneCellAnchorXpgList }

function  TCT_OneCellAnchorXpgList.GetItems(Index: integer): TCT_OneCellAnchor;
begin
  Result := TCT_OneCellAnchor(inherited Items[Index]);
end;

function  TCT_OneCellAnchorXpgList.Add: TCT_OneCellAnchor;
begin
  Result := TCT_OneCellAnchor.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_OneCellAnchorXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_OneCellAnchorXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag(AName);
end;

procedure TCT_OneCellAnchorXpgList.Assign(AItem: TCT_OneCellAnchorXpgList);
begin
end;

procedure TCT_OneCellAnchorXpgList.CopyTo(AItem: TCT_OneCellAnchorXpgList);
begin
end;

{ TCT_AbsoluteAnchor }

function  TCT_AbsoluteAnchor.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FPos <> Nil then 
    Inc(ElemsAssigned,FPos.CheckAssigned);
  if FExt <> Nil then 
    Inc(ElemsAssigned,FExt.CheckAssigned);
  Inc(ElemsAssigned,FEG_ObjectChoices.CheckAssigned);
  if FClientData <> Nil then 
    Inc(ElemsAssigned,FClientData.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_AbsoluteAnchor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashB of
    $1CB4277E: begin
      if FPos = Nil then 
        FPos := TCT_Point2D.Create(FOwner);
      Result := FPos;
    end;
    $6018D3C3: begin
      if FExt = Nil then 
        FExt := TCT_PositiveSize2D.Create(FOwner);
      Result := FExt;
    end;
    $515DD6BF: begin
      if FClientData = Nil then 
        FClientData := TCT_AnchorClientData.Create(FOwner);
      Result := FClientData;
    end;
    else 
    begin
      Result := FEG_ObjectChoices.HandleElement(AReader);
      if Result = Nil then 
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_AbsoluteAnchor.Write(AWriter: TXpgWriteXML);
begin
  if (FPos <> Nil) and FPos.Assigned then 
  begin
    FPos.WriteAttributes(AWriter);
    AWriter.SimpleTag('xdr:pos');
  end
  else 
    AWriter.SimpleTag('xdr:pos');
  if (FExt <> Nil) and FExt.Assigned then 
  begin
    FExt.WriteAttributes(AWriter);
    AWriter.SimpleTag('xdr:ext');
  end
  else 
    AWriter.SimpleTag('xdr:ext');
  FEG_ObjectChoices.Write(AWriter);
  if (FClientData <> Nil) and FClientData.Assigned then 
  begin
    FClientData.WriteAttributes(AWriter);
    AWriter.SimpleTag('xdr:clientData');
  end
  else 
    AWriter.SimpleTag('xdr:clientData');
end;

constructor TCT_AbsoluteAnchor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FEG_ObjectChoices := TEG_ObjectChoices.Create(FOwner);
end;

destructor TCT_AbsoluteAnchor.Destroy;
begin
  if FPos <> Nil then 
    FPos.Free;
  if FExt <> Nil then 
    FExt.Free;
  FEG_ObjectChoices.Free;
  if FClientData <> Nil then 
    FClientData.Free;
end;

procedure TCT_AbsoluteAnchor.Clear;
begin
  FAssigneds := [];
  if FPos <> Nil then 
    FreeAndNil(FPos);
  if FExt <> Nil then 
    FreeAndNil(FExt);
  FEG_ObjectChoices.Clear;
  if FClientData <> Nil then 
    FreeAndNil(FClientData);
end;

procedure TCT_AbsoluteAnchor.Assign(AItem: TCT_AbsoluteAnchor);
begin
end;

procedure TCT_AbsoluteAnchor.CopyTo(AItem: TCT_AbsoluteAnchor);
begin
end;

function  TCT_AbsoluteAnchor.Create_Pos: TCT_Point2D;
begin
  Result := TCT_Point2D.Create(FOwner);
  FPos := Result;
end;

function  TCT_AbsoluteAnchor.Create_Ext: TCT_PositiveSize2D;
begin
  Result := TCT_PositiveSize2D.Create(FOwner);
  FExt := Result;
end;

function  TCT_AbsoluteAnchor.Create_ClientData: TCT_AnchorClientData;
begin
  Result := TCT_AnchorClientData.Create(FOwner);
  FClientData := Result;
end;

{ TCT_AbsoluteAnchorXpgList }

function  TCT_AbsoluteAnchorXpgList.GetItems(Index: integer): TCT_AbsoluteAnchor;
begin
  Result := TCT_AbsoluteAnchor(inherited Items[Index]);
end;

function  TCT_AbsoluteAnchorXpgList.Add: TCT_AbsoluteAnchor;
begin
  Result := TCT_AbsoluteAnchor.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_AbsoluteAnchorXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_AbsoluteAnchorXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_AbsoluteAnchorXpgList.Assign(AItem: TCT_AbsoluteAnchorXpgList);
begin
end;

procedure TCT_AbsoluteAnchorXpgList.CopyTo(AItem: TCT_AbsoluteAnchorXpgList);
begin
end;

{ TEG_Anchor }

function  TEG_Anchor.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FTwoCellAnchorXpgList <> Nil then
    Inc(ElemsAssigned,FTwoCellAnchorXpgList.CheckAssigned);
  if FOneCellAnchorXpgList <> Nil then
    Inc(ElemsAssigned,FOneCellAnchorXpgList.CheckAssigned);
  if FAbsoluteAnchorXpgList <> Nil then
    Inc(ElemsAssigned,FAbsoluteAnchorXpgList.CheckAssigned);
  if FRelSizeAnchorXpgList <> Nil then
    Inc(ElemsAssigned,FRelSizeAnchorXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_Anchor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  case AReader.QNameHashA of
    $000006BD: begin
      if FTwoCellAnchorXpgList = Nil then
        FTwoCellAnchorXpgList := TCT_TwoCellAnchorXpgList.Create(FOwner);
      Result := FTwoCellAnchorXpgList.Add;
    end;
    $000006A5: begin
      if FOneCellAnchorXpgList = Nil then 
        FOneCellAnchorXpgList := TCT_OneCellAnchorXpgList.Create(FOwner);
      Result := FOneCellAnchorXpgList.Add;
    end;
    $00000742: begin
      if FAbsoluteAnchorXpgList = Nil then
        FAbsoluteAnchorXpgList := TCT_AbsoluteAnchorXpgList.Create(FOwner);
      Result := FAbsoluteAnchorXpgList.Add;
    end;
    $000006AC: begin
      if FRelSizeAnchorXpgList = Nil then
        FRelSizeAnchorXpgList := TCT_RelSizeAnchorXpgList.Create(FOwner);
      Result := FRelSizeAnchorXpgList.Add;
    end
    else begin
      Result := Nil;
      Exit;
    end
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

function TEG_Anchor.HasData: boolean;
begin
  Result := (FTwoCellAnchorXpgList.Count > 0) or (FOneCellAnchorXpgList.Count > 0) or (FAbsoluteAnchorXpgList.Count > 0);
end;

procedure TEG_Anchor.Write(AWriter: TXpgWriteXML);
begin
  if FTwoCellAnchorXpgList <> Nil then
    FTwoCellAnchorXpgList.Write(AWriter,'xdr:twoCellAnchor');
  if FOneCellAnchorXpgList <> Nil then 
    FOneCellAnchorXpgList.Write(AWriter,'xdr:oneCellAnchor');
  if FAbsoluteAnchorXpgList <> Nil then
    FAbsoluteAnchorXpgList.Write(AWriter,'xdr:absoluteAnchor');
  if FRelSizeAnchorXpgList <> Nil then
    FRelSizeAnchorXpgList.Write(AWriter,FOwner.NS + ':relSizeAnchor');
end;

constructor TEG_Anchor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;

  FTwoCellAnchorXpgList := TCT_TwoCellAnchorXpgList.Create(FOwner);
  FOneCellAnchorXpgList := TCT_OneCellAnchorXpgList.Create(FOwner);
  FAbsoluteAnchorXpgList := TCT_AbsoluteAnchorXpgList.Create(FOwner);
  FRelSizeAnchorXpgList := TCT_RelSizeAnchorXpgList.Create(FOwner);
end;

destructor TEG_Anchor.Destroy;
begin
  FTwoCellAnchorXpgList.Free;
  FOneCellAnchorXpgList.Free;
  FAbsoluteAnchorXpgList.Free;
  FRelSizeAnchorXpgList.Free;
end;

procedure TEG_Anchor.Clear;
begin
  FAssigneds := [];
  FTwoCellAnchorXpgList.Clear;
  FOneCellAnchorXpgList.Clear;
  FAbsoluteAnchorXpgList.Clear;
end;

procedure TEG_Anchor.Assign(AItem: TEG_Anchor);
begin
end;

procedure TEG_Anchor.CopyTo(AItem: TEG_Anchor);
begin
end;

{ TCT_Drawing }

function  TCT_Drawing.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_Anchor.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Drawing.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FEG_Anchor.HandleElement(AReader);
  if Result = Nil then
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result = Nil then
    Result := Self
  else
    Result.Assigneds := [xaRead];
end;

function TCT_Drawing.HasData: boolean;
begin
  Result := FEG_Anchor.HasData;
end;

procedure TCT_Drawing.Write(AWriter: TXpgWriteXML);
begin
  FEG_Anchor.Write(AWriter);
end;

constructor TCT_Drawing.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FEG_Anchor := TEG_Anchor.Create(FOwner);
end;

destructor TCT_Drawing.Destroy;
begin
  FEG_Anchor.Free;
end;

procedure TCT_Drawing.Clear;
begin
  FAssigneds := [];
  FEG_Anchor.Clear;
end;

procedure TCT_Drawing.Assign(AItem: TCT_Drawing);
begin
end;

procedure TCT_Drawing.CopyTo(AItem: TCT_Drawing);
begin
end;

{ T__ROOT__ }

function  T__ROOT__.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FWsDr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  T__ROOT__.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  i: integer;
begin
  for i := 0 to AReader.Attributes.Count - 1 do
    FRootAttributes.Add(AReader.Attributes.AsXmlText2(i));
  Result := Self;
  case AReader.QNameHashA of
    $00000328: Result := FWsDr;
    else begin
      if AReader.QName = 'c:userShapes' then
        Result := FWsDr
      else
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end;
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure T__ROOT__.Write(AWriter: TXpgWriteXML);
begin
  AWriter.Attributes := FRootAttributes.Text;
  if FWsDr.Assigned then
    if xaElements in FWsDr.Assigneds then
    begin
      AWriter.BeginTag('xdr:wsDr');
      FWsDr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('xdr:wsDr')
  else
    AWriter.SimpleTag('xdr:wsDr');
end;

procedure T__ROOT__.WriteUserShapes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('xmlns:c',OOXML_URI_OFFICEDOC_CHART);
  AWriter.AddAttribute('xmlns:a',OOXML_URI_OFFICEDOC_DRAWING);

  AWriter.BeginTag('c:userShapes');
  FWsDr.Write(AWriter);
  AWriter.EndTag;
end;

constructor T__ROOT__.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FRootAttributes := TStringXpgList.Create;
  FElementCount := 3;
  FAttributeCount := 0;
  FWsDr := TCT_Drawing.Create(FOwner);
end;

destructor T__ROOT__.Destroy;
begin
  FRootAttributes.Free;
  FWsDr.Free;
end;

procedure T__ROOT__.Clear;
begin
  FRootAttributes.Clear;
  FAssigneds := [];
  FWsDr.Clear;
end;

procedure T__ROOT__.Assign(AItem: T__ROOT__);
begin
end;

procedure T__ROOT__.CopyTo(AItem: T__ROOT__);
begin
end;

{ TXPGDocument }

function TXPGDocXLSXDrawing.GetErrors: TXpgPErrors;
begin
  Result := FReader.Errors;
end;

function  TXPGDocXLSXDrawing.GetWsDr: TCT_Drawing;
begin
  Result := FRoot.WsDr;
end;

constructor TXPGDocXLSXDrawing.Create(AGrManager: TXc12GraphicManager);
begin
  FGrManager := AGrManager;
  FRoot := T__ROOT__.Create(Self);
  FErrors := TXpgPErrors.Create;
  FReader := TXPGReader.Create(FErrors,FRoot);
  FWriter := TXpgWriteXML.Create;
end;

destructor TXPGDocXLSXDrawing.Destroy;
begin
  FRoot.Free;
  FReader.Free;
  FWriter.Free;
  FErrors.Free;
  inherited Destroy;
end;

procedure TXPGDocXLSXDrawing.LoadFromFile(AFilename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFilename,fmOpenRead);
  try
    FReader.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXPGDocXLSXDrawing._LoadFromStream(AStream: TStream; AUserShapes: boolean = False);
begin
  if AUserShapes then
    FNS := 'cdr'
  else
    FNS := 'xdr';

  FReader.LoadFromStream(AStream);
end;

procedure TXPGDocXLSXDrawing.SaveToFile(AFilename: AxUCString; AClassToWrite: TClass);
begin
  FRoot.FCurrWriteClass := AClassToWrite;
  FWriter.SaveToFile(AFilename);
  FRoot.CheckAssigned;
  FRoot.Write(FWriter);
end;

procedure TXPGDocXLSXDrawing.SaveToStream(AStream: TStream; AUserShapes: boolean = False);
begin
  FWriter.SaveToStream(AStream);
  FRoot.CheckAssigned;
  if AUserShapes then begin
    FNS := 'cdr';

    FRoot.WriteUserShapes(FWriter);
  end
  else begin
    FNS := 'xdr';

    FRoot.Write(FWriter);
  end;
end;

{ TCT_RelSizeAnchor }

procedure TCT_RelSizeAnchor.AddTextBox(AText: AxUCString);
var
  P: TCT_TextParagraph;
  R: TEG_TextRun;
begin
  P := AddTextBox;

  R := P.TextRuns.Add;
  R.Create_R;
  R.Run.T := AText;
end;

function TCT_RelSizeAnchor.AddTextBox: TCT_TextParagraph;
begin
  FSp.Create_NvSpPr;

  FSp.Create_NvSpPr.Create_CNvPr;
  FSp.NvSpPr.CNvPr.Id := 1;
  FSp.NvSpPr.CNvPr.Name := 'TextBox';

  FSp.NvSpPr.Create_CNvSpPr;
  FSp.NvSpPr.CNvSpPr.TxBox := True;

  FSp.Create_TxBody;
  FSp.TxBody.Create_Paras;
  Result := FSp.TxBody.Paras.Add;
end;

function TCT_RelSizeAnchor.CheckAssigned: integer;
begin
  Result := 3;

  FAssigneds := [xaElements];

  FSp.CheckAssigned;
end;

constructor TCT_RelSizeAnchor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 1;
  FSp := TCT_Shape.Create(FOwner);
  FSp.Create_SpPr;

  FFrom := TCT_XYMarker.Create(FOwner);
  FFrom.X := 0.2;
  FFrom.Y := 0.2;

  FTo := TCT_XYMarker.Create(FOwner);
  FTo.X := 0.8;
  FTo.Y := 0.4;
end;

destructor TCT_RelSizeAnchor.Destroy;
begin
  FSp.Free;

  FFrom.Free;
  FTo.Free;

  inherited;
end;

function TCT_RelSizeAnchor.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  if AReader.QName = 'cdr:from' then
    Result := FFrom
  else if AReader.QName = 'cdr:to' then
    Result := FTo
  else if AReader.QName = 'cdr:sp' then
    Result := FSp
  else
    Result := Self;
end;

procedure TCT_RelSizeAnchor.Write(AWriter: TXpgWriteXML);
begin
  AWriter.BeginTag(FOwner.NS + ':from');
  FFrom.Write(AWriter);
  AWriter.EndTag;

  AWriter.BeginTag(FOwner.NS + ':to');
  FTo.Write(AWriter);
  AWriter.EndTag;

  AWriter.BeginTag(FOwner.NS + ':sp');
  FSp.Write(AWriter);
  AWriter.EndTag;
end;

{ TCT_XYMarker }

function TCT_XYMarker.CheckAssigned: integer;
begin
  Result := 0;
end;

constructor TCT_XYMarker.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
end;

function TCT_XYMarker.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  if AReader.QName = 'cdr:x' then
    FX := XmlStrToFloatDef(Areader.Text,0)
  else if AReader.QName = 'cdr:y' then
    FY := XmlStrToFloatDef(Areader.Text,0);

  Result := Self;
end;

procedure TCT_XYMarker.Write(AWriter: TXpgWriteXML);
begin
  AWriter.Text := XmlFloatToStr(FX);
  AWriter.SimpleTag(FOwner.NS + ':x');

  AWriter.Text := XmlFloatToStr(FY);
  AWriter.SimpleTag(FOwner.NS + ':y');
end;

{ TCT_RelSizeAnchorXpgList }

function TCT_RelSizeAnchorXpgList.Add: TCT_RelSizeAnchor;
begin
  Result := TCT_RelSizeAnchor.Create(FOwner);
  inherited Add(Result);
end;

function TCT_RelSizeAnchorXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

function TCT_RelSizeAnchorXpgList.GetItems(Index: integer): TCT_RelSizeAnchor;
begin
  Result := TCT_RelSizeAnchor(inherited Items[Index]);
end;

procedure TCT_RelSizeAnchorXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    if xaElements in Items[i].Assigneds then
    begin
      AWriter.AddAttribute('xmlns:cdr',OOXML_URI_OFFICEDOC_CHARTDRAWING);
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag(AName);
end;

initialization
  L_DrwCounter := 2000;

end.
