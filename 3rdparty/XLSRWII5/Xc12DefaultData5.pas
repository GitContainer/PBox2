unit Xc12DefaultData5;

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

const XLS_DEFAULT_VML =
'<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
'	<o:shapelayout v:ext="edit">' +
'		<o:idmap v:ext="edit" data="1"/>' +
'	</o:shapelayout>' +
'	<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
'		<v:stroke joinstyle="miter"/>' +
'		<v:path gradientshapeok="t" o:connecttype="rect"/>' +
'	</v:shapetype>' +
'</xml>';


const XLS_DEFAULT_DOCPROPS_CORE =
'<cp:coreProperties ' +
'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
'xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
'xmlns:dcterms="http://purl.org/dc/terms/" ' +
'xmlns:dcmitype="http://purl.org/dc/dcmitype/" ' +
'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' + #13 +
'  <dc:creator></dc:creator>' + #13 +
'  <cp:lastModifiedBy></cp:lastModifiedBy>' + #13 +
'  <dcterms:created xsi:type="dcterms:W3CDTF">2000-01-01T12:00:00Z</dcterms:created>' + #13 +
'  <dcterms:modified xsi:type="dcterms:W3CDTF">2000-01-01T12:00:00Z</dcterms:modified>' + #13 +
'</cp:coreProperties>';

const XLS_DEFAULT_THEME =
'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">' + #13 +
'	<a:themeElements>' + #13 +
'		<a:clrScheme name="Office">' + #13 +
'			<a:dk1>' + #13 +
'				<a:sysClr val="windowText" lastClr="000000"/>' + #13 +
'			</a:dk1>' + #13 +
'			<a:lt1>' + #13 +
'				<a:sysClr val="window" lastClr="FFFFFF"/>' + #13 +
'			</a:lt1>' + #13 +
'			<a:dk2>' + #13 +
'				<a:srgbClr val="1F497D"/>' + #13 +
'			</a:dk2>' + #13 +
'			<a:lt2>' + #13 +
'				<a:srgbClr val="EEECE1"/>' + #13 +
'			</a:lt2>' + #13 +
'			<a:accent1>' + #13 +
'				<a:srgbClr val="4F81BD"/>' + #13 +
'			</a:accent1>' + #13 +
'			<a:accent2>' + #13 +
'				<a:srgbClr val="C0504D"/>' + #13 +
'			</a:accent2>' + #13 +
'			<a:accent3>' + #13 +
'				<a:srgbClr val="9BBB59"/>' + #13 +
'			</a:accent3>' + #13 +
'			<a:accent4>' + #13 +
'				<a:srgbClr val="8064A2"/>' + #13 +
'			</a:accent4>' + #13 +
'			<a:accent5>' + #13 +
'				<a:srgbClr val="4BACC6"/>' + #13 +
'			</a:accent5>' + #13 +
'			<a:accent6>' + #13 +
'				<a:srgbClr val="F79646"/>' + #13 +
'			</a:accent6>' + #13 +
'			<a:hlink>' + #13 +
'				<a:srgbClr val="0000FF"/>' + #13 +
'			</a:hlink>' + #13 +
'			<a:folHlink>' + #13 +
'				<a:srgbClr val="800080"/>' + #13 +
'			</a:folHlink>' + #13 +
'		</a:clrScheme>' + #13 +
'		<a:fontScheme name="Office">' + #13 +
'			<a:majorFont>' + #13 +
'				<a:latin typeface="Cambria"/>' + #13 +
'				<a:ea typeface=""/>' + #13 +
'				<a:cs typeface=""/>' + #13 +
'				<a:font script="Jpan" typeface="?? ?????"/>' + #13 +
'				<a:font script="Hang" typeface="?? ??"/>' + #13 +
'				<a:font script="Hans" typeface="??"/>' + #13 +
'				<a:font script="Hant" typeface="????"/>' + #13 +
'				<a:font script="Arab" typeface="Times New Roman"/>' + #13 +
'				<a:font script="Hebr" typeface="Times New Roman"/>' + #13 +
'				<a:font script="Thai" typeface="Tahoma"/>' + #13 +
'				<a:font script="Ethi" typeface="Nyala"/>' + #13 +
'				<a:font script="Beng" typeface="Vrinda"/>' + #13 +
'				<a:font script="Gujr" typeface="Shruti"/>' + #13 +
'				<a:font script="Khmr" typeface="MoolBoran"/>' + #13 +
'				<a:font script="Knda" typeface="Tunga"/>' + #13 +
'				<a:font script="Guru" typeface="Raavi"/>' + #13 +
'				<a:font script="Cans" typeface="Euphemia"/>' + #13 +
'				<a:font script="Cher" typeface="Plantagenet Cherokee"/>' + #13 +
'				<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>' + #13 +
'				<a:font script="Tibt" typeface="Microsoft Himalaya"/>' + #13 +
'				<a:font script="Thaa" typeface="MV Boli"/>' + #13 +
'				<a:font script="Deva" typeface="Mangal"/>' + #13 +
'				<a:font script="Telu" typeface="Gautami"/>' + #13 +
'				<a:font script="Taml" typeface="Latha"/>' + #13 +
'				<a:font script="Syrc" typeface="Estrangelo Edessa"/>' + #13 +
'				<a:font script="Orya" typeface="Kalinga"/>' + #13 +
'				<a:font script="Mlym" typeface="Kartika"/>' + #13 +
'				<a:font script="Laoo" typeface="DokChampa"/>' + #13 +
'				<a:font script="Sinh" typeface="Iskoola Pota"/>' + #13 +
'				<a:font script="Mong" typeface="Mongolian Baiti"/>' + #13 +
'				<a:font script="Viet" typeface="Times New Roman"/>' + #13 +
'				<a:font script="Uigh" typeface="Microsoft Uighur"/>' + #13 +
'			</a:majorFont>' + #13 +
'			<a:minorFont>' + #13 +
'				<a:latin typeface="Calibri"/>' + #13 +
'				<a:ea typeface=""/>' + #13 +
'				<a:cs typeface=""/>' + #13 +
'				<a:font script="Jpan" typeface="?? ?????"/>' + #13 +
'				<a:font script="Hang" typeface="?? ??"/>' + #13 +
'				<a:font script="Hans" typeface="??"/>' + #13 +
'				<a:font script="Hant" typeface="????"/>' + #13 +
'				<a:font script="Arab" typeface="Arial"/>' + #13 +
'				<a:font script="Hebr" typeface="Arial"/>' + #13 +
'				<a:font script="Thai" typeface="Tahoma"/>' + #13 +
'				<a:font script="Ethi" typeface="Nyala"/>' + #13 +
'				<a:font script="Beng" typeface="Vrinda"/>' + #13 +
'				<a:font script="Gujr" typeface="Shruti"/>' + #13 +
'				<a:font script="Khmr" typeface="DaunPenh"/>' + #13 +
'				<a:font script="Knda" typeface="Tunga"/>' + #13 +
'				<a:font script="Guru" typeface="Raavi"/>' + #13 +
'				<a:font script="Cans" typeface="Euphemia"/>' + #13 +
'				<a:font script="Cher" typeface="Plantagenet Cherokee"/>' + #13 +
'				<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>' + #13 +
'				<a:font script="Tibt" typeface="Microsoft Himalaya"/>' + #13 +
'				<a:font script="Thaa" typeface="MV Boli"/>' + #13 +
'				<a:font script="Deva" typeface="Mangal"/>' + #13 +
'				<a:font script="Telu" typeface="Gautami"/>' + #13 +
'				<a:font script="Taml" typeface="Latha"/>' + #13 +
'				<a:font script="Syrc" typeface="Estrangelo Edessa"/>' + #13 +
'				<a:font script="Orya" typeface="Kalinga"/>' + #13 +
'				<a:font script="Mlym" typeface="Kartika"/>' + #13 +
'				<a:font script="Laoo" typeface="DokChampa"/>' + #13 +
'				<a:font script="Sinh" typeface="Iskoola Pota"/>' + #13 +
'				<a:font script="Mong" typeface="Mongolian Baiti"/>' + #13 +
'				<a:font script="Viet" typeface="Arial"/>' + #13 +
'				<a:font script="Uigh" typeface="Microsoft Uighur"/>' + #13 +
'			</a:minorFont>' + #13 +
'		</a:fontScheme>' + #13 +
'		<a:fmtScheme name="Office">' + #13 +
'			<a:fillStyleLst>' + #13 +
'				<a:solidFill>' + #13 +
'					<a:schemeClr val="phClr"/>' + #13 +
'				</a:solidFill>' + #13 +
'				<a:gradFill rotWithShape="1">' + #13 +
'					<a:gsLst>' + #13 +
'						<a:gs pos="0">' + #13 +
'							<a:schemeClr val="phClr">' + #13 +
'								<a:tint val="50000"/>' + #13 +
'								<a:satMod val="300000"/>' + #13 +
'							</a:schemeClr>' + #13 +
'						</a:gs>' + #13 +
'						<a:gs pos="35000">' + #13 +
'							<a:schemeClr val="phClr">' + #13 +
'								<a:tint val="37000"/>' + #13 +
'								<a:satMod val="300000"/>' + #13 +
'							</a:schemeClr>' + #13 +
'						</a:gs>' + #13 +
'						<a:gs pos="100000">' + #13 +
'							<a:schemeClr val="phClr">' + #13 +
'								<a:tint val="15000"/>' + #13 +
'								<a:satMod val="350000"/>' + #13 +
'							</a:schemeClr>' + #13 +
'						</a:gs>' + #13 +
'					</a:gsLst>' + #13 +
'					<a:lin ang="16200000" scaled="1"/>' + #13 +
'				</a:gradFill>' + #13 +
'				<a:gradFill rotWithShape="1">' + #13 +
'					<a:gsLst>' + #13 +
'						<a:gs pos="0">' + #13 +
'							<a:schemeClr val="phClr">' + #13 +
'								<a:shade val="51000"/>' + #13 +
'								<a:satMod val="130000"/>' + #13 +
'							</a:schemeClr>' + #13 +
'						</a:gs>' + #13 +
'						<a:gs pos="80000">' + #13 +
'							<a:schemeClr val="phClr">' + #13 +
'								<a:shade val="93000"/>' + #13 +
'								<a:satMod val="130000"/>' + #13 +
'							</a:schemeClr>' + #13 +
'						</a:gs>' + #13 +
'						<a:gs pos="100000">' + #13 +
'							<a:schemeClr val="phClr">' + #13 +
'								<a:shade val="94000"/>' + #13 +
'								<a:satMod val="135000"/>' + #13 +
'							</a:schemeClr>' + #13 +
'						</a:gs>' + #13 +
'					</a:gsLst>' + #13 +
'					<a:lin ang="16200000" scaled="0"/>' + #13 +
'				</a:gradFill>' + #13 +
'			</a:fillStyleLst>' + #13 +
'			<a:lnStyleLst>' + #13 +
'				<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + #13 +
'					<a:solidFill>' + #13 +
'						<a:schemeClr val="phClr">' + #13 +
'							<a:shade val="95000"/>' + #13 +
'							<a:satMod val="105000"/>' + #13 +
'						</a:schemeClr>' + #13 +
'					</a:solidFill>' + #13 +
'					<a:prstDash val="solid"/>' + #13 +
'				</a:ln>' + #13 +
'				<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">' + #13 +
'					<a:solidFill>' + #13 +
'						<a:schemeClr val="phClr"/>' + #13 +
'					</a:solidFill>' + #13 +
'					<a:prstDash val="solid"/>' + #13 +
'				</a:ln>' + #13 +
'				<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">' + #13 +
'					<a:solidFill>' + #13 +
'						<a:schemeClr val="phClr"/>' + #13 +
'					</a:solidFill>' + #13 +
'					<a:prstDash val="solid"/>' + #13 +
'				</a:ln>' + #13 +
'			</a:lnStyleLst>' + #13 +
'			<a:effectStyleLst>' + #13 +
'				<a:effectStyle>' + #13 +
'					<a:effectLst>' + #13 +
'						<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">' + #13 +
'							<a:srgbClr val="000000">' + #13 +
'								<a:alpha val="38000"/>' + #13 +
'							</a:srgbClr>' + #13 +
'						</a:outerShdw>' + #13 +
'					</a:effectLst>' + #13 +
'				</a:effectStyle>' + #13 +
'				<a:effectStyle>' + #13 +
'					<a:effectLst>' + #13 +
'						<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">' + #13 +
'							<a:srgbClr val="000000">' + #13 +
'								<a:alpha val="35000"/>' + #13 +
'							</a:srgbClr>' + #13 +
'						</a:outerShdw>' + #13 +
'					</a:effectLst>' + #13 +
'				</a:effectStyle>' + #13 +
'				<a:effectStyle>' + #13 +
'					<a:effectLst>' + #13 +
'						<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">' + #13 +
'							<a:srgbClr val="000000">' + #13 +
'								<a:alpha val="35000"/>' + #13 +
'							</a:srgbClr>' + #13 +
'						</a:outerShdw>' + #13 +
'					</a:effectLst>' + #13 +
'					<a:scene3d>' + #13 +
'						<a:camera prst="orthographicFront">' + #13 +
'							<a:rot lat="0" lon="0" rev="0"/>' + #13 +
'						</a:camera>' + #13 +
'						<a:lightRig rig="threePt" dir="t">' + #13 +
'							<a:rot lat="0" lon="0" rev="1200000"/>' + #13 +
'						</a:lightRig>' + #13 +
'					</a:scene3d>' + #13 +
'					<a:sp3d>' + #13 +
'						<a:bevelT w="63500" h="25400"/>' + #13 +
'					</a:sp3d>' + #13 +
'				</a:effectStyle>' + #13 +
'			</a:effectStyleLst>' + #13 +
'			<a:bgFillStyleLst>' + #13 +
'				<a:solidFill>' + #13 +
'					<a:schemeClr val="phClr"/>' + #13 +
'				</a:solidFill>' + #13 +
'				<a:gradFill rotWithShape="1">' + #13 +
'					<a:gsLst>' + #13 +
'						<a:gs pos="0">' + #13 +
'							<a:schemeClr val="phClr">' + #13 +
'								<a:tint val="40000"/>' + #13 +
'								<a:satMod val="350000"/>' + #13 +
'							</a:schemeClr>' + #13 +
'						</a:gs>' + #13 +
'						<a:gs pos="40000">' + #13 +
'							<a:schemeClr val="phClr">' + #13 +
'								<a:tint val="45000"/>' + #13 +
'								<a:shade val="99000"/>' + #13 +
'								<a:satMod val="350000"/>' + #13 +
'							</a:schemeClr>' + #13 +
'						</a:gs>' + #13 +
'						<a:gs pos="100000">' + #13 +
'							<a:schemeClr val="phClr">' + #13 +
'								<a:shade val="20000"/>' + #13 +
'								<a:satMod val="255000"/>' + #13 +
'							</a:schemeClr>' + #13 +
'						</a:gs>' + #13 +
'					</a:gsLst>' + #13 +
'					<a:path path="circle">' + #13 +
'						<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>' + #13 +
'					</a:path>' + #13 +
'				</a:gradFill>' + #13 +
'				<a:gradFill rotWithShape="1">' + #13 +
'					<a:gsLst>' + #13 +
'						<a:gs pos="0">' + #13 +
'							<a:schemeClr val="phClr">' + #13 +
'								<a:tint val="80000"/>' + #13 +
'								<a:satMod val="300000"/>' + #13 +
'							</a:schemeClr>' + #13 +
'						</a:gs>' + #13 +
'						<a:gs pos="100000">' + #13 +
'							<a:schemeClr val="phClr">' + #13 +
'								<a:shade val="30000"/>' + #13 +
'								<a:satMod val="200000"/>' + #13 +
'							</a:schemeClr>' + #13 +
'						</a:gs>' + #13 +
'					</a:gsLst>' + #13 +
'					<a:path path="circle">' + #13 +
'						<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>' + #13 +
'					</a:path>' + #13 +
'				</a:gradFill>' + #13 +
'			</a:bgFillStyleLst>' + #13 +
'		</a:fmtScheme>' + #13 +
'	</a:themeElements>' + #13 +
'	<a:objectDefaults/>' + #13 +
'	<a:extraClrSchemeLst/>' + #13 +
'</a:theme>';

const XLS_DEFAULT_FILE_97: array[0..1739] of byte =(
$50,$4B,$03,$04,$14,$00,$00,$00,$08,$00,$76,$8D,$73,$42,$18,$EC,$86,$AE,$34,$06,$00,$00,$00,$36,$00,$00,
$09,$00,$00,$00,$45,$6D,$70,$74,$79,$2E,$78,$6C,$73,$ED,$5B,$5B,$6C,$14,$55,$18,$FE,$66,$BB,$BB,$6D,$A1,
$76,$A7,$E5,$B6,$72,$59,$37,$6D,$8D,$DB,$CB,$92,$29,$16,$6C,$BD,$F4,$02,$A5,$84,$04,$B4,$E1,$22,$04,$8B,
$4B,$29,$25,$2D,$BD,$EC,$66,$5B,$0D,$A2,$E2,$2A,$20,$3E,$A0,$96,$60,$94,$98,$0A,$A2,$4F,$C6,$44,$31,$18,
$11,$CB,$43,$EB,$25,$F0,$20,$17,$45,$22,$E8,$4B,$21,$26,$BE,$F0,$E0,$3D,$42,$2A,$EB,$FF,$9F,$33,$B3,$9D,
$DD,$52,$68,$49,$C4,$18,$CF,$B7,$39,$B7,$EF,$FC,$E7,$3F,$FF,$CC,$99,$39,$E7,$FC,$33,$B3,$A7,$4F,$E5,$0C,
$1E,$F8,$60,$FA,$05,$A4,$A0,$12,$69,$B8,$1A,$CF,$84,$DB,$C6,$69,$14,$6E,$B7,$0A,$3A,$A8,$3E,$1E,$E7,$AC,
$95,$7A,$29,$C4,$15,$FE,$53,$C8,$CC,$A0,$81,$74,$BB,$70,$46,$FF,$32,$7D,$80,$C6,$8F,$C7,$FB,$02,$1C,$38,
$E8,$1C,$A0,$18,$B8,$48,$A1,$01,$11,$1E,$5E,$FF,$AD,$C4,$7C,$61,$43,$A3,$C6,$36,$3C,$40,$B1,$86,$5E,$62,
$B2,$E9,$FA,$63,$26,$57,$C4,$93,$44,$FC,$BE,$90,$39,$2A,$25,$A9,$A6,$45,$DB,$89,$3D,$ED,$81,$39,$E5,$E6,
$35,$BB,$DA,$51,$2D,$E4,$5E,$14,$71,$9E,$88,$B3,$C1,$1A,$0F,$8B,$36,$DF,$09,$A6,$14,$D3,$71,$9C,$AF,$DF,
$67,$7A,$34,$D1,$10,$2E,$AD,$06,$51,$B4,$A2,$11,$ED,$FF,$60,$AD,$CF,$E9,$83,$8B,$EF,$AA,$FC,$92,$FC,$7C,
$A3,$C1,$9F,$D7,$16,$CD,$BB,$AF,$21,$68,$2B,$F9,$9C,$F9,$34,$2A,$BE,$64,$89,$47,$96,$35,$6F,$58,$9B,$22,
$56,$80,$74,$DC,$61,$8A,$CD,$36,$52,$74,$25,$08,$9F,$B3,$10,$19,$28,$18,$21,$67,$D7,$68,$13,$AE,$44,$11,
$2A,$80,$50,$B0,$C8,$6F,$EB,$2C,$14,$24,$BD,$A9,$0C,$CB,$E4,$05,$F3,$6C,$E5,$EA,$50,$90,$35,$14,$A6,$68,
$08,$B5,$85,$A2,$C9,$1A,$4C,$26,$A1,$21,$51,$16,$1A,$EA,$50,$82,$1A,$9B,$86,$84,$79,$76,$25,$76,$D2,$D4,
$53,$55,$95,$6A,$4B,$1D,$8A,$47,$6A,$4A,$35,$C7,$4E,$DA,$34,$D9,$6D,$1A,$C4,$64,$31,$92,$BF,$C7,$FD,$18,
$46,$BF,$9F,$79,$CD,$E2,$7F,$1B,$1B,$EF,$18,$27,$8F,$FF,$21,$AF,$C1,$8F,$61,$58,$E7,$B3,$58,$CA,$FF,$99,
$CA,$17,$8E,$C2,$97,$8C,$C2,$17,$8D,$C2,$67,$8E,$E0,$F7,$38,$9C,$D0,$63,$69,$71,$4E,$73,$62,$6E,$91,$E6,
$C6,$9C,$22,$9D,$14,$4B,$17,$29,$62,$10,$E9,$E4,$98,$2B,$BE,$4E,$CC,$5C,$3B,$68,$BE,$D9,$95,$C9,$1A,$DC,
$58,$DE,$D2,$DC,$DC,$5D,$BA,$0B,$4E,$E2,$67,$63,$88,$6E,$45,$0B,$71,$5C,$49,$E3,$D2,$DE,$BA,$88,$51,$4B,
$69,$A6,$53,$F2,$E7,$0F,$77,$88,$72,$7B,$BF,$17,$3D,$33,$F7,$19,$69,$94,$EF,$3D,$E2,$C5,$95,$8F,$3B,$44,
$1E,$66,$DD,$B7,$03,$1D,$06,$A7,$DC,$D1,$BD,$36,$9E,$67,$79,$CE,$B3,$EE,$05,$B4,$46,$7F,$F6,$54,$C4,$48,
$A7,$3C,$87,$59,$4E,$B9,$AA,$FF,$E4,$3E,$A3,$B5,$1F,$F1,$26,$6C,$39,$3A,$37,$1F,$5B,$B3,$1D,$8E,$09,$14,
$C0,$D3,$25,$E1,$E1,$6F,$B6,$66,$93,$AA,$1F,$39,$CF,$BC,$8B,$52,$96,$61,$59,$0B,$81,$D3,$6F,$18,$B8,$01,
$1C,$E9,$42,$E7,$98,$41,$7D,$E5,$72,$BA,$D5,$4C,$1B,$B7,$44,$8C,$8B,$64,$2F,$DB,$FF,$28,$D9,$CC,$29,$1F,
$83,$55,$B6,$EA,$C6,$02,$2D,$A5,$BC,$FD,$8F,$E5,$C6,$4C,$D2,$D1,$6D,$96,$1D,$38,$66,$00,$5F,$18,$76,$19,
$EF,$27,$5E,$D1,$EE,$C2,$FD,$1D,$46,$36,$D5,$7F,$05,$D9,$EF,$10,$A5,$7C,$EE,$4F,$4C,$71,$8A,$36,$97,$DE,
$3E,$66,$54,$4C,$81,$61,$9D,$E3,$54,$9B,$07,$A9,$CC,$FA,$D6,$F5,$0D,$F7,$F7,$D2,$CB,$1E,$D1,$57,$83,$29,
$C3,$63,$17,$21,$B9,$DE,$3E,$39,$36,$E9,$A8,$47,$33,$AD,$1D,$4D,$14,$77,$26,$5A,$19,$58,$8B,$9B,$C3,$E3,
$B4,$10,$F0,$60,$34,$EB,$30,$38,$FF,$EA,$A6,$15,$06,$5F,$76,$E7,$A9,$CC,$F9,$65,$91,$15,$C6,$A7,$C7,$75,
$91,$F2,$28,$FB,$29,$4C,$9A,$11,$A0,$FD,$40,$01,$4A,$E8,$3C,$70,$DB,$8A,$9D,$1E,$83,$79,$3E,$98,$CF,$A7,
$07,$84,$5E,$5E,$60,$86,$B4,$82,$44,$3F,$F9,$33,$02,$49,$FD,$56,$99,$69,$84,$74,$54,$3C,$EF,$31,$DC,$66,
$DB,$20,$E4,$38,$BA,$1D,$05,$D8,$47,$C7,$CC,$C7,$CD,$BA,$98,$5B,$4D,$E1,$E4,$9A,$15,$06,$B7,$63,$7B,$23,
$64,$DF,$25,$67,$00,$4B,$20,$ED,$E5,$3C,$73,$4B,$4C,$B9,$2D,$D4,$36,$D6,$37,$7C,$4D,$33,$DE,$3A,$A6,$1B,
$D6,$7D,$C5,$60,$0B,$8F,$7E,$EF,$C5,$AF,$DB,$3C,$46,$56,$9F,$1C,$D7,$B2,$03,$C3,$B6,$5A,$ED,$2D,$EE,$43,
$B2,$B7,$73,$9B,$1C,$A3,$1B,$61,$89,$4B,$8E,$1B,$CF,$60,$83,$14,$5A,$28,$CF,$7D,$31,$CE,$ED,$0F,$60,$3B,
$95,$5F,$F7,$C3,$E0,$7C,$57,$CF,$04,$62,$E5,$5E,$4C,$4F,$DA,$8B,$4D,$74,$E8,$49,$5A,$6B,$49,$F0,$36,$B1,
$7B,$C9,$A2,$78,$03,$3C,$22,$9F,$23,$46,$42,$A7,$73,$35,$F4,$CE,$CF,$5F,$2F,$5D,$5F,$5F,$15,$12,$7C,$91,
$E0,$8B,$45,$FC,$9C,$60,$62,$B6,$19,$E7,$4E,$9E,$AD,$68,$DE,$79,$96,$6A,$06,$9C,$6C,$E7,$14,$0A,$DB,$84,
$F4,$76,$11,$1F,$A0,$FD,$12,$4B,$68,$E2,$E7,$14,$63,$22,$31,$58,$65,$A5,$27,$D7,$AC,$24,$59,$D2,$EA,$C8,
$46,$32,$2A,$E9,$86,$FD,$48,$38,$10,$D5,$36,$76,$16,$D9,$6C,$CE,$5B,$D6,$ED,$37,$01,$0A,$0A,$0A,$0A,$0A,
$B7,$18,$57,$69,$6F,$E8,$D6,$46,$6E,$C8,$78,$6A,$1E,$DC,$B1,$FF,$97,$CB,$0F,$B5,$E8,$EF,$EE,$CE,$40,$F1,
$5D,$87,$CE,$1B,$C4,$F5,$42,$AE,$4B,$5C,$CF,$1E,$36,$AF,$20,$D5,$90,$BB,$4B,$5E,$A1,$79,$8B,$16,$A1,$90,
$45,$E1,$05,$B0,$8F,$0E,$BC,$02,$88,$D5,$E4,$07,$12,$26,$3F,$96,$D6,$2F,$DA,$49,$D8,$00,$6C,$BE,$2E,$AF,
$53,$58,$DA,$DA,$14,$0D,$77,$85,$37,$76,$FB,$17,$6E,$6E,$6A,$6E,$17,$7D,$C6,$4E,$F5,$14,$F6,$17,$9C,$D4,
$AC,$5D,$B0,$82,$82,$82,$82,$82,$82,$82,$82,$82,$82,$82,$82,$82,$82,$C2,$B5,$31,$9A,$FF,$CF,$8C,$E3,$EC,
$89,$B3,$BD,$B3,$67,$E8,$7B,$5E,$23,$FF,$BF,$E4,$F2,$7B,$FC,$16,$CA,$95,$C2,$CD,$D4,$80,$73,$90,$6F,$3E,
$F8,$99,$40,$3D,$05,$0F,$E4,$B3,$80,$69,$60,$0F,$1E,$98,$08,$7E,$29,$26,$FD,$78,$EB,$99,$40,$0F,$85,$A9,
$14,$F6,$42,$FA,$FD,$6F,$42,$3E,$33,$38,$84,$E4,$67,$05,$2C,$33,$EC,$FB,$2F,$08,$47,$23,$E1,$68,$63,$77,
$6B,$B8,$13,$DD,$5D,$EC,$F7,$6F,$D2,$33,$84,$7E,$98,$FD,$5C,$2B,$F5,$E9,$C3,$6F,$88,$E4,$FB,$37,$64,$E9,
$B2,$1B,$9F,$29,$B6,$2A,$1C,$6D,$EB,$E2,$AA,$2E,$F1,$6C,$5A,$3E,$98,$66,$DB,$AC,$67,$0B,$7E,$93,$9D,$07,
$D9,$B0,$D2,$2C,$73,$9E,$1F,$5F,$87,$EA,$17,$D7,$86,$16,$AD,$5C,$5C,$9B,$B0,$BE,$86,$D2,$07,$29,$3C,$89,
$85,$94,$AF,$41,$19,$B5,$2D,$C7,$7C,$2A,$05,$71,$37,$4A,$29,$3F,$8F,$72,$65,$C4,$94,$52,$39,$48,$12,$15,
$98,$8B,$05,$94,$33,$30,$87,$F8,$72,$DC,$43,$4C,$05,$D5,$97,$53,$A9,$0C,$B5,$78,$1A,$0A,$0A,$0A,$0A,$0A,
$0A,$0A,$0A,$0A,$0A,$0A,$0A,$0A,$0A,$37,$07,$CB,$87,$65,$3F,$97,$DF,$E5,$F3,$57,$9D,$FC,$7E,$9F,$7D,$65,
$FE,$5F,$07,$FB,$F5,$EC,$DF,$B2,$93,$CC,$FE,$39,$FB,$EA,$FC,$9D,$97,$C7,$AC,$CF,$81,$7C,$E7,$CF,$3E,$BD,
$F5,$05,$19,$FB,$EC,$D3,$CC,$FA,$BF,$28,$5C,$1D,$F5,$DF,$07,$0A,$FF,$36,$96,$21,$4C,$BF,$6E,$F8,$B1,$50,
$7C,$51,$1A,$C5,$13,$18,$0F,$A6,$C2,$A5,$59,$BA,$F8,$3A,$F2,$67,$C8,$67,$49,$FD,$B2,$BA,$CE,$2E,$7B,$70,
$D1,$E9,$DD,$FC,$BD,$86,$F5,$7F,$21,$C6,$2A,$EA,$3D,$8A,$36,$AC,$17,$76,$B4,$61,$BC,$C8,$85,$23,$D1,$7F,
$DC,$A6,$F7,$86,$D0,$65,$E2,$C2,$72,$3C,$86,$0E,$FA,$35,$8A,$63,$5F,$4C,$67,$61,$A3,$B0,$89,$99,$6E,$B4,
$52,$BE,$F3,$3A,$6A,$02,$D4,$3F,$DF,$43,$7C,$FF,$8C,$B5,$7F,$F1,$05,$A6,$2E,$F3,$2E,$D4,$52,$0F,$4D,$C2,
$06,$F9,$4D,$EF,$F8,$EC,$29,$BF,$89,$E3,$D7,$13,$11,$F0,$37,$50,$4B,$01,$02,$14,$0B,$14,$00,$00,$00,$08,
$00,$76,$8D,$73,$42,$18,$EC,$86,$AE,$34,$06,$00,$00,$00,$36,$00,$00,$09,$00,$24,$00,$00,$00,$00,$00,$00,
$00,$20,$00,$00,$00,$00,$00,$00,$00,$45,$6D,$70,$74,$79,$2E,$78,$6C,$73,$0A,$00,$20,$00,$00,$00,$00,$00,
$01,$00,$18,$00,$9F,$40,$B0,$EA,$C0,$24,$CE,$01,$AB,$4D,$CA,$59,$C0,$24,$CE,$01,$C1,$6F,$C3,$59,$C0,$24,
$CE,$01,$50,$4B,$05,$06,$00,$00,$00,$00,$01,$00,$01,$00,$5B,$00,$00,$00,$5B,$06,$00,$00,$00,$00);

implementation

end.
