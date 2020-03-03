unit XLSDefaultDataXLSX5;

interface

const DEFAULT_DOCPROPS_CORE =
'<cp:coreProperties ' +
'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
'xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
'xmlns:dcterms="http://purl.org/dc/terms/" ' +
'xmlns:dcmitype="http://purl.org/dc/dcmitype/" ' +
'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' + #13 +
'  <dc:creator>John Doe II</dc:creator>' + #13 +
'  <cp:lastModifiedBy>John Doe II</cp:lastModifiedBy>' + #13 +
'  <dcterms:created xsi:type="dcterms:W3CDTF">2000-01-01T12:00:00Z</dcterms:created>' + #13 +
'  <dcterms:modified xsi:type="dcterms:W3CDTF">2000-01-01T12:00:00Z</dcterms:modified>' + #13 +
'</cp:coreProperties>';

const DEFAULT_THEME =
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

implementation

end.
