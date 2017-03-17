echo off
set path=%path%;C:\Program Files\Microsoft SDKs\Windows\v7.0A\bin
cd oox2uof

xsltc /nologo /settings:dtd+,document-,script+ "metadata.xsl" "rowtrans.xslt" "hyperlinks.xsl"  "object.xsl" "oox2uof.xsl" "sheetbreak.xsl" "sheetcommon.xsl" "sheetfilter.xsl" "sheetprop.xsl" "sheets.xsl" "style.xsl" "table.xsl" /out:"../../../excel_oox2uof.dll"

cd..
cd uof2oox

xsltc /nologo /settings:dtd+,document-,script+ "chart.xsl" "contentTypes.xsl" "drawing.xsl" "hyperlink.xsl" "metadata.xsl" "package_relationships.xsl" "pPr.xsl" "predefined.xsl" "relationships.xsl" "sharedStrings.xsl" "sheet.xsl" "style.xsl" "theme.xsl"  "txBody.xsl"  "workbook.xsl"  "uof2oox.xsl"  /out:"../../../excel_uof2oox.dll"

pause
echo on