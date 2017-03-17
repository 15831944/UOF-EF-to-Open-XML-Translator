echo off
set path=%path%;C:\Program Files\Microsoft SDKs\Windows\v7.0A\bin
cd oox2uof

xsltc /nologo /settings:dtd+,document-,script+ "metadata.xsl" "rowtrans.xslt" "hyperlinks.xsl"  "object.xsl" "oox2uof.xsl" "sheetbreak.xsl" "sheetcommon.xsl" "sheetfilter.xsl" "sheetprop.xsl" "sheets.xsl" "style.xsl" "table.xsl" /out:"../../../excel_oox2uof.dll"

pause
echo on