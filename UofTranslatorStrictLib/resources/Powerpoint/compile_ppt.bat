echo off
set path=%path%;C:\Program Files\Microsoft SDKs\Windows\v7.0A\bin
cd oox2uof
xsltc /nologo /settings:dtd+,document-,script+  "animation.xsl" "table.xsl" "CommenRule.xsl" "fill.xsl" "hyperlinks.xsl" "Master.xsl" "metadata.xsl" "numbering.xsl" "oox2uof.xsl" "p.xsl" "ph.xsl" "PPr-commen.xsl" "pre2.xsl" "shapes.xsl" "sldLayout.xsl" "slide.xsl" "styles.xsl" "txStyles.xsl" /out:"../../../ppt_oox2uof.dll"
pause
echo on