echo off
set path=%path%;C:\Program Files\Microsoft SDKs\Windows\v7.0A\bin
cd oox2uof
xsltc /nologo /settings:dtd+,document-,script+  "animation.xsl" "table.xsl" "CommenRule.xsl" "fill.xsl" "hyperlinks.xsl" "Master.xsl" "metadata.xsl" "numbering.xsl" "oox2uof.xsl" "p.xsl" "ph.xsl" "PPr-commen.xsl" "pre2.xsl" "shapes.xsl" "sldLayout.xsl" "slide.xsl" "styles.xsl" "txStyles.xsl" /out:"../../../ppt_oox2uof.dll"
cd..
cd uof2oox
xsltc /nologo /settings:dtd+,document-,script+  "table.xsl" "animation.xsl" "contentTypes.xsl" "cSld.xsl" "handoutMasterRels.xsl" "nMasterRels1.xsl" "hyperlink.xsl" "LayoutRels.xsl" "metadata.xsl" "noteMasterRels.xsl" "noteRels.xsl" "numbering.xsl" "package_relationships.xsl" "pPr.xsl" "slideMasterRels.xsl" "SlideRels.xsl" "txBody.xsl" "uof2oox.xsl"  /out:"../../../ppt_uof2oox.dll"
pause
echo on