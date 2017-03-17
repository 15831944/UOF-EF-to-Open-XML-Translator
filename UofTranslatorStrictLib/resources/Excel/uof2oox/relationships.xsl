<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:uof="http://schemas.uof.org/cn/2003/uof" xmlns:表="http://schemas.uof.org/cn/2003/uof-spreadsheet" xmlns:演="http://schemas.uof.org/cn/2003/uof-slideshow" xmlns:字="http://schemas.uof.org/cn/2003/uof-wordproc" xmlns:图="http://schemas.uof.org/cn/2003/graph" xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:ws="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:pzip="urn:u2o:xmlns:post-processings:special" xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
                xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<xsl:template match="uof:电子表格" mode="rels">
	<!--这个文档没用-->
		<!--Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings" Target="sharedStrings.xml"/>
    <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId3" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/theme" Target="theme/theme1.xml"/>



	<Relationship Id="rId0" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="worksheets/sheet2.xml"/>
	
	<Relationship Id="rId0" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="worksheets/sheet1.xml"/>
	<Relationship Id="rId0" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet" Target="chartsheets/sheet1.xml"/>
	
	
	<Relationship Id="rId0" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships-->
		<!--workbook.xl.rels-->
		<pzip:entry pzip:target="xl/_rels/workbook.xml.rels">
			<pr:Relationships>
				<pr:Relationship pr:Id="rId1" pr:Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings" pr:Target="sharedStrings.xml"/>
				<pr:Relationship pr:Id="rId2" pr:Type="http://purl.oclc.org/ooxml/officeDocument/relationships/styles" pr:Target="styles.xml"/>
				<pr:Relationship pr:Id="rId3" pr:Type="http://purl.oclc.org/ooxml/officeDocument/relationships/theme" pr:Target="theme/theme1.xml"/>
				
			</pr:Relationships>
		</pzip:entry>
		<!--chartsheets里的relationship-->
		<!--worksheets里的relationship-->
		<!--drawings里的relationship-->
	</xsl:template>
</xsl:stylesheet>
