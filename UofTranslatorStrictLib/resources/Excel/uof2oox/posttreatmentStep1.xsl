<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
                xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"
                xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:规则="http://schemas.uof.org/cn/2009/rules"
                xmlns:元="http://schemas.uof.org/cn/2009/metadata"
                xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
                xmlns:图表="http://schemas.uof.org/cn/2009/chart"
                xmlns:对象="http://schemas.uof.org/cn/2009/objects"
                xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
                xmlns:式样="http://schemas.uof.org/cn/2009/styles"
                xmlns:pzip="urn:u2o:xmlns:post-processings:special"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
                xmlns:xsd="http://www.ord.com"
				xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
				xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
				xmlns:dc="http://purl.org/dc/elements/1.1/"
				xmlns:dcterms="http://purl.org/dc/terms/"
				xmlns:dcmitype="http://purl.org/dc/dcmitype/"
				xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
				xmlns:p="http://purl.oclc.org/ooxml/presentationml/main">
  <xsl:param name="outputFile"/>
  <xsl:template match="//pzip:archive">
    <pzip:archive pzip:target="{$outputFile}">
      <xsl:for-each select="child::node()">
        <xsl:copy-of select="."/>
      </xsl:for-each>
    </pzip:archive>
  </xsl:template>
</xsl:stylesheet>
