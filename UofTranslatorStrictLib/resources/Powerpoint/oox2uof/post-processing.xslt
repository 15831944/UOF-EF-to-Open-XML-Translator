<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
xmlns:元="http://schemas.uof.org/cn/2009/metadata"
xmlns:规则="http://schemas.uof.org/cn/2009/rules"
xmlns:式样="http://schemas.uof.org/cn/2009/styles"
xmlns:对象="http://schemas.uof.org/cn/2009/objects"
xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:fo="http://www.w3.org/1999/XSL/Format"
xmlns:dc="http://purl.org/dc/elements/1.1/"
xmlns:dcterms="http://purl.org/dc/terms/"
xmlns:dcmitype="http://purl.org/dc/dcmitype/"
xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                
xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"                
xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
xmlns:vt="http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes"
xmlns:pic="http://purl.oclc.org/ooxml/drawingml/picture"
xmlns:cus="http://purl.oclc.org/ooxml/officeDocument/customProperties"
xmlns:dgm="http://purl.oclc.org/ooxml/drawingml/diagram"
                
xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
xmlns:pzip="urn:u2o:xmlns:post-processings:special">
  <xsl:import href="metadata.xsl"/>
  <!---->
  <xsl:import href="hyperlinks.xsl"/>
  <!---->
  <!--<xsl:import href="txstyles.xsl"/>-->
  <xsl:import href="styles.xsl"/>
  <xsl:import href="shapes.xsl"/>
  <xsl:import href="Master.xsl"/>
  <xsl:import href="slide.xsl"/>
  <xsl:import href="CommenRule.xsl"/>
  <xsl:param name="outputFile"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
  <xsl:template name="presentation" match="/">
    <pzip:archive pzip:target="{$outputFile}">
      <!--<uof:UOF xmlns:uof="http://schemas.uof.org/cn/2003/uof" xmlns:图="http://schemas.uof.org/cn/2003/graph" xmlns:字="http://schemas.uof.org/cn/2003/uof-wordproc" xmlns:表="http://schemas.uof.org/cn/2003/uof-spreadsheet" xmlns:演="http://schemas.uof.org/cn/2003/uof-slideshow" uof:language="cn" uof:version="1.1" uof:mimetype="vnd.uof.presentation" uof:locID="u0000" uof:attrList="language version mimetype">-->
      <xsl:for-each select="//child::pzip:archive">
        <xsl:copy-of select="*"/>
      </xsl:for-each>
    </pzip:archive>
  </xsl:template>
</xsl:stylesheet>

