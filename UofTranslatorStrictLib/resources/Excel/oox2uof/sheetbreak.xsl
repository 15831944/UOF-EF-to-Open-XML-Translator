<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                
                xmlns:ws="http://purl.oclc.org/ooxml/spreadsheetml/main" 
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
  <xsl:template name="sheetbreak">
    <xsl:if test="ws:worksheet/ws:colBreaks or ws:worksheet/ws:rowBreaks">
      <表:分页符集_E81E>
        <!--yx,caution:@id not @id+1,oox2uof ok breakpos ok!-->
        <xsl:for-each select="ws:worksheet/ws:rowBreaks/ws:brk">
          <表:分页符_E81F>
            <xsl:attribute name="行号_E820">
              <xsl:value-of select="@id"/>
            </xsl:attribute>
          </表:分页符_E81F>
        </xsl:for-each>
        <xsl:for-each select="ws:worksheet/ws:colBreaks/ws:brk">
          <表:分页符_E81F>
            <xsl:attribute name="列号_E821">
              <xsl:value-of select="@id"/>
            </xsl:attribute>
          </表:分页符_E81F>
        </xsl:for-each>
      </表:分页符集_E81E>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>
