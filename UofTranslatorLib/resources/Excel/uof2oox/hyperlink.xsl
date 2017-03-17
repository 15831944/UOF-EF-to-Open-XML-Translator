<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                 xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:图表="http://schemas.uof.org/cn/2009/chart"
                xmlns:式样="http://schemas.uof.org/cn/2009/styles"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" 
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" 
                >

  
	<xsl:template name="HyperLink2">
		<hyperlink>
			<xsl:variable name="colSeqTemp" select="./@表:列号"/>
			<xsl:variable name="rowSeq" select="./parent::表:行/@表:行号"/>
      <xsl:variable name="colSeq">
        <xsl:call-template name="ColIndex">
          <xsl:with-param name="colSeq" select="$colSeqTemp"/>
        </xsl:call-template>
      </xsl:variable>
			<xsl:attribute name="ref">
        <xsl:value-of select="concat($colSeq,$rowSeq)"/>
      </xsl:attribute>
      <xsl:variable name="pos" select="translate(@表:超链接引用,'ID','')"/>
			<xsl:variable name="refId" select="@表:超链接引用"/>
			<xsl:if test="ancestor::uof:电子表格/preceding-sibling::uof:链接集/uof:超级链接[@uof:标识符=$refId]">
				<xsl:variable name="tar" select="ancestor::uof:电子表格/preceding-sibling::uof:链接集/uof:超级链接[@uof:标识符=$refId]/@uof:目标"/>
        <xsl:choose>
          <xsl:when test="contains($tar,'!')">
            <xsl:attribute name="location">
            <xsl:value-of select="$tar"/>
          </xsl:attribute>
          </xsl:when>
          <xsl:when test="not(contains($tar,'!'))">
            <xsl:attribute name="r:id">
            <xsl:value-of select="concat('rId_HyperLink_',$pos)"/>
          </xsl:attribute>
          </xsl:when>
        </xsl:choose>
        <xsl:if test="ancestor::uof:电子表格/preceding-sibling::uof:链接集/uof:超级链接[@uof:标识符=$refId]/@uof:提示">
          <xsl:attribute name="tooltip">
            <xsl:value-of select="ancestor::uof:电子表格/preceding-sibling::uof:链接集/uof:超级链接[@uof:标识符=$refId]/@uof:提示"/>
          </xsl:attribute>
        </xsl:if>
			</xsl:if>
		</hyperlink>
	</xsl:template>

  <!--模板功能：将列号的数字表示转换成字母表示-->
  <!--最多只能转换以两个字母表示的列，如AF,D等-->
  <!--Modified by LDM in 2010/12/18-->
  <!--ColIndex-->
  
  <!--2014-3-19, update by Qihy, 增加两个字母表示的列的转换，三个字母的暂时未考虑，修复BUG3117 互操作-ooxml-uof-oo（2010​2013）需要修复， start-->
  <!--<xsl:template name="ColIndex">
    <xsl:param name="colSeq"/>
    <xsl:choose>
      <xsl:when test="$colSeq &lt; 27">
        <xsl:choose>
          <xsl:when test="$colSeq &lt; 10">
            <xsl:value-of select="translate($colSeq,'123456789','ABCDEFGHI')"/>
          </xsl:when>
          <xsl:when test="($colSeq &lt;19) and ($colSeq &gt; 9)">
            <xsl:value-of select="translate($colSeq - 9,'123456789','JKLMNOPQR')"/>
          </xsl:when>
          <xsl:when test="($colSeq &lt;27) and ($colSeq &gt; 18)">
            <xsl:value-of select="translate($colSeq - 18,'12345678','STUVWXYZ')"/>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>-->
  <xsl:template name="ColIndex">
    <xsl:param name="colSeq"/>
    <xsl:choose>
      <xsl:when test = "$colSeq &gt; 26">
        <xsl:choose>
          <xsl:when test ="$colSeq mod 26 = 0">
            <xsl:variable name ="col_front">
              <xsl:call-template name ="ColIndex1">
                <xsl:with-param name ="colSeq" select ="floor($colSeq div 26) - 1"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name ="col_rear">
              <xsl:value-of select ="'Z'"/>
            </xsl:variable>
            <xsl:value-of select ="concat($col_front, $col_rear)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name ="col_front">
              <xsl:call-template name ="ColIndex1">
                <xsl:with-param name ="colSeq" select ="floor($colSeq div 26)"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name ="col_rear">
              <xsl:call-template name ="ColIndex1">
                <xsl:with-param name ="colSeq" select ="$colSeq mod 26"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select ="concat($col_front, $col_rear)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name ="ColIndex1">
          <xsl:with-param name ="colSeq" select ="$colSeq"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="ColIndex1">
    <xsl:param name="colSeq"/>
      <xsl:choose>
        <xsl:when test="$colSeq &lt; 10">
          <xsl:value-of select="translate($colSeq,'123456789','ABCDEFGHI')"/>
        </xsl:when>
        <xsl:when test="($colSeq &lt;19) and ($colSeq &gt; 9)">
          <xsl:value-of select="translate($colSeq - 9,'123456789','JKLMNOPQR')"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="translate($colSeq - 18,'12345678','STUVWXYZ')"/>
        </xsl:otherwise>
      </xsl:choose>
  </xsl:template>
  <!--2014-3-19 end-->
</xsl:stylesheet>
