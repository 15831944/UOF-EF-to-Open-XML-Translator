<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0" 
xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
xmlns:fo="http://www.w3.org/1999/XSL/Format" 
xmlns:xs="http://www.w3.org/2001/XMLSchema" 
xmlns:fn="http://www.w3.org/2005/xpath-functions" 
xmlns:xdt="http://www.w3.org/2005/xpath-datatypes"
xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships" 
xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" 
xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:图="http://schemas.uof.org/cn/2009/graph"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:元="http://schemas.uof.org/cn/2009/metadata"
xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
xmlns:规则="http://schemas.uof.org/cn/2009/rules"
xmlns:式样="http://schemas.uof.org/cn/2009/styles">


  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>

  <xsl:template name="footer">
    <字:页脚_41F7>
      <xsl:for-each select="./w:footerReference ">

        <xsl:variable name="temp" select="@r:id"/>
        <xsl:variable name="target" select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$temp]/@Target"/>

        <xsl:choose>
          <!--偶页页脚-->
          <xsl:when test="@w:type='even'">
            <xsl:apply-templates select="document(concat('word/',$target))/w:ftr" mode="evenFooter"/>
          </xsl:when>
          <!--奇页页脚-->
          <xsl:when test="@w:type='default'">
            <xsl:apply-templates select="document(concat('word/',$target))/w:ftr" mode="oddFooter"/>
          </xsl:when>
          <!--首页页脚-->
          <xsl:when test="@w:type='first'">
            <xsl:apply-templates select="document(concat('word/',$target))/w:ftr" mode="coverFooter"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </字:页脚_41F7>
  </xsl:template>

  <xsl:template match="w:ftr" mode="coverFooter">
	  <字:首页页脚_41FA>
      <xsl:for-each select="w:p | w:tbl">
        <xsl:choose>
          <xsl:when test="name(.)='w:p'">
            <xsl:call-template name="paragraph"/>
          </xsl:when>
          <xsl:when test="name(.)='w:tbl'">
            <xsl:call-template name="table"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </字:首页页脚_41FA>
  </xsl:template>

  <xsl:template match="w:ftr" mode="oddFooter">
	  <字:奇数页页脚_41F8>
      <xsl:for-each select="w:p | w:tbl">
        <xsl:choose>
          <xsl:when test="name(.)='w:p'">
            <xsl:call-template name="paragraph"/>
          </xsl:when>
          <xsl:when test="name(.)='w:tbl'">
            <xsl:call-template name="table"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </字:奇数页页脚_41F8>
  </xsl:template>

  <xsl:template match="w:ftr" mode="evenFooter">
	  <字:偶数页页脚_41F9>
      <xsl:for-each select="w:p | w:tbl">
        <xsl:choose>
          <xsl:when test="name(.)='w:p'">
            <xsl:call-template name="paragraph"/>
          </xsl:when>
          <xsl:when test="name(.)='w:tbl'">
            <xsl:call-template name="table"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </字:偶数页页脚_41F9>
  </xsl:template>


</xsl:stylesheet>
<!--

<xsl:template name="footerReference">
	    <xsl:choose>
	        <xsl:when test="w:footerReference/@w:type='even'">
	            <xsl:apply-templates select="document('footer1.xml')/w:ftr" mode="footer1part"/>
	        </xsl:when>
	        <xsl:when test="w:footerReference/@w:type='default'">
	            <xsl:apply-templates select="document('footer2.xml')/w:ftr" mode="footer2part"/>
	        </xsl:when>
	        <xsl:when test="w:footerReference/@w:type='first'">
	            <xsl:apply-templates select="document('footer3.xml')/w:ftr" mode="footer3part"/>
	        </xsl:when>
	    </xsl:choose>
	</xsl:template>
	
	<xsl:template match="w:ftr" mode="footer1part">
	    <字:页脚 uof:locID="t0031">
	        <字:偶数页页脚 uof:locID="t0033">
	            <xsl:for-each select="node()">
	                <xsl:if test="name(.)='w:p'">
	                    <xsl:call-template name="paragraph"/>
	                </xsl:if>
	                <xsl:if test="name(.)='w:tbl'">
	                    <xsl:call-template name="table"/>
	                </xsl:if>
	            </xsl:for-each>
	        </字:偶数页页脚>
	    </字:页脚>
	</xsl:template>
	<xsl:template match="w:ftr" mode="footer2part">
	    <字:页脚 uof:locID="t0031">
	        <字:奇数页页脚 uof:locID="t0032">
	            <xsl:for-each select="node()">
	                <xsl:if test="name(.)='w:p'">
	                    <xsl:call-template name="paragraph"/>
	                </xsl:if>
	                <xsl:if test="name(.)='w:tbl'">
	                    <xsl:call-template name="table"/>
	                </xsl:if>
	            </xsl:for-each>
	        </字:奇数页页脚>
	    </字:页脚>
	</xsl:template>
	<xsl:template match="w:ftr" mode="footer3part">
	    <字:页脚 uof:locID="t0031">
	        <字:首页页脚 uof:locID="t0034">
	            <xsl:for-each select="node()">
	                <xsl:if test="name(.)='w:p'">
	                    <xsl:call-template name="paragraph"/>
	                </xsl:if>
	                <xsl:if test="name(.)='w:tbl'">
	                    <xsl:call-template name="table"/>
	                </xsl:if>
	            </xsl:for-each>
	        </字:首页页脚>
	    </字:页脚>
	</xsl:template>
</xsl:stylesheet>-->
