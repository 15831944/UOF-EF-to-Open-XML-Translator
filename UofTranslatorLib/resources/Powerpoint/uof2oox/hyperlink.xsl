<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
  xmlns:pzip="urn:u2o:xmlns:post-processings:special"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
 xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph"
xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks">
  <xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
  <xsl:template name="hyperlink">
    <a:hlinkClick>
      <xsl:variable name="cid">
        <xsl:value-of select="@标识符_AA0A"/>
      </xsl:variable>
      <xsl:variable name="ctarget">
        <!--<xsl:value-of select="@uof:目标"/>-->
		  <xsl:value-of select="超链:目标_AA01"/>
      </xsl:variable>
      <xsl:variable name="return">
        <xsl:if test="@return">
          <xsl:value-of select="concat('&amp;return=',@return)"/>
        </xsl:if>
      </xsl:variable>
      <xsl:if test ="@标识符_AA0A">
        <xsl:attribute name ="r:id">
          <xsl:choose>
            <xsl:when test="$ctarget!=''">
              <xsl:if test ="not(starts-with(超链:目标_AA01,'Custom Show') or 超链:目标_AA01='First Slide' or 超链:目标_AA01='Last Slide' or 超链:目标_AA01='Previous Slide' or 超链:目标_AA01='End Show' or 超链:目标_AA01='Next Slide')">
                <!-- 09.10.23 马有旭 修改            
            <xsl:choose>
              <xsl:when test="//uof:超级链接[@uof:目标=$ctarget and @uof:标识符!=$cid]">
                <xsl:value-of select="//uof:超级链接[@uof:目标=$ctarget][1]/@uof:标识符"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="@uof:标识符"/>
              </xsl:otherwise>
            </xsl:choose>
            <xsl:value-of select ="@uof:标识符"/>
            -->
                <xsl:value-of select ="@标识符_AA0A"/>
              </xsl:if>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="''"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:for-each select ="超链:目标_AA01">
        <xsl:choose>
          <xsl:when test ="starts-with(current(),'Slide:')">
            <xsl:attribute name ="action">ppaction://hlinksldjump</xsl:attribute>
          </xsl:when>
          <xsl:when test =".='First Slide'">
            <xsl:attribute name ="action">ppaction://hlinkshowjump?jump=firstslide</xsl:attribute>
          </xsl:when>
          <xsl:when test =".='Last Slide'">
            <xsl:attribute name ="action">ppaction://hlinkshowjump?jump=lastslide</xsl:attribute>
          </xsl:when>
          <xsl:when test =".='Previous Slide'">
            <xsl:attribute name ="action">ppaction://hlinkshowjump?jump=previousslide</xsl:attribute>
          </xsl:when>
          <xsl:when test =".='End Show'">
            <xsl:attribute name ="action">ppaction://hlinkshowjump?jump=endshow</xsl:attribute>
          </xsl:when>
          <xsl:when test =".='Next Slide'">
            <xsl:attribute name ="action">ppaction://hlinkshowjump?jump=nextslide</xsl:attribute>
          </xsl:when>
          <xsl:when test="''">
            <xsl:attribute name ="action">ppaction://noaction</xsl:attribute>
          </xsl:when>
          <xsl:when test ="starts-with(current(),'Custom Show:')">
            <xsl:attribute name ="action">
              <!--<xsl:value-of select ="concat('ppaction://customshow?id=',//演:幻灯片序列[@演:名称=substring-after(current(),'Custom Show:')]/@演:标识符,'&amp;return=true')"/>-->
              <xsl:variable name ="id">
                <xsl:call-template name="index">
                  <xsl:with-param name="name" select="//演:幻灯片序列[@名称_6B0B=substring-after(current(),'Custom Show:')]/@标识符_6B0A"/>
                </xsl:call-template>
              </xsl:variable>
              <xsl:value-of select="concat('ppaction://customshow?id=',$id,$return)"/>
            </xsl:attribute>
          </xsl:when>

          <xsl:otherwise>
            <!--文件地址与网络地址怎么区分 用“/”和“\”区分？？？>
                <xsl:if test="contains(current(),'')"></xsl:if-->
            <xsl:if test="not(contains(current(),'/')) and not(starts-with(current(),'mailto:'))">
              <xsl:attribute name ="action">
                <xsl:value-of select ="'ppaction://hlinkfile'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
      <xsl:for-each select ="超链:提示_AA05">
        <xsl:attribute name ="tooltip">
          <xsl:value-of select ="."/>
        </xsl:attribute>
      </xsl:for-each>

    </a:hlinkClick>
  </xsl:template>

  <xsl:template name="index">
    <xsl:param name="name"/>
    <xsl:for-each select="//演:幻灯片序列">
      <xsl:if test="@演:标识符=$name">
        <xsl:value-of select="position()"/>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>