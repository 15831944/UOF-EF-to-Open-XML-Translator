<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
  xmlns:pzip="urn:u2o:xmlns:post-processings:special"
  xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
  xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
  xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
 xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph"
xmlns:规则="http://schemas.uof.org/cn/2009/rules"
xmlns:对象="http://schemas.uof.org/cn/2009/objects">
  <xsl:import href ="cSld.xsl"/>
  <xsl:import href="slideMasterRels.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes" standalone ="yes"/>
<!--修改标签 李娟 11.12.30-->
  <xsl:template name="noteMasterRels1">

    <!--noteMaster1-->
    <pzip:entry>
      <xsl:attribute name="pzip:target">
        <xsl:value-of select="concat('ppt/notesMasters/notesMaster',substring-after(.//@图形引用_C62E,'Obj'),'.xml')"/>
      </xsl:attribute>
      <p:notesMaster xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:p="http://purl.oclc.org/ooxml/presentationml/main">

        <xsl:call-template name="cSld"/>
        <!--可能有问题-->
        <xsl:call-template name="clrMap"/>
        <!--2010-11-8 罗文甜：修改页眉页脚-->
        <!--<xsl:apply-templates select="/uof:UOF/uof:演示文稿/演:公用处理规则/演:页眉页脚集/演:讲义和备注页眉页脚" mode="hf"/>-->
		  <xsl:apply-templates select="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C[1]" mode="hf"/>
        <!-- 09.11.27 马有旭 >
        <xsl:apply-templates select="/uof:UOF/uof:扩展区/uof:扩展/uof:扩展内容[uof:路径='uof/演示文稿']/uof:内容/uof:页眉页脚[@uof:类型='notesandhandout']" mode="hf"/-->
        <xsl:call-template name="notesStyle"/>
      </p:notesMaster>
    </pzip:entry>

    <pzip:entry>
      <xsl:attribute name="pzip:target">
        <xsl:value-of select="concat('ppt/notesMasters/_rels/notesMaster',substring-after(.//@图形引用_C62E,'Obj'),'.xml.rels')"/>
      </xsl:attribute>
      <xsl:call-template name="noteMaster1.xml.rels"/>
    </pzip:entry>
  </xsl:template>

  <xsl:template name="notesStyle">
    
  </xsl:template>

  <xsl:template name="noteMaster1.xml.rels">
    <xsl:for-each select="ancestor::uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='slide']">
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <!-- 09.10.12 马有旭 修改
      <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/theme">
        <xsl:attribute name="Id">
          <xsl:value-of select="'rId'"/>
          <xsl:value-of select="substring-after(@演:配色方案引用,'ID')"/>
        </xsl:attribute>
        <xsl:attribute name="Target">
          <xsl:value-of select="concat('../theme/theme',substring-after(@演:配色方案引用,'ID'),'.xml')"/>
        </xsl:attribute>
      </Relationship>-->
        <Relationship Type="http://purl.oclc.org/ooxml/officeDocument/relationships/theme">
          <xsl:attribute name="Id">
            <xsl:choose>
              <xsl:when test="contains(@配色方案引用_6BEB,'clr_')">
                <xsl:value-of select="concat('rIdTheme',substring-after(@配色方案引用_6BEB,'clr_'))"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="concat('rIdTheme',substring-after(@配色方案引用_6BEB,'Id'))"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>

          <xsl:attribute name="Target">
            <xsl:choose>
              <xsl:when test="contains(@配色方案引用_6BEB,'clr_')">
                <xsl:value-of select="concat('../theme/theme',substring-after(@配色方案引用_6BEB,'clr_'),'.xml')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="concat('../theme/theme',substring-after(@配色方案引用_6BEB,'Id'),'.xml')"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </Relationship>
        <xsl:if test=".//@图:图形引用_8007 or .//图:图片_8005">
          <xsl:for-each select="ancestor::uof:UOF/uof:对象集/uof:其他对象/对象:对象数据集_D700/对象:对象数据_D701">
            <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/image">
              <xsl:attribute name="Id">
                <xsl:value-of select="concat('rId',substring-after(@标识符_D704,'Obj'))"/>
              </xsl:attribute>
              <xsl:choose>
                <xsl:when test="@公共类型_D706">
                  <xsl:attribute name="Target">
                    <xsl:value-of select="concat('../media/',@标识符_D704,'.',@公共类型_D706)"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="Target">
                    <xsl:value-of select="concat('../media/',@标识符_D704,'.','jpg')"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </Relationship>
          </xsl:for-each>
        </xsl:if>
      </Relationships>
    </xsl:for-each>

  </xsl:template>

</xsl:stylesheet>
