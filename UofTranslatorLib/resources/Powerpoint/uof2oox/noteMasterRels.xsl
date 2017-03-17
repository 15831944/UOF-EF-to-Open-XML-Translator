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
xmlns:规则="http://schemas.uof.org/cn/2009/rules"
	xmlns:对象="http://schemas.uof.org/cn/2009/objects">
  <xsl:import href ="cSld.xsl"/>
  <xsl:import href="slideMasterRels.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes" standalone ="yes"/>
  <!--李娟 修改标签 2012.01.03-->
  <xsl:template name="noteMasterRels">
    <xsl:param name="idhf"/>

    <!--noteMaster-->
    <pzip:entry>
      <xsl:attribute name="pzip:target">
        <!-- 09.10.12 马有旭 修改
        <xsl:value-of select="concat('ppt/notesMasters/notesMaster',substring-after(@演:标识符,'ID'),'.xml')"/>
        -->
        <xsl:choose>
          <xsl:when test="contains(@标识符_6BE8,'ID')">
            <xsl:value-of select="concat('ppt/notesMasters/notesMaster',substring-after(@标识符_6BE8,'ID'),'.xml')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('ppt/notesMasters/',@标识符_6BE8,'.xml')"/>
          </xsl:otherwise>
        </xsl:choose>
        <!--2月23日改-->
        <!--xsl:choose>
        <xsl:if test="@演:类型='notes'">
            <xsl:value-of select="concat('ppt/notesMasters/notesMaster',substring-after(@演:标识符,'ID'),'.xml')"/>
          </xsl:if>
        <xsl:if test="not(@演:类型='notes') and ancestor::uof:演示文稿/演:主体/演:幻灯片集/演:幻灯片/演:幻灯片备注">
            <xsl:for-each select="ancestor::uof:演示文稿/演:主体/演:幻灯片集/演:幻灯片/演:幻灯片备注">
              <xsl:value-of select="concat('ppt/notesMasters/notesMaster',substring-after(.//@uof:图形引用,'OBJ'),'.xml')"/>
            </xsl:for-each>
          </xsl:if>
        </xsl:choose-->
      </xsl:attribute>
      <p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">

        <xsl:call-template name="cSld"/>
        <!--可能有问题-->
        <xsl:call-template name="clrMap"/>
        <!--2010-11-8 罗文甜：修改页眉页脚-->
        <xsl:apply-templates select="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C[1]" mode="hf"/>
        <!-- 09.11.27 马有旭 >
        <xsl:apply-templates select="/uof:UOF/uof:扩展区/uof:扩展/uof:扩展内容[uof:路径='uof/演示文稿']/uof:内容/uof:页眉页脚[@uof:类型='notesandhandout']" mode="hf"/>
        <xsl:call-template name="notesStyle"/-->
      </p:notesMaster>
    </pzip:entry>

    <pzip:entry>
      <xsl:attribute name="pzip:target">
        <!-- 09.10.13 马有旭 修改
        <xsl:value-of select="concat('ppt/notesMasters/_rels/notesMaster',substring-after(@演:标识符,'ID'),'.xml.rels')"/>
        -->
        <xsl:choose>
          <xsl:when test="contains(@标识符_6BE8,'ID')">

            <xsl:value-of select="concat('ppt/notesMasters/_rels/notesMaster',substring-after(@标识符_6BE8,'ID'),'.xml.rels')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('ppt/notesMasters/_rels/',@标识符_6BE8,'.xml.rels')"/>
          </xsl:otherwise>
        </xsl:choose>
        <!--xsl:choose>
        <xsl:when test="@演:类型='notes'">
            <xsl:value-of select="concat('ppt/notesMasters/_rels/notesMaster',substring-after(@演:标识符,'ID'),'.xml.rels')"/>
          </xsl:when>
          <xsl:when test="not(@演:类型='notes') and ancestor::uof:演示文稿/演:主体/演:幻灯片集/演:幻灯片/演:幻灯片备注">
            <xsl:for-each select="ancestor::uof:演示文稿/演:主体/演:幻灯片集/演:幻灯片/演:幻灯片备注">
              <xsl:value-of select="concat('ppt/notesMasters/_rels/notesMaster',substring-after(.//@uof:图形引用,'OBJ'),'.xml.rels')"/>
            </xsl:for-each>
          </xsl:when>
        </xsl:choose-->
      </xsl:attribute>
      <xsl:call-template name="noteMaster.xml.rels"/>
    </pzip:entry>
  </xsl:template>

  

  <xsl:template name="noteMaster.xml.rels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <!-- 09.10.12  马有旭 修改
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme">
        <xsl:attribute name="Id">
          <xsl:value-of select="'rId'"/>
          <xsl:value-of select="substring-after(@演:配色方案引用,'ID')"/>
        </xsl:attribute>
        <xsl:attribute name="Target">
          <xsl:value-of select="concat('../theme/theme',substring-after(@演:配色方案引用,'ID'),'.xml')"/>
        </xsl:attribute>
      </Relationship>
      -->
      <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme">
        <xsl:attribute name="Id">
			<xsl:choose>
				<xsl:when test="@配色方案引用_6BEB">
					<xsl:choose>
						<xsl:when test="contains(@配色方案引用_6BEB,'clr_')">
							<xsl:value-of select="concat('rIdTheme',substring-after(@配色方案引用_6BEB,'clr_'),'.xml')"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="concat('rIdTheme',substring-after(@配色方案引用_6BEB,'Id'),'.xml')"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="concat('rIdTheme',substring-after(@标识符_6BE8,'ID'),'.xml')"/>
				</xsl:otherwise>


			</xsl:choose>
        </xsl:attribute>

        <xsl:attribute name="Target">
			<xsl:choose>
				<xsl:when test="@配色方案引用_6BEB">
					<xsl:choose>
						<xsl:when test="contains(@配色方案引用_6BEB,'clr_')">
							<xsl:value-of select="concat('../theme/theme',substring-after(@配色方案引用_6BEB,'clr_'),'.xml')"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="concat('../theme/theme',substring-after(@配色方案引用_6BEB,'Id'),'.xml')"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="concat('../theme/theme',substring-after(@标识符_6BE8,'ID'),'.xml')"/>
				</xsl:otherwise>


			</xsl:choose>
        </xsl:attribute>
      </Relationship>
      <!--<xsl:if test=".//@图:其他对象 or .//图:图片">-->
		<xsl:if test=".//图:其他对象引用_8038 or .//图:图片数据引用_8037">
        <xsl:for-each select="ancestor::uof:UOF/uof:对象集/uof:其他对象/对象:对象数据集_D700/对象:对象数据_D701">
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
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
  </xsl:template>

</xsl:stylesheet>
