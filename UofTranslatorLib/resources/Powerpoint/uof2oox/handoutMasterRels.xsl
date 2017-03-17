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
	
	<xsl:template name="handoutMasterRels">
    <!--handoutMaster-->
 <!--修改 李娟 2012.01.07-->
    <pzip:entry>
 
      <xsl:attribute name="pzip:target">
		  
        <xsl:choose>
			<xsl:when test="contains(@标识符_6BE8,'ID')">
				
            <xsl:value-of select="concat('ppt/handoutMasters/handoutMaster',substring-after(@标识符_6BE8,'ID'),'.xml')"/>
          </xsl:when>
          <xsl:otherwise>
			 
            <xsl:value-of select="concat('ppt/handoutMasters/',@标识符_6BE8,'.xml')"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <p:handoutMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        
        <xsl:call-template name="cSld"/>
        <!--可能有问题-->
        <xsl:call-template name="clrMap"/>
        <!--2010-11-8 罗文甜：修改页眉页脚-->
        <xsl:apply-templates select="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C[1]" mode="hf"/>
      <!-- 09.11.27 马有旭 >
        <xsl:apply-templates select="/uof:UOF/uof:扩展区/uof:扩展/uof:扩展内容[uof:路径='uof/演示文稿']/uof:内容/uof:页眉页脚[@uof:类型='notesandhandout']" mode="hf"/-->
      </p:handoutMaster>
    </pzip:entry>
    <!--handoutMasterRels-->
    <pzip:entry>
      <!-- 黎美秀更改 10.29      
       <xsl:attribute name="pzip:target">
        <xsl:value-of select="concat('ppt/handoutMasters/_rels/',@演:类型,@演:标识符,'.xml.rels')"/>
      </xsl:attribute>
      <xsl:call-template name="handoutMaster.xml.rels"/>
      -->
      <xsl:attribute name="pzip:target">
		  
		  <xsl:choose>
			  <xsl:when test="contains(@标识符_6BE8,'ID')">
				  <xsl:value-of select="concat('ppt/handoutMasters/_rels/handoutMaster',substring-after(@标识符_6BE8,'ID'),'.xml.rels')"/>
			  </xsl:when>
			  <xsl:otherwise>
				  <xsl:value-of select="concat('ppt/handoutMasters/_rels/',@标识符_6BE8,'.xml.rels')"/>
			  </xsl:otherwise>
		  </xsl:choose>
        
      </xsl:attribute>
      <xsl:call-template name="handoutMaster.xml.rels"/>
    </pzip:entry>
  </xsl:template>
  
  <xsl:template name="handoutMaster.xml.rels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme">
     
        <xsl:attribute name="Id">
          <xsl:choose>
			  <xsl:when test="@配色方案引用_6BEB">
				  <xsl:choose>
					  <xsl:when test="contains(@配色方案引用_6BEB,'clr_')">
						<xsl:value-of select="concat('rIdTheme',substring-after(@配色方案引用_6BEB,'clr_'))"/>
					</xsl:when>
					  <xsl:otherwise>
						<xsl:value-of select="concat('rIdTheme',substring-after(@配色方案引用_6BEB,'Id'))"/>
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
		<xsl:if test=".//图:图片数据引用_8037 or .//图:图片_8005/@图形引用_8007">
			<xsl:variable name ="picref" select=".//图:图片数据引用_8037|.//图:图片_8005/@图形引用_8007"/>
        <xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picref]">
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
            <xsl:attribute name="Id">
              <xsl:value-of select="concat('rId',@标识符_D704)"/>
            </xsl:attribute>
            <xsl:choose>
              <xsl:when test="@公共类型_D706">
                <xsl:attribute name="Target">
                  <xsl:variable name="objPath" select="./对象:路径_D703"/>
                  <xsl:choose>
                    <xsl:when test="contains($objPath, '/data/' )">
                      <xsl:value-of select="concat('../media/',substring-after($objPath,'/data/'))"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="concat('../media/',substring-after($objPath,'\data\'))"/>
                    </xsl:otherwise>
                  </xsl:choose>
                  <!--<xsl:value-of select="concat('../media/',@标识符_D704,'.',@公共类型_D706)"/>-->
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
