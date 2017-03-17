<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
xmlns:对象="http://schemas.uof.org/cn/2009/objects">
  <xsl:import href ="cSld.xsl"/>
  <xsl:import href="SlideRels.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8"/>
  <xsl:template name="NoteRels">
    <pzip:entry>
      <xsl:attribute name="pzip:target">
        <!--2月16日改 xsl:value-of select="concat('ppt/notesSlides/','note',parent::演:幻灯片/@演:标识符,'.xml')"/-->
        <xsl:value-of select="concat('ppt/notesSlides/','notesSlide',substring-after(parent::演:幻灯片_6C0F/演:幻灯片备注_6B1D/uof:锚点_C644/@图形引用_C62E,'Obj'),'.xml')"/>
      </xsl:attribute>
      <p:notes>

        <xsl:call-template name="cSld"/>

        <xsl:call-template name="clrMapOvr"/>

        <!--xsl:if test="演:切换">
          <xsl:call-template name="transition"/>
        </xsl:if>
        
        <xsl:if test="演:动画">
          <xsl:for-each select=".//演:动画">
            <xsl:call-template name="timing"/>
          </xsl:for-each>
        </xsl:if-->

      </p:notes>
    </pzip:entry>
    <pzip:entry>
      <xsl:attribute name="pzip:target">
        <!--2月16日改 xsl:value-of select="concat('ppt/notesSlides/_rels/','note',parent::演:幻灯片/@演:标识符,'.xml.rels')"/-->
        <xsl:value-of select="concat('ppt/notesSlides/_rels/','notesSlide',substring-after(parent::演:幻灯片_6C0F/演:幻灯片备注_6B1D/uof:锚点_C644/@图形引用_C62E,'Obj'),'.xml.rels')"/>
      </xsl:attribute>
      <xsl:call-template name="notesSlide.xml.rels"/>
    </pzip:entry>
  </xsl:template>

  <xsl:template name="notesSlide.xml.rels">
    <Relationships>
      <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/slide">
        <xsl:attribute name="Id">
          <xsl:value-of select="concat('rId',parent::演:幻灯片_6C0F/@标识符_6B0A)"/>
        </xsl:attribute>
        <xsl:attribute name="Target">
          <xsl:value-of select="concat('../slides/',parent::演:幻灯片_6C0F/@标识符_6B0A,'.xml')"/>
        </xsl:attribute>
      </Relationship>
      <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/notesMaster">

        <!--10.29 黎美秀修改 增加判断 -->
        <xsl:if test="not(//演:母版_6C0D[@类型_6BEA='notes']) and //演:幻灯片备注_6B1D">

          <xsl:attribute name="Id">
            <xsl:value-of select="concat('rIdnotsMaster',substring-after(//演:幻灯片_6C0F/演:幻灯片备注_6B1D/uof:锚点_C644/@图形引用_C62E,'Obj'))"/>

          </xsl:attribute>
          <xsl:attribute name="Target">
            <xsl:value-of select="concat('../notesMasters/notesMaster',substring-after(//演:幻灯片备注_6B1D/uof:锚点_C644/@图形引用_C62E,'Obj'),'.xml')"/>

          </xsl:attribute>

        </xsl:if>

        <xsl:for-each select="ancestor::uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='notes']">
          <!-- 09.10.13 马有旭 修改 -->
          <xsl:attribute name="Id">
            <!--xsl:value-of select="concat('rId',substring-after(.//@uof:图形引用,'OBJ'))"/-->
            <!--<xsl:value-of select="concat('rId',substring-after(@演:标识符,'ID'))"/>-->
            <xsl:choose>
              <xsl:when test="contains(@标识符_6BE8,'ID')">
                <xsl:value-of select="concat('rIdnotsMaster',substring-after(@标识符_6BE8,'ID'))"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="concat('rId',@标识符_6BE8)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name="Target">
            <!--xsl:value-of select="concat('../notesMasters/notesMaster',substring-after(.//@uof:图形引用,'OBJ'),'.xml')"/-->
            <!-- 09.10.13 马有旭 修改 -->
            <!--<xsl:value-of select="concat('../notesMasters/notesMaster',substring-after(@演:标识符,'ID'),'.xml')"/>-->
            <xsl:choose>
              <xsl:when test="contains(@标识符_6BE8,'ID')">
                <xsl:value-of select="concat('../notesMasters/notesMaster',substring-after(@标识符_6BE8,'ID'),'.xml')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="concat('../notesMasters/',@标识符_6BE8,'.xml')"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:for-each>
      </Relationship>
      <!--
      2010.4.7 黎美秀修改 增加超级链接  
      
      -->
      <xsl:for-each select=".//字:区域开始_4165[@类型_413B='hyperlink']">
        <!--<xsl:if test="not(current()/@字:标识符 = preceding::字:区域开始/@字:标识符)">-->
		  <xsl:if test="not(current()/@标识符_4169 = preceding::字:区域开始_4165/@标识符_4100)">
          <xsl:for-each select="//uof:链接集/超链:链接集_AA0B/超链:超级链接_AA0C[@超链:链源_AA00=current()/@标识符_4100]">
           
            <xsl:variable name="ctarget">
              <!--<xsl:value-of select="@uof:目标"/>-->
				<xsl:value-of select="超链:超级链接_AA0C/超链:目标_AA01"/>
            </xsl:variable>
            <xsl:variable name="cid">
              <xsl:value-of select="超链:超级链接_AA0C/@标识符_AA0A"/>
            </xsl:variable>
  
            <xsl:if test="not(starts-with(@超链:目标_AA01,'Custom Show:')) and ./@超链:目标_AA01!='First Slide' and ./@超链:目标_AA01!='Last Slide' and ./@超链:目标_AA01!='Previous Slide' and ./@超链:目标_AA01!='End Show' and ./@超链:目标_AA01!='Next Slide'">
              
              <Relationship>
                <xsl:attribute name="Type">
                  <xsl:choose>
                    <xsl:when test="starts-with(./@超链:目标_AA01,'Slide:')">http://purl.oclc.org/ooxml/officeDocument/relationships/slide</xsl:when>
                    <xsl:otherwise>http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink</xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
                <xsl:attribute name="Id">
                  <xsl:value-of select="超链:超级链接_AA0C/@标识符_AA0A"/>
                </xsl:attribute>
                <xsl:attribute name="Target">
                  <xsl:choose>
                    <xsl:when test="starts-with(@超链:目标_AA01,'Slide:')">
                      <xsl:value-of select="concat(substring-after(@超链:目标_AA01,'Slide:'),'.xml')"/>
                    </xsl:when>
                    <xsl:when test="not(starts-with(@超链:目标_AA01,'Slide:'))">
                      <xsl:variable name="target">
                        <xsl:choose>
                          <xsl:when test="contains(@超链:目标_AA01,' ')">
                            <xsl:call-template name="replace">
                              <xsl:with-param name="target" select="@超链:目标_AA01"/>
                            </xsl:call-template>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="@超链:目标_AA01"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:variable>
                      <xsl:choose>
                        <xsl:when test="contains(@超链:目标_AA01,':') and string-length(substring-before(@超链:目标_AA01,':'))=1">
                          <xsl:value-of select="concat('file:///',$target)"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$target"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                  </xsl:choose>
                </xsl:attribute>
                <xsl:if test="not(starts-with(@超链:目标_AA01,'Slide:'))">
                  <xsl:attribute name="TargetMode">External</xsl:attribute>
                </xsl:if>
              </Relationship>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:for-each>



		<xsl:if test=".//图:其他对象引用_8038 | .//图:图片数据引用_8037 | .//图:图片_8005/@图形引用_8007">
			
			<xsl:variable name="picref" select=".//图:其他对象引用_8038 |.//图:图片数据引用_8037|.//图:图片_8005/@图形引用_8007"/>
        <xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picref]">
			<xsl:variable name="obdata">
				<xsl:value-of select="@标识符_D704"/>
			</xsl:variable>
			<xsl:if test="not($obdata=following::对象:对象数据_D701/@标识符_D704)">
				<Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/image">

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
							</xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="Target">
								<xsl:value-of select="concat('../media/',@标识符_D704,'.','jpg')"/>
							</xsl:attribute>
						</xsl:otherwise>
					</xsl:choose>
				</Relationship>
			</xsl:if>
        </xsl:for-each>
      </xsl:if>
    </Relationships>
  </xsl:template>

</xsl:stylesheet>
