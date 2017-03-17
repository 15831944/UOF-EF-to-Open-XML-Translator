<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
xmlns:元="http://schemas.uof.org/cn/2009/metadata"
xmlns:规则="http://schemas.uof.org/cn/2009/rules"
xmlns:式样="http://schemas.uof.org/cn/2009/styles"
xmlns:对象="http://schemas.uof.org/cn/2009/objects"
xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
xmlns:公式="http://schemas.uof.org/cn/2009/equations"
xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph"

xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:fo="http://www.w3.org/1999/XSL/Format"
xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
xmlns:dc="http://purl.org/dc/elements/1.1/"
xmlns:dcterms="http://purl.org/dc/terms/"
xmlns:dcmitype="http://purl.org/dc/dcmitype/"
xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
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

      <!--修改图片背景丢失  2013-4-1 liqiuling start-->	
		<xsl:if test="//w:media">
	<!--		<xsl:if test=" not(//mc:AlternateContent[mc:Choice/p:sp/p:txBody/a:p/a14:m/m:oMathPara]/mc:Fallback/p:sp/p:spPr/a:blipFill/a:blip)">-->
			<xsl:copy-of select="//w:media/*"/>
      <!--		</xsl:if>-->
		</xsl:if>
      <!--修改图片背景丢失  2013-4-1 liqiuling end-->
      <!--liuyangyang 2015-03-06 修改ole对象丢失 start-->
      <xsl:if test="//w:ole">
        <xsl:copy-of select="//w:ole/*"/>
      </xsl:if>
      <!--end liuyangyang 2015-03-06 修改ole对象丢失-->
      <!--元数据-->
		<pzip:entry pzip:target="content.xml">
			<演:演示文稿文档_6C10  xmlns:图="http://schemas.uof.org/cn/2009/graph">
				<演:母版集_6C0C>
					<xsl:apply-templates select="p:presentation/p:sldMaster" mode="ph"/>
					<xsl:apply-templates select="p:presentation/p:notesMaster" mode="ph"/>
					<xsl:apply-templates select="p:presentation/p:handoutMaster" mode="ph"/>
				</演:母版集_6C0C>
				<演:幻灯片集_6C0E>
					<xsl:apply-templates select="p:presentation/p:sld" mode="ph"/>
				</演:幻灯片集_6C0E>
			</演:演示文稿文档_6C10>
		</pzip:entry>
      
      <!--liuyangyang 2015-05-21 添加公式转换 start-->
      <pzip:entry pzip:target="equations.xml">
        <公式:公式集_C200>
          <xsl:for-each select="p:presentation/p:sld/p:cSld">
            <xsl:for-each select="./p:spTree/mc:AlternateContent">
              <xsl:if test=".//m:oMath">
                <公式:数学公式_C201>
                  <xsl:attribute name="标识符_C202">
                    <xsl:value-of select="./@id"/>
                  </xsl:attribute>
                  <xsl:copy-of select=".//m:oMath"/>
                </公式:数学公式_C201>
              </xsl:if>
            </xsl:for-each>
          </xsl:for-each>
        </公式:公式集_C200>
      </pzip:entry>
      <!--end   liuyangyang 2015-05-21 添加公式转换-->

      <pzip:entry pzip:target="_meta/meta.xml">
        <xsl:call-template name="metaData"/>
        <!---->
      </pzip:entry>
      <pzip:entry pzip:target="uof.xml">
        <uof:UOF_0000 xmlns:uof="http://schemas.uof.org/cn/2009/uof">
          <xsl:attribute name="mimetype_0001">vnd.uof.presentation</xsl:attribute>
          <xsl:attribute name="language_0002">cn</xsl:attribute>
          <xsl:attribute name="version_0003">2.0</xsl:attribute>
          <!--<xsl:if test="//w:media">
            <xsl:copy-of select="//w:media/*"/>
          </xsl:if>-->

        </uof:UOF_0000>
      </pzip:entry>

	
		
      <pzip:entry pzip:target="styles.xml">
        <式样:式样集_990B>
          <xsl:call-template name="styles"/>
          <!--<xsl:apply-templates select="//p:txStyles"/>-->
			<xsl:apply-templates select="p:presentation/p:sldMaster/p:txStyles">
				<xsl:with-param name="ID" select="p:presentation/p:sldMaster/@id"/>
			</xsl:apply-templates>
          <!--<xsl:call-template name="//p:txStyles"/>-->
          <!--
				  <xsl:apply-templates select="//p:txStyles"/-->
        </式样:式样集_990B>
      </pzip:entry>

		<!--李娟 添加 data 下的图片 信息  11.12.07-->

		
      <pzip:entry pzip:target="graphics.xml">

        <图形:图形集_7C00>
          <xsl:apply-templates select="//p:spTree" mode="sp"/>
        </图形:图形集_7C00>
      </pzip:entry>
      <pzip:entry pzip:target="theme.xml">
		  <规则:配色方案集_6C11>
		 <xsl:apply-templates select="//p:clrMap"/>
        </规则:配色方案集_6C11>
      </pzip:entry>
		<!--<xsl:if test ="//p:sld//p:pic  |//p:sld//a:blipFill | //p:notesMaster//a:blipFill | //p:handoutMaster//a:blipFill | //p:notesMaster//p:pic | //p:handoutMaster//p:pic |//p:sld//a:buBlip |//p:handoutMaster//a:buBlip  |//p:notesMaster//a:buBlip">-->
		<pzip:entry pzip:target="objectdata.xml">

				<对象:对象数据集_D700>
					<xsl:call-template name="object">
						<xsl:with-param name="picFrom" select="'slides'"/>
					</xsl:call-template>
          <xsl:call-template name="chartpic">
            <xsl:with-param name="picFrom" select="'slides'"/>
          </xsl:call-template>
					<!--<xsl:apply-templates select="//p:spTree" mode="sp"/>-->
				</对象:对象数据集_D700>
			</pzip:entry>
		<!--</xsl:if>-->
		<!--</xsl:if>-->
		<!--</xsl:if>-->
		<!--新添加批注集 李娟 2012.03.20-->
      <pzip:entry pzip:target="rules.xml">
        <规则:公用处理规则_B665 xmlns:uof="http://schemas.uof.org/cn/2009/uof">
          <规则:长度单位_B666>pt</规则:长度单位_B666>
			<xsl:if test="p:presentation/p:cmAuthorLst">
				<规则:用户集_B667>
					<xsl:for-each select=".//p:cmAuthor">
						<规则:用户_B668>
							<xsl:if test="@id">
								<xsl:attribute name="标识符_4100">
									<xsl:value-of select="@id"/>
								</xsl:attribute>
							</xsl:if>
							<xsl:if test="@name">
								<xsl:attribute name="姓名_41DC">
									<xsl:value-of select="@name"/>
								</xsl:attribute>
							</xsl:if>
						</规则:用户_B668>
					</xsl:for-each>
				</规则:用户集_B667>
			</xsl:if>
			<xsl:if test="p:presentation/p:cmLst">
				<规则:批注集_B669>
          <xsl:for-each select=".//p:cmLst">
            <xsl:variable name="cmSID" select="substring-before(@refBy,'.xml')"/>
            <xsl:for-each select=".//p:cm">
              <规则:批注_B66A>
                <xsl:variable name="authorID" select="@authorId"/>
                <!--<xsl:if test="@authorId">-->
                <xsl:attribute name="作者_41DD">
                  <xsl:value-of select="@authorId"/>
                </xsl:attribute>
                <!--</xsl:if>-->
                <xsl:if test="@dt">
                  <xsl:attribute name="日期_41DE">
                    <xsl:value-of select="@dt"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="../../p:cmAuthorLst/p:cmAuthor[@id=$authorID]/@initials">
                  <xsl:attribute name="作者缩写_41DF">
                    <xsl:value-of select="../../p:cmAuthorLst/p:cmAuthor[@id=$authorID]/@initials"/>
                  </xsl:attribute>
                </xsl:if>
                <!--2015.03.06 guoyongbin 批注只在第一第一张幻灯片显示，故为批注加区域引用编号-->
                <xsl:variable name="idx">
                  <xsl:value-of select="@idx"/>
                </xsl:variable>
                <xsl:attribute name="区域引用_41CE">
                  <xsl:value-of  select="concat('cmt_',$cmSID,'_',$idx)"/>
                </xsl:attribute>
                <!--end 2015.03.06 guoyongbin 批注只在第一张幻灯片显示，故为批注加区域引用编号-->
                <!--<xsl:variable name="pcm">
								<xsl:value-of select="count(preceding-sibling::p:cm)"/>
							</xsl:variable>
							-->
                <!--<xsl:if test="p:pos">-->
                <!--
								<xsl:attribute name="区域引用_41CE">
									<xsl:value-of  select="concat('cmt_',$pcm)"/>
								</xsl:attribute>-->
                <!--</xsl:if>-->
                <xsl:if test="p:text">
                  <字:段落_416B>
                    <字:段落属性_419B/>
                    <字:句_419D>
                      <字:文本串_415B>
                        <xsl:value-of select="p:text"/>
                      </字:文本串_415B>
                    </字:句_419D>
                  </字:段落_416B>
                </xsl:if>
              </规则:批注_B66A>
            </xsl:for-each>
          </xsl:for-each>
				</规则:批注集_B669>
			</xsl:if>
          <xsl:call-template name="CommenRule"/>
        </规则:公用处理规则_B665>
      </pzip:entry>
		
      
      <!-- 09.10.19 马有旭 添加 扩展区-->
      <pzip:entry pzip:target="extend.xml">
        <扩展:扩展区_B200>
          <!-- 09.11.19 马有旭 添加-->
          <xsl:call-template name="p:showPr">
            <xsl:with-param name="showPr" select="/p:presentation/p:presentationPr/p:showPr"/>
          </xsl:call-template>

          <!--2014-03-03, tangjiang, 添加OOXML到UOF用户自定义数据集的转换 start -->
          <xsl:if test="//w:document/w:customXML">
            <xsl:copy-of select="//w:document/w:customXML/*"/>
          </xsl:if>
          <!-- end -->
            
        </扩展:扩展区_B200>
      </pzip:entry>
		<xsl:if test="//a:hlinkClick | //a:hlinkMouseOver">
		  
		  <pzip:entry pzip:target="hyperlinks.xml">
			  <xsl:call-template name="hyperlinks"/>
		  </pzip:entry>
	  </xsl:if>
    </pzip:archive>
	  
	 
  </xsl:template>

	<!-- 李娟添加对象集-->
	<xsl:template name="object">
		<xsl:param name="picFrom"/>

      <!--start liuyin 20121204 修改“音视频丢失”-->
    <xsl:variable name="imageType" select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'"/>
    <xsl:variable name="audioType" select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'"/>
    <xsl:variable name="videoType" select="'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'"/>
    <xsl:variable name="mediaType" select="'http://schemas.microsoft.com/office/2007/relationships/media'"/>
    <xsl:for-each select=".//rel:Relationship[./@Type=$imageType or ./@Type=$audioType or ./@Type=$videoType  or ./@Type=$mediaType]">
      <!--end liuyin 20121204 修改“音视频丢失”-->
        
        <!--增加smartart中图片对象  2013-03-22 liqiuling start--> 
        		
			<xsl:variable name="slideName">
				<xsl:value-of select="substring-before(ancestor::rel:Relationships/@id, '.xml.rels')"/>
			</xsl:variable>
        
        <xsl:variable name="rel">
					<xsl:value-of select=".//rel:Relationship[@Type=$imageType]"/>
				</xsl:variable>
				<xsl:if test="not($rel=following::rel:Relationship/@Type)">
					<xsl:variable name="rID" select="./@Id"/>
					<xsl:variable name="target" select="./@Target"/>
          <xsl:variable name="fileObj" select="substring-after($target,'../media/')"/>
          <xsl:variable name="fileName" select="substring-before($fileObj,'.')"/>
          <xsl:variable name="fileType" select="substring-after($fileObj,'.')"/>
					<xsl:variable name="imageName" select="substring-before(substring-after($target,'../media/'),'.')"/>
          
          <!--增加smartart中图片对象  2013-03-22 liqiuling end-->
          
					<对象:对象数据_D701 是否内嵌_D705="true">
				<xsl:attribute name="标识符_D704">
          <xsl:value-of select="concat($slideName, $rID)"/>
				</xsl:attribute>
				<xsl:choose>
					<xsl:when test="contains($fileType,'emf')">
						<xsl:attribute name="私有类型_D707">
							<xsl:value-of select="$fileType"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="公共类型_D706">
							<xsl:value-of select="$fileType"/>
						</xsl:attribute>
					</xsl:otherwise>
				</xsl:choose>
				<xsl:element name ="对象:路径_D703">
          
          <!--start liuyin 20121204 修改“音视频丢失”-->
					<!--<xsl:value-of select ="concat('.\data\',$fileObj)"/>-->
          <xsl:value-of select ="concat('/data/',$fileObj)"/>
          <!--end liuyin 20121204 修改“音视频丢失”-->
        
				</xsl:element>
			</对象:对象数据_D701>
				</xsl:if>
				
				
				
			<!--<xsl:variable name="number">
			<xsl:number format="1" level="any" count="p:anchor"/>
		</xsl:variable>-->
			
		</xsl:for-each>
    <!--liuyangyang 2015-04-09 修复主题图片引用丢失 start-->
    <xsl:for-each select="/p:presentation/a:theme">
      <xsl:variable name="themeRef" select="concat(@id,'.rels')"/>
      <xsl:if test="./rel:Relationships">
        <xsl:for-each select="./rel:Relationships/rel:Relationship[./@Type=$imageType or ./@Type=$audioType or ./@Type=$videoType  or ./@Type=$mediaType]">
          <xsl:variable name="slideName">
            <xsl:value-of select="substring-before($themeRef, '.xml')"/>
          </xsl:variable>

          <xsl:variable name="rel">
            <xsl:value-of select=".//rel:Relationship[@Type=$imageType]"/>
          </xsl:variable>
          <xsl:if test="not($rel=following::rel:Relationship/@Type)">
            <xsl:variable name="rID" select="./@Id"/>
            <xsl:variable name="target" select="./@Target"/>
            <xsl:variable name="fileObj" select="substring-after($target,'../media/')"/>
            <xsl:variable name="fileName" select="substring-before($fileObj,'.')"/>
            <xsl:variable name="fileType" select="substring-after($fileObj,'.')"/>
            <xsl:variable name="imageName" select="substring-before(substring-after($target,'../media/'),'.')"/>

            <对象:对象数据_D701 是否内嵌_D705="true">
              <xsl:attribute name="标识符_D704">
                <xsl:value-of select="concat($slideName, $rID)"/>
              </xsl:attribute>
              <xsl:choose>
                <xsl:when test="contains($fileType,'emf')">
                  <xsl:attribute name="私有类型_D707">
                    <xsl:value-of select="$fileType"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="公共类型_D706">
                    <xsl:value-of select="$fileType"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
              <xsl:element name ="对象:路径_D703">
                <xsl:value-of select ="concat('/data/',$fileObj)"/>
              </xsl:element>
            </对象:对象数据_D701>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:for-each>
    <!--end liuyangyang 2015-04-09 修复主题图片引用丢失-->
	</xsl:template>
  <!--增加chart图片引用  liqiuling 2013-03-26 start-->
  <xsl:template name="chartpic">
    <xsl:param name="picFrom"/>

    <xsl:variable name="slideName">
      <xsl:value-of select=".//p:sld/@id"/>
    </xsl:variable>

    <xsl:for-each select=".//rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart']">
      <xsl:variable name="rel">
        <xsl:value-of select=".//rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart']"/>
      </xsl:variable>

      <xsl:if test="not($rel=following::rel:Relationship/@Type)">
        <xsl:variable name="rID">
          <xsl:value-of select="./@Id"/>
        </xsl:variable>
        <xsl:variable name="target">
          <xsl:value-of select="./@Target"/>
        </xsl:variable>
        <xsl:variable name="imageName">
          <xsl:value-of select="substring-after($target,'../charts/')"/>
        </xsl:variable>
        <对象:对象数据_D701 是否内嵌_D705="true">
          <xsl:attribute name="标识符_D704">
            <xsl:value-of select="substring-before($imageName,'.xml')"/>
          </xsl:attribute>
         
              <xsl:attribute name="公共类型_D706">
                    <xsl:value-of select="'jpg'"/>
              </xsl:attribute>
        
          <xsl:element name ="对象:路径_D703">
            
            <xsl:value-of select ="concat('/data/',$imageName,'.jpg')"/>
            
          </xsl:element>
        </对象:对象数据_D701>
      </xsl:if>

    </xsl:for-each>
  </xsl:template>

  <!--增加chart图片引用  liqiuling 2013-03-26 end -->
 
  <!-- 09.10.19 马有旭 添加 -->
  <xsl:template name="p:showPr">
    <xsl:param name="showPr"/>
    <扩展:扩展_B201>
      <扩展:软件名称_B202>EIOffice</扩展:软件名称_B202>
      <扩展:软件版本_B203>2011</扩展:软件版本_B203>
      <扩展:扩展内容_B204>
        <扩展:路径_B205>图形集/图形/预定义图形/属性/阴影</扩展:路径_B205>
        <扩展:内容_B206>
          <!--这部分怎么写，shema没定义-->
          <!--<uof:旁白声音 uof:locID="up001">
            <xsl:choose> 
              <xsl:when test="$showPr/@showNarration='0' or ($showPr and not($showPr/@showNarration))">false</xsl:when>                          
              <xsl:otherwise>true</xsl:otherwise>
            </xsl:choose>
          </uof:旁白声音>
          <xsl:if test="$showPr/* and not($showPr/p:present)">
            <uof:播放类型 uof:locID="up002" uof:attrList="类型" uof:类型="kiosk">
              <xsl:if test="$showPr/p:browse">
                <xsl:choose>
                  <xsl:when test="$showPr/p:browse/@showScrollbar='0'">
                    <xsl:attribute name="showScrollbar">0</xsl:attribute>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:attribute name="showScrollbar">1</xsl:attribute>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:if>
            </uof:播放类型>
          </xsl:if>
          -->
          <!-- 10.01.06 马有旭 <xsl:if test="document('ppt/presentation.xml')/p:presentation/@firstSlideNum">-->
          <!--
          <xsl:if test="/p:presentation/@firstSlideNum">
            <uof:幻灯片起始编号 uof:locID="up003">
              <xsl:value-of select="/p:presentation/@firstSlideNum"/>
            </uof:幻灯片起始编号>
          </xsl:if>
          -->
          <!--xsl:call-template name="hf">
            <xsl:with-param name="type">presentation</xsl:with-param>
          </xsl:call-template>
          <xsl:call-template name="hf">
            <xsl:with-param name="type">notesandhandout</xsl:with-param>
          </xsl:call-template>
          <xsl:call-template name="MasterSlidehf"/-->
        </扩展:内容_B206>
      </扩展:扩展内容_B204>
    </扩展:扩展_B201>
  </xsl:template>

  

  <xsl:template name="datetype">
    <xsl:param name="type"/>
    <xsl:attribute name="uof:格式索引">
      <xsl:choose>
        <xsl:when test="$type='datetime1'">1</xsl:when>
        <xsl:when test="$type='datetime2'">0</xsl:when>
        <xsl:when test="$type='datetime3'">10</xsl:when>
        <xsl:when test="$type='datetime5'">1</xsl:when>
        <xsl:when test="$type='datetime6'">8</xsl:when>
        <xsl:when test="$type='datetime7'">4</xsl:when>
        <xsl:when test="$type='datetime8'">11</xsl:when>
        <xsl:when test="$type='datetime9'">11</xsl:when>
        <xsl:when test="$type='datetime10'">20</xsl:when>
        <xsl:when test="$type='datetime11'">22</xsl:when>
        <xsl:when test="$type='datetime12'">23</xsl:when>
        <xsl:when test="$type='datetime13'">25</xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>
</xsl:stylesheet>





