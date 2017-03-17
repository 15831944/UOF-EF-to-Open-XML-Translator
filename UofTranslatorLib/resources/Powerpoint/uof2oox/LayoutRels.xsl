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
xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
				xmlns:对象="http://schemas.uof.org/cn/2009/objects">
  <xsl:import href ="cSld.xsl"/>
<!--  对于字：标识符这块不太确认  李娟 2012.01.04-->
  <!--<xsl:import href ="metadata.xsl"/>
	<xsl:import href ="numbering.xsl"/>
	<xsl:import href ="region.xsl"/>-->
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes" standalone ="yes"/>
  <xsl:template name="LayoutRels">
	 
    <pzip:entry>
      <xsl:attribute name="pzip:target">
        <xsl:value-of select="concat('ppt/slideLayouts/',@标识符_6B0D,'.xml')"/>
      </xsl:attribute>
      <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" preserve="1">
		 <xsl:call-template name="cSld"/>
        <xsl:call-template name="clrMapOvr"/>
        
      </p:sldLayout>
    </pzip:entry>
    <pzip:entry>
      <xsl:attribute name="pzip:target">
        <xsl:value-of select="concat('ppt/slideLayouts/_rels/',@标识符_6B0D,'.xml.rels')"/>
      </xsl:attribute>
		
		<xsl:call-template name="slideLayout.xml.rels">
			
		</xsl:call-template>
    </pzip:entry>
  </xsl:template>
  <xsl:template name="clrMapOvr">
    <p:clrMapOvr>
      <a:masterClrMapping/>
    </p:clrMapOvr>
  </xsl:template>
  <xsl:template name="slideLayout.xml.rels">
	  <xsl:variable name="layoutrel" select="@标识符_6B0D"/>    
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <!--为使多母版情况可以打开，修改此变量值 丽娟 5月18-->
		<xsl:variable name="masterRef">
      
      <!--start liuyin 20130115 修改幻灯片母版引用文件需要修复才能打开,此处还可能存在bug，但为了解决bug_2591只能这么写了-->

      <!--start liuyin 20130316 修改bug2771需要修复才能打开-->
      <xsl:choose>
        <xsl:when test=" $layoutrel ">
          <xsl:value-of select="//演:幻灯片_6C0F[@页面版式引用_6B27=$layoutrel]/@母版引用_6B26"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="//演:母版_6C0D[1]/@标识符_6BE8"/>
        </xsl:otherwise>
      </xsl:choose>
      <!--<xsl:value-of select="//演:母版_6C0D[1]/@标识符_6BE8"/>-->
      <!--end liuyin 20130316 修改bug2771需要修复才能打开-->
      
			<!--<xsl:value-of select="//演:幻灯片_6C0F[@页面版式引用_6B27=$layoutrel]/@母版引用_6B26"/>-->
      <!--end liuyin 20130115 修改幻灯片母版引用文件需要修复才能打开-->
      
		</xsl:variable>
		
      <Relationship Id="rId1"  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster">
        <xsl:attribute name="Id">
		 <xsl:value-of select="concat('rId',$masterRef)"/>
        </xsl:attribute>
        <xsl:attribute name="Target">
          <xsl:value-of select="concat('../slideMasters/',$masterRef,'.xml')"/>
        </xsl:attribute>
		  <!--<xsl:attribute name="Id">
			  <xsl:value-of select="concat('rId',master[1]/@masterRef)"/>
		  </xsl:attribute>
		  <xsl:attribute name="Target">
			  <xsl:value-of select="concat('../slideMasters/',master[1]/@masterRef,'.xml')"/>
		  </xsl:attribute>-->
      </Relationship>
      <!--12.4  黎美秀修改 增加超级链接-->
      <xsl:for-each select="//字:区域开始_4165[@类型_413B='hyperlink']">
        <xsl:if test="not(current()/@标识符_4100 = preceding::字:区域开始_4165/@标识符_4100)">
          <xsl:for-each select="//uof:链接集/超链:链接集_AA0B/超链:超级链接_AA0C[超链:链源_AA00=current()/@标识符_4100]">

            <xsl:variable name="ctarget">
              <xsl:value-of select="超链:目标_AA01"/>
            </xsl:variable>
            <xsl:variable name="cid">
              <xsl:value-of select="@标识符_AA0A"/>
            </xsl:variable>
            <xsl:if test="$ctarget!=''">
              <xsl:if test="not(starts-with(超链:目标_AA01,'Custom Show:')) and ./超链:目标_AA01!='First Slide' and ./超链:目标_AA01!='Last Slide' and ./超链:目标_AA01!='Previous Slide' and ./超链:目标_AA01!='End Show' and ./超链:目标_AA01!='Next Slide'">
                <!--
           10.1.14 黎美秀修改
           <xsl:if test="not(//uof:超级链接[@uof:目标=$ctarget and @uof:标识符!=$cid]) or //uof:超级链接[@uof:目标=$ctarget][1]/@uof:标识符=$cid">
              
            </xsl:if>
           -->
                <Relationship>
                  <xsl:attribute name="Type">
                    <xsl:choose>
                      <xsl:when test="starts-with(./超链:目标_AA01,'Slide:')">http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide</xsl:when>
                      <xsl:otherwise>http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink</xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                  <xsl:attribute name="Id">
                    <xsl:value-of select="@标识符_AA0A"/>
                  </xsl:attribute>
                  <xsl:attribute name="Target">
                    <xsl:choose>
                      <xsl:when test="starts-with(超链:目标_AA01,'Slide:')">
                        <xsl:value-of select="concat(substring-after(超链:目标_AA01,'Slide:'),'.xml')"/>
                      </xsl:when>
                      <xsl:when test="not(starts-with(超链:目标_AA01,'Slide:'))">
                        <xsl:variable name="target">
                          <xsl:choose>
                            <xsl:when test="contains(超链:目标_AA01,' ')">
                              <xsl:call-template name="replace">
                                <xsl:with-param name="target" select="超链:目标_AA01"/>
                              </xsl:call-template>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="超链:目标_AA01"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:variable>
                        <xsl:choose>
                          <xsl:when test="contains(超链:目标_AA01,':') and string-length(substring-before(超链:目标_AA01,':'))=1">
                            <xsl:value-of select="concat('file:///',$target)"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="$target"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:attribute>
                  <xsl:if test="not(starts-with(超链:目标_AA01,'Slide:'))">
                    <xsl:attribute name="TargetMode">External</xsl:attribute>
                  </xsl:if>
                </Relationship>
              </xsl:if>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
      </xsl:for-each>

		<xsl:if test=".//图:图片数据引用_8037 |.//图:图片_8005/@图形引用_8007 ">
			<!--<xsl:variable name="picref" select=".//@图:其他对象| .//@图:图形引用"/>-->
			<xsl:variable name="picref" select=".//图:图片数据引用_8037|.//图:图片_8005/@图形引用_8007"/>
			<xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picref]">
				<xsl:variable name="obdata">
					<xsl:value-of select="@标识符_D704"/>
				</xsl:variable>
				<xsl:if test="not($obdata=following::对象:对象数据_D701/@标识符_D704)">
					<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
						<!--10.23 黎美秀修改 此处存在问题，转出来的格式可能名称为image -->
						<xsl:attribute name="Id">
							<!--xsl:value-of select="concat('rId',substring-after(@uof:标识符,'OBJ'))"/>  -->
							<xsl:value-of select="concat('rId',@标识符_D704)"/>
						</xsl:attribute>
						<xsl:choose>
							<xsl:when test="@公共类型_D706 or @私有类型_D707 ">
								<xsl:attribute name="Target">
                  <!--修复互操作需要修复  liqiuling  2013-5-10 start-->
                  <xsl:variable name="objPath" select="./对象:路径_D703"/>
                  <xsl:choose>
                    <xsl:when test="contains($objPath, '/data/' )">
                      <xsl:value-of select="concat('../media/',substring-after($objPath,'/data/'))"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="concat('../media/',substring-after($objPath,'\data\'))"/>
                    </xsl:otherwise>
                  </xsl:choose>
                  <!--修复互操作需要修复  liqiuling  2013-5-10 end-->
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

		<!-- 09.12.19 马有旭 添加 -->
        <xsl:for-each select="ancestor::uof:UOF/uof:对象集/uof:其他对象/对象:对象数据集_D700/对象:对象数据_D701">
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
            <xsl:attribute name="Id">
              
              <!--
              12.25 黎美秀修改
              
                <xsl:if test="contains(@uof:标识符,'OBJ')">
                <xsl:value-of select="concat('rId',substring-after(@uof:标识符,'OBJ'))"/>
              </xsl:if>
              <xsl:if test="not(contains(@uof:标识符,'OBJ'))">
                <xsl:value-of select="concat('rId',@uof:标识符)"/>
              </xsl:if>
              
              -->
              <xsl:value-of select="concat('rId',@标识符_D704)"/>
             
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
      
      <!--start liuyin 20121203 修改音视频丢失-->
      <xsl:if test=".//图:其他对象引用_8038">
        <xsl:variable name="picref" select=".//图:其他对象引用_8038"/>
        <xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picref]">
          <xsl:choose>
            <xsl:when test="@公共类型_D706='mp3'">
              <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio">
                <xsl:attribute name="Id">
                  <!--xsl:value-of select="concat('rId',substring-after(@uof:标识符,'OBJ'))"/>  -->
                  <xsl:value-of select="concat('rId',@标识符_D704)"/>
                </xsl:attribute>
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
              </Relationship>
              <Relationship Type="http://schemas.microsoft.com/office/2007/relationships/media">
                <xsl:attribute name="Id">
                  <!--xsl:value-of select="concat('rId',substring-after(@uof:标识符,'OBJ'))"/>  -->
                  <xsl:value-of select="concat('rId',@标识符_D704,'1')"/>
                </xsl:attribute>
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
              </Relationship>
            </xsl:when>
            <xsl:when test="@公共类型_D706='avi'">
              <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video">
                <xsl:attribute name="Id">
                  <!--xsl:value-of select="concat('rId',substring-after(@uof:标识符,'OBJ'))"/>  -->
                  <xsl:value-of select="concat('rId',@标识符_D704)"/>
                </xsl:attribute>
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
              </Relationship>
              <Relationship Type="http://schemas.microsoft.com/office/2007/relationships/media">
                <xsl:attribute name="Id">
                  <!--xsl:value-of select="concat('rId',substring-after(@uof:标识符,'OBJ'))"/>  -->
                  <xsl:value-of select="concat('rId',@标识符_D704,'1')"/>
                </xsl:attribute>
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
              </Relationship>
            </xsl:when>
            <xsl:when test="@公共类型_D706='wmv'">
              <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video">
                <xsl:attribute name="Id">
                  <!--xsl:value-of select="concat('rId',substring-after(@uof:标识符,'OBJ'))"/>  -->
                  <xsl:value-of select="concat('rId',@标识符_D704)"/>
                </xsl:attribute>
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
              </Relationship>
              <Relationship Type="http://schemas.microsoft.com/office/2007/relationships/media">
                <xsl:attribute name="Id">
                  <!--xsl:value-of select="concat('rId',substring-after(@uof:标识符,'OBJ'))"/>  -->
                  <xsl:value-of select="concat('rId',@标识符_D704,'1')"/>
                </xsl:attribute>
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
              </Relationship>
            </xsl:when>
          </xsl:choose>
        </xsl:for-each>
      </xsl:if>
      <!--end liuyin 20121203 修改音视频丢失-->
      
    </Relationships>
  </xsl:template>
  <xsl:template name="replace">
    <xsl:param name="target"/>
    <xsl:choose>
      <xsl:when test="contains($target,' ')">
        <xsl:call-template name="replace">
          <xsl:with-param name="target">
            <xsl:value-of select="concat(substring-before($target,' '),'%20',substring-after($target,' '))"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$target"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>