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
xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
xmlns:对象="http://schemas.uof.org/cn/2009/objects">
  <xsl:import href ="cSld.xsl"/>
  <xsl:import href="animation.xsl"/>
  <!--<xsl:import href ="metadata.xsl"/>
	<xsl:import href ="numbering.xsl"/>
	<xsl:import href ="region.xsl"/>-->
	<!--修改标签 李娟  2012.01.02-->
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes" standalone ="yes"/>
  <xsl:template name="slideMasterRels">

    <!--slideMaster-->
    <pzip:entry>
      <xsl:attribute name="pzip:target">
        <xsl:value-of select="concat('ppt/slideMasters/',@标识符_6BE8,'.xml')"/>
      </xsl:attribute>
      <p:sldMaster xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:p="http://purl.oclc.org/ooxml/presentationml/main">

        <xsl:call-template name="cSld"/>

        <xsl:call-template name="clrMap"/>
        
        <!--start liuyin 20130115 修改幻灯片母版引用文件需要修复才能打开-->
        
        <!--start liuyin 20130130 修改bug_2659,转换器出错-->
        <!--<xsl:if test ="@文本式样引用_6BED>-->
        
        <!--start liuyin 20130307 修改幻灯片母版引用等，转换后需要修复才能打开-->
        <!--<xsl:if test ="@文本式样引用_6BED and 演:页面版式引用">-->
          <xsl:if test ="演:页面版式引用">
        <!--end liuyin 20130307 修改幻灯片母版引用等，转换后需要修复才能打开-->
            
        <!--end liuyin 20130130 修改bug_2659,转换器出错-->
            
          <xsl:call-template name="sldLayoutIdLst"/>
        </xsl:if>
        <!--<xsl:call-template name="sldLayoutIdLst"/>-->
        <!--end liuyin 20130115 修改幻灯片母版引用文件需要修复才能打开-->
		        
       
        <!--2010-11-23 罗文甜:增加动画的判断-->
        <xsl:if test="演:动画_6B1A">
          <xsl:call-template name="animation"/>
        </xsl:if>
        <!--李娟 修改 页眉页脚 2012.02.22-->
		  
        <xsl:apply-templates select="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641[@类型_B645='presentation']" mode="hf"/>
		 
        <xsl:if test ="@文本式样引用_6BED">
          <xsl:call-template name="textStyles"/>
        </xsl:if>
      </p:sldMaster>
    </pzip:entry>

    <!--slideMasterRels-->
    <pzip:entry>
      <xsl:attribute name="pzip:target">
        <xsl:value-of select="concat('ppt/slideMasters/_rels/',@标识符_6BE8,'.xml.rels')"/>
      </xsl:attribute>
      <xsl:call-template name="slideMaster.xml.rels"/>
    </pzip:entry>
  </xsl:template>

  <xsl:template name="clrMap">
    <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
    
  </xsl:template>

	<!-- lijuan xiugai 2012 04 05-->
  <xsl:template name="sldLayoutIdLst" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:p="http://purl.oclc.org/ooxml/presentationml/main">
    <p:sldLayoutIdLst>
      
      <!--start liuyin 20130307 修改 段落式样中的对齐方式文档转换后需要修复-->
      <!--<xsl:for-each select="//演:幻灯片_6C0F">-->
     <xsl:for-each select="演:页面版式引用">
			<!--<xsl:if test="./@页面版式引用_6B27">-->
				<xsl:variable name="slideId">
					<xsl:value-of select="."/>
				</xsl:variable>
        <!--<xsl:for-each select="@页面版式引用_6B27[not(.=../following::演:幻灯片_6C0F/@页面版式引用)]">-->
			<!--<xsl:if test="not($slideId=following::演:幻灯片_6C0F/@页面版式引用_6B27)">-->
       <!--end liuyin 20130307 修改 段落式样中的对齐方式文档转换后需要修复-->
       
				<!--<xsl:variable name ="slideId">
					<xsl:value-of select="$refLayout"/>
				</xsl:variable>-->
			
			<!--<xsl:variable name="reflayout" select="$slideId[1]"></xsl:variable>-->
				<p:sldLayoutId>

					<!--<xsl:attribute name="id">
				<xsl:value-of select="2247483648 + substring-after($slideId,'slideLayout')"/>
			</xsl:attribute>
			<xsl:attribute name="r:id">
				<xsl:value-of select="concat('rId',$slideId)"/>
			</xsl:attribute>-->
					<!--注销这部分 李娟 2012.01.09-->
					<xsl:choose>
						<xsl:when test ="contains($slideId,'ID')">
							<xsl:attribute name="id">
								<xsl:value-of select="2247483648 + substring-after($slideId,'ID')"/>
							</xsl:attribute>
							<xsl:attribute name="r:id">
								<xsl:value-of select="concat('rId',substring-after($slideId,'ID'))"/>
							</xsl:attribute>
						</xsl:when>
						<xsl:when test="contains($slideId,'slideLayout')">
							<xsl:attribute name="id">
								<xsl:value-of select="2247483649 + substring-after($slideId,'slideLayout')"/>
							</xsl:attribute>
							<xsl:attribute name="r:id">
								<xsl:value-of select="concat('rId',$slideId)"/>
							</xsl:attribute>
						</xsl:when>
					</xsl:choose>
				</p:sldLayoutId>
       
       <!--start liuyin 20130307 修改 段落式样中的对齐方式文档转换后需要修复-->
			<!--</xsl:if>
			</xsl:if>-->
       <!--end liuyin 20130307 修改 段落式样中的对齐方式文档转换后需要修复-->
       
		</xsl:for-each>
    </p:sldLayoutIdLst>
  </xsl:template>
 
  
  <xsl:template name="textStyles">
    <p:txStyles>
      <p:titleStyle>
        <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:spcBef>
            <a:spcPct val="0"/>
          </a:spcBef>
          <a:buNone/>
          <a:defRPr sz="3600" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mj-lt"/>
            <a:ea typeface="+mj-ea"/>
            <a:cs typeface="+mj-cs"/>
          </a:defRPr>
        </a:lvl1pPr>
      </p:titleStyle>
      <p:bodyStyle>
        <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:spcBef>
            <a:spcPct val="20%"/>
          </a:spcBef>
          <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
          <a:buChar char="•"/>
          <a:defRPr sz="3200" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl1pPr>
        <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:spcBef>
            <a:spcPct val="20%"/>
          </a:spcBef>
          <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
          <a:buChar char="–"/>
          <a:defRPr sz="2800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl2pPr>
        <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:spcBef>
            <a:spcPct val="20%"/>
          </a:spcBef>
          <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
          <a:buChar char="•"/>
          <a:defRPr sz="2400" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl3pPr>
        <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:spcBef>
            <a:spcPct val="20%"/>
          </a:spcBef>
          <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
          <a:buChar char="–"/>
          <a:defRPr sz="2000" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl4pPr>
        <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:spcBef>
            <a:spcPct val="20%"/>
          </a:spcBef>
          <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
          <a:buChar char="»"/>
          <a:defRPr sz="2000" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl5pPr>
        <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:spcBef>
            <a:spcPct val="20%"/>
          </a:spcBef>
          <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
          <a:buChar char="•"/>
          <a:defRPr sz="2000" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl6pPr>
        <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:spcBef>
            <a:spcPct val="20%"/>
          </a:spcBef>
          <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
          <a:buChar char="•"/>
          <a:defRPr sz="2000" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl7pPr>
        <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:spcBef>
            <a:spcPct val="20%"/>
          </a:spcBef>
          <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
          <a:buChar char="•"/>
          <a:defRPr sz="2000" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl8pPr>
        <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:spcBef>
            <a:spcPct val="20%"/>
          </a:spcBef>
          <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
          <a:buChar char="•"/>
          <a:defRPr sz="2000" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl9pPr>
      </p:bodyStyle>
      <p:otherStyle>
        <a:defPPr>
          <a:defRPr lang="zh-CN"/>
        </a:defPPr>
        <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl1pPr>
        <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl2pPr>
        <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl3pPr>
        <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl4pPr>
        <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl5pPr>
        <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl6pPr>
        <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl7pPr>
        <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl8pPr>
        <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
          <a:defRPr sz="1800" kern="1200">
            <a:solidFill>
              <a:schemeClr val="tx1"/>
            </a:solidFill>
            <a:latin typeface="+mn-lt"/>
            <a:ea typeface="+mn-ea"/>
            <a:cs typeface="+mn-cs"/>
          </a:defRPr>
        </a:lvl9pPr>
      </p:otherStyle>
    </p:txStyles>
  </xsl:template>

  <xsl:template name="slideMaster.xml.rels">
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<!--是否增加判断 母版 是否存在页面版式应用属性暂时先不加  @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@李娟 2012.01.09-->
      <xsl:for-each select="//演:幻灯片_6C0F">
		  <xsl:if test="./@页面版式引用_6B27">
			  <xsl:variable name="slideId"><!--liuyin20130130-->
				  <xsl:value-of select="./@页面版式引用_6B27"/>
			  </xsl:variable>
			  <xsl:if test="not($slideId=following::演:幻灯片_6C0F/@页面版式引用_6B27)">
				  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout">
					  <xsl:attribute name="Id">
						  <xsl:value-of select="'rId'"/>
						  <xsl:choose>
							  <!--<xsl:when test ="contains($slideId,'0CLID')">
								  <xsl:value-of select="substring-after($slideId,'0CLID')"/>
							  </xsl:when>-->
							  <xsl:when test ="contains($slideId,'ID')">
								  <xsl:value-of select="substring-after($slideId,'ID')"/>
							  </xsl:when>
							  <!--4月15日蒋俊彦改-->
							  <!--处理空白版式-->
							  <xsl:when test="contains($slideId,'slideLayout')">
								  <xsl:value-of select="$slideId"/>
							  </xsl:when>
							  <xsl:otherwise>
								  <xsl:value-of select="$slideId"/>
							  </xsl:otherwise>
						  </xsl:choose>

					  </xsl:attribute>
					  <xsl:attribute name="Target">
						  <xsl:value-of select="concat('../slideLayouts/',$slideId,'.xml')"/>
					  </xsl:attribute>
				  </Relationship>
			  </xsl:if>
		  </xsl:if>
      </xsl:for-each>
      <!--12.4  黎美秀修改 增加超级链接-->
		<!--修改这块 李娟 2012.01.02-->
      <!--<xsl:for-each select=".//字:区域开始[@字:类型='hyperlink']">-->
		<xsl:for-each select=".//字:区域开始_4165[@类型_413B='hyperlink']">
			<xsl:if test="not(current()/@标识符_4169 = preceding::字:区域开始_4165/@标识符_4100)">
			<!--<xsl:if test="not(current()/@字:标识符 = preceding::字:区域开始/@字:标识符)">-->
				<!--<xsl:for-each select="//uof:链接集/uof:超级链接[@uof:链源=current()/@字:标识符]">-->
				<xsl:for-each select="//uof:链接集/超链:链接集_AA0B/超链:超级链接_AA0C[@超链:链源_AA00=current()/标识符_4100]">
					<xsl:variable name="ctarget">
						<xsl:value-of select="超链:目标_AA01"/>
					</xsl:variable>
					<xsl:variable name="cid">
						<xsl:value-of select="@标识符_AA0A"/>
					</xsl:variable>

					<xsl:if test="not(starts-with(超链:目标_AA01,'Custom Show:')) and ./超链:目标_AA01!='First Slide' and ./超链:目标_AA01!='Last Slide' and ./超链:目标_AA01!='Previous Slide' and ./超链:目标_AA01!='End Show' and ./超链:目标_AA01!='Next Slide'">
						<!--
              10.1.14 黎美秀修改
              <xsl:if test="not(//uof:超级链接[超链:目标_AA01=$ctarget and @标识符_AA0A!=$cid]) or //uof:超级链接[超链:目标_AA01=$ctarget][1]/@标识符_AA0A=$cid">
               </xsl:if>
              -->
						<Relationship>
							<xsl:attribute name="Type">
								<xsl:choose>
									<xsl:when test="starts-with(./超链:目标_AA01,'Slide:')">http://purl.oclc.org/ooxml/officeDocument/relationships/slide</xsl:when>
									<xsl:otherwise>http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink</xsl:otherwise>
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
				</xsl:for-each>
			</xsl:if>
		</xsl:for-each>

      <xsl:for-each select=".">
        <!-- 09.10.12 马有旭 修改
        <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/theme">
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
      </xsl:for-each>


      <!--   12.25 黎美秀增加 图片项目符号的引用-->

      <xsl:if test =".//@编号引用_4187=ancestor::uof:UOF/uof:式样集//字:自动编号_4124[字:级别_4112/字:图片符号_411B]/@标识符_4100">

        <xsl:variable name ="numref" select =".//@编号引用_4187"/>
        <xsl:for-each select ="//字:自动编号_4124[字:级别_4112/字:图片符号_411B]">

          <xsl:if test ="not(current()/字:级别_4112/字:图片符号_411B=preceding::*/字:图片符号_411B)">
            <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/image" Target="../media/image1.gif">
              <xsl:attribute name="Id">
                <xsl:value-of select="concat('rId',./字:级别_4112/字:图片符号_411B)"/>
              </xsl:attribute>
              <xsl:for-each select="//uof:其他对象/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=current()/字:级别_4112/字:图片符号_411B]">
                <xsl:attribute name="Target">
                  <xsl:value-of select="concat('../media/',@标识符_AA0A,'.',@公共类型_D706)"/>
                </xsl:attribute>
              </xsl:for-each>
            </Relationship>
          </xsl:if>

        </xsl:for-each>

      </xsl:if>

      <xsl:if test=".//图:图片数据引用_8037|.//图:图片_8005/@图形引用_8007">    
       <xsl:variable name ="picref" select=".//图:图片数据引用_8037|.//图:图片_8005/@图形引用_8007"/>
		  <xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picref]">
			  <xsl:variable name="obdata">
				  <xsl:value-of select="@标识符_D704"/>
			  </xsl:variable>
			  <xsl:if test="not($obdata=following::对象:对象数据_D701/@标识符_D704)">
				  <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/image">
					  <xsl:attribute name="Id">
						  <xsl:value-of select="concat('rId',@标识符_D704)"/>
					  </xsl:attribute>
					  <!--<xsl:choose>
						  <xsl:when test="@公共类型_D706">-->
							  <xsl:attribute name="Target">

                  <!--start liuyin 20121231 修改母版背景图片文件经转换后需要修复才能打开，打开后图片丢失-->
                  <xsl:variable name="objPath" select="./对象:路径_D703"/>
                  <xsl:choose>
                    <xsl:when test="contains($objPath, '/data/' )">
                      <xsl:value-of select="concat('../media/',substring-after($objPath,'/data/'))"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="concat('../media/',substring-after($objPath,'\data\'))"/>
                    </xsl:otherwise>
                  </xsl:choose>
                  <!--end liuyin 20121231 修改母版背景图片文件经转换后需要修复才能打开，打开后图片丢失-->
                  
							  </xsl:attribute>
						  <!--</xsl:when>
						  <xsl:otherwise>
							  <xsl:attribute name="Target">

								  <xsl:value-of select="concat('../media/',@标识符_D704,'.','jpg')"/>
							  </xsl:attribute>
						  </xsl:otherwise>
					  </xsl:choose>-->
				  </Relationship>
			  </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </Relationships>
  </xsl:template>
  <!--李娟 修改页眉页脚 2012.01.02-->
	<xsl:template match="规则:幻灯片_B641" mode="hf">
		<!--<xsl:template match="/uof:UOF/uof:演示文稿/演:公用处理规则/演:页眉页脚集/演:幻灯片页眉页脚" mode="hf">-->
		<xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641[@类型_B645='presentation']">
			<p:hf>
				<xsl:if test="not(@是否显示日期和时间_B647) or @是否显示日期和时间_B647='false' or (not(规则:日期和时间字符串_B643) and not(@是否自动更新日期和时间_B649)) or (规则:日期和时间字符串_B643='' and @是否自动更新日期和时间_B649='false')">
					<xsl:attribute name="dt">0</xsl:attribute>
				</xsl:if>
				<xsl:attribute name="hdr">0</xsl:attribute>
				<xsl:if test="not(@是否显示页脚_B648) or @是否显示页脚_B648='false' or not(规则:页脚_B644) or 规则:页脚_B644=''">
					<xsl:attribute name="ftr">0</xsl:attribute>
				</xsl:if>
				<xsl:if test="not(@是否显示幻灯片编号_B64A) or @是否显示幻灯片编号_B64A='false'">
					<xsl:attribute name="sldNum">0</xsl:attribute>
				</xsl:if>
			</p:hf>
		</xsl:if>
	</xsl:template>
  <xsl:template match="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:讲义和备注_B64C" mode="hf">
    <p:hf>
      <xsl:if test="not(@是否显示日期和时间_B647) or @是否显示日期和时间_B647='false' or not(@是否自动更新日期和时间_B649) or @是否自动更新日期和时间_B649='false'">
        <xsl:attribute name="dt">0</xsl:attribute>
      </xsl:if>
      <xsl:if test="not(@是否显示页眉_B64F) or @是否显示页眉_B64F='false' or not(规则:页眉_B64D) or 规则:页眉_B64D=''">
        <xsl:attribute name="hdr">0</xsl:attribute>
      </xsl:if>      
      <xsl:if test="not(@是否显示页脚_B648) or @是否显示页脚_B648='false' or not(规则:页脚_B644) or 规则:页脚_B644=''">
        <xsl:attribute name="ftr">0</xsl:attribute>
      </xsl:if>
      <xsl:if test="not(@是否显示页码_B650) or @是否显示页码_B650='false'">
        <xsl:attribute name="sldNum">0</xsl:attribute>
      </xsl:if>
    </p:hf>
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

