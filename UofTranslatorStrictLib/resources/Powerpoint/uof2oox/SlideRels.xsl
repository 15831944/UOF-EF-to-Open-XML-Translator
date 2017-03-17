<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xsl:stylesheet version="1.0" xmlns:pzip="urn:u2o:xmlns:post-processings:special"
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
xmlns:对象="http://schemas.uof.org/cn/2009/objects"
xmlns:规则="http://schemas.uof.org/cn/2009/rules">
  <xsl:import href="cSld.xsl"/>
  <xsl:import href="animation.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8"/>
	<!--修改标签 李娟11.12。31-->
  <xsl:template name="SlideRels">
    <pzip:entry>
      <xsl:attribute name="pzip:target">
        <xsl:value-of select="concat('ppt/slides/',@标识符_6B0A,'.xml')"/>
		  <!--<xsl:value-of select="concat('ppt/slides/',@演:标识符,'.xml')"/>-->
      </xsl:attribute>
      <p:sld xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:p="http://purl.oclc.org/ooxml/presentationml/main">
        <!--4月3日忽略母版背景-->
        <xsl:if test="@是否显示背景对象_6B2A='false'">
          <xsl:attribute name="showMasterSp">0</xsl:attribute>
        </xsl:if>
        <!-- 09.10.20 马有旭 幻灯片隐藏 -->
        <xsl:if test="@是否显示_6B28='false'">
          <xsl:attribute name="show">0</xsl:attribute>
        </xsl:if>
        <xsl:call-template name="cSld"/>
        <xsl:call-template name="clrMapOvr"/>
        <xsl:if test="演:切换_6B1F">
          <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
            <mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">
              <xsl:call-template name="transition"/>
            </mc:Choice>
          </mc:AlternateContent>
        </xsl:if>

        <!--start 2013-1-7 liuyin 音视频、幻灯片动画-->
        <xsl:if test="演:动画_6B1A and uof:锚点_C644/图:图形_8062/图:其他对象引用_8038">
          <xsl:call-template name="animation_other"/>
        </xsl:if>
        <!--end 2013-1-7 liuyin 音视频、幻灯片动画-->

        <!--start 2013-1-10 liuyin 音视频修复-->
        <!--2010-11-23 罗文甜：增加判断-->
        <xsl:if test="演:动画_6B1A and not(uof:锚点_C644/图:图形_8062/图:其他对象引用_8038)">
          <xsl:call-template name="animation"/>
        </xsl:if>
        <!--end 2013-1-10 liuyin 音视频修复-->
        
        <!--2011-2-18 罗文甜：增加母版动画-->
        <!--<xsl:if test="not(演:动画) and ../../演:母版集/演:母版[@演:类型='slide']/演:动画">-->
		  <xsl:if test="not(演:动画_6B1A) and ../../演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='slide']/演:动画_6B1A">
          <p:timing>
            <p:tnLst>
              <p:par>
                <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"/>
              </p:par>
            </p:tnLst>
          </p:timing>
        </xsl:if>
      </p:sld>
    </pzip:entry>
    <pzip:entry>
      <xsl:attribute name="pzip:target">
        <xsl:value-of select="concat('ppt/slides/_rels/',@标识符_6B0A,'.xml.rels')"/>
      </xsl:attribute>
      <xsl:call-template name="slide.xml.rels"/>
    </pzip:entry>
  </xsl:template>
  <xsl:template name="clrMapOvr">
    <p:clrMapOvr>
      <a:masterClrMapping/>
    </p:clrMapOvr>
  </xsl:template>
  <!--飞入、缓慢进入-->
  <!--随机效果-->
  <!--幻灯片切换-->
  <xsl:template name="transition">
    <p:transition>
      <xsl:if test="演:切换_6B1F/演:速度_6B21!='fast'">
        <xsl:attribute name="spd">
          <xsl:choose>
            <xsl:when test="演:切换_6B1F/演:速度_6B21='middle'">med</xsl:when>
            <xsl:when test="演:切换_6B1F/演:速度_6B21='slow'">slow</xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <!-- 09.12.18 -->
      <xsl:if test="演:切换_6B1F/演:方式_6B23/演:单击鼠标_6B24='false'">
        <xsl:attribute name="advClick">0</xsl:attribute>
      </xsl:if>
      <xsl:if test="演:切换_6B1F/演:方式_6B23/演:时间间隔_6B25">
        <xsl:attribute name="advTm">
          <!--<xsl:value-of select="substring-before($advTmtmp,'0')"></xsl:value-of>-->
          
          <!--start liuyin 20130409 修改bug2814 自动换片时间由00.30变为00.00-->
          <!--<xsl:value-of select="substring-before(substring-after(演:切换_6B1F/演:方式_6B23/演:时间间隔_6B25,'P0Y0M0DT0H0M'),'S')*100"/>-->
          <xsl:value-of select="substring-before(substring-after(演:切换_6B1F/演:方式_6B23/演:时间间隔_6B25,'P0Y0M0DT0H0M'),'S')*1000"/>
          <!--end liuyin 20130409 修改bug2814 自动换片时间由00.30变为00.00-->
          
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="演:切换_6B1F/演:效果_6B20">
        <xsl:choose>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='blinds-horizontal'">
            <p:blinds dir="horz"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='blinds-vertical'">
            <p:blinds dir="vert"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='checkerboard-across'">
            <p:checker dir="horz"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='checkerboard-down'">
            <p:checker dir="vert"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='shape-circle'">
            <p:circle/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='comb-vertical'">
            <p:comb dir="vert"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='comb-horizontal'">
            <p:comb dir="horz"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='cover-up'">
            <p:cover dir="u"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='cover-down'">
            <p:cover dir="d"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='cover-left'">
            <p:cover dir="l"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='cover-right'">
            <p:cover dir="r"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='cover-left-up'">
            <p:cover dir="lu"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='cover-left-down'">
            <p:cover dir="ld"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='cover-right-up'">
            <p:cover dir="ru"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='cover-right-down'">
            <p:cover dir="rd"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='cut-through-black'">
            <p:cut thruBlk="1"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='cut'">
            <p:cut thruBlk="0"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='shape-diamond'">
            <p:diamond/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='dissolve'">
            <p:dissolve/>
          </xsl:when>
          <!--p:extLst-->
          <xsl:when test="演:切换_6B1F/演:效果_6B20='fade-through-black'">
            <p:fade thruBlk="1"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='fade-smoothly'">
            <p:fade thruBlk="0"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='newsflash'">
            <p:newsflash/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='shape-plus'">
            <p:plus/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='plush-up'">
            <p:plus dir="u"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='plush-down'">
            <p:plus dir="d"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='plush-left'">
            <p:plus dir="l"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='plush-right'">
            <p:plus dir="r"/>
          </xsl:when>
          <!--p:pull/-->
          <xsl:when test="演:切换_6B1F/演:效果_6B20='uncover-up'">
            <p:pull dir="u"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='uncover-down'">
            <p:pull dir="d"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='uncover-right'">
            <p:pull dir="r"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='uncover-left'">
            <p:pull dir="l"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='uncover-right-down'">
            <p:pull dir="rd"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='uncover-left-down'">
            <p:pull dir="ld"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='uncover-right-up'">
            <p:pull dir="ru"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='uncover-left-up'">
            <p:pull dir="lu"/>
          </xsl:when>
          <!--p:push/-->
          <xsl:when test="演:切换_6B1F/演:效果_6B20='push-up'">
            <p:push dir="u"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='push-down'">
            <p:push dir="d"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='push-right'">
            <p:push dir="r"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='push-left'">
            <p:push dir="l"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='random-transition'">
            <p:random/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='random-bars-vertical'">
            <p:randomBar dir="vert"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='random-bars-horizontal'">
            <p:randomBar dir="horz"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='split-vertical-in'">
            <p:split orient="vert" dir="in"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='split-vertical-out'">
            <!--2014-04-06, tangjiang, 修复幻灯片门切换效果 -->
            <p14:doors dir="vert" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/>
            <!--<p:split orient="vert" dir="out"/>-->
            <!-- end 2014-04-06, tangjiang, 修复幻灯片门切换效果 -->
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='split-horizontal-in'">
            <p:split orient="horz" dir="in"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='split-horizontal-out'">
            <p:split orient="horz" dir="out"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='strips-right-up'">
            <p:strips dir="ru"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='strips-right-down'">
            <p:strips dir="rd"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='strips-left-up'">
            <p:strips dir="lu"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='strips-left-down'">
            <p:strips dir="ld"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='wedge'">
            <p:wedge/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='wheel-clockwise–1spoke'">
            <p:wheel spokes="1"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='wheel-clockwise–2spoke'">
			 
			  <p:wheel spokes="2"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='wheel-clockwise–3spoke'">
            <p:wheel spokes="3"/>
          </xsl:when>
          <!--可能有问题-->
          <xsl:when test="演:切换_6B1F/演:效果_6B20='wheel-clockwise–4spoke'">
            <p:wheel spokes="4"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='wheel-clockwise–8spoke'">
            <p:wheel spokes="8"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='wipe-up'">
            <p:wipe dir="u"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='wipe-down'">
            <p:wipe dir="d"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='wipe-left'">
            <p:wipe dir="l"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='wipe-right'">
            <p:wipe dir="r"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='box-in'">
            <p:zoom dir="in"/>
          </xsl:when>
          <xsl:when test="演:切换_6B1F/演:效果_6B20='box-out'">
            <p:zoom dir="out"/>
			  
          </xsl:when>
			<xsl:when test="演:切换_6B1F/演:效果_6B20='newsflash'">
				<p:newsflash />
			</xsl:when>
			
        </xsl:choose>
      </xsl:if>
      <!--没有切换声音对应，未转-->
      <!--
      2010.1.9 黎美秀修改 转换声音
       <xsl:if test="演:切换_6B1F/演:声音_6B17_6B17">
        <p:sndAc/>
      </xsl:if> 
      <xsl:if test="演:切换_6B1F/演:声音_6B17_6B17">

        <p:sndAc>
        <p:stSnd>
          <p:snd>
            <xsl:attribute name="r:embed">
              <xsl:value-of select="concat('rid',演:切换_6B1F/演:声音_6B17_6B17/@自定义声音_C632)"/>
            </xsl:attribute>
          </p:snd>
        </p:stSnd>  
        </p:sndAc>
        
      </xsl:if>
      -->
      <!--2010.03.30 马有旭 添加 切换声音 -->
      <xsl:if test="演:切换_6B1F/演:声音_6B17">
        <xsl:choose>
          <xsl:when test="演:切换_6B1F/演:声音_6B17/@自定义声音_C632='stop-previous-sound'">
            <p:sndAc>
              <p:endSnd/>
            </p:sndAc>
          </xsl:when>
          <xsl:otherwise>
            <p:sndAc>
              <p:stSnd>
                <!--2010-11-1罗文甜：增加切换声音是否循环播放-->
                <xsl:if test="演:切换_6B1F/演:声音_6B17/@是否循环播放_C633='true'">
                  <xsl:attribute name="loop">
                    <xsl:value-of select="'1'"/>
                  </xsl:attribute>
                </xsl:if>
                <p:snd>
                  <xsl:attribute name="r:embed">
                    <xsl:value-of select="演:切换_6B1F/演:声音_6B17/@自定义声音_C632"/>
                  </xsl:attribute>
                  <xsl:attribute name="name">
                    <xsl:value-of select="'applause.wav'"/>
                  </xsl:attribute>
                </p:snd>
              </p:stSnd>
            </p:sndAc>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </p:transition>
  </xsl:template>
  <xsl:template name="slide.xml.rels">
    <!--
    2010.1.12 黎美秀添加 
    处理超级链接的情况 不同幻灯片中链接目标地址相同
    
    -->
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout">
        <xsl:attribute name="Id">

          <!-- 2014-04-21, tangjiang, 修复LayoutId重复引起的文件需要修复问题 Start -->
          <!--
          <xsl:choose>
            <xsl:when test="contains(@页面版式引用_6B27,'ID')">
              <xsl:value-of select="concat('rId',substring-after(@页面版式引用_6B27,'ID'))"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="concat('rId',@页面版式引用_6B27)"/>
            </xsl:otherwise>
          </xsl:choose>
          -->

          <xsl:value-of select="concat('rId',@页面版式引用_6B27)"/>
          <!-- end 2014-04-21, tangjiang, 修复LayoutId重复引起的文件需要修复问题 -->
          
        </xsl:attribute>
        <xsl:attribute name="Target">
          <xsl:value-of select="concat('../slideLayouts/',@页面版式引用_6B27,'.xml')"/>
        </xsl:attribute>
      </Relationship>
      
		<!--  李娟 11.12.30-->
		<xsl:for-each select=".//字:区域开始_4165[@类型_413B='hyperlink' and not(ancestor::演:幻灯片备注_6B1D)]">
			<xsl:if test="not(current()/@标识符_4100 = preceding::字:区域开始_4165/@标识符_4100)">
				<!--<xsl:for-each select="//uof:链接集/uof:超级链接[@uof:链源=current()/@字:标识符]">-->
				<xsl:if test="//uof:链接集/超链:链接集_AA0B/超链:超级链接_AA0C[超链:链源_AA00=current()/@标识符_4100]">
					<xsl:for-each select="//uof:链接集/超链:链接集_AA0B/超链:超级链接_AA0C[超链:链源_AA00=current()/@标识符_4100]">


						<!--<xsl:for-each select=".//字:区域开始[@字:类型='hyperlink' and not(ancestor::演:幻灯片备注)]">
        <xsl:if test="not(current()/@字:标识符 = preceding::字:区域开始/@字:标识符)">
          <xsl:for-each select="//uof:链接集/uof:超级链接[@uof:链源=current()/@字:标识符]">-->
						<!-- 09.10.23 马有旭 修改-->
						<xsl:variable name="ctarget">
							<xsl:value-of select="超链:目标_AA01"/>
						</xsl:variable>
						<xsl:variable name="cid">
							<xsl:value-of select="@标识符_AA0A"/>
						</xsl:variable>
            <xsl:if test="$ctarget!=''">
              <!-- 09.10.24 马有旭 修改 -->
              <xsl:if test="not(starts-with(超链:目标_AA01,'Custom Show:')) and ./超链:目标_AA01!='First Slide' and ./超链:目标_AA01!='Last Slide' and ./超链:目标_AA01!='Previous Slide' and ./超链:目标_AA01!='End Show' and ./超链:目标_AA01!='Next Slide'">
                <!--
              <xsl:if test="not(//uof:超级链接[@uof:目标=$ctarget and @uof:标识符!=$cid]) or //uof:超级链接[@uof:目标=$ctarget][1]/@uof:标识符=$cid">
              
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
            </xsl:if>
					</xsl:for-each>
				</xsl:if>
			</xsl:if>
		</xsl:for-each>
		
      
		<xsl:if test="演:幻灯片备注_6B1D">
			<xsl:for-each select="演:幻灯片备注_6B1D">
				<Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml">
					<xsl:attribute name="Id">
						<xsl:value-of select="concat('rId',substring-after(uof:锚点_C644/@图形引用_C62E,'Obj'))"/>
					</xsl:attribute>
					<xsl:attribute name="Target">
						<xsl:value-of select="concat('../notesSlides/notesSlide',substring-after(uof:锚点_C644/@图形引用_C62E,'Obj'),'.xml')"/>
					</xsl:attribute>
				</Relationship>
			</xsl:for-each>
		</xsl:if>
    
      <!--<xsl:if test=".//@字:编号引用=ancestor::uof:UOF/uof:式样集//字:自动编号[字:级别/字:图片符号引用]/@字:标识符">-->
		<xsl:if test=".//@编号引用_4187=ancestor::uof:UOF/uof:式样集//字:自动编号_4124[字:级别_4112/字:图片符号_411B/@引用_411C]/@标识符_4100">
        
			<xsl:variable name="numref" select=".//@编号引用_4187"/>
			<xsl:for-each select="//字:自动编号_4124[字:级别_4112/字:图片符号_411B/@引用_411C]">
        <!--<xsl:for-each select="//字:自动编号[字:级别/字:图片符号引用]">-->
          <!--<xsl:if test="not(current()/字:级别/字:图片符号引用=preceding::*/字:图片符号引用)">-->
				<xsl:if test="not(current()/字:级别_4112/字:图片符号_411B/@引用_411C=preceding::*/字:图片符号_411B/@引用_411C)">
				<Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/image" Target="../media/image1.gif">
              <xsl:attribute name="Id">
	
                <xsl:value-of select="concat('rId',./字:级别_4112/字:图片符号_411B/@引用_411C)"/>
              </xsl:attribute>
				<xsl:for-each select="//对象:对象数据_D701[@标识符_D704=current()/字:级别_4112/字:图片符号_411B/@引用_411C]">

						<xsl:attribute name="Target">
							<xsl:choose>
								<xsl:when test="contains(@标识符_D704,'Obj')">
									<xsl:value-of select="concat('../media/',substring-after(@标识符_D704,'Obj'),'.',@公共类型_D706)"/>
								</xsl:when>
								<xsl:otherwise>
                  <xsl:variable name="objPath" select="./对象:路径_D703"/>
                  <xsl:choose>
                    <xsl:when test="contains($objPath, '/data/' )">
                      <xsl:value-of select="concat('../media/',substring-after($objPath,'/data/'))"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="concat('../media/',substring-after($objPath,'\data\'))"/>
                    </xsl:otherwise>
                  </xsl:choose>
								</xsl:otherwise>
							</xsl:choose>
							
						</xsl:attribute>
					</xsl:for-each>
            </Relationship>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
		<!-- 李娟 11.12.30···········@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
		<!--<xsl:if test=".//图:其他对象引用_8038 or .//@图:图形引用_8007 or .//图:图片数据引用_8037">-->
      <xsl:if test=".//图:图形_8062/ole">
        <xsl:for-each select=".//ole/olerel">
          <xsl:copy-of select="./*"/>
        </xsl:for-each>
      </xsl:if>
		<xsl:if test=".//图:图片数据引用_8037 |.//图:图片_8005/@图形引用_8007 ">
        <!--<xsl:variable name="picref" select=".//@图:其他对象| .//@图:图形引用"/>-->
			  <xsl:variable name="picref" select=".//图:图片数据引用_8037|.//图:图片_8005/@图形引用_8007"/>
        <xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picref]">
			    <xsl:variable name="obdata">
				    <xsl:value-of select="@标识符_D704"/>
			    </xsl:variable>
			    <xsl:if test="not($obdata=following::对象:对象数据_D701/@标识符_D704)">
				    <Relationship Type="http://purl.oclc.org/ooxml/officeDocument/relationships/image">
					    <!--10.23 黎美秀修改 此处存在问题，转出来的格式可能名称为image -->
					    <xsl:attribute name="Id">
						    <!--xsl:value-of select="concat('rId',substring-after(@uof:标识符,'OBJ'))"/>  -->
						    <xsl:value-of select="concat('rId',@标识符_D704)"/>
					    </xsl:attribute>
					    <xsl:choose>
						    <xsl:when test="@公共类型_D706 or @私有类型_D707 ">
							    <xsl:attribute name="Target">
                    
                    <!--start liuyin 20121203 修改音视频丢失-->
                    <xsl:variable name="objPath" select="./对象:路径_D703"/>
                    <xsl:choose>
                      <xsl:when test="contains($objPath, '/data/' )">
                        <xsl:value-of select="concat('../media/',substring-after($objPath,'/data/'))"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="concat('../media/',substring-after($objPath,'\data\'))"/>
                      </xsl:otherwise>
                    </xsl:choose>
                    <!--end liuyin 20121203 修改音视频丢失-->
                
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

      <!--start liuyin 20121203 修改音视频丢失-->
      <xsl:if test=".//图:其他对象引用_8038">
        <xsl:variable name="picref" select=".//图:其他对象引用_8038"/>
        <xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$picref]">
          <xsl:choose>
            <xsl:when test="@公共类型_D706='mp3'">
              <Relationship Type="http://purl.oclc.org/ooxml/officeDocument/relationships/audio">
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
              <Relationship Type="http://purl.oclc.org/ooxml/officeDocument/relationships/video">
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
              <Relationship Type="http://purl.oclc.org/ooxml/officeDocument/relationships/video">
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

      <!-- 2010.03.29 myx add 动画声音-->
		<xsl:if test=".//演:声音_6B22[@自定义声音_C632!='stop-previous-sound'] | .//演:声音_6B17[@自定义声音_C632!='none']">
        <!--2010.04.11 myx -->
      <!--start liuyin 20130304 修改动画序列的动画增强，转换后的ooxml文档无法打开-->
        <xsl:variable name="tgt" select=".//演:声音_6B22/@自定义声音_C632 | .//演:声音_6B17/@自定义声音_C632"/>
      <xsl:for-each select="ancestor::uof:UOF/uof:对象集/对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$tgt]">
        <xsl:variable name="slideid" select="ancestor::演:幻灯片_6C0F/@标识符_6B0A"/>
        
        <xsl:variable name="tgtobdata">
          <xsl:value-of select="@标识符_D704"/>
        </xsl:variable>
        
        <!--2010.04.27<xsl:if test="not(ancestor::演:序列/preceding-sibling::演:序列[//演:声音_6B17_6B17/@自定义声音_C632=$tgt])">-->
        <xsl:if test="not(preceding::*[@自定义声音_C632=$tgt and ancestor::演:幻灯片_6C0F/@标识符_6B0A=$slideid])">
          <Relationship Type="http://purl.oclc.org/ooxml/officeDocument/relationships/audio">
            <xsl:attribute name="Id">
              <xsl:value-of select="$tgtobdata"/>
            </xsl:attribute>
            <!--修改声音target 路径的引用 李娟 2012.02.26-->
            <!--<xsl:for-each select="//对象:对象数据集_D700/对象:对象数据_D701">-->
            <xsl:variable  name="exType">
              <xsl:choose>
                <xsl:when test="//对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$tgt]/@公共类型_D706">

                  <!--start liuyin 20121229 修改动画序列的动画增强，转换后的ooxml文档无法打开-->
                  <!--<xsl:value-of select="concat('../media/',substring-after(//对象:对象数据_D701[@标识符_D704=$tgt]/对象:路径_D703,'.\data\'))"/>-->

                  <xsl:variable name="objPath" select="//对象:对象数据_D701[@标识符_D704=$tgtobdata]/对象:路径_D703"/>
                  <xsl:choose>
                    <xsl:when test="contains($objPath, '/data/' )">
                      <xsl:value-of select="concat('../media/',substring-after($objPath,'/data/'))"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="concat('../media/',substring-after($objPath,'\data\'))"/>
                    </xsl:otherwise>
                  </xsl:choose>
                  <!--end liuyin 20121229 修改动画序列的动画增强，转换后的ooxml文档无法打开-->

                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="concat('../media/',//对象:对象数据集_D700/对象:对象数据_D701[@标识符_D704=$tgtobdata]/对象:路径_D703)"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:attribute name="Target">
              <xsl:value-of select="$exType"/>
            </xsl:attribute>
            <!--</xsl:for-each>-->

          </Relationship>
        </xsl:if>
      </xsl:for-each>
      <!--end liuyin 20130304 修改动画序列的动画增强，转换后的ooxml文档无法打开-->
      </xsl:if>
		 
      
		<!--添加批注引用 李娟 2012.02.25-->
		<xsl:if test="uof:锚点_C644/图:图形_8062/图:文本_803C/图:内容_8043/字:段落_416B/字:句_419D/字:区域开始_4165[@类型_413B='annotation']">
			
			<Relationship  Type="http://purl.oclc.org/ooxml/officeDocument/relationships/comments">
				<xsl:attribute name="Id">
					<xsl:value-of select="concat('rId','comment',substring-after(@标识符_6B0A,'slideId'))" />
				</xsl:attribute>
				<xsl:attribute name="Target">
					<xsl:value-of select="concat('../comments/comment',substring-after(@标识符_6B0A,'slideId'),'.xml')"/>
				</xsl:attribute>
			</Relationship>
		</xsl:if>
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
  <!-- 09.12.23 马有旭 添加 -->
  <xsl:template name="spTgt">
    <!--<xsl:variable name="tgt" select="@演:动画对象"/>-->
	  <xsl:variable name="tgt" select="@对象引用_6C28"/>
    <p:spTgt>
      <xsl:attribute name="spid">
        <xsl:choose>
          <xsl:when test="//图:图形_8062[@标识符_804B=$tgt]/其他对象引用_8038">
            <xsl:value-of select="100+substring-after(//图:图形_8062[@标识符_804B=$tgt]/其他对象引用_8038,'image')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="substring(@对象引用_6C28,4,5)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </p:spTgt>
  </xsl:template>
</xsl:stylesheet>
