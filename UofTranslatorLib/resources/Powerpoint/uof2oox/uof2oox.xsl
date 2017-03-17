<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!--xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:uof="http://schemas.uof.org/cn/2003/uof" xmlns:图="http://schemas.uof.org/cn/2003/graph" xmlns:字="http://schemas.uof.org/cn/2003/uof-wordproc" xmlns:表="http://schemas.uof.org/cn/2003/uof-spreadsheet" xmlns:演="http://schemas.uof.org/cn/2003/uof-slideshow" xmlns:pzip="urn:u2o:xmlns:post-processings:special"-->
<xsl:stylesheet version="1.0"
                xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
xmlns:元="http://schemas.uof.org/cn/2009/metadata"
xmlns:规则="http://schemas.uof.org/cn/2009/rules"
xmlns:式样="http://schemas.uof.org/cn/2009/styles"
xmlns:对象="http://schemas.uof.org/cn/2009/objects"
xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:fo="http://www.w3.org/1999/XSL/Format"
xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
xmlns:dc="http://purl.org/dc/elements/1.1/"
xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
xmlns:dcterms="http://purl.org/dc/terms/"
xmlns:dcmitype="http://purl.org/dc/dcmitype/"
xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"			
xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph"

 xmlns:图表="http://schemas.uof.org/cn/2009/chart"
xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
xmlns:pzip="urn:u2o:xmlns:post-processings:special">
	<xsl:import href="package_relationships.xsl"/>
	<xsl:import href="contentTypes.xsl"/>
	<xsl:import href="metadata.xsl"/>
	<xsl:import href="LayoutRels.xsl"/>
	<xsl:import href="slideMasterRels.xsl"/>
	<xsl:import href="SlideRels.xsl"/>
	<xsl:import href="noteMasterRels.xsl"/>
  <xsl:import href="nMasterRels1.xsl"/>
	<xsl:import href="handoutMasterRels.xsl"/>
	<xsl:import href="noteRels.xsl"/>
	<xsl:import href="comments.xslt"/>
	<xsl:param name="outputFile"/>
	<xsl:output method="xml" encoding="UTF-8"/>
	<xsl:template match="/uof:UOF">
		<!--2012.02.17 lijuan-->
		<pzip:archive pzip:target="{$outputFile}">

      <xsl:if test="//w:media">
		  
        <xsl:copy-of select="//w:media/*"/>
      </xsl:if>
       <!--liuyangyang 2015-02-05 拷贝ole对象信息 start-->
      <xsl:if test="//w:ole">

        <xsl:copy-of select="//w:ole/*"/>
      </xsl:if>
      <!--end liuyangyang 2015-02-05 拷贝ole对象信息-->
      <!-- content types -->
			<pzip:entry pzip:target="[Content_Types].xml">
				<xsl:call-template name="contentTypes"/>
				<xsl:for-each select="/uof:UOF/uof:对象集"><!--注销下面的copyof 否则ppt 下面的文件无法生成 修改路径 李娟 11.12.26-->
					<!--<xsl:copy-of select="child::*"/>-->
				</xsl:for-each>
			</pzip:entry>
			<!-- package relationship item -->
			<pzip:entry pzip:target="_rels/.rels">
				<xsl:call-template name="package-relationships"/>
			</pzip:entry>
			
			<!--metadata-->
			<xsl:if test="uof:元数据">
				<pzip:entry pzip:target="docProps/app.xml">
					<xsl:apply-templates select="uof:元数据" mode="zip1"/>
				</pzip:entry>
				<pzip:entry pzip:target="docProps/core.xml">
					<xsl:apply-templates select="uof:元数据" mode="zip2"/>
				</pzip:entry>
				<!--<xsl:if test="uof:元数据/uof:用户自定义元数据集">-->
				<xsl:if test="uof:元数据/元:元数据_5200/元:用户自定义元数据集_520F">
					<pzip:entry pzip:target="docProps/custom.xml">
						<xsl:apply-templates select="uof:元数据" mode="zip3"/>
					</pzip:entry>
				</xsl:if>
			</xsl:if>

      <!--2014-03-03, tangjiang, 添加OOXML到UOF用户自定义数据集的转换 start -->
      <!--custom xml-->
      <xsl:if test="//扩展:扩展区_B200/w:customItem">
        <xsl:for-each select="//扩展:扩展区_B200/w:customItem">
          <pzip:entry>
            <xsl:attribute name="pzip:target">
              <xsl:value-of select="@fileName"/>
            </xsl:attribute>
            <xsl:copy-of select="./*"/>
          </pzip:entry>
        </xsl:for-each>
      </xsl:if>
      <!-- end 2014-03-03, tangjiang -->
      
			<pzip:entry pzip:target="ppt/presentation.xml">
				<xsl:call-template name="presentation"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/_rels/presentation.xml.rels">
				<xsl:call-template name="presentation.xml.rels"/>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/tableStyles.xml">
				<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
					<xsl:attribute name="def">{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</xsl:attribute>
				</a:tblStyleLst>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/presProps.xml">
				<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
					<!--<xsl:for-each select="uof:演示文稿/演:公用处理规则/演:放映设置">--><!--修改路径 有待确认@@@@@@@@@@@2 李娟 11.12.29-->
						<xsl:for-each select="uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:放映设置_B653">
						<p:showPr>
							<!--<xsl:if test="演:循环放映='true' or 演:循环放映='on' or 演:循环放映='1'">-->
							<xsl:if test="规则:是否循环放映_B65A='true' or 规则:是否循环放映_B65A='on' or 规则:是否循环放映_B65A='1'">
								<xsl:attribute name="loop">1</xsl:attribute>
							</xsl:if>

							<xsl:choose>
								<!--修改是否放映动画及下面内容 李娟 11.12.29-->
								<!--<xsl:when test="not(演:放映动画) or (演:放映动画='true') or (演:放映动画='1') or (演:放映动画='on')">-->
								<xsl:when test="not(规则:是否放映动画_B65E) or 规则:是否放映动画_B65E='true' or 规则:是否放映动画_B65E='1' or 规则:是否放映动画_B65E='on'">
									<xsl:attribute name="showAnimation">1</xsl:attribute>
								</xsl:when>
								<xsl:otherwise>
									<xsl:attribute name="showAnimation">0</xsl:attribute>
								</xsl:otherwise>
							</xsl:choose>
							<!--<xsl:if test="not(演:手动方式) or 演:手动方式='true' or 演:手动方式='1' or 演:手动方式='on'">-->
							<xsl:if test="not(规则:是否手动方式_B65C) or 规则:是否手动方式_B65C='true' or 规则:是否手动方式_B65C='1' or 规则:是否手动方式_B65C='on'">
								<xsl:attribute name="useTimings">0</xsl:attribute>
							</xsl:if>
							<!-- 09.10.19 马有旭 修改 -->
							<xsl:attribute name="showNarration">
								<xsl:choose>
									<!--<xsl:when test="/uof:UOF/uof:扩展区/扩展:扩展区_B200/扩展:扩展_B201/扩展:扩展内容_B204[扩展:路径_B205='式样集/字体集']/扩展:内容_B206/扩展:字体信息_F001='true'">1</xsl:when>
									<xsl:when test="not(/uof:UOF/uof:扩展区)">1</xsl:when>-->
									<xsl:when test="规则:是否使用声音旁白_B662= 'true' or 规则:是否使用声音旁白_B662= 'on' or 规则:是否使用声音旁白_B662= '1'">1</xsl:when>
									<xsl:otherwise>0</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>

              <!--start liuyin 20130103 修改放映方式播放类型-展台浏览转换错误-->              
              <xsl:choose>
                <xsl:when test="../规则:幻灯片播放类型_B663='speaker'">
                  <p:present/>
                </xsl:when>
                <!--<xsl:when test="../规则:幻灯片播放类型_B663=' '">
                  <p:browse/>
                </xsl:when> 永中2012无此类型-->
                <xsl:when test="../规则:幻灯片播放类型_B663='kiosk'">
                  <p:kiosk/>
                </xsl:when>
                <xsl:otherwise>
                  <p:browse/>
                </xsl:otherwise>
              </xsl:choose>
              <!--end liuyin 20130103 修改放映方式播放类型-展台浏览转换错误-->

              <!-- 09.10.19 马有旭 添加 -->
							<!--注销对扩展区调用这块 李娟 2012.01.10-->
							<!--<xsl:apply-templates select="/uof:UOF/uof:扩展区/uof:扩展[uof:扩展内容/uof:路径='uof/演示文稿']" mode="showType"/>-->
							<!--<xsl:apply-templates select="/uof:UOF/uof:扩展区/扩展:扩展区_B200/扩展:扩展_B201[扩展:扩展内容_B204/扩展:路径_B205='uof/演示文稿']" mode="showType"/>-->
							<xsl:choose>
								<xsl:when test="规则:放映顺序_B658">
									<!-- 自定义放映 或者 序列放映 -->
									<xsl:choose>
										<xsl:when test="规则:放映顺序_B658='sequentListId'">
											<!--<xsl:when test="演:放映顺序/@演:序列引用='sequentListId'">-->
											<!--序列放映-->
											<p:sldRg>
												<xsl:call-template name="sequent">
													<xsl:with-param name="stid" select="normalize-space(substring-before(规则:幻灯片序列_B654[@标识符_B655='sequentListId']/text(),' '))"/>
													<xsl:with-param name="endid" select="normalize-space(substring-after(规则:幻灯片序列_B654[@标识符_B655='sequentListId']/text(),' '))"/>
												</xsl:call-template>
											</p:sldRg>
										</xsl:when>
										<xsl:otherwise>
											<!--自定义放映-->
											<xsl:call-template name="custShow">
												<xsl:with-param name="id" select="规则:放映顺序_B658"/>
												<!--<xsl:with-param name="id" select="演:放映顺序/@演:序列引用"/>-->
											</xsl:call-template>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:when>
								<xsl:otherwise>
									<!-- 全部放映 -->
									<p:sldAll/>
								</xsl:otherwise>
							</xsl:choose>
							<!--罗文甜 增加：绘图笔颜色-->
							<xsl:if test="规则:绘图笔颜色_B661">
								<!--<xsl:if test="演:绘图笔颜色">-->
								<p:penClr>
									<a:srgbClr>
										<xsl:attribute name="val">
											<xsl:value-of select="substring-after(规则:绘图笔颜色_B661,'#')"/>
										</xsl:attribute>
									</a:srgbClr>
								</p:penClr>
							</xsl:if>
						</p:showPr>
					</xsl:for-each>
				</p:presentationPr>
			</pzip:entry>
			<pzip:entry pzip:target="ppt/viewProps.xml">
				<!--<xsl:for-each select="uof:演示文稿/演:公用处理规则">-->
				<xsl:for-each select="uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665">
					<p:viewPr>
						<xsl:choose>
							<xsl:when test="规则:演示文稿_B66D/规则:最后视图_B639/规则:类型_B63A!='normal'">
								<!--<xsl:when test="演:最后视图/@演:类型!='normal'">-->
								<xsl:attribute name="lastView">
									<xsl:choose>
										<!--修改这块是否符合要求 李娟 11.12.29-->
										<xsl:when test="规则:演示文稿_B66D/规则:最后视图_B639/规则:类型_B63A='slide-master'">sldMasterView</xsl:when>
										<xsl:when test="规则:演示文稿_B66D/规则:最后视图_B639/规则:类型_B63A='note-page'">notesView</xsl:when>
										<xsl:when test="规则:演示文稿_B66D/规则:最后视图_B639/规则:类型_B63A='handout-master'">handoutView</xsl:when>
										<xsl:when test="规则:演示文稿_B66D/规则:最后视图_B639/规则:类型_B63A='note-master'">notesMasterView</xsl:when>
										<xsl:when test="规则:演示文稿_B66D/规则:最后视图_B639/规则:类型_B63A='sort'">sldSorterView</xsl:when>
                    
                    <!--start liuyin 20130102 修改幻灯片折叠显示大纲需要修复才能打开-->
                    <xsl:when test="规则:演示文稿_B66D/规则:最后视图_B639/规则:类型_B63A='outline'">sldThumbnailView</xsl:when>
                    <!--end liuyin 20130102 修改幻灯片折叠显示大纲需要修复才能打开-->
                    
										<!--<xsl:when test="演:最后视图/@演:类型='slide-master'">sldMasterView</xsl:when>
										<xsl:when test="演:最后视图/@演:类型='note-page'">notesView</xsl:when>
										<xsl:when test="演:最后视图/@演:类型='handout-master'">handoutView</xsl:when>
										<xsl:when test="演:最后视图/@演:类型='note-master'">notesMasterView</xsl:when>
										<xsl:when test="演:最后视图/@演:类型='sort'">sldSorterView</xsl:when>-->
									</xsl:choose>
								</xsl:attribute>
							</xsl:when>
							<xsl:when test="规则:最后视图_B639/规则:类型_B63A!='normal'">
								<!--<xsl:when test="演:最后视图/@演:类型!='normal'">-->
								<xsl:attribute name="lastView">
									<xsl:choose>
										<!--修改这块是否符合要求 李娟 11.12.29-->
										<xsl:when test="规则:最后视图_B639/规则:类型_B63A='slide-master'">sldMasterView</xsl:when>
										<xsl:when test="规则:最后视图_B639/规则:类型_B63A='note-page'">notesView</xsl:when>
										<xsl:when test="规则:最后视图_B639/规则:类型_B63A='handout-master'">handoutView</xsl:when>
										<xsl:when test="规则:最后视图_B639/规则:类型_B63A='note-master'">notesMasterView</xsl:when>
										<xsl:when test="规则:最后视图_B639/规则:类型_B63A='sort'">sldSorterView</xsl:when>
										<!--<xsl:when test="演:最后视图/@演:类型='slide-master'">sldMasterView</xsl:when>
										<xsl:when test="演:最后视图/@演:类型='note-page'">notesView</xsl:when>
										<xsl:when test="演:最后视图/@演:类型='handout-master'">handoutView</xsl:when>
										<xsl:when test="演:最后视图/@演:类型='note-master'">notesMasterView</xsl:when>
										<xsl:when test="演:最后视图/@演:类型='sort'">sldSorterView</xsl:when>-->
									</xsl:choose>
								</xsl:attribute>
							</xsl:when>
						</xsl:choose>
						<p:normalViewPr showOutlineIcons="0">
							<p:restoredLeft sz="15620"/>
							<p:restoredTop sz="94660"/>
						</p:normalViewPr>
						<p:slideViewPr>
							<p:cSldViewPr>
								<p:cViewPr>
									<p:scale>
										<!--2011-3-20罗文甜修改默认比例-->
										<a:sx>
											<xsl:attribute name="n">
												<xsl:if test="规则:演示文稿_B66D/规则:显示比例_B63F">
													<xsl:value-of select="round(规则:演示文稿_B66D/规则:显示比例_B63F)"/>
												</xsl:if>
												<xsl:if test="not(规则:演示文稿_B66D/规则:显示比例_B63F)">66</xsl:if>
											</xsl:attribute>
												
										<xsl:attribute name="d">100</xsl:attribute>
										</a:sx>
										<a:sy>
											<xsl:attribute name="n">
												<xsl:if test="规则:演示文稿_B66D/规则:显示比例_B63F">
													<xsl:value-of select="round(规则:演示文稿_B66D/规则:显示比例_B63F)"/>
												</xsl:if>
												<xsl:if test="not(规则:演示文稿_B66D/规则:显示比例_B63F)">66</xsl:if>
											</xsl:attribute>


											<xsl:attribute name="d">100</xsl:attribute>
										</a:sy>
									</p:scale>
									<p:origin x="-606" y="-90"/>
								</p:cViewPr>
								<p:guideLst>
									<p:guide orient="horz" pos="2160"/>
									<p:guide pos="2880"/>
								</p:guideLst>
							</p:cSldViewPr>
						</p:slideViewPr>
						<p:notesTextViewPr>
							<p:cViewPr>
								<p:scale>
									<a:sx n="100" d="100"/>
									<a:sy n="100" d="100"/>
								</p:scale>
								<p:origin x="0" y="0"/>
							</p:cViewPr>
						</p:notesTextViewPr>
						
						<xsl:if test="规则:演示文稿_B66D/规则:页面设置集_B670/规则:页面设置_B638/演:纸张方向_6BE1='landscape'">

							<p:notesViewPr>
								<p:cSldViewPr>
									<p:cViewPr varScale="1">
										<p:scale>
											<!--2011-3-20罗文甜修改默认比例-->
											<a:sx>
												<xsl:attribute name="n">
													<xsl:if test="规则:演示文稿_B66D/规则:显示比例_B63F">
														<xsl:value-of select="round(规则:演示文稿_B66D/规则:显示比例_B63F)"/>
													</xsl:if>
													<xsl:if test="not(规则:演示文稿_B66D/规则:显示比例_B63F)">66</xsl:if>
												</xsl:attribute>

												<xsl:attribute name="d">100</xsl:attribute>
											</a:sx>
											<a:sy>
												<xsl:attribute name="n">
													<xsl:if test="规则:演示文稿_B66D/规则:显示比例_B63F">
														<xsl:value-of select="round(规则:演示文稿_B66D/规则:显示比例_B63F)"/>
													</xsl:if>
													<xsl:if test="not(规则:演示文稿_B66D/规则:显示比例_B63F)">66</xsl:if>
												</xsl:attribute>


												<xsl:attribute name="d">100</xsl:attribute>
											</a:sy>
										</p:scale>
										<p:origin x="-606" y="-90"/>
									</p:cViewPr>
									<p:guideLst>
										<p:guide orient="horz" pos="2880"/>
										<p:guide pos="2160"/>
									</p:guideLst>
									
								</p:cSldViewPr>
							</p:notesViewPr>
						</xsl:if>

            <!--2014-04-09, tangjiang, 添加网格线的转换 start -->
            <xsl:choose>
              <xsl:when test="规则:演示文稿_B66D/规则:绘图网格与参考线_B602/@绘图网格间距_C603">
                <xsl:variable name="gridSpace" select="规则:演示文稿_B66D/规则:绘图网格与参考线_B602/@绘图网格间距_C603"/>
                <p:gridSpacing>
                  <xsl:attribute name="cx">
                    <xsl:value-of select="round($gridSpace * 12700)"/>
                  </xsl:attribute>
                  <xsl:attribute name="cy">
                    <xsl:value-of select="round($gridSpace * 12700)"/>
                  </xsl:attribute>
                </p:gridSpacing>
              </xsl:when>
              <xsl:otherwise>
                <p:gridSpacing cx="78028800" cy="78028800"/>
              </xsl:otherwise>
            </xsl:choose>
            <!--end 2014-04-09, tangjiang, 添加网格线的转换 -->
            
					</p:viewPr>
				</xsl:for-each>
			</pzip:entry>

			<!--comments 李娟 2012.02.22-->
			
			<!--<xsl:for-each select="uof:演示文稿/演:主体/演:母版集/演:母版">--><!--李娟 11.12.29-->
			<xsl:for-each select="uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D">
				<!-- 09.10.12 马有旭 修改<pzip:entry pzip:target="ppt/theme/theme1.xml">-->
				<pzip:entry>
					<xsl:attribute name="pzip:target">
						<xsl:choose>
							<xsl:when test="@配色方案引用_6BEB">
								<xsl:choose>
									<xsl:when test="contains(@配色方案引用_6BEB,'clr_')">
										<xsl:value-of select="concat('ppt/theme/theme',substring-after(@配色方案引用_6BEB,'clr_'),'.xml')"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="concat('ppt/theme/theme',substring-after(@配色方案引用_6BEB,'Id'),'.xml')"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="concat('ppt/theme/theme',substring-after(@标识符_6BE8,'ID'),'.xml')"/>
							</xsl:otherwise>
							
						</xsl:choose>
					</xsl:attribute>
					<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 主题">
						<a:themeElements>
							<a:clrScheme name="Office">
								<a:dk1>
									<a:sysClr val="windowText" lastClr="000000"/>
								</a:dk1>
								<a:lt1>
									<a:sysClr val="window" lastClr="FFFFFF"/>
								</a:lt1>
								<a:dk2>
									<a:srgbClr val="1F497D"/>
								</a:dk2>
								<a:lt2>
									<a:srgbClr val="EEECE1"/>
								</a:lt2>
								<a:accent1>
									<a:srgbClr val="4F81BD"/>
								</a:accent1>
								<a:accent2>
									<a:srgbClr val="C0504D"/>
								</a:accent2>
								<a:accent3>
									<a:srgbClr val="9BBB59"/>
								</a:accent3>
								<a:accent4>
									<a:srgbClr val="8064A2"/>
								</a:accent4>
								<a:accent5>
									<a:srgbClr val="4BACC6"/>
								</a:accent5>
								<a:accent6>
									<a:srgbClr val="F79646"/>
                </a:accent6>
                <!--start liuyin 20130417 修改2839 uof->ooxml回归测试的功能测试中，超链接式样转换后效果不一致-->

                <!--start liuyin 20130110 修改幻灯片配色方案效果丢失-->
                <xsl:variable name="colorSchemename">
                  <xsl:choose>
                    <xsl:when test="../../演:幻灯片集_6C0E/演:幻灯片_6C0F/@配色方案引用_6C12 != ''">
                      <xsl:value-of select="../../演:幻灯片集_6C0E/演:幻灯片_6C0F/@配色方案引用_6C12"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="./@配色方案引用_6BEB"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>
                <xsl:variable name="colorSchemename1">
                  <xsl:value-of select="../../演:幻灯片集_6C0E/演:幻灯片_6C0F/@页面版式引用_6B27"/>
                </xsl:variable>
                  <xsl:choose>
                    <xsl:when test="../../演:幻灯片集_6C0E/演:幻灯片_6C0F/@配色方案引用_6C12 or ./@配色方案引用_6BEB">
                      <a:hlink>
                      <a:srgbClr>
                        <xsl:attribute name="val">
                          <xsl:value-of select="substring(/uof:UOF/uof:演示文稿/演:公用处理规则/规则:配色方案集_6C11/规则:配色方案_6BE4[@标识符_6B0A=$colorSchemename]/规则:强调和超级链接_6B08,2,6)"/>
                        </xsl:attribute>
                      </a:srgbClr>
                      </a:hlink>
                      <a:folHlink>
                        <a:srgbClr>
                          <xsl:attribute name="val">
                            <xsl:value-of select="substring(/uof:UOF/uof:演示文稿/演:公用处理规则/规则:配色方案集_6C11/规则:配色方案_6BE4[@标识符_6B0A=$colorSchemename]/规则:强调和尾随超级链接_6B09,2,6)"/>
                          </xsl:attribute>
                        </a:srgbClr>
                      </a:folHlink>
                    </xsl:when>
                    <xsl:when test="../../演:幻灯片集_6C0E/演:幻灯片_6C0F/@页面版式引用_6B27">
                      <a:hlink>
                      <a:srgbClr>
                        <xsl:attribute name="val">
                          <xsl:value-of select="substring(/uof:UOF/uof:演示文稿/演:公用处理规则/规则:配色方案集_6C11/规则:配色方案_6BE4/规则:强调和超级链接_6B08,2,6)"/>
                        </xsl:attribute>
                      </a:srgbClr>
                      </a:hlink>
                      <a:folHlink>
                        <a:srgbClr>
                          <xsl:attribute name="val">
                            <xsl:value-of select="substring(/uof:UOF/uof:演示文稿/演:公用处理规则/规则:配色方案集_6C11/规则:配色方案_6BE4/规则:强调和尾随超级链接_6B09,2,6)"/>
                          </xsl:attribute>
                        </a:srgbClr>
                      </a:folHlink>
                    </xsl:when>
                    <xsl:otherwise>
                      <a:hlink>
                      <a:srgbClr val="0000FF"/>
                      </a:hlink>
                      <a:folHlink>
                        <a:srgbClr val="800080"/>
                      </a:folHlink>
                    </xsl:otherwise>
                  </xsl:choose>
                  <!--start liuyin 20130110 修改幻灯片配色方案效果丢失-->
                <!--end liuyin 20130417 修改2839 uof->ooxml回归测试的功能测试中，超链接式样转换后效果不一致-->

              </a:clrScheme>
							<a:fontScheme name="Office">
								<a:majorFont>
									<a:latin typeface="Calibri"/>
									<a:ea typeface=""/>
									<a:cs typeface=""/>
									<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
									<a:font script="Hang" typeface="맑은 고딕"/>
									<a:font script="Hans" typeface="宋体"/>
									<a:font script="Hant" typeface="新細明體"/>
									<a:font script="Arab" typeface="Times New Roman"/>
									<a:font script="Hebr" typeface="Times New Roman"/>
									<a:font script="Thai" typeface="Angsana New"/>
									<a:font script="Ethi" typeface="Nyala"/>
									<a:font script="Beng" typeface="Vrinda"/>
									<a:font script="Gujr" typeface="Shruti"/>
									<a:font script="Khmr" typeface="MoolBoran"/>
									<a:font script="Knda" typeface="Tunga"/>
									<a:font script="Guru" typeface="Raavi"/>
									<a:font script="Cans" typeface="Euphemia"/>
									<a:font script="Cher" typeface="Plantagenet Cherokee"/>
									<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
									<a:font script="Tibt" typeface="Microsoft Himalaya"/>
									<a:font script="Thaa" typeface="MV Boli"/>
									<a:font script="Deva" typeface="Mangal"/>
									<a:font script="Telu" typeface="Gautami"/>
									<a:font script="Taml" typeface="Latha"/>
									<a:font script="Syrc" typeface="Estrangelo Edessa"/>
									<a:font script="Orya" typeface="Kalinga"/>
									<a:font script="Mlym" typeface="Kartika"/>
									<a:font script="Laoo" typeface="DokChampa"/>
									<a:font script="Sinh" typeface="Iskoola Pota"/>
									<a:font script="Mong" typeface="Mongolian Baiti"/>
									<a:font script="Viet" typeface="Times New Roman"/>
									<a:font script="Uigh" typeface="Microsoft Uighur"/>
								</a:majorFont>
								<a:minorFont>
									<a:latin typeface="Calibri"/>
									<a:ea typeface=""/>
									<a:cs typeface=""/>
									<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
									<a:font script="Hang" typeface="맑은 고딕"/>
									<a:font script="Hans" typeface="宋体"/>
									<a:font script="Hant" typeface="新細明體"/>
									<a:font script="Arab" typeface="Arial"/>
									<a:font script="Hebr" typeface="Arial"/>
									<a:font script="Thai" typeface="Cordia New"/>
									<a:font script="Ethi" typeface="Nyala"/>
									<a:font script="Beng" typeface="Vrinda"/>
									<a:font script="Gujr" typeface="Shruti"/>
									<a:font script="Khmr" typeface="DaunPenh"/>
									<a:font script="Knda" typeface="Tunga"/>
									<a:font script="Guru" typeface="Raavi"/>
									<a:font script="Cans" typeface="Euphemia"/>
									<a:font script="Cher" typeface="Plantagenet Cherokee"/>
									<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
									<a:font script="Tibt" typeface="Microsoft Himalaya"/>
									<a:font script="Thaa" typeface="MV Boli"/>
									<a:font script="Deva" typeface="Mangal"/>
									<a:font script="Telu" typeface="Gautami"/>
									<a:font script="Taml" typeface="Latha"/>
									<a:font script="Syrc" typeface="Estrangelo Edessa"/>
									<a:font script="Orya" typeface="Kalinga"/>
									<a:font script="Mlym" typeface="Kartika"/>
									<a:font script="Laoo" typeface="DokChampa"/>
									<a:font script="Sinh" typeface="Iskoola Pota"/>
									<a:font script="Mong" typeface="Mongolian Baiti"/>
									<a:font script="Viet" typeface="Arial"/>
									<a:font script="Uigh" typeface="Microsoft Uighur"/>
								</a:minorFont>
							</a:fontScheme>
							<a:fmtScheme name="Office">
								<a:fillStyleLst>
									<a:solidFill>
										<a:schemeClr val="phClr"/>
									</a:solidFill>
									<a:gradFill rotWithShape="1">
										<a:gsLst>
											<a:gs pos="0">
												<a:schemeClr val="phClr">
													<a:tint val="50000"/>
													<a:satMod val="300000"/>
												</a:schemeClr>
											</a:gs>
											<a:gs pos="35000">
												<a:schemeClr val="phClr">
													<a:tint val="37000"/>
													<a:satMod val="300000"/>
												</a:schemeClr>
											</a:gs>
											<a:gs pos="100000">
												<a:schemeClr val="phClr">
													<a:tint val="15000"/>
													<a:satMod val="350000"/>
												</a:schemeClr>
											</a:gs>
										</a:gsLst>
										<a:lin ang="16200000" scaled="1"/>
									</a:gradFill>
									<a:gradFill rotWithShape="1">
										<a:gsLst>
											<a:gs pos="0">
												<a:schemeClr val="phClr">
													<a:shade val="51000"/>
													<a:satMod val="130000"/>
												</a:schemeClr>
											</a:gs>
											<a:gs pos="80000">
												<a:schemeClr val="phClr">
													<a:shade val="93000"/>
													<a:satMod val="130000"/>
												</a:schemeClr>
											</a:gs>
											<a:gs pos="100000">
												<a:schemeClr val="phClr">
													<a:shade val="94000"/>
													<a:satMod val="135000"/>
												</a:schemeClr>
											</a:gs>
										</a:gsLst>
										<a:lin ang="16200000" scaled="0"/>
									</a:gradFill>
								</a:fillStyleLst>
								<a:lnStyleLst>
									<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
										<a:solidFill>
											<a:schemeClr val="phClr">
												<a:shade val="95000"/>
												<a:satMod val="105000"/>
											</a:schemeClr>
										</a:solidFill>
										<a:prstDash val="solid"/>
									</a:ln>
									<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
										<a:solidFill>
											<a:schemeClr val="phClr"/>
										</a:solidFill>
										<a:prstDash val="solid"/>
									</a:ln>
									<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
										<a:solidFill>
											<a:schemeClr val="phClr"/>
										</a:solidFill>
										<a:prstDash val="solid"/>
									</a:ln>
								</a:lnStyleLst>
								<a:effectStyleLst>
									<a:effectStyle>
										<a:effectLst>
											<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
												<a:srgbClr val="000000">
													<a:alpha val="38000"/>
												</a:srgbClr>
											</a:outerShdw>
										</a:effectLst>
									</a:effectStyle>
									<a:effectStyle>
										<a:effectLst>
											<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
												<a:srgbClr val="000000">
													<a:alpha val="35000"/>
												</a:srgbClr>
											</a:outerShdw>
										</a:effectLst>
									</a:effectStyle>
									<a:effectStyle>
										<a:effectLst>
											<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
												<a:srgbClr val="000000">
													<a:alpha val="35000"/>
												</a:srgbClr>
											</a:outerShdw>
										</a:effectLst>
										<a:scene3d>
											<a:camera prst="orthographicFront">
												<a:rot lat="0" lon="0" rev="0"/>
											</a:camera>
											<a:lightRig rig="threePt" dir="t">
												<a:rot lat="0" lon="0" rev="1200000"/>
											</a:lightRig>
										</a:scene3d>
										<a:sp3d>
											<a:bevelT w="63500" h="25400"/>
										</a:sp3d>
									</a:effectStyle>
								</a:effectStyleLst>
								<a:bgFillStyleLst>
									<a:solidFill>
										<a:schemeClr val="phClr"/>
									</a:solidFill>
									<a:gradFill rotWithShape="1">
										<a:gsLst>
											<a:gs pos="0">
												<a:schemeClr val="phClr">
													<a:tint val="40000"/>
													<a:satMod val="350000"/>
												</a:schemeClr>
											</a:gs>
											<a:gs pos="40000">
												<a:schemeClr val="phClr">
													<a:tint val="45000"/>
													<a:shade val="99000"/>
													<a:satMod val="350000"/>
												</a:schemeClr>
											</a:gs>
											<a:gs pos="100000">
												<a:schemeClr val="phClr">
													<a:shade val="20000"/>
													<a:satMod val="255000"/>
												</a:schemeClr>
											</a:gs>
										</a:gsLst>
										<a:path path="circle">
											<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
										</a:path>
									</a:gradFill>
									<a:gradFill rotWithShape="1">
										<a:gsLst>
											<a:gs pos="0">
												<a:schemeClr val="phClr">
													<a:tint val="80000"/>
													<a:satMod val="300000"/>
												</a:schemeClr>
											</a:gs>
											<a:gs pos="100000">
												<a:schemeClr val="phClr">
													<a:shade val="30000"/>
													<a:satMod val="200000"/>
												</a:schemeClr>
											</a:gs>
										</a:gsLst>
										<a:path path="circle">
											<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
										</a:path>
									</a:gradFill>
								</a:bgFillStyleLst>
							</a:fmtScheme>
						</a:themeElements>
						<a:objectDefaults/>
						<a:extraClrSchemeLst/>
					</a:theme>
				</pzip:entry>
			</xsl:for-each>
      <!--2011-3-20罗文甜，增加判断，修改bug-->
			<!-- 修改路径 李娟 11.12.29-->
			<xsl:if test="uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D">
				<xsl:choose>
					<xsl:when test="uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651/规则:页面版式_B652">
						<xsl:for-each select="uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651/规则:页面版式_B652">
							
							<xsl:call-template name="LayoutRels">
								
							</xsl:call-template>
							<!--</pzip:entry>-->
						</xsl:for-each>
					</xsl:when>
					
				</xsl:choose>
			<!--<xsl:if test="uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651/规则:页面版式_B652">
				 <xsl:for-each select="uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651/规则:页面版式_B652">
           
					 <xsl:call-template name="LayoutRels"/>
		 
				</xsl:for-each>
		</xsl:if>-->
		</xsl:if>
		<xsl:for-each select="uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D">
			<!--<xsl:for-each select="uof:演示文稿/演:主体/演:母版集/演:母版">-->
				<!--<xsl:if test="@类型_6BEA='slide'">-->
          <xsl:if test="@类型_6BEA='slide' or not(@类型_6BEA)">
					<!--<pzip:entry pzip:target="ppt/slideMasters/_rels/slideMasters1.xml.rels">-->
						<xsl:call-template name="slideMasterRels"/>
					<!--</pzip:entry>-->
				</xsl:if>
				<!--xsl:if test="@演:类型='notes'">
        <此处2月16日加-->
				<xsl:if test="@类型_6BEA='notes'"> 
          <xsl:call-template name="noteMasterRels">
          </xsl:call-template>
				</xsl:if>
				<xsl:if test="@类型_6BEA='handout'">
          <xsl:call-template name="handoutMasterRels">
         
          </xsl:call-template>
				</xsl:if>
			</xsl:for-each>
		<xsl:for-each select="uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F">
			
				<xsl:call-template name="SlideRels"/>
				
				<xsl:if test="./演:幻灯片备注_6B1D">
					<xsl:for-each select="演:幻灯片备注_6B1D">
						<!--<xsl:if test="not(ancestor::uof:演示文稿/演:主体/演:母版集/演:母版[@演:类型='notes'])">-->
						<xsl:if test="not(ancestor::uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='notes'])">
              <xsl:call-template name="noteMasterRels1">
            
              </xsl:call-template>
						</xsl:if>
						<!--xsl:call-template name="noteMasterRels"/>
            <此处2月16日加-->
						<xsl:call-template name="NoteRels"/>
					</xsl:for-each>
				</xsl:if>
			</xsl:for-each>

			<xsl:if test="uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:用户集_B667">
				<pzip:entry pzip:target="ppt/commentAuthors.xml">
					<xsl:for-each select="uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:用户集_B667/规则:用户_B668">
						<xsl:variable name="ID">
							<xsl:value-of select="@标识符_4100"/>
						</xsl:variable>
					</xsl:for-each>
					<xsl:call-template name="commentAuthors"/>
				</pzip:entry>
				<xsl:if test="uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:批注集_B669">
					<xsl:for-each select="//规则:批注_B66A">
						<xsl:variable name="Field">
							<xsl:value-of select="@区域引用_41CE"/>
						</xsl:variable>

					</xsl:for-each>

						<xsl:call-template name="comment"/>
				</xsl:if>
				
			</xsl:if>
		</pzip:archive>
	</xsl:template>
	<xsl:template match="uof:元数据" mode="zip1">
		<xsl:call-template name="metadataApp"/>
	</xsl:template>
	<xsl:template match="uof:元数据" mode="zip2">
		<xsl:call-template name="metadataCore"/>
	</xsl:template>
	<xsl:template match="uof:元数据" mode="zip3">
		<xsl:call-template name="metadataCustom"/>
	</xsl:template>
	<xsl:template name="presentation">
		<p:presentation saveSubsetFonts="1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <!-- 2010.03.25 马有旭 -->
			<xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面设置集_B670/规则:页面设置_B638/演:页码格式_6BDF/@起始编号_6BE0">
				<xsl:attribute name="firstSlideNum">
					<xsl:variable name="initnum" select="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面设置集_B670/规则:页面设置_B638/演:页码格式_6BDF/@起始编号_6BE0"/>
					<xsl:choose>
						<xsl:when test="string(number($initnum))!='NaN'">
							<xsl:value-of select="$initnum"/>
						</xsl:when>
						<xsl:otherwise>1</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
			</xsl:if>
      <!--2010-11-8 罗文甜：增添属性-->
			<!--修改标签 李娟 11.12.26-->
			<xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页眉页脚集_B640/规则:幻灯片_B641/@标题幻灯片中是否显示_B64B='false'">
      <!--<xsl:if test="/uof:UOF/uof:演示文稿/演:公用处理规则/演:页眉页脚集/演:幻灯片页眉页脚/@演:标题幻灯片中不显示='true'">-->
        <xsl:attribute name="showSpecialPlsOnTitleSld">0</xsl:attribute>
      </xsl:if>
			<xsl:if test="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='slide' or not(@类型_6BEA)]">
			<!--<xsl:if test="/uof:UOF/uof:演示文稿/演:主体/演:母版集/演:母版[@演:类型='slide' or not(@演:类型)]">-->
				<p:sldMasterIdLst>
					<xsl:for-each select="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='slide' or not(@类型_6BEA)]">
					<!--<xsl:for-each select="/uof:UOF/uof:演示文稿/演:主体/演:母版集/演:母版[@演:类型='slide' or not(@演:类型)]">-->
						<p:sldMasterId>
							<xsl:attribute name="id">
								<xsl:choose>
									<xsl:when test="contains(@标识符_6BE8,'masterId')">
										<xsl:value-of select="2147483648 + substring-after(@标识符_6BE8,'masterId')"/>
										</xsl:when>
											<xsl:otherwise>
												<xsl:value-of select="2147483648 + substring-after(@标识符_6BE8,'slideMaster')"/>
											</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
								<xsl:attribute name="r:id">
									<xsl:value-of select="concat('rId',@标识符_6BE8)"/>
								</xsl:attribute>
						</p:sldMasterId>
					</xsl:for-each>
				</p:sldMasterIdLst>
			</xsl:if>
			<!--<xsl:if test="/uof:UOF/uof:演示文稿/演:主体/演:母版集/演:母版[@演:类型='notes']|/uof:UOF/uof:演示文稿/演:主体/演:幻灯片集/演:幻灯片/演:幻灯片备注">-->
			<xsl:if test="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='notes']|/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/演:幻灯片备注_6B1D">
					<p:notesMasterIdLst>
					<xsl:choose>
						<xsl:when test="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='notes']">
							<xsl:for-each select="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='notes']">
						<!--<xsl:when test="/uof:UOF/uof:演示文稿/演:主体/演:母版集/演:母版[@演:类型='notes']">
							<xsl:for-each select="/uof:UOF/uof:演示文稿/演:主体/演:母版集/演:母版[@演:类型='notes']">-->
								<p:notesMasterId>
                  <!--2011-5-30罗文甜修改编号引用Bug-->
									<xsl:attribute name="r:id">
										<xsl:choose>
											<xsl:when test="contains(@标识符_6BE8,'ID')">
												<xsl:value-of select="concat('rId',substring-after(@标识符_6BE8,'ID'))"/>
												</xsl:when>
													<xsl:otherwise>
												<xsl:value-of select="concat('rId',@标识符_6BE8)"/>
											</xsl:otherwise>
										</xsl:choose>
								</xsl:attribute>
								</p:notesMasterId>
							</xsl:for-each>
						</xsl:when>
						<xsl:when test="not(/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='notes']) and /uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/演:幻灯片备注_6B1D">
							<xsl:for-each select="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/演:幻灯片备注_6B1D">
						<!--<xsl:when test="not(/uof:UOF/uof:演示文稿/演:主体/演:母版集/演:母版[@演:类型='notes']) and /uof:UOF/uof:演示文稿/演:主体/演:幻灯片集/演:幻灯片/演:幻灯片备注">
							<xsl:for-each select="/uof:UOF/uof:演示文稿/演:主体/演:幻灯片集/演:幻灯片/演:幻灯片备注">-->
								<p:notesMasterId>
									<xsl:attribute name="r:id">
										<xsl:value-of select="concat('rId',substring-after(.//@图形引用_C62E,'Obj'))"/>
									</xsl:attribute>
								</p:notesMasterId>
							</xsl:for-each>
						</xsl:when>
					</xsl:choose>
				</p:notesMasterIdLst>
			</xsl:if>
      <!--添加讲义母版 引用 李娟 2012 04 。09-->
			<xsl:if test="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='handout']">
				<p:handoutMasterIdLst>
					<xsl:for-each select="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='handout']">
						<p:handoutMasterId>

							<xsl:attribute name="r:id">
								<xsl:choose>
									<xsl:when test="contains(@标识符_6BE8,'ID')">
										<xsl:value-of select="concat('rId',substring-after(@标识符_6BE8,'ID'))"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="concat('rId',@标识符_6BE8)"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
						</p:handoutMasterId>
					</xsl:for-each>
				</p:handoutMasterIdLst>
			</xsl:if>
			<xsl:if test="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F">
				<!--<xsl:if test="/uof:UOF/uof:演示文稿/演:主体/演:幻灯片集/演:幻灯片">-->
				<p:sldIdLst>
					<xsl:for-each select="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E//演:幻灯片_6C0F">
						<p:sldId>
							<xsl:attribute name="id">
								<xsl:choose>
									<xsl:when test="contains(@标识符_6B0A,'slideId')">
										<xsl:value-of select="256 + substring-after(@标识符_6B0A,'slideId')"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="256 + substring-after(@标识符_6B0A,'slide')"/>
									</xsl:otherwise>
								</xsl:choose>
								</xsl:attribute>
									<xsl:attribute name="r:id">
										<xsl:value-of select="concat('rId',@标识符_6B0A)"/>
									</xsl:attribute>
						</p:sldId>
					</xsl:for-each>
				</p:sldIdLst>
			</xsl:if>
			<!--<xsl:for-each select="/uof:UOF/uof:演示文稿/演:公用处理规则/演:页面设置集/演:页面设置/演:纸张">-->
		<xsl:for-each select="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面设置集_B670/规则:页面设置_B638/演:纸张_6BDD">
			 <p:sldSz>
				<xsl:choose>
              <!-- A3 -->
				<xsl:when test="@宽_C605='756' and @长_C604='1008'">
					<xsl:attribute name="cx">
						<xsl:value-of select="9601200"/>
					</xsl:attribute>
					<xsl:attribute name="cy">
						<xsl:value-of select="12801600"/>
					</xsl:attribute>
					<xsl:attribute name="type">A3</xsl:attribute>
				</xsl:when>
              <xsl:when test="@长_C604='756' and @宽_C605='1008'">

                <xsl:attribute name="cx">
                  <xsl:value-of select="12801600"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="9601200"/>
                </xsl:attribute>
                <xsl:attribute name="type">A3</xsl:attribute>
              </xsl:when>

              <!-- A4 -->
              <xsl:when test="@宽_C605='779' and @长_C604='540'">
                <xsl:attribute name="cx">
                  <xsl:value-of select="9906000"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="6858000"/>
                </xsl:attribute>
                <xsl:attribute name="type">A4</xsl:attribute>
              </xsl:when>
              <xsl:when test="@长_C604='779' and @宽_C605='540'">
               
                <xsl:attribute name="cx">
                  <xsl:value-of select="6858000"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="9906000"/>
                </xsl:attribute>
                <xsl:attribute name="type">A4</xsl:attribute>
              </xsl:when>

              <!-- B4 -->
              <xsl:when test="@宽_C605='639' and @长_C604='852'">
                <xsl:attribute name="cx">
                  <xsl:value-of select="8120063"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="10826750"/>
                </xsl:attribute>
                <xsl:attribute name="type">B4ISO</xsl:attribute>
              </xsl:when>
              <xsl:when test="@长_C604='639' and @宽_C605='852'">
               
                <xsl:attribute name="cx">
                  <xsl:value-of select="10826750"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="8120063"/>
                </xsl:attribute>
                <xsl:attribute name="type">B4ISO</xsl:attribute>
              </xsl:when>

              <!-- B5 -->
              <xsl:when test="@宽_C605='423' and @长_C604='564'">
                <xsl:attribute name="cx">
                  <xsl:value-of select="5376863"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="7169150"/>
                </xsl:attribute>
                <xsl:attribute name="type">B5ISO</xsl:attribute>
              </xsl:when>
              <xsl:when test="@长_C604='423' and @宽_C605='564'">                
                <xsl:attribute name="cx">
                  <xsl:value-of select="7169150"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="5376863"/>
                </xsl:attribute>
                <xsl:attribute name="type">B5ISO</xsl:attribute>
              </xsl:when>

              <!-- 信纸 8.5X11英寸 -->
              <xsl:when test="@宽_C605='720' and @长_C604='540'">
                <xsl:attribute name="cx">
                  <xsl:value-of select="9144000"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="6858000"/>
                </xsl:attribute>
                <xsl:attribute name="type">letter</xsl:attribute>
              </xsl:when>
              <xsl:when test="@长_C604='720' and @宽_C605='540'">               
                <xsl:attribute name="cx">
                  <xsl:value-of select="6858000"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="9144000"/>
                </xsl:attribute>
                <xsl:attribute name="type">letter</xsl:attribute>
              </xsl:when>
              <!-- 分类帐纸张 -->
              <!--2010.03.29 myx<xsl:when test="@uof:宽度='718' and @uof:高度='959'">-->
              <xsl:when test="@宽_C605='719' and @长_C604='959'">
                <xsl:attribute name="cx">
                  <xsl:value-of select="9134475"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="12179300"/>
                </xsl:attribute>
                <xsl:attribute name="type">ledger</xsl:attribute>
              </xsl:when>
              <!--2010.03.29 myx<xsl:when test="@uof:高度='718' and @uof:宽度='959'">-->
              <xsl:when test="@长_C604='719' and @宽_C605='959'">
                
                <xsl:attribute name="cx">
                  <xsl:value-of select="12179300"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="9134475"/>
                </xsl:attribute>
                <xsl:attribute name="type">ledger</xsl:attribute>
              </xsl:when>

              <!-- 35毫米幻灯片 -->
              <xsl:when test="@宽_C605='540' and @长_C604='809'">
                <xsl:attribute name="cx">
                  <xsl:value-of select="6858000"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="10287000"/>
                </xsl:attribute>
                <xsl:attribute name="type">35mm</xsl:attribute>
              </xsl:when>
              <xsl:when test="@长_C604='540' and @宽_C605='809'">
               
                <xsl:attribute name="cx">
                  <xsl:value-of select="10287000"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="6858000"/>
                </xsl:attribute>
                <xsl:attribute name="type">35mm</xsl:attribute>
              </xsl:when>

              <!-- 横幅 -->
              <xsl:when test="@宽_C605='72' and @长_C604='576'">
                <xsl:attribute name="cx">
                  <xsl:value-of select="914400"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="7315200"/>
                </xsl:attribute>
                <xsl:attribute name="type">banner</xsl:attribute>
              </xsl:when>
              <xsl:when test="@长_C604='72' and @宽_C605='576'">
               
                <xsl:attribute name="cx">
                  <xsl:value-of select="7315200"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="914400"/>
                </xsl:attribute>
                <xsl:attribute name="type">banner</xsl:attribute>
              </xsl:when>

              <xsl:otherwise>
                <xsl:attribute name="cx">
                  <xsl:value-of select="round(number(@宽_C605) * 12700)"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="round(number(@长_C604) * 12700)"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>
            <!--<xsl:choose>
						<xsl:when test="@uof:宽度='720' and @uof:高度='540'">
							<xsl:attribute name="type">letter</xsl:attribute>
						</xsl:when>
						<xsl:when test="@uof:宽度='959' and @uof:高度='719'">
							<xsl:attribute name="type">ledger</xsl:attribute>
						</xsl:when>
						<xsl:when test="@uof:宽度='1008' and @uof:高度='756'">
							<xsl:attribute name="type">A3</xsl:attribute>
						</xsl:when>
						<xsl:when test="@uof:宽度='779' and @uof:高度='540'">
							<xsl:attribute name="type">A4</xsl:attribute>
						</xsl:when>
						<xsl:when test="@uof:宽度='852' and @uof:高度='639'">
							<xsl:attribute name="type">B4ISO</xsl:attribute>
						</xsl:when>
						<xsl:when test="@uof:宽度='564' and @uof:高度='423'">
							<xsl:attribute name="type">B5ISO</xsl:attribute>
						</xsl:when>
						<xsl:when test="@uof:宽度='809' and @uof:高度='540'">
							<xsl:attribute name="type">35mm</xsl:attribute>
						</xsl:when>
						<xsl:when test="@uof:宽度='576' and @uof:高度='72'">
							<xsl:attribute name="type">banner</xsl:attribute>
						</xsl:when>
					</xsl:choose>-->
          </p:sldSz>
      
			</xsl:for-each>
      <!--liuyin 20121109 讲义备注大纲方向错误-->
      <xsl:choose>
        <xsl:when test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面设置集_B670/规则:页面设置_B638/演:纸张方向_6BE1='landscape'">
			    <p:notesSz cx="9144000" cy="6858000"/>
        </xsl:when>
        <xsl:when test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面设置集_B670/规则:页面设置_B638/演:纸张方向_6BE1='portrait'">
          <p:notesSz cx="6858000" cy="9144000"/>
        </xsl:when>
      </xsl:choose>
			<!--<xsl:for-each select="uof:演示文稿/演:公用处理规则/演:放映设置">-->
			<xsl:for-each select="uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:放映设置_B653">
				<p:custShowLst>
					<xsl:for-each select="规则:幻灯片序列_B654">
						<xsl:if test="@是否自定义_B657='true' or @是否自定义_B657='on' or @是否自定义_B657='1'">
							<p:custShow>
								<xsl:for-each select="@名称_B656">
									<xsl:attribute name="name">
										<xsl:value-of select="."/>
									</xsl:attribute>
								</xsl:for-each>
								<!--p:custShow 的 @id 取值类型为 "unsignedInt" 所以不能用原来的 @标识符 给 @id 赋值 取其位置-->
								<xsl:attribute name="id">
									<xsl:value-of select="position()"/>
								</xsl:attribute>
								<p:sldLst>
									<xsl:if test="normalize-space(.)">
										<xsl:call-template name="sld_lst">
											<xsl:with-param name="left_sldLst" select="concat(normalize-space(.),' ')"/>
										</xsl:call-template>
									</xsl:if>
								</p:sldLst>
							</p:custShow>
						</xsl:if>
					</xsl:for-each>
				</p:custShowLst>
			</xsl:for-each>
			<p:defaultTextStyle>
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
			</p:defaultTextStyle>
		</p:presentation>
	</xsl:template>
	<!--幻灯片放映设置  从字符串中提取幻灯片存放地址信息-->
	<xsl:template name="sld_lst">
		<xsl:param name="left_sldLst"/>
		<p:sld>
			<xsl:attribute name="r:id">
        <!--2010.05.04<xsl:value-of select="concat('rId',substring-before($left_sldLst,' '))"/>-->
        <xsl:variable name="slideId" select="substring-before($left_sldLst,' ')"/>
        <xsl:choose>
          <xsl:when test="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F[@标识符_6B0A=$slideId]">
            <xsl:value-of select="concat('rId',$slideId)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:for-each select="/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F">
              <xsl:if test="position()=number($slideId)+1">
                <xsl:value-of select="concat('rId',@标识符_6B0A)"/>
              </xsl:if>
            </xsl:for-each>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
		</p:sld>
		<xsl:if test="substring-after($left_sldLst,' ')">
			<xsl:call-template name="sld_lst">
				<xsl:with-param name="left_sldLst" select="substring-after($left_sldLst,' ')"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="presentation.xml.rels">
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>
			<!--<xsl:for-each select="uof:演示文稿/演:主体/演:母版集/演:母版[@演:类型='slide' or not(@演:类型)]"-->
			<xsl:for-each select="uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='slide' or not(@类型_6BEA)]">
			<Relationship>
				<xsl:attribute name="Id">
					<xsl:value-of select="'rId'"/>
					<xsl:value-of select="@标识符_6BE8"/>
				</xsl:attribute>
				<xsl:attribute name="Type">http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster</xsl:attribute>
				<xsl:attribute name="Target">
					<xsl:value-of select="'slideMasters/'"/>
					<xsl:value-of select="@标识符_6BE8"/>
					<xsl:value-of select="'.xml'"/>
				</xsl:attribute>
			</Relationship>
		</xsl:for-each>
			<xsl:for-each select="uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F">
				<Relationship>
					<xsl:attribute name="Id">
						<xsl:value-of select="'rId'"/>
						<xsl:value-of select="@标识符_6B0A"/>
					</xsl:attribute>
					<xsl:attribute name="Type">http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide</xsl:attribute>
					<xsl:attribute name="Target">
						<xsl:value-of select="'slides/'"/>
						<xsl:value-of select="@标识符_6B0A"/>
						<xsl:value-of select="'.xml'"/>
					</xsl:attribute>
				</Relationship>
			</xsl:for-each>
			<xsl:if test="not(uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='notes']) and uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/演:幻灯片备注_6B1D">
				<!--<xsl:if test="not(uof:演示文稿/演:主体/演:母版集/演:母版[@演:类型='notes']) and uof:演示文稿/演:主体/演:幻灯片集/演:幻灯片/演:幻灯片备注">-->
				<xsl:for-each select="uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/演:幻灯片备注_6B1D">
					<Relationship>
						<xsl:attribute name="Id">
							<xsl:value-of select="'rId'"/>
							<xsl:value-of select="substring-after(.//@图形引用_C62E,'Obj')"/>
						</xsl:attribute>
						<xsl:attribute name="Type">http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster</xsl:attribute>
						<xsl:attribute name="Target">
							<xsl:value-of select="'notesMasters/notesMaster'"/>
							<xsl:value-of select="substring-after(.//@图形引用_C62E,'Obj')"/>
							<xsl:value-of select="'.xml'"/>
						</xsl:attribute>
					</Relationship>
				</xsl:for-each>
			</xsl:if>
			<!--4月16日jjy改 删除没有“演:类型”属性母版的配色方案引用-->
			<!--4月8 黎美秀修改 添加[1]-->
			<xsl:for-each select="uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA][1]">
				<Relationship>
					<!-- 09.10.12 马有旭 修改 -->
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
				<xsl:attribute name="Type">http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme</xsl:attribute>
				<xsl:attribute name="Target">
					<xsl:choose>
						<xsl:when test="contains(@配色方案引用_6BEB,'clr_')">
							<xsl:value-of select="concat('theme/theme',substring-after(@配色方案引用_6BEB,'clr_'),'.xml')"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="concat('theme/theme',substring-after(@配色方案引用_6BEB,'Id'),'.xml')"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				</Relationship>
			</xsl:for-each>
      <!--luowentian 2011-6-1先去掉讲义母版，下期改善>-->
			<xsl:for-each select="uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='handout']">
				<Relationship>
					<xsl:attribute name="Id">
						<xsl:choose>
							<xsl:when test="contains(@标识符_6BE8,'ID')">
								<xsl:value-of select="concat('rId',substring-after(@标识符_6BE8,'ID'))"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="concat('rId',@标识符_6BE8)"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
					<xsl:attribute name="Type">http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster</xsl:attribute>
					<xsl:attribute name="Target">
						<xsl:choose>
							<xsl:when test="contains(@标识符_6BE8,'ID')">
								<xsl:value-of select="'handoutMasters/handoutMaster'"/>
								<xsl:value-of select="substring-after(@标识符_6BE8,'ID')"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'handoutMasters/'"/>
								<xsl:value-of select="@标识符_6BE8"/>
							</xsl:otherwise>
						</xsl:choose>
						<xsl:value-of select="'.xml'"/>
					</xsl:attribute>
				</Relationship>
			</xsl:for-each>
			<xsl:for-each select="uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='notes']">
			
				<Relationship>
					<xsl:attribute name="Id">
						<xsl:choose>
							<xsl:when test="contains(@标识符_6BE8,'ID')">
								<xsl:value-of select="concat('rId',substring-after(@标识符_6BE8,'ID'))"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="concat('rId',@标识符_6BE8)"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				<xsl:attribute name="Type">http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster</xsl:attribute>
				<xsl:attribute name="Target">
					<xsl:choose>
						<xsl:when test="contains(@标识符_6BE8,'ID')">
							<xsl:value-of select="'notesMasters/notesMaster'"/>
							<xsl:value-of select="substring-after(@标识符_6BE8,'ID')"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'notesMasters/'"/>
							<xsl:value-of select="@标识符_6BE8"/>
						</xsl:otherwise>
					</xsl:choose>
					<xsl:value-of select="'.xml'"/>
				</xsl:attribute>
				</Relationship>
			</xsl:for-each>
			<!--下边3个未验证-->
			<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>
			<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>
			<!--添加批注引用 李娟 2012.02.25-->
			<xsl:if  test="//规则:用户集_B667">

				<Relationship Id="rIdcomment" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml">
					
				</Relationship>
			</xsl:if>
	</Relationships>
	</xsl:template>
	<!-- 09.10.17 马有旭 添加 -->
	<!--修改标签 李娟 11.12.30-->
	<xsl:template name="custShow">
		<xsl:param name="id"/>
		<!--<xsl:for-each select="//演:幻灯片序列">-->
		<xsl:for-each select="//规则:幻灯片序列_B654">
			<xsl:if test="@标识符_B655=$id">
				<!--<xsl:if test="@演:标识符=$id">-->
				<p:custShow>
					<xsl:attribute name="id">
						<xsl:value-of select="position()"/>
					</xsl:attribute>
				</p:custShow>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="sequent">
		<xsl:param name="stid"/>
		<xsl:param name="endid"/>
		<!--<xsl:for-each select="//演:幻灯片">-->
		<xsl:for-each select="//演:幻灯片_6C0F">
			<xsl:if test="@标识符_6B0A=$stid">
			<!--<xsl:if test="@标识符_6B0A=$stid">-->
				<xsl:attribute name="st">
					<xsl:value-of select="position()"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test="@标识符_6B0A=$endid">
				<xsl:attribute name="end"><xsl:value-of select="position()"/></xsl:attribute>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>
	<!-- 09.10.19 马有旭 添加 -->
	<!--<xsl:template match="/uof:UOF/uof:扩展区/uof:扩展[uof:扩展内容/uof:路径='uof/演示文稿']" mode="showType">-->
	<xsl:template match="/uof:UOF/uof:扩展区/扩展:扩展区_B200/扩展:扩展_B201[扩展:内容_B206/扩展:路径_B205='uof/演示文稿']" mode="showType">
		<xsl:choose>
			<xsl:when test="uof:扩展内容/uof:内容/uof:播放类型">
				<xsl:choose>
					<xsl:when test="uof:扩展内容/uof:内容/uof:播放类型/@showScrollbar">
						<p:browse>
							<xsl:attribute name="showScrollbar"><xsl:value-of select="uof:扩展内容/uof:内容/uof:播放类型/@showScrollbar"/></xsl:attribute>
						</p:browse>
					</xsl:when>
					<xsl:otherwise>
						<p:kiosk/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<p:present/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>
