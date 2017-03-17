<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
        xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" 
        xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
				xmlns:fo="http://www.w3.org/1999/XSL/Format" 
				xmlns:dc="http://purl.org/dc/elements/1.1/" 
				xmlns:dcterms="http://purl.org/dc/terms/" 
				xmlns:dcmitype="http://purl.org/dc/dcmitype/"
				xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
				xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                
        xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
				xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" 
				xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" 
				xmlns:p="http://purl.oclc.org/ooxml/presentationml/main" 
                 
				xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
				xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
        xmlns:uof="http://schemas.uof.org/cn/2009/uof"
				xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
				xmlns:演="http://schemas.uof.org/cn/2009/presentation"
				xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
				xmlns:图="http://schemas.uof.org/cn/2009/graph">
	<xsl:import href="ph.xsl"/>
	<xsl:import href="fill.xsl"/>
	<xsl:import href="animation.xsl"/>
	<xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="p:sld" mode="ph">
    <!--2010-11-8 罗文甜：删除幻灯片下面的“配色方案引用”属性-->
		<!--是否显示背景,是否显示折叠大纲,是否显示配色方案引用 没有添加 李娟 11.11.11-->
		<演:幻灯片_6C0F>
			<xsl:variable name="sldID" select="@id"/>
			<xsl:if test="p:cSld/@name">
				<xsl:attribute name="名称_6B0B">
					<xsl:value-of select="p:cSld/@name"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:attribute name="标识符_6B0A">
        <!--liuyangyang 2014-12-31 修改幻灯片编号，与永中生成编号前置相同-->
				<xsl:value-of select="concat('slideId',substring-after(substring-before($sldID,'.xml'),'slide'))"/>
        <!--<xsl:value-of select="substring-before($sldID,'.xml')"/>-->
        <!--end-->
			</xsl:attribute>
			<xsl:variable name="sldLayoutID" select="@sldLayoutID"/>
			<xsl:variable name="sldMasterID" select="substring-before(//p:sldMaster[p:sldLayout/@id=$sldLayoutID]/@id,'.xml')"/>
			<xsl:attribute name="母版引用_6B26">
				<xsl:value-of select="$sldMasterID"/>
			</xsl:attribute>
			
			<xsl:attribute name="页面版式引用_6B27">
				<xsl:value-of select="substring-before($sldLayoutID,'.xml')"/>
			</xsl:attribute>
			<xsl:choose>
				<xsl:when test="@show='0' or @show='false' or show='off'">
					<xsl:attribute name="是否显示_6B28">false</xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
					<xsl:attribute name="是否显示_6B28">true</xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>
			<!--<xsl:if test="@show='0' or @show='false' or show='off'">
				<xsl:attribute name="是否显示_6B28">false</xsl:attribute>
			
			</xsl:if>-->
      
      <!-- 2014-05-05, tangjiang, 修复页眉页脚引用引起的UOF内容不显示 Start-->
      <!--<xsl:if test="//p:hf">-->
      <xsl:if test="/p:presentation/p:sldMaster//p:hf">
			<xsl:attribute name="页眉页脚引用_6C15">
				<xsl:value-of select="generate-id(/p:presentation/p:notes)"/>
			</xsl:attribute>
			</xsl:if>
      <xsl:attribute name="页眉页脚引用_6C15">
        <xsl:value-of select="generate-id(.)"/>
      </xsl:attribute>
      <!-- end 2014-05-05, tangjiang, 修复页眉页脚引用引起的UOF内容不显示 -->
      
      <!--2011-3-17罗文甜：背景对象-->
			<xsl:choose>
				<xsl:when test="@showMasterSp='0'">
					<xsl:attribute name="是否显示背景对象_6B2A">false</xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
					<xsl:attribute name="是否显示背景对象_6B2A">true</xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>
			<!--锚点-->
			<!--添加批注的锚点-->
      <!--guoyongbin 修改批注锚点添加-->
			<xsl:if test="//p:cmAuthorLst and  ./@commentID">
				<uof:锚点_C644>
					<xsl:attribute name="图形引用_C62E">
						<xsl:value-of select="substring-before(./@commentID,'.xml')"/>
					</xsl:attribute>
				
						<uof:位置_C620>
							<uof:水平_4106>
								<uof:绝对_4107 值_4108="540.70605"/>
							</uof:水平_4106>
							<uof:垂直_410D>
								<uof:绝对_4107 值_4108="9.285573"/>
							</uof:垂直_410D>
						</uof:位置_C620>
						<uof:大小_C621 长_C604="35.699997" 宽_C605="173.5688"/>
					
				</uof:锚点_C644>
			</xsl:if>
      <!--end guoyongbin-->
      <!--2011-01-04 罗文甜：修改BUG，幻灯片中对页眉、页脚、编号、日期的图形引用去掉--><!--jdslfkjljsfdlkfjlsdajflijuasndfajflj·············@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@lijuan-->
		   <!--liuyangyang 2014-12-22 添加对公式锚点的转换-->
      <xsl:for-each select="p:cSld/p:spTree/p:sp[(.//p:ph/@type!='hdr' and .//p:ph/@type!='sldNum' and .//p:ph/@type!='dt' and .//p:ph/@type!='ftr') or not(.//p:ph) or not(.//p:ph/@type)] 
                    | p:cSld/p:spTree/p:cxnSp|p:cSld/p:spTree/p:grpSp|p:cSld/p:spTree/p:pic|p:cSld/p:spTree/p:graphicFrame | p:cSld/p:spTree/mc:AlternateContent">
			<!--<xsl:for-each select="p:cSld/p:spTree/p:sp[(.//p:ph/@type!='hdr' and .//p:ph/@type!='sldNum' and .//p:ph/@type!='dt' and .//p:ph/@type!='ftr') or not(.//p:ph)]|p:cSld/p:spTree/p:cxnSp|p:cSld/p:spTree/p:grpSp|p:cSld/p:spTree/p:pic|p:cSld/p:spTree/p:graphicFrame">-->
			<xsl:choose>
					<xsl:when test="name(.)='p:pic'">
						<xsl:apply-templates select="." mode="sldph">
							<xsl:with-param name="layoutref" select="$sldLayoutID"/>
						</xsl:apply-templates>
					</xsl:when>
        
					<xsl:otherwise>
						<xsl:apply-templates select="." mode="ph"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
		
			<xsl:choose>
				<xsl:when test="./p:cSld/p:bg">
					<xsl:call-template name="backgroundFill"/>
				</xsl:when>
        <xsl:otherwise>
          <!--liuyangyang 2015-04-09 修改背景填充错误 start-->
          <xsl:choose>
            <xsl:when test="preceding-sibling::p:sldMaster[substring-before(./@id,'.xml')=$sldMasterID]/p:sldLayout[./@id=$sldLayoutID]/p:cSld/p:bg">
              <xsl:for-each select="preceding-sibling::p:sldMaster[substring-before(./@id,'.xml')=$sldMasterID]/p:sldLayout[./@id=$sldLayoutID]">
                <!--<xsl:if test="preceding-sibling::p:sldMaster[substring-before(./@id,'.xml')=$sldMasterID]/p:cSld/p:bg">
						<xsl:for-each select="preceding-sibling::p:sldMaster[substring-before(./@id,'.xml')=$sldMasterID]">-->
                <xsl:call-template name="backgroundFill"/>
              </xsl:for-each>
            </xsl:when>
            <xsl:otherwise>
              <xsl:for-each select="preceding-sibling::p:sldMaster[substring-before(./@id,'.xml')=$sldMasterID]">
                <!--<xsl:if test="preceding-sibling::p:sldMaster[substring-before(./@id,'.xml')=$sldMasterID]/p:cSld/p:bg">
						<xsl:for-each select="preceding-sibling::p:sldMaster[substring-before(./@id,'.xml')=$sldMasterID]">-->
                <xsl:call-template name="backgroundFill"/>
              </xsl:for-each>
            </xsl:otherwise>
          </xsl:choose>
          <!--end  liuyangyang 2015-04-09 修改背景填充错误-->
        </xsl:otherwise>
			</xsl:choose>

      <!--start liuyin 20121211 修改音视频丢失-->
			<xsl:if test=".//p:timing">
        <!--<xsl:if test=".//p:timing">-->
      <!--end liuyin 20121211 修改音视频丢失-->

			<xsl:apply-templates select=".//p:timing"/>
			</xsl:if>
      <!--动画-->
				
			<!--<演:动画_6B1A>
      <xsl:apply-templates select="//p:timing"/>
				</演:动画_6B1A>-->
			<!--备注-->
			<!--10.16 黎美秀修改 增加判断
      
      <xsl:apply-templates select="following-sibling::p:notes[1]" mode="note"/>
      -->
			<xsl:if test="following-sibling::rel:Relationships[1]/rel:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/notesSlide']">
				<xsl:apply-templates select="following-sibling::p:notes[1]" mode="note"/>
			</xsl:if>
			<!--切换2012-4-24李娟，修改-->
			
			<xsl:choose>
				<xsl:when test=".//p:transition">
					<xsl:for-each select=".//p:transition">
						<xsl:if test="position()='1'">
							<xsl:apply-templates select="."/>
						</xsl:if>
						</xsl:for-each>
				</xsl:when>
				<!--<xsl:when test=".//mc:AlternateContent">
					<xsl:for-each select=".//mc:AlternateContent">
						
							<xsl:apply-templates select="."/>
						
					</xsl:for-each>
				</xsl:when>-->
			</xsl:choose>
			<!--<xsl:if test=".//p:transition">
				<xsl:for-each select=".//p:transition">
				 <xsl:if test="position()='1'">
				 <xsl:apply-templates select="."/>
			</xsl:if>-->
	       
				
			<!--</xsl:if>-->
		</演:幻灯片_6C0F>
	</xsl:template>
	<xsl:template match="p:notes" mode="note">
		<演:幻灯片备注_6B1D>
			<xsl:attribute name="备注母版引用_6B2D">
				<xsl:value-of select="substring-before(substring-after(following-sibling::rel:Relationships[1]/rel:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/notesMaster']/@Target,'notesMasters/'),'.xml')"/>
			</xsl:attribute>
      <!--2012-12-20, liqiuling, 解决OOXML到UOF备注页页脚丢失  start -->
      <xsl:attribute name="页眉页脚引用_6C17">
        <xsl:value-of select="generate-id(.)"/>
      </xsl:attribute>
      <!--end-->
			<!--<xsl:if test="">
				<xsl:attribute name="页眉页脚引用_6C17">
					
				</xsl:attribute>
			</xsl:if>-->
			<!--      
      10.1.15 黎美秀添加 备注的图形等
      <xsl:apply-templates select="p:cSld/p:spTree/p:pic" mode="ph"/>
       -->
			<!--      
      2010.2.5 黎美秀修改 处理图形层次问题，改变调用顺序
      原：      
     <xsl:apply-templates select="p:cSld/p:spTree/p:sp" mode="ph"/>
      <xsl:apply-templates select="p:cSld/p:spTree/p:grpSp" mode="ph"/>
      
      <xsl:apply-templates select="p:cSld/p:spTree/p:cxnSp" mode="ph"/>
      <xsl:apply-templates select="p:cSld/p:spTree/p:pic" mode="ph"/>
       -->
      <!--2011-02-22 罗文甜：修改BUG，备注中对页眉、页脚、编号、日期的图形引用去掉-->
			<xsl:for-each select="p:cSld/p:spTree/p:sp[(.//p:ph/@type!='hdr' and .//p:ph/@type!='sldNum' and .//p:ph/@type!='dt' and .//p:ph/@type!='ftr') or not(.//p:ph)]|p:cSld/p:spTree/p:cxnSp|p:cSld/p:spTree/p:grpSp|p:cSld/p:spTree/p:pic">
				<xsl:apply-templates select="." mode="ph"/>                                   
			</xsl:for-each>
			<!-- 李娟 修改背景色 11.12.19-->
			<xsl:for-each select="ancestor::p:presentation">
				<xsl:if test="./p:sldMaster/p:cSld/p:bg">
					<演:背景_6B2C>
						<xsl:choose>
							<xsl:when test="./p:sldMaster/p:cSld/p:bg/p:bgPr">
								<xsl:apply-templates select="p:bgPr"/>
							</xsl:when>
							<xsl:when test="./p:sldMaster/p:cSld/p:bg/p:bgRef">
								<xsl:if test ="./p:sldMaster/p:cSld/p:bg/p:bgRef/a:schemeClr">
								<图:颜色_8004>
									<xsl:variable name="schemeClrName">
										<xsl:value-of select="./p:sldMaster/p:cSld/p:bg/p:bgRef/a:schemeClr/@val"/>
									</xsl:variable>
									<xsl:choose>
										<xsl:when test="$schemeClrName='bg1'">
											<xsl:value-of select="./p:sldMaster/p:clrMap/@bg1"/>
										</xsl:when>
										<xsl:when test="$schemeClrName='bg2'">
											<xsl:value-of select="./p:sldMaster/p:clrMap/@bg2"/>
										</xsl:when>
										<xsl:when test="$schemeClrName='tx1'">
											<xsl:value-of select="./p:sldMaster/p:clrMap/@tx1"/>
										</xsl:when>
										<xsl:when test="$schemeClrName='tx2'">
											<xsl:value-of select="./p:sldMaster/p:clrMap/@tx2"/>
										</xsl:when>
										<xsl:when test="$schemeClrName='accent1'">
											<xsl:value-of select="./p:sldMaster/p:clrMap/@accent1"/>
										</xsl:when>
										<xsl:when test="$schemeClrName='accent2'">
											<xsl:value-of select="./p:sldMaster/p:clrMap/@accent2"/>
										</xsl:when>
										<xsl:when test="$schemeClrName='accent3'">
											<xsl:value-of select="./p:sldMaster/p:clrMap/@accent3"/>
										</xsl:when>
										<xsl:when test="$schemeClrName='accent4'">
											<xsl:value-of select="./p:sldMaster/p:clrMap/@accent4"/>
										</xsl:when>
										<xsl:when test="$schemeClrName='accent5'">
											<xsl:value-of select="./p:sldMaster/p:clrMap/@accent5"/>
										</xsl:when>
										<xsl:when test="$schemeClrName='accent6'">
											<xsl:value-of select="./p:sldMaster/p:clrMap/@accent6"/>
										</xsl:when>
										<xsl:when test="$schemeClrName='hlink'">
											<xsl:value-of select="./p:sldMaster/p:clrMap/@hlink"/>
										</xsl:when>
										<xsl:when test="$schemeClrName='folHlink'">
											<xsl:value-of select="./p:sldMaster/p:clrMap/@folHlink"/>
										</xsl:when>
									</xsl:choose>
								</图:颜色_8004>
								</xsl:if>
							</xsl:when>
						</xsl:choose>
					</演:背景_6B2C>
				</xsl:if>
			</xsl:for-each>
		</演:幻灯片备注_6B1D>
	</xsl:template>
	<xsl:template  match="mc:AlternateContent">
    <!--这个模板没有用到  liqiuling 2013-1-15 -->
			<!--<xsl:if test="mc:Choice">-->
		<xsl:for-each select="mc:Choice/p:transition">
			<演:切换_6B1F>
				<演:效果_6B20>
					<xsl:choose>
						<xsl:when test="p14:flash">
							<xsl:value-of select="'newsflash'"/>
						</xsl:when>
						<xsl:when test="p14:doors[@dir='vert']">
							<xsl:value-of select="'split-vertical-out'"/>
						</xsl:when>
						<xsl:when test="p14:pan[@dir='u']">
							<xsl:value-of select="'push-up'"/>
						</xsl:when>
						<xsl:otherwise>'none'</xsl:otherwise>
					</xsl:choose>
				</演:效果_6B20>
				<演:速度_6B21>
					<xsl:choose>
						<xsl:when test="@spd='med'">middle</xsl:when>
						<xsl:when test="name(@spd)=''">fast</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="@spd"/>
						</xsl:otherwise>
					</xsl:choose>
				</演:速度_6B21>
				<xsl:apply-templates select="p:sndAc"/>
				<演:方式_6B23>
					<演:单击鼠标_6B24>'true'</演:单击鼠标_6B24>
				</演:方式_6B23>
			</演:切换_6B1F>
		</xsl:for-each>
			<!--</xsl:if>-->
		
	</xsl:template>


	<!--切换-->
  <!--2013-1-15 修复幻灯片切换效果丢失 liqiuling start-->
	<xsl:template match="p:transition">
		<演:切换_6B1F>
			<演:效果_6B20>
			<!--<xsl:attribute name="演:效果_6B20">-->
				<xsl:choose>

              <xsl:when test="parent::mc:Choice and p14:flash">
                <!--<xsl:value-of select="'newsflash'"/>-->
                <xsl:value-of select="'fade-through-black'"/>
              </xsl:when>
              <xsl:when test="parent::mc:Choice and p14:doors[@dir='vert']">
                <xsl:value-of select="'split-vertical-out'"/>
              </xsl:when>
              <xsl:when test="parent::mc:Choice and p14:pan[@dir='u']">
                <xsl:value-of select="'push-up'"/>
              </xsl:when>
              <xsl:when test="parent::mc:Choice and (@p14:dur='700' or p:fade or p14:window)">
                <xsl:value-of select="'fade-smoothly'"/>
              </xsl:when>
              <xsl:when test="parent::mc:Choice and p14:flash">

                <xsl:value-of  select="'wedge'"/>
              </xsl:when>
              <xsl:when test="parent::mc:Choice and p:checker">
                <xsl:value-of select="'checkerboard-across'"/>
              </xsl:when>
              <xsl:when test="parent::mc:Choice and p:cut">
                <xsl:value-of select="'cut'"/>
              </xsl:when>


          <xsl:when test="parent::mc:Choice and p14:reveal">
            <xsl:value-of select="'cut-through-black'"/>
          </xsl:when>
          <xsl:when test="parent::mc:Choice and (p14:flip or p14:switch)">
            <xsl:value-of select="'box-out'"/>
          </xsl:when>
          <xsl:when test="parent::mc:Choice and p14:prism">
            <xsl:value-of select="'shape-diamond'"/>
          </xsl:when>
   					<xsl:when test="p:blinds">
						<xsl:choose>
							<xsl:when test="p:blinds/@dir='vert'">blinds-vertical</xsl:when>
							<xsl:otherwise>blinds-horizontal</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="p:checker">
						<xsl:choose>
							<xsl:when test="p:checker/@dir='vert'">checkerboard-down</xsl:when>
							<xsl:otherwise>checkerboard-across</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="p:circle">shape-circle</xsl:when>
					<xsl:when test="p:dissolve">dissolve</xsl:when>
					<xsl:when test="p:comb">
						<xsl:choose>
							<xsl:when test="p:comb/@dir='vert'">comb vertical</xsl:when>
							<xsl:otherwise>comb-horizontal</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<!--cover在oox中只给了一个默认值‘l’-->
					<xsl:when test="p:cover">
						<xsl:choose>
							<xsl:when test="p:cover/@dir='r'">cover-right</xsl:when>
							<xsl:when test="p:cover/@dir='d'">cover-down</xsl:when>
							<xsl:when test="p:cover/@dir='lu'">cover-left-up</xsl:when>
							<xsl:when test="p:cover/@dir='ru'">cover-right-up</xsl:when>
							<xsl:when test="p:cover/@dir='rd'">cover-right-down</xsl:when>
							<xsl:when test="p:cover/@dir='ld'">cover-left-down</xsl:when>
							<xsl:when test="p:cover/@dir='u'">cover-up</xsl:when>
							<xsl:otherwise>cover-left</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="p:cut">
						<xsl:choose>
							<xsl:when test="p:cut/@thruBlk='true' or p:cut/@thruBlk='1'or p:cut/@thruBlk='on'">cut-through-black</xsl:when>
							<xsl:otherwise>cut</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="p:diamond">shape-diamond</xsl:when>
					<xsl:when test="p:fade">
						<xsl:choose>
							<xsl:when test="p:fade/@thruBlk='true' or p:fade/@thruBlk='1' or p:fade/@thruBlk='on'">fade-through-black</xsl:when>
							<xsl:otherwise>fade-smoothly</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="p:newsflash">newsflash</xsl:when>
					<xsl:when test="p:plus">shape-plus</xsl:when>
					<!--xsl:when test="pull">
					</xsl:when-->
					 <!--罗文甜 增加lu ld ru rd对应-->
					<xsl:when test="p:push|p:pull">
						<xsl:choose>
							<xsl:when test="p:push/@dir='u' or p:pull/@dir='u'">push-up</xsl:when>
							<xsl:when test="p:push/@dir='r' or p:pull/@dir='r'">push-right</xsl:when>
							<xsl:when test="p:push/@dir='d' or p:pull/@dir='d'">push-down</xsl:when>
              <xsl:when test="p:push/@dir='l' or p:pull/@dir='l'">push left</xsl:when>
              <xsl:when test="p:pull/@dir='lu'">cover-left-up</xsl:when>
              <xsl:when test="p:pull/@dir='ld'">cover-left-down</xsl:when>
              <xsl:when test="p:pull/@dir='ru'">cover-right-up</xsl:when>
              <!--2011-5-12 luowentian-->
              <xsl:when test="p:pull/@dir='rd'">cover-right-down</xsl:when>
							<xsl:otherwise>push-left</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="p:random">random-transition</xsl:when>
					<xsl:when test="p:randomBar">
						<xsl:choose>
							<xsl:when test="p:randomBar/@dir='vert'">random-bars-vertical</xsl:when>
							<xsl:otherwise>random-bars-horizontal</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="p:split">
						<xsl:choose>
							<xsl:when test="p:split/@orient='vert'">
								<xsl:choose>
									<xsl:when test="p:split/@dir='in'">split-vertical-in</xsl:when>
									<xsl:otherwise>split-vertical-out</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
								<xsl:choose>
									<xsl:when test="p:split/@dir='in'">split-horizontal-in</xsl:when>
									<xsl:otherwise>split-horizontal-out</xsl:otherwise>
								</xsl:choose>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="p:strips">
						<xsl:choose>
							<xsl:when test="p:strips/@dir='ru'">strips-right-up</xsl:when>
							<xsl:when test="p:strips/@dir='ld'">strips-left-down</xsl:when>
							<xsl:when test="p:strips/@dir='rd'">strips-right-down</xsl:when>
							<xsl:otherwise>strips-left-up</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="p:wedge">wedge</xsl:when>
					<xsl:when test="p:wheel">
						<xsl:choose>
							<xsl:when test="p:wheel/@spokes='1'">
								<xsl:value-of select="'wheel-clockwise – 1 spoke'"/>
							</xsl:when>
							<xsl:when test="p:wheel/@spokes='2'">wheel-clockwise–2spoke</xsl:when>
							<xsl:when test="p:wheel/@spokes='3'">wheel-clockwise–3spoke</xsl:when>
							<xsl:when test="p:wheel/@spokes='8'">wheel-clockwise–8spoke</xsl:when>
							<xsl:otherwise>wheel-clockwise-4spoke</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="p:wipe">
						<xsl:choose>
							<xsl:when test="p:wipe/@dir='u'">wipe-up</xsl:when>
							<xsl:when test="p:wipe/@dir='r'">wipe-right</xsl:when>
							<xsl:when test="p:wipe/@dir='d'">wipe-down</xsl:when>
							<xsl:otherwise>wipe-left</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<!--这项转的不知道对不对-->
					<xsl:when test="p:zoom">
						<xsl:choose>
							<xsl:when test="p:zoom/@dir='in'">box-in</xsl:when>
							<xsl:otherwise>box-out</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>none</xsl:otherwise>
				</xsl:choose>
				<!--	</xsl:otherwise>
				</xsl:choose>-->
				
			<!--</xsl:attribute>-->
			</演:效果_6B20>
			<!--<xsl:attribute name="演:速度_6B21">-->
			<演:速度_6B21>
				<xsl:choose>
					<xsl:when test="@spd='med'">middle</xsl:when>
					<xsl:when test="name(@spd)=''">fast</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="@spd"/>
					</xsl:otherwise>
				</xsl:choose>
			</演:速度_6B21>
			<xsl:apply-templates select="p:sndAc"/>
			<演:方式_6B23>
				<演:单击鼠标_6B24>
					<xsl:choose>
						<xsl:when test="@advClick='0' or @advClick='false' or @advClick='off'">false</xsl:when>
						<xsl:otherwise>true</xsl:otherwise>
					</xsl:choose>
				</演:单击鼠标_6B24>
				<!--时间间隔-->
				<xsl:apply-templates select="@advTm"/>
			</演:方式_6B23>
		</演:切换_6B1F>
	</xsl:template>
  <!--2013-1-15 修复幻灯片切换效果丢失 liqiuling end-->
	<xsl:template match="p:sndAc">
		<!--OOX中声音是嵌入的，对应永中中是嵌入的二进制代码 切换声音循环放映无对应-->
		<!-- 2010.03.30 myx-->
    <!--2010-11-1罗文甜 增加声音是否循环播放-->
		<!---->
		<!--<xsl:if test="p:endSnd">
			
			--><!--<演:声音 uof:locID="p0061" uof:attrList="预定义声音 自定义声音 是否循环播放" 演:自定义声音="stop previous sound" 演:是否循环播放="false"/>--><!--
		</xsl:if>-->
		<!-- 2010.03.30 马有旭 添加 切换声音 -->
		<xsl:if test="p:stSnd/p:snd">
			<演:声音_6B17>
				<xsl:attribute name="自定义声音_C632">
					<xsl:variable name="tgt">
						<xsl:value-of select="p:stSnd/p:snd/@r:embed"/>
					</xsl:variable>
					<xsl:variable name="sldid">
						<xsl:value-of select="ancestor::p:sld/@id"/>
					</xsl:variable>
          <xsl:value-of select="concat(substring-before($sldid,'.xml'), $tgt)"/>
				</xsl:attribute>
        <!--2010-11-1罗文甜：增加声音是否循环播放-->
        <xsl:attribute name="是否循环播放_C633">
          <xsl:choose>
            <xsl:when test="p:stSnd/@loop='1'">true</xsl:when>
            <xsl:otherwise>false</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
			</演:声音_6B17>
		</xsl:if>
	</xsl:template>
  <!--2013-1-15 修复幻灯片切换时间转换不正确 liqiuling start-->
	<xsl:template match="@advTm">
    <xsl:variable name="time">
      <xsl:value-of select="."/>
    </xsl:variable>
		<演:时间间隔_6B25>
			<xsl:value-of select="concat('P0Y0M0DT0H0M',($time div 1000),'S')"/>
		</演:时间间隔_6B25>
	</xsl:template>
  <!--2013-1-15 修复幻灯片切换时间转换不正确 liqiuling end-->
	<!-- 09.12.04 马有旭 添加 计算动作路径值-->
	<xsl:template name="pathvalue">
		<xsl:param name="value"/>
		<xsl:param name="pagex"/>
		<xsl:param name="pagey"/>
		<xsl:param name="count"/>
		<xsl:if test="contains($value,' ')">
			<xsl:variable name="var" select="substring-before($value,' ')"/>
			<xsl:if test="$var=' '">
				<xsl:call-template name="pathvalue">
					<xsl:with-param name="value" select="substring-after($value,' ')"/>
					<xsl:with-param name="pagex" select="$pagex"/>
					<xsl:with-param name="pagey" select="$pagey"/>
					<xsl:with-param name="count" select="number($count)"/>
				</xsl:call-template>
			</xsl:if>
			<xsl:choose>
				<xsl:when test="$var='M' or $var='Z' or $var='C' or $var='L' or $var='E'">
					<xsl:value-of select="concat($var,' ')"/>
					<xsl:call-template name="pathvalue">
						<xsl:with-param name="value" select="substring-after($value,' ')"/>
						<xsl:with-param name="pagex" select="$pagex"/>
						<xsl:with-param name="pagey" select="$pagey"/>
						<xsl:with-param name="count" select="$count"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<xsl:variable name="newvalue">
						<xsl:choose>
							<xsl:when test="($count mod 2)=0">
								<xsl:value-of select="number($var)*number($pagex)"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="number($var)*number($pagey)"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:variable>
					<xsl:if test="string($newvalue)!='NaN'">
						<xsl:value-of select="$newvalue"/>
					</xsl:if>
					<xsl:value-of select="string(' ')"/>
					<xsl:call-template name="pathvalue">
						<xsl:with-param name="value" select="substring-after($value,' ')"/>
						<xsl:with-param name="pagex" select="$pagex"/>
						<xsl:with-param name="pagey" select="$pagey"/>
						<xsl:with-param name="count" select="number($count)+1"/>
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
	</xsl:template>
</xsl:stylesheet>
