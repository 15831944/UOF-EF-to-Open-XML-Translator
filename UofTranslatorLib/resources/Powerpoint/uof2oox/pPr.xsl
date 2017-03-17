<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:pzip="urn:u2o:xmlns:post-processings:special"
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
xmlns:式样="http://schemas.uof.org/cn/2009/styles"
xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks">
	<xsl:import href="hyperlink.xsl"/>
	<xsl:import href="numbering.xsl"/>
	<xsl:template name="pPr">
		<!-- 10.27 黎美秀修改    
        <xsl:for-each select ="字:大纲级别">
      <xsl:attribute name ="lvl">
        <xsl:value-of select ="."/>
      </xsl:attribute>
    </xsl:for-each>
    -->
		<!--修改标签 李娟 2012.01.07-->
		<xsl:variable name="styleref">
			<xsl:value-of select="@式样引用_419C"/>
		</xsl:variable>
		<xsl:variable name="lvlno">
			<xsl:choose>
				<xsl:when test="字:大纲级别_417C">
				
					<xsl:value-of select="字:大纲级别_417C"/>
				</xsl:when>
				<xsl:when test="//式样:段落式样_9905[@标识符_4100=$styleref]">
					
					<xsl:value-of select="//式样:段落式样_9905[@标识符_4100=$styleref]/字:大纲级别_417C"/>
				</xsl:when>
				<xsl:otherwise>
					<!-- 11.03 黎美秀修改 有可能在UOF：段落式样中-->
					<xsl:value-of select="//式样:段落式样_9912[@标识符_4100=$styleref]/字:大纲级别_417C"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		
		<!-- 09.11.11 马有旭 添加-->
		<!--2011.2.21 罗文甜 修改大纲级别-->
		<xsl:if test="$lvlno!=''">
			<xsl:attribute name="lvl">
				<xsl:choose>
          <!--liuyangyagn 2015-03-30 修改大纲级别错误 start-->
          <xsl:when test="name(.)='字:段落属性_419B'">
            <xsl:choose>
              <xsl:when test="$lvlno='0'">
                <xsl:value-of select="'0'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$lvlno - 1"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <!--end liuyangyagn 2015-03-30 修改大纲级别错误-->
					<!--xsl:when test="ancestor::uof:锚点/@uof:占位符='title' or ancestor::uof:锚点/@uof:占位符='centertitle' or ancestor::uof:锚点/@uof:占位符='vertical_title'
                    or ancestor::uof:锚点/@uof:占位符='date' or ancestor::uof:锚点/@uof:占位符='footer' or ancestor::uof:锚点/@uof:占位符='number' or ancestor::uof:锚点/@uof:占位符='header'"-->
					<xsl:when test="ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627='notes' or ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627='title' or ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627='centertitle' or ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627='vertical_title' or ancestor::uof:锚点_C644/uof:占位符_C626/@类型_C627='text'">
						<xsl:value-of select="$lvlno"/>
					</xsl:when>
					<xsl:otherwise>
						<!--此块有问题，下期修改-->
						<xsl:choose>
							<xsl:when test="$lvlno='0'">
								<xsl:value-of select="'0'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="$lvlno - 1"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
		</xsl:if>
		<!--xsl:choose>
      <xsl:when test="not(字:对齐)">
        <xsl:attribute name="algn">ctr</xsl:attribute>
      </xsl:when>
      <xsl:otherwise-->
		<!-- 李娟 修改文字对齐 2012.04.24-->
		<xsl:choose>
			<xsl:when test="字:对齐_417D">
				<xsl:for-each select="字:对齐_417D">
			<xsl:if test="@水平对齐_421D">
				<xsl:attribute name="algn">
					<xsl:choose>
						<xsl:when test="@水平对齐_421D='left'">l</xsl:when>
						<xsl:when test="@水平对齐_421D='right'">r</xsl:when>
						<xsl:when test="@水平对齐_421D='center'">ctr</xsl:when>
						<xsl:when test="@水平对齐_421D='justified'">just</xsl:when>
						<xsl:when test="@水平对齐_421D='distributed'">dist</xsl:when>
					</xsl:choose>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test="@文字对齐_421E">
				<xsl:attribute name="fontAlgn">
					<xsl:choose>
						<xsl:when test="@文字对齐_421E='auto'">auto</xsl:when>
						<xsl:when test="@文字对齐_421E='top'">t</xsl:when>
						<xsl:when test="@文字对齐_421E='center'">ctr</xsl:when>
						<xsl:when test="@文字对齐_421E='base'">base</xsl:when>
						<xsl:when test="@文字对齐_421E='bottom'">b</xsl:when>
					</xsl:choose>
				</xsl:attribute>
			</xsl:if>
		</xsl:for-each>
			</xsl:when>
			<xsl:when test="not(字:对齐_417D)">
				<xsl:variable name="fontAlgn" select="//式样:段落式样_9905[@标识符_4100=$styleref]/字:对齐_417D"/>
					<xsl:choose>
						<xsl:when test="$fontAlgn/@水平对齐_421D">
							<xsl:attribute name="algn">
								<xsl:choose>
									<xsl:when test="$fontAlgn/@水平对齐_421D='left'">l</xsl:when>
									<xsl:when test="$fontAlgn/@水平对齐_421D='right'">r</xsl:when>
									<xsl:when test="$fontAlgn/@水平对齐_421D='center'">ctr</xsl:when>
									<xsl:when test="$fontAlgn/@水平对齐_421D='justified'">just</xsl:when>
									<xsl:when test="$fontAlgn/@水平对齐_421D='distributed'">dist</xsl:when>
								</xsl:choose>
							</xsl:attribute>
						</xsl:when>
						<xsl:when test="$fontAlgn/文字对齐_421E">
							<xsl:attribute name="fontAlgn">
								<xsl:choose>
									<xsl:when test="$fontAlgn/文字对齐_421E='auto'">auto</xsl:when>
									<xsl:when test="$fontAlgn/文字对齐_421E='top'">t</xsl:when>
									<xsl:when test="$fontAlgn/文字对齐_421E='center'">ctr</xsl:when>
									<xsl:when test="$fontAlgn/文字对齐_421E='base'">base</xsl:when>
									<xsl:when test="$fontAlgn/文字对齐_421E='bottom'">b</xsl:when>
								</xsl:choose>
							</xsl:attribute>
						</xsl:when>
					</xsl:choose>
				
			</xsl:when>
		</xsl:choose>

    <!-- 2014-05-06, tangjiang, 添加默认的缩进方式，修复互操作缩进方式错误 start -->
    <!--2015-03-26,注释代码修复多级编号缩进错误的bug start-->
    <!--<xsl:if test="not(字:缩进_411D)">
      <xsl:attribute name="indent">
        <xsl:value-of select="'0'"/>
      </xsl:attribute>
      <xsl:attribute name="marL">
        <xsl:value-of select="'0'"/>
      </xsl:attribute>
      <xsl:attribute name="hangingPunct">
        <xsl:value-of select="'0'"/>
      </xsl:attribute>
    </xsl:if>-->
    <!-- end 2015-03-26,注释代码修复多级编号缩进错误的bug-->
    <!-- end 2014-05-06, tangjiang, 添加默认的缩进方式，修复互操作缩进方式错误 -->
    
    <xsl:for-each select="字:缩进_411D">
			<xsl:for-each select="字:首行_4111/字:绝对_4107">
				<xsl:attribute name="indent">
					<xsl:value-of select="round(number(@值_410F) * 12700)"/>
				</xsl:attribute>
			</xsl:for-each>
			<!--金山中luowntian 2011-5-26-->
			<xsl:for-each select="字:左_410E/字:绝对_4107">
				<xsl:attribute name="marL">
					<xsl:choose>
						<xsl:when test="@值_410F &gt;0">
							<xsl:value-of select="round(number(@值_410F) * 12700)"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'0'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
			</xsl:for-each>
			<!--2011-3-7罗文甜，修改缩进bug-永中-->
      <!--
			<xsl:for-each select="字:左/字:绝对">
				<xsl:attribute name="marL">
          <xsl:choose>
            <xsl:when test="../../字:首行/字:绝对">
              <xsl:choose>
                <xsl:when test="../../字:首行/字:绝对/@字:值 &gt; 0">
                  <xsl:value-of select="round((number(@字:值) + number(../../字:首行/字:绝对/@字:值)) * 12700)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="round((number(@字:值) - number(../../字:首行/字:绝对/@字:值)) * 12700)"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="round(number(@字:值) * 12700)"/>
            </xsl:otherwise>
          </xsl:choose>
				</xsl:attribute>
			</xsl:for-each>
      -->
      
			<xsl:for-each select="字:右_4110/字:绝对_4107">
				<xsl:attribute name="marR">
					<xsl:value-of select="round(number(@值_410F) * 12700)"/>
				</xsl:attribute>
			</xsl:for-each>
      
		</xsl:for-each>
		<!--中文习惯首尾字符 默认为“1”-->
		<xsl:if test="字:是否采用中文习惯首尾字符_4197='false' or 字:是否采用中文习惯首尾字符_4197='0' or 字:是否采用中文习惯首尾字符_4197='off'">
			<xsl:attribute name="eaLnBrk">0</xsl:attribute>
		</xsl:if>
		
		<xsl:if test="字:是否允许单词断字_4194='true' or 字:是否允许单词断字_4194='1'or 字:是否允许单词断字_4194='on'">
			<xsl:attribute name="latinLnBrk">1</xsl:attribute>
		</xsl:if>
		<!--行首尾标点控制，永中默认为0，office2007为1-->

		<xsl:attribute name="hangingPunct">
			<xsl:choose>
				<xsl:when test="字:是否行首尾标点控制_4195='true' or 字:是否行首尾标点控制_4195='1' or 字:是否行首尾标点控制_4195='on'">1</xsl:when>
				<!--lwt-->
				<xsl:when test="not(字:是否行首尾标点控制_4195)">0</xsl:when>
				<xsl:otherwise>0</xsl:otherwise>
			</xsl:choose>
		</xsl:attribute>
    
		<!--xsl:value-of select ="number(.) * 12700"/-->
		
		<xsl:for-each select="字:行距_417E">
			
			<!--搞不懂多倍行距到底怎么算的。。。
      This type specifies the range that the spacing percent, in terms of a line. Represented in the file format from 0-
1000 line hundredths, maps to 0.00-10.00 lines.This simple type has a minimum value of greater than or equal to 0.
 This simple type has a maximum value of less than or equal to 13200000.  <a:spcPct val="200000"/> - - -200% of the size of the largest text-->
			<xsl:choose>
				<xsl:when test="@类型_417F='multi-lines'">
					
					<a:lnSpc>
						<a:spcPct>
							<xsl:attribute name="val">
								<xsl:value-of select="ceiling((@值_4108) * 100000)"/>
							</xsl:attribute>
						</a:spcPct>
					</a:lnSpc>
				</xsl:when>
				<xsl:when test="@类型_417F='fixed'">
					<a:lnSpc>
						<a:spcPts>
							<xsl:attribute name="val">
								<xsl:value-of select="((@值_4108) * 100)"/>
							</xsl:attribute>
						</a:spcPts>
					</a:lnSpc>
				</xsl:when>
				<!--最小值、行间距 无对应-->
				<!--最小值不转的话出问题，默认转为固定值-->
				<xsl:when test="@类型_417F='at-least'">
					<a:lnSpc>
						<a:spcPts>
							<xsl:attribute name="val">
								<xsl:value-of select="((@值_4108) * 100)"/>
							</xsl:attribute>
						</a:spcPts>
					</a:lnSpc>
				</xsl:when>
			</xsl:choose>
		</xsl:for-each>
		<xsl:for-each select="字:段间距_4180">
			<xsl:for-each select="字:段前距_4181">
				<a:spcBef>
					<xsl:if test="name(*)='字:相对值_4184'">
						<a:spcPct>
							<!--10.22 黎美秀修改 增加round 两处
              <xsl:attribute name ="val">
                <xsl:value-of select ="number(字:相对值/@字:值) * 100000"/>
              </xsl:attribute>
              -->
							<xsl:attribute name="val">
								<xsl:value-of select="round(number(字:相对值_4184) * 100000)"/>
							</xsl:attribute>
						</a:spcPct>
					</xsl:if>
					<xsl:if test="name(*)='字:绝对值_4183'">
						<a:spcPts>
							<xsl:attribute name="val">
								<xsl:value-of select="round(number(字:绝对值_4183) * 100)"/>
							</xsl:attribute>
						</a:spcPts>
					</xsl:if>
					<!--罗文甜 自动 无对应-->
					<xsl:if test="name(*)='字:自动_4182'">
						<a:spcPts val="0"/>
					</xsl:if>

				</a:spcBef>
			</xsl:for-each>
			<xsl:for-each select="字:段后距_4185">
				<a:spcAft>
					<xsl:if test="name(*)='字:相对值_4184'">
						<a:spcPct>
							<xsl:attribute name="val">
								<xsl:value-of select="number(字:相对值_4184) * 100000"/>
							</xsl:attribute>
						</a:spcPct>
					</xsl:if>
					<xsl:if test="name(*)='字:绝对值_4183'">
						<a:spcPts>
							<xsl:attribute name="val">
								<xsl:value-of select="number(字:绝对值_4183) * 100"/>
							</xsl:attribute>
						</a:spcPts>
					</xsl:if>
					<!--2011-4-20 罗文甜 自动 无对应-->
					<xsl:if test="name(*)='字:自动_4182'">
						<a:spcPts val="0"/>
					</xsl:if>
				</a:spcAft>
			</xsl:for-each>
		</xsl:for-each>
    
    <!--start liuyin 20130510 修改2915 S1的副标题-->
    <xsl:variable name="tempsz">
      <xsl:value-of select="../../图:文本_803C/图:内容_8043/字:段落_416B/字:句_419D/字:句属性_4158/字:字体_4128/@字号_412D"/>
    </xsl:variable>
    <!--end liuyin 20130510 修改2915 S1的副标题-->
    <xsl:variable name="tempcol">
      <xsl:value-of select="../../图:文本_803C/图:内容_8043/字:段落_416B/字:句_419D/字:句属性_4158/字:字体_4128/@颜色_412F"/>
    </xsl:variable>
    
    
    
		<!--    2010.2.24 黎美秀修改 
     <xsl:for-each select="字:自动编号信息[@字:编号级别!=0]">
     改为
      <xsl:for-each select="字:自动编号信息">
      在numbering中处理@字:编号级别!=0 即bunone的情况    
    -->
    
    <!--start liuyin 20130327 修改bug2785 项目符号不正确-->
    <!--<xsl:for-each select="字:自动编号信息_4186">
      --><!--*<xsl:value-of select="name(.)"/>*--><!--
      <xsl:variable name="numID" select="@编号引用_4187"/>
      --><!--2010.04.10 myx add --><!--
      <xsl:variable name="level" select="@编号级别_4188"/>
      <xsl:for-each select="//uof:UOF/uof:式样集/式样:式样集_990B/式样:自动编号集_990E/字:自动编号_4124[@标识符_4100=$numID]">
        --><!--2010.04.10 myx<xsl:call-template name="numbering"/>--><!--
        <xsl:call-template name="numbering">
          <xsl:with-param name="numberLevel" select="$level"/>
        </xsl:call-template>
      </xsl:for-each>
    </xsl:for-each>-->

    
    <!-- 2014-04-20, tangjiang, 修复互操作多出项目符号错误 start -->
    <xsl:choose>
      <xsl:when test="字:自动编号信息_4186[@编号级别_4188='0']">
        <xsl:element name="a:buNone"></xsl:element>
      </xsl:when>
      <xsl:when test="not(字:自动编号信息_4186/@编号级别_4188) and not(../../../../演:lstStyle/式样:段落式样_9905/字:自动编号信息_4186/@编号级别_4188)">
        <xsl:element name="a:buNone"></xsl:element>
      </xsl:when>
      <xsl:when test="字:自动编号信息_4186[@编号级别_4188!='0']">
        <xsl:for-each select="字:自动编号信息_4186">
          <!--*<xsl:value-of select="name(.)"/>*-->
          <xsl:variable name="numID" select="@编号引用_4187"/>
          <!--2010.04.10 myx add -->
          <xsl:variable name="level" select="@编号级别_4188"/>
          <xsl:for-each select="//uof:UOF/uof:式样集/式样:式样集_990B/式样:自动编号集_990E/字:自动编号_4124[@标识符_4100=$numID]">
            <!--2010.04.10 myx<xsl:call-template name="numbering"/>-->
            <xsl:call-template name="numbering">
              <xsl:with-param name="numberLevel" select="$level"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="../../../../演:lstStyle/式样:段落式样_9905/字:自动编号信息_4186[@编号级别_4188!='0']">
        <!--start liuyin 20130417 修改bug2837 uof->ooxml回归测试的功能测试中，字体集功能点转换后的文档需要修复-->
        <xsl:variable name="relstyle" select="@式样引用_419C"/>
          <xsl:for-each select="//式样:文本式样集_9913/式样:文本式样_9914/式样:段落式样_9905[@标识符_4100=$relstyle]/字:自动编号信息_4186">
         <!--end liuyin 20130417 修改bug2837 uof->ooxml回归测试的功能测试中，字体集功能点转换后的文档需要修复-->
            <xsl:variable name="numID" select="@编号引用_4187"/>
              <xsl:variable name="level" select="@编号级别_4188"/>
              <xsl:for-each select="//uof:UOF/uof:式样集/式样:式样集_990B/式样:自动编号集_990E/字:自动编号_4124[@标识符_4100=$numID]">
                <xsl:call-template name="numbering">
                  <xsl:with-param name="numberLevel" select="$level"/>
                </xsl:call-template>
              </xsl:for-each>
            <!--</xsl:if>-->
          </xsl:for-each>
        <!--</xsl:if>-->
      </xsl:when>
      <xsl:otherwise>
        <xsl:element name="a:buNone"></xsl:element>
      </xsl:otherwise>
    </xsl:choose>
    <!-- 2014-04-20, tangjiang, 修复互操作多出项目符号错误 -->
    
    <!--end liuyin 20130327 修改bug2785 项目符号不正确-->
		<!--
    2010.4.7 黎美秀增加 增加制表符的转换    
    -->
		<xsl:if test="字:制表位设置_418F">
			<a:tabLst>
				<xsl:for-each select="字:制表位设置_418F/字:制表位_4171">
					<a:tab>
						<xsl:attribute name="pos">
              <!--liuyangyang 2015-03-09 pos数值为整形-->
							<!--<xsl:value-of select="@位置_4172 * 12698.4"/>-->
              <xsl:value-of select="round(@位置_4172 * 12698.4)"/>
              <!--end-->
						</xsl:attribute>
						<xsl:attribute name="algn">
							<xsl:choose>
								<xsl:when test="@类型_4173='left'">
									<xsl:value-of select="'l'"/>
								</xsl:when>
								<xsl:when test="@类型_4173='right'">
									<xsl:value-of select="'r'"/>
								</xsl:when>
								<xsl:when test="@类型_4173='decimal'">
									<xsl:value-of select="'dec'"/>
								</xsl:when>
								<xsl:when test="@类型_4173='center'">
									<xsl:value-of select="'ctr'"/>
								</xsl:when>
							</xsl:choose>
						</xsl:attribute>
					</a:tab>
				</xsl:for-each>
			</a:tabLst>
		</xsl:if>
    <!--start liuyin 20130510 修改2915 S1的副标题-->
    <xsl:apply-templates select="字:句属性_4158" mode="pPr">
      <xsl:with-param name="tempsz" select="$tempsz"/>
      <xsl:with-param name="tempcol" select="$tempcol"/>
    </xsl:apply-templates>
	</xsl:template>
	<xsl:template match="字:句属性_4158" mode="pPr">
    <xsl:param name="tempsz"/>
    <xsl:param name="tempcol"/>
		<xsl:element name="a:defRPr">
      <xsl:call-template name="rPr">
        <xsl:with-param name="tempsz" select="$tempsz"/>
        <xsl:with-param name="tempcol" select="$tempcol"/>
      </xsl:call-template>
		</xsl:element>
	</xsl:template>
	<xsl:template name="rPr">
    <xsl:param name="tempsz"/>
    <xsl:param name="tempcol"/>
    <!--end liuyin 20130510 修改2915 S1的副标题-->
		<xsl:attribute name="lang">
			<xsl:choose>
				<xsl:when test =".//字:字体_4128/@是否西文绘制_412C='true'">
					<xsl:value-of select ="'en-US'"/>
				</xsl:when>
				<xsl:when test =".//字:字体_4128/@西文字体引用_4129 and not(字:字体_4128/@中文字体引用_412A)">
					<xsl:value-of select ="'en-US'"/>
				</xsl:when>
				<xsl:otherwise >
					<xsl:value-of select ="'zh-CN'"/>
				</xsl:otherwise>
			</xsl:choose>

		</xsl:attribute>
		<xsl:attribute name="altLang">
			<xsl:choose>
				<xsl:when test ="字:字体_4128/@是否西文绘制_412C='true'">
					<xsl:value-of select ="'zh-CN'"/>
				</xsl:when>
				<xsl:when test ="字:字体_4128/@西文字体引用_4129 and not(字:字体_4128/@中文字体引用_412A)">
					<xsl:value-of select ="'zh-CN'"/>
				</xsl:when>
				<xsl:otherwise >
					<xsl:value-of select ="'en-US'"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:attribute>
    
    <!--start liuyin 20130103 修改字体集中的字体大小，转换错误-->
    <!--start liuyin 20130112 修改字体集中的字体大小，转换错误-->
    <!-- 字号 -->
    <!-- 2014-03-24, tangjiang, 修复字号转换 start -->
    <xsl:choose>
      <xsl:when test="./字:字体_4128/@字号_412D">
        <xsl:attribute name="sz">
          <xsl:value-of select="number(./字:字体_4128/@字号_412D) * 100"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="not(字:字体_4128/@字号_412D)">
        <xsl:variable name="sz_rel">
          <xsl:value-of select="ancestor::字:段落_416B/字:段落属性_419B/@式样引用_419C"/>
        </xsl:variable>
        <xsl:attribute name="sz">
          <xsl:choose>
            <xsl:when test="//式样:段落式样_9912[@标识符_4100=$sz_rel]/字:句属性_4158/字:字体_4128/@字号_412D">
              <xsl:value-of select="round(number(//式样:段落式样_9912[@标识符_4100=$sz_rel]/字:句属性_4158/字:字体_4128/@字号_412D) * 100)"/>
            </xsl:when>

            <!--start liuyin 20130112 修改bug2720的字体大小-->
            <xsl:when test="//式样:段落式样_9905[@标识符_4100=$sz_rel]/字:句属性_4158/字:字体_4128/@字号_412D">
              <xsl:value-of select="round(number(//式样:段落式样_9905[@标识符_4100=$sz_rel]/字:句属性_4158/字:字体_4128/@字号_412D) * 100)"/>
            </xsl:when>
            <!--end liuyin 20130112 修改bug2720的字体大小-->

            <xsl:otherwise>
              <xsl:value-of select="3200"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
    <!-- end 2014-03-24, tangjiang, 修复字号转换 -->
    <!--end liuyin 20130112 修改字体集中的字体大小，转换错误-->
    <!--end liuyin 20130103 修改字体集中的字体大小，转换错误-->

    <!--start liuyin 20130510 修改2915 S1的副标题，微软雅黑20变为18-->
    <xsl:choose>
      <xsl:when test="$tempsz !=''">
       <xsl:attribute name="sz">
        <xsl:value-of select="round($tempsz * 100)"/>
      </xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
     <xsl:for-each select="字:字体_4128/@字号_412D">
        <xsl:variable name="szVal" select="current()"/>
       <xsl:if test=" $szVal != 'NaN' ">
         <xsl:attribute name="sz">
           <xsl:value-of select="round(number($szVal) *100)"/>
         </xsl:attribute>
       </xsl:if>
     </xsl:for-each>
      </xsl:otherwise>
    </xsl:choose>
    <!--end liuyin 20130510 修改2915 S1的副标题，微软雅黑20变为18-->

    <!-- 2014-04-16, tangjiang, 修复粗体、斜体转换 start -->
    <xsl:variable name="styleRefId">
      <xsl:value-of select="ancestor::字:段落_416B/字:段落属性_419B/@式样引用_419C"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="字:是否粗体_4130='off' or 字:是否粗体_4130='false' or 字:是否粗体_4130='0'">
        <xsl:attribute name="b">0</xsl:attribute>
      </xsl:when>
      <xsl:when test="字:是否粗体_4130='on' or 字:是否粗体_4130='true' or 字:是否粗体_4130='1'">
        <xsl:attribute name="b">1</xsl:attribute>
      </xsl:when>
      <xsl:when test="$styleRefId !='' and  //式样:段落式样_9905[@标识符_4100=$styleRefId]/字:句属性_4158/字:是否粗体_4130 = 'true'">
        <xsl:attribute name="b">1</xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="b">0</xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:choose>
      <xsl:when test="字:是否斜体_4131='on' or 字:是否斜体_4131='true' or 字:是否斜体_4131='1'">
        <xsl:attribute name="i">1</xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="i">0</xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
    <!-- end 2014-04-16, tangjiang, 修复粗体、斜体转换 -->
    
		<!--2011-3-30罗文甜，修改下划线类型-->
		<xsl:for-each select="字:下划线_4136">
			<xsl:attribute name="u">
				<xsl:choose>
					<xsl:when test="not(@字:下划线_4136)">
						<xsl:choose>
              
              <!--start liuyin 20130107 修改句式样中的下划线类型转换错误。-->
              <xsl:when test="(@线型_4137='single' and not(@虚实_4138) and @是否字下划线_4139='true') or (@线型_4137='single' and @虚实_4138='solid' and @是否字下划线_4139='true')">words</xsl:when>
              <xsl:when test="@线型_4137='thick' and not(@虚实_4138)">heavy</xsl:when>
              <xsl:when test="@线型_4137='thick'and @虚实_4138='square-dot'">dottedHeavy</xsl:when>
              <xsl:when test="@线型_4137='single'and @虚实_4138='dash'">dash</xsl:when>
              <xsl:when test="@线型_4137='thick'and @虚实_4138='dash'">dashHeavy</xsl:when>
              <xsl:when test="@线型_4137='thick'and @虚实_4138='long-dash'">dashLongHeavy</xsl:when>
              <xsl:when test="@线型_4137='thick'and @虚实_4138='dash-dot'">dotDashHeavy</xsl:when>
              <xsl:when test="@线型_4137='thick'and @虚实_4138='dash-dot-dot'">dotDotDashHeavy</xsl:when>
              <xsl:when test="@线型_4137='single'and @虚实_4138='wave'">wavy</xsl:when>
              <xsl:when test="@线型_4137='thick'and @虚实_4138='wave'">wavyHeavy</xsl:when>
              <xsl:when test="@线型_4137='double'and @虚实_4138='wave'">wavyDbl</xsl:when>
              <!--end liuyin 20130107 修改句式样中的下划线类型转换错误。-->

              <xsl:when test="(@线型_4137='single' and not(@虚实_4138)) or (@线型_4137='single' and @虚实_4138='solid')">sng</xsl:when>
							<xsl:when test="(@线型_4137='double' and not(@虚实_4138)) or (@线型_4137='double' and @虚实_4138='solid')">dbl</xsl:when>
							<xsl:when test="@线型_4137='single' and @虚实_4138='square-dot'">dotted</xsl:when>
							<xsl:when test="@线型_4137='single' and @虚实_4138='dash'">dash</xsl:when>
							<xsl:when test="@线型_4137='single' and @虚实_4138='long-dash'">dashLong</xsl:when>
							<xsl:when test="@线型_4137='single' and @虚实_4138='dash-dot'">dotDash</xsl:when>
							<xsl:when test="@线型_4137='single' and @虚实_4138='dash-dot-dot'">dotDotDash</xsl:when>
							<xsl:when test="@线型_4137='none'">none</xsl:when>
							<xsl:when test="@线型_4137='wave-heavy'">wavyHeavy</xsl:when>
							<xsl:when test="@线型_4137='wave-double'">wavyDbl</xsl:when>
							<xsl:when test="@线型_4137='wave'">wavy</xsl:when>
							<xsl:otherwise>
								<!--xsl:value-of select="current()"/-->
								<xsl:value-of select="'sng'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="'sng'"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
		</xsl:for-each>
		<!--删除线-->
		<xsl:for-each select="字:删除线_4135">
			<xsl:attribute name="strike">
				<xsl:choose>
					<xsl:when test=".='single'">sngStrike</xsl:when>
					<xsl:when test=".='double'">dblStrike</xsl:when>
					<xsl:when test=".='none'">noStrike</xsl:when>
					
						
				</xsl:choose>
				
			</xsl:attribute>
		</xsl:for-each>
		<xsl:attribute name="dirty">0</xsl:attribute>
		<xsl:attribute name="smtClean">0</xsl:attribute>
		<!--调整字间距-->
		<xsl:for-each select="字:调整字间距_4146">
			<xsl:attribute name="kern">
				<xsl:value-of select="number(.) *100"/>
			</xsl:attribute>
		</xsl:for-each>
    
		<!--醒目字体-->
    <!-- 2014-03-24, tangjiang, 修复醒目字体转换 start -->
    <xsl:choose>
      <xsl:when test="./字:醒目字体类型_4141">
        <xsl:for-each select="字:醒目字体类型_4141">
          <xsl:attribute name="cap">
            <xsl:choose>
              <xsl:when test=".='none'">none</xsl:when>
              <xsl:when test=".='small-caps'">small</xsl:when>
              <xsl:when test=".='uppercase'">all</xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="not(./字:醒目字体类型_4141)">
        <xsl:variable name="styleId" select="ancestor::字:段落_416B/字:段落属性_419B/@式样引用_419C"/>
        <xsl:choose>
          <xsl:when test="//式样:段落式样_9905[@标识符_4100=$styleId]">
            <xsl:for-each select="//式样:段落式样_9905[@标识符_4100=$styleId]/字:句属性_4158/字:醒目字体类型_4141">
              <xsl:attribute name="cap">
                <xsl:choose>
                  <xsl:when test=".='none'">none</xsl:when>
                  <xsl:when test=".='small-caps'">small</xsl:when>
                  <xsl:when test=".='uppercase'">all</xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:for-each>
          </xsl:when>
          <xsl:when test="//式样:段落式样_9912[@标识符_4100=$styleId]">
            <xsl:for-each select="//式样:段落式样_9912[@标识符_4100=$styleId]/字:句属性_4158/字:醒目字体类型_4141">
              <xsl:attribute name="cap">
                <xsl:choose>
                  <xsl:when test=".='none'">none</xsl:when>
                  <xsl:when test=".='small-caps'">small</xsl:when>
                  <xsl:when test=".='uppercase'">all</xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:for-each>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
    <!-- end 2014-03-24, tangjiang, 修复醒目字体转换 -->
    
		<!--字符间距-->
		<xsl:for-each select="字:字符间距_4145">
			<xsl:attribute name="spc">
        
        <!--start liuyin 20130316 修改bug2771需要修复才能打开-->
        <!--<xsl:value-of select="number(.) *100"/>-->
        
        <!--start liuyin 20130316 修改bug2774需要修复才能打开-->
        <!--<xsl:value-of select="round(number(.) *100)"/>-->
				<xsl:value-of select="round(number(.) *95)"/>
        <!--start liuyin 20130316 修改bug2774需要修复才能打开-->
        
        <!--end liuyin 20130316 修改bug2771需要修复才能打开-->
        
			</xsl:attribute>
		</xsl:for-each>
		<!--上下标-->
		<xsl:for-each select="字:上下标类型_4143">
			<xsl:attribute name="baseline">
				<xsl:choose>
					<xsl:when test=".='sup'">30000</xsl:when>
					<xsl:when test=".='none'">0</xsl:when>
					<xsl:when test=".='sub'">-25000</xsl:when>
				</xsl:choose>
			</xsl:attribute>
		</xsl:for-each>

    <!--start liuyin 20130510 修改2915 S1的副标题，字体颜色RGB(176,31,15)变为黑色-->
    <xsl:choose>
      <!--liuyangyang 2015-03-30 添加句属性字体空心转换 start-->
     <xsl:when test="./字:是否空心_413E = 'true'">
          <xsl:variable name="lnCol">
            <xsl:choose>
              <xsl:when test="$tempcol != ''">
                <xsl:value-of select="substring-after($tempcol,'#')"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="substring-after(./字:字体_4128/@颜色_412F,'#')"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <a:ln>
            <a:solidFill>
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="$lnCol"/>
                </xsl:attribute>
              </a:srgbClr>
            </a:solidFill>
          </a:ln>
          <a:solidFill>
            <a:srgbClr val="ffffff"/>
          </a:solidFill>

      </xsl:when>
      <xsl:when test="./字:字体_4128/@颜色_412F">
        <a:solidFill>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring-after(./字:字体_4128/@颜色_412F,'#')"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:solidFill>
      </xsl:when>
      <xsl:when test="$tempcol='' and not(./字:字体_4128/@颜色_412F)">
        <a:solidFill>
          <a:srgbClr val="000000"/>
        </a:solidFill>
      </xsl:when>
     <!-- <xsl:when test="$tempcol !=''">
        <a:solidFill>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring-after( $tempcol,'#')"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:solidFill>
      </xsl:when>
      <xsl:otherwise>
        <xsl:for-each select="字:字体_4128">
          <xsl:if test="@颜色_412F">
            <a:solidFill>
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-after(@颜色_412F,'#')"/>
                </xsl:attribute>
              </a:srgbClr>
            </a:solidFill>
          </xsl:if>
        </xsl:for-each>
      </xsl:otherwise>-->
      <!--end liuyangyang 2015-03-30 添加句属性字体空心转换-->
      
    </xsl:choose>
    <!--end liuyin 20130510 修改2915 S1的副标题，字体颜色RGB(176,31,15)变为黑色-->
    
		<!--start guoyongbing 2015.4.7 字体阴影转换失败-->
		<!--
    <xsl:if test="字:是否阴影_4140='true' or 字:是否阴影_4140='on' or 字:是否阴影_4140='1'">
			<a:effectLst>
				<a:prstShdw prst="shdw2">
					<a:srgbClr val="000000"/>
				</a:prstShdw>
			</a:effectLst>
		</xsl:if>
    -->
    <xsl:if test="字:是否阴影_4140='true' or 字:是否阴影_4140='on' or 字:是否阴影_4140='1'">
      <a:effectLst>
        <a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl">
          <a:srgbClr val="000000"/>
          <a:alpha val="43137"/>
        </a:outerShdw>
      </a:effectLst>
    </xsl:if>
    <!--end guoyongbing 2015.4.7 字体阴影转换失败-->
	
		<xsl:for-each select="字:突出显示颜色_4132">
			<a:highlight>
				<a:srgbClr>
					<xsl:attribute name="val">
						<xsl:value-of select="substring-after(@字:突出显示颜色_4132,'#')"/>
					</xsl:attribute>
				</a:srgbClr>
			</a:highlight>
		</xsl:for-each>
		<xsl:for-each select="字:下划线_4136">
			<xsl:if test="@颜色_412F">
				<a:uFill>
					<a:solidFill>
						<a:srgbClr>
							<xsl:attribute name="val">
								<xsl:value-of select="substring-after(@颜色_412F,'#')"/>
							</xsl:attribute>
						</a:srgbClr>
					</a:solidFill>
				</a:uFill>
			</xsl:if>
		</xsl:for-each>
		<!--2010-11-19罗文甜：修改字体（字体族）-->
		<xsl:for-each select="字:字体_4128">

      <!--start liuyin 20130103 修改字体集中的英文字体转换错误-->
      <!--start liuyin 20130326 修改2780 字体转换不正确-->
      <!--<xsl:if test="not(@西文字体引用_4129)">-->
      <xsl:if test="not(@西文字体引用_4129) and not(@是否西文绘制_412C)">
      <!--end liuyin 20130326 修改2780 字体转换不正确-->
        <a:latin>
          <xsl:attribute name="typeface">
            <xsl:value-of select="'Times New Roman'"/>
          </xsl:attribute>
         </a:latin>
      </xsl:if>
      <!--end liuyin 20130103 修改字体集中的英文字体转换错误-->
      
      <xsl:for-each select="@西文字体引用_4129">
				<a:latin>
					<!--2010.04.23<xsl:attribute name="typeface">
						<xsl:value-of select="//uof:字体集/uof:字体声明[@uof:标识符=current()]/@uof:字体族"/>
					</xsl:attribute>-->
					<xsl:attribute name="typeface">
						<xsl:choose>
							<xsl:when test="//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=current()]/@名称_9903">
								<xsl:value-of select="string(//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=current()]/@名称_9903)"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="current()"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</a:latin>
			</xsl:for-each>
			<xsl:for-each select="@中文字体引用_412A">
				<a:ea>
					<!--2010.04.23<xsl:attribute name="typeface">
						<xsl:value-of select="//uof:字体集/uof:字体声明[@uof:标识符=current()]/@ uof:字体族"/>
					</xsl:attribute>-->
					<xsl:attribute name="typeface">
						<xsl:choose>
							<xsl:when test="//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=current()]/@名称_9903">
								<xsl:value-of select="string(//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=current()]/@名称_9903)"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="current()"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</a:ea>
			</xsl:for-each>
			<xsl:for-each select="@特殊字体引用_412B">
				<a:cs>
					<!--2010.04.23<xsl:attribute name="typeface">
						<xsl:value-of select="//uof:字体集/uof:字体声明[@uof:标识符=current()]/@ uof:字体族"/>
					</xsl:attribute>-->
					<xsl:attribute name="typeface">
						<xsl:choose>
							<xsl:when test="//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=current()]/@名称_9903">
								<xsl:value-of select="//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=current()]/@名称_9903"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="current()"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</a:cs>
			</xsl:for-each>
		</xsl:for-each>
		<!--填充 - - - -北航做转换 -->
		<!--填充-->
		<!--下划线-->
		
	
			
			<xsl:for-each select="../字:区域开始_4165[@类型_413B='hyperlink']">
				
				<!--hyperlink的主要操作放在预处理中进行-->

				<xsl:for-each select="//超链:超级链接_AA0C[超链:链源_AA00=current()/@标识符_4100]">
					<xsl:call-template name="hyperlink"/>
				</xsl:for-each>
			</xsl:for-each>
		
	</xsl:template>
</xsl:stylesheet>
