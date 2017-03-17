<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:uof="http://schemas.uof.org/cn/2003/uof" xmlns:字="http://schemas.uof.org/cn/2009/uof-wordproc">
	<xsl:import href="revisions.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
  <!--yx,add,2010.5.12-->
  
	<!--
  <xsl:template match="/aa">
  <xsl:if test="w:document">
  <字:主体 uof:locID="t0016">
    <xsl:for-each select="w:document/w:body/node()">
      <xsl:if test="name()='w:p'">
        <xsl:call-template name="paragraph"/>
      </xsl:if>
    </xsl:for-each>
</字:主体>
  </xsl:if>
</xsl:template>-->
	<!--xsl:template name="paragraph">
<字:段落 uof:locID="t0051" uof:attrList="标识符">
<xsl:if test="w:pPr">
<字:段落属性 uof:locID="t0052" uof:attrList="式样引用">
<xsl:call-template name="pPr"/>
</字:段落属性>
</xsl:if>
<xsl:if test="w:r">
<字:句 uof:locID="t0085">
<xsl:call-template name="run"/>
</字:句>
</xsl:if>
</字:段落>
</xsl:template-->
	<xsl:template name="paragraph">
		<字:段落 uof:locID="t0051" uof:attrList="标识符">
			<xsl:for-each select="node()">
				<xsl:choose>
					<xsl:when test="name(.)='w:pPr'">
						<字:段落属性 uof:locID="t0052" uof:attrList="式样引用">
							<xsl:call-template name="pPr"/>
						</字:段落属性>
					</xsl:when>
					<xsl:when test="name(.)='w:r'">
						<字:句 uof:locID="t0085">
							<xsl:call-template name="run"/>
						</字:句>
					</xsl:when>
					<xsl:when test="name(.)='w:bookmarkStart'">
						<xsl:call-template name="bookmarkStart"/>
					</xsl:when>
					<xsl:when test="name(.)='w:bookmarkEnd'">
						<xsl:call-template name="bookmarkEnd"/>
					</xsl:when>
          <xsl:when test="name(.)='w:commentRangeStart'">
           
						<xsl:call-template name="commentStart"/>
					</xsl:when>
					<xsl:when test="name(.)='w:commentRangeEnd'">
           
						<xsl:call-template name="commentEnd"/>
					</xsl:when>
          <!--yx,add,2010.4.23-->
          <!--<xsl:when test="name(.)='w:hyperlink'">-->
          <xsl:when test="name(.)='w:hyperlink'">
            <xsl:call-template name="hyperlinkRegion"/>
          </xsl:when>
          <xsl:when test="name(.)='w:fldSimple'">
            <xsl:if test="./w:hyperlink">
              <xsl:for-each select="./w:hyperlink">
                <xsl:call-template name="hyperlinkRegion"/>
              </xsl:for-each>
            </xsl:if>
            <!--yx,add,2010.5.12-->
            <!--<xsl:if test="@w:instr">
              <xsl:variable name="temp" select="normalize-space(@w:instr)"/>
              <xsl:variable name="type" select="substring-before($temp,' ')"/>
              <xsl:choose>
                <xsl:when test="$type='AUTHOR' or $type='FILENAME' or $type='TIME' or $type='PAGE' or $type='SECTION' or $type='REF' or $type='XE' or $type='SEQ' or $type='TITLE'or $type='SAVEDATE'or $type='CREATEDATE'or $type='NUMPAGES' or $type='NUMCHARS'">
                  <xsl:call-template name="SimpleField">
                    <xsl:with-param name="type" select="translate($type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                    <xsl:with-param name="splfldPartFrom" select="$revPartFrom"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:when test="$type='REVNUM'">
                  <xsl:call-template name="SimpleField">
                    <xsl:with-param name="type" select="'revision'"/>
                    <xsl:with-param name="splfldPartFrom" select="$revPartFrom"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:when test="$type='SECTIONPAGES'">
                  <xsl:call-template name="SimpleField">
                    <xsl:with-param name="type" select="'pageinsection'"/>
                    <xsl:with-param name="splfldPartFrom" select="$revPartFrom"/>
                  </xsl:call-template>
                </xsl:when>
              </xsl:choose>
            </xsl:if>-->
          </xsl:when>
					<xsl:when test="name(.)='w:del'">
						<xsl:call-template name="rprchaP">
							<xsl:with-param name="lx" select="'delete'"/>
						</xsl:call-template>
					</xsl:when>
					<xsl:when test="name(.)='w:ins'">
						<xsl:call-template name="rprchaP">
							<xsl:with-param name="lx" select="'insert'"/>
						</xsl:call-template>
					</xsl:when>
				</xsl:choose>
			</xsl:for-each>
		</字:段落>
	</xsl:template>
	<xsl:template name="pPr">
		<xsl:choose>
			<xsl:when test="./w:pStyle">
				<xsl:attribute name="字:式样引用"><xsl:value-of select="./w:pStyle/@w:val"/></xsl:attribute>
			</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name="字:式样引用"><!--需要用到styles.xml解析正文式样id，引入styles.xml,存在路径问题-->
            <xsl:apply-templates select="document('word/styles.xml')" mode="style">
           </xsl:apply-templates></xsl:attribute>
			</xsl:otherwise>
		</xsl:choose>
		<xsl:if test="./w:spacing">
			<字:段间距 uof:locID="t0058">
				<!--判断属性是否存在-->
				<xsl:if test="./w:spacing/@w:before">
					<字:段前距 uof:locID="t0196">
						<字:绝对值 uof:locID="t0199" uof:attrList="值">
							<xsl:attribute name="字:值"><xsl:value-of select="format-number(./w:spacing/@w:before div 20,'0.0')"/></xsl:attribute>
						</字:绝对值>
					</字:段前距>
				</xsl:if>
				<xsl:if test="./w:spacing/@w:after">
					<字:段后距 uof:locID="t0197">
						<字:绝对值 uof:locID="t0202" uof:attrList="值">
							<xsl:attribute name="字:值"><xsl:value-of select="format-number(./w:spacing/@w:before div 20,'0.0')"/></xsl:attribute>
						</字:绝对值>
					</字:段后距>
				</xsl:if>
				<xsl:if test="./w:spacing/@w:beforeLines">
					<字:段前距 uof:locID="t0196">
						<字:相对值 uof:locID="t0200" uof:attrList="值">
							<xsl:attribute name="字:值"><xsl:value-of select="format-number(./w:spacing/@w:beforeLines div 100,'0.0')"/></xsl:attribute>
						</字:相对值>
					</字:段前距>
				</xsl:if>
				<xsl:if test="./w:spacing/@w:afterLines">
					<字:段后距 uof:locID="t0197">
						<字:相对值 uof:locID="t0200" uof:attrList="值">
							<xsl:attribute name="字:值"><xsl:value-of select="format-number(./w:spacing/@w:afterLines div 100,'0.0')"/></xsl:attribute>
						</字:相对值>
					</字:段后距>
				</xsl:if>
			</字:段间距>
		</xsl:if>
		<xsl:if test="./w:spacing/@w:line">
			<字:行距 uof:locID="t0057" uof:attrList="类型 值">
				<xsl:if test="./w:spacing/@w:lineRule='auto'">
					<xsl:attribute name="字:类型">multi-lines</xsl:attribute>
					<xsl:attribute name="字:值"><xsl:value-of select="format-number(./w:spacing/@w:line div 240,'0.0')"/></xsl:attribute>
				</xsl:if>
				<xsl:if test="./w:spacing/@w:lineRule='atLeast'">
					<xsl:attribute name="字:类型">at-least</xsl:attribute>
					<xsl:attribute name="字:值"><xsl:value-of select="format-number(./w:spacing/@w:line div 20,'0.0')"/></xsl:attribute>
				</xsl:if>
				<xsl:if test="./w:spacing/@w:lineRule='exact'">
					<xsl:attribute name="字:类型">fixed</xsl:attribute>
					<xsl:attribute name="字:值"><xsl:value-of select="format-number(./w:spacing/@w:line div 20,'0.0')"/></xsl:attribute>
				</xsl:if>
			</字:行距>
		</xsl:if>
		<xsl:if test="./w:keepLines">
			<xsl:choose>
				<xsl:when test="./w:keepLines/@w:val='false'">
					<字:段中不分页 uof:locID="t0062" uof:attrList="值" 字:值="false"/>
				</xsl:when>
				<xsl:when test="./w:keepLines/@w:val='0'">
					<字:段中不分页 uof:locID="t0062" uof:attrList="值" 字:值="false"/>
				</xsl:when>
				<xsl:when test="./w:keepLines/@w:val='off'">
					<字:段中不分页 uof:locID="t0062" uof:attrList="值" 字:值="false"/>
				</xsl:when>
				<xsl:otherwise>
					<字:段中不分页 uof:locID="t0062" uof:attrList="值" 字:值="true"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
		<xsl:if test="./w:keepNext">
			<xsl:choose>
				<xsl:when test="./w:keepNext/@w:val='false'">
					<字:与下段同页 uof:locID="t0063" uof:attrList="值" 字:值="false"/>
				</xsl:when>
				<xsl:when test="./w:keepNext/@w:val='0'">
					<字:与下段同页 uof:locID="t0063" uof:attrList="值" 字:值="false"/>
				</xsl:when>
				<xsl:when test="./w:keepNext/@w:val='off'">
					<字:与下段同页 uof:locID="t0063" uof:attrList="值" 字:值="false"/>
				</xsl:when>
				<xsl:otherwise>
					<字:与下段同页 uof:locID="t0063" uof:attrList="值" 字:值="true"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
		<xsl:if test="./w:pageBreakBefore">
			<xsl:choose>
				<xsl:when test="./w:pageBreakBefore/@w:val='false'">
					<字:段前分页 uof:locID="t0064" uof:attrList="值" 字:值="false"/>
				</xsl:when>
				<xsl:when test="./w:pageBreakBefore/@w:val='0'">
					<字:段前分页 uof:locID="t0064" uof:attrList="值" 字:值="false"/>
				</xsl:when>
				<xsl:when test="./w:pageBreakBefore/@w:val='off'">
					<字:段前分页 uof:locID="t0064" uof:attrList="值" 字:值="false"/>
				</xsl:when>
				<xsl:otherwise>
					<字:段前分页 uof:locID="t0064" uof:attrList="值" 字:值="true"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
		<xsl:if test="./w:jc">
			<字:对齐 uof:locID="t0055" uof:attrList="水平对齐 文字对齐">
				<xsl:choose>
					<xsl:when test="./w:jc/@ w:val='left'">
						<xsl:attribute name="字:水平对齐">left</xsl:attribute>
					</xsl:when>
					<xsl:when test="./w:jc/@ w:val='right'">
						<xsl:attribute name="字:水平对齐">right</xsl:attribute>
					</xsl:when>
					<xsl:when test="./w:jc/@ w:val='center'">
						<xsl:attribute name="字:水平对齐">center</xsl:attribute>
					</xsl:when>
					<xsl:when test="./w:jc/@w:val='both'">
						<xsl:attribute name="字:水平对齐">justified</xsl:attribute>
					</xsl:when>
				</xsl:choose>
			</字:对齐>
		</xsl:if>
		<xsl:if test="./w:widowControl">
			<字:孤行控制 uof:locID="t0060">
				<xsl:choose>
					<xsl:when test="./w:widowControl/@w:val='0'">0</xsl:when>
					<xsl:when test="./w:widowControl/@w:val='false'">0</xsl:when>
					<xsl:when test="./w:widowControl/@w:val='off'">0</xsl:when>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="./w:widowControl/@w:val='1'">1</xsl:when>
					<xsl:when test="./w:widowControl/@w:val='true'">1</xsl:when>
					<xsl:when test="./w:widowControl/@w:val='on'">1</xsl:when>
				</xsl:choose>
			</字:孤行控制>
		</xsl:if>
		<xsl:if test="./w:pBdr">
			<字:边框>
            </字:边框>
		</xsl:if>
	</xsl:template>
	<xsl:template name="run">
		<xsl:for-each select="node()">
			<xsl:choose>
				<xsl:when test="name(.)='w:rPr'">
					<xsl:call-template name="rPr"/>
				</xsl:when>
				<xsl:when test="name(.)='w:t'">
					<字:文本串 uof:locID="t0109" uof:attrList="字:标识符">
						<xsl:value-of select="."/>
					</字:文本串>
				</xsl:when>
			</xsl:choose>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="rPr">
		<字:句属性 uof:locID="t0086" uof:attrList="式样引用">
			<xsl:if test="w:rFonts|w:sz">
				<字:字体 uof:locID="t0088">
					<xsl:if test="w:sz">
						<xsl:attribute name="字:字号"><xsl:value-of select="w:sz/@w:val div 2"/></xsl:attribute>
					</xsl:if>
				</字:字体>
			</xsl:if>
			<xsl:if test="w:b">
				<字:粗体 uof:locID="t0089" uof:attrList="值" 字:值="true">
			</字:粗体>
			</xsl:if>
			<xsl:if test="w:i">
				<字:斜体 uof:locID="t0090" uof:attrList="值" 字:值="true">
			</字:斜体>
			</xsl:if>
			<xsl:if test="w:dstrike">
				<字:删除线>
			</字:删除线>
			</xsl:if>
			<xsl:if test="w:eboss">
				<字:浮雕>
			</字:浮雕>
			</xsl:if>
			<xsl:if test="w:em">
				<字:着重号>
			</字:着重号>
			</xsl:if>
			<xsl:if test="w:shadow">
				<字:阴影>
			</字:阴影>
			</xsl:if>
			<xsl:if test="w:outline">
				<字:空心>
			</字:空心>
			</xsl:if>
			<xsl:if test="w:rPrChange">
			    <xsl:call-template name="rprchange"/>
			</xsl:if>
		</字:句属性>
	</xsl:template>
	<xsl:template match="w:styles/w:style" mode="style">
		<xsl:if test="w:name/@w:val='Normal'">
			<xsl:value-of select="@w:styleId"/>
		</xsl:if>
	</xsl:template>
</xsl:stylesheet>
