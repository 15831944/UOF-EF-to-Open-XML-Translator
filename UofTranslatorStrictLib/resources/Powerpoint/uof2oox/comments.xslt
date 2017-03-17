<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
 xmlns:pzip="urn:u2o:xmlns:post-processings:special"
				xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
				
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
	<xsl:import href="table.xsl"/>
	<xsl:import href="txBody.xsl"/>
    <xsl:output method="xml" indent="yes"/>
<!-- 添加批注集 李娟 2012.02.24-->
	
    <xsl:template name="commentAuthors">
		<p:cmAuthorLst xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:p="http://purl.oclc.org/ooxml/presentationml/main">
			<!--<p:cmAuthor id="0" name="Love China" initials="LC" lastIdx="3" clrIdx="0"/>-->
			<!--<xsl:variable name="lastIdx">
				<xsl:number count="//规则:批注集_B669/规则:批注_B66A[@作者_41DD=ancestor::规则:公用处理规则_B665/规则:用户集_B667/规则:用户_B668/@标识符_4100]" format="1" level="single"/>
			</xsl:variable>-->
			<xsl:for-each select="uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:用户集_B667/规则:用户_B668">
			<p:cmAuthor lastIdx="1" clrIdx="1">
				
				<xsl:attribute name="id">
					<xsl:choose>
						<xsl:when test="contains(@标识符_4100,'auth')">
							<xsl:value-of select="substring-after(@标识符_4100,'auth')"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="@标识符_4100"/>
						</xsl:otherwise>
					</xsl:choose>
					
				</xsl:attribute>
				<xsl:attribute name="name">
					<xsl:value-of select="@姓名_41DC"/>
					<!--<xsl:variable name="author" select="@作者_41DD"/>
					<xsl:value-of select="preceding::规则:用户_B668[@标识符_4100=$author]/@姓名_41DC"/>-->
				</xsl:attribute>
				<xsl:attribute name="initials">
					<xsl:choose>
						<xsl:when test="//规则:批注集_B669/规则:批注_B66A/@作者缩写_41DF">
							<xsl:value-of select="//规则:批注集_B669/规则:批注_B66A/@作者缩写_41DF"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="@姓名_41DC"/>
						</xsl:otherwise>
					</xsl:choose>
				
				</xsl:attribute>
				<xsl:variable name="UserID">
					<xsl:value-of select="@标识符_4100"/>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="//规则:批注_B66A">
						<xsl:for-each select="//规则:批注_B66A">
							<xsl:variable name="lastIdx">
								<xsl:number count="//规则:批注_B66A[@作者_41DD=$UserID]" format="1" level="single"/>
							</xsl:variable>
							<xsl:attribute name="lastIdx">
                <xsl:choose>
                  <xsl:when test="$lastIdx!=''">
                    <xsl:value-of select="$lastIdx"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0'"/>
                  </xsl:otherwise>
                </xsl:choose>
							</xsl:attribute>
						</xsl:for-each>
					</xsl:when>
					<xsl:otherwise>
						
						<!--<xsl:variable name="lastName">
							<xsl:number count="@姓名_41DC" format="1" level="single"/>
						</xsl:variable>-->
						<xsl:attribute name="lastIdx">
							<xsl:value-of select="'1'"/>
							<!--<xsl:value-of select="$lastName"/>-->
						</xsl:attribute>
					</xsl:otherwise>
				</xsl:choose>
				<!--<xsl:for-each select="//规则:批注_B66A">
					<xsl:variable name="lastIdx">
						<xsl:number count="//规则:批注_B66A[@作者_41DD=$UserID]" format="1" level="single"/>
					</xsl:variable>
				<xsl:attribute name="lastIdx">
					<xsl:value-of select="$lastIdx"/>
					
				</xsl:attribute>
				</xsl:for-each>-->
				<xsl:attribute name="clrIdx">
					<xsl:choose>
						<xsl:when test="contains(@标识符_4100,'auth')">
							<xsl:value-of select="substring-after(@标识符_4100,'auth')"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="@标识符_4100"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
			</p:cmAuthor>
			</xsl:for-each>
		</p:cmAuthorLst>
		
	</xsl:template>

	<xsl:template name="comment">
		
		<xsl:for-each select="//演:幻灯片_6C0F">
			<xsl:if test="uof:锚点_C644/图:图形_8062/图:文本_803C/图:内容_8043/字:段落_416B/字:句_419D/字:区域开始_4165/@类型_413B='annotation'">
			<!--<xsl:for-each select="//演:幻灯片_6C0F[uof:锚点_C644/图:图形_8062/图:文本_803C/图:内容_8043/字:段落_416B/字:句_419D/字:区域开始_4165/@标识符_4100=$Field]">
			<xsl:if test="uof:锚点_C644/图:图形_8062/图:文本_803C/图:内容_8043/字:段落_416B/字:句_419D/字:区域开始_4165/@标识符_4100="-->
			<pzip:entry>
				<xsl:attribute name="pzip:target">
					<xsl:value-of select="concat('ppt/comments/comment',substring-after(@标识符_6B0A,'slideId'),'.xml')"/>
				</xsl:attribute>
				
				<p:cmLst xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:p="http://purl.oclc.org/ooxml/presentationml/main">
					<!--<xsl:call-template name="规则:批注_B66A"/>-->
					<!--<xsl:apply-templates select="//演:幻灯片_6C0F[uof:锚点_C644/图:图形_8062/图:文本_803C/图:内容_8043/字:段落_416B/字:句_419D/字:区域开始_4165/@类型_413B='annotation']"  mode="comment"/>-->
					<xsl:apply-templates select="." mode="comment"/>
				</p:cmLst>
			</pzip:entry>
				</xsl:if>
			
		</xsl:for-each>
</xsl:template>


	<xsl:template match="演:幻灯片_6C0F"  mode="comment">
		
		<!--<xsl:for-each select="规则:用户_B668">-->


    <xsl:for-each select="uof:锚点_C644[图:图形_8062/图:文本_803C/图:内容_8043/字:段落_416B/字:句_419D/字:区域开始_4165/@类型_413B='annotation']">
      <!--guoyongbin 2014-01-27 修改批注选取-->
      <xsl:for-each select="./图:图形_8062/图:文本_803C/图:内容_8043/字:段落_416B/字:句_419D[字:区域开始_4165/@类型_413B='annotation']/字:区域开始_4165">
        <xsl:variable name="region">
          <xsl:value-of select="./@标识符_4100"/>
        </xsl:variable>
        <xsl:for-each select="//规则:批注_B66A[@区域引用_41CE=$region]">
          <!--<xsl:for-each select="//规则:批注_B66A">-->
          <!--end-->
          <p:cm>

            <xsl:variable name="author">
              <xsl:value-of select="@作者_41DD"/>
            </xsl:variable>
            <xsl:attribute name="authorId">
              <xsl:choose>
                <xsl:when test="contains(@作者_41DD,'auth')">

                  <!--end liuyin 20130310 修改集成测试中，公用处理规则的用例转换后，文档需要修复-->
                  <!--<xsl:value-of select="substring-after(@标识符_4100,'auth')"/>-->
                  <xsl:value-of select="substring-after(@作者_41DD,'auth')"/>
                  <!--end liuyin 20130310 修改集成测试中，公用处理规则的用例转换后，文档需要修复-->

                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="@作者_41DD"/>
                </xsl:otherwise>
              </xsl:choose>
              <!--<xsl:value-of select="substring-after(@作者_41DD,'auth')"/>-->
            </xsl:attribute>
            <xsl:variable name="idx">
              <xsl:value-of select ="count(preceding-sibling::规则:批注_B66A[@作者_41DD=$author]) + 1"/>
            </xsl:variable>

            <xsl:attribute name="dt">
              <xsl:value-of select="@日期_41DE"/>
            </xsl:attribute>
            <!--<xsl:variable name ="author">
					<xsl:value-of select="规则:批注_B66A/@作者_41DD"/>
				</xsl:variable>-->
            <xsl:variable name="num" select="ancestor::uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:用户集_B667/规则:用户_B668/@标识符_4100"></xsl:variable>

            <xsl:attribute name="idx">
              <xsl:value-of select="$idx"/>
              <!--<xsl:number count="//规则:批注_B66A[@作者_41DD=//规则:用户_B668/@标识符_4100]" format="1" level="single"/>-->
            </xsl:attribute>


            <p:pos>
              <!--<xsl:variable name="posx" select="ancestor::uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108[ancestor::uof:锚点_C644//字:区域开始_4165/@标识符_4100=$region]"/>
					<xsl:variable name="posy" select="ancestor::uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/uof:锚点_C644/uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108[ancestor::uof:锚点_C644//字:区域开始_4165/@标识符_4100=$region]"/>-->
              <xsl:variable name="posx" select="ancestor::uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108"/>
              <xsl:variable name="posy" select="ancestor::uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/uof:锚点_C644/uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108"/>
              <xsl:attribute name="x">
                <xsl:value-of select="5222"/>
                <!--<xsl:value-of select="round($posx * 20)"/>-->
                <!--修改批注间隔 2013-03-20 liqiuling-->
              </xsl:attribute>
              <xsl:variable name="py" select="$idx*180"/>
              <xsl:attribute name="y">
                <xsl:value-of select="$py"/>
                <!--<xsl:value-of select="round($posy * 7)"/>-->
              </xsl:attribute>
            </p:pos>

            <!--start liuyin 20130418 修改2847，第一轮回归功能测试uof-oo部分，“批注”测试用例转换后需要修复才能打开oo文档-->
            <!--<xsl:if test=".//字:文本串_415B">-->
            <!--<xsl:if test="../字:文本串_415B">-->
            <p:text>
              <xsl:value-of select=".//字:文本串_415B"/>
            </p:text>
            <!--</xsl:if>-->
            <!--end liuyin 20130418 修改2847，第一轮回归功能测试uof-oo部分，“批注”测试用例转换后需要修复才能打开oo文档-->
          </p:cm>
        </xsl:for-each>
      </xsl:for-each>
    </xsl:for-each>
		
		
		
			<!--<xsl:for-each select="../../规则:批注集_B669/规则:批注_B66A[@作者_41DD=$ID]">
				<xsl:variable name="commentId" select="@区域引用_41CE"/>
				<xsl:if test="//演:幻灯片_6C0F/uof:锚点_C644/图:图形_8062/图:文本_803C/图:内容_8043/字:段落_416B/字:句_419D/字:区域开始_4165/@标识符_4100=$commentId">
				<p:cm>
					--><!--<xsl:for-each select="./规则:批注集_B669/规则:批注_B66A[@作者_41DD=$ID]">--><!--

						<xsl:attribute name="authorId">
							<xsl:value-of select="substring-after(@作者_41DD,'auth')"/>
						</xsl:attribute>
						<xsl:attribute name="dt">
							<xsl:value-of select="@日期_41DE"/>
						</xsl:attribute>
						<xsl:attribute name="idx">
							<xsl:number count="规则:批注_B66A[@作者_41DD=$ID]" format="1" level="single"/>
						</xsl:attribute>
					<p:pos>
						--><!--<xsl:for-each select=".">--><!--

						<xsl:variable name="posx" select="ancestor::uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108[ancestor::uof:锚点_C644//字:区域开始_4165/@标识符_4100=$commentId]"/>
						<xsl:variable name="posy" select="ancestor::uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/uof:锚点_C644/uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108[ancestor::uof:锚点_C644//字:区域开始_4165/@标识符_4100=$commentId]"/>
						<xsl:attribute name="x">
							<xsl:value-of select="round($posx * 20)"/>
						</xsl:attribute>
						<xsl:attribute name="y">
							<xsl:value-of select="round($posy * 7)"/>
						</xsl:attribute>
						--><!--</xsl:for-each>--><!--

					</p:pos>
						<xsl:if test=".//字:文本串_415B">
							<p:text>
								<xsl:value-of select=".//字:文本串_415B"/>
							</p:text>
						</xsl:if>
						
				</p:cm>
				</xsl:if>
		</xsl:for-each>-->	
		<!--</xsl:for-each>-->
	</xsl:template>
</xsl:stylesheet>
