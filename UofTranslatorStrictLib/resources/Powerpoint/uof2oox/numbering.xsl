<?xml version="1.0" encoding="UTF-8"?>
<!--
* Copyright (c) 2006, Beihang University, China
* All rights reserved.
*
* Redistribution and use in source and binary forms, with or without
* modification, are permitted provided that the following conditions are met:
*
*     * Redistributions of source code must retain the above copyright
*       notice, this list of conditions and the following disclaimer.
*     * Redistributions in binary form must reproduce the above copyright
*       notice, this list of conditions and the following disclaimer in the
*       documentation and/or other materials provided with the distribution.
*     * Neither the name of Clever Age, nor the names of its contributors may
*       be used to endorse or promote products derived from this software
*       without specific prior written permission.
*
* THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
* EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
* WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
* DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
* DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
* (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
* LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
* ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
* (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
* SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
-->
<xsl:stylesheet version="1.0" xmlns:pzip="urn:u2o:xmlns:post-processings:special" xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" 
				xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
				xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
				xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" 
				xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" 
				xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" 
				xmlns:p="http://purl.oclc.org/ooxml/presentationml/main" 
				xmlns="http://schemas.openxmlformats.org/package/2006/relationships" 
 xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph"
				xmlns:式样="http://schemas.uof.org/cn/2009/styles">
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
  <xsl:template name="numbering">
    <!--2010.04.10 myx add-->
	  <!--修改标签 李娟 2012.01.07-->
    <xsl:param name="numberLevel"/>
    <xsl:for-each select="字:级别_4112">
      <!--
      2010.2.24 黎美秀修改 处理标题与副标题转换后多一个项目符号的问题
        增加
       <xsl:choose>
        <xsl:when test="not(node())">
          <a:buNone/>
        </xsl:when>
        <xsl:otherwise>
        ...
        </xsl:otherwise>
      -->
      <xsl:choose>
        <!--2010.04.10 myx<xsl:when test="not(node())">-->
        <xsl:when test="not(node()) or $numberLevel='0'">
          <a:buNone/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="字:符号字体_4116/字:字体_4128">
            <!--a:buClr-->
			  
            <xsl:for-each select="@颜色_412F">
              <a:buClr>
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring-after(.,'#')"/>
                  </xsl:attribute>
                </a:srgbClr>
              </a:buClr>
            </xsl:for-each>
            <!--a:buSztPts-->
            <xsl:for-each select="@字号_412D">
              <a:buSzPts>
                <xsl:attribute name="val">
                  <xsl:value-of select=". * 100"/>
                </xsl:attribute>
              </a:buSzPts>
            </xsl:for-each>
            <!-- 修复bug：图形内项目编号转换不正确  liqiuling 2013-03-05 start-->
            <xsl:for-each select="@相对字号_412E">
              <a:buSzPct>
                <xsl:attribute name="val">
                  <xsl:value-of select=". * 1000"/>
                </xsl:attribute>
              </a:buSzPct>
            </xsl:for-each>
            <!-- 修复bug：图形内项目编号转换不正确  liqiuling 2013-03-05 end-->
            <xsl:if test="@西文字体引用_4129">
              <a:buFont>
                <xsl:attribute name="typeface">
                  <xsl:variable name="font" select="@西文字体引用_4129"/>
                  <!--xsl:value-of select ="@字:西文字体引用"/-->
                  <!--09.9.25黎美秀改-->
                  <xsl:value-of select="//式样:字体声明_990D[@标识符_9902=$font]/@名称_9903"/>
                </xsl:attribute>
                <xsl:attribute name="pitchFamily">2</xsl:attribute>
                <!--end liuyin 20130417 修改bug2837 uof->ooxml回归测试的功能测试中，字体集功能点转换后的文档需要修复-->
                <!--<xsl:attribute name="charset">2</xsl:attribute>-->
                <xsl:attribute name="charset">0</xsl:attribute>
                <!--end liuyin 20130417 修改bug2837 uof->ooxml回归测试的功能测试中，字体集功能点转换后的文档需要修复-->
              </a:buFont>
            </xsl:if>
            <xsl:if test="not(@西文字体引用_4129) and @中文字体引用_412A">
              <a:buFont>
                <xsl:attribute name="typeface">
                  <!--09.9.25黎美秀改-->
                  <!--xsl:value-of select ="@字:中文字体引用"/-->
                  <xsl:variable name="font" select="@中文字体引用_412A"/>
                  <xsl:value-of select="//式样:字体声明_990D[@标识符_9902=$font]/@名称_9903"/>
                </xsl:attribute>
              </a:buFont>
            </xsl:if>
          </xsl:for-each>

			<xsl:for-each select="字:符号字体_4116/字:下划线_4136">
				
			</xsl:for-each>
          <!--09.9.25黎美秀改-->
          <xsl:choose>
            <!--09.10.31 黎美秀修改 增加图片项目符号-->
            <xsl:when test="字:编号格式_4119">
              <a:buAutoNum>
                <xsl:for-each select="字:编号格式_4119">
                  <xsl:attribute name="type">
                    <xsl:choose>
                      <!--xsl:when test=".='decimal'">arabicPlain</xsl:when-->
                      <!--2011-4-4罗文甜，修改bug。-->
                      <xsl:when test=".='decimal'">arabicPeriod</xsl:when>
                      <xsl:when test=".='upper-roman'">
                        <xsl:choose>
                          <xsl:when test="../字:编号格式表示_411A='%1)'">romanUcParenR</xsl:when>
                          <xsl:when test="../字:编号格式表示_411A='(%1)'">romanUcParenBoth</xsl:when>
                          <xsl:otherwise>romanUcPeriod</xsl:otherwise>
                        </xsl:choose>
                      </xsl:when>
                      <xsl:when test=".='lower-roman'">
                        <xsl:choose>
                          <xsl:when test="../字:编号格式表示_411A='%1)'">romanLcParenR</xsl:when>
                          <xsl:when test="../字:编号格式表示_411A='(%1)'">romanLcParenBoth</xsl:when>
                          <xsl:otherwise>romanLcPeriod</xsl:otherwise>
                        </xsl:choose>
                      </xsl:when>
                      <xsl:when test=".='upper-letter'">
                        <xsl:choose>
                          <xsl:when test="../字:编号格式表示_411A='%1)'">alphaUcParenR</xsl:when>
                          <xsl:when test="../字:编号格式表示_411A='(%1)'">alphaUcParenBoth</xsl:when>
                          <xsl:otherwise>alphaUcPeriod</xsl:otherwise>
                        </xsl:choose>
                      </xsl:when>
                      <xsl:when test=".='lower-letter'">
                        <xsl:choose>
                          <xsl:when test="../字:编号格式表示_411A='%1)'">alphaLcParenR</xsl:when>
                          <xsl:when test="../字:编号格式表示_411A='(%1)'">alphaLcParenBoth</xsl:when>
                          <xsl:otherwise>alphaLcPeriod</xsl:otherwise>
                        </xsl:choose>
                      </xsl:when>
                      <xsl:when test=".='ordinal' or .='cardinal-text' or .='ordinal-text' or .='hex'">arabicPlain</xsl:when>
                      <xsl:when test=".='decimal-full-width'">
                        <xsl:choose>
                          <xsl:when test="../字:编号格式表示_411A='%1.'">arabicDbPeriod</xsl:when>
                          <xsl:otherwise>arabicDbPlain</xsl:otherwise>
                        </xsl:choose>
                      </xsl:when>
                      <xsl:when test=".='decimal-half-width'">circleNumDbPlain</xsl:when>
                      <xsl:when test=".='decimal-enclosed-circle'">circleNumWdWhitePlain</xsl:when>
                      <xsl:when test=".='decimal-enclosed-fullstop'">arabicDbPeriod</xsl:when>
                      <xsl:when test=".='decimal-enclosed-paren'">arabicParenBoth</xsl:when>
                      <xsl:when test=".='decimal-enclosed-circle-chinese'">circleNumWdWhitePlain</xsl:when>
                      <xsl:when test=".='ideograph-enclosed-circle' or .='ideograph-traditional' or .='ideograph-zodiac' or .='chinese-legal-simplified'">ea1ChsPeriod</xsl:when>
                      <!--xsl:when test=".='chinese-counting'">ea1ChsPlain</xsl:when-->
                      <!--2011-4-4罗文甜 修改bug -->
                      <xsl:when test=".='chinese-counting'">ea1JpnChsDbPeriod</xsl:when>
                    </xsl:choose>
                  </xsl:attribute>
                </xsl:for-each>
                <xsl:for-each select="字:起始编号_411F">
                  <xsl:attribute name="startAt">
                    <xsl:value-of select="."/>
                  </xsl:attribute>
                </xsl:for-each>
              </a:buAutoNum>
            </xsl:when>
            <xsl:when test="字:图片符号_411B">
              <a:buBlip>
                <a:blip>
                  <xsl:attribute name="r:embed">
                    <!--09.12.21 马有旭 修改<xsl:choose >
                  <xsl:when test ="contains(字:图片符号引用,'image')">
                    <xsl:value-of select="concat('rId',substring-after(字:图片符号引用,'image'))"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="concat('rId',substring-after(字:图片符号引用,'OBJ000'))"/>
                  </xsl:otherwise>
                </xsl:choose>-->
                    <xsl:value-of select="concat('rId',字:图片符号_411B/@引用_411C)"/>
                  </xsl:attribute>
                </a:blip>
              </a:buBlip>
            </xsl:when>
            <!--
            2010.2.4 黎美秀修改 增加项目符号为“-”的情况
            去掉
            <xsl:when test="字:项目符号!='–'">
            改为   <xsl:when test="字:项目符号">
            
            -->
            <xsl:when test="字:项目符号_4115">
              <xsl:choose>
                <xsl:when test="字:项目符号_4115=''">
                  <!--圆圈-->
                  <a:buChar char="l"/>
                </xsl:when>
				  <xsl:when test="字:项目符号_4115='•'">
					  <a:buChar char="•"/>
				  </xsl:when>
                <xsl:when test="字:项目符号_4115=''">
                  <!--方块-->
                  <a:buChar char="n"/>
                </xsl:when>
                <xsl:when test="字:项目符号_4115=''">
                  <!--菱形-->
                  <a:buChar char="u"/>
                </xsl:when>
                <xsl:when test="字:项目符号_4115=''">
                  <!--对钩-->
                  <a:buChar char="ü"/>
                </xsl:when>
                <xsl:when test="字:项目符号_4115=''">
                  <!--箭头-->
                  <a:buChar char="Ø"/>
                </xsl:when>
				 
                <xsl:otherwise>
                  <a:buChar>
                    <xsl:attribute name="char">
                      <xsl:value-of select="字:项目符号_4115"/>
                    </xsl:attribute>
                  </a:buChar>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <a:buNone/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>
