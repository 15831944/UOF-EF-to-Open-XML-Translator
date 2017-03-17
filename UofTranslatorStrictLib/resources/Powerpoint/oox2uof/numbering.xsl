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
<xsl:stylesheet version="1.0"
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
xmlns:vt="http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes"
xmlns:cus="http://purl.oclc.org/ooxml/officeDocument/customProperties"
                
 xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph"
xmlns:式样="http://schemas.uof.org/cn/2009/styles">
  <xsl:import href="fill.xsl"/>
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="2.0"/>
  <xsl:template name="numbering">
	  <式样:自动编号集_990E>
		  <!--<字:自动编号_4124>-->
			  <!--<xsl:attribute name="标识符_4100">bn00000</xsl:attribute>
			  <xsl:attribute name="名称_4122">bn00000</xsl:attribute>
			  <xsl:attribute name="是否多级编号_4125">false</xsl:attribute>-->
			  <!--<xsl:attribute name="父编号引用_4123"></xsl:attribute>--><!--
			  --><!--父编号引用有待继续，李娟2011-11-09--><!--
			  --><!--字:标识符="bn0" 字:名称="bn0" 字:多级编号="false"-->
			  <!--<字:级别_4112>
				  <xsl:attribute name="级别值_4121">1</xsl:attribute>
			  </字:级别_4112>-->
			  <!--<字:级别 uof:locID="t0159" uof:attrList="级别值 编号对齐方式 尾随字符" 字:级别值="1" 字:编号对齐方式="left" 字:尾随字符="none"/>-->
      <!-- 修复bug：图形内项目编号转换不正确  liqiuling 2013-03-05 start-->
      <xsl:for-each select="//p:titleStyle|//p:bodyStyle|//p:otherStyle|//a:lstStyle">
          <xsl:call-template name="numSet"/>
      </xsl:for-each>
      <xsl:for-each select="//a:p">
        <xsl:if test="not (a:pPr/a:buAutoNum/@type = preceding-sibling::a:p/a:pPr[a:buAutoNum]/a:buAutoNum/@type)">
        <xsl:call-template name="numSet"/>
        </xsl:if>
      </xsl:for-each>
      <!-- 修复bug：图形内项目编号转换不正确  liqiuling 2013-03-05 end-->
		  <!--</字:自动编号_4124>-->
    </式样:自动编号集_990E>
  </xsl:template>

  <xsl:template name="numSet">
    <xsl:for-each select="*">
      <xsl:if test="a:buChar or a:buAutoNum or a:buClr or a:buSzPct or a:buSztPts or a:buFont or a:buBlip">
        <!--9.29 黎美秀改 此处没有起作用-->
        <!--2010.2.4  黎美秀修改 没有解决引用问题，光解决了定义问题
         <xsl:if test="not(current()/a:buChar/@char =preceding-sibling::*[a:buChar]/a:buChar/@char) and not (current()/a:buAutoNum/@type = preceding-sibling::*[a:buAutoNum]/a:buAutoNum/@type)">
          
        -->
		  <!--修改自动编号元素名称 李娟2011-11-09-->
        <xsl:if test="not (preceding-sibling::a:p/a:pPr[a:buAutoNum]/a:buAutoNum/@type = a:pPr/a:buAutoNum/@type)">
			<字:自动编号_4124>
          <!--<字:自动编号 uof:locID="t0169" uof:attrList="标识符 名称 父编号引用 多级编号">-->
            <xsl:if test="@id">
				
				<xsl:attribute name="标识符_4100">
              <!--<xsl:attribute name="字:标识符">-->
                <xsl:value-of select="concat('bn',@id)"/>
              </xsl:attribute>
              <!--2011-1-20罗文甜：增加名称-->
				<xsl:attribute name="名称_4122">
              <!--<xsl:attribute name="字:名称">-->
                <xsl:value-of select="concat('bn',@id)"/>
              </xsl:attribute>
            </xsl:if>
            <!--p/pPr/bu...的情况 -->
            <xsl:if test="not(@id)">
				
              <xsl:attribute name="标识符_4100">
                <xsl:value-of select="concat('bn',translate(../@id,':',''))"/>
              </xsl:attribute>
              <xsl:attribute name="名称_4122">
                <xsl:value-of select="concat('bn',translate(../@id,':',''))"/>
              </xsl:attribute>
            </xsl:if>
				<!--有待扩充<xsl:attribute name="父编号引用_4123"></xsl:attribute>-->
            <xsl:attribute name="是否多级编号_4125"><xsl:value-of select="'false'"/></xsl:attribute>
				<字:级别_4112>
            <!--<字:级别 uof:locID="t0159" uof:attrList="级别值 编号对齐方式 尾随字符" 字:级别值="1" 字:编号对齐方式="left" 字:尾随字符="none">-->
              <!--项目符号和list在uof中的表示是不同的，是图片的还没有考虑-->
					<xsl:attribute name="级别值_4121">1</xsl:attribute>
              <xsl:choose>
                <xsl:when test="a:buChar">
					<!--<字:编号对齐方式_4113></字:编号对齐方式_4113>
					<字:尾随字符_4114>none</字:尾随字符_4114>-->
         <!--09.9.25黎美秀改-->
					<字:项目符号_4115>
                  <!--<字:项目符号 uof:locID="t0171">-->
                    <xsl:for-each select ="a:buChar/@char">
                      <xsl:choose>
                        <xsl:when test=".='l'">
                          <xsl:value-of select="''"/>
                        </xsl:when>
                        <xsl:when test=".='n'">
                          <xsl:value-of select="''"/>
                        </xsl:when>
                        <xsl:when test=".='u'">
                          <xsl:value-of select="''"/>
                        </xsl:when>
                        <xsl:when test=".='p'">
                          <xsl:value-of select="''"/>
                        </xsl:when>
                        <xsl:when test=".='ü'">
                          <xsl:value-of select="''"/>
                        </xsl:when>
                        <xsl:when test=".='Ø'">
                          <xsl:value-of select="''"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select ="."/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:for-each>
                  <!--</字:项目符号>-->
					</字:项目符号_4115>
                  <xsl:if test="a:buClr|a:buSzPct|a:buSztPts|a:buFont">
					  <字:符号字体_4116>
						  <!--<xsl:attribute name="式样引用_4247"></xsl:attribute>-->
                    <!--<字:符号字体 uof:locID="t0160" uof:attrList="式样引用">-->
              <字:字体_4128>
                      <!--<字:字体 uof:locID="t0088" uof:attrList="西文字体引用 中文字体引用 特殊字体引用 西文绘制 字号 相对字号 颜色">-->
                        <xsl:if test="a:buFont">
                          <xsl:variable name="bufont" select="translate(translate(translate(a:buFont/@typeface,' ',''),'+','s'),'-','d')"/>
                          <xsl:attribute name="西文字体引用_4129">
                            <xsl:value-of select="$bufont"/>
                          </xsl:attribute>
                          <xsl:attribute name="中文字体引用_412A">
                            <xsl:value-of select="$bufont"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="a:buSzPct/@val and a:defRPr/@val">
                          <xsl:attribute name="字号_412D">
                            <xsl:value-of select="number(a:defRPr/@sz * a:buSzPct/@val) div 10000000"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test ="not(a:defRPr)">

                        </xsl:if>
                        <xsl:if test="a:buSzPts">
                          <xsl:attribute name="字号_412D">
                            <xsl:value-of select="number(a:buSzPts/@val) div 100"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="a:buClr">
                          <xsl:attribute name="颜色_412F">
                            <xsl:apply-templates select="a:buClr"/>
                          </xsl:attribute>
                        </xsl:if>
                      <!--</字:字体>-->
						  </字:字体_4128>
                    <!--</字:符号字体>-->
					  </字:符号字体_4116>
                  </xsl:if>
                </xsl:when>

                <xsl:when test="a:buAutoNum">
                  <xsl:if test="a:buClr|a:buSzPct|a:buSzPts|a:buFont">
                    <!--<字:符号字体 uof:locID="t0160" uof:attrList="式样引用">-->
					  <字:符号字体_4116>
						  <字:字体_4128>
					  <!--<字:字体 uof:locID="t0088" uof:attrList="西文字体引用 中文字体引用 特殊字体引用 西文绘制 字号 相对字号 颜色">-->
                        <xsl:if test="a:buFont">
                          <xsl:variable name="bufont" select="translate(translate(translate(a:buFont/@typeface,' ',''),'+','s'),'-','d')"/>
                          <xsl:attribute name="西文字体引用_4129">
                            <!--liuyin 20121112 编号样式不正确<xsl:value-of select="$bufont"/>-->
                            <xsl:value-of select="../a:r/a:rPr/a:latin/@typeface"/>
                          </xsl:attribute>
                          <xsl:attribute name="中文字体引用_412A">
                            <!--liuyin 20121112 编号样式不正确<xsl:value-of select="$bufont"/>-->
                            <xsl:value-of select="../a:r/a:rPr/a:ea/@typeface"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="a:buSzPct/@sz and a:defRPr/@sz">
                          <xsl:attribute name="字号_412D">
                            <xsl:value-of select="number(a:defRPr/@sz * a:buSzPct/@val) div 10000000"/>
                          </xsl:attribute>
                        </xsl:if>
                <!-- 修复bug：图形内项目编号转换不正确  liqiuling 2013-03-05 start-->
                   <xsl:if test="a:buSzPct/@val">
                     
                     <!--2013-11-5, tangjiang, Strict OOXML转到UOF相对字号用百分号表示的处理 start -->
                     <xsl:variable name="buSzPctVal" select="a:buSzPct/@val"/>
                     <xsl:variable name="buSzPct">
                       <xsl:choose>
                         <xsl:when test="contains($buSzPctVal,'%')">
                           <xsl:value-of select="substring($buSzPctVal,1,string-length($buSzPctVal)-1)"/>
                         </xsl:when>
                         <xsl:otherwise>
                           <xsl:value-of select="$buSzPctVal"/>
                         </xsl:otherwise>
                       </xsl:choose>
                     </xsl:variable>
                     <xsl:attribute name="相对字号_412E">
                       <xsl:value-of select="$buSzPct"/>
                     </xsl:attribute>
                     <!--
                     <xsl:attribute name="相对字号_412E">
                       <xsl:value-of select="number(a:buSzPct/@val) div 1000"/>
                     </xsl:attribute>
                     -->
                     <!--end-->
                     
                  </xsl:if>
                <!-- 修复bug：图形内项目编号转换不正确  liqiuling 2013-03-05 end-->
                        <xsl:if test="a:buSzPts/@val">
                          <xsl:attribute name="字号_412D">
                            <xsl:value-of select="number(a:buSzPts/@val) div 100"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="a:buClr">
                          <xsl:attribute name="颜色_412F">
                            <!--xsl:apply-templates select="a:buClr/a:srgbClr"/-->
                            <!--modify by linyh-->
                            <xsl:apply-templates select="a:buClr"/>
                          </xsl:attribute>
                        </xsl:if>
						  </字:字体_4128>
                    <!--</字:符号字体>-->
					  </字:符号字体_4116>
				  </xsl:if>
                  <xsl:for-each select ="a:buAutoNum/@type">
                    <xsl:choose>
                      <xsl:when test =".='alphaLcParenBoth'">
                        <字:编号格式_4119>lower-letter</字:编号格式_4119>
						  <字:编号格式表示_411A>(%1)</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='alphaLcParenR'">
						  <字:编号格式_4119>lower-letter</字:编号格式_4119>
						  <字:编号格式表示_411A>%1)</字:编号格式表示_411A>
                      
				  </xsl:when>
                      <xsl:when test =".='alphaLcPeriod'">
                        <字:编号格式_4119>lower-letter</字:编号格式_4119>
						  <字:编号格式表示_411A>%1.</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='alphaUcParenBoth'">
                        <字:编号格式_4119>upper-letter</字:编号格式_4119>
						  <字:编号格式表示_411A>(%1)</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='alphaUcParenR'">
                        <字:编号格式_4119>upper-letter</字:编号格式_4119>
						  <字:编号格式表示_411A>%1)</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='alphaUcPeriod'">
						  <字:编号格式_4119>upper-letter</字:编号格式_4119>
						  <字:编号格式表示_411A>%1.</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='arabicParenBoth'">
						  <字:编号格式_4119>decimal</字:编号格式_4119>
						  <字:编号格式表示_411A>(%1)</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='arabicParenR'">
						  <字:编号格式_4119>decimal</字:编号格式_4119>
						  <字:编号格式表示_411A>%1)</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='arabicPeriod'">
						  <字:编号格式_4119>decimal</字:编号格式_4119>
						  <字:编号格式表示_411A>%1.</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='arabicPlain'">
						  <字:编号格式_4119>decimal</字:编号格式_4119>
						  <字:编号格式表示_411A>%1</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='romanLcParenBoth'">
                        <字:编号格式_4119>lower-roman</字:编号格式_4119>
						  <字:编号格式表示_411A>(%1)</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='romanLcParenR'">
                        <字:编号格式_4119>lower-roman</字:编号格式_4119>
						  <字:编号格式表示_411A>%1)</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='romanLcPeriod'">
                        <字:编号格式_4119>lower-roman</字:编号格式_4119>
						  <字:编号格式表示_411A>%1.</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='romanUcParenBoth'">
                        <字:编号格式_4119>upper-roman</字:编号格式_4119>
						  <字:编号格式表示_411A>(%1)</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='romanUcParenR'">
                        <字:编号格式_4119>upper-roman</字:编号格式_4119>
						  <字:编号格式表示_411A>%1)</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='romanUcPeriod'">
                        <字:编号格式_4119>upper-roman</字:编号格式_4119>
						  <字:编号格式表示_411A>%1.</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='arabicDbPeriod'">
                        <字:编号格式_4119>ddecimal-enclosed-fullstop</字:编号格式_4119>
                      </xsl:when>
                      <xsl:when test =".='arabicDbPlain'">
                        <字:编号格式_4119>decimal-full-width</字:编号格式_4119>
                      </xsl:when>
                      <xsl:when test =".='circleNumDbPlain'">
                        <字:编号格式_4119>decimal-enclosed-circle-chinese</字:编号格式_4119>
						  <字:编号格式表示_411A>%1</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='circleNumWdWhitePlain' or .='circleNumWdBlackPlain'">
                        <字:编号格式_4119>decimal-enclosed-circle</字:编号格式_4119>
                      </xsl:when>
                      <xsl:when test =".='ea1ChsPeriod' or '.=ea1ChtPeriod' or .='ea1JpnKorPeriod'">
                        <字:编号格式_4119>chinese-counting</字:编号格式_4119>
						  <字:编号格式表示_411A>%1.</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='ea1JpnChsDbPeriod'">
                        <字:编号格式_4119>chinese-counting</字:编号格式_4119>
						  <字:编号格式表示_411A>%1. </字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:when test =".='ea1ChsPlain' or .='ea1ChtPlain' or .='ea1JpnKorPlain'">
                        <字:编号格式_4119>chinese-counting</字:编号格式_4119>
						  <字:编号格式表示_411A>%1</字:编号格式表示_411A>
                      </xsl:when>
                      <xsl:otherwise>
                        <字:编号格式_4119>decimal</字:编号格式_4119>
						  <字:编号格式表示_411A>%1.</字:编号格式表示_411A>
                      </xsl:otherwise>
                      <!--xsl:when test =".='arabicPeriod'">
								<字:编号格式_4119 uof:locID="t0162">decimal</字:编号格式>
								<字:编号格式表示 uof:locID="t0163">(%1)</字:编号格式表示>
							</xsl:when-->
                    </xsl:choose>
                  </xsl:for-each>
                  <!--2011-1-14罗文甜：自动编号修改BUG-->
                  <xsl:if test="a:buAutoNum">
					  
                    <xsl:choose>
                      <xsl:when test="a:buAutoNum/@startAt">
						 
						  <字:起始编号_411F>
                        <!--<字:起始编号 uof:locID="t0167">-->
                          <xsl:value-of select="a:buAutoNum/@startAt"/>
                        <!--</字:起始编号>-->
						  </字:起始编号_411F>
                      </xsl:when>
                      <xsl:otherwise>
                        <字:起始编号_411F>1</字:起始编号_411F>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:if>
                  <!--xsl:if test="a:buAutoNum/@startAt">
                    <字:起始编号 uof:locID="t0167">
                      <xsl:value-of select="a:buAutoNum/@startAt"/>
                    </字:起始编号>
                  </xsl:if-->
                </xsl:when>
                <xsl:when test ="a:buBlip">
					<字:图片符号_411B>
						<xsl:attribute name="引用_411C">
					<xsl:call-template name="findRelatedRelationships">
                      <xsl:with-param name="id">
					<xsl:value-of select=".//a:blip/@r:embed"/>
					</xsl:with-param>
					  </xsl:call-template>
						</xsl:attribute>
					</字:图片符号_411B>
                </xsl:when>
              </xsl:choose>
              <!---add by linyh 项目符号是图片>
          <xsl:if test="name(descendant-or-self::*) = 'a:buBlip'">
            <xsl:apply-templates select="//..//a:buBlip" mode="autoNumbering"/>
          </xsl:if-->

          <!--2014-01-14, tangjiang, PPT自动编号无缩进、制表位置，添上影响显示效果，故注释掉 start -->
          <!--
					<字:缩进_411D>
						<字:左_410E>
							<字:绝对_4107>
								<xsl:attribute name="值_410F">0.0</xsl:attribute>
							</字:绝对_4107>
						</字:左_410E>
						<字:右_4110>
							<字:绝对_4107>
								<xsl:attribute name="值_410F">0.0</xsl:attribute>
							</字:绝对_4107>
						</字:右_4110>
						<字:首行_4111>
							<字:绝对_4107>
								<xsl:attribute name="值_410F">0.0</xsl:attribute>
							</字:绝对_4107>
						</字:首行_4111>
					</字:缩进_411D>
          <字:制表符位置_411E>0.0</字:制表符位置_411E>
					-->
          <!-- end, tangjiang, PPT自动编号无缩进、制表位置，添上影响显示效果，故注释掉 -->
          
				</字:级别_4112>
				</字:自动编号_4124>
			</xsl:if>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <!--add by linyh:项目符号颜色-->
  <xsl:template match="a:buClr">
    <xsl:call-template name="colorChoice"/>
  </xsl:template>

</xsl:stylesheet>
