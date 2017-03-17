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
<!--
<Author>Fang Chunyan(BITI)</Author>
-->
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/xpath-functions"
  xmlns:xdt="http://www.w3.org/2005/xpath-datatypes" 
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main">
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
  <!--<xsl:template match="/">
    <xsl:apply-templates select="uof:UOF" mode="preRev"/>
  </xsl:template>-->
  <!--<xsl:template match="uof:UOF" mode="preRev">
		<uof:UOF uof:language="cn" uof:version="1.0" uof:locID="u0000" xml:space="preserve">
			<xsl:copy-of select="uof:元数据"/>
			<xsl:if test="uof:书签集">
				<xsl:copy-of select="uof:书签集"/>
			</xsl:if>
			<xsl:if test="uof:链接集">
				<xsl:copy-of select="uof:链接集"/>
			</xsl:if>
			<xsl:copy-of select="uof:式样集"/>
			<xsl:if test="uof:对象集">
				<xsl:copy-of select="uof:对象集"/>
			</xsl:if>
			<xsl:if test="uof:用户数据集">
				<xsl:copy-of select="uof:用户数据集"/>
			</xsl:if>
			<uof:文字处理 uof:locID="u0047">
				<xsl:if test="uof:文字处理/字:公用处理规则">
					<xsl:apply-templates select="uof:文字处理/字:公用处理规则" mode="preRev"/>
				</xsl:if>
				<xsl:apply-templates select="uof:文字处理/字:主体" mode="preRev"/>
			</uof:文字处理>
		</uof:UOF>
	</xsl:template>-->
  <xsl:template match="规则:公用处理规则_B665" mode="preRev">
    <规则:公用处理规则_B665>
      <xsl:if test="规则:长度单位_B666">
        <xsl:copy-of select="规则:长度单位_B666"/>
      </xsl:if>
      <xsl:if test=".//规则:文档设置_B600">
        <xsl:copy-of select=".//规则:文档设置_B600"/>
      </xsl:if>
      <xsl:if test="规则:用户集_B667">
        <xsl:copy-of select="规则:用户集_B667"/>
      </xsl:if>
      <xsl:if test=".//规则:修订信息集_B60E">
        <xsl:copy-of select=".//规则:修订信息集_B60E"/>
      </xsl:if>
      <xsl:if test="规则:批注集_B669">
        <规则:批注集_B669>
          <xsl:apply-templates select="规则:批注集_B669" mode="preRev"/>
        </规则:批注集_B669>
      </xsl:if>
    </规则:公用处理规则_B665>
  </xsl:template>
  <xsl:template match="规则:批注集_B669" mode="preRev">
    <xsl:for-each select="规则:批注_B66A">
      <规则:批注_B66A>
        <xsl:if test="@区域引用_41CE">
          <xsl:attribute name="区域引用_41CE">
            <xsl:value-of select="@区域引用_41CE"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="@作者_41DD">
          <xsl:attribute name="作者_41DD">
            <xsl:value-of select="@作者_41DD"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="@日期_41DE">
          <xsl:attribute name="日期_41DE">
            <xsl:value-of select="@日期_41DE"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="@作者缩写_41DF">
          <xsl:attribute name="作者缩写_41DF">
            <xsl:value-of select="@作者缩写_41DF"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:for-each select="字:段落_416B | 字:文字表_416C">
          <xsl:choose>
            <xsl:when test="name(.)='字:段落_416B'">
              <字:段落_416B>
                <xsl:for-each select="./node()">
                  <xsl:call-template name="nodeloop"/>
                </xsl:for-each>
              </字:段落_416B>
            </xsl:when>
            <xsl:when test="name(.)='字:文字表_416C'">
              <xsl:call-template name="revTable"/>
            </xsl:when>
          </xsl:choose>
        </xsl:for-each>
      </规则:批注_B66A>
    </xsl:for-each>
  </xsl:template>
  
  <xsl:template match="字:文字处理文档_4225" mode="preRev">
      <xsl:for-each select="字:段落_416B | 字:分节_416A | 字:文字表_416C">
        <xsl:choose>
          <xsl:when test="name(.)='字:段落_416B'">
            <字:段落_416B>
              <xsl:for-each select="./node()">
                <xsl:call-template name="nodeloop"/>
              </xsl:for-each>
              <!--xsl:apply-templates select="child::node()" mode="rev"/-->
            </字:段落_416B>
          </xsl:when>
          <xsl:when test="name(.)='字:分节_416A'">
            <字:分节_416A>
              <xsl:for-each select="node()">
                <xsl:choose>
                  <xsl:when test="name(.)='字:节属性_421B'">
                    <字:节属性_421B>
                      <xsl:call-template name="sectionPreProcessing"/>
                    </字:节属性_421B>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:copy-of select="."/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </字:分节_416A>              
          </xsl:when>
          <xsl:when test="name(.)='字:文字表_416C'">
            <xsl:call-template name="revTable"/>           
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
  </xsl:template>
  <xsl:template match="字:分节_416A" mode="preRev"><!--cxl,2012.2.16对第一个字:分节_416A进行处理，放于所有段落之后-->
    <!--<xsl:for-each select="node()">
      <xsl:choose>
        <xsl:when test="name(.)='字:分节_416A' and position(parent::node()/字:分节_416A)='1'">-->
          <字:分节_416A>
            <xsl:for-each select="node()">
              <xsl:choose>
                <xsl:when test="name(.)='字:节属性_421B'">
                  <字:节属性_421B>
                    <xsl:call-template name="sectionPreProcessing"/>
                  </字:节属性_421B>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:copy-of select="."/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
          </字:分节_416A>
        <!--</xsl:when>
      </xsl:choose>
    </xsl:for-each>-->
  </xsl:template>
  
  <!--cxl,2012/2/15,for-each./node(),current node:字:段落_416B-->
  <xsl:template name="nodeloop">
    <xsl:choose>
      <xsl:when test="name(.)='字:段落属性_419B'">
        <xsl:if test="not(preceding-sibling::字:修订开始_421F/@标识符_4220 = following::字:修订结束_4223/@开始标识引用_4224)">
          <xsl:call-template name="pPrprocess"/>
        </xsl:if>
        <xsl:if test="preceding-sibling::字:修订开始_421F/@标识符_4220 = following-sibling::字:修订结束_4223/@开始标识引用_4224">
          <!--do nothing-->
        </xsl:if>
      </xsl:when>
      <xsl:when test="name(.)='字:句_419D'">
        <xsl:for-each select=".">
          <xsl:if test="not(preceding-sibling::字:修订开始_421F/@标识符_4220 = following::字:修订结束_4223/@开始标识引用_4224)">
            <xsl:call-template name="runRev"/>
          </xsl:if>
          <xsl:if test="preceding-sibling::字:修订开始_421F/@标识符_4220 = following-sibling::字:修订结束_4223/@开始标识引用_4224">
            <!--do nothing-->
          </xsl:if>
          <!--yx,add pages manage here2010.2.1-->
          <xsl:variable name="pagestrans">
            <xsl:if test="字:文本串_415B">
              <xsl:value-of select="字:文本串_415B"/>
            </xsl:if>

          </xsl:variable>
          <xsl:if test="$pagestrans='PAGE \* MERGEFORMAT'">
            <xsl:copy-of select="."/>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="name(.)='字:修订开始_421F'">
        <xsl:variable name="revid" select="@标识符_4220"/>
        <字:修订开始_421F>
          <xsl:attribute name="标识符_4220">
            <xsl:value-of select="@标识符_4220"/>
          </xsl:attribute>
          <xsl:attribute name="类型_4221">
            <xsl:value-of select="@类型_4221"/>
          </xsl:attribute>
          <xsl:attribute name="修订信息引用_4222">
            <xsl:value-of select="@修订信息引用_4222"/>
          </xsl:attribute>
          <xsl:for-each
            select="following-sibling::node()[following::字:修订结束_4223[1]/@开始标识引用_4224 = $revid]">
            <xsl:if test="name(.)='字:段落属性_419B'">
              <xsl:call-template name="pPrprocess"/>
            </xsl:if>
            <xsl:if test="name(.)='字:句_419D'">
              <xsl:call-template name="runRev"/>
            </xsl:if>
            <xsl:if test="name(.)='字:修订开始_421F'">
              <xsl:call-template name="nodeloop"/>
            </xsl:if>
          </xsl:for-each>
        </字:修订开始_421F>
      </xsl:when>
      <xsl:when test="name(.)='字:修订结束_4223'">
        <!--do nothing-->
      </xsl:when>
      <xsl:when test="name(.)='字:域开始_419E'">
        <字:域开始_419E>
          <xsl:attribute name="类型_416E">
            <xsl:value-of select="./@类型_416E"/>
          </xsl:attribute>
          <xsl:call-template name="regionstart"/>
        </字:域开始_419E>
      </xsl:when>
      <xsl:otherwise>
        <xsl:copy-of select="."/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="runRev">
    <字:句_419D>
      <xsl:for-each select="node()">
        <xsl:choose>
          <xsl:when test="name(.)='字:修订开始_421F'">
            <xsl:variable name="revid1" select="@标识符_4220"/>
            <字:修订开始_421F>
              <xsl:attribute name="标识符_4220">
                <xsl:value-of select="@标识符_4220"/>
              </xsl:attribute>
              <xsl:attribute name="类型_4221">
                <xsl:value-of select="@类型_4221"/>
              </xsl:attribute>
              <xsl:attribute name="修订信息引用_4222">
                <xsl:value-of select="@修订信息引用_4222"/>
              </xsl:attribute>
              <xsl:copy-of
                select="following-sibling::node()[following::字:修订结束_4223[1]/@开始标识引用_4224 = $revid1]"/>
            </字:修订开始_421F>
          </xsl:when>
          <xsl:when test="name(.)='字:修订结束_4223'">
            <!--do nothing-->
          </xsl:when>
          <!--<xsl:when test="name(.)='字:区域结束_4167' and preceding-sibling::字:区域开始_4165[@标识符_4100]/@类型_413B='hyperlink'">
            <xsl:if test="preceding-sibling::字:区域开始_4165[@类型_413B='hyperlink']">
              <xsl:variable name="id" select="@标识符引用_4168"/>
              <xsl:variable name="hyperid"
                select="preceding-sibling::字:区域开始_4165[@类型_413B='hyperlink']/@标识符_4100"/>
              <xsl:if
                test="not(following-sibling::字:区域结束_4167[@标识符引用_4168 = $hyperid]) and preceding-sibling::字:区域开始_4165[@标识符_4100=$id]/@类型_413B!='hyperlink'">
                <xsl:copy-of select="."/>
              </xsl:if>
            </xsl:if>
            <xsl:if test="not(preceding-sibling::字:区域开始_4165[@标识符_4100]/@类型_413B='hyperlink')">
              <xsl:copy-of select="."/>
            </xsl:if>
          </xsl:when>
          <xsl:when test="name(.)='字:区域开始_4165'">
            <xsl:if test="@类型_413B='hyperlink'">
              <xsl:variable name="regionid" select="@标识符_4100"/>
              <字:区域开始_4165>
                <xsl:attribute name="标识符_4100">
                  <xsl:value-of select="@标识符_4100"/>
                </xsl:attribute>
                <xsl:attribute name="类型_413B">
                  <xsl:value-of select="@类型_413B"/>
                </xsl:attribute>
                <xsl:attribute name="名称_4166">
                  <xsl:value-of select="@名称_4166"/>
                </xsl:attribute>
                <xsl:copy-of
                  select="preceding-sibling::字:句属性_4158[not(preceding-sibling::字:修订开始_421F/@标识符_4220 = following-sibling::字:修订结束_4223/@开始标识引用_4224)]"/>
                <xsl:copy-of
                  select="following-sibling::node()[following-sibling::字:区域结束_4167[@标识符引用_4168 = $regionid]]"
                />
              </字:区域开始_4165>
            </xsl:if>
            <xsl:if test="@字:类型!='hyperlink'">
              <xsl:if test="preceding-sibling::字:区域开始_4165[@标识符_4100]/@类型_413B='hyperlink'">
                <xsl:variable name="rsId">
                  <xsl:value-of select="preceding-sibling::字:区域开始_4165[@类型_413B='hyperlink']/@标识符_4100"/>
                </xsl:variable>
                <xsl:if test="not(following-sibling::字:区域结束_4167[@标识符引用_4168 = $rsId])">
                  <xsl:copy-of select="."/>
                </xsl:if>
              </xsl:if>
              <xsl:if test="not(preceding-sibling::字:区域开始_4165[@标识符_4100]/@类型_413B='hyperlink')">
                <xsl:copy-of select="."/>
              </xsl:if>
            </xsl:if>
          </xsl:when>-->
          <xsl:when test="字:句属性_4158">
            <xsl:if
              test="not(preceding-sibling::字:修订开始_421F/@标识符_4220 = following-sibling::字:修订结束_4223/@开始标识引用_4224)">
              <xsl:copy-of select="."/>
            </xsl:if>
          </xsl:when>
          <!--yx,add 字:文本串,2010.2.5-->
          <xsl:when test="字:文本串_415B">
            <xsl:if
              test="not(preceding-sibling::字:修订开始_421F/@标识符_4220 = following-sibling::字:修订结束_4223/@开始标识引用_4224)">
              <xsl:copy-of select="self::*"/>
            </xsl:if>
          </xsl:when>
          <!--yx,add 字:文本串over,2010.2.5-->
          <xsl:when test="name(.)='字:脚注_4159'">
            <字:脚注_4159>
              <xsl:if test="@引文体_4157">
                <xsl:attribute name="引文体_4157">
                  <xsl:value-of select="@引文体_4157"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:for-each select="字:段落_416B">
                <字:段落_416B>
                  <xsl:for-each select="./node()">
                    <xsl:call-template name="nodeloop"/>
                  </xsl:for-each>
                </字:段落_416B>
              </xsl:for-each>
            </字:脚注_4159>
          </xsl:when>
          <xsl:when test="name(.)='字:尾注_415A'">
            <字:尾注_415A>
              <xsl:if test="@引文体_4157">
                <xsl:attribute name="引文体_4157">
                  <xsl:value-of select="@引文体_4157"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:for-each select="字:段落_416B">
                <字:段落_416B>
                  <xsl:for-each select="./node()">
                    <xsl:call-template name="nodeloop"/>
                  </xsl:for-each>
                </字:段落_416B>
              </xsl:for-each>
            </字:尾注_415A>
          </xsl:when>

          <xsl:when
            test="not(preceding-sibling::字:修订开始_421F/@标识符_4220 = following-sibling::字:修订结束_4223/@开始标识引用_4224) and not(preceding-sibling::字:区域开始_4165[@类型_413B='hyperlink']/@标识符_4100 = following-sibling::字:区域结束_4167/@标识符引用_4168)">
            <xsl:copy-of select="."/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </字:句_419D>
  </xsl:template>
  
  <xsl:template name="regionstart">
    <xsl:for-each select="./node()">
      <xsl:choose>
        <xsl:when test="name(.)='字:域代码_419F'">
          <字:域代码_419F>
            <xsl:for-each select="字:段落_416B">
              <字:段落_416B>
                <xsl:for-each select="./node()">
                  <xsl:call-template name="nodeloop"/>
                </xsl:for-each>
              </字:段落_416B>
            </xsl:for-each>
          </字:域代码_419F>
        </xsl:when>
        <xsl:when test="name(.)='字:句_419D'">
          <!--yx,pages call not manage here2010.2.1-->
          <!--yx,add pages translation here current node:字:域代码/字:句2010.2.1-->
          <xsl:if test="not(preceding-sibling::字:修订开始_421F/@标识符_4220 = following-sibling::字:修订结束_4223/@开始标识引用_4224)">
            <xsl:call-template name="runRev"/>
          </xsl:if>
          <xsl:if test="preceding-sibling::字:修订开始_421F/@标识符_4220 = following-sibling::字:修订结束_4223/@开始标识引用_4224">
            <!--do nothing-->
          </xsl:if>
          <!--yx,add pages translation here2010.2.1-->
        </xsl:when>
        <xsl:when test="name(.)='字:修订开始_421F'">
          <xsl:variable name="revid" select="@标识符_4220"/>
          <字:修订开始_421F>
            <xsl:attribute name="标识符_4220">
              <xsl:value-of select="@标识符_4220"/>
            </xsl:attribute>
            <xsl:attribute name="类型_4221">
              <xsl:value-of select="@类型_4221"/>
            </xsl:attribute>
            <xsl:attribute name="修订信息引用_4222">
              <xsl:value-of select="@修订信息引用_4222"/>
            </xsl:attribute>
            <xsl:for-each
              select="following-sibling::node()[following-sibling::字:修订结束_4223/@开始标识引用_4224 = $revid]">
              <xsl:if test="name(.)='字:段落属性_419B'">
                <xsl:call-template name="pPrprocess"/>
              </xsl:if>
              <xsl:if test="name(.)='字:句_419D'">
                <xsl:call-template name="runRev"/>
              </xsl:if>
              <xsl:if test="name(.)='字:修订开始_421F'">
                <xsl:call-template name="nodeloop"/>
              </xsl:if>
            </xsl:for-each>
          </字:修订开始_421F>      
      </xsl:when>
        <xsl:when test="name(.)='字:修订结束_4223'">
          <!--do nothing-->
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>

  </xsl:template>
  <!--yx,add pages translation here2010.2.1-->
  <!--xsl:template name="pagestranslation">
    
    <字:文本串 xml:space="preserve" uof:locID="t0109" uof:attrList="标识符">
     <xsl:copy-of select="./字:文本串"/>
    </字:文本串>
  
  </xsl:template-->
  <xsl:template name="pPrprocess">
    <字:段落属性_419B>
      <xsl:if test="@式样引用_419C">
        <xsl:attribute name="式样引用_419C">
          <xsl:value-of select="@式样引用_419C"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:for-each select="node()">
        <xsl:choose>
          <xsl:when test="name(.)='字:分节_416A'">
            <字:分节_416A>
              <xsl:for-each select="node()">
                <xsl:choose>
                  <xsl:when test="name(.)='字:节属性_421B'">
                    <字:节属性_421B>
                      <xsl:call-template name="sectionPreProcessing"/>
                    </字:节属性_421B>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:copy-of select="."/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:for-each>
            </字:分节_416A>
          </xsl:when>
          <xsl:otherwise>
            <xsl:copy-of select="."/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </字:段落属性_419B>
  </xsl:template>

  <xsl:template name="revTable">
    <字:文字表_416C>
      <xsl:for-each select="字:文字表属性_41CC | 字:行_41CD | 字:修订开始_421F | 字:修订结束_4223">
        <xsl:choose>
          <xsl:when test="name(.)='字:文字表属性_41CC'">
            <xsl:copy-of select="."/>
          </xsl:when>
          <xsl:when test="name(.)='字:行_41CD'">
            <字:行_41CD>
              <xsl:for-each select="字:表行属性_41BD | 字:单元格_41BE">
                <xsl:choose>
                  <xsl:when test="name(.)='字:表行属性_41BD'">
                    <xsl:copy-of select="."/>
                  </xsl:when>
                  <xsl:when test="name(.)='字:单元格_41BE'">
                    <xsl:if test="not(./字:跨列占位单元格)">
                      <字:单元格_41BE>
                        <xsl:for-each select="node()">
                          <xsl:choose>
                            <xsl:when test="name(.)='字:单元格属性_41B7'">
                              <xsl:copy-of select="."/>
                            </xsl:when>
                            <xsl:when test="name(.)='字:段落_416B'">
                              <字:段落_416B>
                                <xsl:for-each select="./node()">
                                  <xsl:call-template name="nodeloop"/>
                                </xsl:for-each>
                              </字:段落_416B>
                            </xsl:when>
                            <xsl:when test="name(.)='字:文字表_416C'">
                              <xsl:call-template name="revTable"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:copy-of select="."/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:for-each>
                      </字:单元格_41BE>
                    </xsl:if>
                  </xsl:when>
                </xsl:choose>          
              </xsl:for-each>
            </字:行_41CD>
          </xsl:when>
          <xsl:when test="name(.)='字:修订开始_421F'">
            <xsl:copy-of select="."/>
          </xsl:when>
          <xsl:when test="name(.)='字:修订结束_4223'">
            <xsl:copy-of select="."/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </字:文字表_416C>
  </xsl:template>
  
  <!--cxl,current node:字:节属性_421B,2012.1.11-->
  <xsl:template name="sectionPreProcessing">
    <xsl:for-each select="node()">
      <xsl:choose>
        <xsl:when test="not(name(.)='字:页眉_41F3' or name(.)='字:页脚_41F7')">
          <xsl:copy-of select="."/>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
    <xsl:choose>
      <xsl:when
        test="((//字:分节_416A/字:节属性_421B/字:是否奇偶页页眉页脚不同_41ED='true')and (字:页眉_41F3) and ( (字:页眉_41F3/字:奇数页页眉_41F4) 
                               and (not(字:页眉_41F3/字:偶数页页眉_41F5/*)) and (not(字:是否奇偶页页眉页脚不同_41ED='true')) ))
                    or ((//字:分节_416A/字:节属性_421B/字:是否奇偶页页眉页脚不同_41ED='true')and (字:页脚_41F7) and ( (字:页脚_41F7/字:奇数页页脚_41F8) 
                              and (not(字:页脚_41F7/字:偶数页页脚_41F9/*)) and (not(字:是否奇偶页页眉页脚不同_41ED='true')) ))">
        <xsl:call-template name="hdrftrPreMethodOne"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="hdrftrPreMethodTwo"/>
      </xsl:otherwise>
    </xsl:choose>

  </xsl:template>

  <xsl:template name="hdrftrPreMethodOne">
    <xsl:if test="字:页眉_41F3">
      <字:页眉_41F3>
        <xsl:if test="字:页眉_41F3/字:奇数页页眉_41F4">
          <字:奇数页页眉_41F4>
            <xsl:apply-templates select="字:页眉_41F3/字:奇数页页眉_41F4" mode="preprocessing"/>
          </字:奇数页页眉_41F4>
          <字:偶数页页眉_41F5>
            <xsl:apply-templates select="字:页眉_41F3/字:偶数页页眉_41F5" mode="preprocessing"/>
          </字:偶数页页眉_41F5>
        </xsl:if>
        <xsl:if test="字:页眉_41F3/字:首页页眉_41F6">
          <字:首页页眉_41F6>
            <xsl:apply-templates select="字:页眉_41F3/字:首页页眉_41F6" mode="preprocessing"/>
          </字:首页页眉_41F6>
        </xsl:if>
      </字:页眉_41F3>
    </xsl:if>
    <xsl:if test="字:页脚_41F7">
      <字:页脚_41F7>
        <xsl:if test="字:页脚_41F7/字:奇数页页脚_41F8">
          <字:奇数页页脚_41F8>
            <xsl:apply-templates select="字:页脚_41F7/字:奇数页页脚_41F8" mode="preprocessing"/>
          </字:奇数页页脚_41F8>
          <字:偶数页页脚_41F9>
            <xsl:apply-templates select="字:页脚_41F7/字:偶数页页脚_41F9" mode="preprocessing"/>
          </字:偶数页页脚_41F9>        
       </xsl:if>
        <xsl:if test="字:页脚_41F7/字:首页页脚_41FA">
          <字:首页页脚_41FA>
            <xsl:apply-templates select="字:页脚_41F7/字:首页页脚_41FA" mode="preprocessing"/>
          </字:首页页脚_41FA>
        </xsl:if>
      </字:页脚_41F7>
    </xsl:if>
  </xsl:template>
  <xsl:template name="hdrftrPreMethodTwo">
    <xsl:if test="字:页眉_41F3">
      <字:页眉_41F3>
        <xsl:if test="字:页眉_41F3/字:奇数页页眉_41F4">
          <字:奇数页页眉_41F4>
            <xsl:apply-templates select="字:页眉_41F3/字:奇数页页眉_41F4" mode="preprocessing"/>
          </字:奇数页页眉_41F4>
        </xsl:if>
        <xsl:if test="字:页眉_41F3/字:偶数页页眉_41F5">
          <字:偶数页页眉_41F5>
            <xsl:apply-templates select="字:页眉_41F3/字:偶数页页眉_41F5" mode="preprocessing"/>
          </字:偶数页页眉_41F5>
        </xsl:if>
        <xsl:if test="字:页眉_41F3/字:首页页眉_41F6">
          <字:首页页眉_41F6>
            <xsl:apply-templates select="字:页眉_41F3/字:首页页眉_41F6" mode="preprocessing"/>
          </字:首页页眉_41F6>      
      </xsl:if>
      </字:页眉_41F3>
    </xsl:if>
    <xsl:if test="字:页脚_41F7">
      <字:页脚_41F7>
        <xsl:if test="字:页脚_41F7/字:奇数页页脚_41F8">
          <字:奇数页页脚_41F8>
            <xsl:apply-templates select="字:页脚_41F7/字:奇数页页脚_41F8" mode="preprocessing"/>
          </字:奇数页页脚_41F8>
        </xsl:if>
        <xsl:if test="字:页脚_41F7/字:偶数页页脚_41F9">
          <字:偶数页页脚_41F9>
            <xsl:apply-templates select="字:页脚_41F7/字:偶数页页脚_41F9" mode="preprocessing"/>
          </字:偶数页页脚_41F9>
        </xsl:if>
        <xsl:if test="字:页脚_41F7/字:首页页脚_41FA">
          <字:首页页脚_41FA>
            <xsl:apply-templates select="字:页脚_41F7/字:首页页脚_41FA" mode="preprocessing"/>
          </字:首页页脚_41FA>
        </xsl:if>
      </字:页脚_41F7>
    </xsl:if>
  </xsl:template>
  <!--CXL,current node:字:页脚/字:奇数页页脚 or......2012.1.11-->
  <xsl:template match="字:奇数页页眉_41F4 | 字:首页页眉_41F6 | 字:偶数页页眉_41F5 | 字:奇数页页脚_41F8 | 字:首页页脚_41FA | 字:偶数页页脚_41F9" mode="preprocessing">
    <xsl:for-each select="字:段落_416B | 字:文字表_416C">
      <xsl:choose>
        <xsl:when test="name(.)='字:段落_416B'">
          <字:段落_416B>
            <xsl:for-each select="node()">
              <xsl:call-template name="nodeloop"/>
            </xsl:for-each>
          </字:段落_416B>
        </xsl:when>
        <xsl:when test="name(.)='字:文字表_416C'">
          <xsl:call-template name="revTable"/>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="sectionheaderfooter">
    <字:分节_416A>
      <xsl:choose>
        <xsl:when test="./字:节属性_421B">
          <!--<xsl:call-template name="sectionPr"/>-->
          <xsl:apply-templates select="./字:节属性_421B" mode="sectionPr"/>
        </xsl:when>
        <xsl:when test="./字:修订开始_421F">
          <xsl:copy-of select="./字:修订开始_421F"/>
        </xsl:when>
        <xsl:when test="./字:修订结束_4223">
          <xsl:copy-of select="./字:修订结束_4223"/>
        </xsl:when>
      </xsl:choose>
    </字:分节_416A>
  </xsl:template>

  <xsl:template match="字:节属性_421B" mode="sectionPr">
    <xsl:if
      test="( (./字:页眉_41F3/字:奇数页页眉_41F4) and (not(./字:页眉_41F3/字:偶数页页眉_41F5))and(not(./字:是否奇偶页页眉页脚不同_41ED)) )
                     or ( (./字:页脚_41F7/字:奇数页页脚_41F8) and (not(./字:页脚_41F7/字:偶数页页脚_41F9))and(not(./字:是否奇偶页页眉页脚不同_41ED)) )">
      <字:节属性_421B>
        <xsl:copy-of select="./*[name(.)!='字:页眉_41F3' and name(.)!='字:页脚_41F7']"/>
        <字:页眉_41F3>
          <xsl:copy-of select="./字:页眉_41F3/字:奇数页页眉_41F4"/>
          <字:偶数页页眉_41F5>
            <xsl:copy-of select="./字:页眉_41F3/字:偶数页页眉_41F5/*"/>
          </字:偶数页页眉_41F5>
          <xsl:copy-of select="./字:页眉_41F3/字:首页页眉_41F6"/>
        </字:页眉_41F3>
        <字:页脚_41F7>
          <xsl:copy-of select="./字:页脚_41F7/字:奇数页页脚_41F8"/>
          <字:偶数页页脚_41F9>
            <xsl:copy-of select="./字:页脚_41F7/字:偶数页页脚_41F9/*"/>
          </字:偶数页页脚_41F9>
          <xsl:copy-of select="./字:页脚_41F7/字:首页页脚_41FA"/>
        </字:页脚_41F7>      
    </字:节属性_421B>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>
