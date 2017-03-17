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
<Author>Li Jingui 2011-02-18</Author>
-->
<xsl:stylesheet version="2.0" 
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:fn="http://www.w3.org/2005/xpath-functions"
  xmlns:xdt="http://www.w3.org/2005/xpath-datatypes"
  xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
  xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
  xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"              
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:对象="http://schemas.uof.org/cn/2009/objects"              
  xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
  xmlns:书签="http://schemas.uof.org/cn/2009/bookmarks"
  xmlns:公式="http://schemas.uof.org/cn/2009/equations"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:pzip="urn:u2o:xmlns:post-processings:special">
	<xsl:import href="metadata.xsl"/>
	<xsl:import href="region.xsl"/>
	<xsl:import href="revisions.xsl"/>
	<xsl:import href="sectPr.xsl"/>
	<xsl:import href="styles.xsl"/>
	<xsl:import href="paragraph.xsl"/>
	<xsl:import href="table.xsl"/>
	<xsl:import href="numbering.xsl"/>
	<xsl:import href="header-footer.xsl"/>
	<xsl:import href="endnote-footnote.xsl"/>
	<xsl:import href="field.xsl"/>
	<xsl:import href="object.xsl"/>
  <xsl:import href="graphics.xsl"/>
  <xsl:import href="hyperlinks.xsl"/>
  <xsl:import href="extend.xsl"/>
  <xsl:import href="bookmark.xsl"/>
	<xsl:param name="outputFile"/>
  <xsl:param name="pPartFrom"/>
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

  <xsl:template match="/">

    <pzip:archive pzip:target="{$outputFile}">

      
      <!-- metaData types -->
      <pzip:entry pzip:target="_meta/meta.xml">
        <xsl:call-template name="metaData"/>
      </pzip:entry>
      
      <pzip:entry pzip:target="uof.xml">
        <uof:UOF_0000 mimetype_0001="vnd.uof.text" language_0002="cn" version_0003="2.0">
        </uof:UOF_0000>      
      </pzip:entry>
         
      <!-- rules types -->
      <pzip:entry pzip:target="rules.xml">
        <规则:公用处理规则_B665>
          <规则:长度单位_B666>pt</规则:长度单位_B666>			
          <规则:文字处理_B66B>
            <规则:文档设置_B600>
              <xsl:apply-templates select="document('word/settings.xml')/w:settings" mode="rule"/>
              <xsl:if test="document('word/settings.xml')/w:settings/w:endnotePr">
                <xsl:apply-templates select="document('word/settings.xml')/w:settings/w:endnotePr" mode="endnotePr-pos"/>
              </xsl:if>
            </规则:文档设置_B600>
			      <xsl:call-template name="revisionSet"/>
          </规则:文字处理_B66B>
          <xsl:call-template name="userSet"/><!--region.xsl-->
          <xsl:call-template name="commentSet"/>
        </规则:公用处理规则_B665>
      </pzip:entry>

      <!--2012-11-22，wudi，OOX到UOF方向的书签实现，start-->
      <pzip:entry pzip:target="bookmarks.xml">
        <书签:书签集_9104 xmlns:书签="http://schemas.uof.org/cn/2009/bookmarks">
          <xsl:apply-templates select="w:document/w:body" mode="bookmark"/>
        </书签:书签集_9104>
      </pzip:entry>
      <!--end-->
   
      <!-- styles types -->
      <pzip:entry pzip:target="styles.xml">
        <xsl:call-template name="styles"/>
      </pzip:entry>

		<!-- graphics types -->
      <!--<xsl:if test="//w:drawing">-->
			  <pzip:entry pzip:target="graphics.xml">
          <xsl:call-template name="graphicDefine"/>
			  </pzip:entry>
      <!--</xsl:if>-->

		<!-- object types -->
      <!--<xsl:if test="//wps:spPr/a:blipFill | //pic:blipFill | //w:background | document('word/numbering.xml')/w:numbering/w:numPicBullet/w:pict">-->
			  <pzip:entry pzip:target="objectdata.xml">
          <xsl:call-template name="objectData"/>       
			  </pzip:entry>
      <!--</xsl:if>-->

      <!--2015-06-02，wudi，增加公式的转换，start-->
      <!--euqations.xml-->
      <pzip:entry pzip:target="equations.xml">
        <公式:公式集_C200>
          <xsl:for-each select="//m:oMath">
            <xsl:variable name="number">
              <xsl:number format="1" level="any" count="m:oMath | m:oMathPara"/>
            </xsl:variable>
            <公式:数学公式_C201>
              <xsl:attribute name="标识符_C202">
                <xsl:value-of select="concat('equation_',$number)"/>
              </xsl:attribute>
              <xsl:copy-of select="."/>
            </公式:数学公式_C201>
          </xsl:for-each>
        </公式:公式集_C200>
      </pzip:entry>
      <!--end-->

		<!-- extend types -->
		<pzip:entry pzip:target="extend.xml">
      <扩展:扩展区_B200>
				<xsl:if test="//w:drawing">
					<xsl:call-template name="shdExtend"/>
				</xsl:if>
        <xsl:if test="//w:framePr">
          <xsl:call-template name="framePrExtend"/>
        </xsl:if>
        <xsl:if test="//w:document/w:customXML">
          <xsl:copy-of select="//w:document/w:customXML/*"/>
        </xsl:if>
			</扩展:扩展区_B200>	
  </pzip:entry>
		
      <!--cxl2011/11/12日增加超级链接的转换，仍然存在问题复杂域***************************-->
      <!--2013-03-27，wudi，注释掉if条件语句-->
      <!--<xsl:if test="//w:hyperlink or contains(//w:p//w:instrText,'HYPERLINK')">-->
        <pzip:entry pzip:target="hyperlinks.xml">
          <超链:链接集_AA0B>
            <xsl:call-template name="hyperlinks">
              <xsl:with-param name ="link" select ="'document'"/>
            </xsl:call-template>
          </超链:链接集_AA0B>
        </pzip:entry>
      <!--</xsl:if>-->
      <!--cxl2011/11/29日增加书签转换-->
      <!--<xsl:if test="//w:bookmarkStart">
        <pzip:entry pzip:target="bookmarks">
          <书签:书签集_9104>
            <xsl:call-template name="bookmark"/>
          </书签:书签集_9104>
        </pzip:entry>
      </xsl:if>-->
      <xsl:if test="//w:media">
        <xsl:copy-of select="//w:media/*"/>
      </xsl:if>

      <pzip:entry pzip:target="content.xml">
        <字:文字处理文档_4225 xmlns:uof="http://schemas.uof.org/cn/2009/uof">
          <xsl:apply-templates select="w:document/w:body" mode="main"/>
        </字:文字处理文档_4225>
      </pzip:entry>
      
    </pzip:archive>
    
  </xsl:template>

	<xsl:template match="w:body" mode="main">
		<xsl:for-each select="node()">
			<xsl:choose>
				<!--转换分节-->
				<xsl:when test="name(.)='w:sectPr'">
					<xsl:call-template name="section"/>
				</xsl:when>
				<!--转换段落-->
        
        <!--2013-03-08，wudi，修复分节符转换后出现多余换行BUG，start-->
        <!--2013-03-15，wudi，增加限制条件not(./w:r)，之前的条件会导致分栏转换出问题-->
				<xsl:when test="name(.)='w:p' and not(name(following-sibling::*[1])='w:sectPr' and not(./w:r))">         
          <!--<xsl:if test="./w:pPr/w:sectPr">
            <xsl:call-template name="section"/>
          </xsl:if>-->
					<xsl:call-template name="paragraph">
						<xsl:with-param name ="pPartFrom" select ="'document'"/>
					</xsl:call-template>
				</xsl:when>
        <!--end-->
        
        <!--2012-11-27，wudi，OOX到UOF方向的书签实现，start-->
        <!--暂时先这样，UOF和OOX处理不一样-->
        <!-- 
        <xsl:when test="name(.)='w:bookmarkStart'">
          <字:段落_416B>
            <xsl:call-template name="bookmarkStart"/>
          </字:段落_416B>
        </xsl:when>
        <xsl:when test="name(.)='w:bookmarkEnd'">
          <字:段落_416B>
            <xsl:call-template name="bookmarkEnd"/>
          </字:段落_416B>
        </xsl:when>
        -->
        <!--
        <xsl:when test="name(.)='w:bookmarkStart'">
            <xsl:call-template name="bookmarkStart"/>
        </xsl:when>
        <xsl:when test="name(.)='w:bookmarkEnd'">
            <xsl:call-template name="bookmarkEnd"/>
        </xsl:when>
        -->
        <!--end-->
        
				<!--转换表格-->
				<xsl:when test="name(.)='w:tbl'">
					<xsl:call-template name="table"><!--位于table.xsl中-->
						<xsl:with-param name ="tblPartFrom" select ="'document'"/>
					</xsl:call-template>
				</xsl:when>

        <!--
				<xsl:when test="name(.)='w:sdt'">
					<xsl:call-template name="sdtContentBlock"/>
				</xsl:when>
        -->
        
        <!--2012-12-03，wudi，OOX到UOF方向的目录与索引实现，start-->
        <xsl:when test="name(.)='w:sdt'">
          <xsl:for-each select="w:sdtContent/w:p">
            <!--<xsl:choose>
              <xsl:when test="name(.)='w:sdtContent'">
                <xsl:for-each select="w:p">-->

                  <!--2014-04-08，wudi，目录区域，增加批注的转换，start-->
                  <xsl:variable name="commentStartNum">
                    <xsl:value-of select="count(preceding-sibling::w:commentRangeStart)"/>
                  </xsl:variable>
                  <xsl:variable name="commentStartNum1">
                    <xsl:value-of select="count(ancestor::w:tc/preceding-sibling::w:commentRangeStart)"/>
                  </xsl:variable>
                  <xsl:variable name="commentStartNum2">
                    <xsl:value-of select="count(ancestor::w:sdt/preceding-sibling::w:commentRangeStart)"/>
                  </xsl:variable>
                  <!--end-->
                  
                  <字:段落_416B>

                    <!--2014-04-08，wudi，目录区域，增加批注的转换，start-->
                    <xsl:if test ="($commentStartNum2 > 0) and (name(ancestor::w:sdt/preceding-sibling::*[1]) = 'w:commentRangeStart') and not(preceding-sibling::w:p)">
                      <字:句_419D>
                        <字:区域开始_4165>
                          <xsl:attribute name="标识符_4100">
                            <xsl:value-of select= "concat('cmt_',ancestor::w:sdt/preceding-sibling::w:commentRangeStart[1]/@w:id)"/>
                          </xsl:attribute>
                          <xsl:attribute name="名称_4166">
                            <xsl:value-of select="'annotation'"/>
                          </xsl:attribute>
                          <xsl:attribute name="类型_413B">
                            <xsl:value-of select="'annotation'"/>
                          </xsl:attribute>
                        </字:区域开始_4165>
                      </字:句_419D>
                    </xsl:if>

                    <xsl:if test ="($commentStartNum1 > 0) and (name(ancestor::w:tc/preceding-sibling::*[1]) = 'w:commentRangeStart')">
                      <字:句_419D>
                        <字:区域开始_4165>
                          <xsl:attribute name="标识符_4100">
                            <xsl:value-of select= "concat('cmt_',ancestor::w:tc/preceding-sibling::w:commentRangeStart[1]/@w:id)"/>
                          </xsl:attribute>
                          <xsl:attribute name="名称_4166">
                            <xsl:value-of select="'annotation'"/>
                          </xsl:attribute>
                          <xsl:attribute name="类型_413B">
                            <xsl:value-of select="'annotation'"/>
                          </xsl:attribute>
                        </字:区域开始_4165>
                      </字:句_419D>
                    </xsl:if>

                    <xsl:if test ="($commentStartNum > 0) and (name(preceding-sibling::*[2]) = 'w:commentRangeStart' and name(preceding-sibling::*[1]) = 'w:commentRangeStart')">
                      <字:句_419D>
                        <字:区域开始_4165>
                          <xsl:attribute name="标识符_4100">
                            <xsl:value-of select= "concat('cmt_',preceding-sibling::w:commentRangeStart[2]/@w:id)"/>
                          </xsl:attribute>
                          <xsl:attribute name="名称_4166">
                            <xsl:value-of select="'annotation'"/>
                          </xsl:attribute>
                          <xsl:attribute name="类型_413B">
                            <xsl:value-of select="'annotation'"/>
                          </xsl:attribute>
                        </字:区域开始_4165>
                      </字:句_419D>
                    </xsl:if>

                    <xsl:if test ="($commentStartNum > 0) and (name(preceding-sibling::*[1]) = 'w:commentRangeStart')">
                      <字:句_419D>
                        <字:区域开始_4165>
                          <xsl:attribute name="标识符_4100">
                            <xsl:value-of select= "concat('cmt_',preceding-sibling::w:commentRangeStart[1]/@w:id)"/>
                          </xsl:attribute>
                          <xsl:attribute name="名称_4166">
                            <xsl:value-of select="'annotation'"/>
                          </xsl:attribute>
                          <xsl:attribute name="类型_413B">
                            <xsl:value-of select="'annotation'"/>
                          </xsl:attribute>
                        </字:区域开始_4165>
                      </字:句_419D>
                    </xsl:if>
                    <!--end-->

                    <xsl:for-each select="node()">
                      <xsl:choose>
                        <xsl:when test="name(.)='w:pPr'">
                          <字:段落属性_419B>
                            <!--
                            <xsl:for-each select="node()">
                              <xsl:choose>
                                <xsl:when test="name(.)='w:pStyle'">
                                  <xsl:apply-templates select="w:pStyle" mode="pPrChildren"/>
                                </xsl:when>
                                <xsl:when test="name(.)='w:tabs'">
                                  <xsl:apply-templates select="w:tabs" mode="pPrChildren"/>
                                </xsl:when>
                                <xsl:when test="name(.)='w:rPr'">
                                  <字:句属性_4158>
                                    <xsl:apply-templates select="." mode="RunProperties"/>
                                  </字:句属性_4158>
                                </xsl:when>
                              </xsl:choose>
                            </xsl:for-each>
                            -->

                            <!--2014-05-06，wudi，修复目录区域段落式样转换BUG，start-->
                            <xsl:if test="not(./w:pStyle)">
                              <xsl:call-template name="ParagraphStyle"/>
                            </xsl:if>

                            <xsl:if test="not(./w:pPrChange)">
                              <xsl:call-template name="pPr"/>
                            </xsl:if>
                            <!--end-->
                            
                          </字:段落属性_419B>
                        </xsl:when>
                        <xsl:when test="name(.)='w:r'">
                          <xsl:if test="./w:instrText[@xml:space='preserve'] and contains(./w:instrText,'TOC')">
                            <字:域开始_419E 类型_416E="toc" 是否锁定_416F="false"/>
                            <字:域代码_419F>
                              <字:段落_416B>
                                <字:句_419D>
                                  <字:句属性_4158/>
                                  <字:文本串_415B>
                                    <xsl:value-of select="./w:instrText"/>
                                  </字:文本串_415B>
                                </字:句_419D>
                              </字:段落_416B>
                            </字:域代码_419F>
                          </xsl:if>
                          
                          <!--2013-04-25，wudi，修复目录与索引转换BUG，增加对超链接以./w:fldChar/@w:fldCharType='begin'开始的情况处理，start-->
                          <xsl:if test ="./w:fldChar/@w:fldCharType='begin' and contains(following-sibling::w:r[1]/w:instrText,'HYPERLINK')">
                            <字:句_419D>
                              <xsl:call-template name ="hlkstart"/>
                            </字:句_419D>
                          </xsl:if>
                          <xsl:if test ="./w:fldChar/@w:fldCharType='begin' and contains(following-sibling::w:r[1]/w:instrText,'PAGEREF')">
                            <字:域开始_419E 类型_416E="pageref" 是否锁定_416F="false"/>
                            <字:域代码_419F>
                              <字:段落_416B>
                                <字:句_419D>
                                  <字:句属性_4158/>
                                  <字:文本串_415B>
                                    <xsl:value-of select ="following-sibling::w:r[1]/w:instrText"/>
                                  </字:文本串_415B>
                                </字:句_419D>
                              </字:段落_416B>
                            </字:域代码_419F>
                          </xsl:if>
                          <!--end-->
                          
                          <xsl:if test="contains(./w:instrText,'XE')">
                            
                            <!--2013-04-10，wudi，修复索引转换的BUG，start-->
                            <xsl:call-template name ="SimpleField">
                              <xsl:with-param name ="type" select ="'xe'"/>
                            </xsl:call-template>
                            <!--end-->
                            
                            <!--2013-04-10，wudi，修复索引转换的BUG，注释掉-->
                            <!--<字:域开始_419E 类型_416E="xe" 是否锁定_416F="false"/>
                            <字:域代码_419F>
                              <字:段落_416B>
                                <字:句_419D>
                                  <字:句属性_4158>
                                    <字:字体_4128 颜色_412F="auto"/>
                                    <字:删除线_4135>none</字:删除线_4135>
                                    <字:下划线_4136 线型_4137="none"/>
                                    <字:是否隐藏文字_413D>false</字:是否隐藏文字_413D>
                                  </字:句属性_4158>
                                  <字:文本串_415B>
                                    <xsl:value-of select="./w:instrText"/>
                                    <xsl:value-of select="following::w:instrText[1]"/>
                                    <xsl:value-of select="following::w:instrText[2]"/>
                                    <xsl:value-of select="following::w:instrText[3]"/>
                                  </字:文本串_415B>
                                </字:句_419D>
                              </字:段落_416B>
                            </字:域代码_419F>
                            <字:句_419D>
                              <字:局属性_4158/>
                              <字:文本串_415B>
                                <xsl:value-of select="following::w:instrText[1]"/>
                                <xsl:value-of select="following::w:instrText[2]"/>
                              </字:文本串_415B>
                            </字:句_419D>-->
                          </xsl:if>
                          
                          <!--2013-04-25，wudi，增加制表位为w:tab节点的情况-->
                          <xsl:if test="./w:ptab or ./w:tab">
                            
                            <!--2013-04-10，wudi，修复目录-制表位互操作转换的BUG，注释掉-->
                            <!--<字:段落属性_419B>
                              <字:制表位设置_418F>
                                <字:制表位_4171>
                                  <xsl:attribute name="位置_4172">
                                    <xsl:value-of select="415.34"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="类型_4173">
                                    <xsl:value-of select="./w:ptab/@w:alignment"/>
                                  </xsl:attribute>
                                  <xsl:attribute name="前导符_4174">
                                    <xsl:choose>
                                      <xsl:when test="./w:ptab/@w:leader='dot'">
                                        <xsl:value-of select="'.'"/>
                                      </xsl:when>
                                    </xsl:choose>
                                  </xsl:attribute>
                                </字:制表位_4171>
                              </字:制表位设置_418F>
                            </字:段落属性_419B>-->
                            <字:句_419D>
                              <字:句属性_4158/>
                              <字:制表符_415E/>
                            </字:句_419D>
                          </xsl:if>
                          <xsl:if test="not(./w:fldChar) and not(./w:instrText) and not(./w:ptab) and not(./w:tab)">
                            <字:句_419D>
                              <xsl:for-each select="node()">
                                <xsl:choose>
                                  <xsl:when test="name(.)='w:rPr'">
                                    <字:句属性_4158>
                                      <xsl:apply-templates select="." mode="RunProperties"/>
                                    </字:句属性_4158>
                                  </xsl:when>
                                  <xsl:when test="name(.)='w:t'">
                                    <字:文本串_415B>
                                      <xsl:value-of select="."/>
                                    </字:文本串_415B>
                                  </xsl:when>

                                  <!--2014-04-08，wudi，增加换行符的转换，start-->
                                  <xsl:when test="name(.)='w:br'">
                                    <xsl:if test="not(@w:type) or (@w:type='textWrapping')">
                                      <字:换行符_415F/>
                                    </xsl:if>
                                    <xsl:if test="@w:type='column'">
                                      <字:分栏符_4160/>
                                    </xsl:if>
                                    <xsl:if test="@w:type='page'">
                                      <字:分页符_4163/>
                                    </xsl:if>
                                  </xsl:when>
                                  <!--end-->

                                  <!--2014-03-28，wudi，增加图片的转换，之前没有遇到过这种情况，start-->
                                  <xsl:when test="name(.)='w:drawing'">

                                    <xsl:for-each select="wp:inline | wp:anchor">
                                      <xsl:choose>
                                        <xsl:when test="name(.)='wp:inline'">

                                          <!--2013-04-11，针对包含MACROBUTTON的情况，部分图片代码有问题，避免graphics.xml里再出现锚点引用-->
                                          <!--2014-03-18，wudi，增加文本框里插入图片的情形，去除条件not($rPartFrom ='txbody')，start-->
                                          <xsl:if test="./a:graphic/a:graphicData/pic:pic">
                                            <xsl:call-template name="bodyAnchorPic">
                                              <xsl:with-param name="picType" select="'inline'"/>
                                              <xsl:with-param name="filename" select="'document'"/>
                                            </xsl:call-template>
                                          </xsl:if>
                                          <!--end-->

                                          <!--2013-01-07，wudi，OOX到UOF方向SmartArt实现，start-->
                                          <xsl:if test="./a:graphic/a:graphicData/wps:wsp or ./a:graphic/a:graphicData/wpg:wgp or ./a:graphic/a:graphicData/dgm:relIds">
                                            <xsl:call-template name="bodyAnchorWps">
                                              <xsl:with-param name="picType" select="'inline'"/>
                                              <xsl:with-param name="filename" select="'document'"/>
                                            </xsl:call-template>
                                          </xsl:if>
                                          <!--end-->

                                          <!--2013-04-19，wudi，增加对chart的转换，start-->
                                          <xsl:if test ="./a:graphic/a:graphicData/c:chart">
                                            <xsl:call-template name="bodyAnchorChart">
                                              <xsl:with-param name="picType" select="'inline'"/>
                                              <xsl:with-param name="filename" select="'document'"/>
                                            </xsl:call-template>
                                          </xsl:if>
                                          <!--end-->

                                        </xsl:when>
                                        <xsl:when test="name(.)='wp:anchor'">

                                          <!--2013-04-11，针对包含MACROBUTTON的情况，部分图片代码有问题，避免graphics.xml里再出现锚点引用-->
                                          <xsl:if test="./a:graphic/a:graphicData/pic:pic">
                                            <xsl:call-template name="bodyAnchorPic">
                                              <xsl:with-param name="picType" select="'anchor'"/>
                                              <xsl:with-param name="filename" select="'document'"/>
                                            </xsl:call-template>
                                          </xsl:if>

                                          <!--2013-01-08，wudi，OOX到UOF方向SmartArt实现，start-->
                                          <xsl:if test="./a:graphic/a:graphicData/wps:wsp or ./a:graphic/a:graphicData/wpg:wgp or ./a:graphic/a:graphicData/dgm:relIds">
                                            <xsl:call-template name="bodyAnchorWps">
                                              <xsl:with-param name="picType" select="'anchor'"/>
                                              <xsl:with-param name="filename" select="'document'"/>
                                            </xsl:call-template>
                                          </xsl:if>
                                          <!--end-->

                                        </xsl:when>
                                      </xsl:choose>
                                    </xsl:for-each>
                                  </xsl:when>
                                  <!--end-->
                                  
                                </xsl:choose>
                              </xsl:for-each>
                            </字:句_419D>
                          </xsl:if>
                          <!--end-->

                          <!--2013-03-20，wudi，修复#2735集成测试OO到UOF2.0目录导航功能失效的BUG，start-->
                          <xsl:if test ="./w:fldChar/@w:fldCharType='end'">
                            <字:域结束_41A0/>
                          </xsl:if>
                          <!--end-->
                          
                        </xsl:when>

                        <!--2014-04-08，wudi，目录区域，增加批注的转换，start-->
                        <xsl:when test="name(.)='w:commentRangeStart'">
                          <xsl:call-template name="commentStart"/>

                          <!--2013-03-12，wudi，修复文字表批注BUG，start-->
                          <xsl:variable name ="commentId">
                            <xsl:value-of select ="@w:id"/>
                          </xsl:variable>
                          <xsl:if test ="not(following::w:commentRangeEnd[@w:id=$commentId])">
                            <字:句_419D>
                              <字:区域结束_4167>
                                <xsl:attribute name ="标识符引用_4168">
                                  <xsl:value-of select ="concat('cmt_',$commentId)"/>
                                </xsl:attribute>
                              </字:区域结束_4167>
                            </字:句_419D>
                          </xsl:if>
                          <!--end-->

                        </xsl:when>

                        <xsl:when test="name(.)='w:commentRangeEnd'">

                          <!--2013-03-13，wudi，修复文字表批注BUG，start-->
                          <xsl:variable name ="commentId">
                            <xsl:value-of select ="@w:id"/>
                          </xsl:variable>
                          <xsl:if test ="not(ancestor::w:tr)">
                            <xsl:call-template name="commentEnd"/>
                          </xsl:if>
                          <!--end-->

                        </xsl:when>
                        <!--end-->
                        
                        <!--2012-12-21，wudi，OOX到UOF方向的目录与索引的实现，start-->
                        <xsl:when test ="name(.)='w:sdt'">
                          <xsl:for-each select ="descendant::w:r">
                            <字:句_419D>
                              <xsl:for-each select="node()">
                                <xsl:choose>
                                  <xsl:when test="name(.)='w:rPr'">
                                    <字:句属性_4158>
                                      <xsl:apply-templates select="." mode="RunProperties"/>
                                    </字:句属性_4158>
                                  </xsl:when>
                                  <xsl:when test="name(.)='w:t'">
                                    <字:文本串_415B>
                                      <xsl:value-of select="."/>
                                    </字:文本串_415B>
                                  </xsl:when>
                                </xsl:choose>
                              </xsl:for-each>
                            </字:句_419D>
                          </xsl:for-each>
                        </xsl:when>
                        
                        <!--end-->
                        
                        <xsl:when test="name(.)='w:hyperlink'">
                          <xsl:call-template name="hyperlinkRegion">
                            <xsl:with-param name="filename" select="$pPartFrom"/>
                          </xsl:call-template>
                        </xsl:when>
                      </xsl:choose>
                    </xsl:for-each>
                  </字:段落_416B>
                <!--</xsl:for-each>
              </xsl:when>
            </xsl:choose>-->
          </xsl:for-each>
        </xsl:when>
        <!--end-->
        
				<xsl:when test="name(.)='w:customXml'">
					<xsl:call-template name="CustomXmlBlock"/>
				</xsl:when>
				<xsl:when test="name(.)='w:commentRangeEnd'">
					<xsl:call-template name="regionEnd">
						<xsl:with-param name="id" select="concat('cmt_',@w:id)"/>
					</xsl:call-template>
				</xsl:when>
				<!--<xsl:when test="name(.)='w:bookmarkEnd'">
					<xsl:call-template name="bookmarkEnd"/>
				</xsl:when>-->
			</xsl:choose>
		</xsl:for-each>
	</xsl:template>
  
  <!--2012-11-23，wudi，OOX到UOF方向的书签实现，start-->
  <xsl:template match="w:body" mode="bookmark">
    <xsl:for-each select=".//w:bookmarkStart">
      <!--<xsl:choose>
        --><!--转换书签--><!--
        <xsl:when test="name(.)='w:p'">
          <xsl:call-template name="bookmarks"/>
        </xsl:when>
        <xsl:when test="name(.)='w:bookmarkStart'">-->
          <书签:书签_9105>
            <xsl:attribute name="名称_9103">
              <xsl:value-of select="@w:name"/>
            </xsl:attribute>
            <书签:区域_9100>
              <xsl:attribute name="区域引用_41CE">
                <xsl:value-of select="concat('bk_',@w:id)"/>
              </xsl:attribute>
            </书签:区域_9100>
          </书签:书签_9105>
        <!--</xsl:when>
      </xsl:choose>-->
    </xsl:for-each>
  </xsl:template>
  <!--end-->

  <!--公用处理规则中文档设置模板，上一期还有一些元素未转换，同时标准里新增一些元素-->
	<xsl:template match="w:settings" mode="rule">
    <xsl:if test="./w:view/@w:val='outline'"><!--cxl,2012.3.22视图转换-->
      <规则:当前视图_B601>outline</规则:当前视图_B601>
    </xsl:if>
    <xsl:if test="./w:view/@w:val='web'">
      <规则:当前视图_B601>web</规则:当前视图_B601>
    </xsl:if>
    <xsl:if test="./w:view/@w:val='normal'">
      <规则:当前视图_B601>normal</规则:当前视图_B601>
    </xsl:if>
    <xsl:if test="not(./w:view) or ./w:view/@w:val='masterPages' or ./w:view/@w:val='print' or ./w:view/@w:val='none'">
      <规则:当前视图_B601>page</规则:当前视图_B601>
    </xsl:if>
    
		<xsl:if test="w:zoom">
			<xsl:if test="w:zoom[@w:percent]">
        <规则:缩放_B603>
					<xsl:value-of select="w:zoom/@w:percent"/>
				</规则:缩放_B603>
			</xsl:if>
			<xsl:if test="not(w:zoom/@w:percent)">
        <规则:缩放_B603>100</规则:缩放_B603>
			</xsl:if>
		</xsl:if>
		<xsl:if test="w:defaultTabStop">
      <规则:默认制表位位置_B604>
				<xsl:value-of select="w:defaultTabStop/@w:val div 20"/>
			</规则:默认制表位位置_B604>
		</xsl:if>
		<xsl:if test="w:trackRevisions">
			<xsl:if test="w:trackRevisions/@w:val='false' or w:trackRevisions/@w:val='off' or w:trackRevisions/@w:val='0'">
        <规则:是否修订_B605>
          <xsl:value-of select="'false'"/>
        </规则:是否修订_B605>
			</xsl:if>
			<xsl:if test="not(w:trackRevisions/@w:val='false' or w:trackRevisions/@w:val='off' or w:trackRevisions/@w:val='0')">
        <规则:是否修订_B605>
          <xsl:value-of select="'true'"/>
        </规则:是否修订_B605>
			</xsl:if>
		</xsl:if>
		<xsl:if test="w:noLineBreaksBefore[@w:lang='zh-CN'] | w:noLineBreaksAfter[@w:lang='zh-CN']">
      <规则:标点禁则_B608>
				<xsl:if test="w:noLineBreaksBefore[@w:lang='zh-CN']">
          <规则:行首字符_B609>
						<xsl:value-of select="w:noLineBreaksBefore/@w:val"/>
					</规则:行首字符_B609>
				</xsl:if>
				<xsl:if test="w:noLineBreaksAfter[@w:lang='zh-CN']">
          <规则:行尾字符_B60A>
						<xsl:value-of select="w:noLineBreaksAfter/@w:val"/>
					</规则:行尾字符_B60A>
				</xsl:if>
			</规则:标点禁则_B608>
		</xsl:if>
    <!--cxl,2012.3.5新增字距调整转换-->
    <规则:字距调整_B606>
      <xsl:if test="w:noPunctuationKerning">
        <xsl:value-of select="'western-only'"/>
      </xsl:if>
      <xsl:if test="not(w:noPunctuationKerning)">
        <xsl:value-of select="'western-and-punctuation'"/>
      </xsl:if>
    </规则:字距调整_B606>
	</xsl:template>

  <!--尾注转化模板-->
	<xsl:template match="w:endnotePr" mode="endnotePr-pos">
    <规则:尾注位置_B607>
				<xsl:choose>
					<xsl:when test="./w:pos/@w:val='sectEnd'">
						<xsl:value-of select="'section-end'"/>
					</xsl:when>
					<xsl:when test="./w:pos/@w:val='docEnd'">
						<xsl:value-of select="'doc-end'"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="'doc-end'"/>
					</xsl:otherwise>
				</xsl:choose>
		</规则:尾注位置_B607>
	</xsl:template>

  <xsl:template name ="graphicDefine">  
    <图形:图形集_7C00>

      <!--2013-04-08，wudi，修复图片转换的BUG，没有考虑节点为w:pict，子节点为v:shape的情况，所有number计数增加对v:shape节点的统计，start-->
      <xsl:if test ="//v:rect | //wp:anchor | //wp:inline | //v:shape">
          <xsl:call-template name="graphic">
            <xsl:with-param name ="objFrom" select ="'document'"/>
          </xsl:call-template>
        </xsl:if>
      <xsl:apply-templates select ="document('word/_rels/document.xml.rels')/rel:Relationships" mode ="graphicDef"/>
     <!--end-->
      
     <!--页面填充的图案填充--><!--
      <xsl:if test ="//v:background/v:fill/@type='pattern'">
        <uof:其他对象 uof:locID="u0036" uof:attrList="标识符 内嵌 公共类型 私有类型" uof:标识符="backgroundpattern006" uof:内嵌="false">
          <xsl:variable name ="BackgroundpatternId">
            <xsl:value-of select ="//v:background/v:fill/@r:id"/>
          </xsl:variable>
          <xsl:variable name ="Backgroundpatternname">
            <xsl:value-of select ="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$BackgroundpatternId]/@Target"/>
          </xsl:variable>
          <xsl:variable name ="Backgroundpatterntype">
            <xsl:value-of select ="substring-after($Backgroundpatternname,'.')"/>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test ="$Backgroundpatterntype='emf'">
              <xsl:attribute name ="uof:私有类型">
                <xsl:value-of select ="'emf'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test ="$Backgroundpatterntype='jpeg'">
              <xsl:attribute name ="uof:私有类型">
                <xsl:value-of select ="'jpeg'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="uof:公共类型">
                <xsl:value-of select ="$Backgroundpatterntype"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <o2upic:picture>
            <xsl:attribute name="target">
              <xsl:value-of select ="concat('word/',$Backgroundpatternname)"/>
            </xsl:attribute>
          </o2upic:picture>
        </uof:其他对象>
      </xsl:if>-->
    </图形:图形集_7C00>
  </xsl:template>

  <xsl:template name ="objectData">
    <对象:对象数据集_D700>

      <!--2013-04-08，wudi，修复图片转换的BUG，没有考虑节点为w:pict，子节点为v:shape的情况，所有number计数增加对v:shape节点的统计，start-->
      
      <!--2103-04-19，wudi，增加对chart的转换，处理c:chart节点，start-->
      
      <!--2014-04-10，wudi，增加对SmartArt里包含图片填充的处理，start-->
      <xsl:if test="//wps:spPr/a:blipFill| //pic:blipFill| //v:background | //v:shape | //c:chart | //dgm:relIds">
        <!--分别对应图片填充（预定义图形）、图片填充（背景）、图案填充、自动编号图片符号-->
        <xsl:call-template name="object">
          <xsl:with-param name ="objFrom" select ="'document'"/>
        </xsl:call-template>
      </xsl:if>
      <xsl:apply-templates select ="document('word/_rels/document.xml.rels')/rel:Relationships" mode ="objDef"/>
      <!--end-->
      
      <!--end-->
      
      <!--end-->
      
    </对象:对象数据集_D700>
  </xsl:template>
	
	<!--<xsl:template match="v:background" mode="pic">
	
		<xsl:variable name="filltype" select="v:background/v:fill/@type"/>
		<xsl:if test="$filltype='frame' or $filltype='tile'">
			<xsl:call-template name="object">
				<xsl:with-param name ="objFrom" select ="'document'"/>
			</xsl:call-template>
		</xsl:if>
	
	</xsl:template>-->
  <xsl:template match ="rel:Relationships" mode ="graphicDef">
    <xsl:for-each select ="rel:Relationship">
      <xsl:choose>
        <xsl:when test ="@Target='comments.xml'">
          <xsl:apply-templates select ="document('word/comments.xml')/w:comments" mode ="graphicDef"/>
        </xsl:when>
        <xsl:when test ="@Target='endnotes.xml'">
          <xsl:apply-templates select ="document('word/endnotes.xml')/w:endnotes" mode ="graphicDef"/>
        </xsl:when>
        <xsl:when test ="@Target='footnotes.xml'">
          <xsl:apply-templates select ="document('word/footnotes.xml')/w:footnotes" mode ="graphicDef"/>
        </xsl:when>
        <!--<xsl:when test ="@Target='numbering.xml'">
          <xsl:apply-templates select ="document('word/numbering.xml')/w:numbering" mode ="objDef"/>
        </xsl:when>-->
        <xsl:when test ="contains(@Target,'header')">
          <xsl:variable name ="hnr" select ="@Target"/>
          <xsl:variable name ="hnn" select ="substring-before($hnr,'.xml')"/>
          <xsl:variable name ="hpath" select ="concat('word/',$hnr)"/>

          <xsl:apply-templates select ="document($hpath)/w:hdr" mode ="graphicDef">
            <xsl:with-param name ="headname" select ="$hnn"/>
          </xsl:apply-templates>
        </xsl:when>
        <xsl:when test ="contains(@Target,'footer')">
          <xsl:variable name ="fnr" select ="@Target"/>
          <xsl:variable name ="fnn" select ="substring-before($fnr,'.xml')"/>
          <xsl:variable name ="fpath" select ="concat('word/',$fnr)"/>

          <xsl:apply-templates select ="document($fpath)/w:ftr" mode ="graphicDef">
            <xsl:with-param name ="footname" select ="$fnn"/>
          </xsl:apply-templates>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <xsl:template match ="rel:Relationships" mode ="objDef">
    <xsl:for-each select ="rel:Relationship">
      <xsl:choose>
        <xsl:when test ="@Target='comments.xml'">
          <xsl:apply-templates select ="document('word/comments.xml')/w:comments" mode ="objDef"/>
        </xsl:when>
        <xsl:when test ="@Target='endnotes.xml'">
          <xsl:apply-templates select ="document('word/endnotes.xml')/w:endnotes" mode ="objDef"/>
        </xsl:when>
        <xsl:when test ="@Target='footnotes.xml'">
          <xsl:apply-templates select ="document('word/footnotes.xml')/w:footnotes" mode ="objDef"/>
        </xsl:when>
        <xsl:when test ="@Target='numbering.xml'">
          <xsl:apply-templates select ="document('word/numbering.xml')/w:numbering" mode ="objDef"/>
        </xsl:when>
        <xsl:when test ="contains(@Target,'header')">
          <xsl:variable name ="hnr" select ="@Target"/>
          <xsl:variable name ="hnn" select ="substring-before($hnr,'.xml')"/>
          <xsl:variable name ="hpath" select ="concat('word/',$hnr)"/>

          <xsl:apply-templates select ="document($hpath)/w:hdr" mode ="objDef">
            <xsl:with-param name ="headname" select ="$hnn"/>
          </xsl:apply-templates>
        </xsl:when>
        <xsl:when test ="contains(@Target,'footer')">
          <xsl:variable name ="fnr" select ="@Target"/>
          <xsl:variable name ="fnn" select ="substring-before($fnr,'.xml')"/>
          <xsl:variable name ="fpath" select ="concat('word/',$fnr)"/>

          <xsl:apply-templates select ="document($fpath)/w:ftr" mode ="objDef">
            <xsl:with-param name ="footname" select ="$fnn"/>
          </xsl:apply-templates>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match ="w:comments" mode ="graphicDef">
    <xsl:call-template name ="graphic">
      <xsl:with-param name ="objFrom" select ="'comments'"/>
    </xsl:call-template>
  </xsl:template>
  <xsl:template match ="w:endnotes" mode ="graphicDef">
    <xsl:call-template name ="graphic">
      <xsl:with-param name ="objFrom" select ="'endnotes'"/>
    </xsl:call-template>
  </xsl:template>
  <xsl:template match ="w:footnotes" mode ="graphicDef">
    <xsl:call-template name ="graphic">
      <xsl:with-param name ="objFrom" select ="'footnotes'"/>
    </xsl:call-template>
  </xsl:template>
  <xsl:template match ="w:hdr" mode ="graphicDef">
    <xsl:param name ="headname"/>
    <xsl:call-template name ="graphic">
      <xsl:with-param name ="objFrom" select ="$headname"/>
    </xsl:call-template>
  </xsl:template>
  <xsl:template match ="w:ftr" mode ="graphicDef">
    <xsl:param name ="footname"/>
    <xsl:call-template name ="graphic">
      <xsl:with-param name ="objFrom" select ="$footname"/>
    </xsl:call-template>
  </xsl:template>
  <xsl:template match ="w:comments" mode ="objDef">
    <xsl:call-template name ="object">
      <xsl:with-param name ="objFrom" select ="'comments'"/>
    </xsl:call-template>
  </xsl:template>
  <xsl:template match ="w:endnotes" mode ="objDef">
    <xsl:call-template name ="object">
      <xsl:with-param name ="objFrom" select ="'endnotes'"/>
    </xsl:call-template>
  </xsl:template>
  <xsl:template match ="w:footnotes" mode ="objDef">
    <xsl:call-template name ="object">
      <xsl:with-param name ="objFrom" select ="'footnotes'"/>
    </xsl:call-template>
  </xsl:template>
  <xsl:template match ="w:numbering" mode ="objDef">
    <xsl:for-each select ="w:numPicBullet">
      <xsl:call-template name ="numberingPicture"/>
    </xsl:for-each>
  </xsl:template>
  <xsl:template match ="w:hdr" mode ="objDef">
    <xsl:param name ="headname"/>
    <xsl:call-template name ="object">
      <xsl:with-param name ="objFrom" select ="$headname"/>
    </xsl:call-template>
  </xsl:template>
  <xsl:template match ="w:ftr" mode ="objDef">
    <xsl:param name ="footname"/>
    <xsl:call-template name ="object">
      <xsl:with-param name ="objFrom" select ="$footname"/>
    </xsl:call-template>
  </xsl:template>
	
  <xsl:template name="hyperlinks">
    <xsl:param name="link"/>
    
    <!--2012-12-10,wudi,OOX到UOF方向的目录与索引的实现，start-->
    <!--<xsl:for-each select="w:document/w:body/descendant::node()">-->
    <xsl:for-each select=".//w:hyperlink | .//w:r[w:fldChar/@w:fldCharType='begin']">
    <!--<xsl:for-each select="w:document/w:body/w:p/descendant::node()"> -->
    <!--end-->
    
      <xsl:choose>
        <xsl:when test="name(.)='w:hyperlink'">
          <xsl:call-template name="hyperlink">
            <xsl:with-param name="filename" select="$link"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="name(.)='w:r'"><!--2012.3.14-->
          <!--2013-03-27，wudi，修复超链接转换-hyperlinks.xml文件出现重复链接信息的BUG，增加限制条件@w:fldCharType='begin'-->
          <!--<xsl:if test="./w:fldChar[@w:fldCharType='begin']">-->
            <xsl:variable name="temp">
              <xsl:for-each select="parent::w:p//w:instrText">
                <xsl:value-of select="concat(.,'')"/>
              </xsl:for-each>
            </xsl:variable>
            <xsl:if test="contains($temp,'HYPERLINK')">
              <xsl:call-template name="hyperlink">
                <xsl:with-param name="filename" select="'field'"/>
              </xsl:call-template>
            </xsl:if>
          <!--</xsl:if>-->
        </xsl:when>
        <!--<xsl:when test="name(.)='w:fldSimple'">
          <xsl:if test="./w:hyperlink">--><!--这种情况我做案例还没有遇到，只遇到由复杂域生成的超链接--><!--
            <xsl:for-each select="./w:hyperlink">
              <xsl:call-template name="hyperlink">
                <xsl:with-param name="filename" select="$link"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>
        </xsl:when>-->
      </xsl:choose> 
    </xsl:for-each>   
  </xsl:template>
  
</xsl:stylesheet>
