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
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
  xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:pic="http://purl.oclc.org/ooxml/drawingml/picture"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <xsl:import href="common.xsl"/>
  <xsl:import href="numbering.xsl"/>
  <xsl:import href="sectPr.xsl"/>
  <xsl:import href="region.xsl"/>
  <xsl:import href="hyperlinks.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
  
  <!-- template for run -->
  <!--cxl,current node:字:句_419D，2012.1.11-->
  <xsl:template name="run">

    <!--2013-04-23，wudi，修复修订转换的BUG，修订内容为空格符，修订类型为"delete"的情况，增加参数revType，start-->
    <xsl:param name ="revType"/>
    <xsl:choose>
      <xsl:when test="字:区域开始_4165 or 字:区域结束_4167">
        <xsl:call-template name="runTempWithRegion"/>
      </xsl:when>
      <xsl:otherwise>
        <w:r>
          <xsl:call-template name="runTempNoRegion">
            <xsl:with-param name ="revType" select ="$revType"/>
          </xsl:call-template>
        </w:r>
      </xsl:otherwise>
    </xsl:choose>
    <!--end-->
    
  </xsl:template>

  <xsl:template name="runTempWithRegion">

    <xsl:variable name="lowercase">
      <xsl:if test="字:句属性_4158/字:醒目字体类型_4141='lowercase'">
        <xsl:value-of select="'true'"/>
      </xsl:if>
    </xsl:variable>
    <!--yx,current node:字:句_419D,2010.3.2-->
    <xsl:for-each select="node()">
      <xsl:choose>
        <xsl:when test="name(.)='字:区域开始_4165'">
          <xsl:choose>
            <xsl:when test="./@类型_413B = 'bookmark'">
              <xsl:call-template name="bkStart"/>
            </xsl:when>
            <xsl:when test="./@类型_413B = 'annotation'">
              <xsl:call-template name="annoStart"/>
            </xsl:when>
            <xsl:when test="./@类型_413B = 'hyperlink'">
              <!--yx,consider more cases below,2010.3.2-->
              <xsl:choose>
                <xsl:when test="not(./ancestor::*/字:段落_416B/字:句_419D/字:区域结束_4167/@标识符引用_4168 = ./@标识符_4100)">
                  
                  <!--2013-04-23，wudi，针对OOX案例中超链接开始节点和结束节点不在同一段落的情况，结束节点转成<字:域结束_41A0/>给互操作带来的问题，start-->
                  <!--2013-04-23，wudi，结束节点应该转成字:区域结束_4167，暂时没有找到方法取开始节点的标识符，此案例比较特殊-->
                  <xsl:choose>
                    <xsl:when test ="not(following::字:段落_416B[1]/字:域结束_41A0)">
                      <xsl:call-template name="hyperLink"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:call-template name="hlkstart"/>
                    </xsl:otherwise>
                  </xsl:choose>
                  <!--end-->
                  
                </xsl:when>
                <xsl:when test="./ancestor::*/字:段落_416B/字:句_419D/字:区域结束_4167/@标识符引用_4168 = ./@标识符_4100">
                  <xsl:call-template name="hlkstart"/>
                </xsl:when>
              </xsl:choose>
              <!--yx,consider more cases above,2010.3.2-->
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="name(.)='字:区域结束_4167'">
          <xsl:variable name="sid" select="@标识符引用_4168"/>
          <xsl:choose>
            <xsl:when test="preceding::字:区域开始_4165[@标识符_4100 = $sid and @类型_413B = 'bookmark']">
              <xsl:call-template name="bkEnd"/>
            </xsl:when>
            <xsl:when test="preceding::字:区域开始_4165[@标识符_4100 = $sid and @类型_413B = 'annotation']">
              <xsl:call-template name="annoEnd"/>
            </xsl:when>
            <!--yx,add one case,the same link in more paragraphs,2010.3.2-->
            <xsl:when test="./ancestor::*/字:段落_416B/字:句_419D/字:区域开始_4165[@标识符_4100 = $sid and @类型_413B = 'hyperlink']">
              <xsl:call-template name="hlkend"/>
            </xsl:when>
          </xsl:choose>
        </xsl:when>


        <xsl:when test="name(.)='字:文本串_415B'">
          <w:r>
            <xsl:apply-templates select="preceding-sibling::字:句属性_4158" mode="preWithRegion"/>
            <xsl:choose>
              <xsl:when test="ancestor::字:修订开始_421F[@类型_4221='delete']">
                <!-- lee -->
                <w:delText xml:space="preserve">
                  <xsl:choose>
                    <xsl:when test="$lowercase='true'">
						          <xsl:value-of select="translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
					          </xsl:when>
					          <xsl:otherwise>
						          <xsl:value-of select="."/>
					          </xsl:otherwise>
                  </xsl:choose>
				        </w:delText>
              </xsl:when>
              <!-- lee -->

              <!--2013-04-18，wudi，修复UOF到OOX方向特殊字符转换BUG，笑脸丢失，部分枚举，start-->
              <xsl:when test ="preceding-sibling::字:句属性_4158/字:字体_4128/@特殊字体引用_412B">
                <xsl:variable name ="tmp">
                  <xsl:value-of select ="preceding-sibling::字:句属性_4158/字:字体_4128/@特殊字体引用_412B"/>
                </xsl:variable>
                <xsl:variable name ="wfont">
                  <xsl:value-of select ="//uof:式样集/式样:字体集_990C/式样:字体声明_990D[@标识符_9902 =$tmp]/@名称_9903"/>
                </xsl:variable>
                <xsl:if test =". =''">
                  <w:sym>
                    <xsl:attribute name ="w:font">
                      <xsl:value-of select ="$wfont"/>
                    </xsl:attribute>
                    <xsl:attribute name ="w:char">
                      <xsl:value-of select ="'F04A'"/>
                    </xsl:attribute>
                  </w:sym>
                </xsl:if>
              </xsl:when>
              <!--end-->
              
              <xsl:otherwise>
                <w:t xml:space="preserve"><xsl:if test="$lowercase='true'"><xsl:value-of select="translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/></xsl:if><xsl:if test="$lowercase != 'true'"><xsl:value-of select="."/></xsl:if></w:t>             
              </xsl:otherwise>
            </xsl:choose>
          </w:r>
        </xsl:when>

        <xsl:when test="name(.)='字:空格符_4161'">
          <w:r>
            <xsl:apply-templates select="preceding-sibling::字:句属性_4158" mode="preWithRegion"/>

            <!--2013-04-10，wudi，修复空格符转换BUG，没有考虑不存在“个数_4162”属性的情况，start-->
            <xsl:choose>
              <xsl:when test ="./@个数_4162">
                <w:t xml:space="preserve"><xsl:call-template name="loop"><xsl:with-param name="Count" select="@个数_4162"/></xsl:call-template></w:t>
              </xsl:when>
              <xsl:when test ="not(./@个数_4162)">
                <w:t xml:space="preserve"><xsl:call-template name="loop"><xsl:with-param name="Count" select="1"/></xsl:call-template></w:t>
              </xsl:when>
            </xsl:choose>
            <!--end-->
            
          </w:r>
        </xsl:when>
        <xsl:when test="name(.)='字:制表符_415E'">
          <w:r>
            <xsl:apply-templates select="preceding-sibling::字:句属性_4158" mode="preWithRegion"/>
            <w:tab/>
          </w:r>
        </xsl:when>
        <xsl:when test="name(.)='字:换行符_415F'">
          <w:r>
            <xsl:apply-templates select="preceding-sibling::字:句属性_4158" mode="preWithRegion"/>
            <w:br>
              <xsl:attribute name="w:type">
                <xsl:value-of select="'textWrapping'"/>
              </xsl:attribute>
            </w:br>
          </w:r>
        </xsl:when>
        <xsl:when test="name(.)='字:分栏符_4160'">
          <w:r>
            <xsl:apply-templates select="preceding-sibling::字:句属性_4158" mode="preWithRegion"/>
            <w:br>
              <xsl:attribute name="w:type">
                <xsl:value-of select="'column'"/>
              </xsl:attribute>
            </w:br>
          </w:r>
        </xsl:when>
        <xsl:when test="name(.)='字:分页符_4163'">
          <w:r>
            <xsl:apply-templates select="preceding-sibling::字:句属性_4158" mode="preWithRegion"/>
            <w:br>
              <xsl:attribute name="w:type">
                <xsl:value-of select="'page'"/>
              </xsl:attribute>
            </w:br>
          </w:r>
        </xsl:when>

        <xsl:when test="name(.)='字:引文符号_4164'">
          <xsl:choose>
            <xsl:when test="ancestor::字:脚注_4159">
              <w:r>
                <xsl:apply-templates select="preceding-sibling::字:句属性_4158" mode="preWithRegion"/>
                <w:footnoteRef/>
              </w:r>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="ancestor::字:尾注_415A">
                <w:r>
                  <xsl:apply-templates select="preceding-sibling::字:句属性_4158" mode="preWithRegion"/>
                  <w:endnoteRef/>
                </w:r>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="name(.)='字:脚注_4159'">
          <xsl:variable name="footnoteNo">
            <xsl:number count="字:脚注_4159" level="any" format="1"/>
          </xsl:variable>
          <w:r>
            <xsl:apply-templates select="preceding-sibling::字:句属性_4158" mode="preWithRegion"/>
            <w:footnoteReference>
              <xsl:if test="./@引文体_4157">
                <xsl:attribute name="w:customMarkFollows">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:attribute name="w:id">
                <xsl:call-template name="addone">
                  <xsl:with-param name="temp" select="$footnoteNo"/>
                </xsl:call-template>
              </xsl:attribute>
            </w:footnoteReference>
            <xsl:if test="./@引文体_4157">
              <w:t><xsl:value-of select="./@引文体_4157"/></w:t>
            </xsl:if>
          </w:r>
        </xsl:when>
        <xsl:when test="name(.)='字:尾注_415A'">
          <xsl:variable name="endnoteNo">
            <xsl:number count="字:尾注_415A" level="any" format="1"/>
          </xsl:variable>
          <w:r>
            <xsl:apply-templates select="preceding-sibling::字:句属性_4158" mode="preWithRegion"/>
            <w:endnoteReference>
              <xsl:if test="./@引文体_4157">
                <xsl:attribute name="w:customMarkFollows">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:attribute name="w:id">
                <xsl:call-template name="addone">
                  <xsl:with-param name="temp" select="$endnoteNo"/>
                </xsl:call-template>
              </xsl:attribute>
            </w:endnoteReference>
            <xsl:if test="./@引文体_4157">
              <w:t><xsl:value-of select="./@引文体_4157"/></w:t>
            </xsl:if>
          </w:r>
        </xsl:when>

        <xsl:when test="name(.)='uof:锚点_C644'">
          <w:r>
            <xsl:variable name="number">
              <xsl:number count="uof:锚点_C644" format="1" level="any"/>
            </xsl:variable>
            <xsl:call-template name="objectPicture">
              <xsl:with-param name="num" select="$number"/>
            </xsl:call-template>
          </w:r>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="runTempNoRegion">
    <xsl:param name ="revType"/>
    <xsl:variable name="lowercase">
      <xsl:if test="字:句属性_4158/字:醒目字体类型_4141='lowercase'">
        <xsl:value-of select="'true'"/>
      </xsl:if>
    </xsl:variable>
    <!--cxl,current node:字:句_419D，2012.1.11-->
    <xsl:for-each select="node()">
      <xsl:choose>
        <xsl:when test="name(.)='字:句属性_4158'">
          <xsl:call-template name="RunPrWithChange"/>
        </xsl:when>
        <xsl:when test="name(.)='字:文本串_415B'">
          <xsl:choose>
            <xsl:when test="ancestor::字:修订开始_421F[1][@类型_4221='delete']">
              <w:delText xml:space="preserve"><xsl:if test="$lowercase='true'"><xsl:value-of select="translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/></xsl:if><xsl:if test="$lowercase != 'true'"><xsl:value-of select="."/></xsl:if></w:delText>
            </xsl:when>
            
            <!--2013-04-18，wudi，修复UOF到OOX方向特殊字符转换BUG，笑脸丢失，部分枚举，start-->
            <xsl:when test ="preceding-sibling::字:句属性_4158/字:字体_4128/@特殊字体引用_412B">
              <xsl:variable name ="tmp">
                <xsl:value-of select ="preceding-sibling::字:句属性_4158/字:字体_4128/@特殊字体引用_412B"/>
              </xsl:variable>
              <xsl:variable name ="wfont">
                <xsl:value-of select ="//uof:式样集/式样:字体集_990C/式样:字体声明_990D[@标识符_9902 =$tmp]/@名称_9903"/>
              </xsl:variable>
              <xsl:if test =". =''">
                <w:sym>
                  <xsl:attribute name ="w:font">
                    <xsl:value-of select ="$wfont"/>
                  </xsl:attribute>
                  <xsl:attribute name ="w:char">
                    <xsl:value-of select ="'F04A'"/>
                  </xsl:attribute>
                </w:sym>
              </xsl:if>
            </xsl:when>
            <!--end-->
            
            <xsl:otherwise>
              <w:t xml:space="preserve"><xsl:if test="$lowercase = 'true'"><xsl:value-of select="translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/></xsl:if><xsl:if test="$lowercase != 'true'"><xsl:value-of select="."/></xsl:if></w:t>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="name(.)='字:空格符_4161'">
          <!--<w:t xml:space="preserve"> </w:t>-->
          <!--cxl2012/2/9空格符的产生这地方有问题(多个空格符的情况)-->

          <!--2013-04-23，wudi，修复修订转换的BUG，修订内容为空格符，修订类型为"delete"的情况，start-->
          <xsl:choose>
            <xsl:when test ="$revType = 'w:del'">
              
              <!--2013-04-10，wudi，修复空格符转换BUG，没有考虑不存在“个数_4162”属性的情况，start-->
              <xsl:choose>
                <xsl:when test ="./@个数_4162">
                  <w:delText xml:space="preserve"><xsl:call-template name="loop"><xsl:with-param name="Count" select="@个数_4162"/></xsl:call-template></w:delText>
                </xsl:when>
                <xsl:otherwise>
                  <w:delText xml:space="preserve"><xsl:call-template name="loop"><xsl:with-param name="Count" select="1"/></xsl:call-template></w:delText>
                </xsl:otherwise>
              </xsl:choose>
              <!--end-->
              
            </xsl:when>
            <xsl:otherwise>

              <!--2013-04-10，wudi，修复空格符转换BUG，没有考虑不存在“个数_4162”属性的情况，start-->
              <xsl:choose>
                <xsl:when test ="./@个数_4162">
                  <w:t xml:space="preserve"><xsl:call-template name="loop"><xsl:with-param name="Count" select="@个数_4162"/></xsl:call-template></w:t>
                </xsl:when>
                <xsl:when test ="not(./@个数_4162)">
                  <w:t xml:space="preserve"><xsl:call-template name="loop"><xsl:with-param name="Count" select="1"/></xsl:call-template></w:t>
                </xsl:when>
              </xsl:choose>
              <!--end-->

            </xsl:otherwise>
          </xsl:choose>
          <!--end-->
          
        </xsl:when>
        <xsl:when test="name(.)='字:制表符_415E'">
          <w:tab/>
        </xsl:when>
        <xsl:when test="name(.)='字:换行符_415F'">
          <w:br>
            <xsl:attribute name="w:type">
              <xsl:value-of select="'textWrapping'"/>
            </xsl:attribute>
          </w:br>
        </xsl:when>
        <xsl:when test="name(.)='字:分栏符_4160'">
          <w:br>
            <xsl:attribute name="w:type">
              <xsl:value-of select="'column'"/>
            </xsl:attribute>
          </w:br>
        </xsl:when>
        <xsl:when test="name(.)='字:分页符_4163'">
          <w:br>
            <xsl:attribute name="w:type">
              <xsl:value-of select="'page'"/>
            </xsl:attribute>
          </w:br>
        </xsl:when>

        <xsl:when test="name(.)='字:引文符号_4164'">
          <xsl:choose>
            <xsl:when test="ancestor::字:脚注_4159">
              <w:footnoteRef/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="ancestor::字:尾注_415A">
                <w:endnoteRef/>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>

        <xsl:when test="name(.)='字:脚注_4159'">
          <xsl:variable name="footnoteNo">
            <xsl:number count="字:脚注_4159" level="any" format="1"/>
          </xsl:variable>
          <w:footnoteReference>
            <xsl:if test="./@引文体_4157">
              <xsl:attribute name="w:customMarkFollows">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:attribute name="w:id">
              <xsl:call-template name="addone">
                <xsl:with-param name="temp" select="$footnoteNo"/>
              </xsl:call-template>
            </xsl:attribute>
          </w:footnoteReference>
          <xsl:if test="./@引文体_4157">
            <w:t><xsl:value-of select="./@引文体_4157"/></w:t>
          </xsl:if>
        </xsl:when>

        <xsl:when test="name(.)='字:尾注_415A'">
          <xsl:variable name="endnoteNo">
            <xsl:number count="字:尾注_415A" level="any" format="1"/>
          </xsl:variable>
          <w:endnoteReference>
            <xsl:if test="./@引文体_4157">
              <xsl:attribute name="w:customMarkFollows">
                <xsl:value-of select="'1'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:attribute name="w:id">
              <xsl:call-template name="addone">
                <xsl:with-param name="temp" select="$endnoteNo"/>
              </xsl:call-template>
            </xsl:attribute>
          </w:endnoteReference>
          <xsl:if test="./@引文体_4157">
            <w:t><xsl:value-of select="./@引文体_4157"/></w:t>
          </xsl:if>
        </xsl:when>

        <xsl:when test="name(.)='uof:锚点_C644'">
          <xsl:variable name="number">
            <xsl:number count="uof:锚点_C644" format="1" level="any"/>
          </xsl:variable>

          <!--yx,add num 锚点的数目，2010.2.6-->
          <xsl:call-template name="objectPicture">
            <xsl:with-param name="num" select="$number"/>
          </xsl:call-template>
        </xsl:when>
        <!--other-->
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>


  <xsl:template match="字:句属性_4158" mode="preWithRegion">
    <xsl:call-template name="RunPrWithChange"/>
  </xsl:template>

  <xsl:template name="lowercaseText">
    <xsl:param name="lowercase"/>
    <xsl:choose>
      <xsl:when test="$lowercase='true'">
        <xsl:value-of
          select="translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="."/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="addone">
    <xsl:param name="temp"/>
    <xsl:value-of select="number($temp)+1"/>
  </xsl:template>


  <xsl:template name="loop">
    <xsl:param name="Count"/>
    <xsl:if test="$Count&gt;=1">
      <xsl:value-of select="' '"/>
      <xsl:call-template name="loop">
        <xsl:with-param name="Count" select="number($Count)-1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <xsl:template name="RunPrWithChange">
    <xsl:choose>
      <xsl:when test="following-sibling::字:修订开始_421F[@类型_4221='format']/字:句属性_4158">
        <xsl:apply-templates select="following-sibling::字:修订开始_421F[@类型_4221='format']/字:句属性_4158" mode="withRev"/>
        <!--<w:rPrChange>
        <xsl:apply-templates select="." mode="rpr"/>
      </w:rPrChange>-->
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="." mode="rpr"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="字:句属性_4158" mode="withRev">
    <xsl:apply-templates select="." mode="rpr"/>
  </xsl:template>


  <!-- run properties,cxl2012.1.12-->
  <xsl:template name="RunProperties" match="字:句属性_4158" mode="rpr">
    <w:rPr>
      <!--双行合一 add by Lee-->
      
      <!--2013-04-17，wudi，修复UOF到OOX方向双行合一的BUG，start-->
      <xsl:choose>
        <xsl:when test="./字:双行合一_4148 and ./字:双行合一_4148/@前置字符_414A = '('">
          <w:eastAsianLayout>
            <xsl:attribute name ="w:id">
              <xsl:value-of select ="./字:双行合一_4148/@标识符_4149"/>
            </xsl:attribute>
            <xsl:attribute name ="w:combine">
              <xsl:value-of select ="'1'"/>
            </xsl:attribute>
            <xsl:attribute name ="w:combineBrackets">
              <xsl:value-of select ="'round'"/>
            </xsl:attribute>
          </w:eastAsianLayout>
        </xsl:when>
        <xsl:when test="./字:双行合一_4148 and (./字:双行合一_4148/@前置字符_414A = '' or not(./字:双行合一_4148/@前置字符_414A))">
          <w:eastAsianLayout>
            <xsl:attribute name ="w:id">
              <xsl:value-of select ="./字:双行合一_4148/@标识符_4149"/>
            </xsl:attribute>
            <xsl:attribute name ="w:combine">
              <xsl:value-of select ="'1'"/>
            </xsl:attribute>
          </w:eastAsianLayout>
        </xsl:when>
      </xsl:choose>
      <!--end-->

      <!-- run style -->
      <xsl:if test="@式样引用_417B">
        <w:rStyle>
          <xsl:attribute name="w:val">
            <xsl:value-of select="@式样引用_417B"/>
          </xsl:attribute>
        </w:rStyle>
      </xsl:if>

      <!--run fonts-->
      <xsl:if test="./字:字体_4128">     
        <xsl:apply-templates select="字:字体_4128"/>
      </xsl:if>

      <!--bold-->
      <xsl:if test="./字:是否粗体_4130">
        <w:b>
          <xsl:if test="./字:是否粗体_4130">
            <xsl:attribute name="w:val">
              <xsl:value-of select="./字:是否粗体_4130"/>
            </xsl:attribute>
          </xsl:if>
        </w:b>
      </xsl:if>
      <!--italics-->
      <xsl:if test="./字:是否斜体_4131">
        <w:i>
          <xsl:if test="./字:是否斜体_4131">
            <xsl:attribute name="w:val">
              <xsl:value-of select="./字:是否斜体_4131"/>
            </xsl:attribute>
          </xsl:if>
        </w:i>
      </xsl:if>

      <!--caps display all characters as capital letters-->
      <xsl:if test="./字:醒目字体类型_4141='uppercase'">
        <w:caps/>
      </xsl:if>
      <!-- small caps-->
      <xsl:if test="./字:醒目字体类型_4141='small-caps'">
        <w:smallCaps/>
      </xsl:if>
      <!-- first letter capital infomation lost-->
      <xsl:if test="./字:醒目字体类型_4141='capital'">
        <xsl:message terminate="no">
          feedback:lost:Capital_Letter_Style_in_Paragraph:first character
          as capital letter
        </xsl:message>
      </xsl:if>

      <!-- strike-->
      <xsl:if test="./字:删除线_4135='single'">
        <w:strike/>
      </xsl:if>
      <!-- &double strike-->
      <xsl:if test="./字:删除线_4135='double'">
        <w:dstrike/>
      </xsl:if>

      <!-- outline -->
      <xsl:if test="./字:是否空心_413E">
        <w:outline>
          <xsl:attribute name="w:val">
            <xsl:value-of select="./字:是否空心_413E"/>
          </xsl:attribute>
        </w:outline>
      </xsl:if>

      <!-- shadow -->
      <!--2015-05-10，wudi，修复项目符合互操作后多出阴影的BUG，start-->
      <xsl:if test="./字:是否阴影_4140">
        <!--<w14:shadow w14:blurRad="50800" w14:dist="114300" w14:dir="2700000"
                        w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl">
			  <w14:srgbClr w14:val="000000">
				  <w14:alpha w14:val="60000"/>
			  </w14:srgbClr>
		  </w14:shadow>-->
        <!--<w14:shadow>
          <xsl:attribute name="w:val">
            <xsl:value-of select="./字:是否阴影_4140"/>
          </xsl:attribute>
        </w14:shadow>-->
      </xsl:if>
      <!--end-->

      <!-- emboss & imprint-->
      <xsl:if test="./字:浮雕_413F='emboss'">
        <w:emboss/>
      </xsl:if>
      <xsl:if test="./字:浮雕_413F='engrave'">
        <w:imprint/>
      </xsl:if>

      <!-- snap to grid-->
      <xsl:if test="./字:是否字符对齐网格_4147">
        <w:snapToGrid>
          <xsl:attribute name="w:val">
            <xsl:value-of select="./字:是否字符对齐网格_4147"/>
          </xsl:attribute>
        </w:snapToGrid>
      </xsl:if>

      <!-- vanish -->
      <xsl:if test="./字:是否隐藏文字_413D">
        <w:vanish>
          <xsl:attribute name="w:val">
            <xsl:value-of select="./字:是否隐藏文字_413D"/>
          </xsl:attribute>
        </w:vanish>
      </xsl:if>

      <!-- run color-->
      <xsl:if test="./字:字体_4128/@颜色_412F">
        <w:color>
          <xsl:attribute name="w:val">
            <xsl:choose>
              <xsl:when test="./字:字体_4128[@颜色_412F='auto']">
                <xsl:value-of select="./字:字体_4128/@颜色_412F"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="substring(./字:字体_4128/@颜色_412F,2)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </w:color>
      </xsl:if>

      <!-- character spacing adjustment-->
      <xsl:if test="./字:字符间距_4145">
        <w:spacing>
          <xsl:attribute name="w:val">
            <xsl:call-template name="twipsMeasure">
              <xsl:with-param name="lengthVal" select="number(./字:字符间距_4145)"/>
            </xsl:call-template>
          </xsl:attribute>
        </w:spacing>
      </xsl:if>

      <!-- w expanded or compressed text-->
      <xsl:if test="./字:缩放_4144">
        <w:w>
          <xsl:attribute name="w:val">
            
            <!--2015-01-23，wudi，取值差异，start-->
            <xsl:value-of select="concat(number(./字:缩放_4144),'%')"/>
            <!--end-->
            
          </xsl:attribute>
        </w:w>
      </xsl:if>

      <!-- kern -->
      <xsl:if test="./字:调整字间距_4146">
        <w:kern>
          <xsl:attribute name="w:val">
            <!--
          <xsl:value-of select="number(./字:调整字间距)* 2"/>
          -->
            <xsl:call-template name="halfPointMeasure">
              <xsl:with-param name="lengthVal" select="number(./字:调整字间距_4146)"/>
            </xsl:call-template>
          </xsl:attribute>
        </w:kern>
      </xsl:if>

      <!-- position-->
      <xsl:if test="./字:位置_4142">
        <w:position>
          <xsl:choose>
            <xsl:when test="contains(./字:位置_4142/字:偏移量_4126,' ')">
              <xsl:attribute name="w:val">
                <xsl:call-template name="halfPointMeasure">
                  <xsl:with-param name="lengthVal" select="number(substring-before(./字:位置_4142/字:偏移量_4126,' '))"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="w:val">
                <xsl:call-template name="halfPointMeasure">
                  <xsl:with-param name="lengthVal" select="number(./字:位置_4142/字:偏移量_4126)"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </w:position>
      </xsl:if>

      <!-- font size-->
      <xsl:if test="./字:字体_4128/@字号_412D">
        <w:sz>
          <xsl:attribute name="w:val">
            <xsl:value-of select="number(./字:字体_4128/@字号_412D)*2"/>
          </xsl:attribute>
        </w:sz>
      </xsl:if>

      <!-- highlight-->
      <xsl:if test="./字:突出显示颜色_4132">
        <xsl:apply-templates select="字:突出显示颜色_4132"/>
      </xsl:if>

      <!-- underline-->
      <xsl:if test="./字:下划线_4136">
        <xsl:apply-templates select="字:下划线_4136"/>
      </xsl:if>

      <!-- text border-->
      <xsl:if test="./字:边框线_4226">
        <xsl:apply-templates select="字:边框线_4226" mode="runBorder"/>
      </xsl:if>

      <!-- subscript or superscript text-->
      <xsl:if test="./字:上下标类型_4143">
        <w:vertAlign>
          <xsl:attribute name="w:val">
            <xsl:choose>
              <xsl:when test="./字:上下标类型_4143='sub'">
                <xsl:value-of select="'subscript'"/>
              </xsl:when>
              <xsl:when test="./字:上下标类型_4143='sup'">
                <xsl:value-of select="'superscript'"/>
              </xsl:when>
              <xsl:when test="./字:上下标类型_4143='none'">
                <xsl:value-of select="'baseline'"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </w:vertAlign>
      </xsl:if>

      <!-- emphasis mark -->
      <xsl:if test="./字:着重号_413A">
        <w:em>
          <xsl:attribute name="w:val">
            <xsl:value-of select="./字:着重号_413A/@类型_413B"/>
          </xsl:attribute>
        </w:em>
      </xsl:if>

      <!-- run shading -->
      <xsl:if test="./字:填充_4134">
        <!--<xsl:apply-templates select="字:填充_4134" mode="runShd"/>-->
        <xsl:apply-templates select="字:填充_4134" mode="shdCommon"/>
      </xsl:if>

		<xsl:if test="../../字:修订开始_421F">
			<w:rPrChange>
				<xsl:variable name="revId">
					<xsl:value-of select="../../字:修订开始_421F/@修订信息引用_4222"/>
				</xsl:variable>
				<xsl:variable name="revDate">
					<xsl:value-of select="preceding::规则:修订信息_B60F[@标识符_B610=$revId]/@日期_B612"/>
				</xsl:variable>
				<xsl:variable name="userId">
					<xsl:value-of select="preceding::规则:修订信息_B60F[@标识符_B610=$revId]/@作者_B611"/>
				</xsl:variable>
				<xsl:variable name="username">
					<xsl:value-of select="preceding::规则:用户_B668[@标识符_4100=$userId]/@姓名_41DC"/>
				</xsl:variable>
				
				<xsl:attribute name="w:id">
					<xsl:apply-templates select="preceding::规则:修订信息_B60F[@标识符_B610=$revId]" mode="rev"/>
				</xsl:attribute>
				<xsl:attribute name="w:author">
					<xsl:value-of select="$username"/>
				</xsl:attribute>
				<xsl:attribute name="w:date">
					<xsl:value-of select="concat($revDate,'Z')"/>
				</xsl:attribute>
				<w:rPr/>
			</w:rPrChange>
		</xsl:if>
    </w:rPr>
  </xsl:template>

  <!--cxl,2012.2.23新增字体转换模板-->
  <!--  run fonts -->
  <xsl:template match="字:字体_4128">
    <w:rFonts>
      <xsl:choose>
        <xsl:when test="@是否西文绘制_412C='false'">
          <xsl:attribute name="w:hint">
            <xsl:value-of select="'eastAsia'"/>
          </xsl:attribute>
        </xsl:when>

        <!--2014-04-16，修复字体转换BUG，处理空格符转换前后显示不一样的问题，start-->
        <xsl:when test="not(@是否西文绘制_412C)">
          <xsl:attribute name="w:hint">
            <xsl:value-of select="'eastAsia'"/>
          </xsl:attribute>
        </xsl:when>
        <!--end-->
        
        <xsl:otherwise>
          <xsl:attribute name="w:hint">
            <xsl:value-of select="'default'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:if test="@西文字体引用_4129">
        <xsl:attribute name="w:ascii">
          <xsl:variable name="temp" select="@西文字体引用_4129"/>
          <xsl:value-of select="/uof:UOF/uof:式样集//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$temp]/@名称_9903"/>
        </xsl:attribute>
        <xsl:attribute name="w:hAnsi">
          <!--for numbering use not sure-->
          <xsl:variable name="temp" select="@西文字体引用_4129"/>
          <xsl:value-of select="/uof:UOF/uof:式样集//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$temp]/@名称_9903"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@中文字体引用_412A">
        <xsl:attribute name="w:eastAsia">
          <xsl:variable name="temp" select="@中文字体引用_412A"/>
          <xsl:value-of select="/uof:UOF/uof:式样集//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$temp]/@名称_9903"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@特殊字体引用_412B">
        <xsl:attribute name="w:cs">
          <xsl:variable name="temp" select="@特殊字体引用_412B"/>
          <xsl:value-of select="/uof:UOF/uof:式样集//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$temp]/@名称_9903"/>
        </xsl:attribute>
      </xsl:if>
    </w:rFonts>
  </xsl:template>

  <!--cxl,2012.2.23新增“字:突出显示颜色_4132”模板-->
  <xsl:template match="字:突出显示颜色_4132">
    <w:highlight>
      <xsl:variable name="color"
        select="translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
      <xsl:attribute name="w:val">

        <xsl:choose>
          <xsl:when test="$color='#ffcc00' or $color='#ffff00' or $color='#ffff99'">
            <xsl:value-of select="'yellow'"/>
          </xsl:when>
          <xsl:when test="$color='#008000' or $color='#99cc00' or $color='#00ff00'">
            <xsl:value-of select="'green'"/>
          </xsl:when>
          <xsl:when test="$color='#00ffff' or $color='#00ccff' or $color='#99ccff'">
            <xsl:value-of select="'cyan'"/>
          </xsl:when>
          <xsl:when test="$color='#ff00ff' or $color='#ff99cc' or $color='#cc99ff'">
            <xsl:value-of select="'magenta'"/>
          </xsl:when>
          <xsl:when test="$color='#333399' or $color='#0000ff' or $color='#3366ff'">
            <xsl:value-of select="'blue'"/>
          </xsl:when>
          <!--10 17 对应red-->
          <xsl:when test="$color='#ff6600' or $color='#ff0000'">
            <xsl:value-of select="'red'"/>
          </xsl:when>
          <xsl:when test="$color='#003366' or $color='#000080'">
            <xsl:value-of select="'darkBlue'"/>
          </xsl:when>
          <!--13 21 对应8-->
          <xsl:when test="$color='#008080' or $color='#33cccc'">
            <xsl:value-of select="'darkCyan'"/>
          </xsl:when>
          <xsl:when test="$color='#003300' or $color='#339966'">
            <xsl:value-of select="'darkGreen'"/>
          </xsl:when>
          <xsl:when test="$color='#800080' or $color='#993366'">
            <xsl:value-of select="'darkMagenta'"/>
          </xsl:when>
          <xsl:when test="$color='#993300' or $color='#800000'">
            <xsl:value-of select="'darkRed'"/>
          </xsl:when>
          <xsl:when
            test="$color='#333300' or $color='#808000' or $color='#ff9900' or $color='#ccffcc'">
            <xsl:value-of select="'darkYellow'"/>
          </xsl:when>
          <xsl:when
            test="$color='#666699' or $color='#808080' or $color='#c0c0c0' or $color='#ffcc99'">
            <xsl:value-of select="'darkGray'"/>
          </xsl:when>
          <xsl:when test="$color='#969696' or $color='#ccffff'">
            <xsl:value-of select="'lightGray'"/>
          </xsl:when>
          <xsl:when test="$color='#000000' or $color='#333333'">
            <xsl:value-of select="'black'"/>
          </xsl:when>
          <xsl:when test="$color='auto'">
            <xsl:value-of select="'yellow'"/>
          </xsl:when>
          <!-- other colors use yellow for substitue-->
          <xsl:otherwise>
            <xsl:value-of select="'yellow'"/>
            <xsl:message terminate="no">
              feedback:lost:Highlight_Style_in_Paragraph:text highlight
            </xsl:message>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </w:highlight>
  </xsl:template>

  <xsl:template name="underline" match="字:下划线_4136">
    <w:u>
      <xsl:variable name="bordertp" select="@线型_4137"/>
      <xsl:variable name="borderxs">
        <xsl:choose>
          <xsl:when test="@虚实_4138">
            <xsl:value-of select="@虚实_4138"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select ="'solid'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:attribute name="w:val">
        <xsl:choose>
          <xsl:when test="$bordertp='none'">none</xsl:when>
          <xsl:when test="$bordertp='single' and $borderxs='solid'">
            <xsl:if test="@是否字下划线_4139='true'">words</xsl:if>
            <xsl:if test="not (@是否字下划线_4139)">single</xsl:if>
          </xsl:when>
          <xsl:when test="$bordertp='double' and $borderxs='solid'">double</xsl:when>
          <xsl:when test="$bordertp='single' and $borderxs='round-dot'">thick</xsl:when>
          <xsl:when test="$bordertp='single' and $borderxs='square-dot'">dotted</xsl:when>
          <xsl:when test="$bordertp='thick-between-thin' and $borderxs='square-dot'"
            >dottedHeavy</xsl:when>
          <xsl:when test="$bordertp='single' and $borderxs='dash'">dash</xsl:when>
          <xsl:when test="$bordertp='thick-between-thin' and $borderxs='dash'"
            >dashedHeavy</xsl:when>
          <xsl:when test="$bordertp='single' and $borderxs='long-dash'">dashLong</xsl:when>
          <xsl:when test="$bordertp='thick-between-thin' and $borderxs='long-dash'"
            >dashLongHeavy</xsl:when>
          <xsl:when test="$bordertp='single' and $borderxs='dash-dot'">dotDash</xsl:when>
          <xsl:when test="$bordertp='thick-between-thin' and $borderxs='dash-dot'"
            >dashDotHeavy</xsl:when>
          <xsl:when test="$bordertp='single' and $borderxs='dash-dot-dot'">dotDotDash</xsl:when>
          <xsl:when test="$bordertp='thick-between-thin' and $borderxs='dash-dot-dot'"
            >dashDotDotHeavy</xsl:when>
          <xsl:when test="$bordertp='single' and $borderxs='wave'">wave</xsl:when>
          <xsl:when test="$bordertp='thick-between-thin' and $borderxs='long-dash-dot'"
            >wavyHeavy</xsl:when>
          <xsl:when test="$bordertp='double' and $borderxs='round-dot'">wavyDouble</xsl:when>
          
          <!--2013-04-25，wudi，修复字体下划线转换BUG，新增两种下划线的转换，start-->
          <xsl:when test ="$bordertp='thick' and $borderxs='long-dash'">dashLongHeavy</xsl:when>
          <xsl:when test ="$bordertp='double' and $borderxs='wave'">wavyDouble</xsl:when>
          <!--end-->
          
          <xsl:otherwise>single</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>

      <xsl:if test="@是否字下划线_4139='true'">
        <xsl:attribute name="w:val">words</xsl:attribute>
      </xsl:if>
      <xsl:if test="@颜色_412F">
        <xsl:attribute name="w:color">
          <xsl:choose>
            <xsl:when test="@颜色_412F='auto'">
              <xsl:value-of select="@颜色_412F"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring(@颜色_412F,2)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
    </w:u>
  </xsl:template>

  <xsl:template match="字:边框线_4226" mode="runBorder">
    <w:bdr>
      <xsl:if test="./@阴影类型_C645='right-bottom'"><!--cxl,2012.3.17-->
        <xsl:attribute name="w:shadow">
          <xsl:value-of select="'1'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:call-template name="SingleBorder"/><!--paragraph.xsl-->
    </w:bdr>
  </xsl:template>

  <!--<xsl:template match="字:填充_4134" mode="runShd">cxl填充，句属性只支持颜色或图案的填充。
    <w:shd>
      <xsl:attribute name="w:val">
        <xsl:value-of select="'solid'"/>
      </xsl:attribute>
      <xsl:if test="./图:颜色_8004">
        <xsl:attribute name="w:color">
          <xsl:choose>
            <xsl:when test="./图:颜色_8004='auto'">
              <xsl:value-of select="./图:颜色_8004"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring(./图:颜色_8004,2)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./图:图案_800A">
        <xsl:attribute name="w:color">
          <xsl:choose>
            <xsl:when test="./图:图案_800A[@图:前景色='auto']">
              <xsl:value-of select="./图:图案_800A/@图:前景色"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring(./图:图案_800A/@图:前景色,2)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="w:fill">
          <xsl:choose>
            <xsl:when test="./图:图案_800A[@图:背景色='auto']">
              <xsl:value-of select="./图:图案_800A/@图:背景色"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring(./图:图案_800A/@图:背景色,2)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
    </w:shd>
  </xsl:template>-->
  
</xsl:stylesheet>