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
<Author>Gu Yueqiong(BUAA)</Author>
<Author>Ban Qianchao(BUAA)</Author>
-->

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:公式="http://schemas.uof.org/cn/2009/equations"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles">

  <xsl:import href="sectPr.xsl"/>
  <!--yx,add,simpleField is in region.xsl2010.2.2-->
  <xsl:import href="region.xsl"/>
  <!--yx,add,"objectPicture" is in object.xsl2010.2.6-->
  <xsl:import href="run.xsl"/>
  <xsl:import href="common.xsl"/>
  <!--yx,add,<xsl:import href="common.xsl"/>,template name="twipsMeasure"is in this xsl2010.2.6-->
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>

  <xsl:preserve-space elements="字:文本串_415B"/>

  <!--template for paragraph -->
  <!--cxl修改段落，2012.1.5-->
  <xsl:template name="paragraph">
    <xsl:message terminate="no">progress:paragraph</xsl:message>
    <w:p>
      <xsl:for-each select="node()">      
        <xsl:choose>

          <!--2014-04-08，wudi，修复批注转换BUG，与下面条件冲突，在此补充，start-->
          <xsl:when test="name(.)='字:句_419D' and ./字:区域开始_4165[@类型_413B='annotation']">
            <xsl:call-template name="runTempWithRegion"/>
          </xsl:when>
          <!--end-->
          
          <!--和字：句并列-->
          <xsl:when test="name(.)='字:段落属性_419B'">
            <xsl:call-template name="pPrWithpStyle"/>
          </xsl:when>
          
          <!--yx,add one special case:header&footer inserted with pages with its information can be found in Object set.2010.2.2-->
          <xsl:when
            test="((name(.)='字:句_419D') and (./uof:锚点_C644/@图形引用_C62E = ./uof:锚点_C644/preceding::uof:对象集/图:图形_8062/@标识符_804B) 
            and (./uof:锚点_C644/preceding::uof:对象集/图:图形_8062[@图形引用_C62E=./uof:锚点_C644/preceding::uof:对象集/图:图形_8062/@标识符_804B]/图:文本_803C/图:内容_8043/字:段落_416B/字:域开始_419E[@类型_416E='page']))">
            <xsl:for-each
              select="./uof:锚点_C644/preceding::uof:对象集/图:图形_8062[1]/图:文本_803C/图:内容_8043/字:段落_416B/字:域开始_419E[@类型_416E='page']  ">
              <!--yx,current node:.../字:域开始,2010.2.2-->
              <xsl:call-template name="pagesinobjectset"/><!--yx,manage regionstart in objectset-->
            </xsl:for-each>
            <!--<xsl:if test="./字:锚点/字:锚点属性/字:位置/字:水平/字:相对/@字:参考点">
               <w:jc>
                  <xsl:attribute name="w:val">
                    <xsl:value-of select="./字:锚点/字:锚点属性/字:位置/字:水平/字:相对/@字:参考点"/>
                  </xsl:attribute>
                </w:jc>
            
            </xsl:if>-->
          </xsl:when>

          <!--2015-06-02，wudi，增加反方向公式转换，start-->
          <xsl:when test="name(.)='字:句_419D' and (./uof:锚点_C644/@图形引用_C62E = following::公式:公式集_C200/公式:数学公式_C201/@标识符_C202)">
            <xsl:variable name="filename">
              <xsl:value-of select="./uof:锚点_C644/@图形引用_C62E"/>
            </xsl:variable>
            <xsl:copy-of select="following::公式:公式集_C200/公式:数学公式_C201[@标识符_C202 = $filename]/m:oMath"/>
          </xsl:when>
          <!--end-->
          
          <xsl:when test="(name(.)='字:修订开始_421F') and (@类型_4221='delete')">
            <xsl:call-template name="paragraphRevision">
              <xsl:with-param name="revType" select="'w:del'"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="(name(.)='字:修订开始_421F') and (@类型_4221='insert')">
            <xsl:call-template name="paragraphRevision">
              <xsl:with-param name="revType" select="'w:ins'"/>
            </xsl:call-template>
          </xsl:when>

          <!--字：域-->
          <xsl:when test="name(.)='字:域开始_419E'">
            <!--2013-03-20，wudi，修复UOF到OOX方向目录与索引的BUG，添加if条件@类型_416E='pageref'-->
            <!--2013-03-27，wudi，修复UOF到OOX方向索引转换的BUG，添加if条件@类型_416E='xe'-->
            <xsl:if test="not(@类型_416E='toc' or @类型_416E='index' or @类型_416E='link' or @类型_416E='pageref' or @类型_416E='xe')">
              <xsl:variable name="type">
                <xsl:choose>
                  <xsl:when test="@类型_416E='revision'">
                    <xsl:value-of select="'revnum'"/>
                  </xsl:when>
                  <xsl:when test="@类型_416E='pageinsection'">
                    <xsl:value-of select="'sectionpages'"/>
                  </xsl:when>
                  <xsl:when test="not(@类型_416E='revision' or @类型_416E='pageinsection')">
                    <xsl:value-of select="@类型_416E"/>
                  </xsl:when>
                  <xsl:when test="not(@类型_416E)">
                    <!--cxl,2012.3.21带圈字符处理-->
                    <xsl:variable name="enclosed">
                      <xsl:for-each select="following-sibling::字:域代码_419F[1]//字:文本串_415B">
                        <xsl:value-of select="concat(.,'')"/>
                      </xsl:for-each>
                    </xsl:variable>
                    <xsl:choose>
                      <xsl:when test="contains($enclosed,'eq')">
                        <xsl:value-of select="'eq'"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="''"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                </xsl:choose>
              </xsl:variable>
              <xsl:call-template name="simplyField">
                <xsl:with-param name="type" select="$type"/>
              </xsl:call-template>
            </xsl:if>
            <!--<xsl:if test="@类型_416E='time'">
              <w:r>
                <w:instrText>
                  <xsl:for-each select="following-sibling::字:域代码_419F/字:段落_416B/字:句_419D/字:文本串_415B">
                    <xsl:value-of select="concat(.,' ')"/>
                  </xsl:for-each>
                </w:instrText>
              </w:r>
            </xsl:if>-->

            <xsl:if test ="@类型_416E='pageref' and not(preceding-sibling::字:句_419D/字:区域开始_4165[@类型_413B='hyperlink'])">
              <w:r>
                <w:rPr/>
                <w:fldChar w:fldCharType="begin"/>
              </w:r>
              <xsl:for-each select ="following-sibling::字:句_419D | following-sibling::字:域代码_419F">
                <xsl:choose>
                  <xsl:when test ="name(.)='字:句_419D'">
                    <xsl:choose>
                      <xsl:when test ="name(following-sibling::*[1]) ='字:域结束_41A0'">
                        <w:r w:rsidRPr="00953832">
                          <w:rPr/>
                          <w:fldChar w:fldCharType="separate"/>
                        </w:r>
                        <xsl:call-template name="run"/>
                        <w:r>
                          <w:rPr/>
                          <w:fldChar w:fldCharType="end"/>
                        </w:r>
                      </xsl:when>

                      <xsl:when test ="not(name(following-sibling::*[1]) ='字:域结束_41A0')">
                        <xsl:call-template name="run"/>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:when test ="name(.)='字:域代码_419F'">
                    <xsl:variable name="str">
                      <xsl:value-of select="./字:段落_416B/字:句_419D/字:文本串_415B"/>
                    </xsl:variable>
                    <xsl:variable name="pageref">
                      <xsl:value-of select="concat(' ',$str,' ')"/>
                    </xsl:variable>
                    <w:r>
                      <w:rPr/>
                      <w:instrText xml:space="preserve"><xsl:value-of select ="$pageref"/></w:instrText>
                    </w:r>
                  </xsl:when>
                </xsl:choose>
              </xsl:for-each>
            </xsl:if>
                        
            <!--2013-04-15，wudi，修复UOF到OOX方向索引转换的BUG，处理INDEX，start-->
            <xsl:if test ="@类型_416E='index'">
              <w:r>
                <w:rPr/>
                <w:fldChar w:fldCharType="begin"/>
              </w:r>
              <xsl:for-each select ="following-sibling::字:句_419D | following-sibling::字:域代码_419F">
                <xsl:choose>
                  <xsl:when test ="name(.)='字:句_419D'">
                    <xsl:choose>
                      <xsl:when test ="name(following-sibling::*[1]) ='字:域结束_41A0'">
                        <w:r w:rsidRPr="00953832">
                          <w:rPr/>
                          <w:fldChar w:fldCharType="separate"/>
                        </w:r>
                        <xsl:call-template name="run"/>
                        <w:r>
                          <w:rPr/>
                          <w:fldChar w:fldCharType="end"/>
                        </w:r>
                      </xsl:when>

                      <xsl:when test ="not(name(preceding-sibling::*[1]) = '字:域代码_419F' or name(following-sibling::*[1]) ='字:域结束_41A0')">
                        <xsl:call-template name="run"/>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:when test ="name(.)='字:域开始_419E'">
                    <xsl:if test ="@类型_416E='pageref'">
                      <w:r>
                        <w:rPr/>
                        <w:fldChar w:fldCharType="begin"/>
                      </w:r>
                    </xsl:if>
                  </xsl:when>
                  <xsl:when test ="name(.)='字:域代码_419F'">
                    <xsl:for-each select ="./字:段落_416B/字:句_419D">
                      <w:r>
                        <xsl:variable name="lowercase">
                          <xsl:if test="字:句属性_4158/字:醒目字体类型_4141='lowercase'">
                            <xsl:value-of select="'true'"/>
                          </xsl:if>
                        </xsl:variable>
                        <xsl:for-each select ="字:句属性_4158 | 字:文本串_415B | 字:空格符_4161">
                          <xsl:choose>
                            <xsl:when test="name(.)='字:句属性_4158'">
                              <xsl:call-template name="RunPrWithChange"/>
                            </xsl:when>
                            <xsl:when test="name(.)='字:文本串_415B'">
                              <xsl:choose>
                                <xsl:when test="ancestor::字:修订开始_421F[1][@类型_4221='delete']">
                                  <w:delText xml:space="preserve"><xsl:if test="$lowercase='true'"><xsl:value-of select="translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/></xsl:if><xsl:if test="$lowercase != 'true'"><xsl:value-of select="."/></xsl:if></w:delText>
                                </xsl:when>
                                <xsl:otherwise>
                                  <w:instrText xml:space="preserve"><xsl:if test="$lowercase = 'true'"><xsl:value-of select="translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/></xsl:if><xsl:if test="$lowercase != 'true'"><xsl:value-of select="."/></xsl:if></w:instrText>
                                </xsl:otherwise>
                              </xsl:choose>
                            </xsl:when>
                            <xsl:when test="name(.)='字:空格符_4161'">
                              <xsl:choose>
                                <xsl:when test ="./@个数_4162">
                                  <w:instrText xml:space="preserve"><xsl:call-template name="loop"><xsl:with-param name="Count" select="@个数_4162"/></xsl:call-template></w:instrText>
                                </xsl:when>
                                <xsl:otherwise>
                                  <w:instrText xml:space="preserve"><xsl:call-template name="loop"><xsl:with-param name="Count" select="1"/></xsl:call-template></w:instrText>
                                </xsl:otherwise>
                              </xsl:choose>
                            </xsl:when>
                          </xsl:choose>
                        </xsl:for-each>
                      </w:r>

                      <!--2014-04-29，wudi，修复目录-互操作转换后内容丢失BUG，start-->
                      <xsl:if test="not(following-sibling::字:域结束_41A0)">
                        <w:r>
                          <w:rPr/>
                          <w:fldChar w:fldCharType="separate"/>
                        </w:r>
                      </xsl:if>
                      <!--end-->
                                            
                    </xsl:for-each>
                  </xsl:when>
                </xsl:choose>
              </xsl:for-each>
            </xsl:if>
            <!--end-->
            
            <!--2013-03-20，wudi，修复UOF到OOX方向目录与索引的BUG，start-->
            <xsl:if test ="@类型_416E='toc'">
              <w:r>
                <w:rPr/>
                <w:fldChar w:fldCharType="begin"/>
              </w:r>
              
              <!--2013-04-15，wudi，考虑到空格符的转换，进行如下优化处理，start-->
              <xsl:for-each select ="following-sibling::字:域代码_419F[1]/字:段落_416B/字:句_419D">
                <w:r>
                  <xsl:variable name="lowercase">
                    <xsl:if test="字:句属性_4158/字:醒目字体类型_4141='lowercase'">
                      <xsl:value-of select="'true'"/>
                    </xsl:if>
                  </xsl:variable>
                  <xsl:for-each select ="字:句属性_4158 | 字:文本串_415B | 字:空格符_4161">
                    <xsl:choose>
                      <xsl:when test="name(.)='字:句属性_4158'">
                        <xsl:call-template name="RunPrWithChange"/>
                      </xsl:when>
                      <xsl:when test="name(.)='字:文本串_415B'">
                        <xsl:choose>
                          <xsl:when test="ancestor::字:修订开始_421F[1][@类型_4221='delete']">
                            <w:delText xml:space="preserve"><xsl:if test="$lowercase='true'"><xsl:value-of select="translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/></xsl:if><xsl:if test="$lowercase != 'true'"><xsl:value-of select="."/></xsl:if></w:delText>
                          </xsl:when>
                          <xsl:otherwise>
                            <w:instrText xml:space="preserve"><xsl:if test="$lowercase = 'true'"><xsl:value-of select="translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/></xsl:if><xsl:if test="$lowercase != 'true'"><xsl:value-of select="."/></xsl:if></w:instrText>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:when>
                      <xsl:when test="name(.)='字:空格符_4161'">
                        <xsl:choose>
                          <xsl:when test ="./@个数_4162">
                            <w:instrText xml:space="preserve"><xsl:call-template name="loop"><xsl:with-param name="Count" select="@个数_4162"/></xsl:call-template></w:instrText>
                          </xsl:when>
                          <xsl:otherwise>
                            <w:instrText xml:space="preserve"><xsl:call-template name="loop"><xsl:with-param name="Count" select="1"/></xsl:call-template></w:instrText>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:for-each>
                </w:r>
              </xsl:for-each>
              <!--end-->
              
              <w:r>
                <w:rPr/>
                <w:fldChar w:fldCharType="separate"/>
              </w:r>
            </xsl:if>
            <!--end-->
            
            <!--2013-03-27，wudi，修复UOF到OOX方向索引转换的BUG，start-->
            <xsl:if test ="@类型_416E='xe'">
              <!--2013-04-10，wudi，修复索引转换的BUG，考虑超链接里包含索引的情况，增加if条件-->
              <xsl:if test ="not(preceding-sibling::字:句_419D/字:区域开始_4165[@类型_413B='hyperlink'])">
                <w:r>
                  <w:rPr>
                    <xsl:apply-templates select ="preceding-sibling::字:句_419D[1]/字:句属性_4158" mode="rpr"/>
                  </w:rPr>
                  <w:fldChar w:fldCharType="begin"/>
                </w:r>
                
                <!--2015-04-26，wudi，修复Xe域转换BUG，start-->
                <w:r>
                  <xsl:variable name="tmp">
                    <xsl:for-each select="following-sibling::字:域代码_419F/字:段落_416B/字:句_419D/字:文本串_415B">
                      <xsl:value-of select="."/>
                    </xsl:for-each>
                  </xsl:variable>
                  <w:instrText xml:space="preserve"><xsl:value-of select="$tmp"/></w:instrText>
                </w:r>
                <!--<w:r>
                  <w:instrText xml:space="preserve"> XE "</w:instrText>
                </w:r>
                <xsl:variable name ="tmp1">
                  <xsl:value-of select ="following-sibling::字:域代码_419F/字:段落_416B/字:句_419D/字:文本串_415B"/>
                </xsl:variable>
                <xsl:variable name ="tmp2">
                  <xsl:value-of select ="substring-after($tmp1,'XE')"/>
                </xsl:variable>
                <xsl:variable name ="tmp3">
                  <xsl:value-of select ="normalize-space($tmp2)"/>
                </xsl:variable>
                <xsl:variable name ="strlgth">
                  <xsl:value-of select ="string-length($tmp3)"/>
                </xsl:variable>
                <xsl:variable name ="lgth">
                  <xsl:value-of select ="$strlgth - 2"/>
                </xsl:variable>
                <w:r>
                  <w:rPr>
                    <xsl:apply-templates select ="preceding-sibling::字:句_419D[1]/字:句属性_4158" mode="rpr"/>
                  </w:rPr>
                  <w:instrText>
                    <xsl:value-of select ="substring($tmp3,2,$lgth)"/>
                  </w:instrText>
                </w:r>
                <w:r>
                  <w:instrText xml:space="preserve">" </w:instrText>
                </w:r>-->
                <!--end-->
                
                <w:r>
                  <w:rPr>
                    <xsl:apply-templates select ="preceding-sibling::字:句_419D[1]/字:句属性_4158" mode="rpr"/>
                  </w:rPr>
                  <w:fldChar w:fldCharType="end"/>
                </w:r>
              </xsl:if>
            </xsl:if>
            <!--end-->
            
          </xsl:when>

          <!--2013-03-20，wudi，修复UOF到OOX方向目录与索引的BUG，start-->
          <!--2013-04-15，wudi，针对部分特例，区域开始和区域结束不在一个段落的情况，增加条件，start-->
          <xsl:when test ="name(.)='字:域结束_41A0' and (not(preceding-sibling::字:域开始_419E) or contains(preceding-sibling::字:句_419D[1]/字:区域结束_4167/@标识符引用_4168,'hlk'))">
            <w:r>
              <w:rPr/>
              <w:fldChar w:fldCharType="end"/>
            </w:r>
          </xsl:when>
          <!--end-->

          <!--2013-04-09，wudi，增加test判断条件and following-sibling::字:域开始_419E[@类型_416E ='pageref']-->
          <xsl:when test ="name(.)='字:句_419D' and child::字:区域开始_4165[@类型_413B='hyperlink'] and following-sibling::字:域开始_419E[@类型_416E ='pageref']">
            <w:hyperlink>

              <!--2013-03-21，wudi，修复UOF到OOX方向目录与索引的BUG，start-->
              <xsl:for-each select ="following-sibling::字:域代码_419F">
                
                <!--2013-04-15，wudi，考虑PageRef的情况，分情况，start-->
                <xsl:choose>
                  <xsl:when test ="contains(./字:段落_416B/字:句_419D/字:文本串_415B,'PAGEREF')">
                    <xsl:variable name ="temp">
                      <xsl:value-of select ="substring-before(./字:段落_416B/字:句_419D/字:文本串_415B,' \h')"/>
                    </xsl:variable>
                    <xsl:variable name ="Wanchor">
                      <xsl:value-of select ="substring-after($temp,'PAGEREF ')"/>
                    </xsl:variable>
                    <xsl:attribute name ="w:anchor">
                      <xsl:value-of select ="$Wanchor"/>
                    </xsl:attribute>
                    <xsl:attribute name ="w:history">
                      <xsl:value-of select ="1"/>
                    </xsl:attribute>
                  </xsl:when>
                  <xsl:when test ="contains(./字:段落_416B/字:句_419D/字:文本串_415B,'PageRef')">
                    <xsl:variable name ="Wanchor">
                      <xsl:value-of select ="./字:段落_416B/字:句_419D/字:文本串_415B[2]"/>
                    </xsl:variable>
                    <xsl:attribute name ="w:anchor">
                      <xsl:value-of select ="$Wanchor"/>
                    </xsl:attribute>
                    <xsl:attribute name ="w:history">
                      <xsl:value-of select ="1"/>
                    </xsl:attribute>
                  </xsl:when>
                </xsl:choose>
              </xsl:for-each>
              
              <!--end-->
              
              <!--end-->

              <xsl:for-each select ="following-sibling::字:句_419D | following-sibling::字:域开始_419E | following-sibling::字:域代码_419F">
                <xsl:choose>

                  <!--2013-04-10，wudi，修复索引转换BUG，考虑目录项里标记索引项的情况-->
                  <xsl:when test ="name(.)='字:句_419D' and not(./字:区域结束_4167)">
                    
                    <!--2013-04-10，wudi，增加if限制条件and not(name(preceding-sibling::*[1]) = '字:域代码_419F')-->
                    <!--2013-04-15，wudi，限制条件更加严格，仅排除XE的情况-->
                    <xsl:if test ="name(following-sibling::*[1]) ='字:域结束_41A0' and not(contains(preceding-sibling::字:域代码_419F[1]/字:段落_416B/字:句_419D/字:文本串_415B,'XE'))">
                      <w:r w:rsidRPr="00953832">
                        <w:rPr/>
                        <w:fldChar w:fldCharType="separate"/>
                      </w:r>
                      <!--<w:r>
                        <w:rPr/>
                        <w:t>
                          <xsl:value-of select ="./字:文本串_415B"/>
                        </w:t>
                      </w:r>-->
                      <xsl:call-template name="run"/>
                      <w:r>
                        <w:rPr/>
                        <w:fldChar w:fldCharType="end"/>
                      </w:r>
                    </xsl:if>
                    
                    <!--2013-04-10，wudi，增加if限制条件or name(following-sibling::*[1]) ='字:域结束_41A0'-->
                    <xsl:if test ="not(name(preceding-sibling::*[1]) = '字:域代码_419F' or name(following-sibling::*[1]) ='字:域结束_41A0')">
                      <xsl:call-template name="run"/>
                    </xsl:if>
                  </xsl:when>
                  <xsl:when test ="name(.)='字:域开始_419E'">
                    <xsl:choose>
                      <xsl:when test ="@类型_416E='pageref'">
                        <w:r>
                          <w:rPr/>
                          <w:fldChar w:fldCharType="begin"/>
                        </w:r>
                      </xsl:when>

                      <!--2013-04-10，wudi，增加对索引项的处理，start-->
                      <xsl:when test ="@类型_416E='xe'">
                        <w:r>
                          <w:rPr>
                            <xsl:apply-templates select ="preceding-sibling::字:句_419D[1]/字:句属性_4158" mode="rpr"/>
                          </w:rPr>
                          <w:fldChar w:fldCharType="begin"/>
                        </w:r>
                      </xsl:when>
                      <!--end-->
                    </xsl:choose>
                  </xsl:when>
                  <xsl:when test ="name(.)='字:域代码_419F'">
                    <xsl:if test ="preceding-sibling::字:域开始_419E[1][@类型_416E='pageref']">
                      
                      <!--2013-04-15，wudi，考虑到空格符的转换，进行如下优化处理，start-->
                      <xsl:for-each select ="./字:段落_416B/字:句_419D">
                        <w:r>
                          <xsl:variable name="lowercase">
                            <xsl:if test="字:句属性_4158/字:醒目字体类型_4141='lowercase'">
                              <xsl:value-of select="'true'"/>
                            </xsl:if>
                          </xsl:variable>
                          <xsl:for-each select ="字:句属性_4158 | 字:文本串_415B | 字:空格符_4161">
                            <xsl:choose>
                              <xsl:when test="name(.)='字:句属性_4158'">
                                <xsl:call-template name="RunPrWithChange"/>
                              </xsl:when>
                              <xsl:when test="name(.)='字:文本串_415B'">
                                <xsl:choose>
                                  <xsl:when test="ancestor::字:修订开始_421F[1][@类型_4221='delete']">
                                    <w:delText xml:space="preserve"><xsl:if test="$lowercase='true'"><xsl:value-of select="translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/></xsl:if><xsl:if test="$lowercase != 'true'"><xsl:value-of select="."/></xsl:if></w:delText>
                                  </xsl:when>
                                  <xsl:otherwise>
                                    <w:instrText xml:space="preserve"><xsl:if test="$lowercase = 'true'"><xsl:value-of select="translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/></xsl:if><xsl:if test="$lowercase != 'true'"><xsl:value-of select="."/></xsl:if></w:instrText>
                                  </xsl:otherwise>
                                </xsl:choose>
                              </xsl:when>
                              <xsl:when test="name(.)='字:空格符_4161'">
                                <xsl:choose>
                                  <xsl:when test ="./@个数_4162">
                                    <w:instrText xml:space="preserve"><xsl:call-template name="loop"><xsl:with-param name="Count" select="@个数_4162"/></xsl:call-template></w:instrText>
                                  </xsl:when>
                                  <xsl:otherwise>
                                    <w:instrText xml:space="preserve"><xsl:call-template name="loop"><xsl:with-param name="Count" select="1"/></xsl:call-template></w:instrText>
                                  </xsl:otherwise>
                                </xsl:choose>
                              </xsl:when>
                            </xsl:choose>
                          </xsl:for-each>
                        </w:r>
                      </xsl:for-each>
                      <!--end-->
                      
                    </xsl:if>

                    <!--2013-04-10，wudi，增加对索引项的处理，start-->
                    <xsl:if test ="preceding-sibling::字:域开始_419E[1][@类型_416E='xe']">
                      <xsl:for-each select ="./字:段落_416B/字:句_419D">
                        <xsl:call-template name="run"/>
                      </xsl:for-each>
                      <w:r>
                        <w:rPr>
                          <xsl:apply-templates select ="preceding-sibling::字:句_419D[1]/字:句属性_4158" mode="rpr"/>
                        </w:rPr>
                        <w:fldChar w:fldCharType="end"/>
                      </w:r>
                    </xsl:if>
                    <!--end-->
                    
                  </xsl:when>
                  <!--end-->
                  
                </xsl:choose>
              </xsl:for-each>
            </w:hyperlink>
          </xsl:when>
          <!--end-->
          
          <!--wcz,修复页眉，页脚文本字段不能显示的问题，2013/2/28-->
          <!--2013-03-20，wudi，添加test判断条件and not(preceding-sibling::字:句_419D/字:区域开始_4165[@类型_413B='hyperlink'])-->
          <!--2013-04-09，wudi，增加test判断条件and not(following-sibling::字:域开始_419E[@类型_416E ='pageref'] or preceding-sibling::字:域开始_419E[@类型_416E ='pageref'])-->
          <!--2014-05-19，wudi，增加test判断条件(name(following-sibling::*[1]) ='字:域结束_41A0' and ./uof:锚点_C644)，可能会有图片转换，start-->
          <xsl:when
            test="name(.)='字:句_419D' and (not(name(preceding-sibling::*[1]) = '字:域代码_419F' or name(following-sibling::*[1]) = '字:域结束_41A0') or (name(following-sibling::*[1]) ='字:域结束_41A0' and ./uof:锚点_C644)) and not(following-sibling::字:域开始_419E[@类型_416E ='pageref'] or preceding-sibling::字:域开始_419E[@类型_416E ='pageref'])">
            <xsl:call-template name="run"/>
          </xsl:when>
          <!--end-->
        </xsl:choose>
      </xsl:for-each>
    </w:p>
  </xsl:template>
  <xsl:template name="pagesinobjectset">
    <!--yx,manage regionstart in objectset,2010.2.2-->
    <!--yx,current node:.../字:域开始,2010.2.2-->

    <xsl:if test="not(@类型_416E='toc' or @类型_416E='index' or @类型_416E='link' or @类型_416E='time')">
      <xsl:variable name="type">
        <xsl:choose>
          <xsl:when test="@类型_416E='revision'">
            <xsl:value-of select="'revnum'"/>
          </xsl:when>
          <xsl:when test="@类型_416E='pageinsection'">
            <xsl:value-of select="'sectionpages'"/>
          </xsl:when>
          <xsl:when test="not(@类型_416E='revision' or @类型_416E='pageinsection')">
            <xsl:value-of select="@类型_416E"/>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>
      <xsl:call-template name="simplyField">
        <xsl:with-param name="type" select="$type"/>
      </xsl:call-template>
    </xsl:if>
    <xsl:if test="@类型_416E='time'">
      <w:r>
        <w:instrText>
          <xsl:for-each select="字:域代码_419/字:段落_416B/字:句_419D/字:文本串_415B">
            <xsl:value-of select="concat(.,' ')"/>
          </xsl:for-each>
        </w:instrText>
      </w:r>
    </xsl:if>
  </xsl:template>

  <xsl:template name="pPrWithpStyle">
    <w:pPr>
      <xsl:call-template name="pPrWithChange"/>
      <!--这里的代码是以前考虑首字下沉写的<xsl:if test="following::字:句_419D[1]/uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:相对_4109/@参考点_410A">
        <w:jc>
          <xsl:attribute name="w:val">
            <xsl:value-of select="following::字:句_419D[1]/uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:相对_4109/@参考点_410A"/>
          </xsl:attribute>
        </w:jc>
      </xsl:if>-->
      <xsl:if test="not(parent::字:段落_416B/字:句_419D/字:文本串_415B) and not(字:句属性_4158)">
        <xsl:apply-templates select="following-sibling::字:句_419D[字:句属性_4158][1]/字:句属性_4158" mode="rpr"/>
      </xsl:if>
    </w:pPr>
  </xsl:template>

  <xsl:template name="pPrWithChange">
    <xsl:choose>
      <xsl:when test="not(following-sibling::字:修订开始_421F[@类型_4221='format']/字:段落属性_419B)">
        <xsl:call-template name="pPr"/>
      </xsl:when>
      <xsl:when test="following-sibling::字:修订开始_421F[@类型_4221='format']/字:段落属性_419B">
        <xsl:apply-templates select="following-sibling::字:修订开始_421F[@类型_4221='format']/字:段落属性_419B" mode="withRev"/>
        <!--<w:pPrChange>
        <xsl:call-template name="pPr"/>
      </w:pPrChange>-->
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="字:段落属性_419B" mode="withRev">
    <xsl:call-template name="pPr"/>
  </xsl:template>
  
  <xsl:template name="pPr">
    <!--段落属性模板，CXL，2012.2.15,  current node:字:段落属性_419B-->
    <xsl:if test="@式样引用_419C">
      <w:pStyle>
        <xsl:attribute name="w:val">
          <xsl:value-of select="@式样引用_419C"/>
        </xsl:attribute>
      </w:pStyle>
    </xsl:if>
    
    <!--2015-05-25，wudi，默认在相同样式的段落间不添加空格，start--><!--
    <w:contextualSpacing/>
    --><!--end-->
    
    <xsl:if test="./字:是否与下段同页_418D">
      <xsl:call-template name="InsertKeepNext"/>
    </xsl:if>
    <xsl:if test="./字:是否段中不分页_418C">
      <xsl:call-template name="InsertKeepLines"/>
    </xsl:if>
    <xsl:if test="./字:是否段前分页_418E">
      <xsl:call-template name="InsertPageBreakBefore"/>
    </xsl:if>
    <xsl:if test="./字:首字下沉_4191"><!--cxl,2012.4.11-->
      <w:framePr>
        <xsl:attribute name="w:dropCap">
          <xsl:choose>
            <xsl:when test="./字:首字下沉_4191/@类型_413B='dropped'">
              <xsl:value-of select="'drop'"/>
            </xsl:when>
            <xsl:when test="./字:首字下沉_4191/@类型_413B='margin'">
              <xsl:value-of select="'margin'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="w:lines">
          <xsl:choose>
            <xsl:when test="./字:首字下沉_4191/@行数_4178">
              <xsl:value-of select="./字:首字下沉_4191/@行数_4178"/>
            </xsl:when>
            <xsl:when test="not(./字:首字下沉_4191/@行数_4178)">
              <xsl:value-of select="3"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="w:wrap">around</xsl:attribute>
        <xsl:attribute name="w:vAnchor">text</xsl:attribute>
        <xsl:attribute name="w:hAnchor">text</xsl:attribute>
      </w:framePr>
    </xsl:if>

    <!--
    根据差异文档，孤行控制_418A和寡行控制均作为widowControl处理
    Office没有实现寡行控制
    -->
    <!-- 
    <xsl:choose>
      <xsl:when test="(./字:孤行控制_418A &gt; 0) or (./字:寡行控制 &gt; 0)">
        <w:widowControl>
          <xsl:attribute name="w:val">true</xsl:attribute>
        </w:widowControl>
      </xsl:when>
      <xsl:when test="(./字:孤行控制_418A = 0) and (./字:寡行控制 = 0)">
        <w:widowControl>
          <xsl:attribute name="w:val">false</xsl:attribute>
        </w:widowControl>
      </xsl:when>
      <xsl:when test="(./字:孤行控制_418A &gt; 0) and not(./字:寡行控制)">
        <w:widowControl>
          <xsl:attribute name="w:val">true</xsl:attribute>
        </w:widowControl>
      </xsl:when>
      <xsl:when test="(./字:孤行控制_418A = 0) and not(./字:寡行控制)">
        <w:widowControl>
          <xsl:attribute name="w:val">false</xsl:attribute>
        </w:widowControl>
      </xsl:when>
      <xsl:when test="(./字:寡行控制 &gt; 0) and not(./字:孤行控制_418A)">
        <w:widowControl>
          <xsl:attribute name="w:val">true</xsl:attribute>
        </w:widowControl>
      </xsl:when>     
    </xsl:choose>-->
    <w:widowControl>
      <xsl:attribute name="w:val">
        <xsl:choose>
          <xsl:when test="./字:孤行控制_418A='1'">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:when test="./字:孤行控制_418A='0' or not(./字:孤行控制_418A)">
            <xsl:value-of select="'false'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </w:widowControl>
    
    <xsl:if test="./字:自动编号信息_4186">
      <xsl:call-template name="pPrNumbering"/>
    </xsl:if>
    <xsl:if test="./字:是否取消行号_4193">
      <xsl:call-template name="InsertSuppressLineNumber"/>
    </xsl:if>
    <xsl:if test="./字:边框_4133">
      <xsl:apply-templates select="字:边框_4133" mode="paragraph"/>
    </xsl:if>
    <xsl:if test="./字:填充_4134">
      <xsl:apply-templates select="字:填充_4134" mode="shdCommon"/>
    </xsl:if>
    <xsl:if test="./字:制表位设置_418F">
      <xsl:apply-templates select="字:制表位设置_418F"/>
    </xsl:if>
    <xsl:if test="./字:是否取消断字_4192">
      <xsl:call-template name="InsertSuppressAutoHyphens"/>
    </xsl:if>
    <xsl:if test="./字:是否采用中文习惯首尾字符_4197">
      <xsl:call-template name="InsertKinsoku"/>
    </xsl:if>
    <xsl:if test="./字:是否允许单词断字_4194">
      <xsl:call-template name="InsertWordwrap"/>
    </xsl:if>
    <xsl:if test="./字:是否行首尾标点控制_4195">
      <xsl:call-template name="InsertOverflowPunct"/>
    </xsl:if>
    <xsl:if test="./字:是否行首标点压缩_4196">
      <xsl:call-template name="InsertTopLinePunct"/>
    </xsl:if>
    <xsl:if test="./字:是否自动调整中英文字符间距_4198">
      <xsl:call-template name="InsertAutoSpaceDE"/>
    </xsl:if>
    <xsl:if test="./字:是否自动调整中文与数字间距_4199">
      <xsl:call-template name="InsertAutoSpaceDN"/>
    </xsl:if>
    <xsl:if test="./字:是否有网格自动调整右缩进_419A">
      <xsl:call-template name="InsertAdjustRightInd"/>
    </xsl:if>
    <xsl:if test="./字:是否对齐网格_4190">
      <xsl:call-template name="InsertSnapToGrid"/>
    </xsl:if>

    <xsl:variable name="LineSpacingExist">
      <xsl:value-of select="boolean(./字:行距_417E)"/>
    </xsl:variable>
    <xsl:variable name="ParaSpacingExist">
      <xsl:value-of select="boolean(./字:段间距_4180)"/>
    </xsl:variable>

    <!--2015-05-15，wudi，修复段落属性-段间距转换BUG，start-->
    <!--<xsl:if test="($LineSpacingExist='true') or ($ParaSpacingExist='true')">
      <w:spacing>-->
        <!--<xsl:choose>-->
          <xsl:if test="$LineSpacingExist='true'">
            <w:spacing>
              <xsl:call-template name="InsertLineSpacing"/>
            </w:spacing>
          </xsl:if>
          <xsl:if test="$ParaSpacingExist='true'">
            <w:spacing>
              <xsl:call-template name="InsertParaSpacing"/>
            </w:spacing>
          </xsl:if>
        <!--</xsl:choose>-->
      <!--</w:spacing>
    </xsl:if>-->
    <!--end-->
    
    <xsl:if test="./字:缩进_411D">
      <xsl:apply-templates select="字:缩进_411D"/>
    </xsl:if>
    <xsl:if test="./字:自动编号信息_4186 and not(./字:缩进_411D) and not(name(.)='式样:段落式样_9912')">
      <w:ind w:firstLineChars="0"/>
    </xsl:if>
    <xsl:if test="./字:对齐_417D"><!--这里的and条件应该是以前为了首字下沉而写的 and (not(following::字:句_419D[1]/uof:锚点_C644/uof:位置_C620/uof:水平_4106/uof:相对_4109/@参考点_410A))-->
      <xsl:call-template name="InsertJcAndTextAlign"/>
    </xsl:if>
    <xsl:if test="./字:大纲级别_417C">
      <xsl:call-template name="InsertOutlineLvl"/>
    </xsl:if>
    <xsl:if test="./字:句属性_4158">
      <xsl:choose>
        <xsl:when test="name(.)='式样:段落式样_9912'">
          <!--do nothing-->
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates select="./字:句属性_4158" mode="rpr"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <!--经过预处理“字:分节_416A”被调整到"字:段落属性__419B"之下-->
    <xsl:if test="./字:分节_416A">
      <xsl:apply-templates select="./字:分节_416A"/>
    </xsl:if>
	  
	  <xsl:if test="../../字:修订开始_421F">
		  <w:pPrChange>
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
				  <xsl:apply-templates select="preceding::规则:修订信息_B60F[@字:标识符=$revId]" mode="rev"/>
			  </xsl:attribute>
			  <xsl:attribute name="w:author">
				  <xsl:value-of select="$username"/>
			  </xsl:attribute>
			  <xsl:attribute name="w:date">
				  <xsl:value-of select="concat($revDate,'Z')"/>
			  </xsl:attribute>
			  <w:pPr/>
		  </w:pPrChange>
	  </xsl:if>
  </xsl:template>

  <xsl:template name="InsertKeepNext">
    <w:keepNext>
      <xsl:attribute name="w:val">
        <xsl:value-of select="./字:是否与下段同页_418D"/>
      </xsl:attribute>
    </w:keepNext>
  </xsl:template>
  
  <xsl:template name="InsertPageBreakBefore">
    <w:pageBreakBefore>
      <xsl:attribute name="w:val">
        <xsl:value-of select="./字:是否段前分页_418E"/>
      </xsl:attribute>
    </w:pageBreakBefore>
  </xsl:template>

  <xsl:template name="InsertSuppressLineNumber">
    <w:suppressLineNumbers>
      <xsl:attribute name="w:val">
        <xsl:value-of select="./字:是否取消行号_4193"/>
      </xsl:attribute>
    </w:suppressLineNumbers>
  </xsl:template>

  <xsl:template match="字:边框_4133" mode="paragraph">
    <w:pBdr>
      <xsl:if test="./uof:上_C614">
        <w:top>
          <xsl:if test="./@阴影类型_C645='right-bottom'"><!--cxl,2012.3.17-->
            <xsl:attribute name="w:shadow">
              <xsl:value-of select="'1'"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:apply-templates select="uof:上_C614" mode="pBdrChildren"/>
        </w:top>
      </xsl:if>
      <xsl:if test="./uof:左_C613">
        <w:left>
          <xsl:if test="./@阴影类型_C645='right-bottom'">
            <xsl:attribute name="w:shadow">
              <xsl:value-of select="'1'"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:apply-templates select="uof:左_C613" mode="pBdrChildren"/>
        </w:left>
      </xsl:if>
      <xsl:if test="./uof:下_C616">
        <w:bottom>
          <xsl:if test="./@阴影类型_C645='right-bottom'">
            <xsl:attribute name="w:shadow">
              <xsl:value-of select="'1'"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:apply-templates select="uof:下_C616" mode="pBdrChildren"/>
        </w:bottom>
      </xsl:if>
      <xsl:if test="./uof:右_C615">
        <w:right>
          <xsl:if test="./@阴影类型_C645='right-bottom'">
            <xsl:attribute name="w:shadow">
              <xsl:value-of select="'1'"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:apply-templates select="uof:右_C615" mode="pBdrChildren"/>
        </w:right>
      </xsl:if>
    </w:pBdr>
  </xsl:template>
  <xsl:template match="uof:上_C614" mode="pBdrChildren">
    <xsl:call-template name="SingleBorder"/>
  </xsl:template>
  <xsl:template match="uof:左_C613" mode="pBdrChildren">
    <xsl:call-template name="SingleBorder"/>
  </xsl:template>
  <xsl:template match="uof:下_C616" mode="pBdrChildren">
    <xsl:call-template name="SingleBorder"/>
  </xsl:template>
  <xsl:template match="uof:右_C615" mode="pBdrChildren">
    <xsl:call-template name="SingleBorder"/>
  </xsl:template>
  <xsl:template name="SingleBorder"><!--解决段落边框式样丢失 2011-04-28--><!--与下划线中的一致-->
    <xsl:variable name="bordertp">
      <xsl:choose>
        <xsl:when test="@线型_C60D">
          <xsl:value-of select="./@线型_C60D"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'none'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="borderxs">
      <xsl:choose>
        <xsl:when test="@虚实_C60E">
          <xsl:value-of select="./@虚实_C60E"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'none'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:attribute name="w:val">
      <xsl:choose>

        <!--in common
        <xsl:when test="$bordertp='dotted' or $bordertp='dotted-heavy'">dotted</xsl:when>
        <xsl:when test="$bordertp='dash'or $bordertp='dashed-heavy'">dashed</xsl:when>
        <xsl:when test="$bordertp='dash-long' or $bordertp='dash-long-heavy'">dashSmallGap</xsl:when>
        <xsl:when test="$bordertp='dot-dash' or $bordertp='dash-dot-heavy'">dotDash</xsl:when>
        <xsl:when test="$bordertp='dot-dot-dash' or $bordertp='dash-dot-dot-heavy'">dotDotDash</xsl:when>
        <xsl:when test="$bordertp='wave' or $bordertp='wavy-heavy'">wave</xsl:when>
        <xsl:when test="$bordertp='wavy-double'">doubleWave</xsl:when>
        -->

        <!--same with table border-->
        <xsl:when test="$bordertp='none'">none</xsl:when>
        <xsl:when test="$bordertp='single' and $borderxs='square-dot'">
          <xsl:value-of select="'dotted'"/>
        </xsl:when>
        <!--wcz,2013/3/26,解决虚边框线bug-->
        <xsl:when test="$bordertp='thick' and $borderxs='square-dot'">
          <xsl:value-of select="'dashSmallGap'"/>
        </xsl:when>
        <!--end-->
        <xsl:when test="$bordertp='single' and $borderxs='dash'">
          <xsl:value-of select="'dashed'"/>
        </xsl:when>
        <xsl:when test="$bordertp='single' and $borderxs='dash-dot'">
          <xsl:value-of select="'dotDash'"/>
        </xsl:when>
        <!--case only have double
        <-->
        <xsl:when test="$bordertp='single' and $borderxs='solid'">
          <xsl:value-of select="'single'"/>
        </xsl:when>
        <xsl:when test="$bordertp='single' and $borderxs='round-dot'">
          <xsl:value-of select="'dashSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bordertp='single' and $borderxs='dash-dot-dot'">
          <xsl:value-of select="'dotDotDash'"/>
        </xsl:when>

        <xsl:when test="$bordertp='double' and $borderxs='solid'">
          <xsl:value-of select="'double'"/>
        </xsl:when>
        <xsl:when test="$bordertp='double' and $borderxs='dash-dot-dot'">
          <xsl:value-of select="'triple'"/>
        </xsl:when>
        <xsl:when test="$bordertp='thin-thick' and $borderxs='solid'">
          <xsl:value-of select="'thickThinSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bordertp='thick-thin' and $borderxs='solid'">
          <xsl:value-of select="'thinThickSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bordertp='thin-thick' and $borderxs='dash'">
          <xsl:value-of select="'thickThinMediumGap'"/>
        </xsl:when>
        <xsl:when test="$bordertp='thick-thin' and $borderxs='dash'">
          <xsl:value-of select="'thinThickMediumGap'"/>
        </xsl:when>
        <xsl:when test="$bordertp='thick-thin' and $borderxs='long-dash'">
          <xsl:value-of select="'thinThickLargeGap'"/>
        </xsl:when>
        <xsl:when test="$bordertp='thin-thick' and $borderxs='long-dash'">
          <xsl:value-of select="'thickThinLargeGap'"/>
        </xsl:when>
        <xsl:when test="$bordertp='thick-between-thin' and $borderxs='solid'">
          <xsl:value-of select="'thinThickThinSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bordertp='thick-between-thin' and $borderxs='dash'">
          <xsl:value-of select="'thinThickThinMediumGap'"/>
        </xsl:when>
        <xsl:when test="$bordertp='thick-between-thin' and $borderxs='long-dash'">
          <xsl:value-of select="'thinThickThinLargeGap'"/>
        </xsl:when>
        <xsl:when test="$bordertp='single' and $borderxs='long-dash'">
          <xsl:value-of select="'wave'"/>
        </xsl:when>
        <!--wcz,2013/3/26,解决波浪型边框线bug，start-->
        <xsl:when test="$bordertp='single' and $borderxs='wave'">
          <xsl:value-of select="'wave'"/>
        </xsl:when>
        <xsl:when test="$bordertp='double' and $borderxs='wave'">
          <xsl:value-of select="'doubleWave'"/>
        </xsl:when>
        <!--end-->
        <xsl:when test="$bordertp='double' and $borderxs='round-dot'">
          <xsl:value-of select="'doubleWave'"/>
        </xsl:when>
        <xsl:when test="$bordertp='single' and $borderxs='long-dash-dot'">
          <xsl:value-of select="'dashDotStroked'"/>
        </xsl:when>
        <xsl:when test="$bordertp='double' and $borderxs='dash'">
          <xsl:value-of select="'threeDEmboss'"/>
        </xsl:when>
        <xsl:when test="$bordertp='double' and $borderxs='dash-dot'">
          <xsl:value-of select="'threeDEngrave'"/>
        </xsl:when>
        <xsl:when test="$bordertp='double' and $borderxs='long-dash'">
          <xsl:value-of select="'outset'"/>
        </xsl:when>
        <xsl:when test="$bordertp='double' and $borderxs='long-dash-dot'">
          <xsl:value-of select="'inset'"/>
        </xsl:when>
        <!--以上基于wps设置进行对应,以下考虑永中中没有虚实的情况-->
        <xsl:when test="$bordertp='single' and $borderxs='dotted-heavy'">
          <xsl:value-of select="'dotted'"/>
        </xsl:when>
        <xsl:when test="$bordertp='wave' and $borderxs='solid'">
          <xsl:value-of select="'wave'"/>
        </xsl:when>
        <xsl:when test="$bordertp='double' and $borderxs='solid'">
          <xsl:value-of select="'doubleWave'"/>
        </xsl:when>
        <xsl:when test="$bordertp='dash-dot'">
          <xsl:value-of select="'dashDotStroked'"/>
        </xsl:when>
        <!--zhaobj-->
        <xsl:when test="$bordertp='thick-thin'">
          <xsl:value-of select="'thinThickSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bordertp='thin-thick'">
          <xsl:value-of select="'thickThinSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bordertp='double'">
          <xsl:value-of select="'double'"/>
        </xsl:when>
        <xsl:when test="$bordertp='thick-between-thin'">
          <xsl:value-of select="'thinThickThinSmallGap'"/>
        </xsl:when>

        <xsl:otherwise>single</xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>

    <xsl:if test="@颜色_C611">
      <xsl:attribute name="w:color">
        <xsl:choose>
          <xsl:when test="@颜色_C611='auto'">auto</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="substring(@颜色_C611,2)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="@宽度_C60F">
      <xsl:attribute name="w:sz">
        <xsl:call-template name="eighthPointMeasure">
          <xsl:with-param name="lengthVal" select="@宽度_C60F"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="@边距_C610">
      <xsl:attribute name="w:space">
        <xsl:call-template name="pointMeasure">
          <xsl:with-param name="lengthVal" select="@边距_C610"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <!--阴影改变较大*********************************************************************-->
    <xsl:if test="@uof:阴影">
      <xsl:attribute name="w:shadow">
        <xsl:value-of select="@uof:阴影"/>
      </xsl:attribute>
    </xsl:if>

  </xsl:template>


  <xsl:template match="字:填充_4134" mode="pPrShd">
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
            <xsl:when test="./图:图案_800A[@前景色_800B='auto']">
              <xsl:value-of select="./图:图案_800A/@前景色_800B"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring(./图:图案_800A/@前景色_800B,2)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="w:fill">
          <xsl:choose>
            <xsl:when test="./图:图案_800A[@背景色_800C='auto']">
              <xsl:value-of select="./图:图案_800A/@背景色_800C"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring(./图:图案_800A/@背景色_800C,2)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
    </w:shd>
  </xsl:template>
  
  <xsl:template match="字:制表位设置_418F">
    <w:tabs>
      <xsl:for-each select="字:制表位_4171">
        <!--去掉下面的这句，会出错，因为uof中子元素制表位可以不出现，此时node()取到的是一个空集合-->
        <!--<xsl:if test="name(.)='字:制表位_4171'">-->
          <w:tab>
            <xsl:attribute name="w:val">
              <xsl:choose>
                <xsl:when test="./@类型_4173">
                  <xsl:value-of select="./@类型_4173"/>
                </xsl:when>
                <xsl:otherwise>left</xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>

            <xsl:attribute name="w:pos">
              <xsl:choose>
                <xsl:when test="./@位置_4172">
                  <xsl:call-template name="twipsMeasure">
                    <xsl:with-param name="lengthVal" select="./@位置_4172"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>420</xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>

            <xsl:if test="./@前导符_4174">
              <xsl:attribute name="w:leader">
                <xsl:choose>
                  <xsl:when test="./@前导符_4174='.' or ./@前导符_4174='&#46;'">dot</xsl:when>
                  <xsl:when test="./@前导符_4174='-' or ./@前导符_4174='&#45;'">hyphen</xsl:when>
                  <xsl:when test="./@前导符_4174='_' or ./@前导符_4174='&#95;'">underscore</xsl:when>
                  <xsl:when test="./@前导符_4174='&#183;'">middleDot</xsl:when>
                  <xsl:otherwise>none</xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </w:tab>
        <!--</xsl:if>-->
      </xsl:for-each>
    </w:tabs>
  </xsl:template>

  <xsl:template name="InsertSuppressAutoHyphens">
    <w:suppressAutoHyphens>
      <xsl:attribute name="w:val">
        <xsl:value-of select="./字:是否取消断字_4192"/>
      </xsl:attribute>
    </w:suppressAutoHyphens>
  </xsl:template>

  <xsl:template name="InsertKinsoku">
    <w:kinsoku>
      <xsl:attribute name="w:val">
        <xsl:value-of select="./字:是否采用中文习惯首尾字符_4197"/>
      </xsl:attribute>
    </w:kinsoku>
  </xsl:template>

  <xsl:template name="InsertWordwrap">
    <w:wordWrap>
      <xsl:attribute name="w:val">
        <xsl:choose>
          <xsl:when test="./字:是否允许单词断字_4194='true'">false</xsl:when>
          <xsl:when test="./字:是否允许单词断字_4194='false'">true</xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </w:wordWrap>
  </xsl:template>

  <xsl:template name="InsertOverflowPunct">
    <w:overflowPunct>
      <xsl:attribute name="w:val">
        <xsl:value-of select="./字:是否行首尾标点控制_4195"/>
      </xsl:attribute>
    </w:overflowPunct>
  </xsl:template>

  <xsl:template name="InsertTopLinePunct">
    <w:topLinePunct>
      <xsl:attribute name="w:val">
        <xsl:value-of select="./字:是否行首标点压缩_4196"/>
      </xsl:attribute>
    </w:topLinePunct>
  </xsl:template>

  <xsl:template name="InsertAutoSpaceDE">
    <w:autoSpaceDE>
      <xsl:attribute name="w:val">
        <xsl:value-of select="./字:是否自动调整中英文字符间距_4198"/>
      </xsl:attribute>
    </w:autoSpaceDE>
  </xsl:template>

  <xsl:template name="InsertAutoSpaceDN">
    <w:autoSpaceDN>
      <xsl:attribute name="w:val">
        <xsl:value-of select="./字:是否自动调整中文与数字间距_4199"/>
      </xsl:attribute>
    </w:autoSpaceDN>
  </xsl:template>

  <xsl:template name="InsertAdjustRightInd">
    <w:adjustRightInd>
      <xsl:attribute name="w:val">
        <xsl:value-of select="./字:是否有网格自动调整右缩进_419A"/>
      </xsl:attribute>
    </w:adjustRightInd>
  </xsl:template>

  <xsl:template name="InsertSnapToGrid">
    <w:snapToGrid>
      <xsl:attribute name="w:val">
        <xsl:value-of select="./字:是否对齐网格_4190"/>
      </xsl:attribute>
    </w:snapToGrid>
  </xsl:template>


  <xsl:template name="InsertLineSpacing">
    <xsl:choose>
      <xsl:when test="./字:行距_417E/@类型_417F='at-least'">
        <xsl:attribute name="w:lineRule">atLeast</xsl:attribute>
        <xsl:attribute name="w:line">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="./字:行距_417E/@值_4108"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="./字:行距_417E/@类型_417F='fixed'">
        <xsl:attribute name="w:lineRule">exact</xsl:attribute>
        <xsl:attribute name="w:line">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="./字:行距_417E/@值_4108"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="./字:行距_417E/@类型_417F='multi-lines'">
        <xsl:attribute name="w:lineRule">auto</xsl:attribute>
        <xsl:attribute name="w:line">
          <xsl:value-of select="round(number(./字:行距_417E/@值_4108 * 240))"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="./字:行距_417E/@类型_417F='line-space'">
        <xsl:attribute name="w:lineRule">exact</xsl:attribute>
        <xsl:attribute name="w:line">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="./字:行距_417E/@值_4108"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="InsertParaSpacing">
    <xsl:if test="./字:段间距_4180/字:段后距_4185/字:自动_4182">
      <xsl:attribute name="w:afterAutospacing">on</xsl:attribute>
      <xsl:attribute name="w:after">100</xsl:attribute>
    </xsl:if>
    <xsl:if test="./字:段间距_4180/字:段后距_4185/字:绝对值_4183">
      <xsl:attribute name="w:after">
        <xsl:call-template name="twipsMeasure">
          <xsl:with-param name="lengthVal" select="./字:段间距_4180/字:段后距_4185/字:绝对值_4183"/>
        </xsl:call-template>
      </xsl:attribute>
      <!--不加照理说也可以，但有时候会出状况-->
      <xsl:attribute name="w:afterAutospacing">
        <xsl:value-of select="0"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="./字:段间距_4180/字:段后距_4185/字:相对值_4184">
      <xsl:attribute name="w:afterLines">
        <xsl:value-of select="round(number(./字:段间距_4180/字:段后距_4185/字:相对值_4184 * 100))"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="./字:段间距_4180/字:段前距_4181/字:自动_4182">
      <xsl:attribute name="w:beforeAutospacing">on</xsl:attribute>
      <xsl:attribute name="w:before">100</xsl:attribute>
    </xsl:if>
    <xsl:if test="./字:段间距_4180/字:段前距_4181/字:绝对值_4183">
      <xsl:attribute name="w:before">
        <xsl:call-template name="twipsMeasure">
          <xsl:with-param name="lengthVal" select="./字:段间距_4180/字:段前距_4181/字:绝对值_4183"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="w:beforeAutospacing">
        <xsl:value-of select="0"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="./字:段间距_4180/字:段前距_4181/字:相对值_4184">
      <xsl:attribute name="w:beforeLines">
        <xsl:value-of select="round(number(./字:段间距_4180/字:段前距_4181/字:相对值_4184 * 100))"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
   
  <xsl:template match="字:缩进_411D">
    <w:ind>
      <xsl:if test="(preceding-sibling::字:自动编号信息_4186 and not(./字:首行_4111/字:相对_4109) and not(ancestor::式样:段落式样_9912))">
        <xsl:attribute name="w:firstLineChars">
          <xsl:value-of select="0"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if
        test="ancestor::字:段落属性_419B/字:自动编号信息_4186 and ((./字:首行_4111/字:绝对_4107/@值_410F = 0) or (./字:首行_4111/字:绝对_4107/@值_410F = 0))">
        <xsl:attribute name="w:firstLine">
          <xsl:value-of select="0"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:variable name="hangingVal">
        <xsl:choose>
          <xsl:when test="./字:首行_4111/字:绝对_4107/@值_410F &lt; 0">
            <xsl:value-of select="-(./字:首行_4111/字:绝对_4107/@值_410F)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'0'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:if test="(./字:首行_4111/字:绝对_4107/@值_410F &lt; 0) and not(字:左_410E)">
        <xsl:attribute name="w:left">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="$hangingVal"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:for-each select="字:首行_4111 | 字:左_410E | 字:右_4110">
        <xsl:choose>
          <xsl:when test="name(.)='字:首行_4111'">
            <xsl:call-template name="FirstLineOrHanging"/>
          </xsl:when>
          <xsl:when test="name(.)='字:左_410E'">
            <xsl:call-template name="ComputeLeftInd">
              <xsl:with-param name="hangingValParam" select="$hangingVal"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name(.)='字:右_4110'">
            <xsl:call-template name="ComputeRightInd"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </w:ind>
  </xsl:template>

  <xsl:template name="ComputeLeftInd">
    <xsl:param name="hangingValParam"/>
    <xsl:if test="./字:绝对_4107/@值_410F ">
      <xsl:attribute name="w:left">
        <xsl:call-template name="twipsMeasure">
          <xsl:with-param name="lengthVal" select="(./字:绝对_4107/@值_410F)+$hangingValParam"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="./字:相对_4109/@值_4108 ">
      <xsl:attribute name="w:leftChars">
        <xsl:value-of select="round(number(./字:相对_4109/@值_4108 * 100))"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
  <xsl:template name="ComputeRightInd">
    <xsl:if test="./字:绝对_4107/@值_410F ">
      <xsl:attribute name="w:right">
        <xsl:call-template name="twipsMeasure">
          <xsl:with-param name="lengthVal" select="./字:绝对_4107/@值_410F"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="./字:相对_4109/@值_4108 ">
      <xsl:attribute name="w:rightChars">
        <xsl:value-of select="round(number(./字:相对_4109/@值_4108 * 100))"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template name="FirstLineOrHanging">
    
    <!--2014-03-26，wudi，悬挂或首行缩进，考虑0值的情况，start-->
    <xsl:if test="./字:绝对_4107/@值_410F &gt;= 0">
      <xsl:attribute name="w:firstLine">
        <xsl:call-template name="twipsMeasure">
          <xsl:with-param name="lengthVal" select="./字:绝对_4107/@值_410F"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <!--end-->
    
    <xsl:if test="./字:绝对_4107/@值_410F &lt; 0">
      <xsl:attribute name="w:hanging">
        <xsl:call-template name="twipsMeasure">
          <xsl:with-param name="lengthVal" select="-(./字:绝对_4107/@值_410F)"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <!--2014-03-26，wudi，悬挂或首行缩进，考虑0值的情况，start-->
    <xsl:if test="./字:相对_4109/@值_4108 &gt;= 0">
      <xsl:attribute name="w:firstLineChars">
        <xsl:value-of select="round(number(./字:相对_4109/@值_4108 * 100))"/>
      </xsl:attribute>
    </xsl:if>
    <!--end-->
    
    <xsl:if test="./字:相对_4109/@值_4108 &lt; 0">
      <xsl:attribute name="w:hangingChars">
        <xsl:value-of select="round(-(number(./字:相对_4109/@值_4108 * 100)))"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertJcAndTextAlign">
    <xsl:if test="./字:对齐_417D/@水平对齐_421D">
      <w:jc>
        <xsl:choose>
          <xsl:when test="./字:对齐_417D/@水平对齐_421D='center'">
            <xsl:attribute name="w:val">center</xsl:attribute>
          </xsl:when>
          <xsl:when test="./字:对齐_417D/@水平对齐_421D='left'">
            <xsl:attribute name="w:val">left</xsl:attribute>
          </xsl:when>
          <xsl:when test="./字:对齐_417D/@水平对齐_421D='right'">
            <xsl:attribute name="w:val">right</xsl:attribute>
          </xsl:when>
          <xsl:when test="./字:对齐_417D/@水平对齐_421D='justified'">
            <xsl:attribute name="w:val">both</xsl:attribute>
          </xsl:when>
          <xsl:when test="./字:对齐_417D/@水平对齐_421D='distributed'">
            <xsl:attribute name="w:val">distribute</xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </w:jc>
    </xsl:if>
    <xsl:if test="./字:对齐_417D/@文字对齐_421E">
      <w:textAlignment>
        <xsl:choose>
          <xsl:when test="./字:对齐_417D/@文字对齐_421E='bottom'">
            <xsl:attribute name="w:val">bottom</xsl:attribute>
          </xsl:when>
          <xsl:when test="./字:对齐_417D/@文字对齐_421E='top'">
            <xsl:attribute name="w:val">top</xsl:attribute>
          </xsl:when>
          <xsl:when test="./字:对齐_417D/@文字对齐_421E='center'">
            <xsl:attribute name="w:val">center</xsl:attribute>
          </xsl:when>
          <xsl:when test="./字:对齐_417D/@文字对齐_421E='base'">
            <xsl:attribute name="w:val">baseline</xsl:attribute>
          </xsl:when>
          <xsl:when test="./字:对齐_417D/@文字对齐_421E='auto'">
            <xsl:attribute name="w:val">auto</xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </w:textAlignment>
    </xsl:if>
  </xsl:template>

  <xsl:template name="InsertOutlineLvl">
    
    <!--2015-03-19，wudi，取值不能为负值，否则会导致转后文档需要修复，start-->
    <xsl:if test="round(./字:大纲级别_417C) != 0">
      <w:outlineLvl>
        <xsl:attribute name="w:val">
          <xsl:value-of select="round(./字:大纲级别_417C - 1)"/>
        </xsl:attribute>
      </w:outlineLvl>
    </xsl:if>
    <!--end-->
    
  </xsl:template>

  <xsl:template name="InsertKeepLines">
    <w:keepLines>
      <xsl:attribute name="w:val">
        <xsl:value-of select="./字:是否段中不分页_418C"/>
      </xsl:attribute>
    </w:keepLines>
  </xsl:template>

  <xsl:template match="字:分节_416A">
    <xsl:call-template name="section"/>
  </xsl:template>
</xsl:stylesheet>
