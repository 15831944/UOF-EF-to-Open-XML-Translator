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
  xmlns:fn="http://www.w3.org/2005/xpath-functions"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
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
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
  <!--yx,add,2010.5.12-->
  <xsl:import href="revisions.xsl"/>
  <xsl:import href="styles.xsl"/>
  <xsl:import href="region.xsl"/>
  <xsl:import href="common.xsl"/>
  <xsl:import href="graphics.xsl"/>
  <xsl:import href="bookmark.xsl"/>
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
  <xsl:variable name="styleId"
    select="document('word/styles.xml')/w:styles/w:style[(@w:type='paragraph') and (@w:default='1'or @w:default='on' or @w:default='true')]/@w:styleId">
  </xsl:variable>

  <xsl:template match="w:r" mode="smartTag"><!--智能标签smartTag模板-->
    <xsl:param name="pPartFrom"/>
    <字:句_419D>
      <xsl:call-template name="run">
        <xsl:with-param name="rPartFrom" select="$pPartFrom"/>
      </xsl:call-template>
    </字:句_419D>
  </xsl:template>
  
  <!--paragraph(w:p)模板11.4日cxl修改*************************************************************************-->
  <xsl:template name="paragraph">
    <xsl:param name="pPartFrom"/>
    
    <!--2014-05-05，wudi，增加postr参数，记录单元格所属行数，修复单元格字体加粗效果转换BUG，start-->
    <xsl:param name="postr"/>
    <xsl:param name="postc"/>
    <!--end-->

    <!--2014-12-02，wudi，注释掉以下代码，解决性能问题，后期调整，start-->
    
    <!--2012-12-17，wudi，OOX到UOF方向的书签的实现，start--><!--
    <xsl:variable name ="bookStartNum">
      <xsl:value-of select="count(preceding-sibling::w:bookmarkStart)"/>
    </xsl:variable>
    <xsl:variable name="bookEndNum1">
      <xsl:value-of select ="count(preceding-sibling::w:bookmarkEnd)"/>
    </xsl:variable>

    --><!--2014-03-27，wudi，针对书签，批注转换出现的特例，新增加的变量，start--><!--
    <xsl:variable name="bookStartNum1">
      <xsl:value-of select ="count(ancestor::w:tc/preceding-sibling::w:bookmarkStart)"/>
    </xsl:variable>
    <xsl:variable name="commentStartNum">
      <xsl:value-of select="count(preceding-sibling::w:commentRangeStart)"/>
    </xsl:variable>
    <xsl:variable name="commentStartNum1">
      <xsl:value-of select="count(ancestor::w:tc/preceding-sibling::w:commentRangeStart)"/>
    </xsl:variable>
    <xsl:variable name="commentStartNum2">
      <xsl:value-of select="count(ancestor::w:sdt/preceding-sibling::w:commentRangeStart)"/>
    </xsl:variable>
    --><!--end--><!--

    --><!--
    <xsl:variable name="bookEndNum2">
      <xsl:value-of select ="count(following::w:bookmarkEnd)"/>
    </xsl:variable>
    --><!--
    --><!--end-->
    
    <字:段落_416B><!--有attrList="标识符_4169">-->

      <!--
      <preceding>
        <xsl:value-of select ="name(preceding-sibling::*[1])"/>
      </preceding>
      <following>
        <xsl:value-of select ="name(following::*[1])"/>
      </following>
      -->
      
      <!--2012-12-17，wudi，OOX到UOF方向的书签的实现，start--><!--
      <xsl:if test ="($bookStartNum > 0) and (name(preceding-sibling::*[1]) = 'w:bookmarkStart' or (name(preceding-sibling::*[2]) = 'w:bookmarkStart' and name(preceding-sibling::*[1]) = 'w:bookmarkEnd'))">
        <字:句_419D>
          <字:区域开始_4165>
            <xsl:attribute name="标识符_4100">
              <xsl:value-of select= "concat('bk_',preceding-sibling::w:bookmarkStart[1]/@w:id)"/>
            </xsl:attribute>
            <xsl:attribute name="名称_4166">
              <xsl:value-of select="preceding-sibling::w:bookmarkStart[1]/@w:name"/>
            </xsl:attribute>
            <xsl:attribute name="类型_413B">
              <xsl:value-of select="'bookmark'"/>
            </xsl:attribute>
          </字:区域开始_4165>
        </字:句_419D> 
      </xsl:if>

      --><!--2014-03-27，wudi，书签转换出现新的特例，不符合UOF标准，start--><!--

      <xsl:if test ="($bookStartNum1 > 0) and (name(ancestor::w:tc/preceding-sibling::*[1]) = 'w:bookmarkStart' or (name(ancestor::w:tc/preceding-sibling::*[2]) = 'w:bookmarkStart' and name(ancestor::w:tc/preceding-sibling::*[1]) = 'w:commentRangeStart'))">
        <字:句_419D>
          <字:区域开始_4165>
            <xsl:attribute name="标识符_4100">
              <xsl:value-of select= "concat('bk_',ancestor::w:tc/preceding-sibling::w:bookmarkStart[1]/@w:id)"/>
            </xsl:attribute>
            <xsl:attribute name="名称_4166">
              <xsl:value-of select="ancestor::w:tc/preceding-sibling::w:bookmarkStart[1]/@w:name"/>
            </xsl:attribute>
            <xsl:attribute name="类型_413B">
              <xsl:value-of select="'bookmark'"/>
            </xsl:attribute>
          </字:区域开始_4165>
        </字:句_419D>
      </xsl:if>

      <xsl:if test ="($bookStartNum >0) and (name(preceding-sibling::*[2])='w:bookmarkStart' and name(preceding-sibling::*[1])='w:commentRangeStart')">
        <字:句_419D>
          <字:区域开始_4165>
            <xsl:attribute name="标识符_4100">
              <xsl:value-of select= "concat('bk_',preceding-sibling::w:bookmarkStart[1]/@w:id)"/>
            </xsl:attribute>
            <xsl:attribute name="名称_4166">
              <xsl:value-of select="preceding-sibling::w:bookmarkStart[1]/@w:name"/>
            </xsl:attribute>
            <xsl:attribute name="类型_413B">
              <xsl:value-of select="'bookmark'"/>
            </xsl:attribute>
          </字:区域开始_4165>
        </字:句_419D>
      </xsl:if>

      <xsl:if test ="($bookStartNum >0) and (name(preceding-sibling::*[3])='w:bookmarkStart' and name(preceding-sibling::*[2])='w:commentRangeStart' and name(preceding-sibling::*[1])='w:commentRangeStart')">
        <字:句_419D>
          <字:区域开始_4165>
            <xsl:attribute name="标识符_4100">
              <xsl:value-of select= "concat('bk_',preceding-sibling::w:bookmarkStart[1]/@w:id)"/>
            </xsl:attribute>
            <xsl:attribute name="名称_4166">
              <xsl:value-of select="preceding-sibling::w:bookmarkStart[1]/@w:name"/>
            </xsl:attribute>
            <xsl:attribute name="类型_413B">
              <xsl:value-of select="'bookmark'"/>
            </xsl:attribute>
          </字:区域开始_4165>
        </字:句_419D>
      </xsl:if>

      <xsl:if test ="($bookEndNum1 > 0) and (name(preceding-sibling::*[1]) = 'w:bookmarkEnd' or (name(preceding-sibling::*[2]) = 'w:bookmarkEnd' and name(preceding-sibling::*[1]) = 'w:commentRangeStart'))">
        <字:句_419D>
          <字:区域结束_4167>
            <xsl:attribute name="标识符引用_4168">
              <xsl:value-of select= "concat('bk_',preceding-sibling::w:bookmarkEnd[1]/@w:id)"/>
            </xsl:attribute>
          </字:区域结束_4167>
        </字:句_419D>
      </xsl:if>
      --><!--end--><!--

      --><!--end--><!--

      --><!--2014-03-27，wudi，批注转换出现新的特例，不符合UOF标准，start--><!--

      <xsl:if test ="($commentStartNum2 > 0) and (name(ancestor::w:sdt/preceding-sibling::*[1]) = 'w:commentRangeStart')">
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
      --><!--end-->
      <!--end-->
      
      
      <字:段落属性_419B><!--有attrList="式样引用_419C">-->

        <!--2014-04-22，wudi，修复表格单元格行距转换BUG，start-->
        <xsl:if test="./w:pPr">
          <xsl:if test="not(./w:pPr/w:pStyle)">
            <xsl:call-template name="ParagraphStyle"/>
          </xsl:if>
          <xsl:if test="not(./w:pPr/w:pPrChange)">
            <xsl:apply-templates select="./w:pPr"/>
            <!--cxl,2012.4.11以下代码为按WPS处理时用-->
            <!--<xsl:if test="not(./w:pPr/w:framePr)">
            <xsl:apply-templates select="./w:pPr"/>
          </xsl:if>
          <xsl:if test="./w:pPr/w:framePr and ./w:pPr[2]">
            <xsl:apply-templates select="./w:pPr[2]">
              <xsl:with-param name="framePrLocation" select="'2'"/>
            </xsl:apply-templates>
          </xsl:if>
          <xsl:if test="./w:pPr/w:framePr and count(./w:pPr)=1">
            <xsl:call-template name="framePr"/>
          </xsl:if>-->
          </xsl:if>
        </xsl:if>

        <xsl:if test ="not(./w:pPr)">
          <xsl:if test ="ancestor::w:tbl/w:tblPr/w:tblStyle">
            <xsl:variable name="styId">
              <xsl:value-of select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
            </xsl:variable>
            <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId=$styId]/w:pPr">
              <xsl:apply-templates select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$styId]/w:pPr"/>
            </xsl:if>
          </xsl:if>
        </xsl:if>
        <!--end-->
        
      </字:段落属性_419B>
      
      <!--<xsl:if test="./w:pPr/w:pPrChange">--><!--Revision Information for Paragraph Properties-->
		  <!--<xsl:apply-templates select="./w:pPr" mode="RevisonPr">
			  <xsl:with-param name="lx" select="'format'"/>
			  <xsl:with-param name="revPartFrom" select="$pPartFrom"/>
		  </xsl:apply-templates>-->
		  <!--<xsl:apply-templates select="./w:pPr/w:pPrChange/w:pPr"/>-->       
      <!--</xsl:if>-->
      
      <xsl:for-each select="node()">
        <xsl:choose>
          <xsl:when test="name(.)='w:pPr'">
			      <xsl:for-each select="w:pPrChange">
              <xsl:call-template name="pPrchangeBody">
                <xsl:with-param name="lx" select="'format'"/>
                <xsl:with-param name="revPartFrom" select="$pPartFrom"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:when>
          <!--cxl2011/12/15添加句判断条件，区分w:r的子结点为域w:fldChar的情况-->
          <xsl:when test="name(.)='w:r'">
            <xsl:if test="not(./w:fldChar) and not(./w:instrText) and not(contains(preceding-sibling::w:r/w:instrText,'HYPERLINK'))">
            <!--cxl,2012.5.14删除这个条件 and not(preceding-sibling::w:r/w:fldChar/@w:fldCharType='separate' or preceding-sibling::w:r/w:fldChar/@w:fldCharType='begin')-->
              <字:句_419D>
                <xsl:call-template name="run">
                  <xsl:with-param name="rPartFrom" select="$pPartFrom"/>

                  <!--2014-05-05，wudi，增加postr参数，记录单元格所属行数，修复单元格字体加粗效果转换BUG，start-->
                  <xsl:with-param name="postr" select="$postr"/>
                  <xsl:with-param name="postc" select="$postc"/>
                  <!--end-->
                  
                </xsl:call-template>
              </字:句_419D>
            </xsl:if>
            <!--CXL,2012.3.6,双行合一的话需要在句之后添加一个句，否则效果不正确-->
            <xsl:if test="./w:rPr/w:eastAsianLayout[@w:combine='1']">
              <字:句_419D>
                <字:句属性_4158/>
              </字:句_419D>
            </xsl:if>
            
            <!--2013-03-12，wudi，修复文字表批注BUG，start-->
            <xsl:if test ="./w:commentReference">
              <xsl:variable name ="commentId">
                <xsl:value-of select ="./w:commentReference/@w:id"/>
              </xsl:variable>
              <xsl:if test ="(ancestor::w:tr) and not(preceding::w:commentRangeEnd[@w:id=$commentId])">
                <xsl:if test ="not(preceding::w:commentRangeStart[@w:id=$commentId])">
                  <字:句_419D>
                    <字:区域开始_4165>
                      <xsl:attribute name="标识符_4100">
                        <xsl:value-of select="concat('cmt_',$commentId)"/>
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
                <字:句_419D>
                  <字:区域结束_4167>
                    <xsl:attribute name="标识符引用_4168">
                      <xsl:value-of select="concat('cmt_',$commentId)"/>
                    </xsl:attribute>
                  </字:区域结束_4167>
                </字:句_419D>
              </xsl:if>
            </xsl:if>
            <!--end-->
            
            <!--cxl2011/12/15-->
            <!--2013-03-20，wudi，删除if判断条件./w:instrText-->
            <xsl:if test="./w:fldChar">

              <!--2013-03-22，wudi，修复OOX到UOF方向页码转换的BUG，第几页和总共的页数统计出错，start-->
              <!--2014-05-07，wudi，增加条件./ancestor::w:hdr，考虑页码在页眉的情况，start-->
              <xsl:if test ="./w:fldChar/@w:fldCharType='begin' and (./ancestor::w:ftr or ./ancestor::w:hdr)">
                <xsl:variable name ="tmp">
                  <xsl:value-of select ="following-sibling::w:r/w:instrText"/>
                </xsl:variable>
                <xsl:variable name ="temp2">
                  <xsl:value-of select ="normalize-space($tmp)"/>
                </xsl:variable>
                <!--<temp2>
                  <xsl:value-of select ="$temp2"/>
                </temp2>-->
                <xsl:variable name ="type2">
                  <xsl:choose>
                    <xsl:when test ="starts-with($temp2,'PAGE')">
                      <xsl:value-of select ="'PAGE'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp2,'NUMPAGES')">
                      <xsl:value-of select="'NUMPAGES'"/>
                    </xsl:when>

                    <!--2014-03-26，wudi，增加域DATE，FILENAME，PRINTDATE的转换，start-->
                    <xsl:when test="starts-with($temp2,'DATE')">
                      <xsl:value-of select="'DATE'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp2,'FILENAME')">
                      <xsl:value-of select="'FILENAME'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp2,'PRINTDATE')">
                      <xsl:value-of select="'PRINTDATE'"/>
                    </xsl:when>
                    <!--end-->
                    
                  </xsl:choose>
                </xsl:variable>
                <字:域开始_419E>
                  <xsl:attribute name="类型_416E">
                    <xsl:value-of select ="translate($type2,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                  </xsl:attribute>
                  <xsl:if test="@fldLock='true'">
                    <xsl:attribute name="是否锁定_416F">
                      <xsl:value-of select ="'true'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:if test ="not(@fldLock) or @fldLock='false'">
                    <xsl:attribute name="是否锁定_416F">
                      <xsl:value-of select ="'false'"/>
                    </xsl:attribute>
                  </xsl:if>
                </字:域开始_419E>
                <xsl:choose>
                  <xsl:when test ="$type2 ='PAGE'">
                    <字:域代码_419F>
                      <字:段落_416B>
                        <字:句_419D>
                          <字:句属性_4158/>
                          <字:文本串_415B>
                            <xsl:value-of select ="'Page'"/>
                          </字:文本串_415B>
                        </字:句_419D>
                      </字:段落_416B>
                    </字:域代码_419F>
                  </xsl:when>
                  <xsl:when test ="$type2 ='NUMPAGES'">
                    <字:域代码_419F>
                      <字:段落_416B>
                        <字:句_419D>
                          <字:句属性_4158/>
                          <字:文本串_415B>
                            <xsl:value-of select ="'NumPages'"/>
                          </字:文本串_415B>
                        </字:句_419D>
                      </字:段落_416B>
                    </字:域代码_419F>
                  </xsl:when>

                  <!--2014-03-26，wudi，增加域DATE，FILENAME，PRINTDATE的转换，start-->
                  <xsl:when test ="$type2 ='DATE'">
                    <字:域代码_419F>
                      <字:段落_416B>
                        <字:句_419D>
                          <字:句属性_4158/>
                          <字:文本串_415B>
                            <xsl:value-of select ="$tmp"/>
                          </字:文本串_415B>
                        </字:句_419D>
                      </字:段落_416B>
                    </字:域代码_419F>
                  </xsl:when>
                  <xsl:when test ="$type2 ='FILENAME'">
                    <字:域代码_419F>
                      <字:段落_416B>
                        <字:句_419D>
                          <字:句属性_4158/>
                          <字:文本串_415B>
                            <xsl:value-of select ="'FileName'"/>
                          </字:文本串_415B>
                        </字:句_419D>
                      </字:段落_416B>
                    </字:域代码_419F>
                  </xsl:when>
                  <xsl:when test ="$type2 ='PRINTDATE'">
                    <字:域代码_419F>
                      <字:段落_416B>
                        <字:句_419D>
                          <字:句属性_4158/>
                          <字:文本串_415B>
                            <xsl:value-of select ="$tmp"/>
                          </字:文本串_415B>
                        </字:句_419D>
                      </字:段落_416B>
                    </字:域代码_419F>
                  </xsl:when>
                  <!--end-->
                  
                </xsl:choose>
              </xsl:if>
              <!--end-->
              <!--end-->
              
              <!--2013-04-11，wudi，加限制条件，MACROBUTTON为本期不转的域，暂时屏蔽它-->
              <!--2014-03-27，wudi，加限制条件，FORMTEXT为本期不转的域，暂时屏蔽它-->
              <!--2014-05-07，wudi，增加条件./ancestor::w:hdr，考虑页码在页眉的情况，start-->
              <xsl:if test="./w:fldChar/@w:fldCharType='begin' and not(./ancestor::w:ftr or ./ancestor::w:hdr) and not(contains(following-sibling::w:r/w:instrText,'MACROBUTTON')) and not(contains(preceding-sibling::w:r/w:instrText,'FORMTEXT'))">
                <!--and(contains(./following::*/w:instrText[1],'FILENAME'))">-->
                
                <!--2013-04-15，wudi，修复索引转换BUG，处理INDEX，TOC，PAGEREF，start-->
                <xsl:variable name="temp1">
                    <xsl:for-each select="./ancestor::w:p//w:instrText">
                      <xsl:value-of select ="."/>
                    </xsl:for-each>
                </xsl:variable>
                <xsl:variable name ="temp2">
                  <xsl:value-of select ="normalize-space($temp1)"/>
                </xsl:variable>
                <xsl:variable name ="temp">
                  <xsl:choose>
                    <xsl:when test ="contains($temp2,'PageRef') and starts-with($temp2,'TOC') and contains(following-sibling::w:r[1]/w:instrText,'PageRef')">
                      <xsl:value-of select ="concat('PageRef',substring-after($temp2,'PageRef'))"/>
                    </xsl:when>
                    <xsl:when test ="contains($temp2,'PAGEREF') and starts-with($temp2,'TOC') and contains(following-sibling::w:r[1]/w:instrText,'PAGEREF')">
                      <xsl:value-of select ="concat('PAGEREF',substring-after($temp2,'PAGEREF'))"/>
                    </xsl:when>
                    <xsl:when test ="contains($temp2,'PageRef') and starts-with($temp2,'INDEX') and contains(following-sibling::w:r[1]/w:instrText,'PageRef')">
                      <xsl:value-of select ="concat('PageRef',substring-after($temp2,'PageRef'))"/>
                    </xsl:when>
                    <xsl:when test ="contains($temp2,'PAGEREF') and starts-with($temp2,'INDEX') and contains(following-sibling::w:r[1]/w:instrText,'PAGEREF')">
                      <xsl:value-of select ="concat('PAGEREF',substring-after($temp2,'PAGEREF'))"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select ="$temp2"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>
                <!--end-->
                
                <xsl:variable name="type">
                  <xsl:choose>
                    <!--<xsl:when test="contains($temp,' ')">
                      <xsl:value-of select="substring-before($temp,' ')"/>
                    </xsl:when>-->
                    <!--CXL,2012.3.17-->
                    <xsl:when test="starts-with($temp,'eq')">
                      <xsl:value-of select="'eq'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'AUTHOR')">
                      <xsl:value-of select="'AUTHOR'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'DATE')">
                      <xsl:value-of select="'DATE'"/>
                    </xsl:when>
                    
                    <!--2013-03-20，wudi，修复#2738集成测试OOX到UOF2.0模板8目录出现差异BUG，start-->
                    <xsl:when test ="starts-with($temp,'TOC')">
                      <xsl:value-of select ="'TOC'"/>
                    </xsl:when>
                    <!--end-->

                    <!--2013-04-15，wudi，修复索引转换BUG，处理INDEX，TOC，PAGEREF，start-->
                    <xsl:when test ="starts-with($temp,'INDEX')">
                      <xsl:value-of select ="'INDEX'"/>
                    </xsl:when>

                    <xsl:when test ="starts-with($temp,'PageRef')">
                      <xsl:value-of select ="'PAGEREF'"/>
                    </xsl:when>

                    <xsl:when test ="starts-with($temp,'PAGEREF')">
                      <xsl:value-of select ="'PAGEREF'"/>
                    </xsl:when>
                    <!--end-->
                    
                    <xsl:when test="starts-with($temp,'FILENAME')">
                      <xsl:value-of select="'FILENAME'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'SECTIONPAGES')">
                      <xsl:value-of select="'SECTIONPAGES'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'REVNUM')">
                      <xsl:value-of select="'REVNUM'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'TIME')">
                      <xsl:value-of select="'TIME'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'PAGE')">
                      <xsl:value-of select="'PAGE'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'SECTION')">
                      <xsl:value-of select="'SECTION'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'REF')">
                      <xsl:value-of select="'REF'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'XE')">
                      <xsl:value-of select="'XE'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'SEQ')">
                      <xsl:value-of select="'SEQ'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'SAVEDATE')">
                      <xsl:value-of select="'SAVEDATE'"/>
                    </xsl:when>
                    <xsl:when test="contains($temp,'TITLE')">
                      <xsl:value-of select="'TITLE'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'CREATEDATE')">
                      <xsl:value-of select="'CREATEDATE'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'NUMPAGES')">
                      <xsl:value-of select="'NUMPAGES'"/>
                    </xsl:when>
                    <xsl:when test="starts-with($temp,'NUMCHARS')">
                      <xsl:value-of select="'NUMCHARS'"/>
                    </xsl:when>
                    <xsl:when test="contains($temp,'HYPERLINK')">
                      <xsl:value-of select="'HYPERLINK'"/>
                    </xsl:when>                  
                    <!--<xsl:otherwise>
                      <xsl:value-of select="$temp"/>
                    </xsl:otherwise>-->
                  </xsl:choose>
                </xsl:variable>
                <xsl:choose>
                  <xsl:when
                    test="$type='AUTHOR' or $type='DATE' or $type='FILENAME' or $type='SECTIONPAGES' or $type='REVNUM' or $type='TIME' or $type='PAGE' or $type='SECTION' or $type='REF' or $type='XE' or $type='SEQ' or $type='TITLE' or $type='SAVEDATE'or $type='CREATEDATE'or $type='NUMPAGES' or $type='NUMCHARS'">
                    <xsl:call-template name="SimpleField">
                      <xsl:with-param name="type"
                        select="translate($type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                      <xsl:with-param name="splfldPartFrom" select="$pPartFrom"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$type='HYPERLINK'">
                    <xsl:call-template name="hyperlinkRegion">
                      <xsl:with-param name="filename" select="$pPartFrom"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$type='eq'">
                    <xsl:call-template name="EQfield">
                      <xsl:with-param name="type"
                        select="translate($type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                      <xsl:with-param name="splfldPartFrom" select="$pPartFrom"/>
                    </xsl:call-template>
                  </xsl:when>
                  
                  <!--2013-03-20，wudi，修复#2738集成测试OOX到UOF2.0模板8目录出现差异BUG，start-->
                  <xsl:when test ="$type='TOC'">
                    <字:域开始_419E 类型_416E="toc" 是否锁定_416F="false"/>
                    <字:域代码_419F>
                      <字:段落_416B>
                        <字:句_419D>
                          <字:句属性_4158/>
                          <字:文本串_415B>
                            
                            <!--2013-04-15，wudi，处理temp带‘PageRef’或‘PAGEREF’的情况，start-->
                            <xsl:choose>
                              <xsl:when test ="contains($temp,'PageRef')">
                                <xsl:value-of select ="substring-before($temp,'PageRef')"/>
                              </xsl:when>
                              <xsl:when test ="contains($temp,'PAGEREF')">
                                <xsl:value-of select ="substring-before($temp,'PAGEREF')"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select ="$temp"/>
                              </xsl:otherwise>
                            </xsl:choose>
                            <!--end-->
                            
                          </字:文本串_415B>
                        </字:句_419D>
                      </字:段落_416B>
                    </字:域代码_419F>
                  </xsl:when>
                  <!--end-->
                  
                  <!--2013-04-15，wudi，修复索引转换BUG，增加对INDEX的处理，start-->
                  <xsl:when test ="$type='INDEX'">
                    <字:域开始_419E 类型_416E="index" 是否锁定_416F="false"/>
                    <字:域代码_419F>
                      <字:段落_416B>
                        <字:句_419D>
                          <字:句属性_4158/>
                          <字:文本串_415B>
                            
                            <!--2013-04-15，wudi，处理temp带‘PageRef’或‘PAGEREF’的情况，start-->
                            <xsl:choose>
                              <xsl:when test ="contains($temp,'PageRef')">
                                <xsl:value-of select ="substring-before($temp,'PageRef')"/>
                              </xsl:when>
                              <xsl:when test ="contains($temp,'PAGEREF')">
                                <xsl:value-of select ="substring-before($temp,'PAGEREF')"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select ="$temp"/>
                              </xsl:otherwise>
                            </xsl:choose>
                            <!--end-->
                            
                          </字:文本串_415B>
                        </字:句_419D>
                      </字:段落_416B>
                    </字:域代码_419F>
                  </xsl:when>
                  <!--end-->
                  
                  <!--2013-04-15，wudi，修复目录与索引转换BUG，增加对PAGEREF的处理，start-->
                  <xsl:when test ="$type='PAGEREF'">
                    <字:域开始_419E 类型_416E="pageref" 是否锁定_416F="false"/>
                    <字:域代码_419F>
                      <字:段落_416B>
                        <字:句_419D>
                          <字:句属性_4158/>
                          <字:文本串_415B>
                            <xsl:value-of select ="$temp"/>
                          </字:文本串_415B>
                        </字:句_419D>
                      </字:段落_416B>
                    </字:域代码_419F>
                  </xsl:when>
                  <!--end-->
                  
                  <!--<xsl:when test="$type='REVNUM'">
                    <xsl:call-template name="SimpleField">
                      <xsl:with-param name="type" select="'revision'"/>
                      <xsl:with-param name="splfldPartFrom" select="$pPartFrom"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$type='SECTIONPAGES'">
                    <xsl:call-template name="SimpleField">
                      <xsl:with-param name="type" select="'pageinsection'"/>
                      <xsl:with-param name="splfldPartFrom" select="$pPartFrom"/>
                    </xsl:call-template>
                  </xsl:when>-->
                </xsl:choose>
              </xsl:if>
              <!--end-->
              
              <!--2013-04-11，wudi，加限制条件，MACROBUTTON为本期不转的域，暂时屏蔽它-->
              <!--2014-03-27，wudi，加限制条件，FORMTEXT为本期不转的域，暂时屏蔽它-->
              <xsl:if test="./w:fldChar/@w:fldCharType='end' and not(contains(preceding-sibling::w:r/w:instrText,'MACROBUTTON')) and not(contains(preceding-sibling::w:r/w:instrText,'FORMTEXT'))">
                <字:域结束_41A0/>
              </xsl:if>
            </xsl:if>
          </xsl:when>

          <!--2015-06-02，wudi，增加公式转换锚点，start-->
          <xsl:when test="name(.)='m:oMath'">
            <xsl:variable name="number">
              <xsl:number format="1" level="any" count="m:oMath | m:oMathPara"/>
            </xsl:variable>
            <字:句_419D>
              <字:句属性_4158/>
              <uof:锚点_C644>
                <xsl:attribute name="图形引用_C62E">
                  <xsl:value-of select="concat('equation_',$number)"/>
                </xsl:attribute>
              </uof:锚点_C644>
            </字:句_419D>
          </xsl:when>

          <xsl:when test=".//m:oMath">
            <xsl:variable name="number">
              <xsl:number format="1" level="any" count="m:oMath | m:oMathPara"/>
            </xsl:variable>
            <字:句_419D>
              <字:句属性_4158/>
              <uof:锚点_C644>
                <xsl:attribute name="图形引用_C62E">
                  <xsl:value-of select="concat('equation_',$number + 1)"/>
                </xsl:attribute>
              </uof:锚点_C644>
            </字:句_419D>
          </xsl:when>
          <!--end-->
          
          <!--2012-11-22，wudi，OOX到UOF方向的书签实现，start-->
          <!--
          <xsl:when test="name(.)='w:bookmarkStart'">
           <字:句_419D>
            <字:区域开始_4165>
              <xsl:attribute name="标识符_4100">
                <xsl:value-of select= "concat('bk_',@w:id)"/>
              </xsl:attribute>
              <xsl:attribute name="名称_4166">
                <xsl:value-of select="@w:name"/>
              </xsl:attribute>
              <xsl:attribute name="类型_413B">
                <xsl:value-of select="'bookmark'"/>
              </xsl:attribute>
            </字:区域开始_4165>
           </字:句_419D>
          </xsl:when>
          <xsl:when test="name(.)='w:bookmarkEnd'">
           <字:句_419D>
            <字:区域结束_4167>
              <xsl:attribute name="标识符引用_4168">
                <xsl:value-of select= "concat('bk_',@w:id)"/>
              </xsl:attribute>
            </字:区域结束_4167>
           </字:句_419D>
          </xsl:when>
           -->
          <!--字区域开始-->
          <xsl:when test="name(.)='w:bookmarkStart'">
            <xsl:call-template name="bookmarkStart"/>
          </xsl:when>
          <!--字区域结束-->
          <xsl:when test="name(.)='w:bookmarkEnd'">
            <xsl:call-template name="bookmarkEnd"/>
          </xsl:when>
          <!--end-->

          <xsl:when test="name(.)='w:smartTag'"><!--2011/11/29新增智能标签的匹配，这里面也会有句-->
            <xsl:apply-templates select="./w:r" mode="smartTag">
              <xsl:with-param name="pPartFrom" select="$pPartFrom"/>
            </xsl:apply-templates>
          </xsl:when>
          <!--位于region.xsl中-->                                     
          <!--<xsl:when test="name(.)='w:bookmarkStart'">
            <xsl:call-template name="bookmarkStart"/>
          </xsl:when>
          <xsl:when test="name(.)='w:bookmarkEnd'">
            <字:句_419D>
              <xsl:call-template name="bookmarkEnd"/>
            </字:句_419D>       
          </xsl:when>-->
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
          <xsl:when test="name(.)='w:hyperlink'"><!--通过引用方式生成的超链接-->
            <xsl:call-template name="hyperlinkRegion">
              <xsl:with-param name="filename" select="$pPartFrom"/>
            </xsl:call-template>
          </xsl:when>
          <!--yx,add,2010.4.23-->
          <!--以下这一块代码：简单域的转换？？？？？？？？？？？？？？？？？？-->
          <xsl:when test="name(.)='w:fldSimple'">
            <xsl:if test="./w:hyperlink"><!--这种情况我做案例还没有遇到，只遇到由复杂域生成的超链接-->
              <xsl:for-each select="./w:hyperlink">
                <xsl:call-template name="hyperlinkRegion">
                  <xsl:with-param name="filename" select="$pPartFrom"/>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:if>
            <xsl:if test="@w:instr">
              <xsl:variable name="temp" select="normalize-space(@w:instr)"/>
              <xsl:variable name="type">
                <!--yx,change,2010.5.12-->
                <xsl:choose>
                  <xsl:when test="contains($temp,' ')">
                    <xsl:value-of select="substring-before($temp,' ')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$temp"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:choose>
                <xsl:when
                  test="$type='AUTHOR' or $type='DATE' or $type='FILENAME' or $type='REVNUM' or $type='SECTIONPAGES' or $type='TIME' or $type='PAGE' or $type='SECTION' or $type='REF' or $type='XE' or $type='SEQ' or $type='TITLE'or $type='SAVEDATE'or $type='CREATEDATE'or $type='NUMPAGES' or $type='NUMCHARS' or $type='Page' or $type='NUMWORDS'">
                  <xsl:call-template name="SimpleField">
                    <xsl:with-param name="type"
                      select="translate($type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                    <xsl:with-param name="splfldPartFrom" select="$pPartFrom"/>
                  </xsl:call-template>
                </xsl:when>
                <!--<xsl:when test="$type='REVNUM'">
                  <xsl:call-template name="SimpleField">
                    <xsl:with-param name="type" select="'revision'"/>
                    <xsl:with-param name="splfldPartFrom" select="$pPartFrom"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:when test="$type='SECTIONPAGES'">
                  <xsl:call-template name="SimpleField">
                    <xsl:with-param name="type" select="'pageinsection'"/>
                    <xsl:with-param name="splfldPartFrom" select="$pPartFrom"/>
                  </xsl:call-template>
                </xsl:when>-->
              </xsl:choose>
              <xsl:if test="'/w:fldSimple'">
                <字:域结束_41A0/>
              </xsl:if>
            </xsl:if>
            <xsl:if test=".//w:fldChar or .//w:instrText">
              <xsl:if test=".//w:fldChar/@w:fldCharType='begin'"><!--and ./following::*/w:fldChar[1]/preceding::*)--> 
                <!--<xsl:variable name="temp" select="normalize-space(./following::*/w:instrText[1])"/>-->
                <xsl:variable name="temp">
                  <!--这地方有问题，如何把好几个结点下的值都取出来然后做连接?-->
                  <xsl:for-each select="./ancestor::w:p//w:instrText">
                    <xsl:value-of select ="normalize-space(.)"/>
                  </xsl:for-each>
                </xsl:variable>
                <xsl:variable name="type">
                  <xsl:choose>
                    <xsl:when test="contains($temp,' ')">
                      <xsl:value-of select="substring-before($temp,' ')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$temp"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>
                <xsl:choose>
                  <xsl:when
                    test="$type='AUTHOR' or $type='DATE' or $type='FILENAME' or $type='REVNUM' or $type='SECTIONPAGES' or $type='TIME' or $type='PAGE' or $type='SECTION' or $type='REF' or $type='XE' or $type='SEQ' or $type='TITLE'or $type='SAVEDATE'or $type='CREATEDATE'or $type='NUMPAGES' or $type='NUMCHARS' or $type='Page'">
                    <xsl:call-template name="SimpleField">
                      <xsl:with-param name="type"
                        select="translate($type,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                      <xsl:with-param name="splfldPartFrom" select="$pPartFrom"/>
                    </xsl:call-template>
                  </xsl:when>
                  <!--<xsl:when test="$type='REVNUM'">
                    <xsl:call-template name="SimpleField">
                      <xsl:with-param name="type" select="'revision'"/>
                      <xsl:with-param name="splfldPartFrom" select="$pPartFrom"/>
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$type='SECTIONPAGES'">
                    <xsl:call-template name="SimpleField">
                      <xsl:with-param name="type" select="'pageinsection'"/>
                      <xsl:with-param name="splfldPartFrom" select="$pPartFrom"/>
                    </xsl:call-template>
                  </xsl:when>-->
                </xsl:choose>
              </xsl:if>
              <xsl:if test=".//w:fldChar/@w:fldCharType='end'">
                <字:域结束_41A0/>
              </xsl:if>
            </xsl:if>
          </xsl:when>
          <xsl:when test="name(.)='w:sdt'">
            <xsl:call-template name="sdtContentRun">
              <xsl:with-param name="sdtPartFrom" select="$pPartFrom"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name(.)='w:customXml'">
            <xsl:call-template name="CustomXmlRun"/>
          </xsl:when>
          <xsl:when test="name(.)='w:del'">
            <xsl:call-template name="rprchangeBody">
              <xsl:with-param name="lx" select="'delete'"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name(.)='w:ins'">
            <xsl:call-template name="rprchangeBody">
              <xsl:with-param name="lx" select="'insert'"/>
            </xsl:call-template>
          </xsl:when>
			<xsl:when test="name(.)='w:moveFrom'">
				<xsl:call-template name="rprchangeBody">
					<xsl:with-param name="lx" select="'delete'"/>
				</xsl:call-template>
			</xsl:when>
			<xsl:when test="name(.)='w:moveTo'">
				<xsl:call-template name="rprchangeBody">
					<xsl:with-param name="lx" select="'insert'"/>
				</xsl:call-template>
			</xsl:when>
        </xsl:choose>
      </xsl:for-each>
      
      <!--2012-12-20，wudi，OOX到UOF方向的书签实现，start-->
      <!--
      <xsl:if test ="($bookEndNum2 > 0) and (name(following::*[1]) = 'w:bookmarkEnd')">
        <字:句_419D>
          <字:区域结束_4167>
            <xsl:attribute name="标识符引用_4168">
              <xsl:value-of select= "concat('bk_',following::w:bookmarkEnd[1]/@w:id)"/>
            </xsl:attribute>
          </字:区域结束_4167>
        </字:句_419D>
      </xsl:if>
      -->
      <!--end-->
      
    </字:段落_416B>
  </xsl:template>
  
  
  <!--************************************************************************************************************-->
  <!--yx,add,2010.1.27-->
  <xsl:template name="ParagraphStyle">
    <xsl:attribute name="式样引用_419C">
      <xsl:call-template name="IdProducer"><!--产生ID,在styles.xsl中-->
        <xsl:with-param name="ooxId" select="$styleId"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>
  
  <xsl:template match="w:pPr" mode="tbl"> <!--段落中的表（如果存在为相同元素定义的多个处理方法，那么用 mode 可以区分它们。）-->
    <xsl:choose>
      <xsl:when test="w:cnfStyle and not(./w:cnfStyle/@w:val='000000000000')">
        <xsl:for-each select="w:cnfStyle"><!--table style formatting properties-->
          <xsl:if test="substring(@w:val,5,1) = '1'">
            <xsl:attribute name="式样引用_419C">
              <xsl:value-of
                select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'_band1Vert')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="substring(@w:val,6,1)  = '1'">
            <xsl:attribute name="式样引用_419C">
              <xsl:value-of
                select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'_band2Vert')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="substring(@w:val,7,1) = '1'">
            <xsl:attribute name="式样引用_419C">
              <xsl:value-of
                select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'_band1Horz')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="substring(@w:val,8,1) = '1'">
            <xsl:attribute name="式样引用_419C">
              <xsl:value-of
                select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'_band2Horz')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="substring(@w:val,1,1) = '1'">
            <xsl:attribute name="式样引用_419C">
              <xsl:value-of
                select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'_firstRow')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="substring(@w:val,2,1) = '1'">
            <xsl:attribute name="式样引用_419C">
              <xsl:value-of
                select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'_lastRow')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="substring(@w:val,3,1) = '1'">
            <xsl:attribute name="式样引用_419C">
              <xsl:value-of
                select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'_firstCol')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="substring(@w:val,4,1) = '1'">
            <xsl:attribute name="式样引用_419C">
              <xsl:value-of
                select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'_lastCol')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="substring(@w:val,9,1) = '1'">
            <xsl:attribute name="式样引用_419C">
              <xsl:value-of
                select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'_neCell')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="substring(@w:val,10,1) = '1'">
            <xsl:attribute name="式样引用_419C">
              <xsl:value-of
                select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'_nwCell')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="substring(@w:val,11,1) = '1'">
            <xsl:attribute name="式样引用_419C">
              <xsl:value-of
                select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'_seCell')"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="substring(@w:val,12,1) = '1'">
            <xsl:attribute name="式样引用_419C">
              <xsl:value-of
                select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'_swCell')"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:for-each>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="式样引用_419C">
          <xsl:value-of select="concat('id_',ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val,'Paragraph')"/>
        </xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="pPr" match="w:pPr">
    <xsl:param name="framePrLocation"/>
    <xsl:apply-templates
      select="w:pStyle|w:jc|w:textAlignment|w:numPr|w:snapToGrid|w:adjustRightInd|w:autoSpaceDE
                                |w:autoSpaceDN|w:kinsoku|w:wordWrap|w:overflowPunct
                                |w:topLinePunct|w:outlineLvl|w:keepLines|w:keepNext
                                |w:pageBreakBefore|w:suppressAutoHyphens|w:spacing
                                |w:tabs|w:ind|w:pBdr|w:suppressLineNumbers|w:shd"
      mode="pPrChildren"/>
      <xsl:if test="$framePrLocation='2'">
        <xsl:call-template name="framePr"/>
      </xsl:if>
    <xsl:apply-templates select="w:framePr"/><!--2012.4.11直接处理-->

    <!--2014-04-22，wudi，修复表格单元格行距转换BUG，start-->
    <xsl:if test="not(w:spacing)">
      <xsl:if test ="ancestor::w:tbl/w:tblPr/w:tblStyle">
        <xsl:variable name="styId">
          <xsl:value-of select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
        </xsl:variable>
        <xsl:if test="document('word/styles.xml')/w:styles/w:style[@w:styleId=$styId]/w:pPr/w:spacing">
          <xsl:apply-templates select="document('word/styles.xml')/w:styles/w:style[@w:styleId=$styId]/w:pPr/w:spacing" mode="pPrChildren"/>
        </xsl:if>
      </xsl:if>
    </xsl:if>
    <!--end-->
    
    <!--cxl,2012.3.7修改孤行控制转换问题，因为孤行控制会有继承性-->
    <字:孤行控制_418A>
      <xsl:if test="w:widowControl">
        <xsl:variable name="AttrVal">
          <xsl:choose>
            <xsl:when test="w:widowControl/@w:val='0' or w:widowControl/@w:val='false' or w:widowControl/@w:val='off'">0</xsl:when>
            <xsl:otherwise>1</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>     
          <xsl:value-of select="$AttrVal"/>     
      </xsl:if>
      <xsl:if test="not(w:widowControl)">
        <xsl:value-of select="'0'"/>
      </xsl:if>
    </字:孤行控制_418A>

    <!--2014-05-19，wudi，注释掉此部分代码，部分案例不适用如下代码，start-->
    <!--2013-03-26，wudi，修复表格标题位置转换错误的BUG，之前没有考虑w:framePr节点情况，start-->
    <!--
    <xsl:if test ="./w:framePr[@w:xAlign='center']">
      <字:对齐_417D>
        <xsl:attribute name ="水平对齐_421D">
          <xsl:value-of select ="'center'"/>
        </xsl:attribute>
      </字:对齐_417D>
    </xsl:if>
    -->
    <!--end-->
    <!--end-->

    <!--2013-04-10，wudi，修复目录-制表位互操作转换的BUG，start-->

    <!--2014-04-18，wudi，修复制表符转换的BUG，针对w:ptab的处理，start-->
    <xsl:if test ="following-sibling::w:r/w:ptab">
      <字:制表位设置_418F>

        <xsl:for-each select="following-sibling::w:r/w:ptab">
          <字:制表位_4171>
            <xsl:attribute name="位置_4172">
              <xsl:choose>
                <xsl:when test="@w:leader='dot'">
                  <xsl:value-of select="415.34"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="128"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="类型_4173">
              <xsl:value-of select="@w:alignment"/>
            </xsl:attribute>
            <xsl:attribute name="前导符_4174">
              <xsl:choose>
                <xsl:when test="@w:leader='dot'">
                  <xsl:value-of select="'.'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="''"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </字:制表位_4171>
        </xsl:for-each>

      </字:制表位设置_418F>
    </xsl:if>
    <!--end-->

    <!--end-->

    <!--2013-03-12，wudi，删除条件not(following-sibling::w:r)，start-->
    <xsl:if test="./w:rPr">
      <xsl:apply-templates select="w:rPr" mode="pPrChildren"/>
    </xsl:if>
    <!--end-->
    
  </xsl:template>
  
  <xsl:template match="w:rPr" mode="pPrChildren">
    <字:句属性_4158>
      <xsl:apply-templates select="." mode="RunProperties"/>

      <!--2013-03-14，wudi，优化大纲样式，关键词等转换，永中对UOF2.0标准支持得不是很好，start-->
      <xsl:if test ="preceding-sibling::w:pStyle">
        <xsl:variable name ="pStyleId">
          <xsl:value-of select ="preceding-sibling::w:pStyle/@w:val"/>
        </xsl:variable>
        <!--删除条件<xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId = $pStyleId]/w:pPr/w:outlineLvl">-->
        <!--此段代码已是针对永中文字做优化-->
          <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId = $pStyleId]/w:rPr/w:b and not(./w:b)">
            <字:是否粗体_4130>
              <xsl:choose>
                <xsl:when test="not(document('word/styles.xml')/w:styles/w:style[@w:styleId = $pStyleId]/w:rPr/w:b/@w:val) or (document('word/styles.xml')/w:styles/w:style[@w:styleId = $pStyleId]/w:rPr/w:b/@w:val='on') or (document('word/styles.xml')/w:styles/w:style[@w:styleId = $pStyleId]/w:rPr/w:b/@w:val='1') or (document('word/styles.xml')/w:styles/w:style[@w:styleId = $pStyleId]/w:rPr/w:b/@w:val='true')">
                  <xsl:value-of select="'true'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'false'"/>
                </xsl:otherwise>
              </xsl:choose>
            </字:是否粗体_4130>
          </xsl:if>
          <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId = $pStyleId]/w:rPr/w:i and not(./w:i)">
            <字:是否斜体_4131>
              <xsl:choose>
                <xsl:when test="not(document('word/styles.xml')/w:styles/w:style[@w:styleId = $pStyleId]/w:rPr/w:i/@w:val) or (document('word/styles.xml')/w:styles/w:style[@w:styleId = $pStyleId]/w:rPr/w:i/@w:val='on') or (document('word/styles.xml')/w:styles/w:style[@w:styleId = $pStyleId]/w:rPr/w:i/@w:val='1') or (document('word/styles.xml')/w:styles/w:style[@w:styleId = $pStyleId]/w:rPr/w:i/@w:val='true')">
                  <xsl:value-of select="'true'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'false'"/>
                </xsl:otherwise>
              </xsl:choose>
            </字:是否斜体_4131>
          </xsl:if>
        </xsl:if>
      <!--</xsl:if>-->
      <!--end-->

    </字:句属性_4158>
  </xsl:template>
  
  <xsl:template match="w:pStyle" mode="pPrChildren">
    <xsl:attribute name="式样引用_419C">
      <xsl:call-template name="IdProducer">
        <xsl:with-param name="ooxId" select="@w:val"/>
      </xsl:call-template>
    </xsl:attribute>
    
    <!--2012-01-18，wudi，修复OOX到UOF方向BUG：句式样-编号链接式样引用效果丢失，start-->
      <xsl:variable name ="styleId" select ="@w:val"/>
      <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId = $styleId]/w:pPr/w:outlineLvl">
        <字:大纲级别_417C>
        <xsl:value-of select ="document('word/styles.xml')/w:styles/w:style[@w:styleId = $styleId]/w:pPr/w:outlineLvl/@w:val + 1"/>
        </字:大纲级别_417C>
      </xsl:if>
    <!--end-->

    <!--2013-03-15，wudi，修复集成测试OOX到UOF方向-模板5第一段字体加粗效果丢失BUG，start-->
    <xsl:if test ="not(following-sibling::w:rPr)">
      <字:句属性_4158>
        <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:style[@w:styleId = $styleId]/w:rPr" mode="RunProperties"/>
      </字:句属性_4158>
    </xsl:if>
    <!--end-->
    
  </xsl:template>
  
  <!--段落属性下对齐-->
  <xsl:template match="w:jc" mode="pPrChildren">
    <字:对齐_417D><!--uof:attrList="水平对齐 文字对齐"-->
      <xsl:attribute name="水平对齐_421D">
        <xsl:choose>
          <xsl:when test="@w:val='left'">left</xsl:when>
          <xsl:when test="@w:val='right'">right</xsl:when>
          <xsl:when test="@w:val='center'">center</xsl:when>
          <xsl:when test="@w:val='both'">justified</xsl:when>
          <xsl:when test="@w:val='distribute'">distributed</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="justified"/>
            <xsl:message>feedback:lost:Justification_Style_in_Paragraph:justification</xsl:message>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:if test="following-sibling::w:textAlignment">
        <xsl:call-template name="textAlignment">
          <xsl:with-param name="textAlignVal">
            <xsl:value-of select="following-sibling::w:textAlignment/@w:val"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:if>
    </字:对齐_417D>
  </xsl:template>
  <xsl:template match="w:textAlignment" mode="pPrChildren">
    <xsl:if test="not(preceding-sibling::w:jc)">
      <字:对齐_417D><!--uof:attrList="水平对齐 文字对齐"--> 
        <xsl:call-template name="textAlignment">
          <xsl:with-param name="textAlignVal">
            <xsl:value-of select="@w:val"/>
          </xsl:with-param>
        </xsl:call-template>
      </字:对齐_417D>
    </xsl:if>
  </xsl:template>
  <xsl:template name="textAlignment">
    <xsl:param name="textAlignVal"/>
    <xsl:attribute name="文字对齐_421E">
      <xsl:choose>
        <xsl:when test="$textAlignVal='top'">top</xsl:when>
        <xsl:when test="$textAlignVal='center'">center</xsl:when>
        <xsl:when test="$textAlignVal='bottom'">bottom</xsl:when>
        <xsl:when test="$textAlignVal='baseline'">base</xsl:when>
        <xsl:when test="$textAlignVal='auto'">auto</xsl:when>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>
  <!--段落属性下间距-->
  <xsl:template match="w:spacing" mode="pPrChildren">
    <xsl:if test="@w:line">
      <xsl:call-template name="line"/><!--行距-->
    </xsl:if>
    <xsl:call-template name="spacing"/><!--其它段前距之类-->
  </xsl:template>
  <xsl:template name="line">
    <字:行距_417E><!--uof:attrList="类型 值"--> 
      <xsl:if test="@w:lineRule='auto'">
        <xsl:attribute name="类型_417F">multi-lines</xsl:attribute>
        <xsl:attribute name="值_4108">
          <xsl:value-of select="@w:line div 240"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@w:lineRule='atLeast'">
        <xsl:attribute name="类型_417F">at-least</xsl:attribute>
        <xsl:attribute name="值_4108">
          <xsl:value-of select="@w:line div 20"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@w:lineRule='exact'">
        <xsl:attribute name="类型_417F">fixed</xsl:attribute>
        <xsl:attribute name="值_4108">
          <xsl:value-of select="@w:line div 20"/>
        </xsl:attribute>
      </xsl:if>
    </字:行距_417E>
  </xsl:template>
  <xsl:template name="spacing">
    <字:段间距_4180>
      <xsl:choose>
        <xsl:when
          test="@w:beforeAutospacing='true' or @w:beforeAutospacing='1' or @w:beforeAutospacing='on'">
          <字:段前距_4181>
            <字:自动_4182/>
          </字:段前距_4181>
        </xsl:when>
        <xsl:when test="@w:beforeLines">
          <字:段前距_4181>
            <字:相对值_4184><!--无属性uof:attrList="值"--> 
              <xsl:value-of select="@w:beforeLines div 100"/>
            </字:相对值_4184>
          </字:段前距_4181>
        </xsl:when>
        <xsl:when test="@w:before">
          <字:段前距_4181>
            <字:绝对值_4183>
              <xsl:value-of select="@w:before div 20"/>
            </字:绝对值_4183>
          </字:段前距_4181>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when
          test="@w:afterAutospacing='true' or @w:afterAutospacing='1' or @w:afterAutospacing='on'">
          <字:段后距_4185>
            <字:自动_4182/>
          </字:段后距_4185>
        </xsl:when>
        <xsl:when test="@w:afterLines">
          <字:段后距_4185>
            <字:相对值_4184>
              <xsl:value-of select="@w:afterLines div 100"/>
            </字:相对值_4184>
          </字:段后距_4185>
        </xsl:when>
        <xsl:when test="@w:after">
          <字:段后距_4185>
            <字:绝对值_4183>
              <xsl:value-of select="@w:after div 20"/>             
            </字:绝对值_4183>
          </字:段后距_4185>
        </xsl:when>
      </xsl:choose>
    </字:段间距_4180>
  </xsl:template>
  <xsl:template match="w:numPr" mode="pPrChildren">
    <字:自动编号信息_4186><!--uof:attrList="编号引用 编号级别 重新编号 起始编号"-->
		<!--<xsl:variable name="defaultEffect">
			<xsl:value-of select="../../@w:rsidP"/>
		</xsl:variable>-->
		<!--<xsl:if test="document('word/stylesWithEffects.xml')//w:style[w:rsid/@w:val=$defaultEffect]/w:pPr/w:numPr/w:numId">
			<xsl:attribute name="编号引用_4187">
				<xsl:value-of select="concat('bn_',document('word/stylesWithEffects.xml')//w:style[w:rsid/@w:val=$defaultEffect]/w:pPr/w:numPr/w:numId/@w:val)"/>
			</xsl:attribute>
		</xsl:if>
		<xsl:if test="not(document('word/stylesWithEffects.xml')//w:style[w:rsid/@w:val=$defaultEffect]/w:pPr/w:numPr/w:numId)">
			<xsl:if test="./w:numId">
				<xsl:attribute name="编号引用_4187">
					<xsl:value-of select="concat('bn_',./w:numId/@w:val)"/>
				</xsl:attribute>
			</xsl:if>
		</xsl:if>-->

		<xsl:if test="./w:numId">
			<xsl:variable name="numId" select="./w:numId/@w:val"/>
			<xsl:variable name="abstractNumId">
				<xsl:value-of select="document('word/numbering.xml')/w:numbering/w:num[@w:numId=$numId]/w:abstractNumId/@w:val"/>
			</xsl:variable>
      
      <!--2012-01-18，wudi，修复OOX到UOF方向BUG：句式样-编号链接式样引用效果丢失，start-->
      <xsl:if test ="./w:numId/@w:val = '0'">
        <xsl:attribute name="编号引用_4187">
          <xsl:value-of select="'bn_1'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test ="./w:numId/@w:val != '0'">
        <xsl:if test="document('word/numbering.xml')/w:numbering/w:abstractNum[@w:abstractNumId=$abstractNumId]/w:lvl">
          <xsl:attribute name="编号引用_4187">
            <xsl:value-of select="concat('bn_',$numId)"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="not(document('word/numbering.xml')/w:numbering/w:abstractNum[@w:abstractNumId=$abstractNumId]/w:lvl)">
          <xsl:variable name="numStyleLink" select="document('word/numbering.xml')/w:numbering/w:abstractNum[@w:abstractNumId=$abstractNumId]/w:numStyleLink/@w:val"/>
          <xsl:variable name="realAbstractNumId" select="document('word/numbering.xml')/w:numbering/w:abstractNum[w:styleLink/@w:val=$numStyleLink and w:lvl]/@w:abstractNumId"/>
          <xsl:variable name="realNumId" select="document('word/numbering.xml')/w:numbering/w:num[w:abstractNumId/@w:val=$realAbstractNumId]/@w:numId"/>
          <xsl:attribute name="编号引用_4187">
            <xsl:value-of select="concat('bn_',$realNumId)"/>
          </xsl:attribute>
        </xsl:if>
      </xsl:if>
      <!--end-->
      
		</xsl:if>
      
    <!--2012-01-18，wudi，修复OOX到UOF方向BUG：句式样-编号链接式样引用效果丢失，start-->
    <xsl:if test ="not(./w:numId)">
      <xsl:attribute name="编号引用_4187">
        <xsl:value-of select="'bn_1'"/>
      </xsl:attribute>
    </xsl:if>
    <!--end-->
      
    <xsl:if test="./w:ilvl">

      <!--2012-01-18，wudi，修复OOX到UOF方向BUG：句式样-编号链接式样引用效果丢失，start-->
      <xsl:if test ="./w:numId/@w:val = '0'">
        <xsl:attribute name="编号级别_4188">
          <xsl:value-of select="./w:ilvl/@w:val"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test ="not(./w:numId/@w:val = '0')">
        <xsl:attribute name="编号级别_4188">
          <xsl:value-of select="./w:ilvl/@w:val+1"/>
        </xsl:attribute>
      </xsl:if>
      <!--end-->
      
    </xsl:if>
    <xsl:if test="not(./w:ilvl)">
      <xsl:attribute name="编号级别_4188">
        <xsl:value-of select="'1'"/>
      </xsl:attribute>
    </xsl:if>
      
    </字:自动编号信息_4186>
  </xsl:template>
  <xsl:template match="w:snapToGrid" mode="pPrChildren">
    <字:是否对齐网格_4190>
        <xsl:if test="not(@w:val)">
          <xsl:value-of select="'true'"/>
        </xsl:if>
        <xsl:if test="(@w:val='0')or(@w:val='false')or(@w:val='off')">
          <xsl:value-of select="'false'"/>
        </xsl:if>
        <xsl:if test="(@w:val='1') or (@w:val='true') or (@w:val='on')">
          <xsl:value-of select="'true'"/>
        </xsl:if>
    </字:是否对齐网格_4190>
  </xsl:template>
  <xsl:template match="w:adjustRightInd" mode="pPrChildren">
    <字:是否有网格自动调整右缩进_419A>
        <xsl:choose>
          <xsl:when test="(@w:val='0') or (@w:val='false') or (@w:val='off')">false</xsl:when>
          <xsl:otherwise>true</xsl:otherwise>
        </xsl:choose>
    </字:是否有网格自动调整右缩进_419A>
  </xsl:template>
  <xsl:template match="w:autoSpaceDE" mode="pPrChildren">
    <字:是否自动调整中英文字符间距_4198>
        <xsl:choose>
          <xsl:when test="(@w:val='0') or (@w:val='false') or (@w:val='off')">false</xsl:when>
          <xsl:otherwise>true</xsl:otherwise>
        </xsl:choose>
    </字:是否自动调整中英文字符间距_4198>
  </xsl:template>
  <xsl:template match="w:autoSpaceDN" mode="pPrChildren">
    <字:是否自动调整中文与数字间距_4199>
        <xsl:choose>
          <xsl:when test="(@w:val='0') or (@w:val='false') or (@w:val='off')">false</xsl:when>
          <xsl:otherwise>true</xsl:otherwise>
        </xsl:choose>
    </字:是否自动调整中文与数字间距_4199>
  </xsl:template>
  <xsl:template match="w:kinsoku" mode="pPrChildren">
    <字:是否采用中文习惯首尾字符_4197>
        <xsl:choose>
          <xsl:when test="(@w:val='0') or (@w:val='false') or (@w:val='off')">false</xsl:when>
          <xsl:otherwise>true</xsl:otherwise>
        </xsl:choose>
    </字:是否采用中文习惯首尾字符_4197>
  </xsl:template>
  <xsl:template match="w:wordWrap" mode="pPrChildren">
    <字:是否允许单词断字_4194>
        <xsl:choose>
          <xsl:when test="(@w:val='0') or (@w:val='false') or (@w:val='off')">true</xsl:when>
          <xsl:otherwise>false</xsl:otherwise>
        </xsl:choose>
    </字:是否允许单词断字_4194>
  </xsl:template>
  <xsl:template match="w:overflowPunct" mode="pPrChildren">
    <字:是否行首尾标点控制_4195>
        <xsl:choose>
          <xsl:when test="(@w:val='0') or (@w:val='false') or (@w:val='off')">false</xsl:when>
          <xsl:otherwise>true</xsl:otherwise>
        </xsl:choose>
    </字:是否行首尾标点控制_4195>
  </xsl:template>
  <xsl:template match="w:topLinePunct" mode="pPrChildren">
    <字:是否行首标点压缩_4196>
      <xsl:choose>
        <xsl:when test="(@w:val='0') or (@w:val='false') or (@w:val='off')">false</xsl:when>
        <xsl:otherwise>true</xsl:otherwise>
      </xsl:choose>
    </字:是否行首标点压缩_4196>
  </xsl:template>
  <xsl:template match="w:outlineLvl" mode="pPrChildren">
    <xsl:if test="@w:val">
      <字:大纲级别_417C>
        <xsl:value-of select="@w:val+1"/>
      </字:大纲级别_417C>
    </xsl:if>
  </xsl:template>
  <xsl:template match="w:keepLines" mode="pPrChildren">
    <xsl:choose>
      <xsl:when test="@w:val='false' or @w:val='0' or @w:val='off'">
        <字:是否段中不分页_418C>false</字:是否段中不分页_418C>
      </xsl:when>
      <xsl:otherwise>
        <字:是否段中不分页_418C>true</字:是否段中不分页_418C>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="w:keepNext" mode="pPrChildren">
    <xsl:choose>
      <xsl:when test="@w:val='false' or @w:val='0' or @w:val='off'">
        <字:是否与下段同页_418D>false</字:是否与下段同页_418D>
      </xsl:when>
      <xsl:otherwise>
        <字:是否与下段同页_418D>true</字:是否与下段同页_418D>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="w:pageBreakBefore" mode="pPrChildren">
    <xsl:choose>
      <xsl:when test="@w:val='false' or @w:val='0' or @w:val='off'">
        <字:是否段前分页_418E>false</字:是否段前分页_418E>
      </xsl:when>
      <xsl:otherwise>
        <字:是否段前分页_418E>true</字:是否段前分页_418E>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template match="w:suppressAutoHyphens" mode="pPrChildren">
    <xsl:choose>
      <xsl:when test="@w:val='false' or @w:val='0' or @w:val='off'">
        <字:是否取消断字_4192>false</字:是否取消断字_4192>
      </xsl:when>
      <xsl:otherwise>
        <字:是否取消断字_4192>true</字:是否取消断字_4192>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="w:tabs" mode="pPrChildren">
    <字:制表位设置_418F>
      <xsl:for-each select="w:tab">
        <xsl:if
          test="name(.)='w:tab' and not(@w:val='bar') and not(@w:val='clear') and not(@w:val='num')">
          <字:制表位_4171>  <!--新增属性"取消制表位_4245"-->
            <xsl:attribute name="位置_4172">
              <xsl:value-of select="number(@w:pos) div 20"/>
            </xsl:attribute>
            <xsl:attribute name="类型_4173">
              <xsl:variable name="tabStopVal" select="@w:val"/>
              <xsl:if test="$tabStopVal='left'">left</xsl:if>
              <xsl:if test="$tabStopVal='right'">right</xsl:if>
              <xsl:if test="$tabStopVal='center'">center</xsl:if>
              <xsl:if test="$tabStopVal='decimal'">decimal</xsl:if>
            </xsl:attribute>
            <xsl:if test="@w:leader and not(@w:leader='none')">
              <xsl:attribute name="前导符_4174">
                <xsl:choose>
                  <xsl:when test="./@w:leader='dot'">&#46;</xsl:when>
                  <xsl:when test="./@w:leader='hyphen'">&#45;</xsl:when>
                  <xsl:when test="./@w:leader='underscore'">&#95;</xsl:when>
                  <xsl:when test="./@w:leader='middleDot'">&#183;</xsl:when>
                  <xsl:when test="./@w:leader='heavy'">&#46;</xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
          </字:制表位_4171>
        </xsl:if>
      </xsl:for-each>
    </字:制表位设置_418F>
  </xsl:template>
  <!--缩进-->
  <xsl:template match="w:ind" mode="pPrChildren">
    <字:缩进_411D>
      <xsl:call-template name="computInd"/>
    </字:缩进_411D>
  </xsl:template>
  <xsl:template name="computInd">
    <xsl:variable name="hangingVal">
      <xsl:choose>
        <xsl:when test="@w:hanging">
          <xsl:value-of select="@w:hanging"/>
        </xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="hasLeftInd">
      <xsl:choose>
        <xsl:when test="@w:left and not(@w:leftChars)">true</xsl:when>
        <xsl:otherwise>false</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="hasLeftCharInd">
      <xsl:choose>
        <xsl:when test="@w:leftChars">true</xsl:when>
        <xsl:otherwise>false</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="hasRightInd">
      <xsl:choose>
        <xsl:when test="@w:right and not(@w:rightChars)">true</xsl:when>
        <xsl:otherwise>false</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="hasRightCharInd">
      <xsl:choose>
        <xsl:when test="@w:rightChars">true</xsl:when>
        <xsl:otherwise>false</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="AbsFirstLineInd">
      <xsl:choose>
        <xsl:when test="@w:firstLine and not(@w:hanging) and (not(@w:firstLineChars) or @w:firstLineChars='0')">true</xsl:when>
        <xsl:otherwise>false</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="firstlineCharInd">
      <xsl:choose>
        <xsl:when test="@w:firstLineChars &gt;0 and not(@w:hangingChars)">true</xsl:when>
        <xsl:otherwise>false</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="AbsHangingInd">
      <xsl:choose>
        <xsl:when test="@w:hanging and not(@w:hangingChars)">true</xsl:when>
        <xsl:otherwise>false</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="hangingCharInd">
      <xsl:choose>
        <xsl:when test="@w:hangingChars">true</xsl:when>
        <xsl:otherwise>false</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when
        test="not(following-sibling::w:mirrorIndents) or (following-sibling::w:mirrorIndents/@w:val='false') 
                      or (following-sibling::w:mirrorIndents/@w:val='off') or (following-sibling::w:mirrorIndents/@w:val='0')">
        <xsl:if test="@w:left and not(@w:leftChars)">
          <字:左_410E>
            <字:绝对_4107>
              <xsl:attribute name="值_410F">
                <xsl:value-of select="(@w:left - $hangingVal) div 20"/>
              </xsl:attribute>
            </字:绝对_4107>
          </字:左_410E>
        </xsl:if>
        <xsl:if test="@w:leftChars">
          <字:左_410E>
            <字:相对_4109>
              <xsl:attribute name="值_4108">
                <xsl:value-of select="@w:leftChars div 100"/>
              </xsl:attribute>
            </字:相对_4109>
          </字:左_410E>
        </xsl:if>
        <xsl:if test="@w:right and not(@w:rightChars)">
          <字:右_4110>
            <字:绝对_4107>
              <xsl:attribute name="值_410F">
                <xsl:value-of select="@w:right div 20"/>
              </xsl:attribute>
            </字:绝对_4107>
          </字:右_4110>
        </xsl:if>
        <xsl:if test="@w:rightChars">
          <字:右_4110>
            <字:相对_4109>
              <xsl:attribute name="值_4108">
                <xsl:value-of select="@w:rightChars div 100"/>
              </xsl:attribute>
            </字:相对_4109>
          </字:右_4110>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="OrgInsdVal">
          <xsl:choose>
            <xsl:when test="@w:left">
              <xsl:value-of select="@w:left"/>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="insideVal">
          <xsl:value-of select="$OrgInsdVal - $hangingVal"/>
        </xsl:variable>
        <xsl:variable name="outsideVal">
          <xsl:choose>
            <xsl:when test="@w:right">
              <xsl:value-of select="@w:right"/>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="averageInd">
          <xsl:value-of select="($insideVal + $outsideVal) div 2"/>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="@w:leftChars and @w:rightChars">
            <字:左_410E>
              <字:相对_4109>
                <xsl:attribute name="值_4108">
                  <xsl:value-of select="(@w:leftChars + @w:rightChars) div 200"/>
                </xsl:attribute>
              </字:相对_4109>
            </字:左_410E>
            <字:右_4110>
              <字:相对_4109>
                <xsl:attribute name="值_4108">
                  <xsl:value-of select="(@w:leftChars + @w:rightChars) div 200"/>
                </xsl:attribute>
              </字:相对_4109>
            </字:右_4110>
          </xsl:when>
          <xsl:otherwise>
            <字:左_410E>
              <字:绝对_4107>
                <xsl:attribute name="值_410F">
                  <xsl:value-of select="$averageInd div 20"/>
                </xsl:attribute>
              </字:绝对_4107>
            </字:左_410E>
            <字:右_4110>
              <字:绝对_4107>
                <xsl:attribute name="值_410F">
                  <xsl:value-of select="$averageInd div 20"/>
                </xsl:attribute>
              </字:绝对_4107>
            </字:右_4110>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:if test="@w:firstLine and not(@w:hanging) and (not(@w:firstLineChars) or @w:firstLineChars='0')">
      <字:首行_4111>
        <字:绝对_4107>
          <xsl:attribute name="值_410F">
            <xsl:value-of select="@w:firstLine div 20"/>
          </xsl:attribute>
        </字:绝对_4107>
      </字:首行_4111>
    </xsl:if>
    <xsl:if test="@w:hanging and not(@w:hangingChars)">
      <字:首行_4111>
        <字:绝对_4107>
          <xsl:attribute name="值_410F">
            <xsl:value-of select="-(@w:hanging div 20)"/>
          </xsl:attribute>
        </字:绝对_4107>
      </字:首行_4111>
    </xsl:if>
    <xsl:if test="@w:firstLineChars and not(@w:firstLineChars='0') and not(@w:hangingChars)">
      <字:首行_4111>
        <字:相对_4109>
          <xsl:attribute name="值_4108">
            <xsl:value-of select="@w:firstLineChars div 100"/>
          </xsl:attribute>
        </字:相对_4109>
      </字:首行_4111>
    </xsl:if>
    <xsl:if test="@w:hangingChars">
      <字:首行_4111>
        <字:相对_4109>
          <xsl:attribute name="值_4108">
            <xsl:value-of select="-(@w:hangingChars div 100)"/>
          </xsl:attribute>
        </字:相对_4109>
      </字:首行_4111>
    </xsl:if>
  </xsl:template>
  <!--段落边框 2011-02-24差异文档中有@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
  <xsl:template match="w:pBdr" mode="pPrChildren">
    <字:边框_4133><!--段落边框-->
      <xsl:if test="./w:top/@w:shadow='1'">
        <xsl:attribute name="阴影类型_C645">
          <xsl:value-of select="'right-bottom'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./w:left">
        <!--类型转换为线型和虚实 2011-5-10 12:50:40-->
        <!--关于阴影标准中有变化，1.1是布尔型，现在是枚举类型，并且直接作为边框的属性@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
        <uof:左_C613><!--attrList="线型 虚实 宽度 边距 颜色"--> 
          <xsl:apply-templates select="w:left" mode="pBdrChildren"/>
        </uof:左_C613>
      </xsl:if>
      <xsl:if test="./w:top">
        <uof:上_C614>
          <xsl:apply-templates select="w:top" mode="pBdrChildren"/>
        </uof:上_C614>
      </xsl:if>
      <xsl:if test="./w:right">
        <uof:右_C615>
          <xsl:apply-templates select="w:right" mode="pBdrChildren"/>
        </uof:右_C615>
      </xsl:if>
      <xsl:if test="./w:bottom">
        <uof:下_C616>
          <xsl:apply-templates select="w:bottom" mode="pBdrChildren"/>
        </uof:下_C616>
      </xsl:if>
    </字:边框_4133>
  </xsl:template>
  <xsl:template match="w:left" mode="pBdrChildren">
    <xsl:call-template name="singleBorder"/>
  </xsl:template>
  <xsl:template match="w:top" mode="pBdrChildren">
    <xsl:call-template name="singleBorder"/>
  </xsl:template>
  <xsl:template match="w:right" mode="pBdrChildren">
    <xsl:call-template name="singleBorder"/>
  </xsl:template>
  <xsl:template match="w:bottom" mode="pBdrChildren">
    <xsl:call-template name="singleBorder"/>
  </xsl:template>
  <xsl:template name="singleBorder">
  <!--Lee 线型 虚实 2011年5月10日-->

    <xsl:call-template name="borderStyle"><!--段落边框线型及虚实模板-->
      <xsl:with-param name="styleVal" select="@w:val"/>
    </xsl:call-template>

    <xsl:if test="@w:sz">
      <xsl:attribute name="宽度_C60F">
        <xsl:value-of select="number(@w:sz) div 8"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="@w:space">
      <xsl:attribute name="边距_C610">
        <xsl:value-of select="@w:space"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="@w:color">
      <xsl:attribute name="颜色_C611">
        <xsl:variable name="color" select="@w:color"/>
        <xsl:if test="not($color='auto')">
          <xsl:text>#</xsl:text>
        </xsl:if>
        <xsl:value-of select="$color"/>
      </xsl:attribute>
    </xsl:if>
    <!--边框阴影这里有待修改@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
    <!--<xsl:if test="@w:shadow">
      <xsl:attribute name="阴影">
        <xsl:value-of select="@w:shadow"/>
      </xsl:attribute>
    </xsl:if>-->
  </xsl:template>


  <!--段落边框的转换，对应关系参照差异文档中边框 2011-02-24-->
  <xsl:template name="borderStyle">
    <xsl:param name="styleVal"/>
     
    <xsl:choose>
      <xsl:when test="$styleVal = 'nil' or $styleVal='none'">
        <xsl:attribute name="线型_C60D">none</xsl:attribute>
      </xsl:when>

      <xsl:when test="$styleVal='thinThickThinLargeGap'">
        <xsl:attribute name="线型_C60D">thick-between-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='thinThickSmallGap'">
        <xsl:attribute name ="线型_C60D">thick-thin</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='thickThinSmallGap'">
        <xsl:attribute name ="线型_C60D">thin-thick</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">solid</xsl:attribute>
      </xsl:when> 
      <xsl:when test="$styleVal='thick'">
        <xsl:attribute name ="线型_C60D">single</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">round-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='thinThickThinSmallGap'">
        <xsl:attribute name ="线型_C60D">thick-between-thin</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <!--uof没实现转换为相近的，详见差异文档-->
      <xsl:when test ="$styleVal='thinThickMediumGap'">
        <xsl:attribute name ="线型_C60D">thick-thin</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='thickThinMediumGap'">
        <xsl:attribute name ="线型_C60D">thin-thick</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='thinThickLargeGap'">
        <xsl:attribute name ="线型_C60D">thick-thin</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='thickThinLargeGap'">
        <xsl:attribute name ="线型_C60D">thin-thick</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='thinThickThinMediumGap'">
        <xsl:attribute name ="线型_C60D">thick-between-thin</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$styleVal='dashLongHeavy'">
        <xsl:attribute name="线型_C60D">thick-between-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$styleVal='dashLong'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$styleVal='dashDotHeavy'">
        <xsl:attribute name="线型_C60D">thick-between-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test="$styleVal='wavyHeavy'">
        <xsl:attribute name="线型_C60D">thick-between-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash-dot-dot</xsl:attribute>
      </xsl:when>

      <!--2013-03-26，wudi，修复字边框转换错误的BUG，start-->
      <xsl:when test="$styleVal='doubleWave'">
        <xsl:attribute name="线型_C60D">double</xsl:attribute>
        <xsl:attribute name="虚实_C60E">wave</xsl:attribute>
      </xsl:when>
      
      <xsl:when test="$styleVal='wave'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">wave</xsl:attribute>
      </xsl:when>
      <!--end-->

      <xsl:when test="$styleVal='dash'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='dashed'">
        <xsl:attribute name ="线型_C60D">single</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$styleVal='single'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="$styleVal='dotted'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">square-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='dashSmallGap'">
        <xsl:attribute name ="线型_C60D">single</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="$styleVal='dotDash'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test="$styleVal='dotDotDash'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash-dot-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test="$styleVal='double'">
        <xsl:attribute name="线型_C60D">double</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='triple'">
        <xsl:attribute name ="线型_C60D">double</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">dash-dot-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='dashDotStroked'">
        <xsl:attribute name ="线型_C60D">single</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">long-dash-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='threeDEmboss'">
        <xsl:attribute name ="线型_C60D">single</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='threeDEngrave'">
        <xsl:attribute name ="线型_C60D">single</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='outset'">
        <xsl:attribute name ="线型_C60D">single</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test ="$styleVal='inset'">
        <xsl:attribute name ="线型_C60D">single</xsl:attribute>
        <xsl:attribute name ="虚实_C60E">solid</xsl:attribute>
      </xsl:when>       
      <xsl:otherwise>
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:otherwise>
    </xsl:choose> 
  </xsl:template>
  <xsl:template match="w:suppressLineNumbers" mode="pPrChildren">
    <字:是否取消行号_4193>
        <xsl:choose>
          <xsl:when test="(@w:val='0') or (@w:val='false') or (@w:val='off')">false</xsl:when>
          <xsl:otherwise>true</xsl:otherwise>
        </xsl:choose>
    </字:是否取消行号_4193>
  </xsl:template>
  <xsl:template match="w:shd" mode="pPrChildren">
    <xsl:if test="./@w:val!='nil'">
      <字:填充_4134>
        <xsl:call-template name="shd"/>
      </字:填充_4134>
    </xsl:if>
  </xsl:template>
  <!--2012.4.11首字下沉直接处理-->
  <xsl:template match="w:framePr">
    <xsl:if test="./@w:dropCap = 'drop' or ./@w:dropCap = 'margin'">
      <字:首字下沉_4191>
        <xsl:attribute name="类型_413B">
          <xsl:if test="./@w:dropCap='drop'">
            <xsl:value-of select="'dropped'"/>
          </xsl:if>
          <xsl:if test="./@w:dropCap='margin'">
            <xsl:value-of select="'margin'"/>
          </xsl:if>
        </xsl:attribute>
        <xsl:if test="starts-with(..//w:rFonts/@w:asciiTheme,'minor')">
          <xsl:attribute name="字体引用_4176">
            <xsl:variable name="themefont"
              select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
            <xsl:value-of select="translate($themefont,' ','')"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:attribute name="字符数_4177">
          <xsl:value-of select="'1'"/>
        </xsl:attribute>
        <xsl:attribute name="行数_4178">
          <xsl:choose>
            <xsl:when test="./@w:lines and ./@w:dropCap">
              <xsl:value-of select="./@w:lines"/>
            </xsl:when>
            <xsl:otherwise>1</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="间距_4179">
          <xsl:choose>
            <xsl:when test="./@w:wrap='around' and ./@w:hSpace">
              <xsl:value-of select="number(./@w:hSpace) div 20"/>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </字:首字下沉_4191>
    </xsl:if>
  </xsl:template>
  <!--cxl,2012.3.8首字下沉转换-->
  <xsl:template name="framePr">
    <xsl:if test="preceding::w:pPr/w:framePr/@w:dropCap = 'drop' or preceding::w:pPr/w:framePr/@w:dropCap = 'margin'">
      <字:首字下沉_4191>
        <xsl:attribute name="类型_413B">
          <xsl:if test="preceding::w:pPr/w:framePr/@w:dropCap='drop'">
            <xsl:value-of select="'dropped'"/>
          </xsl:if>
          <xsl:if test="preceding::w:pPr/w:framePr/@w:dropCap='margin'">
            <xsl:value-of select="'margin'"/>
          </xsl:if>
        </xsl:attribute>
        <xsl:if test="starts-with(..//w:rFonts/@w:asciiTheme,'minor')">
          <xsl:attribute name="字体引用_4176">
            <xsl:variable name="themefont"
              select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
            <xsl:value-of select="translate($themefont,' ','')"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:attribute name="字符数_4177">
          <xsl:value-of select="'1'"/>
        </xsl:attribute>
        <xsl:attribute name="行数_4178">
          <xsl:choose>
            <xsl:when test="preceding::w:pPr/w:framePr/@w:lines and preceding::w:pPr/w:framePr/@w:dropCap">
              <xsl:value-of select="preceding::w:pPr/w:framePr/@w:lines"/>
            </xsl:when>
            <xsl:otherwise>1</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="间距_4179">
          <xsl:choose>
            <xsl:when test="(preceding::w:pPr/w:framePr/@w:wrap='around')and preceding::w:pPr/w:framePr/@w:hSpace">
              <xsl:value-of select="number(preceding::w:pPr/w:framePr/@w:hSpace) div 20"/>
            </xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </字:首字下沉_4191>
    </xsl:if>
  </xsl:template>
  
  <!--run模板11.4日cxl修改*******************************************************************************-->
  <xsl:template name="run">
    <xsl:param name="rPartFrom"/>
    
    <!--2014-05-05，wudi，增加postr参数，记录单元格所属行数，修复单元格字体加粗效果转换BUG，start-->
    <xsl:param name="postr"/>
    <xsl:param name="postc"/>
    <!--end-->
    
    <xsl:for-each select="node()">
      <xsl:choose>
        <xsl:when test="name(.)='w:rPr'">
          <字:句属性_4158>
            <xsl:if test="w:rPrChange"><!--修订信息-->
              <xsl:apply-templates select="w:rPrChange/w:rPr" mode="RunProperties"/>
            </xsl:if>
            <xsl:if test="../../w:pPr/w:pPrChange"><!--这是什么修订？-->
              <xsl:apply-templates select="../../w:pPr" mode="RunProperties"/>
            </xsl:if>
            <xsl:if test="not(w:rPrChange)"><!--对所有不是修订信息的句（rPr）中内容进行转换-->
              <!--<xsl:if test="not(preceding::w:p[1]//w:framePr)">--><!--cxl,2012.3.8首字下沉--><!--
                <xsl:apply-templates select="." mode="RunProperties"/>
              </xsl:if>
              <xsl:if test="preceding::w:p[1]//w:framePr">
                <xsl:apply-templates select="self::w:rPr[position()!=1]" mode="RunProperties"/>
              </xsl:if>-->
              <xsl:apply-templates select="." mode="RunProperties">

                <!--2014-05-05，wudi，增加postr参数，记录单元格所属行数，修复单元格字体加粗效果转换BUG，start-->
                <xsl:with-param name="postr" select="$postr"/>
                <xsl:with-param name="postc" select="$postc"/>
                <!--end-->
                
              </xsl:apply-templates>
            </xsl:if>
            
            <!--2013-03-12，wudi，修复脚注-尾注转换编号上标样式丢失BUG，start-->
            <xsl:if test ="name(following-sibling::*[1])='w:footnoteReference' or name(following-sibling::*[1])='w:endnoteReference'">
              <字:上下标类型_4143>sup</字:上下标类型_4143>
            </xsl:if>
            <xsl:if test ="name(following-sibling::*[1])='w:footnoteRef' or name(following-sibling::*[1])='w:endnoteRef'">
              <字:上下标类型_4143>sup</字:上下标类型_4143>
            </xsl:if>
            <!--2013-04-18，wudi，修复特殊字符转换BUG，增加限制条件and (ancestor::w:footnote or ancestor::w:endnote)-->
            <xsl:if test ="name(following-sibling::*[1])='w:sym' and (ancestor::w:footnote or ancestor::w:endnote)">
              <字:上下标类型_4143>sup</字:上下标类型_4143>
            </xsl:if>
            <!--end-->

            <!--2013-03-27，wudi，修复给标题设置超链接，转换后标题字体样式效果丢失的BUG，start-->
            <xsl:if test ="contains(ancestor::w:p//w:instrText,'HYPERLINK')">
              <xsl:variable name ="styId">
                <xsl:value-of select ="ancestor::w:p/w:pPr/w:pStyle/@w:val"/>
              </xsl:variable>
              <!--<styId>
                <xsl:value-of select ="$styId"/>
              </styId>-->
              <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:b and not(w:b)">
                <字:是否粗体_4130>
                  <xsl:choose>
                    <xsl:when test="not(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:b/@w:val) or (document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:b/@w:val='on')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:b/@w:val='1')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:b/@w:val='true')">
                      <xsl:value-of select="'true'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'false'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </字:是否粗体_4130>
              </xsl:if>
              <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:sz and not(w:sz)">
                <xsl:variable name ="size">
                  <xsl:value-of select ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:sz/@w:val"/>
                </xsl:variable>
                <字:字体_4128>
                  <xsl:attribute name ="字号_412D">
                    <xsl:value-of select ="format-number($size div 2,'0.0')"/>
                  </xsl:attribute>
                </字:字体_4128>
              </xsl:if>
            </xsl:if>
            <!--end-->

          </字:句属性_4158>
          <xsl:for-each select="w:rPrChange">
            <xsl:call-template name="rprchangeBody"><!--修订信息模板在revisions.xsl中-->
              <xsl:with-param name="lx" select="'format'"/>
              <xsl:with-param name="revPartFrom" select="$rPartFrom"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="name(.)='w:t'"><!--文本串-->
          <xsl:if
            test="not(preceding-sibling::w:footnoteReference) and not(preceding-sibling::w:endnoteReference)">

            <!--2013-03-21，wudi，修复空格符转换互操作BUG，start-->
            <xsl:variable name ="wtxt">
              <xsl:value-of select ="."/>
            </xsl:variable>
            <xsl:variable name ="pos">
              <xsl:value-of select ="number(string-length(.))"/>
            </xsl:variable>
            <xsl:call-template name="text">
              <xsl:with-param name ="pos1" select ="$pos"/>
              <xsl:with-param name ="cnt1" select ="1"/>
              <xsl:with-param name ="wtxt1" select ="$wtxt"/>
            </xsl:call-template>
            <!--end-->
            
          </xsl:if>
        </xsl:when>
        <xsl:when test="name(.)='w:delText'">

          <!--2013-03-21，wudi，修复空格符转换互操作BUG，start-->
          <xsl:variable name ="wtxt">
            <xsl:value-of select ="."/>
          </xsl:variable>
          <xsl:variable name ="pos">
            <xsl:value-of select ="number(string-length(.))"/>
          </xsl:variable>
          <xsl:call-template name="text">
            <xsl:with-param name ="pos1" select ="$pos"/>
            <xsl:with-param name ="cnt1" select ="1"/>
            <xsl:with-param name ="wtxt1" select ="$wtxt"/>
          </xsl:call-template>
          <!--end-->
          
        </xsl:when>
        <xsl:when test="name(.)='delInsertText'">
          <xsl:if test="(.!='eq \o\ac(')and(.!=',')and(.!=')')and(.!=' ')">
            
            <!--2013-03-21，wudi，修复空格符转换互操作BUG，start-->
            <xsl:variable name ="wtxt">
              <xsl:value-of select ="."/>
            </xsl:variable>
            <xsl:variable name ="pos">
              <xsl:value-of select ="number(string-length(.))"/>
            </xsl:variable>
            <xsl:call-template name="text">
              <xsl:with-param name ="pos1" select ="$pos"/>
              <xsl:with-param name ="cnt1" select ="1"/>
              <xsl:with-param name ="wtxt1" select ="$wtxt"/>
            </xsl:call-template>
            <!--end-->
            
          </xsl:if>
          <xsl:message terminate="no">feedback:lost:Enclose_Characters_in_Paragraph:enclose characters</xsl:message>
        </xsl:when>
        <xsl:when test="name(.)='w:ruby'"><!--拼音指南Phonetic Guide-->
          <xsl:for-each select="./w:rubyBase/w:r/w:t">

            <!--2013-03-21，wudi，修复空格符转换互操作BUG，start-->
            <xsl:variable name ="wtxt">
              <xsl:value-of select ="."/>
            </xsl:variable>
            <xsl:variable name ="pos">
              <xsl:value-of select ="number(string-length(.))"/>
            </xsl:variable>
            <xsl:call-template name="text">
              <xsl:with-param name ="pos1" select ="$pos"/>
              <xsl:with-param name ="cnt1" select ="1"/>
              <xsl:with-param name ="wtxt1" select ="$wtxt"/>
            </xsl:call-template>
            <!--end-->

          </xsl:for-each>
          <xsl:message terminate="no">feedback:lost:Phonetic_Guide_Style_Paragraph:phonetic guide</xsl:message>
        </xsl:when>
        <xsl:when test="name(.)='w:footnoteReference'">
          <xsl:call-template name="footnoteReference"/>
        </xsl:when>
        <xsl:when test="name(.)='w:endnoteReference'">
          <xsl:call-template name="endnoteReference"/>
        </xsl:when>

        <!--2013-03-13，wudi，修复文字表批注BUG，对整个文字表添加批注，start-->
        <xsl:when test ="name(.)='w:commentReference'">
          <xsl:variable name ="commentId">
            <xsl:value-of select ="@w:id"/>
          </xsl:variable>
          <xsl:if test ="(ancestor::w:tr) and (preceding::w:commentRangeEnd[@w:id=$commentId])">
            <字:区域结束_4167>
              <xsl:attribute name="标识符引用_4168">
                <xsl:value-of select="concat('cmt_',$commentId)"/>
              </xsl:attribute>
            </字:区域结束_4167>
          </xsl:if>
        </xsl:when>
        <!--end-->
        
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
        <xsl:when test="name(.)='w:tab'">
          <字:制表符_415E/>
        </xsl:when>

        <!--2014-04-18，wudi，增加对w:ptab的处理，start-->
        <xsl:when test="name(.)='w:ptab'">
          <字:制表符_415E/>
        </xsl:when>
        <!--end-->
        
        <xsl:when test="name(.)='w:footnoteRef' or name(.)='w:endnoteRef'">
          <字:引文符号_4164/>
        </xsl:when>
        <!--杨晓,下面的代码应该放在域中-->
        <xsl:when test="name(.)='w:instrText' or name(.)='w:fldChar'">
          <!--yx,add more conditions,2010.3.1-->
          <xsl:if
            test="(.!='eq \o\ac(')and(.!=',')and(.!=')')and(.!=' ')and(not((./@w:fldCharType='begin')and(contains(./following::*/w:instrText,'HYPERLINK'))))and(not(contains(.,'HYPERLINK')))">
            <xsl:copy-of select="."/>

            <!--2013-03-21，wudi，修复空格符转换互操作BUG，start-->
            <xsl:variable name ="wtxt">
              <xsl:value-of select ="."/>
            </xsl:variable>
            <xsl:variable name ="pos">
              <xsl:value-of select ="number(string-length(.))"/>
            </xsl:variable>
            <xsl:call-template name="text">
              <xsl:with-param name ="pos1" select ="$pos"/>
              <xsl:with-param name ="cnt1" select ="1"/>
              <xsl:with-param name ="wtxt1" select ="$wtxt"/>
            </xsl:call-template>
            <!--end-->

          </xsl:if>
          <!--yx,add one special case below,2010.3.1-->
          <xsl:if
            test="(./@w:fldCharType='begin')and(contains(./following::*/w:instrText[1],'HYPERLINK'))">
            <xsl:call-template name="hlkstart"> </xsl:call-template>
          </xsl:if>
          <xsl:if
            test="(./@w:fldCharType='end')and(contains(./preceding::*/w:instrText[1],'HYPERLINK'))">
            <xsl:call-template name="hlkend"> </xsl:call-template>
          </xsl:if>
          <!--yx,add one special case above,2010.3.1-->
          
        </xsl:when>
        <xsl:when test="name(.)='w:sym'">

          <!--2013-04-18，wudi，修复特殊字符转换BUG，没有枚举全，start-->
          <xsl:if test ="@w:font='Wingdings'">
            <字:文本串_415B>
              <xsl:choose>
                <xsl:when test ="@w:char='F04A'">
                  <xsl:value-of select ="''"/>
                </xsl:when>
              </xsl:choose>
            </字:文本串_415B>
          </xsl:if>
          <!--end-->
          
          <xsl:if
            test="@w:font='Symbol' and not(preceding-sibling::w:footnoteReference) and not(preceding-sibling::w:endnoteReference)">
            <字:文本串_415B>
              
              <!--2013-03-12，wudi，修复自定义脚注-尾注BUG，部分枚举，start-->
              <xsl:choose>
                <xsl:when test ="@w:char='F021'">
                  <xsl:value-of select ="'!'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F023'">
                  <xsl:value-of select ="'#'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F025'">
                  <xsl:value-of select ="'%'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F028'">
                  <xsl:value-of select ="'('"/>
                </xsl:when>
                <xsl:when test ="@w:char='F029'">
                  <xsl:value-of select ="')'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F02A'">
                  <xsl:value-of select ="'*'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F02B'">
                  <xsl:value-of select ="'+'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F02C'">
                  <xsl:value-of select ="','"/>
                </xsl:when>
                <xsl:when test ="@w:char='F02D'">
                  <xsl:value-of select ="'-'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F02E'">
                  <xsl:value-of select ="'.'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F02F'">
                  <xsl:value-of select ="'/'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F030'">
                  <xsl:value-of select ="'0'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F031'">
                  <xsl:value-of select ="'1'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F032'">
                  <xsl:value-of select ="'2'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F033'">
                  <xsl:value-of select ="'3'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F034'">
                  <xsl:value-of select ="'4'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F035'">
                  <xsl:value-of select ="'5'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F036'">
                  <xsl:value-of select ="'6'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F037'">
                  <xsl:value-of select ="'7'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F038'">
                  <xsl:value-of select ="'8'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F039'">
                  <xsl:value-of select ="'9'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F03A'">
                  <xsl:value-of select ="':'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F03B'">
                  <xsl:value-of select ="';'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F03D'">
                  <xsl:value-of select ="'='"/>
                </xsl:when>
                <xsl:when test ="@w:char='F03E'">
                  <xsl:value-of select ="'>'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F03F'">
                  <xsl:value-of select ="'?'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F041'">
                  <xsl:value-of select ="'Α'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F042'">
                  <xsl:value-of select ="'Β'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F043'">
                  <xsl:value-of select ="'Χ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F044'">
                  <xsl:value-of select ="'Δ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F045'">
                  <xsl:value-of select ="'Ε'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F046'">
                  <xsl:value-of select ="'Φ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F047'">
                  <xsl:value-of select ="'Γ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F048'">
                  <xsl:value-of select ="'Η'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F049'">
                  <xsl:value-of select ="'Ι'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F04B'">
                  <xsl:value-of select ="'Κ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F04C'">
                  <xsl:value-of select ="'Λ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F04D'">
                  <xsl:value-of select ="'Μ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F04E'">
                  <xsl:value-of select ="'Ν'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F04F'">
                  <xsl:value-of select ="'Ο'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F050'">
                  <xsl:value-of select ="'Π'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F051'">
                  <xsl:value-of select ="'Θ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F052'">
                  <xsl:value-of select ="'Ρ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F053'">
                  <xsl:value-of select ="'Σ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F054'">
                  <xsl:value-of select ="'Τ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F055'">
                  <xsl:value-of select ="'Υ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F057'">
                  <xsl:value-of select ="'Ω'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F058'">
                  <xsl:value-of select ="'Ξ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F059'">
                  <xsl:value-of select ="'Ψ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F05A'">
                  <xsl:value-of select ="'Ζ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F061'">
                  <xsl:value-of select ="'α'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F062'">
                  <xsl:value-of select ="'β'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F063'">
                  <xsl:value-of select ="'χ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F064'">
                  <xsl:value-of select ="'δ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F065'">
                  <xsl:value-of select ="'ε'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F066'">
                  <xsl:value-of select ="'φ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F067'">
                  <xsl:value-of select ="'γ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F068'">
                  <xsl:value-of select ="'η'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F069'">
                  <xsl:value-of select ="'ι'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F06B'">
                  <xsl:value-of select ="'κ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F06C'">
                  <xsl:value-of select ="'λ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F06D'">
                  <xsl:value-of select ="'μ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F06E'">
                  <xsl:value-of select ="'ν'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F06F'">
                  <xsl:value-of select ="'ο'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F070'">
                  <xsl:value-of select ="'π'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F071'">
                  <xsl:value-of select ="'θ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F072'">
                  <xsl:value-of select ="'ρ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F073'">
                  <xsl:value-of select ="'σ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F074'">
                  <xsl:value-of select ="'τ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F075'">
                  <xsl:value-of select ="'υ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F076'">
                  <xsl:value-of select ="'ϖ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F077'">
                  <xsl:value-of select ="'ω'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F078'">
                  <xsl:value-of select ="'ξ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F079'">
                  <xsl:value-of select ="'ψ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F07A'">
                  <xsl:value-of select ="'ζ'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F0A7'">
                  <xsl:value-of select ="'♣'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F0A8'">
                  <xsl:value-of select ="'♦'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F0A9'">
                  <xsl:value-of select ="'♥'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F0AA'">
                  <xsl:value-of select ="'♠'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F0AC'">
                  <xsl:value-of select ="'←'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F0AD'">
                  <xsl:value-of select ="'↑'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F0AE'">
                  <xsl:value-of select ="'→'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F0AF'">
                  <xsl:value-of select ="'↓'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F0AB'">
                  <xsl:value-of select ="'↔'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F05B'">
                  <xsl:value-of select ="'['"/>
                </xsl:when>
                <xsl:when test ="@w:char='F05D'">
                  <xsl:value-of select ="']'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F07B'">
                  <xsl:value-of select ="'{'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F07C'">
                  <xsl:value-of select ="'|'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F07D'">
                  <xsl:value-of select ="'}'"/>
                </xsl:when>
                <xsl:when test ="@w:char='F07E'">
                  <xsl:value-of select ="'~'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'sym'"/>
                </xsl:otherwise>
              </xsl:choose>
              <!--end-->

            </字:文本串_415B>
          </xsl:if>
        </xsl:when>
  
        <!--2013-03-28，wudi，修复图形，图片转换BUG，start-->
        <xsl:when test="name(.)='w:drawing'">
          
          <xsl:for-each select="wp:inline | wp:anchor">
            <xsl:choose>
              <xsl:when test="name(.)='wp:inline'">

                <!--2013-04-11，针对包含MACROBUTTON的情况，部分图片代码有问题，避免graphics.xml里再出现锚点引用-->
                <xsl:if test="./a:graphic/a:graphicData/pic:pic and not($rPartFrom ='txbody')">
                  <xsl:call-template name="bodyAnchorPic">
                    <xsl:with-param name="picType" select="'inline'"/>
                    <xsl:with-param name="filename" select="$rPartFrom"/>
                  </xsl:call-template>
                </xsl:if>
                
                <!--2013-01-07，wudi，OOX到UOF方向SmartArt实现，start-->
                <xsl:if test="./a:graphic/a:graphicData/wps:wsp or ./a:graphic/a:graphicData/wpg:wgp or ./a:graphic/a:graphicData/dgm:relIds">
                  <xsl:call-template name="bodyAnchorWps">
                    <xsl:with-param name="picType" select="'inline'"/>
                    <xsl:with-param name="filename" select="$rPartFrom"/>
                  </xsl:call-template>
                </xsl:if>
                <!--end-->

                <!--2013-04-19，wudi，增加对chart的转换，start-->
                <xsl:if test ="./a:graphic/a:graphicData/c:chart">
                  <xsl:call-template name="bodyAnchorChart">
                    <xsl:with-param name="picType" select="'inline'"/>
                    <xsl:with-param name="filename" select="$rPartFrom"/>
                  </xsl:call-template>
                </xsl:if>
                <!--end-->
                
              </xsl:when>
              <xsl:when test="name(.)='wp:anchor'">
                
                <!--2013-04-11，针对包含MACROBUTTON的情况，部分图片代码有问题，避免graphics.xml里再出现锚点引用-->
                <xsl:if test="./a:graphic/a:graphicData/pic:pic and not($rPartFrom ='txbody')">
                  <xsl:call-template name="bodyAnchorPic">
                    <xsl:with-param name="picType" select="'anchor'"/>
                    <xsl:with-param name="filename" select="$rPartFrom"/>
                  </xsl:call-template>
                </xsl:if>
                
                <!--2013-01-08，wudi，OOX到UOF方向SmartArt实现，start-->
                <xsl:if test="./a:graphic/a:graphicData/wps:wsp or ./a:graphic/a:graphicData/wpg:wgp or ./a:graphic/a:graphicData/dgm:relIds">
                  <xsl:call-template name="bodyAnchorWps">
                    <xsl:with-param name="picType" select="'anchor'"/>
                    <xsl:with-param name="filename" select="$rPartFrom"/>
                  </xsl:call-template>
                </xsl:if>
                <!--end-->
                
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </xsl:when>
        <!--zhaobj<xsl:when test="name(.)='w:drawing'">-->
        <xsl:when test=".//w:drawing">
          <xsl:for-each select=".//w:drawing/wp:inline | .//w:drawing/wp:anchor">
            
            <xsl:if test ="not(ancestor::mc:Fallback)">
              <xsl:choose>
                <xsl:when test="name(.)='wp:inline'">
                  
                  <!--2013-04-11，针对包含MACROBUTTON的情况，部分图片代码有问题，避免graphics.xml里再出现锚点引用-->
                  <xsl:if test="./a:graphic/a:graphicData/pic:pic and not($rPartFrom ='txbody')">
                    <xsl:call-template name="bodyAnchorPic">
                      <xsl:with-param name="picType" select="'inline'"/>
                      <xsl:with-param name="filename" select="$rPartFrom"/>
                    </xsl:call-template>
                  </xsl:if>

                  <!--2013-01-08，wudi，OOX到UOF方向SmartArt实现，start-->
                  <xsl:if test="./a:graphic/a:graphicData/wps:wsp or ./a:graphic/a:graphicData/wpg:wgp or ./a:graphic/a:graphicData/dgm:relIds">
                    <xsl:call-template name="bodyAnchorWps">
                      <xsl:with-param name="picType" select="'inline'"/>
                      <xsl:with-param name="filename" select="$rPartFrom"/>
                    </xsl:call-template>
                  </xsl:if>
                  <!--end-->

                </xsl:when>
                <xsl:when test="name(.)='wp:anchor'">
                  <!--zhaobj-->
                  
                  <!--2013-04-11，针对包含MACROBUTTON的情况，部分图片代码有问题，避免graphics.xml里再出现锚点引用-->
                  <xsl:if test="./a:graphic/a:graphicData/pic:pic and not($rPartFrom ='txbody')">
                    <xsl:call-template name="bodyAnchorPic">
                      <xsl:with-param name="picType" select="'anchor'"/>
                      <xsl:with-param name="filename" select="$rPartFrom"/>
                    </xsl:call-template>
                  </xsl:if>

                  <!--2013-01-08，wudi，OOX到UOF方向SmartArt实现，start-->
                  <xsl:if test="./a:graphic/a:graphicData/wps:wsp or ./a:graphic/a:graphicData/wpg:wgp or ./a:graphic/a:graphicData/dgm:relIds">
                    <xsl:call-template name="bodyAnchorWps">
                      <xsl:with-param name="picType" select="'anchor'"/>
                      <xsl:with-param name="filename" select="$rPartFrom"/>
                    </xsl:call-template>
                  </xsl:if>
                  <!--end-->

                </xsl:when>
              </xsl:choose>
            </xsl:if>
            
          </xsl:for-each>
        </xsl:when>
        <!--end-->

        <!--2013-04-11，针对包含MACROBUTTON的情况，部分图片代码有问题，避免graphics.xml里再出现锚点引用-->
        <xsl:when test="name(.)='w:pict' and not($rPartFrom ='txbody')">

          <!--2013-04-08，wudi，修复图片转换的BUG，没有考虑节点为w:pict，子节点为v:shape的情况，所有number计数增加对v:shape节点的统计，start-->
          <xsl:variable name="number">
            <xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline | v:shape"/>
          </xsl:variable>
          <xsl:if test ="not(ancestor::w:pict)">
            <xsl:for-each select="v:rect">
              
              <!--2014-03-26，wudi，模板bodyAnchorYdy增加一个参数filename，start-->
              <xsl:call-template name="bodyAnchorYdy">
                <xsl:with-param name="filename" select="$rPartFrom"/>
              </xsl:call-template>
              <!--end-->
              
              <!--Cxl预定义图形锚点模板-->
            </xsl:for-each>
            <xsl:for-each select ="v:shape">
              
              <!--2014-03-26，wudi，模板bodyAnchorYdy增加一个参数filename，start-->
              <xsl:call-template name="bodyAnchorYdy">
                <xsl:with-param name="filename" select="$rPartFrom"/>
              </xsl:call-template>
              <!--end-->
              
            </xsl:for-each>
          </xsl:if>
          <!--end-->
          
        </xsl:when>

        <!--2013-04-19，wudi，增加对对象集-图片，图表对象的转换，start-->
        <xsl:when test ="name(.)='w:object'">
          <xsl:variable name="number">
            <xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline | v:shape"/>
          </xsl:variable>
          <xsl:for-each select ="v:shape">
            
            <!--2014-03-26，wudi，模板bodyAnchorYdy增加一个参数filename，start-->
            <xsl:call-template name="bodyAnchorYdy">
              <xsl:with-param name="filename" select="$rPartFrom"/>
            </xsl:call-template>
            <!--end-->
            
          </xsl:for-each>
        </xsl:when>
        <!--end-->

      </xsl:choose>
    </xsl:for-each>
    <xsl:if test="w:sym">
      <xsl:message terminate="no">feedback:lost:Symbol_Characters_Paragraph:symbol</xsl:message>
    </xsl:if>
    <xsl:if test="w:dayShort|w:monthShort|w:yearShort|w:dayLong|w:monthLong|w:yearLong">
      <xsl:message terminate="no"
        >feedback:lost:day_month_year_information_in_Paragraph:day,month,year</xsl:message>
    </xsl:if>
  </xsl:template>
  <!--*****************************************************************************************************************-->
  <!--yx,add the template:hlkstart,2010.3.1-->
  <xsl:template name="hlkstart">
    <字:区域开始_4165>
      <xsl:variable name="hlkref">
        <xsl:value-of select="concat('hlkref_',generate-id(.))"/>
      </xsl:variable>
      <xsl:attribute name="标识符_4100">
        <xsl:value-of select="$hlkref"/>
      </xsl:attribute>
      <xsl:attribute name="名称_4166">
        <xsl:value-of select="'hyperlink'"/>
      </xsl:attribute>
      <xsl:attribute name="类型_413B">
        <xsl:value-of select="'hyperlink'"/>
      </xsl:attribute>
      <!--xsl:if test="@r:id">
          <xsl:variable name="rid">
            <xsl:value-of select="@r:id"/>
          </xsl:variable>
          <xsl:variable name ="goal">
            <xsl:choose>
              <xsl:when test ="$filename='document'">
                <xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
              </xsl:when>
              <xsl:when test ="$filename='comments'">
                <xsl:value-of select="document('word/_rels/comments.xml.rels')/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
              </xsl:when>
              <xsl:when test ="$filename='endnotes'">
                <xsl:value-of select="document('word/_rels/endnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
              </xsl:when>
              <xsl:when test ="$filename='footnotes'">
                <xsl:value-of select="document('word/_rels/footnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
              </xsl:when>
              <xsl:when test ="contains($filename,'header')">
                <xsl:variable name ="doc">
                  <xsl:value-of select ="concat('word/_rels/',$filename,'.rels')"/>
                </xsl:variable>
                <xsl:value-of select="document($doc)/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
              </xsl:when>
              <xsl:when test ="contains($filename,'footer')">
                <xsl:variable name ="doc">
                  <xsl:value-of select ="concat('word/_rels/',$filename,'.rels')"/>
                </xsl:variable>
                <xsl:value-of select="document($doc)/rel:Relationships/rel:Relationship[@Id = $rid]/@Target"/>
              </xsl:when>
            </xsl:choose>
          </xsl:variable>
          <xsl:attribute name="uof:目标">
            <xsl:value-of select="$goal"/>
          </xsl:attribute>
        </xsl:if>
      <xsl:if test="@w:anchor">
        <超链:书签_AA0D>
          <xsl:value-of select="./@w:anchor"/>
        </超链:书签_AA0D>
      </xsl:if>
      <xsl:if test="@w:tooltip">
        <超链:提示_AA05>
          <xsl:value-of select="./@w:tooltip"/>
        </超链:提示_AA05>
      </xsl:if>
      <xsl:variable name="hlkgoal">
        <xsl:value-of
          select="substring-after(normalize-space(./following::*/w:instrText[1]),' &quot;')"/>
      </xsl:variable>
      <xsl:variable name="goal">
        <xsl:value-of select="substring-before($hlkgoal,'&quot;')"/>
      </xsl:variable>
      <超链:目标_AA01>
        <xsl:value-of select="$goal"/>
      </超链:目标_AA01>
      <超链:链源_AA00>
        <xsl:value-of select="concat('hlkref_',generate-id(.))"/>
      </超链:链源_AA00>-->
    </字:区域开始_4165>
  </xsl:template>
  <!--yx,add the template:hlkend,2010.3.1-->
  <xsl:template name="hlkend">
    <字:区域结束_4167>
      <xsl:attribute name="标识符引用_4168">
        <!--./preceding::*/w:fldChar[@w:fldCharType='begin']-->
        <xsl:value-of
          select="concat('hlkref_',generate-id(ancestor::w:p/preceding-sibling::w:p[w:r/w:fldChar/@w:fldCharType='begin'][1]/w:r/w:fldChar[@w:fldCharType='begin']))"
        />
      </xsl:attribute>
    </字:区域结束_4167>
  </xsl:template>
  
  <!--2013-03-21，修复空格符转换互操作的BUG，start-->
  <xsl:template name="text"><!--转换文本串模板-->
    <xsl:param name ="pos1"/>
    <xsl:param name ="cnt1"/>
    <xsl:param name ="wtxt1"/>
    <!--<调用开始></调用开始>
    <pos1>
      <xsl:value-of select ="$pos1"/>
    </pos1>-->
    <xsl:if test ="$pos1&gt;=1">
      <xsl:variable name="lgth">
        <xsl:value-of select ="string-length(substring($wtxt1,$pos1,$pos1))"/>
      </xsl:variable>
      <xsl:variable name ="tmp">
        <xsl:value-of select ="string-length(substring($wtxt1,$pos1,1))"/>
      </xsl:variable>
      
      <!--2013-03-27，wudi，优化空格符转换的处理方式，start-->
      <xsl:if test ="substring($wtxt1,$pos1,1) =' '">
        <xsl:variable name ="pos2">
            <xsl:value-of select ="number($pos1)-1"/>
        </xsl:variable>
        <xsl:variable name ="cnt2">
            <xsl:value-of select ="number($cnt1)+1"/>
        </xsl:variable>
        <xsl:variable name ="wtxt2">
          <xsl:value-of select ="$wtxt1"/>
        </xsl:variable>
        <xsl:call-template name ="text">
          <xsl:with-param name ="pos1" select ="$pos2"/>
          <xsl:with-param name ="cnt1" select ="$cnt2"/>
          <xsl:with-param name ="wtxt1" select ="$wtxt2"/>
        </xsl:call-template>
      </xsl:if>
      <!--end-->
      
      <!--<xsl:if test ="substring($wtxt1,number($pos1),number($pos1)) =' ' or $lgth&gt;1">
        <xsl:variable name ="pos2">
          <xsl:if test ="substring($wtxt1,number($pos1),number($pos1)) =' '">
            <xsl:value-of select ="number($pos1)-1"/>
          </xsl:if>
          <xsl:if test ="$lgth&gt;1">
            <xsl:value-of select ="number($pos1)-$lgth"/>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name ="cnt2">
          <xsl:if test ="substring($wtxt1,number($pos1),number($pos1)) =' '">
            <xsl:value-of select ="number($cnt1)+1"/>
          </xsl:if>
          <xsl:if test ="$lgth&gt;1">
            <xsl:value-of select ="number($cnt1)+$lgth"/>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name ="wtxt2">
          <xsl:value-of select ="$wtxt1"/>
        </xsl:variable>
        <xsl:call-template name ="text">
          <xsl:with-param name ="pos1" select ="$pos2"/>
          <xsl:with-param name ="cnt1" select ="$cnt2"/>
          <xsl:with-param name ="wtxt1" select ="$wtxt2"/>
        </xsl:call-template>
      </xsl:if>-->
    </xsl:if>
    <xsl:if test ="(number($cnt1) =number(string-length($wtxt1)) + 1)">
      <字:空格符_4161>
        <xsl:attribute name ="个数_4162">
          <xsl:value-of select ="number($cnt1) - 1"/>
        </xsl:attribute>
      </字:空格符_4161>
    </xsl:if>
    <!--<测试01>
      <xsl:value-of select ="number($pos1 =1)"/>
    </测试01>
    <测试02>
      <xsl:value-of select ="number($pos1&gt;1)"/>
    </测试02>
    <测试03>
      <xsl:value-of select ="number(not(starts-with($wtxt1,' ')))"/>
    </测试03>
    <测试04>
      <xsl:value-of select ="number(substring($wtxt1,$pos1,1) !=' ')"/>
    </测试04>-->
    <xsl:if test ="(number($pos1 =1) + number($cnt1 =1) + number(not(starts-with($wtxt1,' '))))&gt;=3">
      <字:文本串_415B>
        <xsl:value-of select ="$wtxt1"/>
      </字:文本串_415B>
    </xsl:if>
    <xsl:if test ="(number(not(starts-with($wtxt1,' '))) + number($cnt1 =1) + number($pos1&gt;1))&gt;=3">
      <字:文本串_415B>
        <xsl:value-of select ="$wtxt1"/>
      </字:文本串_415B>
    </xsl:if>
    
    <!--2013-04-02，wudi，修复文本串转换丢失的BUG，没有考虑开头第一个字符为空格符的情况，start-->
    <xsl:if test ="(number(substring($wtxt1,$pos1,1) !=' ') + number(starts-with($wtxt1,' ')) + number($cnt1 =1) + number($pos1&gt;1))&gt;=4">
      <字:文本串_415B>
        <xsl:value-of select ="$wtxt1"/>
      </字:文本串_415B>
    </xsl:if>
    <!--end-->
    
    <!--2013-04-03，wudi，修复文本串转换丢失的BUG，没有考虑开头和结尾第一个字符均为空格符的情况，start-->
    <!--2013-04-10，wudi，增加限制条件number(substring($wtxt1,$pos1,1) =' ')-->
    <xsl:if test ="(number(substring($wtxt1,$pos1 -1,1) !=' ') + number(starts-with($wtxt1,' ') + number(substring($wtxt1,$pos1,1) =' ')) + number($cnt1 =1) + number($pos1&gt;1))&gt;=5">
      <字:文本串_415B>
        <xsl:value-of select ="$wtxt1"/>
      </字:文本串_415B>
    </xsl:if>
    <!--<调用结束></调用结束>-->
  </xsl:template>
  <!--end-->
  
  <xsl:template match="w:rPr" mode="RunProperties"><!--转换句中除修订信息外的其它内容的模板-->
    
    <!--2014-05-05，wudi，增加postr参数，记录单元格所属行数，修复单元格字体加粗效果转换BUG，start-->
    <xsl:param name="postr"/>
    <xsl:param name="postc"/>
    <!--end-->
    
    <!--2014-03-27，wudi，增加w:szCs，处理存在w:szCs节点没有w:sz节点的情况，start-->
    <xsl:apply-templates
      select="w:rStyle|w:rFonts|w:color|w:sz|w:szCs|w:b|w:i
                                     |w:caps|w:smallCaps|w:strike|w:dstrike
                                     |w:outline|w:shadow|w:emboss|w:imprint
                                     |w:imprint|w:snapToGrid|w:vanish|w:spacing
                                     |w:w|w:kern|w:position|w:highlight|w:u|w:bdr
                                     |w:shd|w:vertAlign|w:em|w14:shadow"
      mode="rPrChildren"/>
    <!--end-->
    
    <!--cxl,2012.3.6双行合一，未考虑加括号的判断-->
    
    <!--2014-03-13，wudi，修改标识符的取值方式，start-->
    <xsl:if test="w:eastAsianLayout/@w:combine='1' or w:eastAsianLayout/@w:combine='on'">
      <字:双行合一_4148>
        <xsl:attribute name="标识符_4149">
          <xsl:value-of select="w:eastAsianLayout/@w:id"/>
        </xsl:attribute>
        <xsl:attribute name="前置字符_414A">
          <xsl:value-of select="''"/>
        </xsl:attribute>
        <xsl:attribute name="后置字符_414B">
          <xsl:value-of select="''"/>
        </xsl:attribute>
        <!--<xsl:if test="w:eastAsianLayout/@w:combineBrackets='round'">
          <xsl:attribute name="前置字符_414A">--><!--这个地方不全面吧！后置字符_414B呢？@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@--><!--
            <xsl:value-of select="'('"/>
          </xsl:attribute>         
        </xsl:if>-->
      </字:双行合一_4148>
    </xsl:if>
    <!--end-->

    <!--2014-11-24，wudi，注释掉以下代码，解决性能问题，可能会带来表格样式的问题，后面处理，start-->
    
    <!--2014-05-05，wudi，修复单元格字体转换前后不一致的BUG，start-->
    <!--2014-05-08，wudi，修复单元格字体颜色转换前后不一致的BUG，start-->
    <!--<xsl:if test ="ancestor::w:tbl/w:tblPr/w:tblStyle and not(ancestor::w:p/w:pPr/w:pStyle or ./w:rStyle) and not(w:sz or w:szCs or w:color) and (not(w:rFonts) or w:rFonts[@w:hint ='eastAsia'])">
      <xsl:variable name ="styId">
        <xsl:value-of select ="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
      </xsl:variable>
      
      --><!--2014-05-19，wudi，注释掉此段代码，可能带来新的问题，start--><!--
      --><!--<xsl:if test="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:rFonts">
        <xsl:apply-templates select="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:rFonts" mode="rPrChildren"/>
      </xsl:if>--><!--
      --><!--end--><!--

      <xsl:if test ="not(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:rFonts) and (document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:color)">
        <字:字体_4128>
          --><!--uof:attrList="西文字体引用 中文字体引用 特殊字体引用 西文绘制 字号 相对字号 颜色"--><!--
          <xsl:attribute name="颜色_412F">
            <xsl:if test="not(@w:val='auto')">
              <xsl:value-of select="'#'"/>
            </xsl:if>
            <xsl:value-of select="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:color/@w:val"/>
          </xsl:attribute>
        </字:字体_4128>
      </xsl:if>

    </xsl:if>-->
    <!--end-->
    <!--end-->
    
    <!--end-->

    <!--2014-03-27，wudi，处理字号转换的BUG，针对软件差异，start-->
    <xsl:if test ="./w:rStyle and not(w:sz or w:szCs or w:rFonts or w:color)">
      <xsl:variable name ="styId">
        <xsl:value-of select ="./w:rStyle/@w:val"/>
      </xsl:variable>
      <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:sz or document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:szCs">
        <xsl:choose>
          <xsl:when test="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:sz">
            <xsl:variable name ="size">
              <xsl:value-of select ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:sz/@w:val"/>
            </xsl:variable>
            <字:字体_4128>
              <xsl:attribute name ="字号_412D">
                <xsl:value-of select ="format-number($size div 2,'0.0')"/>
              </xsl:attribute>
            </字:字体_4128>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name ="size">
              <xsl:value-of select ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:szCs/@w:val"/>
            </xsl:variable>
            <字:字体_4128>
              <xsl:attribute name ="字号_412D">
                <xsl:value-of select ="format-number($size div 2,'0.0')"/>
              </xsl:attribute>
            </字:字体_4128>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </xsl:if>
    <!--end-->

    <!--2014-11-24，wudi，注释掉以下代码，解决性能问题，可能会带来表格样式的问题，后面处理，start-->
    
    <!--2013-03-26，wudi，修复表格文字内容加粗效果丢失BUG，之前没有考虑无w:b节点的情况，start-->
    <xsl:if test ="./w:rStyle and not(w:b)">
      <xsl:variable name ="styId">
        <xsl:value-of select ="./w:rStyle/@w:val"/>
      </xsl:variable>
      <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:b">
        <字:是否粗体_4130>
          <xsl:choose>
            <xsl:when test="not(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:b/@w:val) or (document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:b/@w:val='on')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:b/@w:val='1')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:rPr/w:b/@w:val='true')">
              <xsl:value-of select="'true'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'false'"/>
            </xsl:otherwise>
          </xsl:choose>
        </字:是否粗体_4130>
      </xsl:if>
    </xsl:if>
    <!--end-->
    
    <!--2013-04-11，wudi，修复表格文字内容加粗效果丢失的BUG，考虑没有w:rStyle,又没有w:b的情况，start-->
    <xsl:variable name ="styId">
      <xsl:value-of select ="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
    </xsl:variable>

    <!--2014-05-05，wudi，修复表格字体颜色转换前后不一致的BUG，start-->
    <!--<xsl:if test ="(number(not(./w:rStyle)) + number(not(w:color)) + number(preceding::w:cnfStyle[1]/@w:firstRow ='1' or ancestor::w:tr/w:trPr/w:cnfStyle/@w:firstRow ='1') + number($postr = '1') + number(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr/@w:type ='firstRow'))&gt;=5">
      <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:rPr/w:color">
        <字:字体_4128>
          --><!--uof:attrList="西文字体引用 中文字体引用 特殊字体引用 西文绘制 字号 相对字号 颜色"--><!--
          <xsl:attribute name="颜色_412F">
            <xsl:if test="not(@w:val='auto')">
              <xsl:value-of select="'#'"/>
            </xsl:if>
            <xsl:value-of select="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:rPr/w:color/@w:val"/>
          </xsl:attribute>
        </字:字体_4128>
      </xsl:if>
    </xsl:if>-->
    <!--end-->

    <!--2014-05-09，wudi，修复表格字体字号转换前后不一致的BUG，start-->
    <!--<xsl:if test ="(number(not(./w:rStyle)) + number(not(w:sz or w:szCs or w:rFonts or w:color)) + number(preceding::w:cnfStyle[1]/@w:firstRow ='1' or ancestor::w:tr/w:trPr/w:cnfStyle/@w:firstRow ='1') + number($postr = '1') + number(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr/@w:type ='firstRow'))&gt;=5">
      <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:rPr/w:sz or document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:rPr/w:szCs">
        <xsl:choose>
          <xsl:when test="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:rPr/w:sz">
            <xsl:variable name ="size">
              <xsl:value-of select ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:rPr/w:sz/@w:val"/>
            </xsl:variable>
            <字:字体_4128>
              <xsl:attribute name ="字号_412D">
                <xsl:value-of select ="format-number($size div 2,'0.0')"/>
              </xsl:attribute>
            </字:字体_4128>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name ="size">
              <xsl:value-of select ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:rPr/w:szCs/@w:val"/>
            </xsl:variable>
            <字:字体_4128>
              <xsl:attribute name ="字号_412D">
                <xsl:value-of select ="format-number($size div 2,'0.0')"/>
              </xsl:attribute>
            </字:字体_4128>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </xsl:if>-->
    <!--end-->

    <!--2014-05-05，wudi，增加postr参数，记录单元格所属行数，修复单元格字体加粗效果转换BUG，start-->
    <!--<xsl:if test ="(number(not(./w:rStyle)) + number(not(w:b)) + number(preceding::w:cnfStyle[1]/@w:firstRow ='1') + number($postr = '1') + number(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr/@w:type ='firstRow'))&gt;=5">
      <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:rPr/w:b">
        <字:是否粗体_4130>
          <xsl:choose>
            <xsl:when test="not(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:rPr/w:b/@w:val) or (document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:rPr/w:b/@w:val='on')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:rPr/w:b/@w:val='1')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:rPr/w:b/@w:val='true')">
              <xsl:value-of select="'true'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'false'"/>
            </xsl:otherwise>
          </xsl:choose>
        </字:是否粗体_4130>
      </xsl:if>
    </xsl:if>-->
    <!--end-->
    
    <xsl:if test ="(number(not(./w:rStyle)) + number(not(w:b)) + number(preceding::w:cnfStyle[1]/@w:firstColumn ='1') + number(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr/@w:type ='firstCol'))&gt;=4">
      <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstCol']/w:rPr/w:b">
        <字:是否粗体_4130>
          <xsl:choose>
            <xsl:when test="not(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstCol']/w:rPr/w:b/@w:val) or (document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstCol']/w:rPr/w:b/@w:val='on')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstCol']/w:rPr/w:b/@w:val='1')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstCol']/w:rPr/w:b/@w:val='true')">
              <xsl:value-of select="'true'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'false'"/>
            </xsl:otherwise>
          </xsl:choose>
        </字:是否粗体_4130>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="(number(not(./w:rStyle)) + number(not(w:b)) + number(preceding::w:cnfStyle[1]/@w:lastColumn ='1') + number(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr/@w:type ='lastCol'))&gt;=4">
      <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='lastCol']/w:rPr/w:b">
        <字:是否粗体_4130>
          <xsl:choose>
            <xsl:when test="not(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='lastCol']/w:rPr/w:b/@w:val) or (document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='lastCol']/w:rPr/w:b/@w:val='on')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='lastCol']/w:rPr/w:b/@w:val='1')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='lastCol']/w:rPr/w:b/@w:val='true')">
              <xsl:value-of select="'true'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'false'"/>
            </xsl:otherwise>
          </xsl:choose>
        </字:是否粗体_4130>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="(number(not(./w:rStyle)) + number(not(w:b)) + number(preceding::w:cnfStyle[1]/@w:lastRow ='1') + number(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr/@w:type ='lastRow'))&gt;=4">
      <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='lastRow']/w:rPr/w:b">
        <字:是否粗体_4130>
          <xsl:choose>
            <xsl:when test="not(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='lastRow']/w:rPr/w:b/@w:val) or (document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='lastRow']/w:rPr/w:b/@w:val='on')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='lastRow']/w:rPr/w:b/@w:val='1')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='lastRow']/w:rPr/w:b/@w:val='true')">
              <xsl:value-of select="'true'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'false'"/>
            </xsl:otherwise>
          </xsl:choose>
        </字:是否粗体_4130>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="(number(not(./w:rStyle)) + number(not(w:b)) + number(preceding::w:cnfStyle[1]/@oddVBand ='1') + number(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr/@w:type ='band1Vert'))&gt;=4">
      <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='band1Vert']/w:rPr/w:b">
        <字:是否粗体_4130>
          <xsl:choose>
            <xsl:when test="not(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='band1Vert']/w:rPr/w:b/@w:val) or (document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='band1Vert']/w:rPr/w:b/@w:val='on')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='band1Vert']/w:rPr/w:b/@w:val='1')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='band1Vert']/w:rPr/w:b/@w:val='true')">
              <xsl:value-of select="'true'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'false'"/>
            </xsl:otherwise>
          </xsl:choose>
        </字:是否粗体_4130>
      </xsl:if>
    </xsl:if>
    <xsl:if test ="(number(not(./w:rStyle)) + number(not(w:b)) + number(preceding::w:cnfStyle[1]/@w:oddHBand ='1') + number(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr/@w:type ='band1Horz'))&gt;=4">
      <xsl:if test ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='band1Horz']/w:rPr/w:b">
        <字:是否粗体_4130>
          <xsl:choose>
            <xsl:when test="not(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='band1Horz']/w:rPr/w:b/@w:val) or (document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='band1Horz']/w:rPr/w:b/@w:val='on')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='band1Horz']/w:rPr/w:b/@w:val='1')or(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='band1Horz']/w:rPr/w:b/@w:val='true')">
              <xsl:value-of select="'true'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'false'"/>
            </xsl:otherwise>
          </xsl:choose>
        </字:是否粗体_4130>
      </xsl:if>
    </xsl:if>
    <!--end-->
    
    <!--end-->
    
    <xsl:if test="w:effect|w:oMath">
      <xsl:message terminate="no">feedback:lost:Text_Effect_or_Math_Formula_in_Paragraph:text effect
        or math formula</xsl:message>
    </xsl:if>
  </xsl:template>
  <xsl:template match="w:rStyle" mode="rPrChildren">
    <xsl:attribute name="式样引用_417B">
      <xsl:call-template name="IdProducer">
        <xsl:with-param name="ooxId" select="@w:val"/>
      </xsl:call-template>
    </xsl:attribute>

    <!--2013-03-09，wudi，修复脚注-尾注转换编号上标样式丢失BUG，start-->
    <xsl:if test ="@w:val='EndnoteReference' or @w:val='FootnoteReference'">
      <字:上下标类型_4143>sup</字:上下标类型_4143>
    </xsl:if>
    <!--end-->
    
  </xsl:template>
  <xsl:template match="w:rFonts" mode="rPrChildren">
    <字:字体_4128><!--uof:attrList="西文字体引用 中文字体引用 特殊字体引用 西文绘制 字号 相对字号 颜色"-->

      <!--2014-04-16，修复字体转换BUG，增加此属性，处理空格符转换前后显示不一样的问题，start-->
      <xsl:choose>
        <xsl:when test="@w:hint = 'eastAsia'">
          <xsl:attribute name="是否西文绘制_412C">
            <xsl:value-of select="'false'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="是否西文绘制_412C">
            <xsl:value-of select="'true'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <!--end-->
      
      <xsl:if test="@w:ascii">
        <xsl:attribute name="西文字体引用_4129">
          <xsl:value-of select="translate(@w:ascii,' ','')"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@w:eastAsia">
        <xsl:attribute name="中文字体引用_412A">
          <xsl:value-of select="translate(@w:eastAsia,' ','')"/>
        </xsl:attribute>
      </xsl:if>
      
      <!--2014-04-11，wudi，修复西文字体引用转换BUG，start-->
      <xsl:if test="(not(@w:ascii))and(@w:asciiTheme)">

        <xsl:variable name="themefontlang"
      select="document('word/settings.xml')/w:settings/w:themeFontLang/@w:eastAsia"/>

        <xsl:if test="starts-with(@w:asciiTheme,'minor')">
          <xsl:attribute name="西文字体引用_4129">
            <xsl:if test ="starts-with(@w:asciiTheme,'minorEastAsia')">
              <xsl:choose>
                <xsl:when test="$themefontlang='zh-CN'">
                  <xsl:variable name="themefont"
                  select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:font[@script='Hans']/@typeface"/>
                  <xsl:value-of select="translate($themefont,' ','')"/>
                </xsl:when>
                <xsl:when test="$themefontlang='ko-KR'">
                  <xsl:variable name="themefont"
                  select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:font[@script='Hang']/@typeface"/>
                  <xsl:value-of select="translate($themefont,' ','')"/>
                </xsl:when>
                <xsl:when test="$themefontlang='ja-JP'">
                  <xsl:variable name="themefont"
                  select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:font[@script='Jpan']/@typeface"/>
                  <xsl:value-of select="translate($themefont,' ','')"/>
                </xsl:when>
                <xsl:when test="$themefontlang='zh-TW'">
                  <xsl:variable name="themefont"
                  select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:font[@script='Hant']/@typeface"/>
                  <xsl:value-of select="translate($themefont,' ','')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name="themefont"
                  select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
                  <xsl:value-of select="translate($themefont,' ','')"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
            <xsl:if test ="starts-with(@w:asciiTheme,'minorHAnsi')">
              <xsl:variable name="themefont"
                  select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
              <xsl:value-of select="translate($themefont,' ','')"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="starts-with(@w:asciiTheme,'major')">
          <xsl:attribute name="西文字体引用_4129">
            <xsl:if test ="starts-with(@w:asciiTheme,'majorEastAsia')">
              <xsl:choose>
                <xsl:when test="$themefontlang='zh-CN'">
                  <xsl:variable name="themefont"
                  select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:majorFont/a:font[@script='Hans']/@typeface"/>
                  <xsl:value-of select="translate($themefont,' ','')"/>
                </xsl:when>
                <xsl:when test="$themefontlang='ko-KR'">
                  <xsl:variable name="themefont"
                  select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:majorFont/a:font[@script='Hang']/@typeface"/>
                  <xsl:value-of select="translate($themefont,' ','')"/>
                </xsl:when>
                <xsl:when test="$themefontlang='ja-JP'">
                  <xsl:variable name="themefont"
                  select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:majorFont/a:font[@script='Jpan']/@typeface"/>
                  <xsl:value-of select="translate($themefont,' ','')"/>
                </xsl:when>
                <xsl:when test="$themefontlang='zh-TW'">
                  <xsl:variable name="themefont"
                  select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:majorFont/a:font[@script='Hant']/@typeface"/>
                  <xsl:value-of select="translate($themefont,' ','')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name="themefont"
                  select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
                  <xsl:value-of select="translate($themefont,' ','')"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
            <xsl:if test ="starts-with(@w:asciiTheme,'majorHAnsi')">
              <xsl:variable name="themefont"
                  select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
              <xsl:value-of select="translate($themefont,' ','')"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:if>
      </xsl:if>
      <!--end-->
      
      <xsl:if test="(not(@w:eastAsia))and(@w:eastAsiaTheme)">
        <xsl:call-template name="eastAsiaFontTheme"/>
      </xsl:if>
      <xsl:for-each select="following-sibling::w:color">
        <xsl:attribute name="颜色_412F">
          <xsl:if test="not(@w:val='auto')">
            <xsl:value-of select="'#'"/>
          </xsl:if>
          <xsl:value-of select="@w:val"/>
        </xsl:attribute>
      </xsl:for-each>
      <xsl:for-each select="following-sibling::w:sz">
        <xsl:attribute name="字号_412D">
          <xsl:value-of select="format-number(number(@w:val) div 2,'0.0')"/>
        </xsl:attribute>
      </xsl:for-each>

      <!--2013-04-18，wudi，修复特殊字符转换BUG，start-->
      <xsl:if test ="ancestor::w:r/w:sym">
        <xsl:attribute name ="特殊字体引用_412B">
          <xsl:value-of select ="ancestor::w:r/w:sym/@w:font"/>
        </xsl:attribute>
      </xsl:if>
      <!--end-->

      <!--2013-04-01，wudi，修复字体字号转换的BUG，start-->
      <xsl:if test ="not(following-sibling::w:sz)">
        <xsl:for-each select ="following-sibling::w:szCs">
          <xsl:attribute name="字号_412D">
            <xsl:value-of select="format-number(number(@w:val) div 2,'0.0')"/>
          </xsl:attribute>
        </xsl:for-each>
      </xsl:if>
      <!--end-->
      
    </字:字体_4128>
  </xsl:template>
  <xsl:template name="eastAsiaFontTheme">
    <xsl:variable name="themefontlang"
      select="document('word/settings.xml')/w:settings/w:themeFontLang/@w:eastAsia"/>
    <xsl:if test="starts-with(@w:eastAsiaTheme,'major')">
      <xsl:choose>
        <xsl:when test="$themefontlang='zh-CN'">
          <xsl:call-template name="subForFont">
            <xsl:with-param name="theme" select="'a:majorFont'"/>
            <xsl:with-param name="script" select="'Hans'"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$themefontlang='ko-KR'">
          <xsl:call-template name="subForFont">
            <xsl:with-param name="theme" select="'a:majorFont'"/>
            <xsl:with-param name="script" select="'Hang'"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$themefontlang='ja-JP'">
          <xsl:call-template name="subForFont">
            <xsl:with-param name="theme" select="'a:majorFont'"/>
            <xsl:with-param name="script" select="'Jpan'"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$themefontlang='zh-TW'">
          <xsl:call-template name="subForFont">
            <xsl:with-param name="theme" select="'a:majorFont'"/>
            <xsl:with-param name="script" select="'Hant'"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:message terminate="no">feedback:lost:Font_Style:theme font</xsl:message>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="starts-with(@w:eastAsiaTheme,'minor')">
      <xsl:choose>
        <xsl:when test="$themefontlang='zh-CN'">
          <xsl:call-template name="subForFont">
            <xsl:with-param name="theme" select="'a:minorFont'"/>
            <xsl:with-param name="script" select="'Hans'"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$themefontlang='ko-KR'">
          <xsl:call-template name="subForFont">
            <xsl:with-param name="theme" select="'a:minorFont'"/>
            <xsl:with-param name="script" select="'Hang'"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$themefontlang='ja-JP'">
          <xsl:call-template name="subForFont">
            <xsl:with-param name="theme" select="'a:minorFont'"/>
            <xsl:with-param name="script" select="'Jpan'"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="$themefontlang='zh-TW'">
          <xsl:call-template name="subForFont">
            <xsl:with-param name="theme" select="'a:minorFont'"/>
            <xsl:with-param name="script" select="'Hant'"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:message terminate="no">feedback:lost:Theme_Font_Style:theme font</xsl:message>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>
  <xsl:template name="subForFont">
    <xsl:param name="script"/>
    <xsl:param name="theme"/>
    <xsl:if test="$theme='a:majorFont'">
      <xsl:variable name="majorfont"
        select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:majorFont/a:font[@script=$script]/@typeface"/>
      <xsl:attribute name="中文字体引用_412A">
        <xsl:value-of select="translate($majorfont,' ','')"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="$theme='a:minorFont'">
      <xsl:variable name="minorfont"
        select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:font[@script=$script]/@typeface"/>
      <xsl:attribute name="中文字体引用_412A">
        <xsl:value-of select="translate($minorfont,' ','')"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
  <xsl:template name="fontColor" match="w:color" mode="rPrChildren">
    <xsl:if test="not(preceding-sibling::w:rFonts)">
      <字:字体_4128><!--uof:attrList="西文字体引用 中文字体引用 特殊字体引用 西文绘制 字号 相对字号 颜色"--> 
        <xsl:attribute name="颜色_412F">
          <xsl:if test="not(@w:val='auto')">
            <xsl:value-of select="'#'"/>
          </xsl:if>
          <xsl:value-of select="@w:val"/>
        </xsl:attribute>

        <!--2013-04-18，wudi，修复特殊字符转换BUG，start-->
        <xsl:if test ="ancestor::w:r/w:sym">
          <xsl:attribute name ="特殊字体引用_412B">
            <xsl:value-of select ="ancestor::w:r/w:sym/@w:font"/>
          </xsl:attribute>
        </xsl:if>
        <!--end-->
        
        <xsl:for-each select="following-sibling::w:sz">
          <xsl:attribute name="字号_412D">
            <xsl:value-of select="format-number(number(@w:val) div 2,'0.0')"/>
          </xsl:attribute>
        </xsl:for-each>
      </字:字体_4128>
    </xsl:if>
  </xsl:template>
  <xsl:template name="fontSize" match="w:sz" mode="rPrChildren">
    <xsl:if test="not(preceding-sibling::w:rFonts or preceding-sibling::w:color)">
      <字:字体_4128>
        <xsl:attribute name="字号_412D">
          <xsl:value-of select="format-number(number(@w:val) div 2,'0.0')"/>
        </xsl:attribute>
      </字:字体_4128>
    </xsl:if>
  </xsl:template>

  <!--2014-03-27，wudi，处理存在w:szCs节点没有w:sz节点的情况，start-->
  <xsl:template match="w:szCs" mode="rPrChildren">
    <xsl:if test="not(preceding-sibling::w:rFonts or preceding-sibling::w:color or preceding-sibling::w:sz)">
      <字:字体_4128>
        <xsl:attribute name="字号_412D">
          <xsl:value-of select="format-number(number(@w:val) div 2,'0.0')"/>
        </xsl:attribute>
      </字:字体_4128>
    </xsl:if>
  </xsl:template>
  <!--end-->
  
  <xsl:template match="w:b" mode="rPrChildren">
    <字:是否粗体_4130>
        <xsl:choose>
          <xsl:when test="not(@w:val) or (@w:val='on')or(@w:val='1')or(@w:val='true')">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'false'"/>
          </xsl:otherwise>
        </xsl:choose>
    </字:是否粗体_4130>
  </xsl:template>
  <xsl:template match="w:i" mode="rPrChildren">
    <字:是否斜体_4131>
        <xsl:choose>
          <xsl:when test="not(@w:val) or (@w:val='on')or(@w:val='1')or(@w:val='true')">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'false'"/>
          </xsl:otherwise>
        </xsl:choose>
    </字:是否斜体_4131>
  </xsl:template>
  <xsl:template match="w:caps" mode="rPrChildren">
    <xsl:if test="not(@w:val) or @w:val='on' or @w:val='1' or @w:val='true'">
      <字:醒目字体类型_4141><!--uof:attrList="类型"-->
        <xsl:value-of select="'uppercase'"/>
      </字:醒目字体类型_4141>
    </xsl:if>
  </xsl:template>
  <xsl:template match="w:smallCaps" mode="rPrChildren">
    <xsl:if test="not(@w:val) or @w:val='on' or @w:val='1' or @w:val='true'">
      <字:醒目字体类型_4141>
        <xsl:value-of select="'small-caps'"/>
      </字:醒目字体类型_4141>
    </xsl:if>
  </xsl:template>
  
  <!--2013-05-10，吴迪，修复拼音指南转换BUG，拼音指南变为下划线，start-->
  <xsl:template match="w:strike" mode="rPrChildren">

    <xsl:if test ="not(ancestor::w:p/w:r/w:ruby)">
      <xsl:if test="(not(@w:val)) or (@w:val='1')or(@w:val='on')or(@w:val='true')">
        <字:删除线_4135>
          <!--uof:attrList="类型"-->
          <!--<xsl:attribute name="字:类型">-->
          <xsl:value-of select="'single'"/>
          <!--</xsl:attribute>-->
        </字:删除线_4135>
      </xsl:if>
    </xsl:if>
    
  </xsl:template>
  <xsl:template match="w:dstrike" mode="rPrChildren">

    <xsl:if test ="not(ancestor::w:p/w:r/w:ruby)">
      <xsl:if test="(not(@w:val)) or (@w:val='1')or(@w:val='on')or(@w:val='true')">
        <字:删除线_4135>
          <xsl:value-of select="'double'"/>
        </字:删除线_4135>
      </xsl:if>
    </xsl:if>

  </xsl:template>
  <!--end-->
  
  <xsl:template match="w:outline" mode="rPrChildren">
    <字:是否空心_413E><!-- uof:attrList="值"-->
      <!--<xsl:attribute name="字:值">-->
        <xsl:choose>
          <xsl:when test="not(@w:val) or (@w:val='on')or(@w:val='1')or(@w:val='true')">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'false'"/>
          </xsl:otherwise>
        </xsl:choose>
      <!--</xsl:attribute>-->
    </字:是否空心_413E>
  </xsl:template>
  <xsl:template match="w14:shadow" mode="rPrChildren">
    <字:是否阴影_4140>
        <xsl:choose>
          <xsl:when test="not(@w:val) or (@w:val='on')or(@w:val='1')or(@w:val='true')">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'false'"/>
          </xsl:otherwise>
        </xsl:choose>
    </字:是否阴影_4140>
  </xsl:template>
  <xsl:template match="w:emboss" mode="rPrChildren">
    <xsl:if test="not(@w:val) or (@w:val='on')or(@w:val='1')or(@w:val='true')">
      <字:浮雕_413F>
        <xsl:value-of select="'emboss'"/>
      </字:浮雕_413F>
    </xsl:if>
  </xsl:template>
  <xsl:template match="w:imprint" mode="rPrChildren">
    <xsl:if test="not(@w:val) or (@w:val='on')or(@w:val='1')or(@w:val='true')">
      <字:浮雕_413F>
        <xsl:value-of select="'engrave'"/>
      </字:浮雕_413F>
    </xsl:if>
  </xsl:template>
  <xsl:template match="w:snapToGrid" mode="rPrChildren">
    <字:是否字符对齐网格_4147>
        <xsl:choose>
          <xsl:when test="not(@w:val) or (@w:val='on')or(@w:val='1')or(@w:val='true')">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'false'"/>
          </xsl:otherwise>
        </xsl:choose>
    </字:是否字符对齐网格_4147>
  </xsl:template>
  <xsl:template match="w:vanish" mode="rPrChildren">
    <字:是否隐藏文字_413D>
        <xsl:choose>
          <xsl:when test="not(@w:val) or (@w:val='on')or(@w:val='1')or(@w:val='true')">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'false'"/>
          </xsl:otherwise>
        </xsl:choose>
    </字:是否隐藏文字_413D>
  </xsl:template>
  <xsl:template match="w:spacing" mode="rPrChildren">
    <字:字符间距_4145>
      <xsl:value-of select="format-number(number(@w:val) div 20,'0.0')"/>
    </字:字符间距_4145>
  </xsl:template>
  <xsl:template match="w:w" mode="rPrChildren">
    <字:缩放_4144>
      <xsl:if test="@w:val">
        <xsl:value-of select="@w:val"/>
      </xsl:if>
      <xsl:if test="not(@w:val)">
        <xsl:value-of select="100"/>
      </xsl:if>
    </字:缩放_4144>
  </xsl:template>
  <xsl:template match="w:kern" mode="rPrChildren">
    <字:调整字间距_4146>
      <xsl:value-of select="format-number(number(@w:val) div 2,'0.0')"/>
    </字:调整字间距_4146>
  </xsl:template>
  <xsl:template match="w:position" mode="rPrChildren">
    <字:位置_4142>
      <字:偏移量_4126>
        <xsl:value-of select="format-number(number(@w:val) div 2,'0.0')"/>
      </字:偏移量_4126>
      <字:缩放量_4127>
        <xsl:value-of select="'1'"/>
      </字:缩放量_4127>       
    </字:位置_4142>
  </xsl:template>
  <xsl:template match="w:highlight" mode="rPrChildren">
    <xsl:if test="@w:val!= 'none'">
      <字:突出显示颜色_4132><!--uof:attrList="颜色"--> 
        <xsl:if test="@w:val">
          <!--<xsl:attribute name="字:颜色">-->
            <xsl:if test="@w:val='black'">
              <xsl:value-of select="'#000000'"/>
            </xsl:if>
            <xsl:if test="@w:val='blue'">
              <xsl:value-of select="'#0000ff'"/>
            </xsl:if>
            <xsl:if test="@w:val='cyan'">
              <xsl:value-of select="'#00FFFF'"/>
            </xsl:if>
            <xsl:if test="@w:val='red'">
              <xsl:value-of select="'#ff0000'"/>
            </xsl:if>
            <xsl:if test="@w:val='yellow'">
              <xsl:value-of select="'#ffff00'"/>
            </xsl:if>
            <xsl:if test="@w:val='green'">
              <xsl:value-of select="'#00ff00'"/>
            </xsl:if>
            <xsl:if test="@w:val='white'">
              <xsl:value-of select="'#FFFFFF'"/>
            </xsl:if>
            <xsl:if test="@w:val='magenta'">
              <xsl:value-of select="'#ff00ff'"/>
            </xsl:if>
            <xsl:if test="@w:val='darkBlue'">
              <xsl:value-of select="'#000080'"/>
            </xsl:if>
            <xsl:if test="@w:val='darkCyan'">
              <xsl:value-of select="'#008080'"/>
            </xsl:if>
            <xsl:if test="@w:val='darkGreen'">
              <xsl:value-of select="'#008000'"/>
            </xsl:if>
            <xsl:if test="@w:val='darkMagenta'">
              <xsl:value-of select="'#800080'"/>
            </xsl:if>
            <xsl:if test="@w:val='darkRed'">
              <xsl:value-of select="'#800000'"/>
            </xsl:if>
            <xsl:if test="@w:val='darkYellow'">
              <xsl:value-of select="'#808000'"/>
            </xsl:if>
            <xsl:if test="@w:val='darkGray'">
              <xsl:value-of select="'#808080'"/>
            </xsl:if>
            <xsl:if test="@w:val='lightGray'">
              <xsl:value-of select="'#c0c0c0'"/>
            </xsl:if>
          <!--</xsl:attribute>-->
        </xsl:if>
      </字:突出显示颜色_4132>
    </xsl:if>
  </xsl:template>
  <!--线型 虚实 颜色 字下划线,09.12.11-->
  <xsl:template match="w:u" mode="rPrChildren">
    <字:下划线_4136><!--attrList="线型 虚实 颜色 字下划线"--> 
      <xsl:if test="@w:val">
        <xsl:choose>
          <xsl:when test="@w:val='single'">
            <xsl:attribute name="线型_4137">single</xsl:attribute>
            <xsl:attribute name="虚实_4138">solid</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='double'">
            <xsl:attribute name="线型_4137">double</xsl:attribute>
            <xsl:attribute name="虚实_4138">solid</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='thick'">
            <xsl:attribute name="线型_4137">single</xsl:attribute>
            <xsl:attribute name="虚实_4138">round-dot</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='dotted'">
            <xsl:attribute name="线型_4137">single</xsl:attribute>
            <xsl:attribute name="虚实_4138">square-dot</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='dottedHeavy'">
            <xsl:attribute name="线型_4137">thick-between-thin</xsl:attribute>
            <xsl:attribute name="虚实_4138">square-dot</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='dash'">
            <xsl:attribute name="线型_4137">single</xsl:attribute>
            <xsl:attribute name="虚实_4138">dash</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='dashedHeavy'">
            <xsl:attribute name="线型_4137">thick-between-thin</xsl:attribute>
            <xsl:attribute name="虚实_4138">dash</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='dashLong'">
            <xsl:attribute name="线型_4137">single</xsl:attribute>
            <xsl:attribute name="虚实_4138">long-dash</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='dashLongHeavy'">
            <xsl:attribute name="线型_4137">single</xsl:attribute>
            <xsl:attribute name="虚实_4138">long-dash</xsl:attribute>
          </xsl:when>
          
          <xsl:when test="@w:val='dotDash'">
            <xsl:attribute name="线型_4137">single</xsl:attribute>
            <xsl:attribute name="虚实_4138">dash-dot</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='dashDotHeavy'">
            <xsl:attribute name="线型_4137">thick-between-thin</xsl:attribute>
            <xsl:attribute name="虚实_4138">dash-dot</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='dotDotDash'">
            <xsl:attribute name="线型_4137">single</xsl:attribute>
            <xsl:attribute name="虚实_4138">dash-dot-dot</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='dashDotDotHeavy'">
            <xsl:attribute name="线型_4137">thick-between-thin</xsl:attribute>
            <xsl:attribute name="虚实_4138">dash-dot-dot</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='wave'">
            <xsl:attribute name="线型_4137">single</xsl:attribute>
            <xsl:attribute name="虚实_4138">wave</xsl:attribute>
          </xsl:when>     
          <xsl:when test="@w:val='wavyHeavy'">
            <xsl:attribute name="线型_4137">thick-between-thin</xsl:attribute>
            <xsl:attribute name="虚实_4138">long-dash-dot</xsl:attribute>
          </xsl:when>
          <xsl:when test="@w:val='wavyDouble'">
            <xsl:attribute name="线型_4137">double</xsl:attribute>
            <xsl:attribute name="虚实_4138">round-dot</xsl:attribute>
          </xsl:when>
          
          <xsl:otherwise>
            <xsl:attribute name="线型_4137">single</xsl:attribute>
            <xsl:attribute name="虚实_4138">solid</xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
      <xsl:if test="not(@w:val)">
        <xsl:attribute name="线型_4137">none</xsl:attribute>
      </xsl:if>
      <xsl:if test ="@w:val='words'">
        <xsl:attribute name ="是否字下划线_4139">true</xsl:attribute>
      </xsl:if>
      <xsl:if test="@w:color">
        <xsl:attribute name="颜色_412F">
          <xsl:if test="not(@w:color='auto')">
            <xsl:value-of select="'#'"/>
          </xsl:if>
          <xsl:value-of select="@w:color"/>
        </xsl:attribute>
      </xsl:if>
    </字:下划线_4136>
  </xsl:template>
  <xsl:template match="w:bdr" mode="rPrChildren">
    <字:边框线_4226><!--文字边框-->
      <xsl:if test ="@w:val">
        <xsl:call-template name="borderStyle"><!--页面边框模板-->
          <xsl:with-param name="styleVal" select="@w:val"/>
        </xsl:call-template>
      </xsl:if>
      <xsl:if test="@w:sz">
        <xsl:attribute name="宽度_C60F">
          <xsl:value-of select="format-number(number(@w:sz) div 8,'0.0')"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@w:color">
        <xsl:attribute name="颜色_C611">
          <xsl:if test="not(@w:color='auto')">
            <xsl:value-of select="'#'"/>
          </xsl:if>
          <xsl:value-of select="@w:color"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@w:space">
        <xsl:attribute name="边距_C610">
          <xsl:value-of select="@w:space"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@w:shadow">
        <xsl:attribute name="阴影类型_C645">
          <xsl:value-of select="'right-bottom'"/>
        </xsl:attribute>
      </xsl:if>
    </字:边框线_4226>
  </xsl:template>
  <xsl:template match="w:shd" mode="rPrChildren">
    <xsl:if test="./@w:val!='nil'">
      <字:填充_4134>
        <xsl:call-template name="shd"/><!--调用填充的模板,标准有变化应该要修改%%%%%%%%%%%%%%%%%%%%%%%%%-->
      </字:填充_4134>
    </xsl:if>
  </xsl:template>
  <xsl:template match="w:vertAlign" mode="rPrChildren">
    <字:上下标类型_4143>
        <xsl:if test="@w:val='subscript'">
          <xsl:value-of select="'sub'"/>
        </xsl:if>
        <xsl:if test="@w:val='superscript'">
          <xsl:value-of select="'sup'"/>
        </xsl:if>
        <xsl:if test="@w:val='baseline'">
          <xsl:value-of select="'none'"/>
        </xsl:if>
    </字:上下标类型_4143>
  </xsl:template>
  <xsl:template match="w:em" mode="rPrChildren">
    <字:着重号_413A>
      <xsl:choose>
        <xsl:when test="@w:val='none'">
          <xsl:attribute name="类型_413B">none</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="类型_413B">dot</xsl:attribute>
          <xsl:attribute name="是否字着重号_413C">true</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </字:着重号_413A>
  </xsl:template>

  <!--2013-04-19，wudi，增加对chart的转换，start-->
  <xsl:template name ="bodyAnchorChart">
    <xsl:param name ="picType"/>
    <xsl:param name ="filename"/>
    <xsl:variable name="number">
      <xsl:number format="1" level="any" count="v:rect|wp:anchor|wp:inline|v:shape"/>
    </xsl:variable>
    <uof:锚点_C644>
      <xsl:attribute name="图形引用_C62E">
        <xsl:choose>
          <xsl:when test="$filename='document'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="$filename='comments'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="$filename='endnotes'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="$filename='footnotes'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="contains($filename,'header')">
            <xsl:variable name="hn" select="substring-before($filename,'.xml')"/>
            <xsl:value-of select="concat($hn,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="contains($filename,'footer')">
            <xsl:variable name="fn" select="substring-before($filename,'.xml')"/>
            <xsl:value-of select="concat($fn,'Obj',$number * 2 +1)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
      <uof:位置_C620>
        <uof:水平_4106>
          <xsl:attribute name="相对于_410C">
            <xsl:value-of select="'column'"/>
          </xsl:attribute>
          <uof:绝对_4107>
            <xsl:attribute name="值_4108">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
          </uof:绝对_4107>
        </uof:水平_4106>
        <uof:垂直_410D>
          <xsl:attribute name="相对于_C647">
            <xsl:value-of select="'paragraph'"/>
          </xsl:attribute>
          <uof:绝对_4107>
            <xsl:attribute name="值_4108">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
          </uof:绝对_4107>
        </uof:垂直_410D>
      </uof:位置_C620>
      <uof:大小_C621>
        <xsl:attribute name="宽_C605">
          <xsl:value-of select="wp:extent/@cx div 12700"/>
        </xsl:attribute>
        <xsl:attribute name="长_C604">
          <xsl:value-of select="wp:extent/@cy div 12700"/>
        </xsl:attribute>
      </uof:大小_C621>
      <uof:绕排_C622 环绕文字_C624="both"/>
      <xsl:if test="@distT or @distL or @distR or @distB">
        <uof:边距_C628>
          <xsl:attribute name="上_C609">
            <xsl:value-of select="@distT div 12500"/>
          </xsl:attribute>
          <xsl:attribute name="左_C608">
            <xsl:value-of select="@distL div 12500"/>
          </xsl:attribute>
          <xsl:attribute name="右_C60A">
            <xsl:value-of select="@distR div 12500"/>
          </xsl:attribute>
          <xsl:attribute name="下_C60B">
            <xsl:value-of select="@distB div 12500"/>
          </xsl:attribute>
        </uof:边距_C628>
      </xsl:if>
    </uof:锚点_C644>
    
  </xsl:template>
  <!--end-->
   
  <!--zhaobj 2011/02/23 预定义图形的锚点-->
  <!--这一部分标准变化较大，以后应该要增加代码@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
  <xsl:template name="bodyAnchorWps">
    <xsl:param name="picType"/>
    <xsl:param name="filename"/>
    <xsl:variable name="number">
      <xsl:number format="1" level="any" count="v:rect|wp:anchor|wp:inline|v:shape"/>
    </xsl:variable>
    <uof:锚点_C644> <!--uof:attrList="标识符 类型"-->  
      <!--<xsl:attribute name="字:类型">
        <xsl:if test="$picType='inline'">
          <xsl:value-of select="'inline'"/>
        </xsl:if>
        <xsl:if test="$picType='anchor'">
          <xsl:value-of select="'normal'"/>
        </xsl:if>
      </xsl:attribute>-->
      <xsl:attribute name="图形引用_C62E">
        <xsl:choose>
          <xsl:when test="$filename='document'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="$filename='comments'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="$filename='endnotes'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="$filename='footnotes'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>

          <!--2014-03-25，wudi，修改图形引用_C62E的取值方法，页眉页脚调整，start-->
          <!--2014-04-09，wudi，还原取值方法，新的取值方法会带来新的问题，影响页眉页脚转换，start-->
          <xsl:when test="contains($filename,'header')">
            <xsl:variable name="hn" select="substring-before($filename,'.xml')"/>
            <xsl:value-of select="concat($hn,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="contains($filename,'footer')">
            <xsl:variable name="fn" select="substring-before($filename,'.xml')"/>
            <xsl:value-of select="concat($fn,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <!--end-->
          <!--end-->
          
        </xsl:choose>
      </xsl:attribute>
      <uof:位置_C620>
        <!--这一部分代码需要修改，标准里位置多了类型属性@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
        <xsl:if test=".//wp:positionH or .//wp:positionV">
          <xsl:if test=".//wp:positionH">
            <uof:水平_4106>
              <xsl:attribute name="相对于_410C">
                <xsl:call-template name="positionHrelative2">
                  <xsl:with-param name="val" select=".//wp:positionH/@relativeFrom"/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:if test=".//wp:positionH/wp:posOffset">
                <uof:绝对_4107>
                  <xsl:attribute name="值_4108">
                    <xsl:value-of select=".//wp:positionH/wp:posOffset div 12700"/>
                  </xsl:attribute>
                </uof:绝对_4107>
              </xsl:if>
              <xsl:if test=".//wp:positionH/wp:align">
                <uof:相对_4109>
                  <xsl:attribute name="参考点_410A">
                    <xsl:call-template name="positionHalign">
                      <xsl:with-param name="val" select=".//wp:positionH/wp:align"/>
                      <xsl:with-param name="val2" select=".//wp:positionH/@relativeFrom"/>
                    </xsl:call-template>
                  </xsl:attribute>
                  <!--<xsl:attribute name="字:值">
                    <xsl:value-of select="'0'"/>
                  </xsl:attribute>-->
                </uof:相对_4109>
              </xsl:if>
            </uof:水平_4106>
          </xsl:if>
          <xsl:if test=".//wp:positionV">
            <uof:垂直_410D>

              <!--2015-04-26，wudi，修复文本框位置转换BUG，实际上是针对非标准OOX代码做特殊处理，start-->
              <xsl:choose>
                <xsl:when test=".//wp:positionV/wp14:pctPosVOffset">
                  <xsl:attribute name="相对于_C647">
                    <xsl:call-template name="positionVrelative2">
                      <xsl:with-param name="val" select=".//mc:Fallback/wp:positionV/@relativeFrom"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="相对于_C647">
                    <xsl:call-template name="positionVrelative2">
                      <xsl:with-param name="val" select=".//wp:positionV/@relativeFrom"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
              <!--end-->

              <xsl:if test=".//wp:positionV/wp:posOffset">
                <uof:绝对_4107>
                  <xsl:attribute name="值_4108">
                    <xsl:value-of select=".//wp:positionV/wp:posOffset div 12700"/>
                  </xsl:attribute>
                </uof:绝对_4107>
              </xsl:if>
              <xsl:if test=".//wp:positionV/wp:align">
                <uof:相对_4109>
                  <xsl:attribute name="参考点_410B">
                    <xsl:call-template name="positionValign2">
                      <xsl:with-param name="val" select=".//wp:positionV/wp:align"/>
                      <xsl:with-param name="val2" select=".//wp:positionV/@relativeFrom"/>
                    </xsl:call-template>
                  </xsl:attribute>
                  <!--<xsl:attribute name="字:值">
                    <xsl:value-of select="'0'"/>
                  </xsl:attribute>-->
                </uof:相对_4109>
              </xsl:if>
            </uof:垂直_410D>
          </xsl:if>
        </xsl:if>
      </uof:位置_C620>
      <uof:大小_C621>
        <xsl:attribute name="宽_C605">
          <xsl:value-of select="wp:extent/@cx div 12700"/>
        </xsl:attribute>
        <xsl:attribute name="长_C604">
          <xsl:value-of select="wp:extent/@cy div 12700"/>
        </xsl:attribute>
      </uof:大小_C621>
      <uof:绕排_C622><!--attrList="绕排方式 环绕文字 绕排顶点"--> 
        <xsl:if test="@behindDoc='1'">
          <xsl:attribute name="绕排方式_C623">
            <xsl:value-of select="'behind-text'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="@behindDoc='0'">
          <xsl:attribute name="绕排方式_C623">
            <xsl:value-of select="'infront-of-text'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="wp:wrapTight">
          <xsl:attribute name="绕排方式_C623">
            <xsl:value-of select="'tight'"/>
          </xsl:attribute>
          <xsl:if test="wp:wrapTight/@wrapText">
            <xsl:attribute name="环绕文字_C624">
              <xsl:call-template name="wraptext2">
                <xsl:with-param name="val" select="wp:wrapTight/@wrapText"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
        <xsl:if test="wp:wrapTopAndBottom">
          <xsl:attribute name="绕排方式_C623">
            <xsl:value-of select="'top-bottom'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="wp:wrapThrough">
          <xsl:attribute name="绕排方式_C623">
            <xsl:value-of select="'through'"/>
          </xsl:attribute>
          <xsl:if test="wp:wrapThrough/@wrapText">
            <xsl:attribute name="环绕文字_C624">
              <xsl:call-template name="wraptext2">
                <xsl:with-param name="val" select="wp:wrapThrough/@wrapText"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
        <xsl:if test="wp:wrapSquare">
          <xsl:attribute name="绕排方式_C623">
            <xsl:value-of select="'square'"/>
          </xsl:attribute>
          <xsl:if test="wp:wrapSquare/@wrapText">
            <xsl:attribute name="环绕文字_C624">
              <xsl:call-template name="wraptext2">
                <xsl:with-param name="val" select="wp:wrapSquare/@wrapText"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>          
      </uof:绕排_C622>
      <xsl:if test="@distT or @distL or @distR or @distB">
        <uof:边距_C628>
          <xsl:attribute name="上_C609">
            <xsl:value-of select="@distT div 12500"/>
          </xsl:attribute>
          <xsl:attribute name="左_C608">
            <xsl:value-of select="@distL div 12500"/>
          </xsl:attribute>
          <xsl:attribute name="右_C60A">
            <xsl:value-of select="@distR div 12500"/>
          </xsl:attribute>
          <xsl:attribute name="下_C60B">
            <xsl:value-of select="@distB div 12500"/>
          </xsl:attribute>
        </uof:边距_C628>
      </xsl:if>
      <!--zhaobj
        <xsl:if test="wp:cNvGraphicFramePr/a:graphicFrameLocks/@noResize='1'">
        <字:锁定 uof:locID="t0117" uof:attrList="值" 字:值="true"/>
      </xsl:if> -->
      <xsl:if test="@locked='1'">
        <uof:是否锁定_C629>
          <xsl:value-of select="'true'"/>
        </uof:是否锁定_C629>
      </xsl:if>
      <xsl:if test="@locked='0'">
        <uof:是否锁定_C629>
          <xsl:value-of select="'false'"/>
        </uof:是否锁定_C629>
      </xsl:if>
      <uof:保护_C62A><!--干嘛的？？？？？？？？？？？？？？？？？？？？？？？？？？-->
        <xsl:attribute name="是否保护内容_C641">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
        <xsl:attribute name="是否保护位置_C642">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
        <xsl:attribute name="是否保护大小_C643">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
      </uof:保护_C62A>
      <xsl:if test="@allowOverlap='1'">
        <uof:是否允许重叠_C62B>
          <xsl:value-of select="'true'"/>
        </uof:是否允许重叠_C62B>
      </xsl:if>
    </uof:锚点_C644>
  </xsl:template>
  <xsl:template name="wraptext2"><!--环绕文字-->
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$val='bothSides'">
        <xsl:value-of select="'both'"/>
      </xsl:when>
      <xsl:when test="$val='left'">
        <xsl:value-of select="'left'"/>
      </xsl:when>
      <xsl:when test="$val='right'">
        <xsl:value-of select="'right'"/>
      </xsl:when>
      <xsl:when test="$val='largest'">
        <xsl:value-of select="'largest'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'both'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="positionHrelative2">
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$val='margin'">
        <xsl:value-of select="'margin'"/>
      </xsl:when>
      <xsl:when test="$val='page'">
        <xsl:value-of select="'page'"/>
      </xsl:when>
      <xsl:when test="$val='column'">
        <xsl:value-of select="'column'"/>
      </xsl:when>
      <xsl:when test="$val='character'">
        <xsl:value-of select="'char'"/>
      </xsl:when>
      <xsl:when test="$val='rightMargin' or $val='leftMargin'">
        <xsl:value-of select="'page'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'margin'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="positionHalign">
    <xsl:param name="val"/>
    <xsl:param name="val2"/>
    <xsl:if test="$val2!='rightMargin' and $val2!='leftMargin'">
      <xsl:choose>
        <xsl:when test="$val='left'">
          <xsl:value-of select="'left'"/>
        </xsl:when>
        <xsl:when test="$val='right'">
          <xsl:value-of select="'right'"/>
        </xsl:when>
        <xsl:when test="$val='center'">
          <xsl:value-of select="'center'"/>
        </xsl:when>
        
        <!--2013-05-03，wudi，修复OOX到UOF方向图片位置转换BUG，增加书籍版式的转换，start-->
        <xsl:when test ="$val='inside'">
          <xsl:value-of select ="'inside'"/>
        </xsl:when>
        <xsl:when test ="$val='outside'">
          <xsl:value-of select ="'outside'"/>
        </xsl:when>
        <!--end-->
        
        <xsl:otherwise>
          <xsl:value-of select="'left'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="$val2='rightMargin'">
      <xsl:value-of select="'right'"/>
    </xsl:if>
    <xsl:if test="$val2='leftMargin'">
      <xsl:value-of select="'left'"/>
    </xsl:if>
  </xsl:template>
  <xsl:template name="positionVrelative2">
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$val='margin'">
        <xsl:value-of select="'margin'"/>
      </xsl:when>
      <xsl:when test="$val='page'">
        <xsl:value-of select="'page'"/>
      </xsl:when>
      <xsl:when test="$val='paragraph'">
        <xsl:value-of select="'paragraph'"/>
      </xsl:when>
      <xsl:when test="$val='line'">
        <xsl:value-of select="'line'"/>
      </xsl:when>
      <xsl:when test="$val='topMargin' or $val='bottomMargin'">
        <xsl:value-of select="'page'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'margin'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="positionValign2">
    <xsl:param name="val"/>
    <xsl:param name="val2"/>
    <xsl:if test="$val2!='topMargin' and $val2!='bottomMargin'">
      <xsl:choose>
        <xsl:when test="$val='top'">
          <xsl:value-of select="'top'"/>
        </xsl:when>
        <xsl:when test="$val='bottom'">
          <xsl:value-of select="'bottom'"/>
        </xsl:when>
        <xsl:when test="$val='center'">
          <xsl:value-of select="'center'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'top'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="$val2='topMargin'">
      <xsl:value-of select="'top'"/>
    </xsl:if>
    <xsl:if test="$val2='bottomMargin'">
      <xsl:value-of select="'bottom'"/>
    </xsl:if>
  </xsl:template>
  
</xsl:stylesheet>