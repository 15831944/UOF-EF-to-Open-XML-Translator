<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
<Author>Li Yanyan(BITI)</Author>
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles">

  <xsl:import href="paragraph.xsl"/>
  <xsl:import href="common.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

  <xsl:template name="table">
  <!--cxl,current node:字:文字表_416C,2012/2/9-->
    <xsl:message terminate="no">progress:Table</xsl:message>
    <w:tbl>
      <xsl:for-each select="字:文字表属性_41CC | 字:行_41CD">
        <xsl:choose>
          <xsl:when test="name(.)='字:文字表属性_41CC'">
            
            <!--2014-05-20，wudi，修复文字表式样引用反方向转换丢失的BUG，取值有误，start-->
            <w:tblPr>
              <xsl:if test="./@式样引用_419C">
                <w:tblStyle>
                  <xsl:attribute name="w:val">
                    <xsl:value-of select="./@式样引用_419C"/>
                  </xsl:attribute>
                </w:tblStyle>
              </xsl:if>
              <xsl:call-template name="tblPr"/>
            </w:tblPr>
            <!--end-->

            <xsl:choose>
              <xsl:when test="字:列宽集_41C1">
                <xsl:call-template name="tblGrid"/>
              </xsl:when>
              <xsl:otherwise>
                <w:tblGrid/>
              </xsl:otherwise>
            </xsl:choose>

          </xsl:when>
          <xsl:when test="name(.)='字:行_41CD'">
            
            <!--2014-05-08，wudi，增加postr参数，记录当前单元格所在行数，start-->
            <xsl:variable name="postr">
              <xsl:number count="字:行_41CD" format="1" level="single"/>
            </xsl:variable>
            
            <xsl:call-template name="tblRow">
              <xsl:with-param name="postr" select="$postr"/>
            </xsl:call-template>
            <!--end-->
            
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </w:tbl>
  </xsl:template>

  <xsl:template name="tblPr"><!--current node:字:文字表属性_41CC-->
    <xsl:if test="字:绕排_41C5='around'">
      <xsl:call-template name="tblpPr"/>
    </xsl:if>

    <xsl:choose>
      <xsl:when test="字:宽度_41A1">
        <w:tblW>
          <xsl:choose>
            <xsl:when test="字:宽度_41A1/@相对宽度_41C0 and not(字:宽度_41A1/@绝对宽度_41BF)">
              <xsl:attribute name="w:w">
                <xsl:call-template name="twipsMeasure-pct">
                  <xsl:with-param name="lengthVal" select="字:宽度_41A1/@相对宽度_41C0"/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name="w:type">
                <xsl:value-of select="'pct'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="字:宽度_41A1/@绝对宽度_41BF">
              <xsl:attribute name="w:w">
                <xsl:call-template name="twipsMeasure">
                  <xsl:with-param name="lengthVal" select="字:宽度_41A1/@绝对宽度_41BF"/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name="w:type">
                <xsl:value-of select="'dxa'"/>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </w:tblW>
      </xsl:when>
      <xsl:otherwise>
        <w:tblW w:w="0" w:type="auto"/>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:if test="字:对齐_41C3">
      <w:jc>
        <xsl:attribute name="w:val">
          
          <!--2015-01-23，wudi，左边框start，右边框end，start-->
          <xsl:choose>
            <xsl:when test="字:对齐_41C3 = 'left'">
              <xsl:value-of select="'start'"/>
            </xsl:when>
            <xsl:when test="字:对齐_41C3 = 'right'">
              <xsl:value-of select="'end'"/>
            </xsl:when>
            <xsl:when test="字:对齐_41C3 = 'center'">
              <xsl:value-of select="'center'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'start'"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--end-->

        </xsl:attribute>
      </w:jc>
    </xsl:if>
    
    <xsl:if test="字:默认单元格间距_41CB">
      <w:tblCellSpacing>
        <xsl:attribute name="w:w">
          <xsl:call-template name="tableTwipsMeasure"><!--位于common.xsl中，进行不同单位的换算-->
            <xsl:with-param name="lengthVal" select="./字:默认单元格间距_41CB"/>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="w:type">
          <xsl:value-of select="'dxa'"/>
        </xsl:attribute>
      </w:tblCellSpacing>
    </xsl:if>

    <xsl:if test="字:左缩进_41C4">
      <w:tblInd>
        <xsl:attribute name="w:w">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="字:左缩进_41C4"/>
          </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="w:type">
          <xsl:value-of select="'dxa'"/>
        </xsl:attribute>
      </w:tblInd>
    </xsl:if>

    <xsl:if test="字:文字表边框_4227">
      <w:tblBorders>
        <xsl:call-template name="TableBorder">
          <xsl:with-param name="Bor" select="字:文字表边框_4227"/>
        </xsl:call-template>
      </w:tblBorders>
    </xsl:if>
    <xsl:if test="字:填充_4134">
      <!--xsl:call-template name="shd"/-->
      <xsl:apply-templates select="字:填充_4134" mode="shdCommon"/>
    </xsl:if>

    <xsl:choose>
      <xsl:when test="字:是否自动调整大小_41C8">
        <w:tblLayout>
          <xsl:attribute name="w:type">
            <xsl:choose>
              <xsl:when test="字:是否自动调整大小_41C8='true' or 字:是否自动调整大小_41C8='1'">
                <xsl:value-of select="'autofit'"/>
              </xsl:when>
              <xsl:when test="字:是否自动调整大小_41C8='false' or 字:是否自动调整大小_41C8='0'">
                <xsl:value-of select="'fixed'"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </w:tblLayout>
      </xsl:when>
      <xsl:otherwise>
        <w:tblLayout w:type="fixed"/>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:if test="字:默认单元格边距_41CA">
      <w:tblCellMar>
        <xsl:call-template name="Cellmar">
          <xsl:with-param name="Cellmar" select="字:默认单元格边距_41CA"/>
        </xsl:call-template>
      </w:tblCellMar>
    </xsl:if>

    <!--w:tblLook w:val="04A0"/-->
  </xsl:template>

  <xsl:template name="tblGrid">
    <w:tblGrid>
      <xsl:for-each select="字:列宽集_41C1/字:列宽_41C2">
        <w:gridCol>
          <xsl:attribute name="w:w">
            <!--xsl:value-of select="round(. * 20)"/-->
            <xsl:call-template name="twipsMeasure">
              <xsl:with-param name="lengthVal" select="."/>
            </xsl:call-template>
          </xsl:attribute>
        </w:gridCol>
      </xsl:for-each>
    </w:tblGrid>
  </xsl:template>

  <xsl:template name="tblRow">

    <!--2014-05-08，wudi，增加postr参数，记录当前单元格所在行数，start-->
    <xsl:param name="postr"/>
    <!--end-->
    
    <w:tr>
      <xsl:call-template name="tblRowPr"/><!--表行属性模板-->
      <xsl:for-each select="字:单元格_41BE">
        <!--<xsl:choose>
          
          --><!--2014-04-30，wudi，修复涉及单元格合并的表格边框转换缺失的BUG，start--><!--
          <xsl:when test="name(.)='字:单元格_41BE'">-->

            <xsl:variable name ="postc">
              <xsl:number count="字:单元格_41BE" format="1" level="single"/>
            </xsl:variable>
            
            <xsl:if test="./字:段落_416B or ./字:新建的单元格">
              <!--<test>
                <xsl:value-of select="$postc"/>
              </test>-->
              <xsl:call-template name="tblCell">
                <xsl:with-param name="postr" select="$postr"/>
                <xsl:with-param name="postc" select="$postc"/>
              </xsl:call-template>
            </xsl:if>
          <!--</xsl:when>
          --><!--end--><!--
          
        </xsl:choose>-->
      </xsl:for-each>
    </w:tr>
  </xsl:template>

  <xsl:template name="tblRowPr">
    <w:trPr>
      <xsl:for-each select="./字:单元格_41BE">
        <xsl:variable name="postc">
          <xsl:number count="字:单元格_41BE" format="1" level="single"/>
        </xsl:variable>
        <!--cxl,2012.5.15删除这两个元素的转换，因为两边实现并不一致，uof中没有对应元素-->
        <!--<xsl:if test="not(./字:段落_416B) and ($postc='1')">
          <w:gridBefore>
            <xsl:attribute name="w:val">
              <xsl:if test="./字:单元格属性_41B7/字:跨列_41A7">
                <xsl:value-of select="./字:单元格属性_41B7/字:跨列_41A7"/>
              </xsl:if>
              <xsl:if test="not(./字:单元格属性_41B7/字:跨列_41A7)">
                <xsl:value-of select="'1'"/>
              </xsl:if>
            </xsl:attribute>
          </w:gridBefore>
        </xsl:if>

        <xsl:if test="not(./字:段落_416B) and ($postc!='1')">
          <w:gridAfter>
            <xsl:attribute name="w:val">
              <xsl:if test="./字:单元格属性_41B7/字:跨列_41A7">
                <xsl:value-of select="./字:单元格属性_41B7/字:跨列_41A7"/>
              </xsl:if>
              <xsl:if test="not(./字:单元格属性_41B7/字:跨列_41A7)">
                <xsl:value-of select="'1'"/>
              </xsl:if>
            </xsl:attribute>
          </w:gridAfter>
        </xsl:if>-->
      </xsl:for-each>

      <xsl:choose>
        <xsl:when test="字:表行属性_41BD/字:是否跨页_41BB">
          <w:cantSplit>
            <xsl:attribute name="w:val">
              <xsl:if test="字:表行属性_41BD/字:是否跨页_41BB='true' or 字:表行属性_41BD/字:是否跨页_41BB='1'">
                <xsl:value-of select="'off'"/>
              </xsl:if>
              <xsl:if test="字:表行属性_41BD/字:是否跨页_41BB='false' or 字:表行属性_41BD/字:是否跨页_41BB='0'">
                <xsl:value-of select="'on'"/>
              </xsl:if>
            </xsl:attribute>
          </w:cantSplit>
        </xsl:when>
        <xsl:otherwise>
          <w:cantSplit w:val="on"/>
        </xsl:otherwise>
      </xsl:choose>
      
      <xsl:choose>
        <xsl:when test="字:表行属性_41BD/字:高度_41B8">
          <w:trHeight>
            <xsl:if test="字:表行属性_41BD/字:高度_41B8/@固定值_41B9 and not(字:表行属性_41BD/字:高度_41B8/@最小值_41BA)">
              <xsl:attribute name="w:val">
                <xsl:call-template name="twipsMeasure">
                  <xsl:with-param name="lengthVal" select="字:表行属性_41BD/字:高度_41B8/@固定值_41B9"/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name="w:hRule">
                <xsl:value-of select="'exact'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="字:表行属性_41BD/字:高度_41B8/@最小值_41BA">
              <xsl:attribute name="w:val">
                <xsl:call-template name="twipsMeasure">
                  <xsl:with-param name="lengthVal" select="字:表行属性_41BD/字:高度_41B8/@最小值_41BA"/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name="w:hRule">
                <xsl:value-of select="'atLeast'"/>
              </xsl:attribute>
            </xsl:if>
          </w:trHeight>
        </xsl:when>
        <!--
        <xsl:otherwise>
          <w:trHeight w:val="310" w:hRule="atLeast"/>
          <w:trHeight w:val="0" w:hRule="auto"/>
        </xsl:otherwise>
        -->
      </xsl:choose>
      <xsl:if test="字:表行属性_41BD/字:是否表头行_41BC">
        <w:tblHeader>
          <xsl:attribute name="w:val">
            <xsl:if test="字:表行属性_41BD/字:是否表头行_41BC='true' or 字:表行属性_41BD/字:是否表头行_41BC='1'">
              <xsl:value-of select="'on'"/>
            </xsl:if>
            <xsl:if test="字:表行属性_41BD/字:是否表头行_41BC='false' or 字:表行属性_41BD/字:是否表头行_41BC='0'">
              <xsl:value-of select="'off'"/>
            </xsl:if>
          </xsl:attribute>
        </w:tblHeader>
      </xsl:if>
    </w:trPr>
  </xsl:template>

  <xsl:template name="tblCell">

    <!--2014-05-08，wudi，增加postr参数，记录当前单元格所在行数，start-->
    <xsl:param name="postr"/>
    <!--end-->
    
    <!--2014-04-30，wudi，修复涉及单元格合并的表格边框转换缺失的BUG，记录单元格列数，start-->
    <xsl:param name="postc"/>
    <!--end-->
    
    <w:tc>
      <xsl:for-each select="字:单元格属性_41B7 | 字:段落_416B | 字:文字表_416C | 字:新建的单元格">
        <xsl:choose>
          <xsl:when test="name(.)='字:单元格属性_41B7'">
            <xsl:call-template name="tblCellPr">
              <xsl:with-param name="postr" select="$postr"/>
            </xsl:call-template><!--单元格属性模板-->
          </xsl:when>
          <xsl:when test="name(.)='字:段落_416B'">
            <xsl:call-template name="paragraph"/>
          </xsl:when>
          <xsl:when test="name(.)='字:文字表_416C'">
            <xsl:call-template name="table"/>
          </xsl:when>
          <xsl:when test="name(.)='字:新建的单元格'">
            <w:tcPr>
              <w:tcW w:w="1705" w:type="dxa"/>
              <w:vMerge/>
              
              <!--2014-04-30，wudi，修复涉及单元格垂直合并的表格边框转换丢失的BUG，start-->
              <!--2014-04-30，wudi，因为垂直合并的转换本身就有差异，以下处理方式只是一种折中的办法，与源案例的代码有出入，start-->
              <!--<w:tcBorders>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
              </w:tcBorders>
              <w:vAlign w:val="center"/>
              <w:hideMark/>-->
              <!--end-->
              <!--end-->

              <!--2014-04-30，wudi，修复涉及单元格合并的表格边框转换缺失的BUG，新增模板，处理垂直合并单元格边框的转换，start-->
              <xsl:call-template name="vMerge">
                <xsl:with-param name="postc" select="$postc"/>
                <xsl:with-param name="count" select="1"/>
              </xsl:call-template>
              <!--end-->
              
            </w:tcPr>
            <w:p w:rsidR="00DD6586" w:rsidRDefault="00DD6586"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </w:tc>
  </xsl:template>

  <!--2014-04-30，wudi，修复涉及单元格合并的表格边框转换缺失的BUG，新增模板，处理垂直合并单元格边框的转换，start-->
  <xsl:template name="vMerge">
    <xsl:param name="postc"/>
    <xsl:param name="count"/>

    <!--<postc>
      <xsl:value-of select="$postc"/>
    </postc>
    <count>
      <xsl:value-of select="$count"/>
    </count>-->

    <!--2014-05-08，wudi，增加限制条件，避免死循环，部分案例难以处理，会出现死循环，start-->
    <xsl:if test="$count &lt;16">
      <xsl:if test="ancestor::字:行_41CD/preceding-sibling::字:行_41CD[number($count)]/child::字:单元格_41BE[number($postc)]/字:单元格属性_41B7/字:跨行合并开始">
        <w:tcBorders>
          <xsl:call-template name="TcBorder">
            <xsl:with-param name="Bor" select="ancestor::字:行_41CD/preceding-sibling::字:行_41CD[number($count)]/child::字:单元格_41BE[number($postc)]/字:单元格属性_41B7/字:边框_4133"/>
          </xsl:call-template>
        </w:tcBorders>
      </xsl:if>
      <xsl:if test="not(ancestor::字:行_41CD/preceding-sibling::字:行_41CD[number($count)]/child::字:单元格_41BE[number($postc)]/字:单元格属性_41B7/字:跨行合并开始)">
        <xsl:variable name="tmp">
          <xsl:value-of select="$count + 1"/>
        </xsl:variable>
        <xsl:call-template name="vMerge">
          <xsl:with-param name="postc" select="$postc"/>
          <xsl:with-param name="count" select="$tmp"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:if>
    <!--end-->
    
  </xsl:template>
  <!--end-->

  <xsl:template name="tblCellPr"><!--单元格属性模板-->

    <!--2014-05-08，wudi，增加postr参数，记录当前单元格所在行数，start-->
    <xsl:param name="postr"/>
    <!--end-->
    
    <xsl:variable name="tblId" select="ancestor::字:文字表_416C/字:文字表属性_41CC/@式样引用_419C"/>
    <w:tcPr>
      <xsl:choose>
        <xsl:when test="字:宽度_41A1">
          <w:tcW>
            <xsl:if test="字:宽度_41A1/@相对值_41A3 and not(字:宽度_41A1/@绝对值_41A2)">
              <xsl:attribute name="w:w">
                <xsl:call-template name="twipsMeasure-pct">
                  <xsl:with-param name="lengthVal" select="字:宽度_41A1/@相对值_41A3"/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name="w:type">
                <xsl:value-of select="'pct'"/>
              </xsl:attribute>
            </xsl:if>

            <xsl:if test="字:宽度_41A1/@绝对值_41A2">
              <xsl:attribute name="w:w">
                <xsl:call-template name="twipsMeasure">
                  <xsl:with-param name="lengthVal" select="字:宽度_41A1/@绝对值_41A2"/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:attribute name="w:type">
                <xsl:value-of select="'dxa'"/>
              </xsl:attribute>
            </xsl:if>
          </w:tcW>
        </xsl:when>
        <xsl:otherwise>
          <w:tcW w:w="0" w:type="auto"/>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:if test="字:跨列_41A7">
        <w:gridSpan>
          <xsl:attribute name="w:val">
            <xsl:value-of select="字:跨列_41A7"/>
          </xsl:attribute>
        </w:gridSpan>
      </xsl:if>

      <xsl:if test="字:跨行合并开始">
        <w:vMerge>
          <xsl:attribute name="w:val">
            <xsl:value-of select="'restart'"/>
          </xsl:attribute>
        </w:vMerge>
      </xsl:if>


      <xsl:if test="字:边框_4133 or 字:斜线表头_41AB">
        <w:tcBorders>
          <xsl:call-template name="TcBorder"><!--斜线表头，在这里处理-->
            <xsl:with-param name="Bor" select="字:边框_4133"/>
            <xsl:with-param name="postr" select="$postr"/>
          </xsl:call-template>
        </w:tcBorders>
      </xsl:if>

      
      <xsl:choose>
        <xsl:when test="字:填充_4134">
          <!--xsl:call-template name="shd"/-->
          <xsl:apply-templates select="字:填充_4134" mode="shdCommon"/>
        </xsl:when>
        
        <xsl:otherwise>
          <xsl:if test="ancestor::字:文字表_416C/字:文字表属性_41CC/字:填充_4134">
            <xsl:apply-templates select="ancestor::字:文字表_416C/字:文字表属性_41CC/字:填充_4134"
              mode="shdCommon"/>
          </xsl:if>
          <xsl:if test="ancestor::uof:UOF/uof:式样集//式样:文字表式样_9918[@标识符_4100=$tblId]/字:填充_4134">
            <xsl:apply-templates
              select="ancestor::uof:UOF/uof:式样集//式样:文字表式样_9918[@标识符_4100=$tblId]/字:填充_4134"
              mode="shdCommon"/>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:if test="字:是否自动换行_41A8">
        <w:noWrap>
          <xsl:attribute name="w:val">
            <xsl:if test="字:是否自动换行_41A8='true' or 字:是否自动换行_41A8='1'">
              <xsl:value-of select="'off'"/>
            </xsl:if>
            <xsl:if test="字:是否自动换行_41A8='false' or 字:是否自动换行_41A8='0'">
              <xsl:value-of select="'on'"/>
            </xsl:if>
          </xsl:attribute>
        </w:noWrap>
      </xsl:if>

      <xsl:if test="字:单元格边距_41A4">
        <w:tcMar>
          <xsl:call-template name="Cellmar">
            <xsl:with-param name="Cellmar" select="字:单元格边距_41A4"/>
          </xsl:call-template>
        </w:tcMar>
      </xsl:if>

      <xsl:if test="字:是否适应文字_41A9">
        <w:tcFitText>
          <xsl:attribute name="w:val">
            <xsl:if test="字:是否适应文字_41A9='true' or 字:是否适应文字_41A9='1'">
              <xsl:value-of select="'on'"/>
            </xsl:if>
            <xsl:if test="字:是否适应文字_41A9='false' or 字:是否适应文字_41A9='0'">
              <xsl:value-of select="'off'"/>
            </xsl:if>
          </xsl:attribute>
        </w:tcFitText>
      </xsl:if>

      <xsl:if test="字:垂直对齐方式_41A5">
        <w:vAlign>
          <xsl:attribute name="w:val">
            <xsl:value-of select="字:垂直对齐方式_41A5"/>
          </xsl:attribute>
        </w:vAlign>
      </xsl:if>

      <!--2014-03-27，wudi，表格单元格属性里增加文字排列方向的转换，start-->
      <xsl:if test="字:文字排列方向_41AA">
        <w:textDirection>
          <xsl:attribute name="w:val">
            <xsl:call-template name="textDirection2">
              <xsl:with-param name="dir" select="字:文字排列方向_41AA"/>
            </xsl:call-template>
          </xsl:attribute>
        </w:textDirection>
      </xsl:if>
      <!--end-->
      
    </w:tcPr>
  </xsl:template>

  <!--2014-03-27，wudi，表格单元格属性里增加文字排列方向的转换，start-->
  <xsl:template name="textDirection2">
    <xsl:param name="dir"/>
    <xsl:choose>
      <xsl:when test="$dir='r2l-t2b-0e-90w'">
        <xsl:value-of select="'tbRlV'"/>
      </xsl:when>
      <xsl:when test="$dir='r2l-t2b-90e-90w'">
        <xsl:value-of select="'tbRl'"/>
      </xsl:when>
      <xsl:when test="$dir='t2b-l2r-270e-0w'">
        <xsl:value-of select="'lrTbV'"/>
      </xsl:when>
      <xsl:when test="$dir='l2r-b2t-270e-270w'">
        <xsl:value-of select="'btLr'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:message terminate="no">feedback:lost:textDirection_in_Section:Section</xsl:message>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--end-->

  <xsl:template name="tblpPr"><!--文字表属性->绕排-->
    <w:tblpPr>
      <xsl:if test="字:绕排边距_41C6/@左_C608">
        <xsl:attribute name="w:leftFromText">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="字:绕排边距_41C6/@左_C608"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="字:绕排边距_41C6/@右_C60A">
        <xsl:attribute name="w:rightFromText">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="字:绕排边距_41C6/@右_C60A"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="字:绕排边距_41C6/@上_C609">
        <xsl:attribute name="w:topFromText">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="字:绕排边距_41C6/@上_C609"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="字:绕排边距_41C6/@下_C60B">
        <xsl:attribute name="w:bottomFromText">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="字:绕排边距_41C6/@下_C60B"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="字:位置_41C7/uof:垂直_410D/@相对于_C647">
        <xsl:attribute name="w:vertAnchor">
          <xsl:choose>
            <xsl:when test="字:位置_41C7/uof:垂直_410D/@相对于_C647='margin'">
              <xsl:value-of select="'margin'"/>
            </xsl:when>
            <xsl:when test="字:位置_41C7/uof:垂直_410D/@相对于_C647='page'">
              <xsl:value-of select="'page'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'text'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="字:位置_41C7/uof:水平_4106/@相对于_410C">
        <xsl:attribute name="w:horzAnchor">
          <xsl:choose>
            <xsl:when test="字:位置_41C7/uof:水平_4106/@相对于_410C='margin'">
              <xsl:value-of select="'margin'"/>
            </xsl:when>
            <xsl:when test="字:位置_41C7/uof:水平_4106/@相对于_410C='page'">
              <xsl:value-of select="'page'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'text'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="字:位置_41C7/uof:水平_4106/uof:相对_4109/@参考点_410A">
        <xsl:attribute name="w:tblpXSpec">
          <xsl:value-of select="字:位置_41C7/uof:水平_4106/uof:相对_4109/@参考点_410A"/>
        </xsl:attribute>
      </xsl:if>
      <!--这个是做什么的？cxl-->
      <!--<xsl:if test="字:位置_41C7/uof:水平_4106/字:相对/@字:值">
        <xsl:message terminate="no"
          >feedback:lost:floating_table_relative_horizontal_position_in_Table:Table</xsl:message>
      </xsl:if>-->
      <xsl:if test="字:位置_41C7/uof:水平_4106/uof:绝对_4107/@值_4108">
        <xsl:attribute name="w:tblpX">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="字:位置_41C7/uof:水平_4106/uof:绝对_4107/@值_4108"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="字:位置_41C7/uof:垂直_410D/uof:相对_4109/@参考点_410B">
        <xsl:attribute name="w:tblpYSpec">
          <xsl:value-of select="字:位置_41C7/uof:垂直_410D/uof:相对_4109/@参考点_410B"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="字:位置_41C7/uof:垂直_410D/uof:相对_4109/@值_4108">
        <xsl:message terminate="no"
          >feedback:lost:floating_table_relative_vertical_position_in_Table:Table</xsl:message>
      </xsl:if>
      <xsl:if test="字:位置_41C7/uof:垂直_410D/uof:绝对_4107/@值_4108">
        <xsl:attribute name="w:tblpY">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="字:位置_41C7/uof:垂直_410D/uof:绝对_4107/@值_4108"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:if>
    </w:tblpPr>
  </xsl:template>

  <xsl:template name="TableBorder">
    <xsl:param name="Bor"/>
    <xsl:if test="$Bor/uof:上_C614">
      <w:top>
        <xsl:call-template name="border">
          <xsl:with-param name="bor" select="$Bor/uof:上_C614"/>
        </xsl:call-template>
      </w:top>
    </xsl:if>
    
    <!--2015-01-23，wudi，左边框为start，start-->
    <xsl:if test="$Bor/uof:左_C613">
      <w:start>
        <xsl:call-template name="border">
          <xsl:with-param name="bor" select="$Bor/uof:左_C613"/>
        </xsl:call-template>
      </w:start>
    </xsl:if>
    <!--end-->
    
    <xsl:if test="$Bor/uof:下_C616">
      <w:bottom>
        <xsl:call-template name="border">
          <xsl:with-param name="bor" select="$Bor/uof:下_C616"/>
        </xsl:call-template>
      </w:bottom>
    </xsl:if>

    <!--2015-01-23，wudi，右边框为end，start-->
    <xsl:if test="$Bor/uof:右_C615">
      <w:end>
        <xsl:call-template name="border">
          <xsl:with-param name="bor" select="$Bor/uof:右_C615"/>
        </xsl:call-template>
      </w:end>
    </xsl:if>
    <!--end-->
    
    <!--cxl,2012.2.16增加表格边框内部横线与内部竖线-->
    <xsl:if test="$Bor/uof:内部横线_C619">
      <w:insideH>
        <xsl:call-template name="border">
          <xsl:with-param name="bor" select="$Bor/uof:内部横线_C619"/>
        </xsl:call-template>
      </w:insideH>
    </xsl:if>
    <xsl:if test="$Bor/uof:内部竖线_C61A">
      <w:insideV>
        <xsl:call-template name="border">
          <xsl:with-param name="bor" select="$Bor/uof:内部竖线_C61A"/>
        </xsl:call-template>
      </w:insideV>
    </xsl:if>

  </xsl:template>
  <!--cxl,2012.2.16,修改文字表边框与单元格边框冲突问题，当前结点：字:单元格属性_41B7-->
  <xsl:template name="TcBorder">

    <!--2014-05-08，wudi，增加postr参数，记录当前单元格所在行数，start-->
    <xsl:param name="postr"/>
    <!--end-->
    
    <xsl:param name="Bor"/>
    <xsl:if test="$Bor/uof:上_C614 and not($Bor/uof:上_C614/@线型_C60D = ancestor::字:文字表_416C/字:文字表属性_41CC/字:文字表边框_4227/uof:上_C614/@线型_C60D
              and $Bor/uof:上_C614/@虚实_C60E = ancestor::字:文字表_416C/字:文字表属性_41CC/字:文字表边框_4227/uof:上_C614/@虚实_C60E 
              and $Bor/uof:上_C614/@宽度_C60F = ancestor::字:文字表_416C/字:文字表属性_41CC/字:文字表边框_4227/uof:上_C614/@宽度_C60F
              and $Bor/uof:上_C614/@颜色_C611 = ancestor::字:文字表_416C/字:文字表属性_41CC/字:文字表边框_4227/uof:上_C614/@颜色_C611)">
      <w:top>
        <xsl:if test="$Bor/@阴影类型_C645">
          <xsl:attribute name="w:shadow">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:call-template name="border">
          <xsl:with-param name="bor" select="$Bor/uof:上_C614"/>
        </xsl:call-template>
      </w:top>
    </xsl:if>
    
    <!--2014-05-19，wudi，之前的修改会带来新的问题，还原代码，start-->
    <!--2014-05-08，wudi，修复表格边框转换BUG，内部边框部分丢失的问题，start-->
    <xsl:if test="$Bor/uof:左_C613 and not($Bor/uof:左_C613/@线型_C60D = ancestor::字:文字表_416C//字:文字表边框_4227/uof:左_C613/@线型_C60D
              and $Bor/uof:左_C613/@虚实_C60E = ancestor::字:文字表_416C//字:文字表边框_4227/uof:左_C613/@虚实_C60E 
              and $Bor/uof:左_C613/@宽度_C60F = ancestor::字:文字表_416C//字:文字表边框_4227/uof:左_C613/@宽度_C60F
              and $Bor/uof:左_C613/@颜色_C611 = ancestor::字:文字表_416C//字:文字表边框_4227/uof:左_C613/@颜色_C611)">
    <!--<xsl:if test="$Bor/uof:左_C613">-->

      <!--2015-01-23，wudi，左边框为start，start-->
      <w:start>
        <xsl:if test="$Bor/@阴影类型_C645">
          <xsl:attribute name="w:shadow">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:call-template name="border">
          <xsl:with-param name="bor" select="$Bor/uof:左_C613"/>
        </xsl:call-template>
      </w:start>
      <!--end-->
      
    </xsl:if>
    <!--end-->
    <!--end-->

    <!--2014-05-19，wudi，之前的修改会带来新的问题，还原代码，start-->
    <!--2014-05-08，wudi，修复表格边框转换BUG，首行边框存在丢失的问题，start-->
    <!--2014-05-08，wudi，修复表格边框转换BUG，内部边框部分丢失的问题，start-->
    <xsl:if test="$Bor/uof:下_C616 and not($Bor/uof:下_C616/@线型_C60D = ancestor::字:文字表_416C//字:文字表边框_4227/uof:下_C616/@线型_C60D
              and $Bor/uof:下_C616/@虚实_C60E = ancestor::字:文字表_416C//字:文字表边框_4227/uof:下_C616/@虚实_C60E 
              and $Bor/uof:下_C616/@宽度_C60F = ancestor::字:文字表_416C//字:文字表边框_4227/uof:下_C616/@宽度_C60F
              and $Bor/uof:下_C616/@颜色_C611 = ancestor::字:文字表_416C//字:文字表边框_4227/uof:下_C616/@颜色_C611 and $postr != '1')">
    <!--<xsl:if test="$Bor/uof:下_C616">-->
      <w:bottom>
        <xsl:if test="$Bor/@阴影类型_C645">
          <xsl:attribute name="w:shadow">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:call-template name="border">
          <xsl:with-param name="bor" select="$Bor/uof:下_C616"/>
        </xsl:call-template>
      </w:bottom>
    </xsl:if>
    <!--end-->
    <!--end-->
    <!--end-->

    <!--2014-05-19，wudi，之前的修改会带来新的问题，还原代码，start-->
    <!--2014-05-08，wudi，修复表格边框转换BUG，内部边框部分丢失的问题，start-->
    <xsl:if test="$Bor/uof:右_C615 and not($Bor/uof:右_C615/@线型_C60D = ancestor::字:文字表_416C//字:文字表边框_4227/uof:右_C615/@线型_C60D
              and $Bor/uof:右_C615/@虚实_C60E = ancestor::字:文字表_416C//字:文字表边框_4227/uof:右_C615/@虚实_C60E 
              and $Bor/uof:右_C615/@宽度_C60F = ancestor::字:文字表_416C//字:文字表边框_4227/uof:右_C615/@宽度_C60F
              and $Bor/uof:右_C615/@颜色_C611 = ancestor::字:文字表_416C//字:文字表边框_4227/uof:右_C615/@颜色_C611)">
    <!--<xsl:if test="$Bor/uof:右_C615">-->
      
      <!--2015-01-23，wudi，右边框为end，start-->
      <w:end>
        <xsl:if test="$Bor/@阴影类型_C645">
          <xsl:attribute name="w:shadow">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:call-template name="border">
          <xsl:with-param name="bor" select="$Bor/uof:右_C615"/>
        </xsl:call-template>
      </w:end>
      <!--end-->
      
    </xsl:if>
    <!--end-->
    <!--end-->

    <!--Edit by LEE 修复对角线不显示的问题-->

    <xsl:if test="$Bor/uof:对角线1_C617[@线型_C60D!='none'] or ./字:斜线表头_41AB//字:起点_41AD='top-left'">
      <w:tl2br>
        <xsl:call-template name="border">
          <xsl:with-param name="bor" select="$Bor/uof:对角线1_C617"/>
        </xsl:call-template>
      </w:tl2br>
    </xsl:if>
    <xsl:if test="$Bor/uof:对角线2_C618[@线型_C60D!='none'] or ./字:斜线表头_41AB//字:起点_41AD='top-right'">
      <w:tr2bl>
        <xsl:call-template name="border">
          <xsl:with-param name="bor" select="$Bor/uof:对角线2_C618"/>
        </xsl:call-template>
      </w:tr2bl>
    </xsl:if>

    <!--阴影标准变化较大**********************************************************-->
    <!--<xsl:if test="./@阴影类型_C645">
      <xsl:attribute name="w:shadow">
        <xsl:if test="./@阴影类型_C645='true' or $bor/@uof:阴影='1'">
          <xsl:value-of select="'on'"/>
        </xsl:if>
        <xsl:if test="./@阴影类型_C645='false' or $bor/@uof:阴影='0'">
          <xsl:value-of select="'off'"/>
        </xsl:if>
      </xsl:attribute>
    </xsl:if>-->
  </xsl:template>
 
  <!--
  在ISO中，table属性不需要之名对角线，此模板为空
  <xsl:template name="uof:内部横线">
  </xsl:template>
  <xsl:template name="uof:内部竖线">
  </xsl:template>
  -->


  <!--
  1.1中类型转换为线型和虚实，由虚实判断ISO中的类型
  edit by Lee 2011-04-28
  -->
  <xsl:template name="border">
    <xsl:param name="bor"/>
    <xsl:attribute name="w:val">
      <xsl:choose>
        <!--
        <xsl:when test="$bor/@uof:线型='single' and $bor/@uof:虚实='？'">
          <xsl:value-of select="''"/>
        </xsl:when>
        -->
        <xsl:when test="$bor/@线型_C60D='single' and $bor/@虚实_C60E='square-dot'">
          <xsl:value-of select="'dotted'"/>
        </xsl:when>
        
        <!--2013-05-10，wudi，修改'nil'为'none'-->
        
        <!--2014-04-15，wudi，还原原来的取值，start-->
        <xsl:when test="$bor/@线型_C60D='none'">
          <xsl:value-of select="'nil'"/>
        </xsl:when>
        <!--end-->
        
        <!--end-->
        
        <xsl:when test="$bor/@线型_C60D='single' and $bor/@虚实_C60E='solid'">
          <xsl:value-of select="'single'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='single' and $bor/@虚实_C60E='round-dot'">
          <xsl:value-of select="'dashSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='single' and $bor/@虚实_C60E='dash'">
          <xsl:value-of select="'dashed'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='single' and $bor/@虚实_C60E='dash-dot'">
          <xsl:value-of select="'dotDash'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='single' and $bor/@虚实_C60E='dash-dot-dot'">
          <xsl:value-of select="'dotDotDash'"/>
        </xsl:when>

        <!--dotted-heavy这一条在ISO-UOF中可以不考虑 Lee-->
        <xsl:when test="$bor/@线型_C60D='dotted-heavy' or $bor/@虚实_C60E='dotted-heavy'">
          <xsl:value-of select="'dotted'"/>
        </xsl:when>

        <!--case only have double-->
        <!--2013-03-28，wudi，增加限制条件and not($bor/@虚实_C60E)-->
        <xsl:when test="$bor/@线型_C60D='double' and not($bor/@虚实_C60E)">
          <xsl:value-of select="'double'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='double' and $bor/@虚实_C60E='solid'">
          <xsl:value-of select="'double'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='double' and $bor/@虚实_C60E='dash-dot-dot'">
          <xsl:value-of select="'triple'"/>
        </xsl:when>

        <!--2014-05-19，wudi，修复表格边框转换BUG，之前没考虑这种情况，start-->
        <xsl:when test="$bor/@线型_C60D='thick' and $bor/@虚实_C60E='round-dot'">
          <xsl:value-of select="'dashSmallGap'"/>
        </xsl:when>
        <!--end-->

        <!--2014-01-08，wudi，表格边框，粗细线，细粗线转换有误，修正，start-->
        <xsl:when test="$bor/@线型_C60D='thick-thin' and $bor/@虚实_C60E='solid'">
          <xsl:value-of select="'thinThickSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='thick-thin' and not($bor/@虚实_C60E)">
          <xsl:value-of select="'thinThickSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='thin-thick' and $bor/@虚实_C60E='solid'">
          <xsl:value-of select="'thickThinSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='thin-thick' and not($bor/@虚实_C60E)">
          <xsl:value-of select="'thickThinSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='thick-thin' and $bor/@虚实_C60E='dash'">
          <xsl:value-of select="'thinThickMediumGap'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='thin-thick' and $bor/@虚实_C60E='dash'">
          <xsl:value-of select="'thickThinMediumGap'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='thin-thick' and $bor/@虚实_C60E='long-dash'">
          <xsl:value-of select="'thickThinLargeGap'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='thick-thin' and $bor/@虚实_C60E='long-dash'">
          <xsl:value-of select="'thinThickLargeGap'"/>
        </xsl:when>
        <!--end-->
        
        <xsl:when test="$bor/@线型_C60D='thick-between-thin' and $bor/@虚实_C60E='solid'">
          <xsl:value-of select="'thinThickThinSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='thick-between-thin' and not($bor/@虚实_C60E)='solid'">
          <xsl:value-of select="'thinThickThinSmallGap'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='thick-between-thin' and $bor/@虚实_C60E='long-dash'">
          <xsl:value-of select="'thinThickThinMediumGap'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='thick-between-thin' and $bor/@虚实_C60E='dash'">
          <xsl:value-of select="'thinThickMediumGap'"/>
        </xsl:when>
        
        <!--2013-03-28，wudi，修复表格边框BUG，start-->
        <xsl:when test="$bor/@线型_C60D='single' and $bor/@虚实_C60E='wave'">
          <xsl:value-of select="'wave'"/>
        </xsl:when>
        <xsl:when test="$bor/@线型_C60D='double' and $bor/@虚实_C60E='wave'">
          <xsl:value-of select="'doubleWave'"/>
        </xsl:when>
        <!--end-->
        
        <xsl:when test="$bor/@线型_C60D='dash-dot'">
          <xsl:value-of select="'dashDotStroked'"/>
        </xsl:when>
        <!--以上枚举 测试没问题 Lee-->
        <xsl:otherwise>
          <xsl:value-of select="'single'"/>
          <xsl:message terminate="no">feedback:lost:table_border_style_in_Table:Table</xsl:message>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>

    <xsl:if test="$bor/@颜色_C611">
      <xsl:attribute name="w:color">
        <xsl:if test="$bor/@颜色_C611!='auto'">
          <xsl:value-of select="substring-after($bor/@颜色_C611,'#')"/>
        </xsl:if>
        <xsl:if test="$bor/@颜色_C611='auto'">
          <xsl:value-of select="'auto'"/>
        </xsl:if>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="$bor/@宽度_C60F">
      <xsl:attribute name="w:sz">
        <xsl:call-template name="eighthPointMeasure">
          <xsl:with-param name="lengthVal" select="$bor/@宽度_C60F"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="$bor/@边距_C610">
      <xsl:attribute name="w:space">
        <xsl:call-template name="pointMeasure">
          <xsl:with-param name="lengthVal" select="$bor/@边距_C610"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    
  </xsl:template>

  <xsl:template name="Cellmar"><!--单元格宽度高度模板-->
    <xsl:param name="Cellmar"/>
    <xsl:if test="$Cellmar/@上_C609">
      <w:top>
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="$Cellmar/@上_C609"/>
        </xsl:call-template>
      </w:top>
    </xsl:if>
    <xsl:if test="$Cellmar/@左_C608">
      <w:start>
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="$Cellmar/@左_C608"/>
        </xsl:call-template>
      </w:start>
    </xsl:if>
    <xsl:if test="$Cellmar/@下_C60B">
      <w:bottom>
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="$Cellmar/@下_C60B"/>
        </xsl:call-template>
      </w:bottom>
    </xsl:if>
    <xsl:if test="$Cellmar/@右_C60A">
      <w:end>
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="$Cellmar/@右_C60A"/>
        </xsl:call-template>
      </w:end>
    </xsl:if>
  </xsl:template>

  <xsl:template name="cellmar">
    <xsl:param name="cellmar"/>
    <xsl:attribute name="w:w">
      <!--xsl:value-of select="($cellmar) * 20"/-->
      <xsl:call-template name="twipsMeasure">
        <xsl:with-param name="lengthVal" select="$cellmar"/>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="w:type">
      <xsl:value-of select="'dxa'"/>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name="twipsMeasure-pct"><!--cxl宽度计算模板（根据不同单位）-->
    <xsl:param name="lengthVal"/>
    <xsl:variable name="unit" select="/uof:UOF/uof:文字处理/规则:公用处理规则_B665/规则:长度单位_B666"/>
    <xsl:choose>

      <!--2015-01-23，wudi，取值差异，start-->
      <xsl:when test="$unit='pt'">
        <xsl:value-of select="concat($lengthVal,'%')"/>
      </xsl:when>
      <!--end-->
      
      <xsl:when test="$unit='cm'">
        <xsl:value-of select="round(number($lengthVal) * 28.3 * 50)"/>
      </xsl:when>
      <xsl:when test="$unit='mm'">
        <xsl:value-of select="round(number($lengthVal) * 2.83 * 50)"/>
      </xsl:when>
      <xsl:when test="$unit='inch'">
        <xsl:value-of select="round(number($lengthVal) * 72 * 50)"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>
