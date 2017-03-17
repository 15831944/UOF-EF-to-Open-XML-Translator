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

  <xsl:import href="header-footer.xsl"/>
  <xsl:import href="styles.xsl"/>
  <xsl:import href="paragraph.xsl"/>
  <xsl:import href="common.xsl"/>

  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

  <!--表格转换 2011-02-22--><!--Edit by LiJG 2011-03-18-->
  <xsl:template name="table">
    <xsl:param name="tblPartFrom"/>
    <字:文字表_416C>
      <xsl:for-each select="w:tblPr | w:tr | w:sdt | w:customXml">
        <xsl:choose>
          <xsl:when test="name(.)='w:tblPr'">
            <字:文字表属性_41CC>
              <xsl:if test="w:tblStyle">
                <xsl:attribute name="式样引用_419C">
                  <xsl:call-template name="IdProducer">
                    <xsl:with-param name="ooxId" select="./w:tblStyle/@w:val"/>
                  </xsl:call-template>
                </xsl:attribute>
              </xsl:if>
              <xsl:call-template name="tblPr"/>
            </字:文字表属性_41CC>
          </xsl:when>

          <!--表行属性-->
          <xsl:when test="name(.)='w:tr'">

            <!--2014-06-05，wudi，处理表格中某行被修订删除的情况，此行不做转换，start-->
            <xsl:if test ="not(./w:trPr/w:del)">
              <xsl:call-template name="tblRow">
                <xsl:with-param name="trPartFrom" select="$tblPartFrom"/>
              </xsl:call-template>
            </xsl:if>
            <!--end-->
            
          </xsl:when>

          <xsl:when test="name(.)='w:sdt'">
            <xsl:call-template name="sdtContentRow"/>
          </xsl:when>

          <xsl:when test="name(.)='w:customXml'">
            <xsl:call-template name="CustomXmlRow"/>
          </xsl:when>

          <!--<xsl:when test="name(.)='w:bookmarkEnd'">
            <xsl:call-template name="bookmarkEnd"/>
          </xsl:when>-->

        </xsl:choose>
      </xsl:for-each>
    </字:文字表_416C>
  </xsl:template>

  <!--文字表属性_41CC-->
  <xsl:template name="tblPr">
    <xsl:if test="w:tblW[@w:type!='auto']">
      <字:宽度_41A1>
        <xsl:if test="w:tblW/@w:type='dxa'">
          <xsl:attribute name="绝对宽度_41BF">
            
            <!--2013-11-14，wudi，Strict标准下，w:w取值不同，修正，start-->
            <xsl:if test="w:tblW/@w:w">
              <xsl:value-of select="substring-before(w:tblW/@w:w,'pt')"/>
            </xsl:if>
            <!--end-->
            
            <xsl:if test="not(w:tblW/@w:w)">
              <xsl:value-of select="'0'"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="w:tblW/@w:type='nil'">
          <xsl:attribute name="绝对宽度_41BF">
            <xsl:value-of select="'0'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="w:tblW/@w:type='pct'">
          <xsl:attribute name="相对宽度_41C0">
            
            <!--2013-11-14，wudi，Strict标准下w:w取值不同，修正，start-->
            <xsl:if test="w:tblW/@w:w">
              <xsl:value-of select="substring-before(w:tblW/@w:w,'%')"/>
            </xsl:if>
            <!--end-->
            
            <xsl:if test="not(w:tblW/@w:w)">
              <xsl:value-of select="'0'"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:if>
      </字:宽度_41A1>
    </xsl:if>

    <xsl:if test="following-sibling::w:tblGrid">
      <xsl:call-template name="tblGrid"/>
    </xsl:if>

    <xsl:if test="w:jc">
      <字:对齐_41C3>
        <xsl:choose>
          
          <!--2013-11-14，wudi，Strict标准下，表格对齐方式，左对齐对应start，右对齐对应end，修正，start-->
          <xsl:when test="w:jc/@w:val='start'">
            <xsl:value-of select="'left'"/>
          </xsl:when>
          <xsl:when test="w:jc/@w:val='end'">
            <xsl:value-of select="'right'"/>
          </xsl:when>
          <!--end-->
          
          <xsl:when test="w:jc/@w:val='center'">
            <xsl:value-of select="'center'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'left'"/>
          </xsl:otherwise>
        </xsl:choose>
      </字:对齐_41C3>
    </xsl:if>

    <xsl:if test="w:tblInd[@w:type!='auto']">
      <字:左缩进_41C4>

        <!--2013-10-30,wudi,w:pt在strict和transitonal标准下取值不同，加以区分，start-->
        <xsl:if test="w:tblInd/@w:type='dxa' and w:tblInd/@w:w">
          <!--<xsl:if test ="not(contains(w:tblInd/@w:w,'pt'))">
            <xsl:value-of select="(w:tblInd/@w:w) div 20"/>
          </xsl:if>
          <xsl:if test ="contains(w:tblInd/@w:w,'pt')">-->
            <xsl:value-of select="substring-before(w:tblInd/@w:w,'pt')"/>
          <!--</xsl:if>-->
        </xsl:if>
        <!--end-->

        <xsl:if test="w:tblInd/@w:type='nil' or not(w:tblInd/@w:w)">
          <xsl:value-of select="'0'"/>
        </xsl:if>

        <!--2013-10-30,wudi,w:w在strict和transitonal标准下取值不同，加以区分，start-->
        <xsl:if test="w:tblInd/@w:type='pct' and w:tblInd/@w:w">
          <!--<xsl:if test ="not(contains(w:tblInd/@w:w,'pt'))">
            <xsl:value-of select="(w:tblInd/@w:w) * 0.08522"/>
          </xsl:if>-->
          <!--<xsl:if test ="contains(w:tblInd/@w:w,'pt')">-->
            <xsl:value-of select="substring-before(w:tblInd/@w:w,'pt')"/>
          <!--</xsl:if>-->
        </xsl:if>
        <!--end-->

      </字:左缩进_41C4>
    </xsl:if>
    <xsl:if test="w:tblpPr">
      <xsl:call-template name="tblpPr"/>
    </xsl:if>
    <xsl:call-template name="tblBorder"/>
    <xsl:if test="w:shd[@w:val!='nil']">
      <xsl:apply-templates select="w:shd" mode="tbShd"/>
    </xsl:if>
    <xsl:choose>
      <xsl:when test="w:tblLayout">
        <字:是否自动调整大小_41C8>   
            <xsl:choose>
              <xsl:when test="w:tblLayout/@w:type='autofit'">
                <xsl:value-of select="'true'"/>
              </xsl:when>
              <xsl:when test="w:tblLayout/@w:type='fixed'">
                <xsl:value-of select="'false'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'true'"/>
              </xsl:otherwise>
            </xsl:choose>       
        </字:是否自动调整大小_41C8>
      </xsl:when>
      <xsl:otherwise>
        <字:是否自动调整大小_41C8>true</字:是否自动调整大小_41C8>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:if test="w:tblCellMar">
      <字:默认单元格边距_41CA>
        <xsl:call-template name="Cellmar">
          <xsl:with-param name="Cellmar" select="w:tblCellMar"/>
        </xsl:call-template>
      </字:默认单元格边距_41CA>
    </xsl:if>
    <xsl:if test="w:tblCellSpacing[@w:type!='auto']">
      <字:默认单元格间距_41CB>

        <!--2013-10-30,wudi,w:w在strict和transitonal标准下取值不同，加以区分，start-->
        <xsl:if test="w:tblCellSpacing/@w:type='dxa' and w:tblCellSpacing/@w:w">
          <!--<xsl:if test ="not(contains(w:tblCellSpacing/@w:w,'pt'))">
            <xsl:value-of select="(w:tblCellSpacing/@w:w) div 20"/>
          </xsl:if>-->
          
          <!--2013-11-14，wudi，计算公式有误，需要乘以倍数2，start-->
          <!--<xsl:if test ="contains(w:tblCellSpacing/@w:w,'pt')">-->
            <xsl:value-of select="substring-before(w:tblCellSpacing/@w:w,'pt') * 2"/>
          <!--</xsl:if>-->
          <!--end-->
          
        </xsl:if>
        <!--end-->
        
        <xsl:if test="w:tblCellSpacing/@w:type='nil'or not(w:tblCellSpacing/@w:w)">
          <xsl:value-of select="'0'"/>
        </xsl:if>

        <!--2013-10-30,wudi,w:w在strict和transitonal标准下取值不同，加以区分，start-->
        <xsl:if test="w:tblCellSpacing/@w:type='pct' and w:tblCellSpacing/@w:w">
          <!--<xsl:if test ="not(contains(w:tblCellSpacing/@w:w,'pt'))">
            <xsl:value-of select="(w:tblCellSpacing/@w:w) * 0.08522"/>
          </xsl:if>-->
          <!--<xsl:if test ="contains(w:tblCellSpacing/@w:w,'pt')">-->
            <xsl:value-of select="substring-before(w:tblCellSpacing/@w:w,'pt')"/>
          <!--</xsl:if>-->
        </xsl:if>
        <!--end-->
        
      </字:默认单元格间距_41CB>
    </xsl:if>
  </xsl:template>

  <xsl:template match="w:shd" mode="tbShd">
    <字:填充_4134>
      <xsl:call-template name="shd"/>
    </字:填充_4134>
  </xsl:template>

  <xsl:template name="tblGrid">
    <字:列宽集_41C1>
      <xsl:for-each select="following-sibling::w:tblGrid/w:gridCol">
        <字:列宽_41C2>
          <xsl:if test="@w:w">
            <xsl:value-of select="@w:w div 20"/>
          </xsl:if>
          <xsl:if test="not(@w:w)">
            <xsl:value-of select="'0'"/>
          </xsl:if>
        </字:列宽_41C2>
      </xsl:for-each>
      <xsl:if test="not(following-sibling::w:tblGrid/w:gridCol)">
        <字:列宽_41C2>
          <xsl:value-of select="'0'"/>
        </字:列宽_41C2>
      </xsl:if>
    </字:列宽集_41C1>
  </xsl:template>


  <xsl:template name="tblRow">
    <xsl:param name="trPartFrom"/>
    <字:行_41CD>
      <xsl:variable name="postr">
        <xsl:number count="w:tr" format="1" level="single"/>
      </xsl:variable>
      
      <xsl:call-template name="tblRowPr"/>
      <!--cxl,2012.5.15删除w:gridBefore和w:gridAfter的转换，因为两边实现方式不同，UOF中没有对应的元素-->
      <!--<xsl:if test="./w:trPr/w:gridBefore">
        <字:单元格_41BE>
          <字:单元格属性_41B7>
            <xsl:if test="./w:trPr/w:gridBefore/@w:val!='1'">
              <字:跨列_41A7>               
                  <xsl:value-of select="./w:trPr/w:gridBefore/@w:val"/>    
              </字:跨列_41A7>
            </xsl:if>
          </字:单元格属性_41B7>
        </字:单元格_41BE>
      </xsl:if>-->

      <xsl:for-each select="w:tc | w:sdt | w:customXml">
        <xsl:choose>

          <!--2013-03-25，wudi，修复表格边框转换错误，start-->
          <xsl:when test="name(.)='w:tc'">
            <xsl:variable name ="postc1">
              <xsl:number count="w:tc" format="1" level="single"/>
            </xsl:variable>
            <xsl:variable name ="postd">
              <xsl:if test ="preceding-sibling::w:sdt">
                <xsl:value-of select ="count(preceding-sibling::w:sdt)"/>
              </xsl:if>
              <xsl:if test ="not(preceding-sibling::w:sdt)">
                <xsl:value-of select ="0"/>
              </xsl:if>
            </xsl:variable>
            <xsl:variable name ="postc">
              <xsl:value-of select ="number($postc1) + number($postd)"/>
            </xsl:variable>
            <xsl:call-template name="tblCell">
              <xsl:with-param name="postr" select="$postr"/>
              <xsl:with-param name ="postc2" select ="$postc"/>
              <xsl:with-param name="tcPartFrom" select="$trPartFrom"/>
            </xsl:call-template>
            <!--end-->
            
          </xsl:when>

          <!--<xsl:when test="name(.)='w:trPr'">
            <xsl:if test="./w:gridAfter">
              <字:单元格_41BE>
                <字:单元格属性_41B7>
                  <xsl:if test="./w:gridAfter/@w:val!='1'">
                    <字:跨列_41A7>                     
                        <xsl:value-of select="./w:gridAfter/@w:val"/>                  
                    </字:跨列_41A7>
                  </xsl:if>
                </字:单元格属性_41B7>
              </字:单元格_41BE>
            </xsl:if>
          </xsl:when>-->

          <xsl:when test="name(.)='w:sdt'">

            <!--2013-03-25，wudi，修复表格边框转换错误，start-->
            <xsl:variable name ="postd1">
              <xsl:number count="w:sdt" format="1" level="single"/>
            </xsl:variable>
            <xsl:variable name ="poswtc">
              <xsl:if test ="preceding-sibling::w:tc">
                <xsl:value-of select ="count(preceding-sibling::w:tc)"/>
              </xsl:if>
              <xsl:if test ="not(preceding-sibling::w:tc)">
                <xsl:value-of select ="0"/>
              </xsl:if>
            </xsl:variable>
            <xsl:variable name ="postd">
              <xsl:value-of select ="number($postd1) + number($poswtc)"/>
            </xsl:variable>
            <xsl:call-template name="sdtContentCell">
              <xsl:with-param name ="postr" select ="$postr"/>
              <xsl:with-param name ="postd" select ="$postd"/>
              <xsl:with-param name ="tcPartFrom" select ="$trPartFrom"/>
            </xsl:call-template>
            <!--end-->
            
          </xsl:when>

          <xsl:when test="name(.)='w:customXml'">
            <xsl:call-template name="CustomXmlCell"/>
          </xsl:when>

        </xsl:choose>
      </xsl:for-each>
    </字:行_41CD>
  </xsl:template>

  <!--表格行属性-->
  <xsl:template name="tblRowPr">
    <字:表行属性_41BD>
      <xsl:if test="w:trPr/w:trHeight">
        <字:高度_41B8>
          <xsl:choose>
            <xsl:when test="w:trPr/w:trHeight/@w:hRule='atLeast'">
              <xsl:attribute name="最小值_41BA">
                <xsl:if test="w:trPr/w:trHeight/@w:val">
                  <xsl:value-of select="(w:trPr/w:trHeight/@w:val) div 20"/>
                </xsl:if>
                <xsl:if test="not(w:trPr/w:trHeight/@w:val)">
                  <xsl:value-of select="'0'"/>
                </xsl:if>
              </xsl:attribute>
            </xsl:when>

            <xsl:when test="w:trPr/w:trHeight/@w:hRule='exact'">
              <xsl:attribute name="固定值_41B9">
                <xsl:if test="w:trPr/w:trHeight/@w:val">
                  <xsl:value-of select="(w:trPr/w:trHeight/@w:val) div 20"/>
                </xsl:if>
                <xsl:if test="not(w:trPr/w:trHeight/@w:val)">
                  <xsl:value-of select="'0'"/>
                </xsl:if>
              </xsl:attribute>
            </xsl:when>

            <xsl:when test="w:trPr/w:trHeight/@w:hRule='auto'">
              <xsl:attribute name="最小值_41BA">
                <xsl:if test="w:trPr/w:trHeight/@w:val">
                  <xsl:value-of select="(w:trPr/w:trHeight/@w:val) div 20"/>
                </xsl:if>
                <xsl:if test="not(w:trPr/w:trHeight/@w:val)">
                  <xsl:value-of select="'0'"/>
                </xsl:if>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="最小值_41BA">
                <xsl:if test="w:trPr/w:trHeight/@w:val">
                  <xsl:value-of select="(w:trPr/w:trHeight/@w:val) div 20"/>
                </xsl:if>
                <xsl:if test="not(w:trPr/w:trHeight/@w:val)">
                  <xsl:value-of select="'0'"/>
                </xsl:if>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </字:高度_41B8>
      </xsl:if>

      <xsl:choose>
        <xsl:when test="w:trPr/w:cantSplit">
          <字:是否跨页_41BB>
              <xsl:choose>
                <xsl:when
                  test="w:trPr/w:cantSplit/@w:val='off' or w:trPr/w:cantSplit/@w:val='0' or w:trPr/w:cantSplit/@w:val='false'">
                  <xsl:value-of select="'true'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'false'"/>
                </xsl:otherwise>
              </xsl:choose>
          </字:是否跨页_41BB>
        </xsl:when>
        <xsl:otherwise>
          <字:是否跨页_41BB>true</字:是否跨页_41BB>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:if test="w:trPr/w:tblHeader">
        <字:是否表头行_41BC>
            <xsl:choose>
              <xsl:when
                test="w:trPr/w:tblHeader/@w:val='off' or w:trPr/w:tblHeader/@w:val='0' or w:trPr/w:tblHeader/@w:val='false'">
                <xsl:value-of select="'false'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'true'"/>
              </xsl:otherwise>
            </xsl:choose>
        </字:是否表头行_41BC>
      </xsl:if>
    </字:表行属性_41BD>
  </xsl:template>

  <xsl:template name="tblCell">

    <!--2013-03-25，wudi，修复表格边框转换错误，start-->
    <xsl:param name="postr"/>
    <xsl:param name ="postd"/>
    <xsl:param name ="postc2"/>
    <xsl:param name="tcPartFrom"/>
    <字:单元格_41BE>
      <xsl:variable name="postc">
        <xsl:if test ="ancestor::w:sdt">
          <xsl:value-of select ="$postd"/>
        </xsl:if>
        <xsl:if test ="not(ancestor::w:sdt)">
          <xsl:value-of select ="$postc2"/>
        </xsl:if>
      </xsl:variable>
      <!--end-->
      
      <xsl:for-each select="w:tcPr | w:p | w:tbl | w:sdt | w:customXml">
        <xsl:choose>
          <xsl:when test="name(.)='w:tcPr'">
            <xsl:call-template name="tblCellPr">
              <xsl:with-param name="postc" select="$postc"/>
              <xsl:with-param name="postr" select="$postr"/>
            </xsl:call-template>
          </xsl:when>

          <xsl:when test="name(.)='w:p'">
            <xsl:call-template name="paragraph">
              <xsl:with-param name="pPartFrom" select="$tcPartFrom"/>

              <!--2014-05-05，wudi，增加postr参数，记录单元格所属行数，修复单元格字体加粗效果转换BUG，start-->
              <xsl:with-param name="postr" select="$postr"/>
              <xsl:with-param name="postc" select="$postc"/>
              <!--end-->

            </xsl:call-template>
          </xsl:when>

          <xsl:when test="name(.)='w:tbl'">
            <xsl:call-template name="table">
              <xsl:with-param name="tblPartFrom" select="$tcPartFrom"/>
            </xsl:call-template>
          </xsl:when>

          <xsl:when test="name(.)='w:sdt'">

            <!--2014-03-24，wudi，修复表格转换BUG，单元格里存在sdt节点的情况，start-->
            <xsl:for-each select="w:sdtContent">
              <xsl:call-template name="sdtContentBlock"/>
            </xsl:for-each>
            <!--end-->
            
          </xsl:when>

          <xsl:when test="name(.)='w:customXml'">
            <xsl:call-template name="CustomXmlBlock"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </字:单元格_41BE>
  </xsl:template>

  <xsl:template name="tblCellPr">
    <xsl:param name="postc"/>
    <xsl:param name="postr"/>
    <xsl:variable name="lasttr" select="count(ancestor::w:tbl/w:tr)"/>

    <!--2013-03-25，wudi，修复表格边框转换错误，start-->
    <xsl:variable name ="lasttc">
      <xsl:variable name ="tmp1">
        <xsl:value-of select ="count(ancestor::w:tr/w:sdt)"/>
      </xsl:variable>
      <xsl:variable name ="tmp2">
        <xsl:value-of select ="count(ancestor::w:tr/w:tc)"/>
      </xsl:variable>
      <xsl:value-of select ="number($tmp1) + number($tmp2)"/>
    </xsl:variable>

    <!--<postc>
      <xsl:value-of select ="$postc"/>
    </postc>
    <postr>
      <xsl:value-of select ="$postr"/>
    </postr>
    <lasttr>
      <xsl:value-of select ="$lasttr"/>
    </lasttr>
    <lasttc>
      <xsl:value-of select ="$lasttc"/>
    </lasttc>-->
    <!--end-->

    <字:单元格属性_41B7>
      <!--cxl,2012.3.9新增单元格垂直合并转换-->
      <xsl:if test="./w:vMerge/@rowSpanCount">
        <字:跨行_41A6>
          <xsl:value-of select="./w:vMerge/@rowSpanCount"/>
        </字:跨行_41A6>
      </xsl:if>
      <xsl:if test="w:tcW[@w:type!='auto']">
        <字:宽度_41A1>
          <xsl:if test="w:tcW/@w:type='dxa'">
            <xsl:attribute name="绝对值_41A2">
              
              <!--2013-11-14，wudi，Strict标准下，w:w取值不同，修正，start-->
              <xsl:if test="w:tcW/@w:w">
                <xsl:value-of select="substring-before(w:tcW/@w:w,'pt')"/>
              </xsl:if>
              <!--end-->
              
              <xsl:if test="not(w:tcW/@w:w)">
                <xsl:value-of select="'0'"/>
              </xsl:if>
            </xsl:attribute>
          </xsl:if>

          <xsl:if test="w:tcW/@w:type='nil'">
            <xsl:attribute name="绝对值_41A2">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
          </xsl:if>

          <xsl:if test="w:tcW/@w:type='pct'">
            <xsl:attribute name="相对值_41A3">
              
              <!--2013-11-14，wudi，Strict标准下w:w取值不同，修正，start-->
              <xsl:if test="w:tcW/@w:w">
                <xsl:value-of select="substring-before(w:tcW/@w:w,'%')"/>
              </xsl:if>
              <!--end-->
              
              <xsl:if test="not(w:tcW/@w:w)">
                <xsl:value-of select="'0'"/>
              </xsl:if>
            </xsl:attribute>
          </xsl:if>
        </字:宽度_41A1>
      </xsl:if>

      <xsl:if test="w:tcMar">
        <字:单元格边距_41A4>
          <xsl:call-template name="Cellmar">
            <xsl:with-param name="Cellmar" select="w:tcMar"/>
          </xsl:call-template>
        </字:单元格边距_41A4>
      </xsl:if>

      <xsl:call-template name="tcBorder"><!--单元格边框：边框_4133-->
        <xsl:with-param name="postr" select="$postr"/>
        <xsl:with-param name="postc" select="$postc"/>
        <xsl:with-param name="lasttr" select="$lasttr"/>
        <xsl:with-param name="lasttc" select="$lasttc"/>
      </xsl:call-template>

      <xsl:if test="w:shd[@w:val!='nil']">
        <xsl:apply-templates select="w:shd" mode="tcShd"/>
      </xsl:if>

      <!--2013-03-26，wudi，修复表格填充的BUG，start-->
      <xsl:variable name ="styId">
        <xsl:value-of select ="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
      </xsl:variable>
      <xsl:if test ="(number(not(w:shd)) + number(ancestor::w:tr/w:trPr/w:cnfStyle/@w:firstRow ='1') + number(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr/@w:type ='firstRow'))&gt;=3">
        <xsl:variable name ="wfill">
          <xsl:value-of select ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='firstRow']/w:tcPr/w:shd/@w:fill"/>
        </xsl:variable>
        <xsl:variable name ="color">
          <xsl:value-of select ="concat('#',$wfill)"/>
        </xsl:variable>
        <字:填充_4134>
          <图:图案_800A>
            <xsl:attribute name ="背景色_800C">
              <xsl:value-of select ="$color"/>
            </xsl:attribute>
          </图:图案_800A>
        </字:填充_4134>
      </xsl:if>
      <xsl:if test ="(number(not(w:shd)) + number(ancestor::w:tr/w:trPr/w:cnfStyle/@w:lastRow ='1') + number(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr/@w:type ='lastRow'))&gt;=3">
        <xsl:variable name ="wfill">
          <xsl:value-of select ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='lastRow']/w:tcPr/w:shd/@w:fill"/>
        </xsl:variable>
        <xsl:variable name ="color">
          <xsl:value-of select ="concat('#',$wfill)"/>
        </xsl:variable>
        <字:填充_4134>
          <图:图案_800A>
            <xsl:attribute name ="背景色_800C">
              <xsl:value-of select ="$color"/>
            </xsl:attribute>
          </图:图案_800A>
        </字:填充_4134>
      </xsl:if>
      <xsl:if test ="(number(not(w:shd)) + number(ancestor::w:tr/w:trPr/w:cnfStyle/@w:oddVBand ='1') + number(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr/@w:type ='band1Vert'))&gt;=3">
        <xsl:variable name ="wfill">
          <xsl:value-of select ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='band1Vert']/w:tcPr/w:shd/@w:fill"/>
        </xsl:variable>
        <xsl:variable name ="color">
          <xsl:value-of select ="concat('#',$wfill)"/>
        </xsl:variable>
        <字:填充_4134>
          <图:图案_800A>
            <xsl:attribute name ="背景色_800C">
              <xsl:value-of select ="$color"/>
            </xsl:attribute>
          </图:图案_800A>
        </字:填充_4134>
      </xsl:if>
      <xsl:if test ="(number(not(w:shd)) + number(ancestor::w:tr/w:trPr/w:cnfStyle/@w:oddHBand ='1') + number(document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr/@w:type ='band1Horz'))&gt;=3">
        <xsl:variable name ="wfill">
          <xsl:value-of select ="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblStylePr[@w:type ='band1Horz']/w:tcPr/w:shd/@w:fill"/>
        </xsl:variable>
        <xsl:variable name ="color">
          <xsl:value-of select ="concat('#',$wfill)"/>
        </xsl:variable>
        <字:填充_4134>
          <图:图案_800A>
            <xsl:attribute name ="背景色_800C">
              <xsl:value-of select ="$color"/>
            </xsl:attribute>
          </图:图案_800A>
        </字:填充_4134>
      </xsl:if>
      <!--end-->

      <xsl:if test="w:vAlign">
        <字:垂直对齐方式_41A5>
          <xsl:choose>
            <xsl:when test="w:vAlign/@w:val='both'">
              <xsl:value-of select="'top'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="w:vAlign/@w:val"/>
            </xsl:otherwise>
          </xsl:choose>
        </字:垂直对齐方式_41A5>
      </xsl:if>

      <xsl:if test="w:gridSpan">
        <字:跨列_41A7>
          <xsl:value-of select="w:gridSpan/@w:val"/>
        </字:跨列_41A7>
      </xsl:if>

      <xsl:choose>
        <xsl:when test="w:noWrap">
          <字:是否自动换行_41A8>
        
              <xsl:choose>
                <xsl:when
                  test="w:noWrap/@w:val='off' or w:noWrap/@w:val='0' or w:noWrap/@w:val='false'">
                  <xsl:value-of select="'true'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'false'"/>
                </xsl:otherwise>
              </xsl:choose>
    
          </字:是否自动换行_41A8>
        </xsl:when>
        <xsl:otherwise>
          <字:是否自动换行_41A8>true</字:是否自动换行_41A8>
        </xsl:otherwise>
      </xsl:choose>

      <xsl:choose>
        <xsl:when test="w:tcFitText">
          <字:是否适应文字_41A9>
          
              <xsl:choose>
                <xsl:when
                  test="w:tcFitText/@w:val='off' or w:tcFitText/@w:val='0' or w:tcFitText/@w:val='false'">
                  <xsl:value-of select="'false'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'true'"/>
                </xsl:otherwise>
              </xsl:choose>
            
          </字:是否适应文字_41A9>
        </xsl:when>
        <xsl:otherwise>
          <字:是否适应文字_41A9>false</字:是否适应文字_41A9>
        </xsl:otherwise>
      </xsl:choose>

      <!--2014-03-10，wudi，增加文字表-单元格属性-文字排列方向的转换，start-->
      <xsl:choose>
        <xsl:when test="w:textDirection">
          <字:文字排列方向_41AA>
          <xsl:call-template name ="textDirection2">
            <xsl:with-param name ="dir" select ="./w:textDirection/@w:val"/>
          </xsl:call-template>
          </字:文字排列方向_41AA>
        </xsl:when>
      </xsl:choose>
      <!--end-->
      
    </字:单元格属性_41B7>
  </xsl:template>

  <!--2014-03-10，wudi，增加文字表-单元格属性-文字排列方向的转换，start-->
  <xsl:template name="textDirection2">
    <xsl:param name="dir"/>
    <xsl:choose>
      <xsl:when test="$dir='rlV'">
        <xsl:value-of select="'r2l-t2b-0e-90w'"/>
      </xsl:when>
      <xsl:when test="$dir='rl'">
        <xsl:value-of select="'r2l-t2b-90e-90w'"/>
      </xsl:when>
      <xsl:when test="$dir='tbV'">
        <xsl:value-of select="'t2b-l2r-270e-0w'"/>
      </xsl:when>
      <xsl:when test="$dir='lr'">
        <xsl:value-of select="'l2r-b2t-270e-270w'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'l2r-t2b-0e-0w'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--end-->

  <xsl:template match="w:shd" mode="tcShd">
    <字:填充_4134>
      <xsl:call-template name="shd"/>
    </字:填充_4134>
  </xsl:template>

  <xsl:template name="tblpPr">

    <字:绕排_41C5>around</字:绕排_41C5>

    <xsl:if
      test="w:tblpPr/@w:topFromText or w:tblpPr/@w:leftFromText or w:tblpPr/@w:rightFromText or w:tblpPr/@w:bottomFromText">
      <字:绕排边距_41C6>
        <xsl:if test="w:tblpPr/@w:topFromText">
          <xsl:attribute name="上_C609">
            <xsl:value-of select="(w:tblpPr/@w:topFromText) div 20"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="w:tblpPr/@w:leftFromText">
          <xsl:attribute name="左_C608">
            <xsl:value-of select="(w:tblpPr/@w:leftFromText) div 20"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="w:tblpPr/@w:rightFromText">
          <xsl:attribute name="右_C60A">
            <xsl:value-of select="(w:tblpPr/@w:rightFromText) div 20"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="w:tblpPr/@w:bottomFromText">
          <xsl:attribute name="下_C60B">
            <xsl:value-of select="(w:tblpPr/@w:bottomFromText) div 20"/>
          </xsl:attribute>
        </xsl:if>
      </字:绕排边距_41C6>
    </xsl:if>

    <xsl:if
      test="w:tblpPr/@w:horzAnchor or w:tblpPr/@w:tblpX or w:tblpPr/@w:tblpXSpec or w:tblpPr/@w:vertAnchor or w:tblpPr/@w:tblpY or w:tblpPr/@w:tblpYSpec">
      <字:位置_41C7>
        <uof:水平_4106>
          <xsl:if test="w:tblpPr/@w:horzAnchor or w:tblpPr/@w:tblpX or w:tblpPr/@w:tblpXSpec">
            <xsl:if test="w:tblpPr/@w:horzAnchor">
              <xsl:attribute name="相对于_410C">
                <xsl:choose>
                  <xsl:when test="w:tblpPr/@w:horzAnchor='margin'">
                    <xsl:value-of select="'margin'"/>
                  </xsl:when>
                  <xsl:when test="w:tblpPr/@w:horzAnchor='page'">
                    <xsl:value-of select="'page'"/>
                  </xsl:when>
                  <xsl:when test="w:tblpPr/@w:horzAnchor='text'">
                    <xsl:value-of select="'column'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(w:tblpPr/@w:horzAnchor)">
              <xsl:attribute name="相对于_410C">
                <xsl:value-of select="'column'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:choose>
              <xsl:when test="w:tblpPr/@w:tblpX">
                <uof:绝对_4107>
                  <xsl:attribute name="值_4108">
                    <xsl:value-of select="(w:tblpPr/@w:tblpX) div 20"/>
                  </xsl:attribute>
                </uof:绝对_4107>
              </xsl:when>
              <xsl:when test="w:tblpPr/@w:tblpXSpec and not(w:tblpPr/@w:tblpX)">
                <uof:相对_4109>
                  <xsl:attribute name="参考点_410A">
                    <xsl:choose>
                      <xsl:when test="w:tblpPr/@w:tblpXSpec='left'">
                        <xsl:value-of select="'left'"/>
                      </xsl:when>
                      <xsl:when test="w:tblpPr/@w:tblpXSpec='center'">
                        <xsl:value-of select="'center'"/>
                      </xsl:when>
                      <xsl:when test="w:tblpPr/@w:tblpXSpec='right'">
                        <xsl:value-of select="'right'"/>
                      </xsl:when>
                      <xsl:when test="w:tblpPr/@w:tblpXSpec='inside'">
                        <xsl:value-of select="'inside'"/>
                      </xsl:when>
                      <xsl:when test="w:tblpPr/@w:tblpXSpec='outside'">
                        <xsl:value-of select="'outside'"/>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:attribute>
                </uof:相对_4109>
              </xsl:when>
              <xsl:otherwise>
                <uof:相对_4109 参考点_410A="'left'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:if>
          <!--cxl,2012.3.7修改表格绕排水平位置，若默认状态下OOX中不出现这段代码-->
          <xsl:if test="not(w:tblpPr/@w:horzAnchor) and not(w:tblpPr/@w:tblpX) and not(w:tblpPr/@w:tblpXSpec)">
            
            <!--2014-05-27，wudi，column应该加引号，start-->
            <xsl:attribute name="相对于_410C">
              <xsl:value-of select="'column'"/>
            </xsl:attribute>
            <!--end-->
            
            <uof:相对_4109>
              <xsl:attribute name="参考点_410A">
                <xsl:value-of select="'left'"/>
              </xsl:attribute>
            </uof:相对_4109>
          </xsl:if>
        </uof:水平_4106>
        <uof:垂直_410D>
          <xsl:if test="w:tblpPr/@w:vertAnchor or w:tblpPr/@w:tblpY or w:tblpPr/@w:tblpYSpec">
            <xsl:if test="w:tblpPr/@w:vertAnchor">
              <xsl:attribute name="相对于_C647">
                <xsl:choose>
                  <xsl:when test="w:tblpPr/@w:vertAnchor='margin'">
                    <xsl:value-of select="'margin'"/>
                  </xsl:when>
                  <xsl:when test="w:tblpPr/@w:vertAnchor='page'">
                    <xsl:value-of select="'page'"/>
                  </xsl:when>
                  <xsl:when test="w:tblpPr/@w:vertAnchor='text'">
                    <xsl:value-of select="'paragraph'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="not(w:tblpPr/@w:vertAnchor)">
              <xsl:attribute name="相对于_C647">
                <xsl:value-of select="'margin'"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:choose>
              <xsl:when test="w:tblpPr/@w:tblpY">
                <uof:绝对_4107>
                  <xsl:attribute name="值_4108">
                    <xsl:value-of select="(w:tblpPr/@w:tblpY) div 20"/>
                  </xsl:attribute>
                </uof:绝对_4107>
              </xsl:when>
              <xsl:when test="w:tblpPr/@w:tblpYSpec and not(w:tblpPr/@w:tblpY)">
                <uof:相对_4109>
                  <xsl:attribute name="参考点_410B">
                    <xsl:choose>
                      <xsl:when test="w:tblpPr/@w:tblpYSpec='top'">
                        <xsl:value-of select="'top'"/>
                      </xsl:when>
                      <xsl:when test="w:tblpPr/@w:tblpYSpec='center'">
                        <xsl:value-of select="'center'"/>
                      </xsl:when>
                      <xsl:when test="w:tblpPr/@w:tblpYSpec='bottom'">
                        <xsl:value-of select="'bottom'"/>
                      </xsl:when>
                      <xsl:when test="w:tblpPr/@w:tblpYSpec='inside'">
                        <xsl:value-of select="'inside'"/>
                      </xsl:when>
                      <xsl:when test="w:tblpPr/@w:tblpYSpec='outside'">
                        <xsl:value-of select="'outside'"/>
                      </xsl:when>
                      <xsl:when test="w:tblpPr/@w:tblpYSpec='inline'">
                        <xsl:value-of select="'top'"/>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:attribute>
                </uof:相对_4109>
              </xsl:when>
              <xsl:otherwise>
                <uof:相对_4109 参考点_410B="top"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:if>
        </uof:垂直_410D>
      </字:位置_41C7>
    </xsl:if>
  </xsl:template>

  <xsl:template name="tblBorder"><!--存在BUG，内部横线与内部竖线的转换-->
    <xsl:variable name="tblId" select="./w:tblStyle/@w:val"/>
    <字:文字表边框_4227>
      
      <!--2013-11-14，wudi，Strict标准下，左边框对应start，start-->
      <xsl:if test="w:tblBorders/w:start">
        <xsl:apply-templates select="w:tblBorders/w:start"/>
      </xsl:if>
      <xsl:if test="not(w:tblBorders/w:start)">
        <xsl:apply-templates
          select="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:start"
        />
      </xsl:if>
      <!--end-->
      
      <xsl:if test="w:tblBorders/w:top">
        <xsl:apply-templates select="w:tblBorders/w:top"/>
      </xsl:if>
      <xsl:if test="not(w:tblBorders/w:top)">
        <xsl:apply-templates
          select="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:top"
        />
      </xsl:if>
      
      <!--2013-11-14，wudi，Strict标准下，右边框对应end，修正，start-->
      <xsl:if test="w:tblBorders/w:end">
        <xsl:apply-templates select="w:tblBorders/w:end"/>
      </xsl:if>
      <xsl:if test="not(w:tblBorders/w:end)">
        <xsl:apply-templates
          select="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:end"
        />
      </xsl:if>
      <!--end-->
      
      <xsl:if test="w:tblBorders/w:bottom">
        <xsl:apply-templates select="w:tblBorders/w:bottom"/>
      </xsl:if>
      <xsl:if test="not(w:tblBorders/w:bottom)">
        <xsl:apply-templates
          select="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:bottom"
        />
      </xsl:if>
      <!--cxl,2012.2.25增加表格内部横线与内部竖线转换代码-->
      <xsl:if test="w:tblBorders/w:insideH">
        <xsl:apply-templates select="w:tblBorders/w:insideH" mode="table"/>
      </xsl:if>
      <xsl:if test="not(w:tblBorders/w:insideH)">
        <xsl:apply-templates
          select="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideH"
         mode="table"/>
      </xsl:if>
      <xsl:if test="w:tblBorders/w:insideV">
        <xsl:apply-templates select="w:tblBorders/w:insideV" mode="table"/>
      </xsl:if>
      <xsl:if test="not(w:tblBorders/w:insideV)">
        <xsl:apply-templates
          select="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideV"
         mode="table"/>
      </xsl:if>    
    </字:文字表边框_4227>
  </xsl:template>
  <!--CXL,2012.2.25,表格内部横线与内部竖线-->
  <xsl:template match="w:insideH" mode="table">
    <uof:内部横线_C619>
      <xsl:call-template name="border"/>
    </uof:内部横线_C619>
  </xsl:template>
  <xsl:template match="w:insideV" mode="table">
    <uof:内部竖线_C61A>
      <xsl:call-template name="border"/>
    </uof:内部竖线_C61A>
  </xsl:template>
  
  <!--转换单元格边框 2011-11-13-->
  <!--2014-05-08，wudi，修改所有if条件中出现的'none'为'nil'，OOX中为nil，UOF中为none，start-->
  <!--2014-05-08，wudi，修改所有if条件中出现的'nil'为'nil'或'none'，OOX中可能为nil或none，UOF中为none-->
  <!--2014-05-19，wudi，修改所有if条件中出现的@w:val !='nil' and @w:val !='none'为@w:val !='none'，之前的修改会带来新的问题，start-->
  <xsl:template name="tcBorder">
    <xsl:param name="postr"/>
    <xsl:param name="postc"/>
    <xsl:param name="lasttr"/>
    <xsl:param name="lasttc"/>
    
    <xsl:variable name="tblId" select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
    <字:边框_4133>
      
      <!--2013-03-26，wudi，修复表格边框BUG，start-->
      
      <!--2013-11-14，wudi，Strict标准下，左边框对应start，修正，start-->
      <xsl:if test="w:tcBorders/w:start[@w:val !='none']">
        <xsl:apply-templates select="w:tcBorders/w:start"/>
      </xsl:if>
      <xsl:if test="not(w:tcBorders/w:start[@w:val !='none'])">
        <xsl:choose>
          <xsl:when test="(ancestor::w:tbl/w:tblPr/w:tblBorders/w:start) and ($postc='1')">
            <xsl:apply-templates select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:start"/>
          </xsl:when>
          <xsl:when test="(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideV) and ($postc!='1')">
            <xsl:apply-templates select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideV" mode="left" />
          </xsl:when>
          
          <!--2013-04-25，wudi，修复OOX到UOF方向表格边框转换BUG，针对document.xml某些表格属性里没有边框信息，而引用的styles.xml里有边框信息的情况，start-->
          <xsl:when test ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideV and ($postc!='1')">
            <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideV" mode ="left"/>
          </xsl:when>
          
          <xsl:when test ="not(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideV) and not(document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideV) and ($postc!='1')">
            <uof:左_C613>
              <xsl:attribute name ="线型_C60D">
                <xsl:value-of select ="'none'"/>
              </xsl:attribute>
              <xsl:attribute name ="虚实_C60E">
                <xsl:value-of select ="'none'"/>
              </xsl:attribute>
              <xsl:attribute name ="宽度_C60F">
                <xsl:value-of select ="'0.0'"/>
              </xsl:attribute>
              <xsl:attribute name ="边距_C610">
                <xsl:value-of select ="'0.0'"/>
              </xsl:attribute>
              <xsl:attribute name ="颜色_C611">
                <xsl:value-of select ="'auto'"/>
              </xsl:attribute>
            </uof:左_C613>
          </xsl:when>
          <!--end-->
          
          <xsl:otherwise>
            <xsl:apply-templates
              select="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:start"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
      <!--end-->
      
      <xsl:if test="w:tcBorders/w:top[@w:val !='none']">
        <xsl:apply-templates select="w:tcBorders/w:top"/>
      </xsl:if>
      <xsl:if test="not(w:tcBorders/w:top[@w:val !='none'])">
        <xsl:choose>
          <xsl:when test="(ancestor::w:tbl/w:tblPr/w:tblBorders/w:top) and ($postr='1')">
            <xsl:apply-templates select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:top"/>
          </xsl:when>
          <xsl:when test="(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideH) and ($postr!='1')">
            <xsl:apply-templates select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideH" mode="top"/>
          </xsl:when>
          
          <!--2013-04-25，wudi，修复OOX到UOF方向表格边框转换BUG，针对document.xml某些表格属性里没有边框信息，而引用的styles.xml里有边框信息的情况，start-->
          <xsl:when test ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideH and ($postc!='1')">
            <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideH" mode ="top"/>
          </xsl:when>

          <!--2014-05-08，wudi，修复表格边框转换BUG，首行边框存在丢失的问题，start-->
          <xsl:when test ="not(ancestor::w:tbl/w:tblPr/w:tblBorders/w:top) and not(document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:top) and (ancestor::w:tr/w:trPr/w:cnfStyle/@w:firstRow = '1') and ($postr='1')">
            <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblStylePr[@w:type='firstRow']/w:tcPr/w:tcBorders/w:top"/>
          </xsl:when>
          <!--end-->

          <!--2014-05-09，wudi，修复表格边框转换BUG，内部边框存在丢失的问题，start-->
          <xsl:when test="not(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideH) and not(document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideH) and (ancestor::w:tr/w:trPr/w:cnfStyle/@w:oddHBand = '1') and ($postr!='1')">
            <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblStylePr[@w:type='band1Horz']/w:tcPr/w:tcBorders/w:top"/>
          </xsl:when>

          <xsl:when test="not(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideH) and not(document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideH) and (ancestor::w:tr/w:trPr/w:cnfStyle/@w:evenHBand = '1') and ($postr!='1')">
            <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblStylePr[@w:type='band2Horz']/w:tcPr/w:tcBorders/w:top"/>
          </xsl:when>
          <!--end-->
          
          <xsl:when test ="not(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideH) and not(document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideH) and not(ancestor::w:tr/w:trPr/w:cnfStyle/@w:oddHBand = '1') and not(ancestor::w:tr/w:trPr/w:cnfStyle/@w:evenHBand = '1') and ($postr!='1')">
            <uof:上_C614>
              <xsl:attribute name ="线型_C60D">
                <xsl:value-of select ="'none'"/>
              </xsl:attribute>
              <xsl:attribute name ="虚实_C60E">
                <xsl:value-of select ="'none'"/>
              </xsl:attribute>
              <xsl:attribute name ="宽度_C60F">
                <xsl:value-of select ="'0.0'"/>
              </xsl:attribute>
              <xsl:attribute name ="边距_C610">
                <xsl:value-of select ="'0.0'"/>
              </xsl:attribute>
              <xsl:attribute name ="颜色_C611">
                <xsl:value-of select ="'auto'"/>
              </xsl:attribute>
            </uof:上_C614>
          </xsl:when>
          <!--end-->
          
          <xsl:otherwise>
            <xsl:apply-templates
              select="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:top"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
      
      <!--2013-11-14，wudi，Strict标准下，右边框对应end，修正，start-->
      <xsl:if test="w:tcBorders/w:end[@w:val !='none']">
        <xsl:apply-templates select="w:tcBorders/w:end"/>
      </xsl:if>
      <xsl:if test="not(w:tcBorders/w:end[@w:val !='none'])">
        <xsl:choose>
          <xsl:when test="(ancestor::w:tbl/w:tblPr/w:tblBorders/w:end) and ($postc=$lasttc)">
            <xsl:apply-templates select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:end"/>
          </xsl:when>
          <xsl:when test="(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideV) and ($postc!=$lasttc)">
            <xsl:apply-templates select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideV" mode="right"/>
          </xsl:when>

          <!--2013-04-25，wudi，修复OOX到UOF方向表格边框转换BUG，针对document.xml某些表格属性里没有边框信息，而引用的styles.xml里有边框信息的情况，start-->
          <xsl:when test ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideV and ($postc!='1')">
            <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideV" mode ="right"/>
          </xsl:when>
          
          <xsl:when test ="not(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideV) and not(document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideV) and ($postc!=$lasttc)">
            <uof:右_C615>
              <xsl:attribute name ="线型_C60D">
                <xsl:value-of select ="'none'"/>
              </xsl:attribute>
              <xsl:attribute name ="虚实_C60E">
                <xsl:value-of select ="'none'"/>
              </xsl:attribute>
              <xsl:attribute name ="宽度_C60F">
                <xsl:value-of select ="'0.0'"/>
              </xsl:attribute>
              <xsl:attribute name ="边距_C610">
                <xsl:value-of select ="'0.0'"/>
              </xsl:attribute>
              <xsl:attribute name ="颜色_C611">
                <xsl:value-of select ="'auto'"/>
              </xsl:attribute>
            </uof:右_C615>
          </xsl:when>
          <!--end-->
          
          <xsl:otherwise>
            <xsl:apply-templates
              select="document('word/styles.xml')/w:styles/w:style[@w:type='table'and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:end"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
      <!--end-->
      
      <xsl:if test="w:tcBorders/w:bottom[@w:val !='none']">
        <xsl:apply-templates select="w:tcBorders/w:bottom"/>
      </xsl:if>
      <xsl:if test="not(w:tcBorders/w:bottom[@w:val !='none'])">
        <xsl:choose>
          <xsl:when test="(ancestor::w:tbl/w:tblPr/w:tblBorders/w:bottom) and ($postr=$lasttr)">
            <xsl:apply-templates select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:bottom"/>
          </xsl:when>
          <xsl:when test="(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideH) and ($postr!=$lasttr)">
            <xsl:apply-templates select="ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideH" mode="bottom"/>
          </xsl:when>

          <!--2013-04-25，wudi，修复OOX到UOF方向表格边框转换BUG，针对document.xml某些表格属性里没有边框信息，而引用的styles.xml里有边框信息的情况，start-->
          <xsl:when test ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideH and ($postc!='1')">
            <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideH" mode ="bottom"/>
          </xsl:when>

          <!--2014-05-08，wudi，修复表格边框转换BUG，首行边框存在丢失的问题，start-->
          <xsl:when test ="not(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideH) and not(document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideH) and (ancestor::w:tr/w:trPr/w:cnfStyle/@w:firstRow = '1') and ($postr='1')">
            <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblStylePr[@w:type='firstRow']/w:tcPr/w:tcBorders/w:bottom"/>
          </xsl:when>
          <!--end-->

          <!--2014-05-09，wudi，修复表格边框转换BUG，内部边框存在丢失的问题，start-->
          <xsl:when test="not(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideH) and not(document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideH) and (ancestor::w:tr/w:trPr/w:cnfStyle/@w:oddHBand = '1') and ($postr!='1')">
            <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblStylePr[@w:type='band1Horz']/w:tcPr/w:tcBorders/w:bottom"/>
          </xsl:when>

          <xsl:when test="not(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideH) and not(document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideH) and (ancestor::w:tr/w:trPr/w:cnfStyle/@w:evenHBand = '1') and ($postr!='1')">
            <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblStylePr[@w:type='band2Horz']/w:tcPr/w:tcBorders/w:bottom"/>
          </xsl:when>
          <!--end-->
          
          <xsl:when test ="not(ancestor::w:tbl/w:tblPr/w:tblBorders/w:insideH) and not(document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:insideH) and ($postr!=$lasttr)">
            <uof:下_C616>
              <xsl:attribute name ="线型_C60D">
                <xsl:value-of select ="'none'"/>
              </xsl:attribute>
              <xsl:attribute name ="虚实_C60E">
                <xsl:value-of select ="'none'"/>
              </xsl:attribute>
              <xsl:attribute name ="宽度_C60F">
                <xsl:value-of select ="'0.0'"/>
              </xsl:attribute>
              <xsl:attribute name ="边距_C610">
                <xsl:value-of select ="'0.0'"/>
              </xsl:attribute>
              <xsl:attribute name ="颜色_C611">
                <xsl:value-of select ="'auto'"/>
              </xsl:attribute>
            </uof:下_C616>
          </xsl:when>
          <!--end-->
          
          <xsl:otherwise>
            <xsl:apply-templates
              select="document('word/styles.xml')/w:styles/w:style[@w:type='table'and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:bottom"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
      <!--end-->
      
      <xsl:if test="w:tcBorders/w:tl2br">
        <xsl:apply-templates select="w:tcBorders/w:tl2br"/>
      </xsl:if>
      <xsl:if test="w:tcBorders/w:tr2bl">
        <xsl:apply-templates select="w:tcBorders/w:tr2bl"/>
      </xsl:if>
    </字:边框_4133>
  </xsl:template>
  <!--end-->
  <!--end-->
  <!--end-->
  <!--yx,09.12.11-->
  

  <xsl:template match="w:start">
    <uof:左_C613>
      <xsl:call-template name="border"/>
    </uof:左_C613>
  </xsl:template>
  <xsl:template match="w:top">
    <uof:上_C614>
      <xsl:call-template name="border"/>
    </uof:上_C614>
  </xsl:template>
  <xsl:template match="w:end">
    <uof:右_C615>
      <xsl:call-template name="border"/>
    </uof:右_C615>
  </xsl:template>
  <xsl:template match="w:bottom">
    <uof:下_C616>
      <xsl:call-template name="border"/>
    </uof:下_C616>
  </xsl:template>
  <xsl:template match="w:tl2br">
    <uof:对角线1_C617>
      <xsl:call-template name="border"/>
    </uof:对角线1_C617>
  </xsl:template>
  <xsl:template match="w:tr2bl">
    <uof:对角线2_C618>
      <xsl:call-template name="border"/>
    </uof:对角线2_C618>
  </xsl:template>
  <xsl:template match="w:insideV" mode="left">
    <uof:左_C613>
      <xsl:call-template name="border"/>
    </uof:左_C613>
  </xsl:template>
  <xsl:template match="w:insideV" mode="right">
    <uof:右_C615>
      <xsl:call-template name="border"/>
    </uof:右_C615>
  </xsl:template>
  <xsl:template match="w:insideH" mode="top">
    <uof:上_C614>
      <xsl:call-template name="border"/>
    </uof:上_C614>
  </xsl:template>
  <xsl:template match="w:insideH" mode="bottom">
    <uof:下_C616>
      <xsl:call-template name="border"/>
    </uof:下_C616>
  </xsl:template>

  <xsl:template name="border">
    <!--修改线型虚实的转换-->
    <xsl:choose>
      <xsl:when test="./@w:val='nil' or ./@w:val='none'">
        <xsl:attribute name="线型_C60D">none</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='single'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='dotted'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">square-dot</xsl:attribute>
      </xsl:when>
      
      <!--2013-05-28，修复表格边框转换BUG，UOF2.0标准支持此边框转换，start-->
      <xsl:when test="./@w:val='dashSmallGap'">
        <xsl:attribute name="线型_C60D">thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">round-dot</xsl:attribute>
      </xsl:when>
      <!--end-->
      
      <xsl:when test="./@w:val='dashed'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='dotDash'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='dotDotDash'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash-dot-dot</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='double'">
        <xsl:attribute name="线型_C60D">double</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='triple'">
        <xsl:attribute name="线型_C60D">double</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash-dot-dot</xsl:attribute>
      </xsl:when>
      
      <!--2014-01-08，wudi，表格边框，粗细线，细粗线转换有误，修正，start-->
      <xsl:when test="./@w:val='thinThickSmallGap'">
        <xsl:attribute name="线型_C60D">thick-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='thickThinSmallGap'">
        <xsl:attribute name="线型_C60D">thin-thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='thinThickMediumGap'">
        <xsl:attribute name="线型_C60D">thick-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='thickThinMediumGap'">
        <xsl:attribute name="线型_C60D">thin-thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='thickThinLargeGap'">
        <xsl:attribute name="线型_C60D">thin-thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='thinThickLargeGap'">
        <xsl:attribute name="线型_C60D">thick-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      <!--end-->
      
      <xsl:when test="./@w:val='thinThickThinSmallGap'">
        <xsl:attribute name="线型_C60D">thick-between-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='thinThickThinMediumGap'">
        <xsl:attribute name="线型_C60D">thick-between-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='thinThickThinLargeGap'">
        <xsl:attribute name="线型_C60D">thick-between-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      
      <!--2013-03-28，修复表格边框BUG，start-->
      <xsl:when test="./@w:val='wave'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">wave</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='doubleWave'">
        <xsl:attribute name="线型_C60D">double</xsl:attribute>
        <xsl:attribute name="虚实_C60E">wave</xsl:attribute>
      </xsl:when>
      <!--end-->
      
      <!--2013-05-28，修复表格边框BUG，没有按差异文档转换，start-->
      <xsl:when test="./@w:val='dashDotStroked'">
        <xsl:attribute name ="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash-dot</xsl:attribute>
      </xsl:when>
      <!--end-->
      
      <xsl:otherwise>
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>

    <xsl:if test="./@w:color">
      <xsl:attribute name="颜色_C611">
        <xsl:if test="./@w:color!='auto'">
          <xsl:value-of select="concat('#',./@w:color)"/>
        </xsl:if>
        <xsl:if test="./@w:color='auto'">
          <xsl:value-of select="'auto'"/>
        </xsl:if>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="./@w:sz">
      <xsl:attribute name="宽度_C60F">
        <xsl:value-of select="(./@w:sz) div 8"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="./@w:space">
      <xsl:attribute name="边距_C610">
        <xsl:value-of select="./@w:space"/>
      </xsl:attribute>
    </xsl:if>
    <!--  
    <xsl:if test="./@w:shadow">
      <xsl:attribute name="uof:阴影">
        <xsl:if test="./@w:shadow='on' or ./@w:shadow='1' or ./@w:shadow='true'">
          <xsl:value-of select="'true'"/>
        </xsl:if>
        <xsl:if test="./@w:shadow='off' or ./@w:shadow='0' or ./@w:shadow='false'">
          <xsl:value-of select="'false'"/>
        </xsl:if>
      </xsl:attribute>
    </xsl:if> -->
  </xsl:template>
   
  <xsl:template name="Cellmar">
    <xsl:param name="Cellmar"/>
    <xsl:if test="$Cellmar/w:top[@w:type!='auto']">
      <xsl:attribute name="上_C609">
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="$Cellmar/w:top"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <!--2014-03-10，wudi，单元格属性不全，继承表格式样，start-->
    <xsl:if test="not($Cellmar/w:top)">
      <xsl:variable name="styId">
        <xsl:value-of select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
      </xsl:variable>
      <xsl:attribute name="上_C609">
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblPr/w:tblCellMar/w:top"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <!--end-->
    
    <!--2013-11-14，wudi，Strict标准下左右边框分别是start和end，Transtional标准下左右边框分别是left和right，修正，start-->
    <xsl:if test="$Cellmar/w:start[@w:type!='auto']">
      <xsl:attribute name="左_C608">
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="$Cellmar/w:start"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <!--2014-03-10，wudi，单元格属性不全，继承表格式样，start-->
    <xsl:if test="not($Cellmar/w:start)">
      <xsl:variable name="styId">
        <xsl:value-of select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
      </xsl:variable>
      <xsl:attribute name="左_C608">
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblPr/w:tblCellMar/w:start"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <!--end-->
    
    <xsl:if test="$Cellmar/w:end[@w:type!='auto']">
      <xsl:attribute name="右_C60A">
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="$Cellmar/w:end"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <!--2014-03-10，wudi，单元格属性不全，继承表格式样，start-->
    <xsl:if test="not($Cellmar/w:end)">
      <xsl:variable name="styId">
        <xsl:value-of select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
      </xsl:variable>
      <xsl:attribute name="右_C60A">
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblPr/w:tblCellMar/w:end"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <!--end-->
    
    <!--end-->
    
    <xsl:if test="$Cellmar/w:bottom[@w:type!='auto']">
      <xsl:attribute name="下_C60B">
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="$Cellmar/w:bottom"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <!--2014-03-10，wudi，单元格属性不全，继承表格式样，start-->
    <xsl:if test="not($Cellmar/w:bottom)">
      <xsl:variable name="styId">
        <xsl:value-of select="ancestor::w:tbl/w:tblPr/w:tblStyle/@w:val"/>
      </xsl:variable>
      <xsl:attribute name="下_C60B">
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="document('word/styles.xml')/w:styles/w:style[@w:styleId =$styId]/w:tblPr/w:tblCellMar/w:bottom"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <!--end-->
    
  </xsl:template>

  <xsl:template name="cellmar">
    <xsl:param name="cellmar"/>

    <!--2013-10-30,wudi,w:w在strict和transitonal标准下取值不同，加以区分，start-->
    <xsl:if test="$cellmar/@w:type='dxa' and $cellmar/@w:w">
      <xsl:if test ="not(contains($cellmar/@w:w,'pt'))">
        <xsl:value-of select="($cellmar/@w:w) div 20"/>
      </xsl:if>
      <xsl:if test ="contains($cellmar/@w:w,'pt')">
        <xsl:value-of select="substring-before($cellmar/@w:w,'pt')"/>
      </xsl:if>
    </xsl:if>
    <!--end-->
    
    <xsl:if test="$cellmar/@w:type='nil'or not($cellmar/@w:w)">
      <xsl:value-of select="'0'"/>
    </xsl:if>

    <!--2013-10-30,wudi,w:w在strict和transitonal标准下取值不同，加以区分，start-->
    <xsl:if test="$cellmar/@w:type='pct' and $cellmar/@w:w">
      <xsl:if test ="not(contains($cellmar/@w:w,'pt'))">
        <xsl:value-of select="($cellmar/@w:w) * 0.08522"/>
      </xsl:if>
      <xsl:if test ="contains($cellmar/@w:w,'pt')">
        <xsl:value-of select="substring-before($cellmar/@w:w,'pt')"/>
      </xsl:if>
    </xsl:if>
    <!--end-->
    
  </xsl:template>

  <xsl:template name="IdProducer">
    <xsl:param name="ooxId"/>
    <xsl:choose>
      <xsl:when test="starts-with($ooxId,'id_')">
        <xsl:value-of select="$ooxId"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="concat('id_',$ooxId)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>
