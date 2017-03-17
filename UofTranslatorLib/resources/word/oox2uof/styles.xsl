<?xml version="1.0" encoding="UTF-8"?>
<!--
* Copyright (c) 2006, BeiHang University, China
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
<Author>Ban Qianchao(BUAA)</Author>
<Author>Li Jingui</Author>
-->
<xsl:stylesheet version="2.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
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
  xmlns:式样="http://schemas.uof.org/cn/2009/styles">
  <xsl:import href="numbering.xsl"/>
  <!--yx,why not add import,because the template: <xsl:apply-templates select="./w:pPr"/>in that xsl,2010.3.3-->
  <xsl:output method="xml" version="1.1" encoding="UTF-8" indent="yes"/>

  <xsl:variable name="docDefaults" select="document('word/styles.xml')/w:styles/w:docDefaults"/>
  <xsl:variable name="tableStyle" select="document('word/styles.xml')/w:styles/w:style[@w:type='table']"/>

  <xsl:key name="StyleId" match="w:style" use="@w:styleId"/>
  <xsl:key name="default-styles" match="w:style[@w:default = 1 or @w:default = 'true' or @w:default = 'on']" use="@w:type"/>


  <xsl:template name="styles">
    <式样:式样集_990B>
      <!--式样集:字体集，2011-02-18-->
      <xsl:call-template name="fontTable"/>
      
      <!--式样集:编号集，2011-02-19-->
      <!--template定义在numbering.xsl-->
      
      <xsl:call-template name="numberingSet"/>

      <xsl:for-each select="document('word/styles.xml')">
        <式样:句式样集_990F>
          <xsl:call-template name ="InsertDefaultTextStyle"/><!--句式样集-->
          <xsl:apply-templates select="w:styles/w:style[@w:type= 'character']" mode ="characterStyle"/>
        </式样:句式样集_990F>
        <式样:段落式样集_9911>
          <xsl:call-template name ="InsertDefaultParagraphStyle"/>
          <xsl:apply-templates select="w:styles/w:style[@w:type= 'paragraph']" mode ="paragraphStyle"/>
        </式样:段落式样集_9911>
        <式样:文字表式样集_9917>
          <xsl:apply-templates select="$tableStyle" mode ="tableStyle"/>
        </式样:文字表式样集_9917>
      </xsl:for-each>
    </式样:式样集_990B>
  </xsl:template>

  <!--字体集 2011-02-18-->
  <xsl:template name="fontTable">
    <式样:字体集_990C>
      <xsl:for-each select="document('word/fontTable.xml')/w:fonts/w:font">
        <!--<xsl:choose>
          <xsl:when test="name(.)='w:font'">-->
            <式样:字体声明_990D><!--属性与元素有变化@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
              <xsl:attribute name="标识符_9902">
                <xsl:value-of select="translate(@w:name,' ','')"/>
              </xsl:attribute>
              <xsl:attribute name="名称_9903">
                <xsl:value-of select="@w:name"/>
              </xsl:attribute>
              <!--字体族属性暂不转换 2011-02-18 L-->
              <!--xsl:attribute name="字体族">
                <xsl:value-of select="@w:name"/>
              </xsl:attribute-->
            </式样:字体声明_990D>
          <!--</xsl:when>
        </xsl:choose>-->
      </xsl:for-each>
    </式样:字体集_990C>
  </xsl:template>
  
  
  <!--cxl,current doc:word/styles.xml,2011.11.10-->
  <!--段落式样集模板*********************************************************************************************************-->
  <xsl:template name="InsertDefaultParagraphStyle">
      <xsl:if
        test="w:styles/w:docDefaults[w:pPrDefault or w:rPrDefault] or key('default-styles', 'paragraph')">
        <xsl:for-each select="key('default-styles', 'paragraph')[last()]">
          <式样:段落式样_9912><!--attrList="标识符 名称 类型 别名 基式样引用 后继式样引用"--> 
            <xsl:call-template name ="InsertStyleAttr"/>

            <字:句属性_4158>
              <xsl:choose>
                <xsl:when test ="./w:rPr/node()">
                  <xsl:apply-templates select ="./w:rPr" mode="RunProperties"/>
                  <xsl:for-each select="document('word/styles.xml')/w:styles/w:docDefaults/w:rPrDefault/w:rPr/child::node()">
                    <xsl:variable name="elementName" select="name()"/>
                    <xsl:if
                      test="not(key('default-styles', 'paragraph')[last()]/w:rPr/*[name() = $elementName])">
                      <xsl:apply-templates select ="w:rStyle|w:rFonts|w:color|w:sz|w:b|w:i
                                       |w:caps|w:smallCaps|w:strike|w:dstrike
                                       |w:outline|w:shadow|w:emboss|w:imprint
                                       |w:imprint|w:snapToGrid|w:vanish|w:spacing
                                       |w:w|w:kern|w:position|w:highlight|w:u|w:bdr
                                       |w:shd|w:vertAlign|w:em" mode="rPrChildren"/>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:docDefaults/w:rPrDefault/w:rPr" mode="RunProperties"/>
                </xsl:otherwise>
              </xsl:choose>
            </字:句属性_4158>

            <xsl:choose>
              <xsl:when test ="./w:pPr/node()">
                <xsl:apply-templates select ="./w:pPr"/>
                <xsl:for-each select="document('word/styles.xml')/w:styles/w:docDefaults/w:pPrDefault/w:pPr/child::node()">
                  <xsl:variable name="elementName" select="name()"/>                
                  <xsl:if
                    test="not(key('default-styles', 'paragraph')[last()]/w:pPr/*[name() = $elementName])">                 
                    <xsl:apply-templates select="." mode="pPrChildren"/>
                  
                  </xsl:if>
                </xsl:for-each>
              </xsl:when>
              <xsl:otherwise>
                <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:docDefaults/w:pPrDefault/w:pPr"/>
              </xsl:otherwise>
            </xsl:choose>
          </式样:段落式样_9912>
        </xsl:for-each>
      </xsl:if>
    </xsl:template>
  <!--句式样集转换*********************************************************************************************************-->
  <xsl:template name="InsertDefaultTextStyle">   
      <xsl:if
      test="w:styles/w:docDefaults[w:rPrDefault] or key('default-styles', 'character') or key('default-styles', 'paragraph')/w:rPr">
        <xsl:for-each select="key('default-styles', 'character')[last()]">
          <式样:句式样_9910><!--attrList="标识符 名称 类型 别名 基式样引用 后继式样引用"-->
            <xsl:call-template name ="InsertStyleAttr"/>

            <xsl:choose>
              <xsl:when test ="./w:rPr/node()">
                <xsl:apply-templates select ="./w:rPr" mode="RunProperties"/>
                <xsl:for-each select="document('word/styles.xml')/w:styles/w:docDefaults/w:rPrDefault/w:rPr/child::node()">
                  <xsl:variable name="elementName" select="name()"/>
                  <xsl:if
                    test="not(key('default-styles', 'paragraph')[last()]/w:rPr/*[name() = $elementName])">
                    <xsl:apply-templates select="." mode="rPrChildren"/>
                  </xsl:if>
                </xsl:for-each>
              </xsl:when>
              <xsl:otherwise>
                <xsl:apply-templates select ="document('word/styles.xml')/w:styles/w:docDefaults/w:rPrDefault/w:rPr" mode="RunProperties"/>
              </xsl:otherwise>
            </xsl:choose>
          </式样:句式样_9910>
        </xsl:for-each>
      </xsl:if>
  </xsl:template>

  <xsl:template match ="w:style" mode="characterStyle">
    <xsl:if test ="not(@w:default = 1 or @w:default = 'true' or @w:default = 'on')">
      <式样:句式样_9910>
        <xsl:call-template name ="InsertStyleAttr"/>
        <xsl:apply-templates select ="./w:rPr" mode="RunProperties"/>
      </式样:句式样_9910>
    </xsl:if>
  </xsl:template>
  
  <!--<xsl:apply-templates select="w:styles/w:style[@w:type= 'paragraph']" mode ="paragraphStyle"/>-->
  <xsl:template match ="w:style" mode="paragraphStyle">
    <xsl:if test ="not(@w:default = 1 or @w:default = 'true' or @w:default = 'on')">
      <式样:段落式样_9912>
        <xsl:call-template name ="InsertStyleAttr"/>
        <xsl:apply-templates select="./w:pPr"/>
        <xsl:if test="w:rPr">
          <字:句属性_4158>
            <xsl:apply-templates select ="./w:rPr" mode="RunProperties"/>
          </字:句属性_4158>
        </xsl:if>
      </式样:段落式样_9912>
    </xsl:if>
  </xsl:template>
  <!--下面这几个模板是匹配啥的？？？？？？？？？？？？？？？？？？@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
  <xsl:template match="w:style" mode="tableStylePr">
    <式样:段落式样_9912>
      <xsl:attribute name="标识符_4100">
        <xsl:value-of select="concat('id_',@w:styleId,'Paragraph')"/>
      </xsl:attribute>
      <xsl:attribute name="名称_4101">
        <xsl:value-of select="concat(@w:styleId,'Paragraph')"/>
      </xsl:attribute>
      <xsl:attribute name="类型_4102">
        <xsl:value-of select="'custom'"/>
      </xsl:attribute>
      <xsl:attribute name="基式样引用_4104">
        <xsl:variable name ="styleId" select="document('word/styles.xml')/w:styles/w:style[@w:type='paragraph' and @w:default='1']/@w:styleId"/>
        <xsl:value-of select="concat('id_',$styleId)"/>
      </xsl:attribute>
      <xsl:attribute name="后继式样引用_4105">
        <xsl:value-of select="concat('id_',@w:styleId,'Paragraph')"/>
      </xsl:attribute>

      <xsl:apply-templates select="./w:pPr"/>
      <xsl:if test="w:rPr">
        <字:句属性_4158>
          <xsl:apply-templates select ="./w:rPr" mode="RunProperties"/>
        </字:句属性_4158>
      </xsl:if>
    </式样:段落式样_9912>
  </xsl:template>

  <xsl:template match="w:style" mode="tableParagraphStyle">
    <xsl:variable name="styleId" select="@w:styleId"/>

    <xsl:for-each select="./w:tblStylePr">
      <式样:段落式样_9912>
        <xsl:attribute name="标识符_4100">
          <xsl:value-of select="concat('id_',$styleId,'_',@w:type)"/>
        </xsl:attribute>
        <xsl:attribute name="名称_4101">
          <xsl:value-of select="concat($styleId,@w:type)"/>
        </xsl:attribute>
        <xsl:attribute name="类型_4102">
          <xsl:value-of select="'custom'"/>
        </xsl:attribute>
        <xsl:attribute name="基式样引用_4104">
          <xsl:value-of select="concat('id_',$styleId,'Paragraph')"/>
        </xsl:attribute>
        <xsl:attribute name="后继式样引用_4105">
          <xsl:value-of select="concat('id_',$styleId,'_',@w:type)"/>
        </xsl:attribute>

        <xsl:apply-templates select="./w:pPr"/>
        <xsl:if test="w:rPr">
          <字:句属性_4158>
            <xsl:apply-templates select ="./w:rPr" mode="RunProperties"/>
          </字:句属性_4158>
        </xsl:if>
      </式样:段落式样_9912>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match ="w:style" mode="tableStyle">    
      <式样:文字表式样_9918>
        <xsl:call-template name ="InsertStyleAttr"/>
        <xsl:for-each select="node()">
          <xsl:if test="name(.)='w:tblPr'">
            <xsl:call-template name="tblPr"/>
          </xsl:if>
          <xsl:if test ="w:trPr|w:tcPr|w:tblStylePr">
            <xsl:message terminate="no">feedback:lost:Table_Style:table style</xsl:message>
          </xsl:if>
        </xsl:for-each>
      </式样:文字表式样_9918>
  </xsl:template>
  <!--文字表属性(与table.xsl中重复)FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF-->
  <!--<xsl:template name="tblPr">

    <xsl:if test="w:tblW[@w:type!='auto']">
      <字:宽度_41A1>

        <xsl:if test="w:tblW/@w:type='dxa'">
          <xsl:attribute name="绝对宽度_41BF">
            <xsl:if test="w:tblW/@w:w">
              <xsl:value-of select="(w:tblW/@w:w) div 20"/>
            </xsl:if>
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
            <xsl:if test="w:tblW/@w:w">
              <xsl:value-of select="(w:tblW/@w:w) div 50"/>
            </xsl:if>
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
          <xsl:when test="w:jc/@w:val='left'">
            <xsl:value-of select="'left'"/>
          </xsl:when>
          <xsl:when test="w:jc/@w:val='right'">
            <xsl:value-of select="'right'"/>
          </xsl:when>
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

        <xsl:if test="w:tblInd/@w:type='dxa' and w:tblInd/@w:w">
          <xsl:value-of select="(w:tblInd/@w:w) div 20"/>
        </xsl:if>

        <xsl:if test="w:tblInd/@w:type='nil' or not(w:tblInd/@w:w)">
          <xsl:value-of select="'0'"/>
        </xsl:if>

        <xsl:if test="w:tblInd/@w:type='pct' and w:tblInd/@w:w">
          <xsl:value-of select="(w:tblInd/@w:w) * 0.08522"/>
        </xsl:if>

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

        <xsl:if test="w:tblCellSpacing/@w:type='dxa' and w:tblCellSpacing/@w:w">
          <xsl:value-of select="(w:tblCellSpacing/@w:w) div 10"/>
        </xsl:if>

        <xsl:if test="w:tblCellSpacing/@w:type='nil'or not(w:tblCellSpacing/@w:w)">
          <xsl:value-of select="'0'"/>
        </xsl:if>

        <xsl:if test="w:tblCellSpacing/@w:type='pct' and w:tblCellSpacing/@w:w">
          <xsl:value-of select="(w:tblCellSpacing/@w:w) * 0.08522"/>
        </xsl:if>
      </字:默认单元格间距_41CB>
    </xsl:if>
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
        <xsl:if test="w:tblpPr/@w:horzAnchor or w:tblpPr/@w:tblpX or w:tblpPr/@w:tblpXSpec">
          <uof:水平_4106>
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
                        <xsl:value-of select="'left'"/>
                      </xsl:when>
                      <xsl:when test="w:tblpPr/@w:tblpXSpec='outside'">
                        <xsl:value-of select="'right'"/>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:attribute>
                </uof:相对_4109>
              </xsl:when>

              <xsl:otherwise>
                <uof:相对_4109 参考点_410A="left"/>
              </xsl:otherwise>
            </xsl:choose>
          </uof:水平_4106>
        </xsl:if>
        <xsl:if test="w:tblpPr/@w:vertAnchor or w:tblpPr/@w:tblpY or w:tblpPr/@w:tblpYSpec">
          <uof:垂直_410D>
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
                        <xsl:value-of select="'top'"/>
                      </xsl:when>
                      <xsl:when test="w:tblpPr/@w:tblpYSpec='outside'">
                        <xsl:value-of select="'bottom'"/>
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
          </uof:垂直_410D>
        </xsl:if>
      </字:位置_41C7>  
  </xsl:if>
  </xsl:template>-->
  <!--<xsl:template name="tblBorder">
    <xsl:variable name="tblId" select="./w:tblStyle/@w:val"/>
    <字:文字表边框_4227>
      <xsl:if test="w:tblBorders/w:left">
        <xsl:apply-templates select="w:tblBorders/w:left"/>
      </xsl:if>
      <xsl:if test="not(w:tblBorders/w:left)">
        <xsl:apply-templates
          select="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:left"
        />
      </xsl:if>
      <xsl:if test="w:tblBorders/w:top">
        <xsl:apply-templates select="w:tblBorders/w:top"/>
      </xsl:if>
      <xsl:if test="not(w:tblBorders/w:top)">
        <xsl:apply-templates
          select="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:top"
        />
      </xsl:if>
      <xsl:if test="w:tblBorders/w:right">
        <xsl:apply-templates select="w:tblBorders/w:right"/>
      </xsl:if>
      <xsl:if test="not(w:tblBorders/w:right)">
        <xsl:apply-templates
          select="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:right"
        />
      </xsl:if>
      <xsl:if test="w:tblBorders/w:bottom">
        <xsl:apply-templates select="w:tblBorders/w:bottom"/>
      </xsl:if>
      <xsl:if test="not(w:tblBorders/w:bottom)">
        <xsl:apply-templates
          select="document('word/styles.xml')/w:styles/w:style[@w:type='table' and @w:styleId=$tblId]/w:tblPr/w:tblBorders/w:bottom"
        />
      </xsl:if>
    </字:文字表边框_4227>
  </xsl:template>
  <xsl:template match="w:left">
    <uof:左_C613>
      <xsl:call-template name="border"/>
    </uof:左_C613>
  </xsl:template>
  <xsl:template match="w:top">
    <uof:上_C614>
      <xsl:call-template name="border"/>
    </uof:上_C614>
  </xsl:template>
  <xsl:template match="w:right">
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
    <uof:内部横线_C619>
      <xsl:call-template name="border"/>
    </uof:内部横线_C619>
  </xsl:template>
  <xsl:template match="w:tr2bl">
    <uof:内部竖线_C61A>
      <xsl:call-template name="border"/>
    </uof:内部竖线_C61A>
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

    --><!--修改线型虚实的转换--><!--

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
      <xsl:when test="./@w:val='dashSmallGap'">
        <xsl:attribute name="线型_C60D">single</xsl:attribute>
        <xsl:attribute name="虚实_C60E">round-dot</xsl:attribute>
      </xsl:when>
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
      <xsl:when test="./@w:val='thinThickSmallGap'">
        <xsl:attribute name="线型_C60D">thin-thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='thickThinSmallGap'">
        <xsl:attribute name="线型_C60D">thick-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='thinThickMediumGap'">
        <xsl:attribute name="线型_C60D">thin-thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='thickThinMediumGap'">
        <xsl:attribute name="线型_C60D">thick-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='thickThinLargeGap'">
        <xsl:attribute name="线型_C60D">thick-thin</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='thinThickLargeGap'">
        <xsl:attribute name="线型_C60D">thin-thick</xsl:attribute>
        <xsl:attribute name="虚实_C60E">long-dash</xsl:attribute>
      </xsl:when>
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
      <xsl:when test="./@w:val='wave'">
        <xsl:attribute name="线型_C60D">wave</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='doubleWave'">
        <xsl:attribute name="线型_C60D">double</xsl:attribute>
        <xsl:attribute name="虚实_C60E">solid</xsl:attribute>
      </xsl:when>
      <xsl:when test="./@w:val='dashDotStroked'">
        <xsl:attribute name="线型_C60D">dash-dot</xsl:attribute>
      </xsl:when>
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
    --><!--  
    <xsl:if test="./@w:shadow">
      <xsl:attribute name="uof:阴影">
        <xsl:if test="./@w:shadow='on' or ./@w:shadow='1' or ./@w:shadow='true'">
          <xsl:value-of select="'true'"/>
        </xsl:if>
        <xsl:if test="./@w:shadow='off' or ./@w:shadow='0' or ./@w:shadow='false'">
          <xsl:value-of select="'false'"/>
        </xsl:if>
      </xsl:attribute>
    </xsl:if> --><!--
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
    <xsl:if test="$Cellmar/w:left[@w:type!='auto']">
      <xsl:attribute name="左_C608">
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="$Cellmar/w:left"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="$Cellmar/w:right[@w:type!='auto']">
      <xsl:attribute name="右_C60A">
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="$Cellmar/w:right"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="$Cellmar/w:bottom[@w:type!='auto']">
      <xsl:attribute name="下_C60B">
        <xsl:call-template name="cellmar">
          <xsl:with-param name="cellmar" select="$Cellmar/w:bottom"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
  <xsl:template name="cellmar">
    <xsl:param name="cellmar"/>
    <xsl:if test="$cellmar/@w:type='dxa' and $cellmar/@w:w">
      <xsl:value-of select="($cellmar/@w:w) div 20"/>
    </xsl:if>
    <xsl:if test="$cellmar/@w:type='nil'or not($cellmar/@w:w)">
      <xsl:value-of select="'0'"/>
    </xsl:if>
    <xsl:if test="$cellmar/@w:type='pct' and $cellmar/@w:w">
      <xsl:value-of select="($cellmar/@w:w) * 0.08522"/>
    </xsl:if>
  </xsl:template>-->
  <!--FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF-->
  <!--@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
  <!--句式样、段落式样、文字表式样属性模板-->
  <xsl:template name ="InsertStyleAttr">
    <xsl:attribute name="标识符_4100">
      <xsl:call-template name="IdProducer">
        <xsl:with-param name ="ooxId" select ="@w:styleId"/>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name ="名称_4101">
      <xsl:value-of select ="./w:name/@w:val"/>
    </xsl:attribute>
    <xsl:choose >
      <xsl:when test="(@w:default='1')or(@w:default='on')or(@w:default='true')">
        <xsl:attribute name ="类型_4102">default</xsl:attribute>
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name ="类型_4102">custom</xsl:attribute>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:attribute name ="别名_4103">
      <xsl:if test="./w:aliases">
        <xsl:value-of select ="./w:aliases"/>
      </xsl:if>
      <xsl:if test ="not(./w:aliases)">
        <xsl:value-of select ="./w:name/@w:val"/>
      </xsl:if>
    </xsl:attribute>
    <xsl:if test="./w:basedOn">
      <xsl:attribute name="基式样引用_4104">
        <xsl:call-template name="IdProducer">
          <xsl:with-param name ="ooxId" select ="./w:basedOn/@w:val"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:attribute name ="后继式样引用_4105">
      <xsl:if test ="./w:next">
        <xsl:call-template name="IdProducer">
          <xsl:with-param name ="ooxId" select ="./w:next/@w:val"/>
        </xsl:call-template>
      </xsl:if>
      <xsl:if test ="not(./w:next)">
        <xsl:call-template name="IdProducer">
          <xsl:with-param name ="ooxId" select ="@w:styleId"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:attribute>
  </xsl:template>

  <xsl:template name ="IdProducer">
    <xsl:param name ="ooxId"/>
    <xsl:choose>
      <xsl:when test="starts-with($ooxId,'id_')">
        <!--用于比较两个字符串，若第一个是以第二个开始，返回True-->
        <xsl:value-of select ="$ooxId"/>
        <!--用于提取选定节点的值-->
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select ="concat('id_',$ooxId)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
</xsl:stylesheet>
