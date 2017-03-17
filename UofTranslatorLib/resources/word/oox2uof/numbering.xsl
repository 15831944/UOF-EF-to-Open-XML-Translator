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
<Author>fangchunyan(BITI)</Author>
<Author>guyueqiong</Author>
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships" 
  xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office" 
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" 
  xmlns:v="urn:schemas-microsoft-com:vml" 
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" 
  xmlns:w10="urn:schemas-microsoft-com:office:word" 
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" 
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles">
  
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>
  <xsl:template name="numberingSet">
    
    <!--转换word:numbering.xml到UOF中的自动编号 2011-02-19-->
    <xsl:if test="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Target='numbering.xml']">
      <xsl:apply-templates select="document('word/numbering.xml')/w:numbering" mode="numSet"/>
    </xsl:if>
    
  </xsl:template>
  
  <!--处理自动编号集-->
  <xsl:template match="w:numbering" mode="numSet">
    <式样:自动编号集_990E>
      <xsl:for-each select="w:num">
        <xsl:variable name="id" select="@w:numId"/>
        <xsl:variable name="absid" select="w:abstractNumId/@w:val"/>
        <xsl:apply-templates select="preceding-sibling::w:abstractNum[@w:abstractNumId = $absid]" mode="set">
          <xsl:with-param name="idv" select="$id"/>
        </xsl:apply-templates>
      </xsl:for-each>
    </式样:自动编号集_990E>
  </xsl:template>
  <xsl:template match="w:abstractNum" mode="set">
    <xsl:param name="idv"/>
    <字:自动编号_4124>
      <xsl:attribute name="标识符_4100">
        <xsl:value-of select="concat('bn_',$idv)"/>
      </xsl:attribute>
      <xsl:attribute name="父编号引用_4123">
        <xsl:choose>
          <xsl:when test="./w:multiLevelType/@w:val='singleLevel'">
            <xsl:value-of select="'false'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'true'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:if test="w:name">
        <xsl:attribute name="名称_4122">
          <xsl:value-of select="./w:name/@w:val"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:for-each select="w:lvl">
        <字:级别_4112>
          
          <xsl:attribute name="级别值_4121">
            <!--杨晓，+1，2010.1.25-->
            <xsl:value-of select="./@w:ilvl+1"/>
          </xsl:attribute>
          
          <xsl:if test="w:lvlJc">
            <字:编号对齐方式_4113>
              <xsl:choose>
                <xsl:when test="w:lvlJc/@w:val='left' or w:lvlJc/@w:val='right' or w:lvlJc/@w:val='center'">
                  <xsl:value-of select="w:lvlJc/@w:val"/>
                </xsl:when>
                <xsl:otherwise>left</xsl:otherwise>
              </xsl:choose>
            </字:编号对齐方式_4113>              
          </xsl:if>

          <字:尾随字符_4114>
            <xsl:choose>
              <xsl:when test ="w:suff">
                <xsl:if test="w:suff/@w:val='space'">space</xsl:if>
                <xsl:if test="w:suff/@w:val='tab'">tab</xsl:if>
                <xsl:if test="w:auff/@w:val='nothing'">none</xsl:if>
              </xsl:when>
              <xsl:otherwise>tab</xsl:otherwise>
            </xsl:choose>
          </字:尾随字符_4114>                 

          <xsl:choose>
            <xsl:when test="./w:numFmt/@w:val='bullet'">
              <字:项目符号_4115>
                <xsl:value-of select="./w:lvlText/@w:val"/>
              </字:项目符号_4115>
              <xsl:if test="./w:rPr">
                <字:符号字体_4116>
                  <xsl:apply-templates select="w:rPr" mode="RunProperties"/>
                </字:符号字体_4116>
              </xsl:if>
              <xsl:if test="w:pStyle">
                <字:链接式样引用_4118>
                  <xsl:variable name ="id" select ="./w:pStyle/@w:val"/>
                  <xsl:choose>
                    <xsl:when test="starts-with($id,'id_')">
                      <xsl:value-of select ="$id"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select ="concat('id_',$id)"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </字:链接式样引用_4118>
              </xsl:if>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="w:rPr">
                <字:符号字体_4116>
                  <xsl:apply-templates select="w:rPr" mode="RunProperties"/>
                </字:符号字体_4116>
              </xsl:if>
              <xsl:if test="w:pStyle">
                <字:链接式样引用_4118>
                  <xsl:variable name ="id" select ="./w:pStyle/@w:val"/>
                  <xsl:choose>
                    <xsl:when test="starts-with($id,'id_')">
                      <xsl:value-of select ="$id"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select ="concat('id_',$id)"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </字:链接式样引用_4118>
              </xsl:if>
              <字:编号格式_4119>
                <xsl:call-template name ="numFmtVal">
                  <xsl:with-param name ="val" select ="./w:numFmt/@w:val"/>
                </xsl:call-template>
              </字:编号格式_4119>
         <字:编号格式表示_411A>
				  
           <!--2012-01-17，修复OOX到UOF方向BUG：句式样-编号样式中01、001、0001等样式丢失，start-->
           <xsl:choose>
             <xsl:when test ="./w:numFmt/@w:val ='decimalZero'">
               <xsl:value-of select="concat('0',./w:lvlText/@w:val)"/>
             </xsl:when>
             <xsl:when test ="contains(./mc:AlternateContent/mc:Choice/w:numFmt/@w:format,'001, 002, 003, ...') ='true' and ./mc:AlternateContent/mc:Fallback//w:numFmt/@w:val ='decimal'">
               <xsl:value-of select="concat('00',./w:lvlText/@w:val)"/>
             </xsl:when>
             <xsl:when test ="contains(./mc:AlternateContent/mc:Choice/w:numFmt/@w:format,'0001, 0002, 0003, ...') ='true' and ./mc:AlternateContent/mc:Fallback//w:numFmt/@w:val ='decimal'">
               <xsl:value-of select="concat('000',./w:lvlText/@w:val)"/>
             </xsl:when>
             <xsl:when test ="contains(./mc:AlternateContent/mc:Choice/w:numFmt/@w:format,'00001, 00002, 00003, ...') ='true' and ./mc:AlternateContent/mc:Fallback//w:numFmt/@w:val ='decimal'">
               <xsl:value-of select="concat('0000',./w:lvlText/@w:val)"/>
             </xsl:when>
             <xsl:otherwise>
               <xsl:value-of select="./w:lvlText/@w:val"/>
             </xsl:otherwise>
           </xsl:choose>
           <!--end-->
           
			  </字:编号格式表示_411A>
            </xsl:otherwise>
          </xsl:choose>

          <xsl:if test="w:lvlPicBulletId">
            <xsl:variable name ="bulletId">
              <xsl:value-of select ="w:lvlPicBulletId/@w:val"/>
            </xsl:variable>
            <字:图片符号_411B>
              <xsl:apply-templates select ="document('word/numbering.xml')/w:numbering/w:numPicBullet[@w:numPicBulletId=$bulletId]" mode ="numbering"/>
              <xsl:attribute name="引用_411C">
                <xsl:value-of select ="concat('numbering','Obj',$bulletId * 2)"/>
              </xsl:attribute>              
            </字:图片符号_411B>
          </xsl:if>

          <xsl:if test="w:pPr">
            <xsl:if test="w:pPr/w:ind">
              <xsl:apply-templates select="w:pPr/w:ind" mode="Numbering"/>
            
            </xsl:if>
            <xsl:if test="w:pPr/w:tabs">
              <字:制表符位置_411E>
                <xsl:value-of select="./w:pPr/w:tabs/w:tab/@w:pos div 20"/>
              </字:制表符位置_411E>
            </xsl:if>

          </xsl:if>
          <字:起始编号_411F>
            <xsl:value-of select="./w:start/@w:val"/>
          </字:起始编号_411F>
          <xsl:if test="w:isLgl">
            <字:是否使用正规格式_4120>             
                <xsl:value-of select="./w:isLgl/@w:val"/>              
            </字:是否使用正规格式_4120>
          
        </xsl:if>
        </字:级别_4112>
      </xsl:for-each>
    </字:自动编号_4124>
  
</xsl:template>
  <xsl:template match ="w:numPicBullet" mode ="numbering">
    <xsl:if test ="w:pict/v:shape/@style">
      <xsl:variable name ="style">
        <xsl:value-of select ="w:pict/v:shape/@style"/>
      </xsl:variable>
      <xsl:variable name ="tmp1" select ="substring-after($style,'width:')"/>
      <xsl:variable name ="width" select ="substring-before($tmp1,';height')"/>
      <xsl:variable name ="height" select ="substring-after($tmp1,'height:')"/>
      <xsl:attribute name ="长_C604">
        <xsl:choose>
          <xsl:when test ="contains($width,'in')">
            <xsl:variable name ="widthval">
              <xsl:value-of select ="substring-before($width,'in')"/>
            </xsl:variable>
            <xsl:value-of select ="round($widthval * 72)"/>
          </xsl:when>
          <xsl:when test ="contains($width,'pt')">
            <xsl:variable name ="widthval">
              <xsl:value-of select ="substring-before($width,'pt')"/>
            </xsl:variable>
            <xsl:value-of select ="round($widthval)"/>
          </xsl:when>
          <xsl:when test ="contains($width,'cm')">
            <xsl:variable name ="widthval">
              <xsl:value-of select ="substring-before($width,'cm')"/>
            </xsl:variable>
            <xsl:value-of select ="round($widthval * 28.3)"/>
          </xsl:when>
          <xsl:when test ="contains($width,'mm')">
            <xsl:variable name ="widthval">
              <xsl:value-of select ="substring-before($width,'mm')"/>
            </xsl:variable>
            <xsl:value-of select ="round($widthval * 2.83)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name ="宽_C605">
        <xsl:choose>
          <xsl:when test ="contains($height,'in')">
            <xsl:variable name ="heightval">
              <xsl:value-of select ="substring-before($height,'in')"/>
            </xsl:variable>
            <xsl:value-of select ="round($heightval * 72)"/>
          </xsl:when>
          <xsl:when test ="contains($height,'pt')">
            <xsl:variable name ="heightval">
              <xsl:value-of select ="substring-before($height,'pt')"/>
            </xsl:variable>
            <xsl:value-of select ="round($heightval)"/>
          </xsl:when>
          <xsl:when test ="contains($height,'cm')">
            <xsl:variable name ="heightval">
              <xsl:value-of select ="substring-before($height,'cm')"/>
            </xsl:variable>
            <xsl:value-of select ="round($heightval * 28.3)"/>
          </xsl:when>
          <xsl:when test ="contains($height,'mm')">
            <xsl:variable name ="heightval">
              <xsl:value-of select ="substring-before($height,'mm')"/>
            </xsl:variable>
            <xsl:value-of select ="round($heightval * 2.83)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>
  <xsl:template match="w:body" mode="numfield">
    <xsl:for-each select="./w:p">
      <字:段落_416B 标识符_4169="段落">
        <xsl:if test="./w:pPr">
          <字:段落属性_419B>
            <xsl:if test="./w:pPr/w:numPr">
              <字:自动编号信息_4186>
                <xsl:variable name="level">
                  <xsl:value-of select="./w:pPr/w:numPr/w:ilvl/@w:val"/>
                </xsl:variable>
                <xsl:attribute name="编号引用_4187">
                  <xsl:value-of select="./w:pPr/w:numPr/w:numId/@w:val"/>
                </xsl:attribute>
                <xsl:attribute name="编号级别_4188">             
                  <xsl:value-of select="$level"/>
                </xsl:attribute>
              </字:自动编号信息_4186>
            </xsl:if>
          </字:段落属性_419B>
        </xsl:if>
      </字:段落_416B>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="w:ind" mode="Numbering">
    <字:缩进_411D>
      <xsl:call-template name="computInd"/>
    </字:缩进_411D>
  </xsl:template>
  <xsl:template name ="numFmtVal">
    <xsl:param name ="val"/>
    <xsl:choose>
      <xsl:when test="$val='decimal'">
        <xsl:value-of select="'decimal'"/>
      </xsl:when>
      <xsl:when test="$val='upperRoman'">
        <xsl:value-of select="'upper-roman'"/>
      </xsl:when>
      <xsl:when test="$val='lowerRoman'">
        <xsl:value-of select="'lower-roman'"/>
      </xsl:when>
      <xsl:when test="$val='upperLetter'">
        <xsl:value-of select="'upper-letter'"/>
      </xsl:when>
      <xsl:when test="$val='lowerLetter'">
        <xsl:value-of select="'lower-letter'"/>
      </xsl:when>
      <xsl:when test="$val='ordinal'">
        <xsl:value-of select="'ordinal'"/>
      </xsl:when>
      <xsl:when test="$val='cardinalText'">
        <xsl:value-of select="'cardinal-text'"/>
      </xsl:when>
      <xsl:when test="$val='ordinalText'">
        <xsl:value-of select="'ordinal-text'"/>
      </xsl:when>
      <xsl:when test="$val='hex'">
        <xsl:value-of select="'hex'"/>
      </xsl:when>
      <xsl:when test="$val='decimalFullWidth'">
        <xsl:value-of select="'decimal-full-width'"/>
      </xsl:when>
      <xsl:when test="$val='decimalHalfWidth'">
        <xsl:value-of select="'decimal-half-width'"/>
      </xsl:when>
      <xsl:when test="$val='decimalEnclosedCircle'">
        <xsl:value-of select="'decimal-enclosed-circle'"/>
      </xsl:when>
      <xsl:when test="$val='decimalEnclosedFullstop'">
        <xsl:value-of select="'decimal-enclosed-fullstop'"/>
      </xsl:when>
      <xsl:when test="$val='decimalEnclosedParen'">
        <xsl:value-of select="'decimal-enclosed-paren'"/>
      </xsl:when>
      <xsl:when test="$val='decimalEnclosedCircleChinese'">
        <xsl:value-of select="'decimal-enclosed-circle-chinese'"/>
      </xsl:when>
      <xsl:when test="$val='ideographEnclosedCircle'">
        <xsl:value-of select="'ideograph-enclosed-circle'"/>
      </xsl:when>
      <xsl:when test="$val='ideographTraditional'">
        <xsl:value-of select="'ideograph-traditional'"/>
      </xsl:when>
      <xsl:when test="$val='ideographZodiac'">
        <xsl:value-of select="'ideograph-zodiac'"/>
      </xsl:when>
      <xsl:when test="$val='chineseCountingThousand'">
        <xsl:value-of select="'chinese-counting'"/>
      </xsl:when>
      <xsl:when test="$val='chineseCounting'">
        <xsl:value-of select="'chinese-counting'"/>
      </xsl:when>
      <xsl:when test="$val='chineseLegalSimplified'">
        <xsl:value-of select="'chinese-legal-simplified'"/>
      </xsl:when>

      <!--2014-04-16，wudi，修复项目编号转换BUG，UOF标准中不存在japanese-counting枚举值，故转成chinese-counting，start-->
      <xsl:when test="$val='japaneseCounting'">
        <xsl:value-of select="'chinese-counting'"/>
      </xsl:when>
      <!--end-->
      
      <xsl:otherwise>
        <xsl:value-of select="'decimal'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>
