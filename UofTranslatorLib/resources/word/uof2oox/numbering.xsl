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
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
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
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:对象="http://schemas.uof.org/cn/2009/objects">
  <xsl:output encoding="UTF-8" method="xml" version="1.0"/>
  <xsl:template name="numbering">
    <!--wcz,2013/3/24，start-->
    <w:numbering>
      
      <!--2013-04-17，wudi，注释掉以下代码，修复图片编号样式丢失的BUG-->
      <!--<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
        <v:stroke joinstyle="miter"/>
        <v:formulas>
          <v:f eqn="if lineDrawn pixelLineWidth 0"/>
          <v:f eqn="sum @0 1 0"/>
          <v:f eqn="sum 0 0 @1"/>
          <v:f eqn="prod @2 1 2"/>
          <v:f eqn="prod @3 21600 pixelWidth"/>
          <v:f eqn="prod @3 21600 pixelHeight"/>
          <v:f eqn="sum @0 0 1"/>
          <v:f eqn="prod @6 1 2"/>
          <v:f eqn="prod @7 21600 pixelWidth"/>
          <v:f eqn="sum @8 21600 0"/>
          <v:f eqn="prod @7 21600 pixelHeight"/>
          <v:f eqn="sum @10 21600 0"/>
        </v:formulas>
        <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
        <o:lock v:ext="edit" aspectratio="t"/>
      </v:shapetype>-->
      <!--end-->
      
		  <xsl:for-each select="//对象:对象数据_D701">
        <xsl:variable name="numPicBulletId">
          <xsl:value-of select="@标识符_D704"/>
        </xsl:variable>
        
        <!--2014-05-07，wudi，修复图片符号转换时w:numPicBulletId取值为空的BUG，start-->
			  <xsl:if test="/uof:UOF/uof:式样集/式样:自动编号集_990E/字:自动编号_4124/字:级别_4112/字:图片符号_411B[@引用_411C = $numPicBulletId]">
          <!--<test>
            <xsl:value-of select="$numPicBulletId"/>
          </test>-->
				  <w:numPicBullet>
					  <xsl:attribute name="w:numPicBulletId">
              
              <!--2013-04-17，wudi，修复图片编号样式丢失的BUG，start-->
              <xsl:choose>
                <xsl:when test ="contains($numPicBulletId,'Obj0000')">
                  <xsl:value-of select="substring-after($numPicBulletId,'Obj0000')"/>
                </xsl:when>
                <xsl:when test ="contains($numPicBulletId,'Obj000')">
                  <xsl:value-of select="substring-after($numPicBulletId,'Obj000')"/>
                </xsl:when>
                <xsl:when test ="contains($numPicBulletId,'Obj00')">
                  <xsl:value-of select="substring-after($numPicBulletId,'Obj00')"/>
                </xsl:when>
                <xsl:when test ="contains($numPicBulletId,'Obj0') and substring-after($numPicBulletId,'Obj0') != ''">
                  <xsl:value-of select="substring-after($numPicBulletId,'Obj0')"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="substring-after($numPicBulletId,'Obj')"/>
                </xsl:otherwise>
              </xsl:choose>
              <!--end-->
              
              <!--end-->
						  
					  </xsl:attribute>
					  <w:pict>
              
              <!--2013-04-17，wudi，修复图片编号样式丢失的BUG，修改style值为width:15pt;height:15pt-->
              
              <!--2014-05-07，wudi，修改style值为width:11.25pt;height:11.25pt，start-->
						  <v:shape type="#_x0000_t75" style="width:11.25pt;height:11.25pt" o:bullet="t">
							  <xsl:attribute name="id">
								  <xsl:value-of select="concat('_x0000_i11',substring-after($numPicBulletId,'Obj'))"/>
							  </xsl:attribute>
							  <v:imagedata>
                  <xsl:attribute name="o:title">
                    <xsl:value-of select="concat('BD',substring-after($numPicBulletId,'Obj'),'_')"/>
                  </xsl:attribute>
								  <xsl:attribute name="r:id">
									  <xsl:value-of select="concat('rId',substring-after($numPicBulletId,'Obj'))"/>
								  </xsl:attribute>
							  </v:imagedata>
						  </v:shape>
              <!--end-->
              
					  </w:pict>
				  </w:numPicBullet>
			  </xsl:if>
		  </xsl:for-each>
      <!--end-->
      <xsl:for-each select="字:自动编号_4124">
        <w:abstractNum>
          <xsl:attribute name="w:abstractNumId">
            <xsl:call-template name="getNumberingId"/>
          </xsl:attribute>
          <w:multiLevelType>
            <xsl:attribute name="w:val">
              <xsl:value-of select="'multilevel'"/>
            </xsl:attribute>
          </w:multiLevelType>
          <xsl:if test="@名称_4122">
            <w:name>
              <xsl:attribute name="w:val">
                <xsl:value-of select="@名称_4122"/>
              </xsl:attribute>
            </w:name>
          </xsl:if>
          <xsl:for-each select="字:级别_4112">
            <w:lvl>
              <xsl:if test="@级别值_4121">
                <xsl:attribute name="w:ilvl">
                  <!--杨晓，-1，2010.1.25-->
                  <xsl:value-of select="./@级别值_4121 - 1"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:choose>
                <xsl:when test="字:起始编号_411F">
                  <w:start>
                    <xsl:attribute name="w:val">
                      <xsl:value-of select="./字:起始编号_411F"/>
                    </xsl:attribute>
                  </w:start>
                </xsl:when>
                <xsl:otherwise>
                  <w:start>
                    <xsl:attribute name="w:val">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </w:start>
                </xsl:otherwise>
              </xsl:choose>
              <xsl:if test="字:编号格式_4119">
                <w:numFmt>
                  <xsl:attribute name="w:val">
                    <xsl:call-template name="numberFormat">
                      <xsl:with-param name="val" select="字:编号格式_4119"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </w:numFmt>
              </xsl:if>
              <xsl:if test="字:项目符号_4115">
                <w:numFmt>
                  <xsl:attribute name="w:val">
                    <xsl:value-of select="'bullet'"/>
                  </xsl:attribute>
                </w:numFmt>
              </xsl:if>

              <!--2013-04-17，wudi，修复图片编号样式丢失的BUG，start-->
              <xsl:if test ="字:图片符号_411B">
                
                <xsl:variable name ="lvlPicBulletId">
                  <xsl:value-of select ="字:图片符号_411B/@引用_411C"/>
                </xsl:variable>
                <w:numFmt>
                  <xsl:attribute name="w:val">
                    <xsl:value-of select="'bullet'"/>
                  </xsl:attribute>
                </w:numFmt>
                <w:lvlText w:val="l"/>
                <w:lvlPicBulletId>
                  <xsl:attribute name ="w:val">
                    <xsl:choose>
                      <xsl:when test ="contains($lvlPicBulletId,'0000')">
                        <xsl:value-of select="substring-after($lvlPicBulletId,'Obj0000')"/>
                      </xsl:when>
                      <xsl:when test ="contains($lvlPicBulletId,'000')">
                        <xsl:value-of select="substring-after($lvlPicBulletId,'Obj000')"/>
                      </xsl:when>
                      <xsl:when test ="contains($lvlPicBulletId,'00')">
                        <xsl:value-of select="substring-after($lvlPicBulletId,'Obj00')"/>
                      </xsl:when>
                      <xsl:when test ="contains($lvlPicBulletId,'0')">
                        <xsl:value-of select="substring-after($lvlPicBulletId,'Obj0')"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="substring-after($lvlPicBulletId,'Obj')"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                </w:lvlPicBulletId>
                <w:lvlJc w:val="left"/>
              </xsl:if>
              <!--end-->
              
              <!--lvlRestart-->
              <!--pStyle-->
              <xsl:if test="字:链接式样引用_4118">
                <w:pStyle>
                  <xsl:attribute name="w:val">
                    <xsl:value-of select="./字:链接式样引用_4118"/>
                  </xsl:attribute>
                </w:pStyle>
              </xsl:if>
              <xsl:if test="字:是否使用正规格式_4120">
                <w:isLgl>
                  <xsl:attribute name="w:val">
                    <xsl:value-of select="./字:是否使用正规格式_4120"/>
                  </xsl:attribute>
                </w:isLgl>
              </xsl:if>

              <!--suff-->
              <xsl:if test="字:尾随字符_4114">
                <w:suff>
                  <xsl:attribute name="w:val">
                    <xsl:choose>
                      <xsl:when test="字:尾随字符_4114='space'">space</xsl:when>
                      <xsl:when test="字:尾随字符_4114='tab'">tab</xsl:when>
                      <xsl:when test="字:尾随字符_4114='none'">nothing</xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'tab'"/>
                        <xsl:message terminate="no"
                          >feedback:lost:lvl_and_suff_in_Numbering:Numbering</xsl:message>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                </w:suff>
              </xsl:if>
              <xsl:if test="./字:编号格式表示_411A">
                <w:lvlText>
                  <xsl:attribute name="w:val">
                    <xsl:value-of select="./字:编号格式表示_411A"/>
                  </xsl:attribute>
                </w:lvlText>
              </xsl:if>
              <xsl:if test="字:项目符号_4115">
                <w:lvlText>
                  <xsl:attribute name="w:val">
                    <xsl:value-of select="./字:项目符号_4115"/>
                  </xsl:attribute>
                </w:lvlText>
              </xsl:if>
              <!--wcz，2013/3/24
              <xsl:if test="字:图片符号_411B">
                <w:lvlPicBulletId>
                  <xsl:variable name="numPicBulletId">
                    <xsl:value-of select="substring-after(./字:图片符号_411B/@引用_411C,'Obj')"/>
                  </xsl:variable>
					        <xsl:attribute name="w:val">
						        <xsl:value-of select="$numPicBulletId"/>
					        </xsl:attribute>
                  <w:pict>
                    <v:shape type="#_x0000_t75" style="width:11.25pt;height:11.25pt" o:bullet="t">
                      <xsl:attribute name="id">
                        <xsl:value-of select="concat('_x0000_i11',$numPicBulletId)"/>
                      </xsl:attribute>
                      <v:imagedata>
                        <xsl:attribute name="r:id">
                          <xsl:value-of select="concat('rId',$numPicBulletId)"/>
                        </xsl:attribute>
                      </v:imagedata>
                    </v:shape>
                  </w:pict>
                </w:lvlPicBulletId>
              </xsl:if>  -->      
              <xsl:if test="字:编号对齐方式_4113">
                <w:lvlJc>
                  <xsl:attribute name="w:val">
                    <xsl:value-of select="字:编号对齐方式_4113"/>
                  </xsl:attribute>
                </w:lvlJc>
              </xsl:if>
              <!--legacy-->
              <!--lvlJc-->
              <!--pPr-->
              <xsl:if test="字:缩进_411D | 字:制表符位置_411E">
                <w:pPr>
                  <xsl:choose>
                    <xsl:when test="字:制表符位置_411E">
                      <xsl:apply-templates select="字:制表符位置_411E"/>
                    </xsl:when>
                    <xsl:when test="字:缩进_411D">
                      <xsl:apply-templates select="字:缩进_411D" mode="numbering"/>
                    </xsl:when>
                  </xsl:choose>
                </w:pPr>
              </xsl:if>
              <!--rPr-->
              <xsl:if test="字:符号字体_4116">
                <xsl:for-each select="字:符号字体_4116">
                  <xsl:call-template name="RunProperties"/>
                </xsl:for-each>
              </xsl:if>

            </w:lvl>
          </xsl:for-each>
        </w:abstractNum>
      </xsl:for-each>
      <!--MARK：自动编号 2011-03-16-->
      <xsl:for-each select="字:自动编号_4124">
        <w:num>
          <xsl:variable name="num">
            <xsl:call-template name="getNumberingId"/>
          </xsl:variable>
          <xsl:attribute name="w:numId">

            <xsl:value-of select="$num"/>
          </xsl:attribute>
          <w:abstractNumId>
            <xsl:attribute name="w:val">

              <xsl:value-of select="$num"/>
            </xsl:attribute>
          </w:abstractNumId>

          <!--2013-04-17，wudi，修复图片编号样式丢失的BUG，start-->
          <xsl:if test ="字:级别_4112/字:图片符号_411B">
            <xsl:variable name ="lvlPicBulletId">
              <xsl:value-of select ="字:级别_4112/字:图片符号_411B/@引用_411C"/>
            </xsl:variable>
            <!--2013-04-18，wudi，修改w:ilvl取值方式，设为1，start-->
            <w:lvlOverride>
              <xsl:attribute name ="w:ilvl">
                <xsl:value-of select ="'1'"/>
              </xsl:attribute>
            </w:lvlOverride>
            <!--end-->
          </xsl:if>
          <!--end-->
          
        </w:num>
      </xsl:for-each>
    </w:numbering>
  </xsl:template>

 
  <xsl:template name="pPrNumbering">
    <w:numPr>

      <!--2012-01-18，wudi，修复UOF到OOX方向自动编号的BUG，start-->
      <w:ilvl>
        <xsl:choose>
          <xsl:when test ="字:自动编号信息_4186/@编号级别_4188 = '0' and 字:自动编号信息_4186/@编号引用_4187 = 'bn_1'">
            <xsl:attribute name="w:val">
              <xsl:value-of select="字:自动编号信息_4186/@编号级别_4188"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="w:val">
              <xsl:value-of select="字:自动编号信息_4186/@编号级别_4188 - 1"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </w:ilvl>
      
      <w:numId>
        <xsl:variable name="id">
          <xsl:value-of select="字:自动编号信息_4186/@编号引用_4187"/>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test ="字:自动编号信息_4186/@编号级别_4188 = '0' and 字:自动编号信息_4186/@编号引用_4187 = 'bn_1'">
            <xsl:attribute name ="w:val">
              <xsl:value-of select ="'0'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="w:val">
              <xsl:apply-templates select="preceding::字:自动编号_4124[@标识符_4100=$id]" mode="getId"/>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
      </w:numId>
      <!--
      <w:ilvl>
          <xsl:attribute name="w:val">
            <xsl:value-of select="字:自动编号信息_4186/@编号级别_4188 - 1"/>
          </xsl:attribute>
      </w:ilvl>
      <w:numId>
        <xsl:variable name="id">
          <xsl:value-of select="字:自动编号信息_4186/@编号引用_4187"/>
        </xsl:variable>
        <xsl:attribute name="w:val">
          <xsl:apply-templates select="preceding::字:自动编号_4124[@标识符_4100=$id]" mode="getId"/>
        </xsl:attribute>
      </w:numId>
      -->
      <!--end-->

    </w:numPr>
  </xsl:template>
  <xsl:template match="字:自动编号_4124" mode="getId">
    <xsl:call-template name="getNumberingId"/>
  </xsl:template>

	<!--<xsl:template name="getNumPicBulletId">
		<xsl:variable name="num">
			<xsl:number count="字:图片符号_411B" format="1" level="single"/>
		</xsl:variable>
		<xsl:value-of select="$num"/>
	</xsl:template>-->
  <xsl:template name="getNumberingId">
    <xsl:variable name="num">
      <xsl:number count="字:自动编号_4124" format="1" level="single"/> 
    </xsl:variable>
    <xsl:value-of select="$num"/>
  </xsl:template>
  <xsl:template match="字:制表符位置_411E">
    <w:tabs>
      <w:tab>
        <xsl:attribute name="w:val">num</xsl:attribute>
        <xsl:attribute name="w:pos">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="."/>
          </xsl:call-template>
        </xsl:attribute>
      </w:tab>
    </w:tabs>
  </xsl:template>

  <xsl:template match="字:缩进_411D" mode="numbering">
    <w:ind>
      <xsl:variable name="firstValue">
        <xsl:choose>
          <xsl:when test="字:首行_4111/字:绝对_4107/@值_410F &lt; 0">
            <xsl:value-of select="-(字:首行_4111/字:绝对_4107/@值_410F)"/>
          </xsl:when>
          <xsl:when test="(字:首行_4111/字:绝对_4107/@值_410F &gt; 0)or(字:首行_4111/字:绝对_4107/@值_410F = 0)">
            <xsl:value-of select="'0'"/>
          </xsl:when>
          <xsl:when test="not(字:首行_4111/字:绝对_4107) or not(字:首行_4111)">
            <xsl:value-of select="'0'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>

      <xsl:for-each select="字:首行_4111 | 字:左_410E | 字:右_4110">
        <xsl:choose>
          <xsl:when test="name(.)='字:首行_4111'">
            <xsl:call-template name="FirstLineOrHangingForNumbering"/>
          </xsl:when>
          <xsl:when test="name(.)='字:左_410E'">
            <xsl:call-template name="ComputeLeftIndForNumbering">
              <xsl:with-param name="firstHang" select="$firstValue"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name(.)='字:右_4110'">
            <xsl:call-template name="ComputeRightIndForNumbering"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </w:ind>
  </xsl:template>

  <xsl:template name="ComputeLeftIndForNumbering">
    <xsl:param name="firstHang"/>
    <xsl:if test="./字:绝对_4107/@值_410F ">
      <xsl:attribute name="w:left">
        <xsl:call-template name="twipsMeasure">
          <xsl:with-param name="lengthVal" select="(./字:绝对_4107/@值_410F) + $firstHang"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="./字:相对_4109/@字:值_4108 ">
      <xsl:attribute name="w:leftChars">
        <xsl:value-of select="number(./字:相对_4109/@字:值_4108 * 100)"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template name="ComputeRightIndForNumbering">
    <xsl:if test="./字:绝对_4107/@值_410F ">
      <xsl:attribute name="w:right">
        <xsl:call-template name="twipsMeasure">
          <xsl:with-param name="lengthVal" select="./字:绝对_4107/@值_410F"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="./字:相对_4109/@字:值_4108 ">
      <xsl:attribute name="w:rightChars">
        <xsl:value-of select="number(./字:相对_4109/@字:值_4108 * 100)"/>
      </xsl:attribute>
    </xsl:if>
  </xsl:template>

  <xsl:template name="FirstLineOrHangingForNumbering">
    <xsl:choose>
      <xsl:when test="./字:绝对_4107/@值_410F &gt; 0">
        <xsl:attribute name="w:firstLine">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="./字:绝对_4107/@值_410F"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="./字:绝对_4107/@值_410F &lt; 0">
        <xsl:attribute name="w:hanging">
          <xsl:call-template name="twipsMeasure">
            <xsl:with-param name="lengthVal" select="-(./字:绝对_4107/@值_410F)"/>
          </xsl:call-template>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="./字:相对_4109/@字:值_4108 &gt; 0">
        <xsl:attribute name="w:firstLineChars">
          <xsl:value-of select="number(./字:相对_4109/@字:值_4108 * 100)"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="./字:相对_4109/@字:值_4108 &lt; 0">
        <xsl:attribute name="w:hangingChars">
          <xsl:value-of select="-(number(./字:相对_4109/@字:值_4108 * 100))"/>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="numberFormat">
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$val='decimal'">
        <xsl:value-of select="'decimal'"/>
      </xsl:when>
      <xsl:when test="$val='upper-roman'">
        <xsl:value-of select="'upperRoman'"/>
      </xsl:when>
      <xsl:when test="$val='lower-roman'">
        <xsl:value-of select="'lowerRoman'"/>
      </xsl:when>
      <xsl:when test="$val='upper-letter'">
        <xsl:value-of select="'upperLetter'"/>
      </xsl:when>
      <xsl:when test="$val='lower-letter'">
        <xsl:value-of select="'lowerLetter'"/>
      </xsl:when>
      <xsl:when test="$val='ordinal'">
        <xsl:value-of select="'ordinal'"/>
      </xsl:when>
      <xsl:when test="$val='cardinal-text'">
        <xsl:value-of select="'cardinalText'"/>
      </xsl:when>
      <xsl:when test="$val='ordinal-text'">
        <xsl:value-of select="'ordinalText'"/>
      </xsl:when>
      <xsl:when test="$val='hex'">
        <xsl:value-of select="'hex'"/>
      </xsl:when>
      <xsl:when test="$val='decimal-full-width'">
        <xsl:value-of select="'decimalFullWidth'"/>
      </xsl:when>
      <xsl:when test="$val='decimal-half-width'">
        <xsl:value-of select="'decimalHalfWidth'"/>
      </xsl:when>
      <xsl:when test="$val='decimal-enclosed-circle'">
        <xsl:value-of select="'decimalEnclosedCircle'"/>
      </xsl:when>
      <xsl:when test="$val='decimal-enclosed-fullstop'">
        <xsl:value-of select="'decimalEnclosedFullstop'"/>
      </xsl:when>
      <xsl:when test="$val='decimal-enclosed-paren'">
        <xsl:value-of select="'decimalEnclosedParen'"/>
      </xsl:when>
      <xsl:when test="$val='decimal-enclosed-circle-chinese'">
        <xsl:value-of select="'decimalEnclosedCircleChinese'"/>
      </xsl:when>
      <xsl:when test="$val='ideograph-enclosed-circle'">
        <xsl:value-of select="'ideographEnclosedCircle'"/>
      </xsl:when>
      <xsl:when test="$val='ideograph-traditional'">
        <xsl:value-of select="'ideographTraditional'"/>
      </xsl:when>
      <xsl:when test="$val='ideograph-zodiac'">
        <xsl:value-of select="'ideographZodiac'"/>
      </xsl:when>
      <xsl:when test="$val='chinese-counting'">
        <xsl:value-of select="'chineseCountingThousand'"/>
      </xsl:when>
      <xsl:when test="$val='chinese-legal-simplified'">
        <xsl:value-of select="'chineseLegalSimplified'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'decimal'"/>
        <xsl:message terminate="no"
          >feedback:lost:lvl_and_numFmt_in_Numbering:Numbering</xsl:message>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>
