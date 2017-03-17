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
  xmlns:对象="http://schemas.uof.org/cn/2009/objects"
  xmlns:fn="http://www.w3.org/2005/xpath-functions"
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
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:pic="http://purl.oclc.org/ooxml/drawingml/picture">
  <xsl:import href="common.xsl"/>
  <xsl:import href="numbering.xsl"/>
  <xsl:import href="sectPr.xsl"/>
  <xsl:import href="region.xsl"/>
  <xsl:import href="paragraph.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

  <xsl:template name="objectPicture">
    <!--cxl,2012/2/14,current node: uof:锚点_C644-->
    <xsl:param name="num"/>
    <!--<xsl:variable name="anchorType">
      <xsl:value-of select="@字:类型"/>
    </xsl:variable>-->

    <xsl:variable name="objref" select="./@图形引用_C62E"/>
    <xsl:variable name="isPic">
      <xsl:choose>
        <xsl:when test="//对象:对象数据_D701/@标识符_D704 = $objref">
          <xsl:value-of select="'datapic'"/>
        </xsl:when>
        <xsl:when test="//uof:对象集/图:图形_8062[@标识符_804B = $objref]/图:图片数据引用_8037">
          <xsl:value-of select="'objpic'"/>
        </xsl:when>
        <!--zhaobj 预定义图形-->
        <xsl:when test="//uof:对象集/图:图形_8062[@标识符_804B = $objref]">
          <xsl:value-of select="'prstpic'"/>
        </xsl:when>
        <!--
      <xsl:when test ="not((//uof:对象集/uof:其他对象/@uof:标识符 = $objref) or (//uof:对象集/图:图形[@标识符_804B = $objref]))">
        <xsl:value-of select ="'false'"/>
      </xsl:when>-->
        <xsl:otherwise>
          <xsl:value-of select="'false'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>

    <xsl:choose>
      <xsl:when test="$isPic='datapic' or $isPic='objpic'">
        <w:drawing>
          <xsl:choose>
            <xsl:when test="$isPic='datapic'">
              <xsl:variable name="id" select="./@图形引用_C62E"/>
              <!--<xsl:copy-of select="//对象:对象数据_D701[@标识符_D704 = $id]/u2opic:picture"/>-->
            </xsl:when>
            <xsl:when test="$isPic='objpic'">
              <xsl:variable name="id" select="./@图形引用_C62E"/>
              <xsl:variable name="oid" select="//uof:对象集/图:图形_8062[@标识符_804B = $id]/图:图片数据引用_8037"/>
              <!--<xsl:copy-of select="//对象:对象数据_D701[@标识符_D704 = $oid]/u2opic:picture"/>-->
            </xsl:when>
          </xsl:choose>

          <!--<xsl:when test="$anchorType='inline'">
          <wp:inline>
            <xsl:call-template name="anchorPictureInline">
              <xsl:with-param name="desType" select="$isPic"/>
              <xsl:with-param name="pos" select="$num"/>
            </xsl:call-template>
          </wp:inline>
        </xsl:when>
        <xsl:when test="$anchorType='normal'">-->

          <!--2013-04-01，修复UOF到OOX方向图片转换绕排方式转换错误的BUG，start-->
          <xsl:choose>
            <xsl:when test="not(./uof:绕排_C622/@绕排方式_C623)">
              <wp:inline>
                <xsl:call-template name="anchorPictureInline">
                  <xsl:with-param name="desType" select="$isPic"/>
                  <xsl:with-param name="pos" select="$num"/>
                </xsl:call-template>
              </wp:inline>
            </xsl:when>
            <xsl:otherwise>
              <wp:anchor>
                <xsl:call-template name="anchorPictureNormal">
                  <xsl:with-param name="desType" select="$isPic"/>
                  <xsl:with-param name="pos" select="$num"/>
                </xsl:call-template>
              </wp:anchor>
            </xsl:otherwise>
          </xsl:choose>

          <!--end-->

        </w:drawing>
      </xsl:when>

      <!--zhaobj 预定义图形-->
      <xsl:when test="$isPic='prstpic'">
        <w:drawing>
          <xsl:variable name="id" select="./@图形引用_C62E"/>
          <!--有图片填充的情况********************************************************************-->
          <xsl:if test="//uof:对象集/图:图形_8062[@标识符_804B = $id]/图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图片_8005/@图形引用_8007">
            <xsl:variable name="oid"
              select="//uof:对象集/图:图形_8062[@标识符_804B = $id]/图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图片_8005/@图形引用_8007"/>
            <!--<xsl:copy-of select="//对象:对象数据_D701[@标识符_D704 = $oid]/u2opic:picture"/>-->
          </xsl:if>
          <!--cxl2012/3/17,绕排方式为嵌入型时，OOX中对应为wp:inline，此元素子元素与wp:anchor下不同-->
          <!--<xsl:when test="$anchorType='inline'">-->

          <xsl:choose>
            <xsl:when test="not(./uof:绕排_C622/@绕排方式_C623)">
              <wp:inline>
                <xsl:call-template name="anchorPrstInline">
                  <xsl:with-param name="desType" select="$isPic"/>
                  <xsl:with-param name="pos" select="$num"/>
                </xsl:call-template>
              </wp:inline>
            </xsl:when>
            <!--<xsl:when test="$anchorType='normal'">-->
            <xsl:otherwise>
              <wp:anchor>
                <xsl:call-template name="anchorPrstNormal">
                  <xsl:with-param name="desType" select="$isPic"/>
                  <xsl:with-param name="pos" select="$num"/>
                </xsl:call-template>
              </wp:anchor>
            </xsl:otherwise>
          </xsl:choose>

        </w:drawing>
      </xsl:when>
    </xsl:choose>

    
  </xsl:template>
  <!--cxl,2012.3.17嵌入式预定义图形的模板-->
  <xsl:template name="anchorPrstInline">
    <!--当前结点：uof:锚点_C644-->
    <xsl:param name="desType"/>
    <xsl:param name="pos"/>

    <xsl:if test="./uof:边距_C628">
      <xsl:apply-templates select="./uof:边距_C628"/>
    </xsl:if>

    <xsl:variable name="ObjId" select="./@图形引用_C62E"/>
    <xsl:variable name="DataObjId" select="//图:图形_8062[@标识符_804B = $ObjId]/图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图片_8005/@图形引用_8007"/>

    <xsl:if test="./uof:大小_C621">
      
      <!--2013-04-11，wudi，改"../uof:锚点_C644"为"."-->
      <xsl:apply-templates select="." mode="pictureProperties"/>
      <!--产生wp:extent-->
    </xsl:if>
    <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D">
      <xsl:variable name="prstrot">
        <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D"/>
      </xsl:variable>
      <xsl:variable name="prstheight">
        <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605"/>
      </xsl:variable>
      <xsl:variable name="prstweight">
        <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604"/>
      </xsl:variable>
      <xsl:variable name="extentcha">
        <xsl:choose>
          <xsl:when test="$prstheight &gt; $prstweight">
            <xsl:value-of select="round(($prstheight - $prstweight) * 6350)"/>
          </xsl:when>
          <xsl:when test="$prstheight &lt; $prstweight">
            <xsl:value-of select="round(($prstweight - $prstheight) * 6350)"/>
          </xsl:when>
          <xsl:when test="$prstheight=$prstweight">
            <xsl:value-of select="round(0)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <xsl:when
          test="$prstrot='0.0' or $prstrot='90.0' or $prstrot='180.0' or $prstrot='270.0' or $prstrot='360.0'">
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
        </xsl:when>

        <xsl:when test="$prstrot &gt; 0.0 and $prstrot &lt; 45.0">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot = '45.0'">
          <wp:effectExtent t="0" b="0">
            <xsl:attribute name="l">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="r">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot &gt; 45.0 and $prstrot &lt; 135.0">
          <wp:effectExtent t="0" b="0">
            <xsl:attribute name="l">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="r">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot = '135.0'">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot &gt; 135.0 and $prstrot &lt; 225.0">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot = '225.0'">
          <wp:effectExtent t="0" b="0">
            <xsl:attribute name="l">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="r">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot &gt; 225.0 and $prstrot &lt; 315.0  ">
          <wp:effectExtent t="0" b="0">
            <xsl:attribute name="l">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="r">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot = '315.0'">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot &gt; 315.0 and $prstrot &lt; 360.0  ">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
      </xsl:choose>
    </xsl:if>

    <wp:docPr>
      <xsl:attribute name="id">

        <!--2014-04-09，Smiley Face的id取值需特殊处理，不然会带来文档需要修复的问题，start-->
        <xsl:choose>
          <xsl:when test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A = 'Smiley Face'">
            <xsl:value-of select="number(substring-after($ObjId,'Obj')) + 1"/>
          </xsl:when>

          <!--2014-05-20，wudi，修复页眉中插入图片时，转换后文档需要修复的BUG，start-->
          <!--2014-05-27，wudi，修复页眉中插入图片时，转换后文档需要修复的BUG，将+2改成+4，start-->
          <xsl:when test="starts-with($ObjId,'header')">
            <xsl:value-of select="number(substring-after($ObjId,'Obj')) + 4"/>
          </xsl:when>
          <!--end-->
          <!--end-->
          
          <xsl:otherwise>
            <xsl:value-of select="substring-after($ObjId,'Obj')"/>
          </xsl:otherwise>
        </xsl:choose>
        <!--end-->
        
        <!--<xsl:apply-templates select="//uof:其他对象[@uof:标识符 = $DataObjId]" mode="num"/>-->
      </xsl:attribute>
      <!--考虑组合图形的情况******************************************************-->
      <xsl:attribute name="name">
        <xsl:choose>
          <xsl:when test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/@组合列表_8064">

            <!--2013-05-03，wudi，修复组合图形，流程图，SmartArt互操作BUG，三种“组合图形”用不同的标识符，start-->
            <xsl:choose>
              <xsl:when test ="contains(//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/@组合列表_8064,'grpspObj')">
                <xsl:value-of select ="'组合'"/>
              </xsl:when>
              <xsl:when test ="contains(//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/@组合列表_8064,'GrpspObj')">
                <xsl:value-of select="'Group'"/>
              </xsl:when>
              <xsl:when test ="contains(//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/@组合列表_8064,'SmartArtObj')">
                <xsl:value-of select ="'图示'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="'组合'"/>
              </xsl:otherwise>
            </xsl:choose>
            <!--end-->
            
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </wp:docPr>
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks>
        <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:缩放是否锁定纵横比_8055">
          <xsl:attribute name="noChangeAspect">
            <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:缩放是否锁定纵横比_8055"/>
          </xsl:attribute>
        </xsl:if>
      </a:graphicFrameLocks>
    </wp:cNvGraphicFramePr>
    <a:graphic>
      <xsl:choose>
        <xsl:when test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/@组合列表_8064">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
            <wpg:wgp>

              <!--2013-05-02，wudi，增加参数isSmartArt，标识是否是SmartArt转换，目前SmartArt按组合图形转，start-->

              <!--2014-03-27，wudi，增加变量isSmartArt，用于判断组合图形是否是SmartArt，start-->
              <xsl:variable name="isSmartArt">
                <xsl:choose>
                  <xsl:when test="contains(//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/@组合列表_8064,'SmartArtObj')">
                    <xsl:value-of select="'1'"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:apply-templates select="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId and @组合列表_8064]"
                mode="wgpprstGeom">
                <xsl:with-param name="num" select="$pos"/>
                <xsl:with-param name ="tag" select ="'1'"/>
                <xsl:with-param name ="isSmartArt" select ="$isSmartArt"/>
                <xsl:with-param name ="anchorx" select ="'0'"/>
                <xsl:with-param name ="anchory" select ="'0'"/>
              </xsl:apply-templates>
              <!--end-->

              <!--end-->

            </wpg:wgp>

          </a:graphicData>
        </xsl:when>
        <xsl:otherwise>
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>

              <!--2014-03-13，wudi，原取值方式有误，修正，去掉‘图:文字排列方向_8042’前的@，条件值为't2b-l2r-270e-0w'，start-->
              <xsl:if test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/图:文本_803C/图:文字排列方向_8042='t2b-l2r-270e-0w'">
                <xsl:attribute name="normalEastAsianFlow">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
              </xsl:if>
              <!--end-->

              <xsl:choose>
                <xsl:when test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:类别_8019 = '61'">
                  <wps:cNvCnPr/>
                </xsl:when>
                <xsl:when test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:类别_8019 = '71'">
                  <wps:cNvCnPr/>
                </xsl:when>
                <xsl:otherwise>
                  <wps:cNvSpPr>
                    <xsl:if test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/图:文本_803C">
                      <xsl:attribute name="txBox">
                        <xsl:value-of select="'1'"/>
                      </xsl:attribute>
                    </xsl:if>
                  </wps:cNvSpPr>
                </xsl:otherwise>
              </xsl:choose>
              <wps:spPr>
                <a:xfrm>
                  <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D">
                    <xsl:attribute name="rot">
                      <xsl:value-of select="round(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D * 60000)"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!--cxl,2012.3.17图形翻转-->

                  <!--2013-03-26，wudi，修复预定义图形方向不对的BUG，此差异与永中软件有关，start-->
                  <!--2013-03-28，wudi，修复功能测试OOX到UOF2方向ShapeStyle右边最下行的黄色矩形方向与原始不一致的BUG，增加限制条件Rectangular Callout-->
                  <xsl:if test="(//图:图形_8062[@标识符_804B=$ObjId]/图:翻转_803A='x') and not(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A ='Line Callout4') and not(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A ='Line Callout4(Accent Bar)') and not(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A ='Line Callout4(No Border)') and not(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A ='Line Callout4(Border and Accent Bar)') and not(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A ='Rectangular Callout')">
                    <xsl:attribute name="flipH">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!--end-->

                  <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:翻转_803A='y'">
                    <xsl:attribute name="flipV">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <a:off x="0" y="0"/>

                  <!--2013-04-11，wudi，改"../uof:锚点_C644"为"."-->
                  <xsl:apply-templates select="." mode="xfrm"/>
                </a:xfrm>
                <xsl:apply-templates select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018" mode="prstGeom">
                  <xsl:with-param name="num" select="$pos"/>
                </xsl:apply-templates>
              </wps:spPr>
              <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:文本_803C">
                <xsl:apply-templates select="//图:图形_8062[@标识符_804B=$ObjId]/图:文本_803C" mode="wpstxbx"/>
              </xsl:if>
              <!--cxl,2012.2.25新增艺术字转换-->
              <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]//图:艺术字_802D">
                <xsl:apply-templates select="//图:图形_8062[@标识符_804B=$ObjId]//图:艺术字_802D" mode="wpstxbx"/>
              </xsl:if>
              <wps:bodyPr>
                <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:文本_803C">
                  <xsl:apply-templates select="//图:图形_8062[@标识符_804B=$ObjId]/图:文本_803C" mode="txbody">
                    <xsl:with-param name="objectid" select="$ObjId"/>
                  </xsl:apply-templates>
                </xsl:if>
                <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]//图:艺术字_802D">
                  <xsl:apply-templates select="//图:图形_8062[@标识符_804B=$ObjId]//图:艺术字_802D" mode="txbody"/>
                </xsl:if>
              </wps:bodyPr>
              <!-- <xsl:if test ="not(//uof:扩展区/uof:扩展/uof:扩展内容/uof:内容/uof:文本框属性[@uof:对象索引 = $ObjId]/@uof:是否随图形旋转文字)">
              <xsl:attribute name ="upright">
                <xsl:value-of select ="1"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="//uof:扩展区/uof:扩展/uof:扩展内容/uof:内容/uof:文本框属性[@uof:对象索引 = $ObjId]/@uof:是否随图形旋转文字">
              <xsl:attribute name ="upright">
                <xsl:value-of select ="0"/>
              </xsl:attribute>
            </xsl:if>-->

            </wps:wsp>
          </a:graphicData>
        </xsl:otherwise>
      </xsl:choose>
      
    </a:graphic>
  </xsl:template>

  <xsl:template name="anchorPrstNormal">
    <!--当前结点：uof:锚点_C644-->
    <xsl:param name="desType"/>
    <xsl:param name="pos"/>

    <xsl:if test="./uof:边距_C628">
      <xsl:apply-templates select="./uof:边距_C628"/>
    </xsl:if>
    <xsl:attribute name="simplePos">
      <xsl:value-of select="'0'"/>
    </xsl:attribute>

    <xsl:variable name="ObjId" select="./@图形引用_C62E"/>
    <xsl:variable name="DataObjId" select="//图:图形_8062[@标识符_804B = $ObjId]/图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图片_8005/@图形引用_8007"/>
    <xsl:attribute name="relativeHeight">
      <xsl:choose>
        <xsl:when test="//图:图形_8062[@标识符_804B=$ObjId]/@层次_8063">
          <xsl:variable name="arrg">
            <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/@层次_8063"/>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="$arrg &lt; 0">
              <xsl:value-of select="'251658240'"/>
            </xsl:when>
            <xsl:when test="$arrg &gt; 0">
              <xsl:value-of select="$arrg"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$arrg"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'251658240'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>

    <xsl:attribute name="behindDoc">
      <xsl:variable name="roundMethod">
        <xsl:value-of select="./uof:绕排_C622/@绕排方式_C623"/>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$roundMethod='behind-text'">
          <xsl:value-of select="'1'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'0'"/>
        </xsl:otherwise>
      </xsl:choose>
      
    </xsl:attribute>
    <xsl:attribute name="locked">
      <xsl:choose>
        <xsl:when test="./uof:是否锁定_C629">
          <xsl:value-of select="./uof:是否锁定_C629"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'false'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="layoutInCell">0</xsl:attribute>
    <xsl:attribute name="allowOverlap">
      <xsl:choose>
        <xsl:when test="./uof:是否允许重叠_C62B">
          <xsl:value-of select="./uof:是否允许重叠_C62B"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'false'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <wp:simplePos x="0" y="0"/>
    <xsl:if test="./uof:位置_C620">
      <!--产生wp:positionH和wp:positionV-->
      <xsl:apply-templates select="./uof:位置_C620" mode="anchorPic"/>
    </xsl:if>
    <xsl:if test="./uof:大小_C621">

      <!--2013-04-11，wudi，改"../uof:锚点_C644"为"."-->
      <xsl:apply-templates select="." mode="pictureProperties"/>
      <!--产生wp:extent-->
    </xsl:if>
    <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D">
      <xsl:variable name="prstrot">
        <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D"/>
      </xsl:variable>
      <xsl:variable name="prstheight">
        <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605"/>
      </xsl:variable>
      <xsl:variable name="prstweight">
        <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604"/>
      </xsl:variable>
      <xsl:variable name="extentcha">
        <xsl:choose>
          <xsl:when test="$prstheight &gt; $prstweight">
            <xsl:value-of select="round(($prstheight - $prstweight) * 6350)"/>
          </xsl:when>
          <xsl:when test="$prstheight &lt; $prstweight">
            <xsl:value-of select="round(($prstweight - $prstheight) * 6350)"/>
          </xsl:when>
          <xsl:when test="$prstheight=$prstweight">
            <xsl:value-of select="0"/>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <xsl:when
          test="$prstrot='0.0' or $prstrot='90.0' or $prstrot='180.0' or $prstrot='270.0' or $prstrot='360.0'">
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
        </xsl:when>

        <xsl:when test="$prstrot &gt; 0.0 and $prstrot &lt; 45.0">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot = '45.0'">
          <wp:effectExtent t="0" b="0">
            <xsl:attribute name="l">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="r">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot &gt; 45.0 and $prstrot &lt; 135.0">
          <wp:effectExtent t="0" b="0">
            <xsl:attribute name="l">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="r">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot = '135.0'">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot &gt; 135.0 and $prstrot &lt; 225.0">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot = '225.0'">
          <wp:effectExtent t="0" b="0">
            <xsl:attribute name="l">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="r">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot &gt; 225.0 and $prstrot &lt; 315.0  ">
          <wp:effectExtent t="0" b="0">
            <xsl:attribute name="l">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="r">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot = '315.0'">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot &gt; 315.0 and $prstrot &lt; 360.0  ">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="./uof:绕排_C622">
      <xsl:variable name="method">
        <!--cxl,2012.3.17添加默认方式-->
        <xsl:choose>
          <xsl:when test="./uof:绕排_C622/@绕排方式_C623">
            <xsl:value-of select="./uof:绕排_C622/@绕排方式_C623"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'default'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$method='tight'">
          <wp:wrapTight>
            <xsl:attribute name="wrapText">
              <xsl:if test="uof:绕排_C622/@环绕文字_C624">
                <xsl:call-template name="wraptext">
                  <xsl:with-param name="val" select="uof:绕排_C622/@环绕文字_C624"/>
                </xsl:call-template>
              </xsl:if>
              <!--<xsl:if test="not(uof:绕排_C622/@环绕文字_C624)">
                <xsl:value-of select="'bothSides'"/>
              </xsl:if>-->
            </xsl:attribute>
            <wp:wrapPolygon edited="0">
              <wp:start x="10294" y="0"/>
              <wp:lineTo x="-169" y="21375"/>
              <wp:lineTo x="21600" y="21375"/>
              <wp:lineTo x="21600" y="21150"/>
              <wp:lineTo x="11138" y="0"/>
              <wp:lineTo x="10294" y="0"/>
            </wp:wrapPolygon>
          </wp:wrapTight>
        </xsl:when>
        <xsl:when test="$method='top-bottom'">
          <wp:wrapTopAndBottom/>
        </xsl:when>
        <xsl:when test="$method='through'">
          <wp:wrapThrough>
            <xsl:attribute name="wrapText">
              <xsl:choose>
                <xsl:when test="uof:绕排_C622/@环绕文字_C624">
                  <xsl:call-template name="wraptext">
                    <xsl:with-param name="val" select="uof:绕排_C622/@环绕文字_C624"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'bothSides'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <wp:wrapPolygon edited="0">
              <wp:start x="-338" y="0"/>
              <wp:lineTo x="-338" y="21375"/>
              <wp:lineTo x="21769" y="21375"/>
              <wp:lineTo x="21769" y="0"/>
              <wp:lineTo x="-338" y="0"/>
            </wp:wrapPolygon>
          </wp:wrapThrough>
        </xsl:when>
        <xsl:when test="$method='infront-of-text'">
          <wp:wrapNone/>
        </xsl:when>
        <xsl:when test="$method='behind-text'">
          <wp:wrapNone/>
        </xsl:when>
        <xsl:when test="$method='square'">
          <wp:wrapSquare>
            <xsl:attribute name="wrapText">
              <xsl:choose>
                <xsl:when test="uof:绕排_C622/@环绕文字_C624">
                  <xsl:call-template name="wraptext">
                    <xsl:with-param name="val" select="uof:绕排_C622/@环绕文字_C624"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'bothSides'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </wp:wrapSquare>
        </xsl:when>
        <!--<xsl:otherwise><wp:wrapNone/></xsl:otherwise>-->
      </xsl:choose>
    </xsl:if>
    <wp:docPr>
      <xsl:attribute name="id">

        <!--2014-04-09，Smiley Face的id取值需特殊处理，不然会带来文档需要修复的问题，start-->
        <xsl:choose>
          <xsl:when test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A = 'Smiley Face'">
            <xsl:value-of select="number(substring-after($ObjId,'Obj')) + 1"/>
          </xsl:when>

          <!--2014-05-20，wudi，修复页眉中插入图片时，转换后文档需要修复的BUG，start-->
          <!--2014-05-27，wudi，修复页眉中插入图片时，转换后文档需要修复的BUG，将+2改成+4，start-->
          <xsl:when test="starts-with($ObjId,'header')">
            <xsl:value-of select="number(substring-after($ObjId,'Obj')) + 4"/>
          </xsl:when>
          <!--end-->
          <!--end-->
          
          <xsl:otherwise>
            <xsl:value-of select="substring-after($ObjId,'Obj')"/>
          </xsl:otherwise>
        </xsl:choose>
        <!--end-->
        
        <!--<xsl:apply-templates select="//uof:其他对象[@uof:标识符 = $DataObjId]" mode="num"/>-->
      </xsl:attribute>
      <!--考虑组合图形的情况******************************************************-->

      <xsl:attribute name="name">
        <xsl:choose>
          <xsl:when test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/@组合列表_8064">

            <!--2013-05-03，wudi，修复组合图形，流程图，SmartArt互操作BUG，三种“组合图形”用不同的标识符，start-->
            <xsl:choose>
              <xsl:when test ="contains(//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/@组合列表_8064,'grpspObj')">
                <xsl:value-of select ="'组合'"/>
              </xsl:when>
              <xsl:when test ="contains(//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/@组合列表_8064,'GrpspObj')">
                <xsl:value-of select="'Group'"/>
              </xsl:when>
              <xsl:when test ="contains(//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/@组合列表_8064,'SmartArtObj')">
                <xsl:value-of select ="'图示'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="'组合'"/>
              </xsl:otherwise>
            </xsl:choose>
            <!--end-->
            
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </wp:docPr>
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks>
        <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:缩放是否锁定纵横比_8055">
          <xsl:attribute name="noChangeAspect">
            <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:缩放是否锁定纵横比_8055"/>
          </xsl:attribute>
        </xsl:if>
      </a:graphicFrameLocks>
    </wp:cNvGraphicFramePr>
    <a:graphic>
      <xsl:choose>
        <xsl:when test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/@组合列表_8064">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
            <wpg:wgp>

              <!--2013-05-02，wudi，增加参数isSmartArt，标识是否是SmartArt转换，目前SmartArt按组合图形转，start-->

              <!--2014-03-27，wudi，增加变量isSmartArt，用于判断组合图形是否是SmartArt，start-->
              <xsl:variable name="isSmartArt">
                <xsl:choose>
                  <xsl:when test="contains(//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/@组合列表_8064,'SmartArtObj')">
                    <xsl:value-of select="'1'"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:variable>
              <xsl:apply-templates select="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId and @组合列表_8064]"
                mode="wgpprstGeom">
                <xsl:with-param name="num" select="$pos"/>
                <xsl:with-param name ="tag" select ="'1'"/>
                <xsl:with-param name ="isSmartArt" select ="$isSmartArt"/>
                <xsl:with-param name ="anchorx" select ="'0'"/>
                <xsl:with-param name ="anchory" select ="'0'"/>
              </xsl:apply-templates>
              <!--end-->

              <!--end-->

            </wpg:wgp>

          </a:graphicData>
        </xsl:when>
        <xsl:otherwise>
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>

              <!--2014-03-13，wudi，原取值方式有误，修正，去掉‘图:文字排列方向_8042’前的@，条件值为't2b-l2r-270e-0w'，start-->
              <xsl:if test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/图:文本_803C/图:文字排列方向_8042='t2b-l2r-270e-0w'">
                <xsl:attribute name="normalEastAsianFlow">
                  <xsl:value-of select="'1'"/>
                </xsl:attribute>
              </xsl:if>
              <!--end-->

              <xsl:choose>
                <xsl:when test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:类别_8019 = '61'">
                  <wps:cNvCnPr/>
                </xsl:when>
                <xsl:when test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:类别_8019 = '71'">
                  <wps:cNvCnPr/>
                </xsl:when>
                <xsl:otherwise>
                  <wps:cNvSpPr>
                    <xsl:if test="//uof:对象集/图:图形_8062[@标识符_804B=$ObjId]/图:文本_803C">
                      <xsl:attribute name="txBox">
                        <xsl:value-of select="'1'"/>
                      </xsl:attribute>
                    </xsl:if>
                  </wps:cNvSpPr>
                </xsl:otherwise>
              </xsl:choose>
              <wps:spPr>
                <a:xfrm>
                  <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D">
                    <xsl:attribute name="rot">
                      <xsl:value-of select="round(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D * 60000)"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!--cxl,2012.3.17图形翻转-->

                  <!--2013-03-26，wudi，修复预定义图形方向不对的BUG，此差异与永中软件有关，start-->
                  <xsl:if test="(//图:图形_8062[@标识符_804B=$ObjId]/图:翻转_803A='x') and not(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A ='Line Callout4') and not(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A ='Line Callout4(Accent Bar)') and not(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A ='Line Callout4(No Border)') and not(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:名称_801A ='Line Callout4(Border and Accent Bar)')">
                    <xsl:attribute name="flipH">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <!--end-->

                  <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:翻转_803A='y'">
                    <xsl:attribute name="flipV">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <a:off x="0" y="0"/>

                  <!--2013-04-11，wudi，改"../uof:锚点_C644"为"."-->
                  <xsl:apply-templates select="." mode="xfrm"/>
                </a:xfrm>
                <xsl:apply-templates select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018" mode="prstGeom">
                  <xsl:with-param name="num" select="$pos"/>
                </xsl:apply-templates>
              </wps:spPr>
              <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:文本_803C">
                <xsl:apply-templates select="//图:图形_8062[@标识符_804B=$ObjId]/图:文本_803C" mode="wpstxbx"/>
              </xsl:if>
              <!--cxl,2012.2.25新增艺术字转换-->
              <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]//图:艺术字_802D">
                <xsl:apply-templates select="//图:图形_8062[@标识符_804B=$ObjId]//图:艺术字_802D" mode="wpstxbx"/>
              </xsl:if>
              <wps:bodyPr>
                <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:文本_803C">
                  <xsl:apply-templates select="//图:图形_8062[@标识符_804B=$ObjId]/图:文本_803C" mode="txbody">
                    <xsl:with-param name="objectid" select="$ObjId"/>
                  </xsl:apply-templates>
                </xsl:if>
                <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]//图:艺术字_802D">
                  <xsl:apply-templates select="//图:图形_8062[@标识符_804B=$ObjId]//图:艺术字_802D" mode="txbody"/>
                </xsl:if>
              </wps:bodyPr>
              <!-- <xsl:if test ="not(//uof:扩展区/uof:扩展/uof:扩展内容/uof:内容/uof:文本框属性[@uof:对象索引 = $ObjId]/@uof:是否随图形旋转文字)">
              <xsl:attribute name ="upright">
                <xsl:value-of select ="1"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="//uof:扩展区/uof:扩展/uof:扩展内容/uof:内容/uof:文本框属性[@uof:对象索引 = $ObjId]/@uof:是否随图形旋转文字">
              <xsl:attribute name ="upright">
                <xsl:value-of select ="0"/>
              </xsl:attribute>
            </xsl:if>-->

            </wps:wsp>
          </a:graphicData>
        </xsl:otherwise>
      </xsl:choose>
    </a:graphic>
  </xsl:template>

  <!--  <xsl:variable name="ObjId" select="字:图形/@字:图形引用"/>-->
  <!--
          <xsl:if test="//图:图形[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:锁定纵横比">
            <xsl:attribute name="noChangeAspect">
              <xsl:value-of select="//图:图形[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:锁定纵横比"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
      </a:graphicFrameLocks>
    </wp:cNvGraphicFramePr>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr>
              <xsl:attribute name="id">
                <xsl:apply-templates select="//uof:其他对象[@uof:标识符=$dataObjId]" mode="num"/>
              </xsl:attribute>
              <xsl:attribute name="name">
                <xsl:value-of select="//uof:其他对象[@uof:标识符=$dataObjId]/u2opic:picture/@u2opic:target"/>
              </xsl:attribute>
            </pic:cNvPr>
            <pic:cNvPicPr/>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip>
              <xsl:attribute name="r:embed">
                <xsl:value-of select="concat('rIdObj',$pos)"/>
              </xsl:attribute>
            </a:blip>
            <a:stretch/>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm>
              <xsl:if test="$desType='objpic'">
                -->
  <!--<xsl:variable name="ObjId" select="字:图形/@字:图形引用"/>-->
  <!--
                <xsl:if test="//图:图形[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度">
                  <xsl:attribute name="rot">
                    <xsl:value-of select="//图:图形[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度 * 60000"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
              <a:off x="0" y="0"/>
              <xsl:apply-templates select="uof:锚点_C644" mode="xfrm"/>
            </a:xfrm>
            <a:prstGeom prst="rect">
              <a:avLst/>
            </a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </xsl:template>-->

  <xsl:template name="anchorPictureInline">
    <xsl:param name="desType"/>
    <xsl:param name="pos"/>
    <xsl:variable name="dataObjId">
      <xsl:choose>
        <xsl:when test="$desType='datapic'">
          <xsl:value-of select="./@图形引用_C62E"/>
        </xsl:when>
        <xsl:when test="$desType='objpic'">
          <xsl:variable name="temp" select="@图形引用_C62E"/>
          <xsl:value-of select="//图:图形_8062[@标识符_804B=$temp]/图:图片数据引用_8037"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    <xsl:if test="uof:边距_C628">
      <xsl:apply-templates select="./uof:边距_C628"/>
    </xsl:if>
    <xsl:attribute name="simplePos">
      <xsl:value-of select="'0'"/>
    </xsl:attribute>
    <xsl:if test="$desType='objpic'">
      <xsl:variable name="ObjId" select="./@图形引用_C62E"/>
      <xsl:attribute name="relativeHeight">
        <xsl:choose>
          <xsl:when test="//图:图形_8062[@标识符_804B=$ObjId]/@层次_8063">
            <xsl:variable name="arrg">
              <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/@层次_8063"/>
            </xsl:variable>
            <xsl:choose>
              <xsl:when test="$arrg &lt; 0">
                <xsl:value-of select="'251658240'"/>
              </xsl:when>
              <xsl:when test="$arrg &gt; 0">
                <xsl:value-of select="$arrg"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$arrg"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'251658240'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
    <xsl:attribute name="behindDoc">
      <xsl:variable name="roundMethod">
        <xsl:value-of select="./uof:绕排_C622/@绕排方式_C623"/>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$roundMethod='behind-text'">
          <xsl:value-of select="'1'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'0'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="locked">
      <xsl:choose>
        <xsl:when test="./uof:是否锁定_C629">
          <xsl:value-of select="./uof:是否锁定_C629"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'false'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="layoutInCell">0</xsl:attribute>
    <xsl:attribute name="allowOverlap">
      <xsl:choose>
        <xsl:when test="./uof:是否允许重叠_C62B">
          <xsl:value-of select="./uof:是否允许重叠_C62B"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'false'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <wp:simplePos x="0" y="0"/>
    <xsl:if test="./uof:位置_C620">
      <xsl:apply-templates select="./uof:位置_C620" mode="anchorPic"/>
    </xsl:if>
    <xsl:apply-templates select="." mode="pictureProperties"/>
    <xsl:if test="./uof:绕排_C622">
      <xsl:variable name="method">
        <!--cxl,2012.3.17添加默认方式-->
        <xsl:choose>
          <xsl:when test="./uof:绕排_C622/@绕排方式_C623">
            <xsl:value-of select="./uof:绕排_C622/@绕排方式_C623"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'default'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$method='tight'">
          <wp:wrapTight>
            <xsl:attribute name="wrapText">
              <xsl:choose>
                <xsl:when test="uof:绕排_C622/@环绕文字_C624">
                  <xsl:call-template name="wraptext">
                    <xsl:with-param name="val" select="uof:绕排_C622/@环绕文字_C624"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'bothSides'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <wp:wrapPolygon edited="0">
              <wp:start x="10294" y="0"/>
              <wp:lineTo x="-169" y="21375"/>
              <wp:lineTo x="21600" y="21375"/>
              <wp:lineTo x="21600" y="21150"/>
              <wp:lineTo x="11138" y="0"/>
              <wp:lineTo x="10294" y="0"/>
            </wp:wrapPolygon>
          </wp:wrapTight>
        </xsl:when>
        <xsl:when test="$method='top-bottom'">
          <wp:wrapTopAndBottom/>
        </xsl:when>
        <xsl:when test="$method='through'">
          <wp:wrapThrough>
            <xsl:attribute name="wrapText">
              <xsl:choose>
                <xsl:when test="uof:绕排_C622/@环绕文字_C624">
                  <xsl:call-template name="wraptext">
                    <xsl:with-param name="val" select="uof:绕排_C622/@环绕文字_C624"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'bothSides'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <wp:wrapPolygon edited="0">
              <wp:start x="-338" y="0"/>
              <wp:lineTo x="-338" y="21375"/>
              <wp:lineTo x="21769" y="21375"/>
              <wp:lineTo x="21769" y="0"/>
              <wp:lineTo x="-338" y="0"/>
            </wp:wrapPolygon>
          </wp:wrapThrough>
        </xsl:when>
        <xsl:when test="$method='infront-of-text'">
          <wp:wrapNone/>
        </xsl:when>
        <xsl:when test="$method='behind-text'">
          <wp:wrapNone/>
        </xsl:when>
        <xsl:when test="$method='square'">
          <wp:wrapSquare>
            <xsl:attribute name="wrapText">
              <xsl:if test="uof:绕排_C622/@环绕文字_C624">
                <xsl:call-template name="wraptext">
                  <xsl:with-param name="val" select="uof:绕排_C622/@环绕文字_C624"/>
                </xsl:call-template>
              </xsl:if>
              <!--<xsl:if test="not(uof:绕排_C622/@环绕文字_C624)">
                <xsl:value-of select="'bothSides'"/>
              </xsl:if>-->
            </xsl:attribute>
          </wp:wrapSquare>
        </xsl:when>
        <!--<xsl:otherwise><wp:wrapNone/></xsl:otherwise>-->
      </xsl:choose>
    </xsl:if>
    <xsl:variable name="ObjId" select="./@图形引用_C62E"/>

    <!--2014-04-10，wudi，反方向图片的转换，增加对wp:effectExtent的处理，start-->
    <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D">
      <xsl:variable name="prstrot">
        <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D"/>
      </xsl:variable>
      <xsl:variable name="prstheight">
        <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:大小_8060/@宽_C605"/>
      </xsl:variable>
      <xsl:variable name="prstweight">
        <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:大小_8060/@长_C604"/>
      </xsl:variable>
      <xsl:variable name="extentcha">
        <xsl:choose>
          <xsl:when test="$prstheight &gt; $prstweight">
            <xsl:value-of select="round(($prstheight - $prstweight) * 6350)"/>
          </xsl:when>
          <xsl:when test="$prstheight &lt; $prstweight">
            <xsl:value-of select="round(($prstweight - $prstheight) * 6350)"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="0"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <xsl:when
          test="$prstrot='0.0' or $prstrot='90.0' or $prstrot='180.0' or $prstrot='270.0' or $prstrot='360.0'">
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
        </xsl:when>

        <xsl:when test="$prstrot &gt; 0.0 and $prstrot &lt; 45.0">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot = '45.0'">
          <wp:effectExtent t="0" b="0">
            <xsl:attribute name="l">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="r">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot &gt; 45.0 and $prstrot &lt; 135.0">
          <wp:effectExtent t="0" b="0">
            <xsl:attribute name="l">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="r">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot = '135.0'">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot &gt; 135.0 and $prstrot &lt; 225.0">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot = '225.0'">
          <wp:effectExtent t="0" b="0">
            <xsl:attribute name="l">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="r">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot &gt; 225.0 and $prstrot &lt; 315.0  ">
          <wp:effectExtent t="0" b="0">
            <xsl:attribute name="l">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="r">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot = '315.0'">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
        <xsl:when test="$prstrot &gt; 315.0 and $prstrot &lt; 360.0  ">
          <wp:effectExtent l="0" r="0">
            <xsl:attribute name="t">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
            <xsl:attribute name="b">
              <xsl:value-of select="$extentcha"/>
            </xsl:attribute>
          </wp:effectExtent>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <!--end-->
    
    <wp:docPr>
      <xsl:attribute name="id">
        <xsl:value-of select="substring-after($ObjId,'Obj')"/>
        <!--<xsl:apply-templates select="//uof:其他对象[@uof:标识符=$dataObjId]" mode="num"/>-->
      </xsl:attribute>
      <xsl:attribute name="name">
        <xsl:value-of select="concat('图片',$dataObjId)"/>
      </xsl:attribute>
    </wp:docPr>
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks>
        <xsl:if test="$desType='objpic'">
          <!-- <xsl:variable name="ObjId" select="字:图形/@字:图形引用"/>-->
          <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:缩放是否锁定纵横比_8055">
            <xsl:attribute name="noChangeAspect">
              <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:缩放是否锁定纵横比_8055"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
      </a:graphicFrameLocks>
    </wp:cNvGraphicFramePr>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr>
              <xsl:attribute name="id">
                <xsl:apply-templates select="//对象:对象数据_D701[@标识符_D704=$dataObjId]" mode="num"/>
              </xsl:attribute>
              <xsl:attribute name="name">

                <!--2012-01-23，wcz，UOF到OOX图片转换BUG：修改'data\'为'data/'，start-->
                <xsl:choose>
                  <xsl:when test="substring-after(//对象:对象数据_D701[@标识符_D704=$dataObjId]/对象:路径_D703,'data/') != ''">
                    <xsl:value-of select ="substring-after(//对象:对象数据_D701[@标识符_D704=$dataObjId]/对象:路径_D703,'data/')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select ="substring-after(//对象:对象数据_D701[@标识符_D704=$dataObjId]/对象:路径_D703,'data\')"/>
                  </xsl:otherwise>
                </xsl:choose>
                <!--end-->

              </xsl:attribute>
            </pic:cNvPr>
            <pic:cNvPicPr/>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip>
              <xsl:attribute name="r:embed">
                <xsl:value-of select="concat('rIdObj',$pos)"/>
              </xsl:attribute>
            </a:blip>
            <a:stretch/>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm>
              <xsl:if test="$desType='objpic'">
                <!--<xsl:variable name="ObjId" select="字:图形/@字:图形引用"/>-->
                <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D">
                  <xsl:attribute name="rot">
                    <xsl:value-of select="round(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D * 60000)"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
              <a:off x="0" y="0"/>
              <xsl:apply-templates select="." mode="xfrm"/>
            </a:xfrm>

            <!--2014-04-10，wudi，增加反方向图片边框的转换，start-->
            <xsl:apply-templates select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018" mode="prstGeom">
              <xsl:with-param name="num" select="$pos"/>
            </xsl:apply-templates>
            <!--end-->
                        
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </xsl:template>

  <!--2013-04-01，修复UOF到OOX方向图片转换绕排方式转换错误的BUG，start-->
  <xsl:template name ="anchorPictureNormal">
    <xsl:param name="desType"/>
    <xsl:param name="pos"/>
    <xsl:variable name="dataObjId">
      <xsl:choose>
        <xsl:when test="$desType='datapic'">
          <xsl:value-of select="./@图形引用_C62E"/>
        </xsl:when>
        <xsl:when test="$desType='objpic'">
          <xsl:variable name="temp" select="@图形引用_C62E"/>
          <xsl:value-of select="//图:图形_8062[@标识符_804B=$temp]/图:图片数据引用_8037"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    <xsl:if test="uof:边距_C628">
      <xsl:apply-templates select="./uof:边距_C628"/>
    </xsl:if>
    <xsl:attribute name="simplePos">
      <xsl:value-of select="'0'"/>
    </xsl:attribute>
    <xsl:if test="$desType='objpic'">
      <xsl:variable name="ObjId" select="./@图形引用_C62E"/>
      <xsl:attribute name="relativeHeight">
        <xsl:choose>
          <xsl:when test="//图:图形_8062[@标识符_804B=$ObjId]/@层次_8063">
            <xsl:variable name="arrg">
              <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/@层次_8063"/>
            </xsl:variable>
            <xsl:choose>
              <xsl:when test="$arrg &lt; 0">
                <xsl:value-of select="'251658240'"/>
              </xsl:when>
              <xsl:when test="$arrg &gt; 0">
                <xsl:value-of select="$arrg"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$arrg"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:when test="not(//图:图形_8062[@标识符_804B=$ObjId]/@层次_8063)">
            <xsl:value-of select="'251658240'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
    <xsl:attribute name="behindDoc">
      <xsl:variable name="roundMethod">
        <xsl:value-of select="./uof:绕排_C622/@绕排方式_C623"/>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$roundMethod='behind-text'">
          <xsl:value-of select="'1'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'0'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="locked">
      <xsl:choose>
        <xsl:when test="./uof:是否锁定_C629">
          <xsl:value-of select="./uof:是否锁定_C629"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'false'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    
    <!--2013-05-10，wudi，修复图片转换BUG，图片插入到表格中，start-->
    <xsl:attribute name="layoutInCell">
      <xsl:choose>
        <xsl:when test ="ancestor::字:文字表_416C">
          <xsl:value-of select ="'1'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select ="'0'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <!--end-->
    
    <xsl:attribute name="allowOverlap">
      <xsl:choose>
        <xsl:when test="./uof:是否允许重叠_C62B">
          <xsl:value-of select="./uof:是否允许重叠_C62B"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'false'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <wp:simplePos x="0" y="0"/>
    <xsl:if test="./uof:位置_C620">
      <xsl:apply-templates select="./uof:位置_C620" mode="anchorPic"/>
    </xsl:if>
    <xsl:apply-templates select="." mode="pictureProperties"/>
    <xsl:if test="./uof:绕排_C622">
      <xsl:variable name="method">
        <!--cxl,2012.3.17添加默认方式-->
        <xsl:choose>
          <xsl:when test="./uof:绕排_C622/@绕排方式_C623">
            <xsl:value-of select="./uof:绕排_C622/@绕排方式_C623"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'default'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$method='tight'">
          <wp:wrapTight>
            <xsl:attribute name="wrapText">
              <xsl:if test="uof:绕排_C622/@环绕文字_C624">
                <xsl:call-template name="wraptext">
                  <xsl:with-param name="val" select="uof:绕排_C622/@环绕文字_C624"/>
                </xsl:call-template>
              </xsl:if>
              <!--<xsl:if test="not(uof:绕排_C622/@环绕文字_C624)">
                <xsl:value-of select="'bothSides'"/>
              </xsl:if>-->
            </xsl:attribute>
            <wp:wrapPolygon edited="0">
              <wp:start x="10294" y="0"/>
              <wp:lineTo x="-169" y="21375"/>
              <wp:lineTo x="21600" y="21375"/>
              <wp:lineTo x="21600" y="21150"/>
              <wp:lineTo x="11138" y="0"/>
              <wp:lineTo x="10294" y="0"/>
            </wp:wrapPolygon>
          </wp:wrapTight>
        </xsl:when>
        <xsl:when test="$method='top-bottom'">
          <wp:wrapTopAndBottom/>
        </xsl:when>
        <xsl:when test="$method='through'">
          <wp:wrapThrough>
            <xsl:attribute name="wrapText">
              <xsl:choose>
                <xsl:when test="uof:绕排_C622/@环绕文字_C624">
                  <xsl:call-template name="wraptext">
                    <xsl:with-param name="val" select="uof:绕排_C622/@环绕文字_C624"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'bothSides'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <wp:wrapPolygon edited="0">
              <wp:start x="-338" y="0"/>
              <wp:lineTo x="-338" y="21375"/>
              <wp:lineTo x="21769" y="21375"/>
              <wp:lineTo x="21769" y="0"/>
              <wp:lineTo x="-338" y="0"/>
            </wp:wrapPolygon>
          </wp:wrapThrough>
        </xsl:when>
        <xsl:when test="$method='infront-of-text'">
          <wp:wrapNone/>
        </xsl:when>
        <xsl:when test="$method='behind-text'">
          <wp:wrapNone/>
        </xsl:when>
        <xsl:when test="$method='square'">
          <wp:wrapSquare>
            <xsl:attribute name="wrapText">
              <xsl:choose>
                <xsl:when test="uof:绕排_C622/@环绕文字_C624">
                  <xsl:call-template name="wraptext">
                    <xsl:with-param name="val" select="uof:绕排_C622/@环绕文字_C624"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'bothSides'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </wp:wrapSquare>
        </xsl:when>
        <!--<xsl:otherwise><wp:wrapNone/></xsl:otherwise>-->
      </xsl:choose>
    </xsl:if>
    <xsl:variable name="ObjId" select="./@图形引用_C62E"/>
    <wp:docPr>
      <xsl:attribute name="id">
        <xsl:value-of select="substring-after($ObjId,'Obj')"/>
        <!--<xsl:apply-templates select="//uof:其他对象[@uof:标识符=$dataObjId]" mode="num"/>-->
      </xsl:attribute>
      <xsl:attribute name="name">
        <xsl:value-of select="concat('图片',$dataObjId)"/>
      </xsl:attribute>
    </wp:docPr>
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks>
        <xsl:if test="$desType='objpic'">
          <!-- <xsl:variable name="ObjId" select="字:图形/@字:图形引用"/>-->
          <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:缩放是否锁定纵横比_8055">
            <xsl:attribute name="noChangeAspect">
              <xsl:value-of select="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:缩放是否锁定纵横比_8055"/>
            </xsl:attribute>
          </xsl:if>
        </xsl:if>
      </a:graphicFrameLocks>
    </wp:cNvGraphicFramePr>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr>
              <xsl:attribute name="id">
                <xsl:apply-templates select="//对象:对象数据_D701[@标识符_D704=$dataObjId]" mode="num"/>
              </xsl:attribute>
              <xsl:attribute name="name">

                <!--2012-01-23，wcz，UOF到OOX图片转换BUG：修改'data\'为'data/'，start-->
                <xsl:choose>
                  <xsl:when test="substring-after(//对象:对象数据_D701[@标识符_D704=$dataObjId]/对象:路径_D703,'data/') != ''">
                    <xsl:value-of select ="substring-after(//对象:对象数据_D701[@标识符_D704=$dataObjId]/对象:路径_D703,'data/')"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select ="substring-after(//对象:对象数据_D701[@标识符_D704=$dataObjId]/对象:路径_D703,'data\')"/>
                  </xsl:otherwise>
                </xsl:choose>
                <!--end-->

              </xsl:attribute>
            </pic:cNvPr>
            <pic:cNvPicPr/>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip>
              <xsl:attribute name="r:embed">
                <xsl:value-of select="concat('rIdObj',$pos)"/>
              </xsl:attribute>
            </a:blip>
            <a:stretch/>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm>
              <xsl:if test="$desType='objpic'">
                <!--<xsl:variable name="ObjId" select="字:图形/@字:图形引用"/>-->
                <xsl:if test="//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D">
                  <xsl:attribute name="rot">
                    <xsl:value-of select="round(//图:图形_8062[@标识符_804B=$ObjId]/图:预定义图形_8018/图:属性_801D/图:旋转角度_804D * 60000)"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
              <a:off x="0" y="0"/>
              <xsl:apply-templates select="." mode="xfrm"/>
            </a:xfrm>
            <a:prstGeom prst="rect">
              <a:avLst/>
            </a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </xsl:template>
  <!--end-->

  <xsl:template name="wraptext">
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$val='both'">
        <xsl:value-of select="'bothSides'"/>
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

  <xsl:template match="uof:位置_C620" mode="anchorPic">
    <xsl:if test="uof:水平_4106">
      <wp:positionH>
        <xsl:attribute name="relativeFrom">
          <!--cxl,2012.3.16-->
          <xsl:choose>
            <xsl:when test="uof:水平_4106/@相对于_410C!='char'">
              <xsl:value-of select="uof:水平_4106/@相对于_410C"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'character'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test="uof:水平_4106/uof:绝对_4107">
          <wp:posOffset>
            <xsl:call-template name="onethousandpointMeasure">
              <xsl:with-param name="lengthVal" select="uof:水平_4106/uof:绝对_4107/@值_4108"/>
            </xsl:call-template>
          </wp:posOffset>
        </xsl:if>
        <xsl:if test="uof:水平_4106/uof:相对_4109">
          <wp:align>
            <xsl:value-of select="uof:水平_4106/uof:相对_4109/@参考点_410A"/>
          </wp:align>

        </xsl:if>
      </wp:positionH>
    </xsl:if>
    <xsl:if test="uof:垂直_410D">
      <wp:positionV>
        <xsl:attribute name="relativeFrom">
          <xsl:value-of select="uof:垂直_410D/@相对于_C647"/>
        </xsl:attribute>
        <xsl:if test="uof:垂直_410D/uof:绝对_4107">
          <wp:posOffset>
            <xsl:call-template name="onethousandpointMeasure">
              <xsl:with-param name="lengthVal" select="uof:垂直_410D/uof:绝对_4107/@值_4108"/>
            </xsl:call-template>
          </wp:posOffset>
        </xsl:if>
        <xsl:if test="uof:垂直_410D/uof:相对_4109">
          <wp:align>
            <xsl:value-of select="uof:垂直_410D/uof:相对_4109/@参考点_410B"/>
          </wp:align>
        </xsl:if>
      </wp:positionV>
    </xsl:if>
  </xsl:template>

  <xsl:template match="uof:锚点_C644" mode="pictureProperties">
    <xsl:if test="./uof:大小_C621">
      <wp:extent>
        <xsl:call-template name="widthAndHeight"/>
      </wp:extent>
    </xsl:if>
  </xsl:template>
  <xsl:template match="uof:锚点_C644" mode="xfrm">
    <xsl:if test="./uof:大小_C621">
      <a:ext>
        <xsl:call-template name="widthAndHeight"/>
      </a:ext>
    </xsl:if>
  </xsl:template>
  <xsl:template name="widthAndHeight">
    <xsl:variable name="width">
      <!--cxl,2012.4.10修改，可能没有这个属性-->
      <xsl:choose>
        <xsl:when test="./uof:大小_C621/@宽_C605">
          <xsl:value-of select="./uof:大小_C621/@宽_C605"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'E'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="height">
      <xsl:choose>
        <xsl:when test="./uof:大小_C621/@长_C604">
          <xsl:value-of select="./uof:大小_C621/@长_C604"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'E'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:attribute name="cx">
      <xsl:call-template name="onethousandpointMeasure">
        <!--common.xsl-->
        <xsl:with-param name="lengthVal" select="$width"/>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="cy">
      <xsl:call-template name="onethousandpointMeasure">
        <xsl:with-param name="lengthVal" select="$height"/>
      </xsl:call-template>
    </xsl:attribute>
  </xsl:template>

  <!--2014-06-04，wudi，修改转换比例，有变化，start-->
  <xsl:template match="uof:边距_C628">
    <xsl:attribute name="distT">
      <xsl:choose>
        <xsl:when test="@上_C609">
          <xsl:variable name="tmp">
            <xsl:value-of select="@上_C609"/>
          </xsl:variable>
          <xsl:value-of select="round(number($tmp) * 12500)"/>
          <!--2013-04-01，修复图片转换绕排方式转换错误的BUG，修改调用模板，原是nineThousandpointMeasure-->
          <!--
        <xsl:call-template name="onethousandpointMeasure">
          <xsl:with-param name="lengthVal" select="@上_C609"/>
        </xsl:call-template>-->
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'0'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="distB">
      <xsl:choose>
        <xsl:when test="@下_C60B">
          <xsl:variable name="tmp">
            <xsl:value-of select="@下_C60B"/>
          </xsl:variable>
          <xsl:value-of select="round(number($tmp) * 12500)"/>
          <!--2013-04-01，修复图片转换绕排方式转换错误的BUG，修改调用模板，原是nineThousandpointMeasure-->
          <!--
        <xsl:call-template name="onethousandpointMeasure">
          <xsl:with-param name="lengthVal" select="@下_C60B"/>
        </xsl:call-template>-->
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'0'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="distL">
      <xsl:choose>
        <xsl:when test="@左_C608">
          <xsl:variable name="tmp">
            <xsl:value-of select="@左_C608"/>
          </xsl:variable>
          <xsl:value-of select="round(number($tmp) * 12500)"/>
          <!--2013-04-01，修复图片转换绕排方式转换错误的BUG，修改调用模板，原是nineThousandpointMeasure-->
          <!--
        <xsl:call-template name="onethousandpointMeasure">
          <xsl:with-param name="lengthVal" select="@左_C608"/>
        </xsl:call-template>-->
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'0'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="distR">
      <xsl:choose>
        <xsl:when test="@右_C60A">
          <xsl:variable name="tmp">
            <xsl:value-of select="@右_C60A"/>
          </xsl:variable>
          <xsl:value-of select="round(number($tmp) * 12500)"/>
          <!--2013-04-01，修复图片转换绕排方式转换错误的BUG，修改调用模板，原是nineThousandpointMeasure-->
          <!--
        <xsl:call-template name="onethousandpointMeasure">
          <xsl:with-param name="lengthVal" select="@右_C60A"/>
        </xsl:call-template>-->
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'0'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>
  <!--end-->

  <xsl:template match="图:预定义图形_8018" mode="prstGeom">
    <xsl:param name="num"/>
    <a:prstGeom>
      <xsl:attribute name="prst">
        <xsl:choose>
          <xsl:when test=".//图:类别_8019|.//图:名称_801A">
            <xsl:call-template name="prstName"/>
          </xsl:when>
          <xsl:otherwise>rect</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <a:avLst>
        <!--<xsl:if
          test=".//图:名称_801A='Elbow Connector' or .//图:名称_801A='Elbow Arrow Connector' or .//图:名称_801A='Curved Arrow Connector' or .//图:名称_801A='Elbow Double-Arrow Connector' or .//图:名称_801A='Curved Double-Arrow Connector' or .//图:名称_801A='Curved Connector'">
          <a:gd name="adj1" fmla="val 50000"/>
        </xsl:if>-->
        
        <!--2014-02-24，wudi，预定义图形-控制点的转换，start-->
        <xsl:choose>
          <!--椭圆形标注，圆角矩形标注，矩形标注，云形标注-->
          <xsl:when test="(.//图:名称_801A='Oval Callout' or .//图:名称_801A='Rounded Rectangular Callout' or .//图:名称_801A='Rectangular Callout' or .//图:名称_801A='Cloud Callout') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="YC607">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round(($XC606 - 10800) * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round(($YC607 - 10800) * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <xsl:if test=".//图:名称_801A='Rounded Rectangular Callout'">
              <a:gd name="adj3">
                <xsl:attribute name="fmla">
                  <xsl:value-of select="'val 16667'"/>
                </xsl:attribute>
              </a:gd>
            </xsl:if>
          </xsl:when>
          <!--线形标注1，线形标注1（带强调线），线形标注1（无边框线），线形标注1（带边框和强调线）-->
          <xsl:when test="(.//图:名称_801A='Line Callout2' or .//图:名称_801A='Line Callout2(Accent Bar)' or .//图:名称_801A='Line Callout2(No Border)' or .//图:名称_801A='Line Callout2(Border and Accent Bar)') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y2C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round($Y2C607 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round($Y1C607 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj4">
              <xsl:value-of select="concat('val',' ',round($X1C606 * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj4">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj4"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--线形标注2，线形标注2（带强调线），线形标注2（无边框线），线形标注2（带边框和强调线）-->
          <xsl:when test="(.//图:名称_801A='Line Callout3' or .//图:名称_801A='Line Callout3(Accent Bar)' or .//图:名称_801A='Line Callout3(No Border)' or .//图:名称_801A='Line Callout3(Border and Accent Bar)') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y2C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X3C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[3]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y3C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[3]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round($Y3C607 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round($X3C606 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round($Y2C607 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj4">
              <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj5">
              <xsl:value-of select="concat('val',' ',round($Y1C607 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj6">
              <xsl:value-of select="concat('val',' ',round($X1C606 * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj4">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj4"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj5">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj5"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj6">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj6"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--线形标注3，线形标注3（带强调线），线形标注3（无边框线），线形标注3（带边框和强调线）-->
          <xsl:when test="(.//图:名称_801A='Line Callout4' or .//图:名称_801A='Line Callout4(Accent Bar)' or .//图:名称_801A='Line Callout4(No Border)' or .//图:名称_801A='Line Callout4(Border and Accent Bar)') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y2C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X3C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[3]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y3C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[3]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X4C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[4]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y4C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[4]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round($Y4C607 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round($X4C606 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round($Y3C607 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj4">
              <xsl:value-of select="concat('val',' ',round($X3C606 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj5">
              <xsl:value-of select="concat('val',' ',round($Y2C607 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj6">
              <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj7">
              <xsl:value-of select="concat('val',' ',round($Y1C607 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj8">
              <xsl:value-of select="concat('val',' ',round($X1C606 * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj4">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj4"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj5">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj5"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj6">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj6"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj7">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj7"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj8">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj8"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--十字星，八角星，十六角星，二十四角星，三十二角星-->
          <!--五角星OOX有控制点，UOF无控制点-->
          <!--六角星，七角星转成八角星，十角星，十二角星转成十六角星-->
          <xsl:when test ="(.//图:名称_801A='4-Point Star' or .//图:名称_801A='8-Point Star' or .//图:名称_801A='16-Point Star' or .//图:名称_801A='24-Point Star' or .//图:名称_801A='32-Point Star') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="adj">
              <xsl:value-of select="concat('val',' ',round((10800 - $XC606) * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--上凸带形-->
          <xsl:when test="(.//图:名称_801A='Up Ribbon') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="YC607">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((21600 - $YC607) * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((10800 - $XC606) * 200000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--前凸带形-->
          <xsl:when test="(.//图:名称_801A='Down Ribbon') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="YC607">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round($YC607 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((10800 - $XC606) * 200000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--上凸弯带形-->
          <xsl:when test="(.//图:名称_801A='Curved Up Ribbon') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((21600 - $Y1C607) * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:choose>
                <xsl:when test ="$X1C606 = -13500">
                  <xsl:value-of select="100000"/>
                </xsl:when>
                <xsl:when test ="$X1C606 = -8099.568">
                  <xsl:value-of select ="25000"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="concat('val',' ',round((10800 - $X1C606) * 200000 div 21600))"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round(($X2C606 - 675) * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--前凸弯带形-->
          <xsl:when test="(.//图:名称_801A='Curved Down Ribbon') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round($Y1C607 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((10800 - $X1C606) * 200000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round((21600 - $X2C606 - 675) * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--竖卷形，横卷形-->
          <xsl:when test ="(.//图:名称_801A='Vertical Scroll' or .//图:名称_801A='Horizontal Scroll') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="adj">
              <xsl:value-of select="concat('val',' ',round($XC606 * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--波形，双波形-->
          <xsl:when test="(.//图:名称_801A='Wave' or .//图:名称_801A='Double Wave') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="YC607">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round($XC606 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round(($YC607 - 10800) * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--右箭头，虚尾箭头，燕尾形箭头-->
          <xsl:when test="(.//图:名称_801A='Right Arrow' or .//图:名称_801A='Striped Right Arrow' or .//图:名称_801A='Notched Right Arrow') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="YC607">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((10800 - $YC607) * 100000 div 10800))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((21600 - $XC606) * 100000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--左箭头，左右箭头-->
          <xsl:when test="(.//图:名称_801A='Left Arrow' or .//图:名称_801A='Left-Right Arrow') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="YC607">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((10800 - $YC607) * 100000 div 10800))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round($XC606 * 100000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--上箭头-->
          <xsl:when test="(.//图:名称_801A='Up Arrow') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="YC607">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round($XC606 * 100000 * $height div 21600 div $width))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((10800 - $YC607) * 100000 div 10800))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--上下箭头-->
          <xsl:when test="(.//图:名称_801A='Up-Down Arrow') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="YC607">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((10800 - $XC606) * 100000 div 10800))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round($YC607 * 100000 * $height div 21600 div $width))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--下箭头-->
          <xsl:when test="(.//图:名称_801A='Down Arrow') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="YC607">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((21600 - $XC606) * 100000 * $height div 21600 div $width))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((10800 - $YC607) * 100000 div 10800))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--十字箭头，丁字箭头-->
          <xsl:when test="(.//图:名称_801A='Quad Arrow' or .//图:名称_801A='Left-Right-Up Arrow') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 div 10800))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((10800 - $X1C606) * 50000 div 10800))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round($X2C606 * 50000 div 10800))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--圆角右箭头，两边有差异-->
          <!--直角双向箭头，直角上箭头-->
          <xsl:when test="(.//图:名称_801A='Left-Up Arrow' or .//图:名称_801A='Bent-Up Arrow') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round(($Y1C607 - 10800 - $X1C606 div 2) * 200000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((21600 - $X1C606) * 50000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round($X2C606 * 50000 div 10800))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--左弧形箭头，显示效果有差异-->
          <xsl:when test="(.//图:名称_801A='Curved Right Arrow') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round(($Y1C607 - 10800 - 125 - $X1C606 div 2) * 200000 * $height div 21600 div $width))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((21600 + 250 - $X1C606) * 100000 * $height div 21600 div $width))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round((21600 - $X2C606) * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--右弧形箭头，显示效果有差异-->
          <xsl:when test="(.//图:名称_801A='Curved Left Arrow') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round(($Y1C607 - 10800 - 125 - $X1C606 div 2) * 200000 * $height div 21600 div $width))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((21600 + 250 - $X1C606) * 100000 * $height div 21600 div $width))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--下弧形箭头，显示效果有差异-->
          <xsl:when test="(.//图:名称_801A='Curved Up Arrow') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round(($Y1C607 - 10800 - 5 - $X1C606 div 2) * 200000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((21600 + 10 - $X1C606) * 100000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--上弧形箭头，显示效果有差异-->
          <xsl:when test="(.//图:名称_801A='Curved Down Arrow') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round(($Y1C607 - 10800 - 5 - $X1C606 div 2) * 200000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((21600 + 10 - $X1C606) * 100000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round((21600 - $X2C606) * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--五边形，燕尾形-->
          <xsl:when test ="(.//图:名称_801A='Pentagon Arrow' or .//图:名称_801A='Chevron Arrow') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj">
              <xsl:value-of select="concat('val',' ',round((21600 - $XC606) * 100000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <a:gd name="adj">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--右箭头标注-->
          <xsl:when test="(.//图:名称_801A='Right Arrow Callout') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y2C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y2C607) * 200000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round((21600 - $X2C606) * 100000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <xsl:variable name="adj4">
              <xsl:value-of select="concat('val',' ',round($X1C606 * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj4">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj4"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--左箭头标注-->
          <xsl:when test="(.//图:名称_801A='Left Arrow Callout') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y2C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y2C607) * 200000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <xsl:variable name="adj4">
              <xsl:value-of select="concat('val',' ',round((21600 - $X1C606) * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj4">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj4"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--上箭头标注-->
          <xsl:when test="(.//图:名称_801A='Up Arrow Callout') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y2C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y2C607) * 200000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 * $height div 21600 div $width))"/>
            </xsl:variable>
            <xsl:variable name="adj4">
              <xsl:value-of select="concat('val',' ',round((21600 - $X1C606) * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj4">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj4"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--下箭头标注-->
          <xsl:when test="(.//图:名称_801A='Down Arrow Callout') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y2C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y2C607) * 200000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round((21600 - $X2C606) * 100000 * $height div 21600 div $width))"/>
            </xsl:variable>
            <xsl:variable name="adj4">
              <xsl:value-of select="concat('val',' ',round($X1C606 * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj4">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj4"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--左右箭头标注-->
          <xsl:when test="(.//图:名称_801A='Left-Right Arrow Callout') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y2C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y2C607) * 200000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <xsl:variable name="adj4">
              <xsl:value-of select="concat('val',' ',round((10800 - $X1C606) * 200000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj4">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj4"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--十字箭头标注-->
          <xsl:when test="(.//图:名称_801A='Quad Arrow Callout') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="X1C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y1C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[1]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="X2C606">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="Y2C607">
              <xsl:value-of select="ancestor::图:图形_8062/child::图:控制点_8039[2]/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y2C607) * 200000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round((10800 - $Y1C607) * 100000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <xsl:variable name="adj3">
              <xsl:value-of select="concat('val',' ',round($X2C606 * 100000 div 21600))"/>
            </xsl:variable>
            <xsl:variable name="adj4">
              <xsl:value-of select="concat('val',' ',round((10800 - $X1C606) * 200000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj3">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj3"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj4">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj4"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--等腰三角形，八边形，十字型，缺角矩形，立方体，棱台，太阳形，新月形，双括号，双大括号-->
          <xsl:when test ="(.//图:名称_801A='Isosceles Triangle' or .//图:名称_801A='Octagon' or .//图:名称_801A='Cross' or .//图:名称_801A='Plaque' or .//图:名称_801A='Cube' or .//图:名称_801A='Bevel' or .//图:名称_801A='Sun' or .//图:名称_801A='Moon' or .//图:名称_801A='Double Bracket' or .//图:名称_801A='Double Brace') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="adj">
              <xsl:value-of select="concat('val',' ',round($XC606 * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--平行四边形，梯形，六边形，同心圆，禁止符-->
          <xsl:when test ="(.//图:名称_801A='Parallelogram' or .//图:名称_801A='Trapezoid' or .//图:名称_801A='Hexagon' or .//图:名称_801A='Donut' or .//图:名称_801A='No Symbol') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj">
              <xsl:value-of select="concat('val',' ',round($XC606 * 100000 * $width div 21600 div $height))"/>
            </xsl:variable>
            <a:gd name="adj">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj"/>
              </xsl:attribute>
            </a:gd>
            
            <!--2015-03-19，wudi，修复六边形控制点反方向转换BUG，start-->
            <xsl:if test=".//图:名称_801A='Hexagon'">
              <a:gd name="vf" fmla="val 115470"/>
            </xsl:if>
            <!--end-->
            
          </xsl:when>
          <!--圆柱形-->
          <xsl:when test ="(.//图:名称_801A='Can') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="adj">
              <xsl:value-of select="concat('val',' ',round($XC606 * 200000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--折角形-->
          <xsl:when test ="(.//图:名称_801A='Folded Corner') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="adj">
              <xsl:value-of select="concat('val',' ',round((21600 - $XC606) * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--笑脸-->
          <xsl:when test ="(.//图:名称_801A='Smiley Face') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="adj">
              <xsl:value-of select="concat('val',' ',round(($XC606 - 16200) * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--左中括号，右中括号-->
          <xsl:when test ="(.//图:名称_801A='Left Bracket' or .//图:名称_801A='Right Bracket') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj">
              <xsl:value-of select="concat('val',' ',round($XC606 * 100000 * $height div 21600 div $width))"/>
            </xsl:variable>
            <a:gd name="adj">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--左大括号，右大括号-->
          <xsl:when test ="(.//图:名称_801A='Left Bracket' or .//图:名称_801A='Right Bracket') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="YC607">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@y_C607"/>
            </xsl:variable>
            <xsl:variable name="width">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@宽_C605"/>
            </xsl:variable>
            <xsl:variable name="height">
              <xsl:value-of select=".//图:属性_801D/图:大小_8060/@长_C604"/>
            </xsl:variable>
            <xsl:variable name="adj1">
              <xsl:value-of select="concat('val',' ',round($XC606 * 100000 * $height div 21600 div $width))"/>
            </xsl:variable>
            <xsl:variable name="adj2">
              <xsl:value-of select="concat('val',' ',round($XC606 * 100000 div 21600))"/>
            </xsl:variable>
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj1"/>
              </xsl:attribute>
            </a:gd>
            <a:gd name="adj2">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj2"/>
              </xsl:attribute>
            </a:gd>
          </xsl:when>
          <!--肘形连接符，肘形箭头连接符，肘形双箭头连接符，曲线连接符，曲线箭头连接符，曲线双箭头连接符-->
          <xsl:when test ="(.//图:名称_801A='Elbow Connector' or .//图:名称_801A='Curved Connector') and ancestor::图:图形_8062/图:控制点_8039">
            <xsl:variable name="XC606">
              <xsl:value-of select="ancestor::图:图形_8062/图:控制点_8039/@x_C606"/>
            </xsl:variable>
            <xsl:variable name="adj">
              <xsl:value-of select="concat('val',' ',round($XC606 * 100000 div 21600))"/>
            </xsl:variable>
            
            <!--2015-03-19，wudi，不是adj，应该是adj1，否则，转后文档打不开，start-->
            <a:gd name="adj1">
              <xsl:attribute name="fmla">
                <xsl:value-of select="$adj"/>
              </xsl:attribute>
            </a:gd>
            <!--end-->
            
          </xsl:when>
        </xsl:choose>
        <!--end-->
        
      </a:avLst>
      <!-- <xsl:if test=".//图:阴影">
        <xsl:call-template name="shadow"/>
      </xsl:if>-->
    </a:prstGeom>
    <!--cxl,2012.2.26修改艺术字所在预定义图形的填充为“无填充”，线条为“无线条”-->
    <xsl:if test=".//图:填充_804C and not(./图:属性_801D/图:艺术字_802D)">
      <xsl:call-template name="fill">
        <xsl:with-param name="fillpos" select="$num"/>
      </xsl:call-template>
    </xsl:if>
    <xsl:if test=".//图:填充_804C and ./图:属性_801D/图:艺术字_802D">
      <a:noFill/>
    </xsl:if>
    <xsl:if test="(.//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F) 
            and not(./图:属性_801D/图:艺术字_802D)">
      <xsl:call-template name="ln"/>
    </xsl:if>
    <xsl:if test="(.//图:线颜色_8058|.//图:线类型_8059|.//图:线粗细_805C|.//图:前端箭头_805E|.//图:后端箭头_805F) 
            and (./图:属性_801D/图:艺术字_802D)">
      <a:ln>
        <a:noFill/>
      </a:ln>
    </xsl:if>
    <!--cxl,2012.2.21新增三维效果转换-->
    <xsl:if test=".//图:三维效果_8061">
      <xsl:apply-templates select=".//图:三维效果_8061"/>
    </xsl:if>
    <!--cxl,2012.2.24新增阴影效果转换-->
    <!--如果是艺术字的阴影效果，则不在此调用这个模板-->
    <xsl:if test=".//图:阴影_8051 and not(./图:属性_801D/图:艺术字_802D)">
      <xsl:apply-templates select=".//图:阴影_8051"/>
    </xsl:if>
  </xsl:template>

  <!--组合图形**************************************************************-->
  <xsl:template match="图:图形_8062" mode="wgpprstGeom">

    <!--2013-05-02，wudi，增加参数isSmartArt，标识是否是SmartArt转换，目前SmartArt按组合图形转，start-->
    <xsl:param name="num"/>
    <xsl:param name ="tag"/>
    <xsl:param name ="isSmartArt"/>
    <xsl:param name ="anchorx"/>
    <xsl:param name ="anchory"/>
    <!--end-->
    
    <xsl:variable name="list">
      <xsl:value-of select="@组合列表_8064"/>
    </xsl:variable>
    <xsl:variable name="id">
      <xsl:value-of select="@标识符_804B"/>
    </xsl:variable>

    <xsl:variable name ="wpganchorx">

      <xsl:choose>
        <xsl:when test ="$tag='1'">
          <xsl:for-each select ="//uof:对象集/图:图形_8062">
            <xsl:sort select ="图:组合位置_803B/@x_C606" data-type ="number" order="ascending"/>
          </xsl:for-each>
          <xsl:value-of select ="//uof:对象集/图:图形_8062[contains($list,@标识符_804B)]/图:组合位置_803B/@x_C606"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select ="$anchorx"/>
        </xsl:otherwise>
      </xsl:choose>
      
    </xsl:variable>
    <xsl:variable name ="wpganchory">
      <xsl:choose>
        <xsl:when test ="$tag='1'">
          <xsl:for-each select ="//uof:对象集/图:图形_8062">

            <xsl:sort select ="图:组合位置_803B/@y_C607" data-type ="number" order="ascending"/>
          </xsl:for-each>
          <xsl:value-of select ="//uof:对象集/图:图形_8062[contains($list,@标识符_804B)]/图:组合位置_803B/@y_C607"/>

        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select ="$anchory"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:if test ="$tag!='1'">
      <wpg:cNvPr>
        <xsl:variable name ="wpgid">
          <xsl:value-of select ="substring-after(@标识符_804B,'Obj')"/>
        </xsl:variable>
        <xsl:attribute name ="id">
          <xsl:value-of select ="$num*100+$tag*10+$wpgid"/>
        </xsl:attribute>
        <xsl:attribute name ="name">
          <xsl:value-of select ="concat('组合',$num*100+$tag*10+$wpgid)"/>
        </xsl:attribute>
      </wpg:cNvPr>
    </xsl:if>
    <wpg:cNvGrpSpPr/>
    <wpg:grpSpPr>
      
      <!--2013-04-12，wudi，修复组合图形转换的BUG，start-->

      <!--2013-05-02，wudi，增加参数isSmartArt，标识是否是SmartArt转换，目前SmartArt按组合图形转，start-->
      
      <!--2013-05-03，wudi，增加Label参数，标识组合图形是否是嵌套组合，start-->
      <xsl:call-template name="grpXfrm">
        <xsl:with-param name ="list" select ="$list"/>
        <xsl:with-param name ="isSmartArt" select ="$isSmartArt"/>
        <!--<xsl:with-param name ="Label" select ="$tag"/>-->
      </xsl:call-template>
      <!--end-->
      
      <!--end-->
      
      <!--end-->
      
      <!--<xsl:if test ="$tag='1'">
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </xsl:if>
      <xsl:if test ="$tag !='1'">
        <xsl:call-template name ="grpXfrm"/>
      </xsl:if>-->
    </wpg:grpSpPr>
    <xsl:for-each select="//uof:对象集/图:图形_8062">
      
      <!--2014-04-09，wudi，增加限制条件，原判断条件遇到类似组合列表_8064=" grpspObj266 grpspObj267 grpspObj268"，同时存在grpspObj2标识符的情况会出问题，start-->
      <xsl:variable name="list1">
        <xsl:value-of select="normalize-space($list)"/>
      </xsl:variable>
      <xsl:variable name="str1">
        <xsl:value-of select="substring-before($list1,' ')"/>
      </xsl:variable>
      <xsl:variable name="list2">
        <xsl:value-of select="substring-after($list1,' ')"/>
      </xsl:variable>
      <xsl:variable name="str2">
        <xsl:value-of select="substring-before($list2,' ')"/>
      </xsl:variable>
      
      <xsl:choose>
        <xsl:when test="contains($list,@标识符_804B) and not(@组合列表_8064) and not(contains($str1,@标识符_804B) and contains($str2,@标识符_804B))">
          <xsl:variable name ="grpnum">
            <xsl:value-of select ="substring-after(@标识符_804B,'Obj')"/>
          </xsl:variable>
          <xsl:apply-templates select="." mode="prtshape1">
            <!--number是该图形的标识符，要保证其唯一，有待完善-->
            <xsl:with-param name="number" select="$num*100+$tag*10+$grpnum"/>
            <xsl:with-param name ="anchorx" select ="$wpganchorx"/>
            <xsl:with-param name ="anchory" select ="$wpganchory"/>
          </xsl:apply-templates>
        </xsl:when>
        <xsl:when test="contains($list,@标识符_804B) and @组合列表_8064 and not(contains($str1,@标识符_804B) and contains($str2,@标识符_804B))">
          <wpg:grpSp>

            <!--2013-05-02，wudi，增加参数isSmartArt，标识是否是SmartArt转换，目前SmartArt按组合图形转，start-->
            <xsl:apply-templates select="." mode="wgpprstGeom">
              <xsl:with-param name="num" select="$num"/>
              <xsl:with-param name ="tag" select ="$tag + 1"/>
              <xsl:with-param name ="isSmartArt" select ="$isSmartArt"/>
              <xsl:with-param name ="anchorx" select ="$wpganchorx"/>
              <xsl:with-param name ="anchory" select ="$wpganchory"/>
            </xsl:apply-templates>
            <!--end-->
            
          </wpg:grpSp>
        </xsl:when>
      </xsl:choose>
      <!--end-->

    </xsl:for-each>
  </xsl:template>

  <xsl:template match="图:图形_8062" mode="prtshape1">
    <xsl:param name="number"/>
    <xsl:param name ="anchorx"/>
    <xsl:param name ="anchory"/>
    <!--组合图形中的图片填充-->
    <xsl:if test=".//图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图片_8005/@图形引用_8007">
      <xsl:variable name="oid" select=".//图:预定义图形_8018/图:属性_801D/图:填充_804C/图:图片_8005/@图形引用_8007"/>
      <xsl:copy-of select="//对象:对象数据_D701[@标识符_D704 = $oid]/u2opic:picture"/>
      <!--*******************************-->
    </xsl:if>
    <xsl:variable name ="shapeid">
      <xsl:value-of select ="@标识符_804B"/>
    </xsl:variable>
    <!--找对应锚点的位置
   <xsl:variable name ="anchorx">
      <xsl:value-of select ="//字:锚点[字:图形/@字:图形引用 = 'Obj00003']/uof:锚点_C644/字:位置/字:水平/字:绝对/@字:值"/>
    </xsl:variable>
    <xsl:variable name ="anchory">
      <xsl:value-of select ="//字:锚点[字:图形/@字:图形引用 = 'Obj00003']/uof:锚点_C644/字:位置/字:垂直/字:绝对/@字:值"/>
    </xsl:variable>-->
    <wps:wsp>
      <wps:cNvPr>
        <xsl:attribute name ="id">
          <xsl:value-of select ="$number"/>
        </xsl:attribute>
        <xsl:attribute name ="name">
          <xsl:value-of select ="concat('图形',$number)"/>
        </xsl:attribute>
      </wps:cNvPr>
      <wps:cNvCnPr/>
      <wps:spPr>
        <a:xfrm>
          <xsl:if test=".//图:预定义图形_8018/图:属性_801D/图:旋转角度_804D">
            <xsl:attribute name="rot">
              <xsl:value-of select="round(.//图:预定义图形_8018/图:属性_801D/图:旋转角度_804D * 60000)"/>
            </xsl:attribute>
          </xsl:if>
          
          <!--2013-04-12，wudi，修复组合图形-单个图形翻转节点不转的BUG，start-->
          <xsl:if test="(./图:翻转_803A='x') and not(./图:预定义图形_8018/图:名称_801A ='Line Callout4') and not(./图:预定义图形_8018/图:名称_801A ='Line Callout4(Accent Bar)') and not(./图:预定义图形_8018/图:名称_801A ='Line Callout4(No Border)') and not(./图:预定义图形_8018/图:名称_801A ='Line Callout4(Border and Accent Bar)') and not(./图:预定义图形_8018/图:名称_801A ='Rectangular Callout')">
            <xsl:attribute name="flipH">
              <xsl:value-of select="'1'"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="./图:翻转_803A='y'">
            <xsl:attribute name="flipV">
              <xsl:value-of select="'1'"/>
            </xsl:attribute>
          </xsl:if>
          <!--end-->
          
          <a:off>
            
            <xsl:attribute name="x">
              <xsl:choose>
                <xsl:when test="contains(.//图:组合位置_803B/@x_C606,'E')">
                  <xsl:value-of select="0"/>
                </xsl:when>
                <xsl:otherwise>
                  <!--<xsl:value-of select="1"/>-->
                  <!--2013-04-12，wudi，修复组合图形转换的BUG，去掉- $anchorx * 12700-->
                  <xsl:value-of select="round(.//图:组合位置_803B/@x_C606 * 12700)"/>
                  <!--<xsl:value-of select="round(.//图:组合位置_803B/@x_C606 * 12700 - $anchorx * 12700)"/>-->
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="y">
              <xsl:choose>
                <xsl:when test="contains(.//图:组合位置_803B/@y_C607,'E')">
                  <xsl:value-of select="0"/>
                </xsl:when>
                <xsl:otherwise>
                  <!--<xsl:value-of select="1"/>-->
                  <!--2013-04-12，wudi，修复组合图形转换的BUG，去掉- $anchorx * 12700-->
                  <xsl:value-of select="round(.//图:组合位置_803B/@y_C607 * 12700)"/>
                  <!--<xsl:value-of select="round(.//图:组合位置_803B/@y_C607 * 12700 - $anchory * 12700)"/>-->
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </a:off>
          <a:ext>
            <xsl:attribute name="cx">
              <xsl:choose>
                <xsl:when test=".//图:大小_8060/@宽_C605">
                  <xsl:choose>
                    <xsl:when test="contains(.//图:大小_8060/@宽_C605,'E')">
                      <xsl:value-of select="0"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="round(.//图:大小_8060/@宽_C605 * 12700)"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="0"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="cy">
              <xsl:choose>
                <xsl:when test=".//图:大小_8060/@长_C604">
                  <xsl:choose>
                    <xsl:when test="contains(.//图:大小_8060/@长_C604,'E')">
                      <xsl:value-of select="0"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="round(.//图:大小_8060/@长_C604 * 12700)"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="0"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </a:ext>
        </a:xfrm>
        <xsl:apply-templates select=".//图:预定义图形_8018" mode="prstGeom">
          <xsl:with-param name="num" select="$number"/>
        </xsl:apply-templates>
      </wps:spPr>
      <xsl:if test="图:文本_803C">
        <!--原来是图:文本内容-->
        <xsl:apply-templates select="图:文本_803C" mode="wpstxbx"/>
      </xsl:if>
      <wps:bodyPr>
        <xsl:if test="图:文本_803C">
          <xsl:apply-templates select="图:文本_803C" mode="txbody">
            <xsl:with-param name="objectid" select="@标识符_804B"/>
          </xsl:apply-templates>
        </xsl:if>
        <!-- <xsl:if test ="not(//uof:扩展区/uof:扩展/uof:扩展内容/uof:内容/uof:文本框属性[@uof:对象索引 = $ObjId]/@uof:是否随图形旋转文字)">
              <xsl:attribute name ="upright">
                <xsl:value-of select ="1"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="//uof:扩展区/uof:扩展/uof:扩展内容/uof:内容/uof:文本框属性[@uof:对象索引 = $ObjId]/@uof:是否随图形旋转文字">
              <xsl:attribute name ="upright">
                <xsl:value-of select ="0"/>
              </xsl:attribute>
            </xsl:if>-->
      </wps:bodyPr>

    </wps:wsp>
  </xsl:template>
  <!--处理图形填充-->
  <xsl:template name="solidFill">
    <a:solidFill>
      <a:srgbClr>
        <xsl:attribute name="val">
          <xsl:choose>
            <xsl:when test=".//图:属性_801D/图:填充_804C/图:颜色_8004!='auto'">
              <xsl:variable name="RGB1">
                <xsl:value-of select=".//图:属性_801D/图:填充_804C/图:颜色_8004"/>
              </xsl:variable>
              <xsl:value-of select="substring-after($RGB1,'#')"/>
            </xsl:when>
            <xsl:otherwise>ffffff</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test=".//图:属性_801D/图:透明度_8050">
          <a:alpha>
            <xsl:attribute name="val">
              <xsl:value-of select="(100-.//图:属性_801D/图:透明度_8050)*1000"/>
            </xsl:attribute>
          </a:alpha>
        </xsl:if>
      </a:srgbClr>
    </a:solidFill>
  </xsl:template>

  <!--处理线属性-->
  <xsl:template name="ln">
    <a:ln>
      <xsl:if test=".//图:线粗细_805C">
        <xsl:attribute name="w">
          <xsl:value-of select="round(.//图:属性_801D/图:线_8057/图:线粗细_805C *12700)"/>
        </xsl:attribute>
      </xsl:if>
      <!--线型对应-->
      <xsl:choose>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@线型_805A='single'">
          <xsl:attribute name="cmpd">sng</xsl:attribute>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@线型_805A='double'">
          <xsl:attribute name="cmpd">dbl</xsl:attribute>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@线型_805A='thin-thick'">
          <xsl:attribute name="cmpd">thickThin</xsl:attribute>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@线型_805A='thick-thin'">
          <xsl:attribute name="cmpd">thinThick</xsl:attribute>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@线型_805A='thick-between-thin'">
          <xsl:attribute name="cmpd">tri</xsl:attribute>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@线型_805A='none'"> </xsl:when>
      </xsl:choose>
      <!--线颜色-->
      <xsl:if test=".//图:属性_801D/图:线_8057/图:线颜色_8058">
        <a:solidFill>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:choose>
                <xsl:when test=".//图:属性_801D/图:线_8057/图:线颜色_8058!='auto'">
                  <xsl:variable name="RGB">
                    <xsl:value-of select=".//图:属性_801D/图:线_8057/图:线颜色_8058"/>
                  </xsl:variable>
                  <xsl:value-of select="substring-after($RGB,'#')"/>
                </xsl:when>
                <xsl:otherwise>000000</xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
          </a:srgbClr>
        </a:solidFill>
      </xsl:if>

      <xsl:choose>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='solid'">
          <a:prstDash>
            <xsl:attribute name="val">solid</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='round-dot'">
          <a:prstDash>
            <xsl:attribute name="val">sysDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='square-dot'">
          <a:prstDash>
            <xsl:attribute name="val">sysDash</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='dash'">
          <a:prstDash>
            <xsl:attribute name="val">dash</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='dash-dot'">
          <a:prstDash>
            <xsl:attribute name="val">dashDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='long-dash'">
          <a:prstDash>
            <xsl:attribute name="val">lgDash</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='long-dash-dot'">
          <a:prstDash>
            <xsl:attribute name="val">lgDashDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>
        <xsl:when test=".//图:属性_801D/图:线_8057/图:线类型_8059/@虚实_805B='dash-dot-dot'">
          <a:prstDash>
            <xsl:attribute name="val">lgDashDotDot</xsl:attribute>
          </a:prstDash>
        </xsl:when>

      </xsl:choose>
      <xsl:if test=".//图:前端箭头_805E">
        <a:headEnd>
          <xsl:attribute name="type">
            <xsl:choose>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='normal'">triangle</xsl:when>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='diamond'">diamond</xsl:when>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='open'">arrow</xsl:when>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='oval'">oval</xsl:when>
              <xsl:when test=".//图:前端箭头_805E/图:式样_8000='stealth'">stealth</xsl:when>
              <xsl:otherwise>arrow</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </a:headEnd>
      </xsl:if>
      <xsl:if test=".//图:后端箭头_805F">
        <a:tailEnd>
          <xsl:attribute name="type">
            <xsl:choose>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='normal'">triangle</xsl:when>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='diamond'">diamond</xsl:when>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='open'">arrow</xsl:when>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='oval'">oval</xsl:when>
              <xsl:when test=".//图:后端箭头_805F/图:式样_8000='stealth'">stealth</xsl:when>
              <xsl:otherwise>arrow</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </a:tailEnd>
      </xsl:if>
    </a:ln>
  </xsl:template>
  <!--处理旋转-->
  <xsl:template name="rot">
    <a:scene3d>
      <a:camera>
        <xsl:attribute name="prst">orthographicFront</xsl:attribute>
        <a:rot>
          <xsl:attribute name="lat">0</xsl:attribute>
          <xsl:attribute name="lon">0</xsl:attribute>
          <xsl:attribute name="rev">
            <!--4月3日，线的方向判断-->
            <!--xsl:choose>
                  <xsl:when test=".//图:翻转">5400000</xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="round(21600000 - .//图:属性_801D/图:旋转角度*60000)"/>
                  </xsl:otherwise>
                </xsl:choose-->
            <xsl:choose>
              <xsl:when test=".//图:属性_801D/图:旋转角度_804D!='0.0'">
                <!--10.02.02 myx<xsl:value-of select="round(21600000 - .//图:属性_801D/图:旋转角度*60000)"/>-->
                <xsl:value-of select="round(.//图:属性_801D/图:旋转角度_804D*60000)"/>
              </xsl:when>
              <xsl:otherwise>0</xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </a:rot>
      </a:camera>
      <a:lightRig>
        <xsl:attribute name="rig">threePt</xsl:attribute>
        <xsl:attribute name="dir">t</xsl:attribute>
      </a:lightRig>
    </a:scene3d>
    
  </xsl:template>
  <!--style
  <xsl:template name="style">
    <wps:style>
      <a:lnRef idx="1">
        <a:schemeClr val="accent1"/>
      </a:lnRef>
      <a:fillRef idx="0">
        <a:schemeClr val="accent1"/>
      </a:fillRef>
      <a:effectRef idx="0">
        <a:schemeClr val="accent1"/>
      </a:effectRef>
      <a:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </a:fontRef>
    </wps:style>
  </xsl:template>-->
  <xsl:template match="对象:对象数据_D701" mode="num">
    <xsl:number count="对象:对象数据_D701" format="1" level="single"/>
  </xsl:template>

  <xsl:template name="prstName">
    <xsl:choose>
      <!--xsl:when test="图:图形/@图:其他对象 and .//图:名称='Rectangle'">actionButtonForwardNext</xsl:when-->
      <xsl:when test=".//图:名称_801A='Rectangle' or .//图:类别_8019='11'">rect</xsl:when>
      <xsl:when test=".//图:名称_801A='Parallelogram' or .//图:类别_8019='12'">parallelogram</xsl:when>
      <xsl:when test=".//图:名称_801A='Trapezoid' or .//图:类别_8019='13'">trapezoid</xsl:when>
      <xsl:when test=".//图:名称_801A='diamond' or .//图:类别_8019='14'">diamond</xsl:when>
      <xsl:when test=".//图:名称_801A='Rounded Rectangle' or .//图:类别_8019='15'">roundRect</xsl:when>
      <xsl:when test=".//图:名称_801A='Octagon'or .//图:类别_8019='16'">octagon</xsl:when>
      <xsl:when test=".//图:名称_801A='Isosceles Triangle' or .//图:类别_8019='17'">triangle</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Triangle' or .//图:类别_8019='18'">rtTriangle</xsl:when>
      <xsl:when test=".//图:名称_801A='Oval' or .//图:类别_8019='19'">ellipse</xsl:when>

      <!--2014-02-24，wudi，添加对云形标注的处理，start-->
      <xsl:when test=".//图:名称_801A='Cloud Callout' or .//图:类别_8019='54'">cloudCallout</xsl:when>
      <!--end-->
      
      <xsl:when test=".//图:名称_801A='Right Arrow' or .//图:类别_8019='21'">rightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Left Arrow' or .//图:类别_8019='22'">leftArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Up Arrow' or .//图:类别_8019='23'">upArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Down Arrow' or .//图:类别_8019='24'">downArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Left-Right Arrow' or .//图:类别_8019='25'">leftRightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Up-Down Arrow' or .//图:类别_8019='26'">upDownArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Quad Arrow' or .//图:类别_8019='27'">quadArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Left-Right-Up Arrow' or .//图:类别_8019='28'">leftRightUpArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Bent Arrow' or .//图:类别_8019='29'">bentArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Process' or .//图:类别_8019='31'">flowChartProcess</xsl:when>
      <xsl:when test=".//图:名称_801A='Alternate Process' or .//图:类别_8019='32'"
        >flowChartAlternateProcess</xsl:when>
      <xsl:when test=".//图:名称_801A='Decision' or .//图:类别_8019='33'">flowChartDecision</xsl:when>
      <xsl:when test=".//图:名称_801A='Data' or .//图:类别_8019='34'">flowChartInputOutput</xsl:when>
      <xsl:when test=".//图:名称_801A='Predefined Process' or .//图:类别_8019='35'"
        >flowChartPredefinedProcess</xsl:when>
      <xsl:when test=".//图:名称_801A='Internal Storage' or .//图:类别_8019='36'"
        >flowChartInternalStorage</xsl:when>
      <xsl:when test=".//图:名称_801A='Document' or .//图:类别_8019='37'">flowChartDocument</xsl:when>
      <xsl:when test=".//图:名称_801A='Multidocument' or .//图:类别_8019='38'">flowChartMultidocument</xsl:when>
      <xsl:when test=".//图:名称_801A='Terminator' or .//图:类别_8019='39'">flowChartTerminator</xsl:when>
      <xsl:when test=".//图:名称_801A='Explosion 1' or .//图:类别_8019='41'">irregularSeal1</xsl:when>
      <xsl:when test=".//图:名称_801A='Explosion 2' or .//图:类别_8019='42'">irregularSeal2</xsl:when>
      <xsl:when test=".//图:名称_801A='4-Point Star' or .//图:类别_8019='43'">star4</xsl:when>
      <xsl:when test=".//图:名称_801A='5-Point Star' or .//图:类别_8019='44'">star5</xsl:when>
      <xsl:when test=".//图:名称_801A='8-Point Star' or .//图:类别_8019='45'">star8</xsl:when>
      <xsl:when test=".//图:名称_801A='16-Point Star' or .//图:类别_8019='46'">star12</xsl:when>
      <xsl:when test=".//图:名称_801A='24-Point Star' or .//图:类别_8019='47'">star24</xsl:when>
      <xsl:when test=".//图:名称_801A='32-Point Star' or .//图:类别_8019='48'">star32</xsl:when>
      <xsl:when test=".//图:名称_801A='Up Ribbon' or .//图:类别_8019='49'">ribbon2</xsl:when>
      <xsl:when test=".//图:名称_801A='Rectangular Callout' or .//图:类别_8019='51'">wedgeRectCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Rounded Rectangular Callout' or .//图:类别_8019='52'"
        >wedgeRoundRectCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Oval Callout' or .//图:类别_8019='53'">wedgeEllipseCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Parallelogram' or .//图:类别_8019='54'">cloudCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout1(No Border)' or .//图:类别_8019='513'">borderCallout1</xsl:when>
      
      <!--2014-02-24，wudi，修改borderCallout2为borderCallout1，start-->
      <xsl:when test=".//图:名称_801A='Line Callout2' or .//图:类别_8019='56'">borderCallout1</xsl:when>
      <!--end-->
      
      <!--线性标注3-->
      <xsl:when test=".//图:名称_801A='Line Callout3' or .//图:类别_8019='57'">borderCallout2</xsl:when>
      <!---->
      <xsl:when test=".//图:名称_801A='Line Callout4' or .//图:类别_8019='58'">borderCallout3</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout1(Accent Bar)' or .//图:类别_8019='59'">accentCallout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout2(Accent Bar)' or .//图:类别_8019='510'"
        >accentCallout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout3(Accent Bar)' or .//图:类别_8019='511'"
        >accentCallout2</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout4(Accent Bar)' or .//图:类别_8019='512'"
        >accentCallout3</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout1' or .//图:类别_8019='55'">callout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout2(No Border)' or .//图:类别_8019='514'">callout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout3(No Border)' or .//图:类别_8019='515'">callout2</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout4(No Border)' or .//图:类别_8019='516'">callout3</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout1(Border and Accent Bar)' or .//图:类别_8019='517'"
        >accentBorderCallout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout2(Border and Accent Bar)' or .//图:类别_8019='518'"
        >accentBorderCallout1</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout3(Border and Accent Bar)' or .//图:类别_8019='519'"
        >accentBorderCallout2</xsl:when>
      <xsl:when test=".//图:名称_801A='Line Callout4(Border and Accent Bar)' or .//图:类别_8019='520'"
        >accentBorderCallout3</xsl:when>
      <xsl:when test=".//图:名称_801A='Line' or .//图:类别_8019='61'">line</xsl:when>
      <xsl:when test=".//图:名称_801A='Arrow' or .//图:类别_8019='62'">straightConnector1</xsl:when>
      <xsl:when test=".//图:名称_801A='Double Arrow' or .//图:类别_8019='63'">straightConnector1</xsl:when>
      <xsl:when test=".//图:名称_801A='Curve' or .//图:类别_8019='64'">curvedConnector3</xsl:when>

      <xsl:when test=".//图:名称_801A='Straight Connector' or .//图:类别_8019='71'">straightConnector1</xsl:when>
      <xsl:when test=".//图:名称_801A='Elbow Connector' or .//图:类别_8019='74'">bentConnector3</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Connector' or .//图:类别_8019='77'">curvedConnector3</xsl:when>
      <xsl:when test=".//图:名称_801A='Hexagon' or .//图:类别_8019='110'">hexagon</xsl:when>
      <xsl:when test=".//图:名称_801A='Cross' or .//图:类别_8019='111'">plus</xsl:when>
      <xsl:when test=".//图:名称_801A='Regual Pentagon' or .//图:类别_8019='112'">pentagon</xsl:when>
      <xsl:when test=".//图:名称_801A='Can' or .//图:类别_8019='113'">can</xsl:when>
      <xsl:when test=".//图:名称_801A='Cube' or .//图:类别_8019='114'">cube</xsl:when>
      <xsl:when test=".//图:名称_801A='Bevel' or .//图:类别_8019='115'">bevel</xsl:when>
      <xsl:when test=".//图:名称_801A='Folded Corner' or .//图:类别_8019='116'">foldedCorner</xsl:when>
      <xsl:when test=".//图:名称_801A='Smiley Face' or .//图:类别_8019='117'">smileyFace</xsl:when>
      <xsl:when test=".//图:名称_801A='Donut' or .//图:类别_8019='118'">donut</xsl:when>
      <xsl:when test=".//图:名称_801A='No Symbol' or .//图:类别_8019='119'">noSmoking</xsl:when>
      <xsl:when test=".//图:名称_801A='Block Arc' or .//图:类别_8019='120'">blockArc</xsl:when>
      <xsl:when test=".//图:名称_801A='Heart' or .//图:类别_8019='121'">heart</xsl:when>
      <xsl:when test=".//图:名称_801A='Lightning' or .//图:类别_8019='122'">lightningBolt</xsl:when>
      <xsl:when test=".//图:名称_801A='Sun' or .//图:类别_8019='123'">sun</xsl:when>
      <xsl:when test=".//图:名称_801A='Moon' or .//图:类别_8019='124'">moon</xsl:when>
      <xsl:when test=".//图:名称_801A='Arc' or .//图:类别_8019='125'">arc</xsl:when>
      <xsl:when test=".//图:名称_801A='Double Bracket' or .//图:类别_8019='126'">bracketPair</xsl:when>
      <xsl:when test=".//图:名称_801A='Double Brace' or .//图:类别_8019='127'">bracePair</xsl:when>
      <xsl:when test=".//图:名称_801A='Plaque' or .//图:类别_8019='128'">plaque</xsl:when>
      <xsl:when test=".//图:名称_801A='Left Bracket' or .//图:类别_8019='129'">leftBracket</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Bracket' or .//图:类别_8019='130'">rightBracket</xsl:when>
      <xsl:when test=".//图:名称_801A='Left Brace' or .//图:类别_8019='131'">leftBrace</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Brace' or .//图:类别_8019='132'">rightBrace</xsl:when>
      <xsl:when test=".//图:名称_801A='U-Turn Arrow' or .//图:类别_8019='210'">uturnArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Left-Up Arrow' or .//图:类别_8019='211'">leftUpArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Bent-Up Arrow' or .//图:类别_8019='212'">bentUpArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Right Arrow' or .//图:类别_8019='213'">curvedRightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Left Arrow' or .//图:类别_8019='214'">curvedLeftArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Up Arrow' or .//图:类别_8019='215'">curvedUpArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Down Arrow' or .//图:类别_8019='216'">curvedDownArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Striped Right Arrow' or .//图:类别_8019='217'">stripedRightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Notched Right Arrow' or .//图:类别_8019='218'">notchedRightArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Pentagon Arrow' or .//图:类别_8019='219'">homePlate</xsl:when>
      <xsl:when test=".//图:名称_801A='Chevron Arrow' or .//图:类别_8019='220'">chevron</xsl:when>
      <xsl:when test=".//图:名称_801A='Right Arrow Callout' or .//图:类别_8019='221'">rightArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Left Arrow Callout' or .//图:类别_8019='222'">leftArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Up Arrow Callout' or .//图:类别_8019='223'">upArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Down Arrow Callout' or .//图:类别_8019='224'">downArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Left-Right Arrow Callout' or .//图:类别_8019='225'"
        >leftRightArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Up-Down Arrow Callout' or .//图:类别_8019='226'"
        >upDownArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Quad Arrow Callout' or .//图:类别_8019='227'">quadArrowCallout</xsl:when>
      <xsl:when test=".//图:名称_801A='Circular Arrow' or .//图:类别_8019='228'">circularArrow</xsl:when>
      <xsl:when test=".//图:名称_801A='Preparation' or .//图:类别_8019='310'">flowChartPreparation</xsl:when>
      <xsl:when test=".//图:名称_801A='Manual Input' or .//图:类别_8019='311'">flowChartManualInput</xsl:when>
      <xsl:when test=".//图:名称_801A='Manual Operation' or .//图:类别_8019='312'"
        >flowChartManualOperation</xsl:when>
      <xsl:when test=".//图:名称_801A='Connector' or .//图:类别_8019='313'">flowChartConnector</xsl:when>
      <xsl:when test=".//图:名称_801A='Off-page Connector' or .//图:类别_8019='314'"
        >flowChartOffpageConnector</xsl:when>
      <xsl:when test=".//图:名称_801A='Card' or .//图:类别_8019='315'">flowChartPunchedCard</xsl:when>
      <xsl:when test=".//图:名称_801A='Punched Tape' or .//图:类别_8019='316'">flowChartPunchedTape</xsl:when>
      <xsl:when test=".//图:名称_801A='Summing Junction' or .//图:类别_8019='317' "
        >flowChartSummingJunction</xsl:when>
      <xsl:when test=".//图:名称_801A='Or' or .//图:类别_8019='318'">flowChartOr</xsl:when>
      <xsl:when test=".//图:名称_801A='Collate' or .//图:类别_8019='319'">flowChartCollate</xsl:when>
      <xsl:when test=".//图:名称_801A='Sort' or .//图:类别_8019='320'">flowChartSort</xsl:when>
      <xsl:when test=".//图:名称_801A='Extract' or .//图:类别_8019='321'">flowChartExtract</xsl:when>
      <xsl:when test=".//图:名称_801A='Merge' or .//图:类别_8019='322'">flowChartMerge</xsl:when>
      <xsl:when test=".//图:名称_801A='Stored Data' or .//图:类别_8019='323'">flowChartOnlineStorage</xsl:when>
      <xsl:when test=".//图:名称_801A='Delay' or .//图:类别_8019='324'">flowChartDelay</xsl:when>
      <xsl:when test=".//图:名称_801A='Sequential Access Storage' or .//图:类别_8019='325'"
        >flowChartMagneticTape</xsl:when>
      <xsl:when test=".//图:名称_801A='Magnetic Disk' or .//图:类别_8019='326'">flowChartMagneticDisk</xsl:when>
      <xsl:when test=".//图:名称_801A='Direct Access Storage' or .//图:类别_8019='327'"
        >flowChartMagneticDrum</xsl:when>
      <xsl:when test=".//图:名称_801A='Display' or .//图:类别_8019='328'">flowChartDisplay</xsl:when>
      <xsl:when test=".//图:名称_801A='Down Ribbon' or .//图:类别_8019='410'">ribbon</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Up Ribbon' or .//图:类别_8019='411'">ellipseRibbon2</xsl:when>
      <xsl:when test=".//图:名称_801A='Curved Down Ribbon' or .//图:类别_8019='412'">ellipseRibbon</xsl:when>
      <xsl:when test=".//图:名称_801A='Vertical Scroll' or .//图:类别_8019='413'">verticalScroll</xsl:when>
      <xsl:when test=".//图:名称_801A='Horizontal Scroll' or .//图:类别_8019='414'">horizontalScroll</xsl:when>
      <xsl:when test=".//图:名称_801A='Wave' or .//图:类别_8019='415'">wave</xsl:when>
      <xsl:when test=".//图:名称_801A='Double Wave' or .//图:类别_8019='416'">doubleWave</xsl:when>
      <xsl:otherwise>rect</xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="fill">
    <xsl:param name="fillpos"/>
    <xsl:if test=".//图:颜色_8004 and substring(.//图:颜色_8004,2,6)!='ffffff'">
      <a:solidFill>
        <a:srgbClr>
          <xsl:choose>
            <xsl:when test=".//图:属性_801D/图:填充_804C/图:颜色_8004='auto'">
              <xsl:attribute name="val">ffffff</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="val">
                <xsl:value-of select="substring(.//图:颜色_8004,2,6)"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:if test=".//图:属性_801D/图:透明度_8050">
            <a:alpha>
              <xsl:attribute name="val">
                <xsl:value-of select="(100-.//图:属性_801D/图:透明度_8050)*1000"/>
              </xsl:attribute>
            </a:alpha>
          </xsl:if>
        </a:srgbClr>
      </a:solidFill>
    </xsl:if>
    <!--处理除白色填充-->
    <xsl:if test="substring(.//图:颜色_8004,2,6)='ffffff'">
      <a:solidFill>
        <a:srgbClr val="ffffff"/>
      </a:solidFill>
    </xsl:if>
    <!--处理渐变填充-->
    <xsl:if test=".//图:渐变_800D">
      <xsl:variable name="angle" select=".//图:渐变_800D/@渐变方向_8013"/>
      <a:gradFill>
        <a:gsLst>
          <xsl:choose>
            <xsl:when test=".//图:渐变_800D/@种子类型_8010='radar'">
              <a:gs pos="0">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(.//图:渐变_800D/@起始色_800E,2,6)"/>
                  </xsl:attribute>
                </a:srgbClr>
              </a:gs>
              <a:gs pos="50000">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(.//图:渐变_800D/@终止色_800F,2,6)"/>
                  </xsl:attribute>
                </a:srgbClr>
              </a:gs>
              <a:gs pos="100000">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(.//图:渐变_800D/@起始色_800E,2,6)"/>
                  </xsl:attribute>
                </a:srgbClr>
              </a:gs>
            </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="$angle='135' or $angle='180' or $angle='225' or $angle='270'">
                  <a:gs pos="100000">
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:value-of select="substring(.//图:渐变_800D/@起始色_800E,2,6)"/>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:gs>
                  <a:gs pos="0">
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:value-of select="substring(.//图:渐变_800D/@终止色_800F,2,6)"/>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:gs>
                </xsl:when>
                <xsl:otherwise>
                  <a:gs pos="0">
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:value-of select="substring(.//图:渐变_800D/@起始色_800E,2,6)"/>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:gs>
                  <a:gs pos="100000">
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:value-of select="substring(.//图:渐变_800D/@终止色_800F,2,6)"/>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:gs>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
          </xsl:choose>
        </a:gsLst>
        <xsl:choose>
          <xsl:when test=".//图:渐变_800D/@种子类型_8010='square'">
            <a:path>
              <xsl:attribute name="path">rect</xsl:attribute>
              <xsl:variable name="x" select=".//图:渐变_800D/@种子X位置_8015"/>
              <xsl:variable name="y" select=".//图:渐变_800D/@种子Y位置_8016"/>
              <xsl:choose>
                <xsl:when test="$x='30' and $y='30'">
                  <a:fillToRect r="100000" b="100000"/>
                </xsl:when>
                <xsl:when test="$x='30' and $y='60'">
                  <a:fillToRect t="100000" r="100000"/>
                </xsl:when>
                <xsl:when test="$x='60' and $y='30'">
                  <a:fillToRect l="100000" b="100000"/>
                </xsl:when>
                <xsl:when test="$x='60' and $y='60'">
                  <a:fillToRect l="100000" t="100000"/>
                </xsl:when>
              </xsl:choose>
            </a:path>
          </xsl:when>
          <xsl:otherwise>
            <a:lin>
              <xsl:attribute name="ang">
                <xsl:choose>
                  <xsl:when test="$angle='0'">
                    <xsl:value-of select="number(90*60000)"/>
                  </xsl:when>
                  <xsl:when test="$angle='45'">
                    <xsl:value-of select="number(45*60000)"/>
                  </xsl:when>
                  <xsl:when test="$angle='90'">
                    <xsl:value-of select="number(0*60000)"/>
                  </xsl:when>
                  <xsl:when test="$angle='135'">
                    <xsl:value-of select="number(315*60000)"/>
                  </xsl:when>
                  <xsl:when test="$angle='180'">
                    <xsl:value-of select="number(270*60000)"/>
                  </xsl:when>
                  <xsl:when test="$angle='225'">
                    <xsl:value-of select="number(225*60000)"/>
                  </xsl:when>
                  <xsl:when test="$angle='270'">
                    <xsl:value-of select="number(180*60000)"/>
                  </xsl:when>
                  <xsl:when test="$angle='315'">
                    <xsl:value-of select="number(135*60000)"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'0'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <xsl:attribute name="scaled">1</xsl:attribute>
            </a:lin>
          </xsl:otherwise>
        </xsl:choose>
        <a:tileRect/>
      </a:gradFill>
    </xsl:if>
    <!--2月20日改:处理纹理和图片填充-->
    <xsl:if test=".//图:图片_8005">
      <a:blipFill>
        <a:blip>

          <xsl:attribute name="r:embed">
            <xsl:value-of select="concat('rIdObj',$fillpos)"/>
          </xsl:attribute>
        </a:blip>
        <xsl:if test=".//图:图片_8005/@位置_8006='tile'">
          <a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>
        </xsl:if>
        <a:srcRect/>
        <xsl:if test=".//图:图片_8005/@位置_8006='stretch'">
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </xsl:if>
      </a:blipFill>
    </xsl:if>
    <!--处理图案填充-->
    <xsl:if test=".//图:图案_800A">
      <a:pattFill>
        <xsl:attribute name="prst">
          <xsl:call-template name="fillName"/>
        </xsl:attribute>
        <a:fgClr>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring(.//图:图案_800A/@前景色_800B,2,6)"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:fgClr>
        <a:bgClr>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring(.//图:图案_800A/@背景色_800C,2,6)"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:bgClr>
      </a:pattFill>
    </xsl:if>
  </xsl:template>
  <xsl:template name="fillName">
    <xsl:choose>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn001'">pct5</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn002'">pct10</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn003'">pct20</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn004'">pct25</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn005'">pct30</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn006'">pct40</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn007'">pct50</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn008'">pct60</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn009'">pct70</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn010'">pct75</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn011'">pct80</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn012'">pct90</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn013'">ltDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn014'">ltUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn015'">dkDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn016'">dkUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn017'">wdDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn018'">wdUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn019'">ltVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn020'">ltHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn021'">narVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn022'">narHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn023'">dkVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn024'">dkHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn025'">dashDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn026'">dashUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn027'">dashHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn028'">dashVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn029'">smConfetti</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn030'">lgConfetti</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn031'">zigZag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn032'">wave</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn033'">diagBrick</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn034'">horzBrick</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn035'">weave</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn036'">plaid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn037'">divot</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn038'">dotGrid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn039'">dotDmnd</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn040'">shingle</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn041'">trellis</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn042'">sphere</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn043'">smGrid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn044'">lgGrid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn045'">smCheck</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn046'">lgCheck</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn047'">openDmnd</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn048'">solidDmnd</xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="grpXfrm">
    
    <!--2013-05-03，wudi，当前节点，图:图形_8062-->
    <!--2013-04-12，wudi，修复组合图形转换的BUG，添加参数list-->
    <!--2013-05-02，wudi，增加参数isSmartArt，标识是否是SmartArt转换，目前SmartArt按组合图形转，start-->
    <!--2013-05-03，wudi，增加Label参数，标识组合图形是否是嵌套组合，start-->
    <!--2013-05-07，wudi，取消Label参数，找到更好的处理方式，start-->
    <xsl:param name ="list"/>
    <xsl:param name ="isSmartArt"/>
    <!--<xsl:param name ="Label"/>-->
    <a:xfrm>
      <xsl:if test=".//图:翻转_803A">
        <!--cxl,2012.3.17做为a:xfrm的属性-->
        <xsl:call-template name="flip"/>
      </xsl:if>
      <!-- 09.10.30 added by myx -->
      <xsl:if test=".//图:旋转角度_804D!='0.0'">
        <xsl:attribute name="rot">
          <!--10.02.02 myx<xsl:value-of select="round(21600000 - .//图:属性/图:旋转角度*60000)"/>-->
          <xsl:value-of select="round(.//图:属性_801D/图:旋转角度_804D*60000)"/>
        </xsl:attribute>
      </xsl:if>
      <a:off>
        
        <!--2013-05-07，wudi，修复UOF到OOX方向组合图形转换BUG，针对组合图形嵌套组合图形的情况，start-->
        <xsl:attribute name="x">
          <xsl:choose>
            <xsl:when test="./图:组合位置_803B">
              <xsl:choose>
                <xsl:when test="contains(./图:组合位置_803B/@x_C606,'E')">
                  <xsl:value-of select="0"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="round(./图:组合位置_803B/@x_C606 * 12700)"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="0"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--<xsl:value-of select ="'0'"/>-->
        </xsl:attribute>
        <xsl:attribute name="y">
          <xsl:choose>
            <xsl:when test="./图:组合位置_803B">
              <xsl:choose>
                <xsl:when test="contains(./图:组合位置_803B/@y_C607,'E')">
                  <xsl:value-of select="0"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="round(./图:组合位置_803B/@y_C607 * 12700)"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="0"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--<xsl:value-of select ="'0'"/>-->
        </xsl:attribute>
      </a:off>
      <!--end-->
      
      <a:ext>
        
        <!--2013-04-12，wudi，修复组合图形转换BUG，原来cx和cy的取值弄反-->
        <!--2013-04-24，wudi，修复组合图形转换的BUG，部分案例@宽_C605或@长_C604属性缺失，start-->
        <xsl:attribute name="cx">
          <xsl:choose>
            <xsl:when test="contains(.//图:大小_8060/@宽_C605,'E') or not(.//图:大小_8060/@宽_C605)">
              <xsl:value-of select="0"/>
            </xsl:when>
            <xsl:when test="not(contains(.//图:大小_8060/@宽_C605,'E')) and (.//图:大小_8060/@宽_C605)">
              <xsl:value-of select="round(.//图:大小_8060/@宽_C605 * 12700)"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="cy">
          <xsl:choose>
            <xsl:when test="contains(.//图:大小_8060/@长_C604,'E') or not(.//图:大小_8060/@长_C604)">
              <xsl:value-of select="0"/>
            </xsl:when>
            <xsl:when test="not(contains(.//图:大小_8060/@长_C604,'E')) and (.//图:大小_8060/@长_C604)">
              <xsl:value-of select="round(.//图:大小_8060/@长_C604 * 12700)"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
        <!--end-->
        <!--end-->
        
      </a:ext>

      <!--2013-05-03，wudi，组合图形，流程图，SmartArt三种情况分别考虑，start-->
      <xsl:choose>
        <xsl:when test ="$isSmartArt ='1'">

          <xsl:variable name ="Tag">
            <xsl:value-of select ="@组合列表_8064"/>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test ="contains($Tag,'SmartArtObj')">
              <a:chOff x="0" y="0"/>
              <a:chExt cx="0" cy="0"/>
            </xsl:when>
            <xsl:otherwise>
              <!--<xsl:if test ="$Label ='1'">-->
              <xsl:variable name ="Num">
                <xsl:number format ="1" level ="any" count ="//uof:对象集/图:图形_8062[contains($list,@标识符_804B)]"/>
              </xsl:variable>
              <xsl:variable name ="Temp1">
                <xsl:value-of select ="substring($list,1,8)"/>
              </xsl:variable>
              <xsl:variable name ="Temp2">
                <xsl:value-of select ="normalize-space($Temp1)"/>
              </xsl:variable>
              <xsl:variable name ="FirstNam">
                <xsl:value-of select ="$Temp2"/>
              </xsl:variable>
              <xsl:call-template name ="ChOff">
                <xsl:with-param name ="num" select ="$Num"/>
                <xsl:with-param name ="chOffX" select ="'1000'"/>
                <xsl:with-param name ="chOffY" select ="'1000'"/>
                <xsl:with-param name ="firstNam" select ="$FirstNam"/>
                <xsl:with-param name ="list" select ="$list"/>
              </xsl:call-template>
              <!--</xsl:if>-->
              <!--<xsl:if test ="$Label !='1'">
              <a:chOff x="0" y="0"/>
            </xsl:if>-->

              <a:chExt>
                <xsl:attribute name="cx">
                  <xsl:choose>
                    <xsl:when test="contains(.//图:大小_8060/@宽_C605,'E') or not(.//图:大小_8060/@宽_C605)">
                      <xsl:value-of select="0"/>
                    </xsl:when>
                    <xsl:when test="not(contains(.//图:大小_8060/@宽_C605,'E')) and (.//图:大小_8060/@宽_C605)">
                      <xsl:choose>
                        <xsl:when test ="contains($Tag,'GrpspObj')">
                          <xsl:value-of select="round(.//图:大小_8060/@宽_C605 * 20)"/>
                        </xsl:when>
                        <xsl:when test ="contains($Tag,'grpspObj')">
                          <xsl:value-of select="round(.//图:大小_8060/@宽_C605 * 12700)"/>
                        </xsl:when>
                        <xsl:when test ="contains($Tag,'SmartArtObj')">
                          <xsl:value-of select="round(.//图:大小_8060/@宽_C605 * 12700)"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="round(.//图:大小_8060/@宽_C605 * 12700)"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                  </xsl:choose>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:choose>
                    <xsl:when test="contains(.//图:大小_8060/@长_C604,'E') or not(.//图:大小_8060/@长_C604)">
                      <xsl:value-of select="0"/>
                    </xsl:when>
                    <xsl:when test="not(contains(.//图:大小_8060/@长_C604,'E')) and (.//图:大小_8060/@长_C604)">
                      <xsl:choose>
                        <xsl:when test ="contains($Tag,'GrpspObj')">
                          <xsl:value-of select="round(.//图:大小_8060/@长_C604 * 20)"/>
                        </xsl:when>
                        <xsl:when test ="contains($Tag,'grpspObj')">
                          <xsl:value-of select="round(.//图:大小_8060/@长_C604 * 12700)"/>
                        </xsl:when>
                        <xsl:when test ="contains($Tag,'SmartArtObj')">
                          <xsl:value-of select="round(.//图:大小_8060/@宽_C605 * 12700)"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="round(.//图:大小_8060/@长_C604 * 12700)"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                  </xsl:choose>
                </xsl:attribute>
              </a:chExt>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test ="not($isSmartArt ='1')">

          <xsl:variable name ="tag">
            <xsl:value-of select ="@组合列表_8064"/>
          </xsl:variable>

          <!--<xsl:if test ="$Label = '1'">-->
          <xsl:choose>
            <xsl:when test ="contains($tag,'GrpspObj')">
              <!--2013-04-12，wudi，修复UOF到OOX方向组合图形转换BUG，以下代码转a:chOff节点-->
              <xsl:variable name ="num">
                <xsl:number format ="1" level ="any" count ="//uof:对象集/图:图形_8062[contains($list,@标识符_804B)]"/>
              </xsl:variable>
              <xsl:variable name ="temp1">
                <xsl:value-of select ="substring($list,2,10)"/>
              </xsl:variable>
              <xsl:variable name ="temp2">
                <xsl:value-of select ="normalize-space($temp1)"/>
              </xsl:variable>
              <xsl:variable name ="firstNam">
                <xsl:value-of select ="$temp2"/>
              </xsl:variable>
              <xsl:call-template name ="chOff">
                <xsl:with-param name ="num" select ="$num"/>
                <xsl:with-param name ="chOffX" select ="'1000'"/>
                <xsl:with-param name ="chOffY" select ="'1000'"/>
                <xsl:with-param name ="firstNam" select ="$firstNam"/>
                <xsl:with-param name ="list" select ="$list"/>
              </xsl:call-template>
              <!--end-->
            </xsl:when>

            <!--2013-05-07，wudi，修复UOF到OOX方向组合图形转换BUG，针对标识符是以'Obj000'开头的，start-->
            <xsl:when test ="contains($tag,'Obj000')">
              <xsl:variable name ="Num">
                <xsl:number format ="1" level ="any" count ="//uof:对象集/图:图形_8062[contains($list,@标识符_804B)]"/>
              </xsl:variable>
              <xsl:variable name ="Temp1">
                <xsl:value-of select ="substring($list,1,8)"/>
              </xsl:variable>
              <xsl:variable name ="Temp2">
                <xsl:value-of select ="normalize-space($Temp1)"/>
              </xsl:variable>
              <xsl:variable name ="FirstNam">
                <xsl:value-of select ="$Temp2"/>
              </xsl:variable>
              <xsl:call-template name ="ChOff">
                <xsl:with-param name ="num" select ="$Num"/>
                <xsl:with-param name ="chOffX" select ="'1000'"/>
                <xsl:with-param name ="chOffY" select ="'1000'"/>
                <xsl:with-param name ="firstNam" select ="$FirstNam"/>
                <xsl:with-param name ="list" select ="$list"/>
              </xsl:call-template>
            </xsl:when>
            <!--end-->

            <xsl:when test ="contains($tag,'grpspObj')">
              <a:chOff x="0" y="0"/>
            </xsl:when>
            <xsl:when test ="contains($tag,'SmartArtObj')">
              <a:chOff x="0" y="0"/>
            </xsl:when>
            <xsl:otherwise>
              <a:chOff x="0" y="0"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--</xsl:if>-->
          <!--<xsl:if test ="$Label != '1'">
          <a:chOff x="0" y="0"/>
        </xsl:if>-->

          <!--2013-04-12，wudi，修复组合图形转换的BUG，调整参数12700为20，start-->

          <!--2013-04-24，wudi，修复组合图形转换的BUG，部分案例@宽_C605或@长_C604属性缺失，start-->
          <a:chExt>
            <xsl:attribute name="cx">
              <xsl:choose>
                <xsl:when test="contains(.//图:大小_8060/@宽_C605,'E') or not(.//图:大小_8060/@宽_C605)">
                  <xsl:value-of select="0"/>
                </xsl:when>
                <xsl:when test="not(contains(.//图:大小_8060/@宽_C605,'E')) and (.//图:大小_8060/@宽_C605)">
                  <xsl:choose>
                    <xsl:when test ="contains($tag,'GrpspObj')">
                      <xsl:value-of select="round(.//图:大小_8060/@宽_C605 * 20)"/>
                    </xsl:when>
                    <xsl:when test ="contains($tag,'grpspObj')">
                      <xsl:value-of select="round(.//图:大小_8060/@宽_C605 * 12700)"/>
                    </xsl:when>
                    <xsl:when test ="contains($tag,'SmartArtObj')">
                      <xsl:value-of select="round(.//图:大小_8060/@宽_C605 * 12700)"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="round(.//图:大小_8060/@宽_C605 * 12700)"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="cy">
              <xsl:choose>
                <xsl:when test="contains(.//图:大小_8060/@长_C604,'E') or not(.//图:大小_8060/@长_C604)">
                  <xsl:value-of select="0"/>
                </xsl:when>
                <xsl:when test="not(contains(.//图:大小_8060/@长_C604,'E')) and (.//图:大小_8060/@长_C604)">
                  <xsl:choose>
                    <xsl:when test ="contains($tag,'GrpspObj')">
                      <xsl:value-of select="round(.//图:大小_8060/@长_C604 * 20)"/>
                    </xsl:when>
                    <xsl:when test ="contains($tag,'grpspObj')">
                      <xsl:value-of select="round(.//图:大小_8060/@长_C604 * 12700)"/>
                    </xsl:when>
                    <xsl:when test ="contains($tag,'SmartArtObj')">
                      <xsl:value-of select="round(.//图:大小_8060/@宽_C605 * 12700)"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="round(.//图:大小_8060/@长_C604 * 12700)"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </a:chExt>
          <!--end-->

          <!--end-->

        </xsl:when>
        
      </xsl:choose>
      <!--end-->
      
    </a:xfrm>
    
    <!--end-->
    
    <!--end-->
    
    <!--end-->

  </xsl:template>

  <!--2013-05-03，wudi，增加模板ChOff，修改实现方法，start-->
  <xsl:template name ="ChOff">
    <xsl:param name ="num"/>
    <xsl:param name ="chOffX"/>
    <xsl:param name ="chOffY"/>
    <xsl:param name ="firstNam"/>
    <xsl:param name ="list"/>
    <xsl:if test ="//uof:对象集/图:图形_8062[@标识符_804B =$firstNam]">
      <xsl:variable name ="tmp">
        <xsl:choose>
          <xsl:when test ="//uof:对象集/图:图形_8062[@标识符_804B =$firstNam]/图:组合位置_803B/@x_C606 &lt;= $chOffX">
            <xsl:value-of select ="//uof:对象集/图:图形_8062[@标识符_804B =$firstNam]/图:组合位置_803B/@x_C606"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select ="$chOffX"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name ="tmp1">
        <xsl:choose>
          <xsl:when test ="//uof:对象集/图:图形_8062[@标识符_804B =$firstNam]/图:组合位置_803B/@y_C607 &lt;= $chOffY">
            <xsl:value-of select ="//uof:对象集/图:图形_8062[@标识符_804B =$firstNam]/图:组合位置_803B/@y_C607"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select ="$chOffY"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name ="num1">
        <xsl:value-of select ="number($num) - 1"/>
      </xsl:variable>
      <xsl:variable name ="temp1">
        <xsl:value-of select ="substring-after($list,$firstNam)"/>
      </xsl:variable>
      <xsl:variable name ="temp2">
        <xsl:value-of select ="normalize-space($temp1)"/>
      </xsl:variable>
      <xsl:variable name ="temp3">
        <xsl:value-of select ="substring($temp2,1,8)"/>
      </xsl:variable>
      <xsl:variable name ="temp4">
        <xsl:value-of select ="normalize-space($temp3)"/>
      </xsl:variable>
      <xsl:call-template name ="ChOff">
        <xsl:with-param name ="num" select ="$num1"/>
        <xsl:with-param name ="chOffX" select ="$tmp"/>
        <xsl:with-param name ="chOffY" select ="$tmp1"/>
        <xsl:with-param name ="firstNam" select ="$temp4"/>
        <xsl:with-param name ="list" select ="$temp2"/>
      </xsl:call-template>
    </xsl:if>
    <xsl:if test ="$num = 0">
      <a:chOff>
        <xsl:attribute name ="x">
          <xsl:value-of select ="round($chOffX * 12700)"/>
        </xsl:attribute>
        <xsl:attribute name ="y">
          <xsl:value-of select ="round($chOffY * 12700)"/>
        </xsl:attribute>
      </a:chOff>
    </xsl:if>
  </xsl:template>
  
  <!--2013-04-12，wudi，修复组合图形转换BUG，新增加的模板，处理a:chOff节点-->
  <xsl:template name ="chOff">
    <xsl:param name ="num"/>
    <xsl:param name ="chOffX"/>
    <xsl:param name ="chOffY"/>
    <xsl:param name ="firstNam"/>
    <xsl:param name ="list"/>
    <xsl:if test ="//uof:对象集/图:图形_8062[@标识符_804B =$firstNam]">
      <xsl:variable name ="tmp">
        <xsl:choose>
          <xsl:when test ="//uof:对象集/图:图形_8062[@标识符_804B =$firstNam]/图:组合位置_803B/@x_C606 &lt;= $chOffX">
            <xsl:value-of select ="//uof:对象集/图:图形_8062[@标识符_804B =$firstNam]/图:组合位置_803B/@x_C606"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select ="$chOffX"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name ="tmp1">
        <xsl:choose>
          <xsl:when test ="//uof:对象集/图:图形_8062[@标识符_804B =$firstNam]/图:组合位置_803B/@y_C607 &lt;= $chOffY">
            <xsl:value-of select ="//uof:对象集/图:图形_8062[@标识符_804B =$firstNam]/图:组合位置_803B/@y_C607"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select ="$chOffY"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name ="num1">
        <xsl:value-of select ="number($num) - 1"/>
      </xsl:variable>
      <xsl:variable name ="temp1">
        <xsl:value-of select ="substring-after($list,$firstNam)"/>
      </xsl:variable>
      <xsl:variable name ="temp2">
        <xsl:value-of select ="normalize-space($temp1)"/>
      </xsl:variable>
      <xsl:variable name ="temp3">
        <xsl:value-of select ="substring($temp2,1,10)"/>
      </xsl:variable>
      <xsl:variable name ="temp4">
        <xsl:value-of select ="normalize-space($temp3)"/>
      </xsl:variable>
      <xsl:call-template name ="chOff">
        <xsl:with-param name ="num" select ="$num1"/>
        <xsl:with-param name ="chOffX" select ="$tmp"/>
        <xsl:with-param name ="chOffY" select ="$tmp1"/>
        <xsl:with-param name ="firstNam" select ="$temp4"/>
        <xsl:with-param name ="list" select ="$temp2"/>
      </xsl:call-template>
    </xsl:if>
    <xsl:if test ="$num = 0">
      <a:chOff>
        <xsl:attribute name ="x">
          <xsl:value-of select ="round($chOffX * 12700)"/>
        </xsl:attribute>
        <xsl:attribute name ="y">
          <xsl:value-of select ="round($chOffY * 12700)"/>
        </xsl:attribute>
      </a:chOff>
    </xsl:if>
  </xsl:template>
  <!--end-->
  
  <!--end-->
  
  <xsl:template name="flip">
    <!--cxl,2012.3.17互换x\y-->
    <xsl:if test=".//图:翻转_803A='x'">
      <xsl:attribute name="flipH">1</xsl:attribute>
    </xsl:if>
    <xsl:if test=".//图:翻转_803A='y'">
      <xsl:attribute name="flipV">1</xsl:attribute>
    </xsl:if>
  </xsl:template>

  <!--cxl,图:文本_803C这里标准改动较大**********************************************************-->
  <xsl:template match="图:文本_803C" mode="txbody">
    <xsl:param name="objectid"/>
    <xsl:attribute name="rot">
      <xsl:value-of select="'0'"/>
    </xsl:attribute>
    <xsl:attribute name="anchor">
      <xsl:choose>
        <xsl:when test="图:对齐_803E/@文字对齐_421E='top'">
          <xsl:value-of select="'t'"/>
        </xsl:when>
        <xsl:when test="图:对齐_803E/@文字对齐_421E='center'">
          <xsl:value-of select="'ctr'"/>
        </xsl:when>
        <xsl:when test="图:对齐_803E/@文字对齐_421E='bottom'">
          <xsl:value-of select="'b'"/>
        </xsl:when>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="wrap">
      <xsl:choose>
        <xsl:when test="@是否自动换行_8047='true'">
          <xsl:value-of select="'square'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'none'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>

    <!--2014-06-04，wudi，修改转换比例，有变化，start-->
    <xsl:attribute name="lIns">
      <xsl:value-of select="round(图:边距_803D/@左_C608*12700)"/>
    </xsl:attribute>
    <xsl:attribute name="rIns">
      <xsl:value-of select="round(图:边距_803D/@右_C60A*12700)"/>
    </xsl:attribute>
    <xsl:attribute name="tIns">
      <xsl:value-of select="round(图:边距_803D/@上_C609*12700)"/>
    </xsl:attribute>
    <xsl:attribute name="bIns">
      <xsl:value-of select="round(图:边距_803D/@下_C60B*12700)"/>
    </xsl:attribute>
    <!--end-->
    
    <!--zhaobj   word中的不旋转文本upright属性
    <xsl:attribute name ="upright">
      <xsl:choose>
        <xsl:when test ="//uof:扩展区/uof:扩展/uof:扩展内容/uof:内容/uof:文本框属性[@uof:对象索引 = $objectid]/@uof:是否随图形旋转文字='true'">
          <xsl:value-of select ="0"/>
        </xsl:when>
        <xsl:when test ="//uof:扩展区/uof:扩展/uof:扩展内容/uof:内容/uof:文本框属性[@uof:对象索引 = $objectid]/@uof:是否随图形旋转文字='false'">
         <xsl:value-of select ="'1'"/>
        </xsl:when>
        <xsl:otherwise >
          <xsl:value-of select ="0"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>  -->
    <xsl:attribute name="vert">
      <xsl:choose>
        <xsl:when test="图:文字排列方向_8042='t2b-l2r-0e-0w'">
          <xsl:value-of select="'horz'"/>
        </xsl:when>
        <xsl:when test="图:文字排列方向_8042='r2l-t2b-0e-90w'">
          <xsl:value-of select="'eaVert'"/>
        </xsl:when>
        <xsl:when test="图:文字排列方向_8042='r2l-t2b-90e-90w'">
          <xsl:value-of select="'vert'"/>
        </xsl:when>
        <xsl:when test="图:文字排列方向_8042='l2r-b2t-270e-270w'">
          <xsl:value-of select="'vert270'"/>
        </xsl:when>
        <xsl:when test="图:文字排列方向_8042='t2b-r2l-0e-0w'">
          <xsl:value-of select="'horz'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'horz'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>

    <a:prstTxWarp prst="textNoShape">
      <a:avLst/>
    </a:prstTxWarp>
    <xsl:choose>
      <xsl:when test="@是否大小适应文字_8048='true'">
        <a:spAutoFit/>
      </xsl:when>
      <xsl:otherwise>
        <a:noAutofit/>
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:template>
  <!--*****************************************************************************************-->

  <xsl:template match="图:文本_803C" mode="wpstxbx">
    <wps:txbx>
      <w:txbxContent>
        
        <xsl:for-each select="./图:内容_8043/字:文字表_416C">
          <xsl:call-template name="table"/>
        </xsl:for-each>
        
        <xsl:for-each select="./图:内容_8043/字:段落_416B">
          <xsl:call-template name="wpsparagraph"/>
        </xsl:for-each>
      </w:txbxContent>
    </wps:txbx>
  </xsl:template>
  <!--cxl,2012.2.26增加艺术字转换-->
  <!--wcz,2013.3.7修复艺术字效果丢失问题-->
  <xsl:template match="图:艺术字_802D" mode="wpstxbx">
    <wps:txbx>
      <w:txbxContent>
        <w:p>
          <w:pPr>
            <w:jc>
              <xsl:attribute name="w:val">
                <xsl:choose>
                  <xsl:when test="图:对齐方式_8031">
                    <xsl:value-of select="图:对齐方式_8031"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="center"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </w:jc>
            <w:rPr>
              <w:b/>
              <w:color>
                <xsl:attribute name="w:val">
                  <xsl:choose>
                    <xsl:when test="preceding-sibling::图:线_8057/图:线颜色_8058">
                      <xsl:value-of select="preceding-sibling::图:线_8057/图:线颜色_8058"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'000000'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </w:color>
              
              <!--2013-04-17，wudi，修复UOF到OOX方向艺术字字号不正确的BUG，start-->
              <w:sz>
                <xsl:attribute name="w:val">
                  <xsl:value-of select="round(图:字体_802E/@字号_412D * 2)"/>
                </xsl:attribute>
              </w:sz>
              <w:szCs>
                <xsl:attribute name="w:val">
                  <xsl:value-of select="round(图:字体_802E/@字号_412D * 2)"/>
                </xsl:attribute>
              </w:szCs>
              <!--end-->
              
            </w:rPr>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts>
                <xsl:attribute name="w:hint">
                  <xsl:value-of select="'eastAsia'"/>
                </xsl:attribute>
                <xsl:if test="图:字体_802E/@西文字体引用_4129">
                  <xsl:attribute name="w:ascii">
                    <xsl:variable name="temp" select="图:字体_802E/@西文字体引用_4129"/>
                    <xsl:value-of select="/uof:UOF/uof:式样集//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$temp]/@名称_9903"/>
                  </xsl:attribute>
                  <xsl:attribute name="w:hAnsi">
                    <xsl:variable name="temp" select="图:字体_802E/@西文字体引用_4129"/>
                    <xsl:value-of select="/uof:UOF/uof:式样集//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$temp]/@名称_9903"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="图:字体_802E/@中文字体引用_412A">
                  <xsl:attribute name="w:eastAsia">
                    <xsl:variable name="temp" select="图:字体_802E/@中文字体引用_412A"/>
                    <xsl:value-of select="/uof:UOF/uof:式样集//式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$temp]/@名称_9903"/>
                  </xsl:attribute>
                </xsl:if>
              </w:rFonts>
              <w:b/>
              <w:color>
                <xsl:attribute name="w:val">
                  <xsl:choose>
                    <xsl:when test="preceding-sibling::图:线_8057/图:线颜色_8058">
                      <xsl:value-of select="preceding-sibling::图:线_8057/图:线颜色_8058"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="000000"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </w:color>

              <!--2013-04-17，wudi，修复UOF到OOX方向艺术字字号不正确的BUG，start-->
              <w:sz>
                <xsl:attribute name="w:val">
                  <xsl:value-of select="round(图:字体_802E/@字号_412D * 2)"/>
                </xsl:attribute>
              </w:sz>
              <w:szCs>
                <xsl:attribute name="w:val">
                  <xsl:value-of select="round(图:字体_802E/@字号_412D * 2)"/>
                </xsl:attribute>
              </w:szCs>
              <!--end-->
       
            </w:rPr>
            <w:t>
              <xsl:value-of select="./图:艺术字文本_8036"/>
            </w:t>
          </w:r>
        </w:p>
      </w:txbxContent>
    </wps:txbx>
  </xsl:template>
  <xsl:template match="图:艺术字_802D" mode="txbody">
    <wps:bodyPr>
      <xsl:attribute name="rot">
        <xsl:value-of select="'0'"/>
      </xsl:attribute>
      <xsl:attribute name="anchor">
        <!--<xsl:choose>
        <xsl:when test="图:对齐方式_8031='top'">
          <xsl:value-of select="'t'"/>
        </xsl:when>
        <xsl:when test="图:对齐方式_8031='center'">-->
        <xsl:value-of select="'ctr'"/>
        <!--</xsl:when>
        <xsl:when test="图:对齐方式_8031='bottom'">
          <xsl:value-of select="'b'"/>
        </xsl:when>
      </xsl:choose>-->
      </xsl:attribute>
      <xsl:attribute name="wrap">
        <!--<xsl:choose>
        <xsl:when test="@是否自动换行_8047='true'">
          <xsl:value-of select="'square'"/>
        </xsl:when>
        <xsl:otherwise>-->
        <xsl:value-of select="'none'"/>
        <!--</xsl:otherwise>
      </xsl:choose>-->
      </xsl:attribute>
      <!--<xsl:attribute name="lIns">
      <xsl:value-of select="图:边距_803D/@左_C608*9525"/>
    </xsl:attribute>
    <xsl:attribute name="rIns">
      <xsl:value-of select="图:边距_803D/@右_C60A*9525"/>
    </xsl:attribute>
    <xsl:attribute name="tIns">
      <xsl:value-of select="图:边距_803D/@上_C609*9525"/>
    </xsl:attribute>
    <xsl:attribute name="bIns">
      <xsl:value-of select="图:边距_803D/@下_C60B*9525"/>
    </xsl:attribute>-->
      <!--zhaobj   word中的不旋转文本upright属性
    <xsl:attribute name ="upright">
      <xsl:choose>
        <xsl:when test ="//uof:扩展区/uof:扩展/uof:扩展内容/uof:内容/uof:文本框属性[@uof:对象索引 = $objectid]/@uof:是否随图形旋转文字='true'">
          <xsl:value-of select ="0"/>
        </xsl:when>
        <xsl:when test ="//uof:扩展区/uof:扩展/uof:扩展内容/uof:内容/uof:文本框属性[@uof:对象索引 = $objectid]/@uof:是否随图形旋转文字='false'">
         <xsl:value-of select ="'1'"/>
        </xsl:when>
        <xsl:otherwise >
          <xsl:value-of select ="0"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>  -->
      <!--<xsl:attribute name="vert">
      <xsl:choose>
        <xsl:when test="图:文字排列方向_8042='t2b-l2r-0e-0w'">
          <xsl:value-of select="'horz'"/>
        </xsl:when>
        <xsl:when test="图:文字排列方向_8042='r2l-t2b-0e-90w'">
          <xsl:value-of select="'eaVert'"/>
        </xsl:when>
        <xsl:when test="图:文字排列方向_8042='r2l-t2b-90e-90w'">
          <xsl:value-of select="'vert'"/>
        </xsl:when>
        <xsl:when test="图:文字排列方向_8042='l2r-b2t-270e-270w'">
          <xsl:value-of select="'vert270'"/>
        </xsl:when>
        <xsl:when test="图:文字排列方向_8042='t2b-r2l-0e-0w'">
          <xsl:value-of select="'horz'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'horz'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>-->

      <a:prstTxWarp prst="textNoShape">
        <a:avLst/>
      </a:prstTxWarp>
      <!--<xsl:if test="@是否大小适应文字_8048='true'">
      <a:spAutoFit/>
    </xsl:if>
    <xsl:if test="@是否大小适应文字_8048='false'">
      <a:noAutofit/>
    </xsl:if>-->
    </wps:bodyPr>

  </xsl:template>

  <xsl:template name="wpsparagraph">
    <w:p>
      <xsl:for-each select="字:段落属性_419B | 字:域开始_419E | 字:句_419D">
        <xsl:choose>
          <xsl:when test="name(.)='字:段落属性_419B'">
            <xsl:call-template name="pPrWithpStyle"/>
          </xsl:when>
          <xsl:when test="name(.)='字:域开始_419E'">
            <xsl:if test="not(@类型_416E='toc' or @类型_416E='index' or @类型_416E='link')">
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

          </xsl:when>
          <xsl:when
            test="name(.)='字:句_419D' and (not((preceding-sibling::字:域代码_419F) and (following-sibling::字:域结束_41A0)))">
            <xsl:call-template name="run"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </w:p>
  </xsl:template>

  <xsl:template match="图:三维效果_8061">
    <a:scene3d>
      <a:camera>
        <xsl:attribute name="prst">
          <xsl:choose>
            <xsl:when test="uof:方向_C63C">
              <xsl:call-template name="scene3dprst"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'orthographicFront'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:if test="not(uof:角度_C635/uof:x方向_C636=0 and uof:角度_C635/uof:y方向_C637=0)">
          <a:rot>
            <xsl:attribute name="lon">
              <xsl:value-of select="((uof:角度_C635/uof:x方向_C636+360) mod 360 ) * 60000"/>
            </xsl:attribute>
            <xsl:attribute name="lat">
              
              <!--2013-04-16，wudi，六角-八角-八角，修复六角矩形转换互操作BUG，start-->
              <xsl:choose>
                <xsl:when test ="ancestor::图:预定义图形_8018/图:名称_801A = '8-Point Star'">
                  <xsl:value-of select ="'0'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="((uof:角度_C635/uof:y方向_C637+360) mod 360 ) * 60000"/>
                </xsl:otherwise>
              </xsl:choose>
              <!--end-->
              
            </xsl:attribute>
            <xsl:attribute name="rev">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
          </a:rot>
        </xsl:if>
      </a:camera>
      <a:lightRig>
        <xsl:attribute name="rig">
          <xsl:choose>
            <xsl:when test="uof:照明_C638/uof:强度_C63A='bright'">
              <xsl:value-of select="'balanced'"/>
            </xsl:when>
            <xsl:when test="uof:照明_C638/uof:强度_C63A='dim'">
              <xsl:value-of select="'harsh'"/>
            </xsl:when>
            <xsl:when test="uof:照明_C638/uof:强度_C63A='normal'">
              <xsl:value-of select="'threePt'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="dir">
          <xsl:value-of select="'t'"/>
        </xsl:attribute>
        <xsl:if test="not(round(uof:照明_C638/uof:角度_C639)=0) and uof:照明_C638/uof:角度_C639">
          <a:rot>
            <xsl:attribute name="lat">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
            <xsl:attribute name="lon">
              <xsl:value-of select="'0'"/>
            </xsl:attribute>
            <xsl:attribute name ="rev">
              <xsl:value-of select="uof:照明_C638/uof:角度_C639"/>
            </xsl:attribute>
          </a:rot>
        </xsl:if>

      </a:lightRig>
    </a:scene3d>
    <a:sp3d>
      <xsl:if test="uof:深度_C63B">
        <xsl:attribute name="extrusionH">
          <xsl:value-of select="uof:深度_C63B * 12700"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="uof:表面效果_C63E">
        <xsl:attribute name="prstMaterial">
          <xsl:choose>
            <xsl:when test="uof:表面效果_C63E='metal'">
              <xsl:value-of select="'metal'"/>
            </xsl:when>
            <xsl:when test="uof:表面效果_C63E='plastic'">
              <xsl:value-of select="'plastic'"/>
            </xsl:when>
            <xsl:when test="uof:表面效果_C63E='matte'">
              <xsl:value-of select="'matte'"/>
            </xsl:when>
            <xsl:when test="uof:表面效果_C63E='wire-frame'">
              <xsl:value-of select="'legacyWireframe'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="uof:颜色_C63F">
        <a:extrusionClr>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring-after(uof:颜色_C63F,'#')"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:extrusionClr>
      </xsl:if>
    </a:sp3d>
  </xsl:template>
  <xsl:template name="scene3dprst">
    <xsl:choose>
      <xsl:when test="not(uof:样式_C634)">
        <xsl:choose>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='none' and uof:方向_C63C/uof:方式_C63D='parallel'">
            <xsl:value-of select="'orthographicFront'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-top-right' and uof:方向_C63C/uof:方式_C63D='parallel'">
            <xsl:value-of select="'obliqueTopRight'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-top-left' and uof:方向_C63C/uof:方式_C63D='parallel'">
            <xsl:value-of select="'obliqueTopLeft'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-top' and uof:方向_C63C/uof:方式_C63D='parallel'">
            <xsl:value-of select="'perspectiveBelow'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-left' and uof:方向_C63C/uof:方式_C63D='parallel'">
            <xsl:value-of select="'perspectiveRight'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-right' and uof:方向_C63C/uof:方式_C63D='parallel'">
            <xsl:value-of select="'perspectiveLeft'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-bottom-left' and uof:方向_C63C/uof:方式_C63D='parallel'">
            <xsl:value-of select="'obliqueBottomLeft'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-bottom-right' and uof:方向_C63C/uof:方式_C63D='parallel'">
            <xsl:value-of select="'isometricLeftDown'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-bottom' and uof:方向_C63C/uof:方式_C63D='parallel'">
            <xsl:value-of select="'perspectiveAbove'"/>
          </xsl:when>

          <xsl:when test="uof:方向_C63C/uof:角度_C639='none' and uof:方向_C63C/uof:方式_C63D='perspective'">
            <xsl:value-of select="'perspectiveFront'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-top-left' and uof:方向_C63C/uof:方式_C63D='perspective'">
            <xsl:value-of select="'obliqueTopLeft'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-top' and uof:方向_C63C/uof:方式_C63D='perspective'">
            <xsl:value-of select="'perspectiveBelow'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-top-right' and uof:方向_C63C/uof:方式_C63D='perspective'">
            <xsl:value-of select="'obliqueTopRight'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-left' and uof:方向_C63C/uof:方式_C63D='perspective'">
            <xsl:value-of select="'perspectiveRight'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-right' and uof:方向_C63C/uof:方式_C63D='perspective'">
            <xsl:value-of select="'perspectiveLeft'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-bottom-left' and uof:方向_C63C/uof:方式_C63D='perspective'">
            <xsl:value-of select="'obliqueBottomLeft'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-bottom' and uof:方向_C63C/uof:方式_C63D='perspective'">
            <xsl:value-of select="'perspectiveAbove'"/>
          </xsl:when>
          <xsl:when test="uof:方向_C63C/uof:角度_C639='to-bottom-right' and uof:方向_C63C/uof:方式_C63D='perspective'">
            <xsl:value-of select="'obliqueBottomRight'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="uof:样式_C634">
        <xsl:choose>
          <xsl:when test="uof:样式_C634='0'">
            <xsl:value-of select="'obliqueTopRight'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='1'">
            <xsl:value-of select="'obliqueTopLeft'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='2'">
            <xsl:value-of select="'obliqueBottomLeft'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='3'">
            <xsl:value-of select="'obliqueTopRight'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='4'">
            <xsl:value-of select="'perspectiveContrastingLeftFacing'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='5'">
            <xsl:value-of select="'perspectiveContrastingRightFacing'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='6'">
            <xsl:value-of select="'perspectiveAbove'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='7'">
            <xsl:value-of select="'obliqueBottomLeft'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='8'">
            <xsl:value-of select="'perspectiveHeroicExtremeLeftFacing'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='9'">
            <xsl:value-of select="'perspectiveHeroicExtremeRightFacing'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='10'">
            <xsl:value-of select="'obliqueTopRight'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='11'">
            <xsl:value-of select="'perspectiveBelow'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='12'">
            <xsl:value-of select="'perspectiveRelaxed'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='13'">
            <xsl:value-of select="'perspectiveRelaxed'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='14'">
            <xsl:value-of select="'perspectiveAbove'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='15'">
            <xsl:value-of select="'obliqueBottomLeft'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='16'">
            <xsl:value-of select="'isometricOffAxis2Left'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='17'">
            <xsl:value-of select="'isometricOffAxis1Right'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='18'">
            <xsl:value-of select="'obliqueTopRight'"/>
          </xsl:when>
          <xsl:when test="uof:样式_C634='19'">
            <xsl:value-of select="'perspectiveBelow'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--2015-03-24，wudi，修复阴影转换BUG，start-->
  <!--cxl,2012.2.24增加阴影效果转换模板-->
  <xsl:template match="图:阴影_8051">
    <xsl:if test="./@是否显示阴影_C61C='true'">
      <a:effectLst>
        <a:outerShdw>
          <xsl:if test="./uof:偏移量_C61B/dist">
            <xsl:attribute name="dist">
              <xsl:value-of select="round(./uof:偏移量_C61B/dist)"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="./uof:偏移量_C61B/dir">
            <xsl:attribute name="dir">
              <xsl:value-of select="round(./uof:偏移量_C61B/dir)"/>
            </xsl:attribute>
          </xsl:if>
          <!--<xsl:choose>
            <xsl:when test="./@类型_C61D='3'or ./@类型_C61D='11'">
              <xsl:attribute name="kx">
                <xsl:value-of select="'-1200000'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="./@类型_C61D='shaperelative'">
              <xsl:attribute name="kx">
                <xsl:value-of select="'1200000'"/>
              </xsl:attribute>
            </xsl:when>
            --><!--<xsl:when test="./@类型_C61D=''">
              <xsl:attribute name="kx">
                <xsl:value-of select="'800400'"/>
              </xsl:attribute>
            </xsl:when>--><!--
            <xsl:when test="./@类型_C61D='single'">
              <xsl:attribute name="kx">
                <xsl:value-of select="'-800400'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
            </xsl:otherwise>
          </xsl:choose>-->
          <a:srgbClr>
            <xsl:choose>
              <xsl:when test="contains(./@颜色_C61E,'#')">
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-after(./@颜色_C61E,'#')"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="./@颜色_C61E='auto' or not(./@颜色_C61E)">
                <xsl:attribute name="val">
                  <xsl:value-of select="'808080'"/>
                </xsl:attribute>
              </xsl:when>
            </xsl:choose>
            <xsl:choose>
              <xsl:when test="./@透明度_C61F='49'">
                <a:alpha>
                  <xsl:attribute name="val">
                    <!--cxl,2012.4.9UOF中只有50\100这两种情况，原来的计算公式是round((./@透明度_C61F div 256) * 100000)-->
                    <xsl:value-of select="'50%'"/>
                  </xsl:attribute>
                </a:alpha>
              </xsl:when>
              <xsl:when test="./@透明度_C61F='0'">
                <a:alpha>
                  <xsl:attribute name="val">
                    <xsl:value-of select="'100%'"/>
                  </xsl:attribute>
                </a:alpha>
              </xsl:when>
            </xsl:choose>
          </a:srgbClr>
        </a:outerShdw>
      </a:effectLst>
    </xsl:if>
  </xsl:template>
  <!--end-->

  <!--<xsl:template name="pPrWithpStyle">
    <w:pPr>      
      <xsl:call-template name="pPrWithChange"/>
      <xsl:if test="following::字:句[1]/字:锚点/uof:锚点_C644/字:位置/字:水平/字:相对/@字:参考点">
        <w:jc>
          <xsl:attribute name="w:val">
            <xsl:value-of select="following::字:句[1]/字:锚点/uof:锚点_C644/字:位置/字:水平/字:相对/@字:参考点"/>
          </xsl:attribute>
        </w:jc>
      </xsl:if>
      <xsl:if test="not(parent::字:段落/字:句/字:文本串) and not(字:句属性)">
        <xsl:apply-templates select="following-sibling::字:句[字:句属性][1]/字:句属性" mode="rpr"/>
      </xsl:if>
    </w:pPr>
  </xsl:template>-->

  <!--<xsl:template name="pPrWithChange">
    <xsl:if test="not(following-sibling::字:修订开始[@字:类型='format']/字:段落属性)">
      <xsl:call-template name="pPr"/>
    </xsl:if>
    <xsl:if test="following-sibling::字:修订开始[@字:类型='format']/字:段落属性">
      <xsl:apply-templates select="following-sibling::字:修订开始[@字:类型='format']/字:段落属性" mode="withRev"/>
      <w:pPrChange>
        <xsl:call-template name="pPr"/>
      </w:pPrChange>
    </xsl:if>
  </xsl:template>
  <xsl:template match="字:段落属性" mode="withRev">
    <xsl:call-template name="pPrWithpStyle"/>
  </xsl:template>-->
</xsl:stylesheet>
