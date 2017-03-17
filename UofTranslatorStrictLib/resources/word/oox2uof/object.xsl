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
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:xs="http://www.w3.org/2001/XMLSchema" 
  xmlns:fn="http://www.w3.org/2005/xpath-functions" 
  xmlns:xdt="http://www.w3.org/2005/xpath-datatypes" 
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:对象="http://schemas.uof.org/cn/2009/objects" 
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
  xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart"
  xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing"
  xmlns:dgm="http://purl.oclc.org/ooxml/drawingml/diagram"
  xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
  xmlns:w10="urn:schemas-microsoft-com:office:word" 
  xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" 
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" 
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
  xmlns:a15="http://schemas.microsoft.com/office/drawing/2012/main"
  xmlns:pic="http://purl.oclc.org/ooxml/drawingml/picture"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
  <!--zhaobj 2011/2/25 加入填充<xsl:import href="fill.xsl"/>-->
 
  <!--<xsl:import href="paragraph.xsl"/>-->
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
  <!--cxl2011/12/1-->
  <xsl:template name="object">
    <xsl:param name="objFrom"/>
    
    <!--2014-03-06，wudi，修改前缀wps为wp，start-->
    <xsl:apply-templates select="//wp:wsp| //pic:pic" mode="obj">
      <xsl:with-param name="picFrom" select="$objFrom"/>
    </xsl:apply-templates>
    <!--end-->

	  <xsl:apply-templates select="//v:background" mode="obj">
		  <xsl:with-param name="picFrom" select="$objFrom"/>
	  </xsl:apply-templates>

    <!--2013-04-08，wudi，修复图片转换的BUG，没有考虑节点为w:pict，子节点为v:shape的情况，所有number计数增加对v:shape节点的统计，start-->
    <xsl:apply-templates select ="//v:shape" mode ="obj">
      <xsl:with-param name ="picFrom" select ="$objFrom"/>
    </xsl:apply-templates>
    <!--end-->

    <!--2013-04-19，wudi，增加对chart的转换，转成图片，start-->
    <xsl:apply-templates select ="//c:chart" mode ="obj">
      <xsl:with-param name ="picFrom" select ="$objFrom"/>
    </xsl:apply-templates>
    <!--end-->

    <!--2014-04-10，wudi，增加对SmartArt里包含图片填充的处理，start-->
    <xsl:apply-templates select ="//dgm:relIds" mode ="obj">
      <xsl:with-param name ="picFrom" select ="'SmartArt'"/>
    </xsl:apply-templates>
    <!--end-->
       
    <!--<xsl:variable name="number">
      <xsl:if test="document('word/document.xml')/w:document//w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic">
        <xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline "/>
      </xsl:if>
    </xsl:variable>-->
    <!--<xsl:variable name="number">
      <xsl:for-each select ="//v:rect | //wp:anchor | //wp:inline">
        <xsl:number format="1" level="any" count=" v:rect | wp:anchor | wp:inline"/>
      </xsl:for-each>
    </xsl:variable>-->
   
      
    <!--<xsl:param name ="objFrom"/>
    <xsl:for-each select="//w:pict|//w:drawing">
      <xsl:choose>
       --><!--zhaobj vml删除
       <xsl:when test="name(.)='w:pict'">
          <xsl:call-template name="pictObj"/>
        </xsl:when>--><!--
        <xsl:when test="name(.)='w:drawing'">
          <xsl:call-template name="drawing">
            <xsl:with-param name ="drawingFrom" select ="$objFrom"/>
          </xsl:call-template>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>-->
  </xsl:template>

  <!--2014-04-10，wudi，增加对SmartArt里包含图片填充的处理，start-->
  <xsl:template match ="dgm:relIds" mode ="obj">
    <xsl:param name ="picFrom"/>

    <xsl:variable name ="rcs" select ="@r:cs"/>
    <xsl:variable name ="rId" select ="substring-after($rcs,'rId')"/>
    <xsl:variable name ="gId" select ="number($rId) + 1"/>
    <xsl:variable name ="number">
      <xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline | v:shape"/>
    </xsl:variable>
    <xsl:for-each select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship">
      <xsl:if test ="@Id = concat('rId',$gId)">
        <xsl:variable name="target">
          <xsl:value-of select="@Target"/>
        </xsl:variable>
        <xsl:variable name="drawing">
          <xsl:value-of select="substring-after($target,'/')"/>
        </xsl:variable>

        <!--2014-04-18，wudi，新增模板，用于处理SmartArt包含图片填充与不包含图片填充差异问题，start-->
        <xsl:apply-templates select="document(concat('word/diagrams/',$drawing))/dsp:drawing" mode="Obj">
          <xsl:with-param name="drawing" select="$drawing"/>
          <xsl:with-param name="number" select="$number"/>
          <xsl:with-param name="picFrom" select="$picFrom"/>
        </xsl:apply-templates>
        <!--end-->
        
      </xsl:if>
      
    </xsl:for-each>

  </xsl:template>

  <!--2014-04-18，wudi，新增模板，用于处理SmartArt包含图片填充与不包含图片填充差异问题，start-->
  <xsl:template match="dsp:drawing" mode="Obj">
    <xsl:param name="drawing"/>
    <xsl:param name="number"/>
    <xsl:param name="picFrom"/>

    <xsl:if test="//a:blip[last()]">
      <xsl:for-each select="document(concat('word/diagrams/_rels/',$drawing,'.rels'))/rel:Relationships/rel:Relationship">
        <对象:对象数据_D701 是否内嵌_D705="true">
          <xsl:attribute name="标识符_D704">
            <xsl:variable name="tmp">
              <xsl:value-of select="number(substring-after(@Id,'rId'))"/>
            </xsl:variable>

            <xsl:value-of select ="concat($picFrom,'Obj',($number - 1) * 9 + $tmp)"/>
          </xsl:attribute>
          <xsl:attribute name="公共类型_D706">
            <xsl:variable name ="tmp2">
              <xsl:value-of select="substring-after(@Target,'media/')"/>
            </xsl:variable>
            <xsl:variable name="picType">
              <xsl:value-of select="substring-after($tmp2,'.')"/>
            </xsl:variable>

            <xsl:call-template name ="objtype">
              <xsl:with-param name ="val" select ="$picType"/>
            </xsl:call-template>
          </xsl:attribute>
          <对象:路径_D703>
            <xsl:value-of select ="concat('.\data\',substring-after(@Target,'media/'))"/>
          </对象:路径_D703>
        </对象:对象数据_D701>
      </xsl:for-each>
    </xsl:if>
    
  </xsl:template>
  <!--end-->

  <!--2013-04-19，wudi，增加对chart的转换，转成图片，start-->
  <xsl:template match ="c:chart" mode ="obj">
    <xsl:param name ="picFrom"/>
    <对象:对象数据_D701 是否内嵌_D705="true">
      <xsl:variable name="number">
        <xsl:number format="1" level="any" count=" v:rect | wp:anchor | wp:inline | v:shape"/>
      </xsl:variable>
      <xsl:attribute name ="标识符_D704">
        <xsl:value-of select ="concat($picFrom,'Obj',$number * 2)"/>
      </xsl:attribute>
      <xsl:attribute name ="公共类型_D706">
        <xsl:value-of select ="'jpg'"/>
      </xsl:attribute>
      <对象:路径_D703>
        <xsl:variable name ="rid">
          <xsl:value-of select ="@r:id"/>
        </xsl:variable>
        <xsl:variable name ="tmp">
          <xsl:value-of select ="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id =$rid]/@Target"/>
        </xsl:variable>
        <xsl:variable name ="path">
          <xsl:value-of select ="concat('.\data\',substring-after($tmp,'charts/'),'.jpg')"/>
        </xsl:variable>
        <xsl:value-of select ="$path"/>
      </对象:路径_D703>
    </对象:对象数据_D701>
  </xsl:template>
  <!--end-->

  <!--2014-05-20，wudi，修复页眉区域图片转换丢失的BUG，start-->
  <!--2013-04-08，wudi，修复图片转换的BUG，没有考虑节点为w:pict，子节点为v:shape的情况，所有number计数增加对v:shape节点的统计，start-->
  <xsl:template match ="v:shape" mode ="obj">
    <xsl:param name="picFrom"/>
    
    <xsl:if test ="./v:imagedata">
      <对象:对象数据_D701 是否内嵌_D705="true">
        <xsl:variable name="picId">
          <xsl:value-of select="./v:imagedata/@r:id"/>
        </xsl:variable>
        <xsl:variable name="picName">
          <xsl:choose>
            <xsl:when test="$picFrom='document'">
              <xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
            </xsl:when>
            <xsl:when test="$picFrom='comments'">
              <xsl:value-of select="document('word/_rels/comments.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
            </xsl:when>
            <xsl:when test="$picFrom='endnotes'">
              <xsl:value-of select="document('word/_rels/endnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
            </xsl:when>
            <xsl:when test="$picFrom='footnotes'">
              <xsl:value-of select="document('word/_rels/footnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
            </xsl:when>
            <xsl:when test="contains($picFrom,'header')">
              <xsl:variable name="hn" select="concat('word/_rels/',$picFrom,'.xml.rels')"/>
              <xsl:value-of select="document($hn)/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
            </xsl:when>
            <xsl:when test="contains($picFrom,'footer')">
              <xsl:variable name="fn" select="concat('word/_rels/',$picFrom,'.xml.rels')"/>
              <xsl:value-of select="document($fn)/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
            </xsl:when>
            <!--暂时先这样-->
            <xsl:when test="$picFrom='grpsp'">
              <xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
            </xsl:when>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="pictype">
          <xsl:value-of select="substring-after($picName,'.')"/>
        </xsl:variable>
        <xsl:variable name="number">
          <xsl:number format="1" level="any" count=" v:rect | wp:anchor | wp:inline | v:shape"/>
        </xsl:variable>
        <xsl:attribute name="标识符_D704">
          <xsl:value-of select ="concat($picFrom,'Obj',$number * 2)"/>
        </xsl:attribute>
        <xsl:attribute name="公共类型_D706">
          <xsl:call-template name ="objtype">
            <xsl:with-param name ="val" select ="$pictype"/>
          </xsl:call-template>
        </xsl:attribute>
        <对象:路径_D703>
          <xsl:value-of select ="concat('.\data\',substring-after($picName,'media/'))"/>
        </对象:路径_D703>
      </对象:对象数据_D701>
    </xsl:if>
  </xsl:template>
  <!--end-->
  <!--end-->
  
  <!--2014-03-06，wudi，修改前缀wps为wp，start-->
  <xsl:template match="wp:wsp|pic:pic" mode="obj">
    <xsl:param name="picFrom"/>
    <xsl:variable name="number">
      <xsl:number format="1" level="any" count=" v:rect | wp:anchor | wp:inline | v:shape"/>
    </xsl:variable>
    <xsl:variable name="style" select="@style"/>

    <!--2013-03-28，wudi，修复BUG#2745集成测试00X到UOF方向-模板1图片由一个变为二个，start-->
    <xsl:if test ="not(ancestor::mc:Fallback) and not(descendant::pic:pic)">
      <xsl:if test=".//pic:blipFill | .//wp:spPr/a:blipFill">
        <对象:对象数据_D701 是否内嵌_D705="true">
          <xsl:variable name="picId">
            <xsl:if test=".//pic:blipFill/a:blip/@r:embed">
              <xsl:value-of select=".//pic:blipFill/a:blip/@r:embed"/>
            </xsl:if>
            <xsl:if test=".//wp:spPr/a:blipFill/a:blip/@r:embed">
              <xsl:value-of select=".//wp:spPr/a:blipFill/a:blip/@r:embed"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name="picName">
            <xsl:choose>
              <xsl:when test="$picFrom='document'">
                <xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="$picFrom='comments'">
                <xsl:value-of select="document('word/_rels/comments.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="$picFrom='endnotes'">
                <xsl:value-of select="document('word/_rels/endnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="$picFrom='footnotes'">
                <xsl:value-of select="document('word/_rels/footnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="contains($picFrom,'header')">
                <xsl:variable name="hn" select="concat('word/_rels/',$picFrom,'.xml.rels')"/>
                <xsl:value-of select="document($hn)/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="contains($picFrom,'footer')">
                <xsl:variable name="fn" select="concat('word/_rels/',$picFrom,'.xml.rels')"/>
                <xsl:value-of select="document($fn)/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <!--暂时先这样-->
              <xsl:when test="$picFrom='grpsp'">
                <xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="pictype">
            <xsl:value-of select="substring-after($picName,'.')"/>
          </xsl:variable>
          <xsl:attribute name="标识符_D704">
            <xsl:if test="$picFrom = 'grpsp'">
              <xsl:if test="document('word/document.xml')//wp:cNvPr/@id">
                <xsl:value-of select="concat('grpsppicObj',document('word/document.xml')//wp:cNvPr/@id)"/>
              </xsl:if>
              <xsl:if test="document('word/document.xml')//wpg:cNvPr/@id">
                <xsl:value-of select="concat('grpsppicObj',document('word/document.xml')//wpg:cNvPr/@id)"/>
              </xsl:if>
            </xsl:if>
            <xsl:if test="$picFrom != 'grpsp'">
              <xsl:value-of select="concat($picFrom,'Obj',$number * 2)"/>
              <!--substring-after($picName,'/')-->
            </xsl:if>
          </xsl:attribute>
          <xsl:attribute name="公共类型_D706">
            <xsl:call-template name="objtype">
              <xsl:with-param name="val" select="$pictype"/>
            </xsl:call-template>
          </xsl:attribute>
          <!--<xsl:attribute name="私有类型_D707">
          -->
          <!--代码内容有待修改，私有类型有哪些格式？？？？？-->
          <!--
          <xsl:call-template name="objtype">
            <xsl:with-param name="val" select="$pictype"/>
          </xsl:call-template>
        </xsl:attribute>-->
          <对象:路径_D703>
            <xsl:value-of select="concat('.\data\',substring-after($picName,'/'))"/>
          </对象:路径_D703>
        </对象:对象数据_D701>
      </xsl:if>
    </xsl:if>
    <!--end-->
    
  </xsl:template>
  <!--end-->


	<!--页面填充中的图片填充-->
	<xsl:template match="v:background" mode="obj">
		<xsl:param name="picFrom"/>
		<xsl:variable name="number">1</xsl:variable>
			<对象:对象数据_D701 是否内嵌_D705="true">
				<xsl:variable name="picId">
						<xsl:value-of select="//v:background/v:fill/@r:id"/>
				</xsl:variable>
				<xsl:variable name="picName">
					<xsl:choose>
						<xsl:when test="$picFrom='document'">
							<xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
						<xsl:when test="$picFrom='comments'">
							<xsl:value-of select="document('word/_rels/comments.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
						<xsl:when test="$picFrom='endnotes'">
							<xsl:value-of select="document('word/_rels/endnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
						<xsl:when test="$picFrom='footnotes'">
							<xsl:value-of select="document('word/_rels/footnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
						<xsl:when test="contains($picFrom,'header')">
							<xsl:variable name="hn" select="concat('word/_rels/',$picFrom,'.xml.rels')"/>
							<xsl:value-of select="document($hn)/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
						<xsl:when test="contains($picFrom,'footer')">
							<xsl:variable name="fn" select="concat('word/_rels/',$picFrom,'.xml.rels')"/>
							<xsl:value-of select="document($fn)/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
						<!--暂时先这样-->
						<xsl:when test="$picFrom='grpsp'">
							<xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
					</xsl:choose>
				</xsl:variable>
				<xsl:variable name="pictype">
					<xsl:value-of select="substring-after($picName,'.')"/>
				</xsl:variable>
				<xsl:attribute name="标识符_D704">
					<xsl:if test="$picFrom != 'grpsp'">
						<xsl:value-of select="concat($picFrom,'Obj',$number)"/>
						<!--substring-after($picName,'/')-->
					</xsl:if>
				</xsl:attribute>
				<xsl:attribute name="公共类型_D706">
					<xsl:call-template name="objtype">
						<xsl:with-param name="val" select="$pictype"/>
					</xsl:call-template>
				</xsl:attribute>
				<对象:路径_D703>
					<xsl:value-of select="concat('.\data\',substring-after($picName,'/'))"/>
				</对象:路径_D703>
			</对象:对象数据_D701>
	
		<!--<xsl:if test ="//v:background/v:fill/@type='frame' or  //v:background/v:fill/@type='tile'">
			<uof:其他对象 uof:locID="u0036" uof:attrList="标识符 内嵌 公共类型 私有类型" uof:标识符="backgroundpic006" uof:内嵌="false">
				<xsl:variable name ="BackgroundpicId">
					<xsl:value-of select ="//v:background/v:fill/@r:id"/>
				</xsl:variable>
				<xsl:variable name ="Backgroundpicname">
					<xsl:value-of select ="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$BackgroundpicId]/@Target"/>
				</xsl:variable>
				<xsl:variable name ="Backgroundpictype">
					<xsl:value-of select ="substring-after($Backgroundpicname,'.')"/>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test ="$Backgroundpictype='emf'">
						<xsl:attribute name ="uof:私有类型">
							<xsl:value-of select ="'emf'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test ="$Backgroundpictype='jpeg'">
						<xsl:attribute name ="uof:私有类型">
							<xsl:value-of select ="'jpeg'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="uof:公共类型">
							<xsl:value-of select ="$Backgroundpictype"/>
						</xsl:attribute>
					</xsl:otherwise>
				</xsl:choose>
				<o2upic:picture>
					<xsl:attribute name="target">
						<xsl:value-of select ="concat('word/',$Backgroundpicname)"/>
					</xsl:attribute>
				</o2upic:picture>
			</uof:其他对象>
		</xsl:if>-->
	</xsl:template>
  <xsl:template name="objtype">
    <xsl:param name="val"/>
    <xsl:choose>
      <xsl:when test="$val='jpeg'">
        <xsl:value-of select="'jpeg'"/>
      </xsl:when>
      <xsl:when test="$val='jpg'">
        <xsl:value-of select="'jpg'"/>
      </xsl:when>
      <xsl:when test="$val='text'">
        <xsl:value-of select="'text'"/>
      </xsl:when>
      <xsl:when test="$val='xml'">
        <xsl:value-of select="'xml'"/>
      </xsl:when>
      <xsl:when test="$val='html'">
        <xsl:value-of select="'html'"/>
      </xsl:when>
      <xsl:when test="$val='wav'">
        <xsl:value-of select="'wav'"/>
      </xsl:when>
      <xsl:when test="$val='midi'">
        <xsl:value-of select="'midi'"/>
      </xsl:when>
      <xsl:when test="$val='ra'">
        <xsl:value-of select="'ra'"/>
      </xsl:when>
      <xsl:when test="$val='au'">
        <xsl:value-of select="'au'"/>
      </xsl:when>
      <xsl:when test="$val='mp3'">
        <xsl:value-of select="'mp3'"/>
      </xsl:when>
      <xsl:when test="$val='snd'">
        <xsl:value-of select="'snd'"/>
      </xsl:when>
      <xsl:when test="$val='png'">
        <xsl:value-of select="'png'"/>
      </xsl:when>
      <xsl:when test="$val='bmp'">
        <xsl:value-of select="'bmp'"/>
      </xsl:when>
      <xsl:when test="$val='pbm'">
        <xsl:value-of select="'pbm'"/>
      </xsl:when>
      <xsl:when test="$val='ras'">
        <xsl:value-of select="'ras'"/>
      </xsl:when>
      <xsl:when test="$val='gif'">
        <xsl:value-of select="'gif'"/>
      </xsl:when>
      <xsl:when test="$val='svg'">
        <xsl:value-of select="'svg'"/>
      </xsl:when>
      <xsl:when test="$val='avi'">
        <xsl:value-of select="'avi'"/>
      </xsl:when>
      <xsl:when test="$val='mpeg1'">
        <xsl:value-of select="'mpeg1'"/>
      </xsl:when>
      <xsl:when test="$val='mpeg2'">
        <xsl:value-of select="'mpeg2'"/>
      </xsl:when>
      <xsl:when test="$val='mpeg4'">
        <xsl:value-of select="'mpeg4'"/>
      </xsl:when>
      <xsl:when test="$val='qt'">
        <xsl:value-of select="'qt'"/>
      </xsl:when>
      <xsl:when test="$val='rm'">
        <xsl:value-of select="'rm'"/>
      </xsl:when>
      <xsl:when test="$val='asf'">
        <xsl:value-of select="'asf'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$val"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--<xsl:template name="bodyAnchorYdy">
    <xsl:variable name ="number">
      <xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline"/>
    </xsl:variable>
    <字:锚点 uof:locID="t0110" uof:attrList="标识符 类型">
      <xsl:attribute name="字:类型">
        <xsl:if test="w10:wrap[type='none'] and w10:anchorlock">
          <xsl:value-of select="'inline'"/>
        </xsl:if>
        <xsl:if test="not(w10:wrap[type='none'] and w10:anchorlock)">
          <xsl:value-of select="'normal'"/>
        </xsl:if>
      </xsl:attribute>
      <字:锚点属性 uof:locID="t0111">
        <xsl:variable name="style" select=" @style"/>
        <xsl:if test="@style">
          <xsl:if test="contains(@style,'width') or contains(@style,'WIDTH') or contains(@style,'height') or contains(@style,'HEIGHT')">
            <xsl:variable name="w-temp" select="substring-after($style,'width:')"/>
            <xsl:variable name="w-temp1" select="substring-before($w-temp,'pt;')"/>
            <xsl:variable name="h-temp" select="substring-after($style,'height:')"/>
            <xsl:variable name="h-temp1" select="substring-before($h-temp,'pt;')"/>
            <字:宽度 uof:locID="t0112">
              <xsl:value-of select="$w-temp1"/>
            </字:宽度>
            <字:高度 uof:locID="t0113">
              <xsl:value-of select="$h-temp1"/>
            </字:高度>
          </xsl:if>
        </xsl:if>
        <字:位置 uof:locID="t0114">
          <xsl:variable name="H-temp">
            <xsl:if test="contains(@style,'mso-position-horizontal')">
              <xsl:variable name="temp" select="substring-after($style,'mso-position-horizontal:')"/>
              <xsl:value-of select="substring-before($temp,';')"/>
            </xsl:if>
            <xsl:if test="not (contains(@style,'mso-position-horizontal'))">
              <xsl:value-of select="'none'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name="HR-temp">
            <xsl:if test="contains(@style,'mso-position-horizontal-relative')">
              <xsl:variable name="temp" select="substring-after($style,'mso-position-horizontal-relative:')"/>
              <xsl:value-of select="substring-before($temp,';')"/>
            </xsl:if>
            <xsl:if test="not (contains(@style,'mso-position-horizontal-relative'))">
              <xsl:value-of select="'none'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name="V-temp">
            <xsl:if test="contains(@style,'mso-position-vertical')">
              <xsl:variable name="temp" select="substring-after($style,'mso-position-vertical:')"/>
              <xsl:value-of select="substring-before($temp,';')"/>
            </xsl:if>
            <xsl:if test="not (contains(@style,'mso-position-vertical'))">
              <xsl:value-of select="'none'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name="VR-temp">
            <xsl:if test="contains(@style,'mso-position-vertical-relative')">
              <xsl:variable name="temp" select="substring-after($style,'mso-position-vertical-relative:')"/>
              <xsl:value-of select="substring-before($temp,';')"/>
            </xsl:if>
            <xsl:if test="not (contains(@style,'mso-position-vertical-relative'))">
              <xsl:value-of select="'none'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name="ML-temp">
            <xsl:variable name="temp" select="substring-after($style,'margin-left:')"/>
            <xsl:value-of select="substring-before($temp,';')"/>
          </xsl:variable>
          <xsl:variable name="MT-temp">
            <xsl:variable name="temp" select="substring-after($style,'margin-top:')"/>
            <xsl:value-of select="substring-before($temp,';')"/>
          </xsl:variable>
          <字:水平 uof:locID="t0176" uof:attrList="相对于">
            <xsl:attribute name="字:相对于">
              <xsl:choose>
                <xsl:when test="$HR-temp='margin'">
                  <xsl:value-of select="'margin'"/>
                </xsl:when>
                <xsl:when test="$HR-temp='page'">
                  <xsl:value-of select="'page'"/>
                </xsl:when>
                <xsl:when test="$HR-temp='char'">
                  <xsl:value-of select="'char'"/>
                </xsl:when>
                <xsl:when test="$HR-temp='text'">
                  <xsl:value-of select="'column'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'column'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:if test="not (contains(@style,'mso-position-horizontal'))">
              <字:绝对 uof:locID="t0177" uof:attrList="值">
                <xsl:attribute name="字:值">
                  <xsl:value-of select ="$ML-temp"/>
                </xsl:attribute>
              </字:绝对>
            </xsl:if>
            <xsl:if test="contains(@style,'mso-position-horizontal')">
              <字:相对 uof:locID="t0178" uof:attrList="参考点 值">
                <xsl:attribute name="字:参考点">
                  <xsl:choose>
                    <xsl:when test="$H-temp='center'">
                      <xsl:value-of select="'center'"/>
                    </xsl:when>
                    <xsl:when test="$H-temp='left'">
                      <xsl:value-of select="'left'"/>
                    </xsl:when>
                    <xsl:when test="$H-temp='right'">
                      <xsl:value-of select="'right'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'left'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
                <xsl:attribute name="字:值">
                  <xsl:value-of select ="$ML-temp"/>
                </xsl:attribute>
              </字:相对>
            </xsl:if>
          </字:水平>
          <字:垂直 uof:locID="t0179" uof:attrList="相对于">
            <xsl:attribute name="字:相对于">
              <xsl:choose>
                <xsl:when test="$VR-temp='margin'">
                  <xsl:value-of select="'margin'"/>
                </xsl:when>
                <xsl:when test="$VR-temp='page'">
                  <xsl:value-of select="'page'"/>
                </xsl:when>
                <xsl:when test="$VR-temp='line'">
                  <xsl:value-of select="'line'"/>
                </xsl:when>
                <xsl:when test="$VR-temp='text'">
                  <xsl:value-of select="'paragraph'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'paragraph'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:if test="not (contains(@style,'mso-position-vertical'))">
              <字:绝对 uof:locID="t0180" uof:attrList="值">
                <xsl:attribute name="字:值">
                  <xsl:value-of select ="$MT-temp"/>
                </xsl:attribute>
              </字:绝对>
            </xsl:if>
            <xsl:if test="contains(@style,'mso-position-vertical')">
              <字:相对 uof:locID="t0181" uof:attrList="参考点 值">
                <xsl:attribute name="字:参考点">
                  <xsl:choose>
                    <xsl:when test="$V-temp='center'">
                      <xsl:value-of select="'center'"/>
                    </xsl:when>
                    <xsl:when test="$V-temp='top'">
                      <xsl:value-of select="'top'"/>
                    </xsl:when>
                    <xsl:when test="$V-temp='bottom'">
                      <xsl:value-of select="'bottom'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'top'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
                <xsl:attribute name="字:值">
                  <xsl:value-of select ="$MT-temp"/>
                </xsl:attribute>
              </字:相对>
            </xsl:if>
          </字:垂直>
        </字:位置>
        <字:绕排 uof:locID="t0115" uof:attrList="绕排方式 环绕文字 绕排顶点" 字:环绕文字="both">
          <xsl:variable name="wrapType">
            <xsl:if test="w10:wrap">
              <xsl:value-of select="w10:wrap/@type"/>
            </xsl:if>
            <xsl:if test="w10:wrap">
              <xsl:value-of select="'none'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:variable name="index">
            <xsl:if test="contains(@style,'z-index')">
              <xsl:variable name="temp" select="substring-after($style,'z-index:')"/>
              <xsl:value-of select="substring-before($temp,';')"/>
            </xsl:if>
            <xsl:if test="not (contains(@style,'z-index'))">
              <xsl:value-of select="'0'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="$wrapType='tight'">
              <xsl:attribute name="字:绕排方式">
                <xsl:value-of select ="'tight'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="$wrapType='topAndBottom'">
              <xsl:attribute name="字:绕排方式">
                <xsl:value-of select ="'top-bottom'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="$wrapType='through'">
              <xsl:attribute name="字:绕排方式">
                <xsl:value-of select ="'through'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="$wrapType='square'">
              <xsl:attribute name="字:绕排方式">
                <xsl:value-of select ="'square'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="$index &lt; '-250000000'">
              <xsl:attribute name="字:绕排方式">
                <xsl:value-of select ="'behindtext'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="$index &gt; '250000000'">
              <xsl:attribute name="字:绕排方式">
                <xsl:value-of select ="'infrontoftext'"/>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </字:绕排>
      </字:锚点属性>
      <字:图形 uof:locID="t0120" uof:attrList="图形引用">
        <xsl:attribute name="字:图形引用">
          <xsl:value-of select ="concat('OBJ',$number * 2 +1)"/>
        </xsl:attribute>
      </字:图形>
    </字:锚点>
  </xsl:template>

  <xsl:template name="bodyAnchorPic">

    <xsl:param name="picType"/>
    <xsl:param name="filename"/>
    <xsl:variable name ="originalId">
      <xsl:value-of select =".//a:blip/@r:embed"/>
    </xsl:variable>
    <xsl:variable name ="number">
      <xsl:if test ="document('word/document.xml')/w:document//w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic">
        <xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline"/>
      </xsl:if>
      <xsl:if test ="not(document('word/document.xml')/w:document//w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic)">
        <xsl:call-template name ="numberForEveryPart">
          <xsl:with-param name ="numPart" select ="$filename"/>
          <xsl:with-param name ="ooxId" select ="$originalId"/>
        </xsl:call-template>
      </xsl:if>
    </xsl:variable>
    <字:锚点 uof:locID="t0110" uof:attrList="标识符 类型">
      <xsl:attribute name="字:类型">
        <xsl:if test="$picType='inline'">
          <xsl:value-of select="'inline'"/>
        </xsl:if>
        <xsl:if test="$picType='anchor'">
          <xsl:value-of select="'normal'"/>
        </xsl:if>
      </xsl:attribute>
      <字:锚点属性 uof:locID="t0111">
        <字:宽度 uof:locID="t0112">
          <xsl:value-of select="wp:extent/@cx div 12700"/>
        </字:宽度>
        <字:高度 uof:locID="t0113">
          <xsl:value-of select="wp:extent/@cy div 12700"/>
        </字:高度>
        <xsl:if test="wp:positionH or wp:positionV">
          <字:位置 uof:locID="t0114">
            <xsl:if test="wp:positionH">
              <字:水平 uof:locID="t0176" uof:attrList="相对于">
                <xsl:attribute name="字:相对于">
                  <xsl:call-template name ="positionHrelative">
                    <xsl:with-param name ="val" select ="wp:positionH/@relativeFrom"/>
                  </xsl:call-template>
                </xsl:attribute>
                <xsl:if test="wp:positionH/wp:posOffset">
                  <字:绝对 uof:locID="t0177" uof:attrList="值">
                    <xsl:attribute name="字:值">
                      <xsl:value-of select ="wp:positionH/wp:posOffset div 12700"/>
                    </xsl:attribute>
                  </字:绝对>
                </xsl:if>
                <xsl:if test="wp:positionH/wp:align">
                  <字:相对 uof:locID="t0178" uof:attrList="参考点 值">
                    <xsl:attribute name="字:参考点">
                      <xsl:call-template name ="positionHalign">
                        <xsl:with-param name ="val" select ="wp:positionH/wp:align"/>
                      </xsl:call-template>
                   </xsl:attribute>
                    <xsl:attribute name="字:值">
                      <xsl:value-of select="'0'"/>
                    </xsl:attribute>
                  </字:相对>
                </xsl:if>
              </字:水平>
            </xsl:if>
            <xsl:if test="wp:positionV">
              <字:垂直 uof:locID="t0179" uof:attrList="相对于">
                <xsl:attribute name="字:相对于">
                  <xsl:call-template name ="positionVrelative">
                    <xsl:with-param name ="val" select ="wp:positionV/@relativeFrom"/>
                  </xsl:call-template>
                </xsl:attribute>
                <xsl:if test="wp:positionV/wp:posOffset">
                  <字:绝对 uof:locID="t0180" uof:attrList="值">
                    <xsl:attribute name="字:值">
                      <xsl:value-of select="wp:positionV/wp:posOffset div 12700"/>
                    </xsl:attribute>
                  </字:绝对>
                </xsl:if>
                <xsl:if test="wp:positionV/wp:align">
                  <字:相对 uof:locID="t0181" uof:attrList="参考点 值">
                    <xsl:attribute name="字:参考点">
                      <xsl:call-template name ="positionValign">
                        <xsl:with-param name ="val" select ="wp:positionV/wp:align"/>
                      </xsl:call-template>
                    </xsl:attribute>
                    <xsl:attribute name="字:值">
                      <xsl:value-of select="'0'"/>
                    </xsl:attribute>
                  </字:相对>
                </xsl:if>
              </字:垂直>
            </xsl:if>
          </字:位置>
        </xsl:if>
        <字:绕排 uof:locID="t0115" uof:attrList="绕排方式 环绕文字 绕排顶点">
          <xsl:if test="@behindDoc='1'">
            <xsl:attribute name="字:绕排方式">
              <xsl:value-of select ="'behindtext'"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@behindDoc='1'">
            <xsl:attribute name="字:绕排方式">
              <xsl:value-of select ="'infrontoftext'"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="wp:wrapTight">
            <xsl:attribute name="字:绕排方式">
              <xsl:value-of select ="'tight'"/>
            </xsl:attribute>
            <xsl:if test ="wp:wrapTight/@wrapText">
              <xsl:attribute name ="字:环绕文字">
                <xsl:call-template name ="wraptext">
                  <xsl:with-param name ="val" select ="wp:wrapTight/@wrapText"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
          <xsl:if test="wp:wrapTopAndBottom">
            <xsl:attribute name="字:绕排方式">
              <xsl:value-of select ="'top-bottom'"/>
            </xsl:attribute>

          </xsl:if>
          <xsl:if test="wp:wrapThrough">
            <xsl:attribute name="字:绕排方式">
              <xsl:value-of select ="'through'"/>
            </xsl:attribute>
            <xsl:if test ="wp:wrapThrough/@wrapText">
              <xsl:attribute name ="字:环绕文字">
                <xsl:call-template name ="wraptext">
                  <xsl:with-param name ="val" select ="wp:wrapThrough/@wrapText"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
          <xsl:if test="wp:wrapSquare">
            <xsl:attribute name="字:绕排方式">
              <xsl:value-of select ="'square'"/>
            </xsl:attribute>
            <xsl:if test ="wp:wrapSquare/@wrapText">
              <xsl:attribute name ="字:环绕文字">
                <xsl:call-template name ="wraptext">
                  <xsl:with-param name ="val" select ="wp:wrapSquare/@wrapText"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
          
        </字:绕排>
        <xsl:if test="@distT or @distL or @distR or @distB">
          <字:边距 uof:locID="t0116" uof:attrList="上 左 右 下">
            <xsl:attribute name="字:上">
              <xsl:value-of select ="@distT div 9525"/>
            </xsl:attribute>
            <xsl:attribute name="字:左">
              <xsl:value-of select ="@distL div 9525"/>
            </xsl:attribute>
            <xsl:attribute name="字:右">
              <xsl:value-of select ="@distR div 9525"/>
            </xsl:attribute>
            <xsl:attribute name="字:下">
              <xsl:value-of select ="@distB div 9525"/>
            </xsl:attribute>
          </字:边距>
        </xsl:if>
        <xsl:if test="wp:cNvGraphicFramePr/a:graphicFrameLocks/@noResize='1'">
          <字:锁定 uof:locID="t0117" uof:attrList="值" 字:值="true"/>
        </xsl:if>
        <字:保护 uof:locID="t0118" uof:attrList="值" 字:值="false"/>
        <xsl:if test="@allowOverlap='1'">
          <字:允许重叠 uof:locID="t0119" uof:attrList="值" 字:值="true"/>
        </xsl:if>
      </字:锚点属性>
      <字:图形 uof:locID="t0120" uof:attrList="图形引用">
        <xsl:attribute name="字:图形引用">
          <xsl:choose>
            <xsl:when test ="$filename='document'">
              <xsl:value-of select ="concat($filename,'OBJ',$number * 2 +1)"/>
            </xsl:when>
            <xsl:when test ="$filename='comments'">
              <xsl:value-of select ="concat($filename,'OBJ',$number * 2 +1)"/>
            </xsl:when>
            <xsl:when test ="$filename='endnotes'">
              <xsl:value-of select ="concat($filename,'OBJ',$number * 2 +1)"/>
            </xsl:when>
            <xsl:when test ="$filename='footnotes'">
              <xsl:value-of select ="concat($filename,'OBJ',$number * 2 +1)"/>
            </xsl:when>
            <xsl:when test ="contains($filename,'header')">
              <xsl:variable name ="hn" select ="substring-before($filename,'.xml')"/>
              <xsl:value-of select ="concat($hn,'OBJ',$number * 2 +1)"/>
            </xsl:when>
            <xsl:when test ="contains($filename,'footer')">
              <xsl:variable name ="fn" select ="substring-before($filename,'.xml')"/>
              <xsl:value-of select ="concat($fn,'OBJ',$number * 2 +1)"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </字:图形>
    </字:锚点>
  </xsl:template>-->

  <!--<xsl:template name ="numberForEveryPart">
    <xsl:param name ="numPart"/>
    <xsl:param name ="ooxId"/>

    <xsl:choose>
      <xsl:when test ="$numPart='document'">
        <xsl:variable name ="nb">
          <xsl:apply-templates select ="document('word/document.xml')/w:document//w:drawing[(wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or(v:rect//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId)]" mode ="objNum"/>
        </xsl:variable>
        <xsl:value-of select ="$nb"/>
      </xsl:when>
      <xsl:when test ="$numPart='comments'">
        <xsl:variable name ="number">
          <xsl:apply-templates select ="document('word/comments.xml')/w:comments//w:drawing[(wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or(v:rect//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId)]" mode ="objNum"/>
        </xsl:variable>
        <xsl:value-of select ="$number"/>
      </xsl:when>
      <xsl:when test ="$numPart='endnotes'">
        <xsl:variable name ="number">
          <xsl:apply-templates select ="document('word/endnotes.xml')/w:endnotes//w:drawing[(wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or(v:rect//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId)]" mode ="objNum"/>
        </xsl:variable>
      </xsl:when>
      <xsl:when test ="$numPart='footnotes'">
        <xsl:variable name ="number">
          <xsl:apply-templates select ="document('word/footnotes.xml')/w:footnotes//w:drawing[(wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or(v:rect//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId)]" mode ="objNum"/>
        </xsl:variable>
        <xsl:value-of select ="$number"/>
      </xsl:when>
      <xsl:when test ="contains($numPart,'header')">
        <xsl:variable name ="hnn" select ="substring-before($numPart,'.xml')"/>
        <xsl:variable name ="hpath" select ="concat('word/',$numPart)"/>
        <xsl:variable name ="number">
          <xsl:apply-templates select ="document($hpath)/w:hdr//w:drawing[(wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or(v:rect//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId)]" mode ="objNum"/>
        </xsl:variable>
        <xsl:value-of select ="$number"/>
      </xsl:when>
      <xsl:when test ="contains($numPart,'footer')">
        <xsl:variable name ="fnn" select ="substring-before($numPart,'.xml')"/>
        <xsl:variable name ="fpath" select ="concat('word/',$numPart)"/>
        <xsl:variable name ="number">
          --><!--<xsl:apply-templates select ="document($fpath)/w:ftr//((wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) 
                               or (wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId)
                               or(v:rect//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId))" mode ="objNum"/>--><!--          
        </xsl:variable>
        <xsl:value-of select ="$number"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>-->
 <!--
  <xsl:template name="numberForEveryPart">
    <xsl:param name ="numPart"/>
    <xsl:param name ="ooxId"/>
    <xsl:value-of select ="'100'"/>
  </xsl:template>
  -->
  <!--cxl,2012.3.2修改自动编号图片符号-->
  <xsl:template name ="numberingPicture">
    <对象:对象数据_D701>
      
      <!--2013-11-20，wudi，Strcit标准下，w:pict节点的位置不同，修正，start-->
      <xsl:variable name="picId">
        <xsl:value-of select="mc:AlternateContent/mc:Choice/w:pict/v:shape/v:imagedata/@r:id"/>
      </xsl:variable>
      <!--end-->
      
      <xsl:variable name="picName">
        <xsl:value-of select="document('word/_rels/numbering.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
      </xsl:variable>
      <xsl:variable name="pictype">
        <xsl:value-of select="substring-after($picName,'.')"/>
      </xsl:variable>
      <xsl:variable name ="number">
        <xsl:value-of select ="@w:numPicBulletId"/>
      </xsl:variable>
      <xsl:attribute name="标识符_D704">
        <xsl:value-of select ="concat('numbering','Obj',$number * 2)"/>
      </xsl:attribute>
      <xsl:attribute name="公共类型_D706">
        <xsl:call-template name ="objtype">
          <xsl:with-param name ="val" select ="$pictype"/>
        </xsl:call-template>
      </xsl:attribute>
      <对象:路径_D703>
        <xsl:value-of select ="concat('.\data\',substring-after($picName,'media/'))"/>
      </对象:路径_D703>       
    </对象:对象数据_D701>
  </xsl:template>
  
  
  <!--<xsl:template name="lineType">
    <xsl:param name="linetype"/>
    <xsl:choose>
      <xsl:when test="$linetype='none'">
        <xsl:value-of select="'none'"/>
      </xsl:when>
      <xsl:when test="$linetype='single'">
        <xsl:value-of select="'single'"/>
      </xsl:when>
      <xsl:when test="$linetype='thick'">
        <xsl:value-of select="'thick'"/>
      </xsl:when>
      <xsl:when test="$linetype='double'">
        <xsl:value-of select="'double'"/>
      </xsl:when>
      <xsl:when test="$linetype='dotted'">
        <xsl:value-of select="'dotted'"/>
      </xsl:when>
      <xsl:when test="$linetype='dashed'">
        <xsl:value-of select="'dashed'"/>
      </xsl:when>
      <xsl:when test="$linetype='dotDash'">
        <xsl:value-of select="'dot-dash'"/>
      </xsl:when>
      <xsl:when test="$linetype='dotDotDash'">
        <xsl:value-of select="'dot-dot-dash'"/>
      </xsl:when>
      <xsl:when test="$linetype='wave'">
        <xsl:value-of select="'wave'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'single'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>-->
 
  <!--zhaobj 箭头-->
  <!--<xsl:template match="wps:spPr/a:ln" mode ="Arrow">
    <xsl:for-each select="node()">
   
      <xsl:if test="name(.)='a:headEnd'">
         <xsl:choose>
             <xsl:when test ="@type='triangle'">
              <图:前端箭头 uof:locID="g0017">
                <图:式样 uof:locID="g0018">normal</图:式样>
                <图:大小 uof:locID="g0019">
                  <xsl:call-template name ="Arrowsize"/>
                </图:大小>
              </图:前端箭头>
            </xsl:when>
            <xsl:when test ="@type='arrow'">
              <图:前端箭头 uof:locID="g0017">
                <图:式样 uof:locID="g0018">open</图:式样>
                <图:大小 uof:locID="g0019">
                  <xsl:call-template name ="Arrowsize"/>
                </图:大小>
              </图:前端箭头>
            </xsl:when>
            <xsl:when test ="@type='stealth'">
              <图:前端箭头 uof:locID="g0017">
                <图:式样 uof:locID="g0018">stealth</图:式样>
                <图:大小 uof:locID="g0019">
                  <xsl:call-template name ="Arrowsize"/>
                </图:大小>
              </图:前端箭头>
            </xsl:when>
            <xsl:when test="@type='diamond'">
              <图:前端箭头 uof:locID="g0017">
                <图:式样 uof:locID="g0018">diamond</图:式样>
                <图:大小 uof:locID="g0019">
                  <xsl:call-template name ="Arrowsize"/>
                </图:大小>
              </图:前端箭头>
            </xsl:when>
            <xsl:when test ="@type='oval'">
              <图:前端箭头 uof:locID="g0017">
                <图:式样 uof:locID="g0018">oval</图:式样>
                <图:大小 uof:locID="g0019">
                  <xsl:call-template name ="Arrowsize"/>
                </图:大小> </图:前端箭头>
            </xsl:when>
            <xsl:otherwise> </xsl:otherwise>
            --><!--<xsl:when test ="@type='none' or not(@type)"></xsl:when>--><!--
         </xsl:choose>
       </xsl:if>
      <xsl:if test="name(.)='a:tailEnd'">
        <xsl:choose>
          <xsl:when test ="@type='triangle'">
            <图:后端箭头 uof:locID="g0020">
              <图:式样 uof:locID="g0021">normal</图:式样>
              <图:大小 uof:locID="g0022">
                <xsl:call-template name ="Arrowsize"/>
              </图:大小>
            </图:后端箭头>
           </xsl:when>
          <xsl:when test ="@type='arrow'">
            <图:后端箭头 uof:locID="g0020">
              <图:式样 uof:locID="g0021">open</图:式样>
              <图:大小 uof:locID="g0022">
                <xsl:call-template name ="Arrowsize"/>
              </图:大小>
            </图:后端箭头>
          </xsl:when>
          <xsl:when test ="@type='stealth'">
            <图:后端箭头 uof:locID="g0020">
              <图:式样 uof:locID="g0021">stealth</图:式样>
              <图:大小 uof:locID="g0022">
                <xsl:call-template name ="Arrowsize"/>
              </图:大小>
            </图:后端箭头>
          </xsl:when>
          <xsl:when test ="@type='diamond'">
            <图:后端箭头 uof:locID="g0020">
              <图:式样 uof:locID="g0021">diamond</图:式样>
              <图:大小 uof:locID="g0022">
                <xsl:call-template name ="Arrowsize"/>
              </图:大小>
            </图:后端箭头>
           </xsl:when>
          <xsl:when test ="@type='oval'">
            <图:后端箭头 uof:locID="g0020">
              <图:式样 uof:locID="g0021">oval</图:式样>
              <图:大小 uof:locID="g0022">
                <xsl:call-template name ="Arrowsize"/>
              </图:大小>
            </图:后端箭头>
          </xsl:when>
          <xsl:otherwise> </xsl:otherwise>
        </xsl:choose>              
      </xsl:if>
            
    </xsl:for-each>
  </xsl:template>
  <xsl:template name ="Arrowsize">
    <xsl:choose>
      <xsl:when test ="@w='sm' and @len='sm'">1</xsl:when>
      <xsl:when test ="@w='sm' and @len='med'">2</xsl:when>
      <xsl:when test ="@w='sm' and @len='lg'">3</xsl:when>
      <xsl:when test ="@w='med' and @len='sm'">4</xsl:when>
      <xsl:when test ="@w='med' and @len='med'">5</xsl:when>
      <xsl:when test ="@w='med' and @len='lg'">6</xsl:when>
      <xsl:when test ="@w='lg' and @len='sm'">7</xsl:when>
      <xsl:when test ="@w='lg' and @len='med'">8</xsl:when>
      <xsl:when test ="@w='lg' and @len='lg'">9</xsl:when>
      <xsl:otherwise>5</xsl:otherwise>
    </xsl:choose>
   
  </xsl:template>-->

  <!--zhaobj 阴影 只转换简单的情况--><!--
  <xsl:template match ="wps:spPr/a:effectLst/a:outerShdw" mode ="shadow">
    <图:阴影 uof:locID="g0048" uof:attrList="是否显示阴影 阴影类型 X偏移量 Y偏移量 阴影颜色 阴影透明度" uof:是否显示阴影="true">
      <xsl:choose>
        <xsl:when test ="@algn='br' and @kx">
          <xsl:attribute name ="uof:阴影类型">3</xsl:attribute>
        </xsl:when>
        <xsl:when test ="@algn='bl' and @kx">
          <xsl:attribute name ="uof:阴影类型">2</xsl:attribute>
        </xsl:when>
        
        <xsl:when test="@algn='tl'">
          <xsl:attribute name="uof:阴影类型">5</xsl:attribute>
        </xsl:when>
        <xsl:when test="@algn='t'">
          <xsl:attribute name="uof:阴影类型">18</xsl:attribute>
        </xsl:when>
        <xsl:when test="@algn='tr'">
          <xsl:attribute name="uof:阴影类型">4</xsl:attribute>
        </xsl:when>
        <xsl:when test="@algn='bl'">
          <xsl:attribute name="uof:阴影类型">1</xsl:attribute>
        </xsl:when>
        <xsl:when test ="@algn='b'">
          <xsl:attribute name ="uof:阴影类型">1</xsl:attribute>
        </xsl:when>
        <xsl:when test ="@algn='l'">
          <xsl:attribute name ="uof:阴影类型">16</xsl:attribute>
        </xsl:when>
        <xsl:when test ="@algn='r'">
          <xsl:attribute name ="uof:阴影类型">17</xsl:attribute>
        </xsl:when>
        <xsl:when test ="@algn='br'">
          <xsl:attribute name ="uof:阴影类型">0</xsl:attribute>
        </xsl:when>
        <xsl:when test ="@algn='ctr'">
          <xsl:attribute name ="uof:阴影类型">18</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="uof:阴影类型">1</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:if test="x">
        <xsl:attribute name="uof:X偏移量">
          <xsl:value-of select="round(x)"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="y">
        <xsl:attribute name="uof:Y偏移量">
          <xsl:value-of select="round(y)"/>
        </xsl:attribute>
      </xsl:if>
     --><!-- <xsl:if test="@sx">
        <xsl:attribute name="uof:X偏移量">
           <xsl:value-of select="@sx div 12700"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test ="@kx">
        <xsl:attribute name ="uof:X偏移量">
          <xsl:value-of select ="'0'"/>
        </xsl:attribute>
      </xsl:if>
      
       
      <xsl:attribute name="uof:X偏移量">
        <xsl:value-of select="math:cos($dir) * $dist"/>
      </xsl:attribute>
      <xsl:if test="@sy">
        <xsl:attribute name="uof:Y偏移量">
         <xsl:if test ="@kx">
            <xsl:value-of select ="'0'"/>
          </xsl:if>
          <xsl:if test ="not (@kx)">
            <xsl:value-of select="@sy div 12700"/>
          </xsl:if>
        </xsl:attribute>
      </xsl:if> --><!--
      <xsl:choose>
        <xsl:when test="a:srgbClr">
          <xsl:attribute name="uof:阴影颜色">
            <xsl:value-of select="concat('#',a:srgbClr/@val)"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="a:prstClr">
          <xsl:attribute name="uof:阴影颜色">
            <xsl:apply-templates select="a:prstClr"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="a:schemeClr">
          <xsl:attribute name="uof:阴影颜色">
            <xsl:apply-templates select="a:schemeClr"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="uof:阴影颜色">
            <xsl:value-of select="'auto'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:if test="//a:alpha">
        <xsl:attribute name="uof:阴影透明度">
         --><!--<xsl:value-of select="round((//a:alpha/@val div (1000 * 100))*256)"/>--><!--
          <xsl:value-of select ="//a:alpha/@val div 1000"/>
        </xsl:attribute>
      </xsl:if>
      
      
      
      
    </图:阴影>
    
  </xsl:template>-->
  <!--图填充-->
  <!--<xsl:template match ="a:solidFill">
    <xsl:call-template name="colorChoice"/>
  </xsl:template>
  <xsl:template match="a:gradFill">
    <图:渐变 uof:locID="g0037" uof:attrList="起始色 终止色 种子类型 起始浓度 终止浓度 渐变方向 边界 种子X位置 种子Y位置 类型">
      <xsl:variable name="angle">
        <xsl:choose>
          <xsl:when test="a:lin">
            <xsl:value-of select="(360 - round(a:lin/@ang div 60000 div 45)*45+90) mod 360"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:attribute name="图:种子类型">
        <xsl:choose>
          <xsl:when test="a:path">
            <xsl:choose>
              <xsl:when test="a:path/@path='rect'">square</xsl:when>
              <xsl:when test="a:path/@path='circle'">oval</xsl:when>      
              <xsl:when test="a:path/@path='shape'">square</xsl:when>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>linear</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="图:起始浓度">1.0</xsl:attribute>
      <xsl:attribute name="图:终止浓度">1.0</xsl:attribute>
      <xsl:attribute name="图:渐变方向">
        <xsl:value-of select="$angle"/>
      </xsl:attribute>
      <xsl:choose>
        <xsl:when test="$angle='135' or $angle='180' or $angle='225' or $angle='270'">
          <xsl:for-each select="a:gsLst/a:gs">
            <xsl:sort select="@pos" data-type="number"/>
            <xsl:if test="position()=1">
              <xsl:attribute name="图:终止色">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="position()=last()">
              <xsl:attribute name="图:起始色">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="a:gsLst/a:gs">
            <xsl:sort select="@pos" data-type="number"/>
            <xsl:if test="position()=1">
              <xsl:attribute name="图:起始色">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="position()=last()">
              <xsl:attribute name="图:终止色">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:attribute name="图:边界">50</xsl:attribute>

      <xsl:choose>
        <xsl:when test="not(a:path)">
          <xsl:attribute name="图:种子X位置">100</xsl:attribute>
          <xsl:attribute name="图:种子Y位置">100</xsl:attribute>
        </xsl:when>
        <xsl:when test="a:path/@path='rect'">
          <xsl:choose>
            <xsl:when test="not(a:path/a:fillToRect/@l) and not(a:path/a:fillToRect/@t)">
              <xsl:attribute name="图:种子X位置">30</xsl:attribute>
              <xsl:attribute name="图:种子Y位置">30</xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:path/a:fillToRect/@l) and not(a:path/a:fillToRect/@b)">
              <xsl:attribute name="图:种子X位置">30</xsl:attribute>
              <xsl:attribute name="图:种子Y位置">60</xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:path/a:fillToRect/@r) and not(a:path/a:fillToRect/@t)">
              <xsl:attribute name="图:种子X位置">60</xsl:attribute>
              <xsl:attribute name="图:种子Y位置">30</xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:path/a:fillToRect/@r) and not(a:path/a:fillToRect/@b)">
              <xsl:attribute name="图:种子X位置">60</xsl:attribute>
              <xsl:attribute name="图:种子Y位置">60</xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
      </xsl:choose>

    </图:渐变>
  </xsl:template>-->

  <!--<xsl:template match="a:fillRef">
    <xsl:call-template name="colorChoice"/>
  </xsl:template>

  <xsl:template match="a:pattFill">
    <图:图案 uof:locID="g0036" uof:attrList="类型 图形引用 前景色 背景色">
      <xsl:attribute name="图:类型">
        <xsl:call-template name="patternFill">
          <xsl:with-param name="pa" select="1"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="图:前景色">
        <xsl:call-template name="patternFill">
          <xsl:with-param name="pa" select="2"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="图:背景色">
        <xsl:call-template name="patternFill">
          <xsl:with-param name="pa" select="3"/>
        </xsl:call-template>
      </xsl:attribute>
    </图:图案>
  </xsl:template>-->

  <!--<xsl:template name="patternFill">
    <xsl:param name="pa"/>
    <xsl:choose>
      <xsl:when test="$pa = '1'">
        <xsl:choose>
          <xsl:when test="@prst ='pct5'">ptn001</xsl:when>
          <xsl:when test="@prst ='pct10'">ptn002</xsl:when>
          <xsl:when test="@prst ='pct20'">ptn003</xsl:when>
          <xsl:when test="@prst ='pct25'">ptn004</xsl:when>
          <xsl:when test="@prst ='pct30'">ptn005</xsl:when>
          <xsl:when test="@prst ='pct40'">ptn006</xsl:when>
          <xsl:when test="@prst ='pct50'">ptn007</xsl:when>
          <xsl:when test="@prst ='pct60'">ptn008</xsl:when>
          <xsl:when test="@prst ='pct70'">ptn009</xsl:when>
          <xsl:when test="@prst ='pct75'">ptn010</xsl:when>
          <xsl:when test="@prst ='pct80'">ptn011</xsl:when>
          <xsl:when test="@prst ='pct90'">ptn012</xsl:when>
          <xsl:when test="@prst ='ltDnDiag'">ptn013</xsl:when>
          <xsl:when test="@prst ='ltUpDiag'">ptn014</xsl:when>
          <xsl:when test="@prst ='dkDnDiag'">ptn015</xsl:when>
          <xsl:when test="@prst ='dkUpDiag'">ptn016</xsl:when>
          <xsl:when test="@prst ='wdDnDiag'">ptn017</xsl:when>
          <xsl:when test="@prst ='wdUpDiag'">ptn018</xsl:when>
          <xsl:when test="@prst ='ltVert'">ptn019</xsl:when>
          <xsl:when test="@prst ='ltHorz'">ptn020</xsl:when>
          <xsl:when test="@prst ='narVert'">ptn021</xsl:when>
          <xsl:when test="@prst ='narHorz'">ptn022</xsl:when>
          <xsl:when test="@prst ='dkVert'">ptn023</xsl:when>
          <xsl:when test="@prst ='dkHorz'">ptn024</xsl:when>
          <xsl:when test="@prst ='dashDnDiag'">ptn025</xsl:when>
          <xsl:when test="@prst ='dashUpDiag'">ptn026</xsl:when>
          <xsl:when test="@prst ='dashHorz'">ptn027</xsl:when>
          <xsl:when test="@prst ='dashVert'">ptn028</xsl:when>
          <xsl:when test="@prst ='smConfetti'">ptn029</xsl:when>
          <xsl:when test="@prst ='lgConfetti'">ptn030</xsl:when>
          <xsl:when test="@prst ='zigZag'">ptn031</xsl:when>
          <xsl:when test="@prst ='wave'">ptn032</xsl:when>
          <xsl:when test="@prst ='diagBrick'">ptn033</xsl:when>
          <xsl:when test="@prst ='horzBrick'">ptn034</xsl:when>
          <xsl:when test="@prst ='weave'">ptn035</xsl:when>
          <xsl:when test="@prst ='plaid'">ptn036</xsl:when>
          <xsl:when test="@prst ='divot'">ptn037</xsl:when>
          <xsl:when test="@prst ='dotGrid'">ptn038</xsl:when>
          <xsl:when test="@prst ='dotDmnd'">ptn039</xsl:when>
          <xsl:when test="@prst ='shingle'">ptn040</xsl:when>
          <xsl:when test="@prst ='trellis'">ptn041</xsl:when>
          <xsl:when test="@prst ='sphere'">ptn042</xsl:when>
          <xsl:when test="@prst ='smGrid'">ptn043</xsl:when>
          <xsl:when test="@prst ='lgGrid'">ptn044</xsl:when>
          <xsl:when test="@prst ='smCheck'">ptn045</xsl:when>
          <xsl:when test="@prst ='lgCheck'">ptn046</xsl:when>
          <xsl:when test="@prst ='openDmnd'">ptn047</xsl:when>
          <xsl:when test="@prst ='solidDmnd'">ptn048</xsl:when>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$pa = '2'">
        <xsl:apply-templates select="a:fgClr"/>
      </xsl:when>
      <xsl:when test="$pa = '3'">
        <xsl:apply-templates select="a:bgClr"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>-->
  <!--fgClr前景色-->
  <!--<xsl:template match="a:fgClr">
    <xsl:call-template name="colorChoice"/>
  </xsl:template>
  --><!--bgClr背景色--><!--
  <xsl:template match="a:bgClr">
    <xsl:call-template name="colorChoice"/>
  </xsl:template>

  <xsl:template match="a:blipFill">
    <xsl:param name ="picturefrom"/>
    <xsl:variable name ="number">
      <xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline"/>
    </xsl:variable>
    <图:图片 uof:locID="g0035" uof:attrList="位置 图形引用 类型 名称">
      <xsl:attribute name="图:位置">
        <xsl:choose>
          <xsl:when test="a:stretch">
            <xsl:value-of select="'stretch'"/>
          </xsl:when>
          <xsl:when test="a:tile">
            <xsl:value-of select="'tile'"/>
          </xsl:when>
          --><!--2010.04.12 myx add --><!--
          <xsl:when test="a:srcRect">tile</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'center'"/>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:attribute>
      <xsl:attribute name="图:图形引用">
        --><!-- <xsl:call-template name="findPicName"/>--><!--
        <xsl:if test ="$picturefrom='grpsp'">
          <xsl:if test ="../../wps:cNvPr/@id">
            <xsl:value-of select ="concat('grpsppicOBJ',../../wps:cNvPr/@id)"/>
          </xsl:if>
          <xsl:if test ="../../wpg:cNvPr/@id">
            <xsl:value-of select ="concat('grpsppicOBJ',../../wpg:cNvPr/@id)"/>
          </xsl:if>
        </xsl:if>
        <xsl:if test ="$picturefrom!='grpsp'">
          <xsl:value-of select ="concat($picturefrom,'OBJ',$number * 2)"/>
        </xsl:if>
      </xsl:attribute>
    </图:图片>
  </xsl:template>

  <xsl:template match ="a:grpFill">
    <xsl:choose>
      <xsl:when test ="../../../wpg:grpSpPr/a:solidFill">
        <xsl:apply-templates select ="../../../wpg:grpSpPr/a:solidFill"/>
      </xsl:when>
      <xsl:when test ="../../../wpg:grpSpPr/a:patternFill">
        <xsl:apply-templates select ="../../../wpg:grpSpPr/a:patternFill"/>
      </xsl:when>
      <xsl:when test ="../../../wpg:grpSpPr/a:pattFill">
        <xsl:apply-templates select ="../../../wpg:grpSpPr/a:pattFill"/>
      </xsl:when>
      <xsl:when test ="../../../wpg:grpSpPr/a:blipFill">
        <xsl:apply-templates select ="../../../wpg:grpSpPr/a:blipFill">
          <xsl:with-param name ="picturefrom" select ="grpsp"/>
        </xsl:apply-templates>
      </xsl:when>
      <xsl:when test ="../../../wpg:grpSpPr/a:grpFill">
        <xsl:apply-templates select ="../../../wpg:grpSpPr/a:grpFill"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template name="findPicName">
    <xsl:variable name="findPictureName">
      <xsl:value-of select=".//a:blip/@r:embed"/>
    </xsl:variable>
    <xsl:call-template name="findRelatedRelationships">
      <xsl:with-param name="id">
        <xsl:value-of select="$findPictureName"/>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="findRelatedRelationships">
    <xsl:param name="id"/>
   --><!--<xsl:for-each select="ancestor::p:cSld | ancestor::p:txStyles">--><!-- 
      <xsl:apply-templates select="document('word/_rels/document.xml.rels')/rel:Relationships">
        <xsl:with-param name="targetID">
          <xsl:value-of select="$id"/>
        </xsl:with-param>
      </xsl:apply-templates>
  --><!--  </xsl:for-each>--><!--
  </xsl:template>

  <xsl:template match="rel:Relationships">
    <xsl:param name="targetID"/>
    <xsl:for-each select="rel:Relationship">
      <xsl:if test="@Id = $targetID">
        <xsl:value-of select="substring-before(substring-after(@Target,'media/'),'.')"/>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>-->
  
  <!--zhaobj 线颜色--><!--
  <xsl:template match ="a:ln" mode="linecolor">
    <xsl:choose>
      <xsl:when test ="a:solidFill">
        <图:线颜色 uof:locID="g0013">
          <xsl:choose>
            <xsl:when test ="a:solidFill/a:srgbClr">
              <xsl:value-of select="concat('#',a:solidFill/a:srgbClr/@val)"/>
            </xsl:when>
            <xsl:when test ="a:solidFill/a:schemeClr">
              <xsl:apply-templates select="a:solidFill" mode="lnschemeClr"/>
            </xsl:when>
            <xsl:otherwise>
              --><!--转换成默认的accent1--><!--
                <xsl:value-of select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val"/>           
            </xsl:otherwise>
          </xsl:choose>       
        </图:线颜色>
      </xsl:when>
      --><!--zhaobj 2/24  a:gradFil 是渐进色，UOF中没有对应，转换为第一个颜色--><!--
      <xsl:when test ="a:gradFill">
        <图:线颜色 uof:locID="g0013">
          <xsl:choose>
            <xsl:when test ="a:gradFill/a:gsLst">
              <xsl:apply-templates select ="a:gradFill/a:gsLst" mode ="firstcolor"/> 
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val)"/>
            </xsl:otherwise>
          </xsl:choose>
        </图:线颜色>
      </xsl:when>
      <xsl:when test ="a:noFill">
        <图:线颜色 uof:locID="g0013">#FFFFFF</图:线颜色>
      </xsl:when>
      --><!--uof中没有图案填充的情况，直接转换为默认--><!--
      <xsl:when test ="a:pattFill">
        <图:线颜色 uof:locID="g0013">
          <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val)"/>
        </图:线颜色>
      </xsl:when>
      <xsl:otherwise>
        --><!--zhaobj 转换到lnRef中找，没有就到主题theme中  <图:线颜色 uof:locID="g0013">#000000</图:线颜色>--><!--
        <xsl:choose>
          <xsl:when test="parent::wps:spPr/parent::wps:wsp/wps:style/a:lnRef/*">
            <xsl:for-each select="parent::wps:spPr/parent::wps:wsp/wps:style/a:lnRef">
              <图:线颜色 uof:locID="g0013">
                <xsl:call-template name="lnRefcolor"/>
              </图:线颜色>
            </xsl:for-each>
          </xsl:when>
          <xsl:when test="parent::wps:spPr/parent::wps:wsp/wps:style/a:lnRef/@idx">
          <xsl:variable name="idx" select ="parent::wps:spPr/parent::wps:wsp/wps:style/a:lnRef/@idx"/>
          <xsl:apply-templates select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position()=$idx]"
                                mode="linethemecolor"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val)"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>      
  </xsl:template>-->

  <!--<xsl:template match ="a:solidFill" mode ="lnschemeClr">
    <xsl:call-template name ="colorChoice"/> 
  </xsl:template>
  
  <xsl:template match ="a:gradFill/a:gsLst" mode ="firstcolor">
   <xsl:for-each select ="node()">
      <xsl:if test ="@pos='0'">
        <xsl:value-of select="concat('#',a:srgbClr/@val)"/>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>-->
   
  <!--lnref中的颜色转换-->
  <!--<xsl:template name ="lnRefcolor">
    <xsl:choose>
      <xsl:when test ="a:schemeClr">
        <xsl:call-template name ="colorChoice"/>
      </xsl:when>
      <xsl:otherwise >
       <xsl:if test="parent::wps:spPr/parent::wps:wsp/wps:style/a:lnRef/@idx">
         <xsl:variable name="idx" select ="parent::wps:spPr/parent::wps:wsp/wps:style/a:lnRef/@idx"/>
         <xsl:apply-templates select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position()=$idx]"
                          mode="linethemecolor"/>
       </xsl:if>
     </xsl:otherwise>
    </xsl:choose>
  </xsl:template>-->
  <!--<xsl:template match ="a:ln" mode ="linethemecolor">
    <xsl:for-each  select ="a:solidFill or a:noFill or a:gradFill or a:blipFill or a:pattFill">
      <图:线颜色 uof:locID="g0013">
       <xsl:call-template name="colorChoice"/>
      </图:线颜色>
    </xsl:for-each>
   </xsl:template>
  <xsl:template name ="colorChoice">
    <xsl:choose>
      <xsl:when test="a:srgbClr">
        <xsl:value-of select="concat('#',a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="a:hslClr">
        <xsl:apply-templates select="a:hslClr"/>
      </xsl:when>
      <xsl:when test="a:prstClr">
        <xsl:apply-templates select="a:prstClr"/>
      </xsl:when>
      <xsl:when test="a:schemeClr">
        <xsl:apply-templates select="a:schemeClr"/>
      </xsl:when>
      <xsl:when test="a:scrgbClr">
        <xsl:apply-templates select="a:scrgbClr"/>
      </xsl:when>
      <xsl:when test="a:sysClr">
        <xsl:apply-templates select="a:sysClr"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>-->
  
  <!--<xsl:template match="a:prstClr">
    <xsl:choose>
      <xsl:when test="@val='aliceBlue'">#f0f8ff</xsl:when>
      <xsl:when test="@val='antiqueWhite'">#faebd7</xsl:when>
      <xsl:when test="@val='aqua'">#00ffff</xsl:when>
      <xsl:when test="@val='aquamarine'">#7fffd4</xsl:when>
      <xsl:when test="@val='azure'">#f0ffff</xsl:when>
      <xsl:when test="@val='beige'">#f5f5dc</xsl:when>
      <xsl:when test="@val='bisque'">#ffe4c4</xsl:when>
      <xsl:when test="@val='black'">#000000</xsl:when>
      <xsl:when test="@val='blanchedAlmond'">#ffebcd</xsl:when>
      <xsl:when test="@val='blue'">#0000ff</xsl:when>
      <xsl:when test="@val='blueViolet'">#8a2be2</xsl:when>
      <xsl:when test="@val='brown'">#a52a2a</xsl:when>
      <xsl:when test="@val='burlyWood'">#deb887</xsl:when>
      <xsl:when test="@val='cadetBlue'">#5f9ea0</xsl:when>
      <xsl:when test="@val='chartreuse'">#7fff00</xsl:when>
      <xsl:when test="@val='chocolate'">#d2691e</xsl:when>
      <xsl:when test="@val='coral'">#ff7f50</xsl:when>
      <xsl:when test="@val='cornflowerBlue'">#6495ed</xsl:when>
      <xsl:when test="@val='cornsilk'">#fff8dc</xsl:when>
      <xsl:when test="@val='crimson'">#dc143c</xsl:when>
      <xsl:when test="@val='cyan'">#00ffff</xsl:when>
      <xsl:when test="@val='deepPink'">#ff1493</xsl:when>
      <xsl:when test="@val='deepSkyBlue'">#00bfff</xsl:when>
      <xsl:when test="@val='dimGray'">#696969</xsl:when>
      <xsl:when test="@val='dkBlue'">#00008b</xsl:when>
      <xsl:when test="@val='dkCyan'">#008b8b</xsl:when>
      <xsl:when test="@val='dkGoldenrod'">#b8860b</xsl:when>
      <xsl:when test="@val='dkGray'">#a9a9a9</xsl:when>
      <xsl:when test="@val='dkGreen'">#006400</xsl:when>
      <xsl:when test="@val='dkKhaki'">#bdb767</xsl:when>
      <xsl:when test="@val='dkMagenta'">#8b008b</xsl:when>
      <xsl:when test="@val='dkOliveGreen'">#556b2f</xsl:when>
      <xsl:when test="@val='dkOrange'">#ff8c00</xsl:when>
      <xsl:when test="@val='dkOrchild'">#9932cc</xsl:when>
      <xsl:when test="@val='dkRed'">#8b0000</xsl:when>
      <xsl:when test="@val='dkSalmon'">#e9967a</xsl:when>
      <xsl:when test="@val='dkSeaGreen'">#8fbc8b</xsl:when>
      <xsl:when test="@val='dkSlateBlue'">#483d8b</xsl:when>
      <xsl:when test="@val='dkSlateGray'">#2f4f4f</xsl:when>
      <xsl:when test="@val='dkTurquoise'">#00cece</xsl:when>
      <xsl:when test="@val='dkViolet'">#9400d3</xsl:when>
      <xsl:when test="@val='dodgerBlue'">#1e90ff</xsl:when>
      <xsl:when test="@val='firebrick'">#b22222</xsl:when>
      <xsl:when test="@val='floralWhite'">#fffaf0</xsl:when>
      <xsl:when test="@val='forestGreen'">#228b22</xsl:when>
      <xsl:when test="@val='fuchsia'">#ff00ff</xsl:when>
      <xsl:when test="@val='gainsboro'">#dcdcdc</xsl:when>
      <xsl:when test="@val='ghostWhite'">#f8f8ff</xsl:when>
      <xsl:when test="@val='gold'">#ffd700</xsl:when>
      <xsl:when test="@val='goldenrod'">#daa520</xsl:when>
      <xsl:when test="@val='gray'">#808080</xsl:when>
      <xsl:when test="@val='green'">#008000</xsl:when>
      <xsl:when test="@val='greenYellow'">#adff2f</xsl:when>
      <xsl:when test="@val='honeydew'">#f0fff0</xsl:when>
      <xsl:when test="@val='hotPink'">#ff69b4</xsl:when>
      <xsl:when test="@val='indianRed'">#cd5c5c</xsl:when>
      <xsl:when test="@val='indigo'">#4b0082</xsl:when>
      <xsl:when test="@val='ivory'">#fffff0</xsl:when>
      <xsl:when test="@val='khaki'">#f0e68c</xsl:when>
      <xsl:when test="@val='lavender'">#e6e6fa</xsl:when>
      <xsl:when test="@val='lavenderBlush'">#fff0f5</xsl:when>
      <xsl:when test="@val='lawnGreen'">#7cfc00</xsl:when>
      <xsl:when test="@val='lemonChiffon'">#fffacd</xsl:when>
      <xsl:when test="@val='lime'">#00ff00</xsl:when>
      <xsl:when test="@val='limeGreen'">#32cd32</xsl:when>
      <xsl:when test="@val='linen'">#faf0e6</xsl:when>
      <xsl:when test="@val='ltBlue'">#add8e6</xsl:when>
      <xsl:when test="@val='ltCoral'">#f08080</xsl:when>
      <xsl:when test="@val='ltCyan'">#e0ffff</xsl:when>
      <xsl:when test="@val='ltGoldenrodYellow'">#fafa78</xsl:when>
      <xsl:when test="@val='ltGray'">#d3d3d3</xsl:when>
      <xsl:when test="@val='ltGreen'">#90ee90</xsl:when>
      <xsl:when test="@val='ltPink'">#ffb6c1</xsl:when>
      <xsl:when test="@val='ltSalmon'">#ffa07a</xsl:when>
      <xsl:when test="@val='ltSeaGreen'">#20b2aa</xsl:when>
      <xsl:when test="@val='ltSkyBlue'">#87cefa</xsl:when>
      <xsl:when test="@val='ltSlateGray'">#778899</xsl:when>
      <xsl:when test="@val='ltSteelBlue'">#b0c4de</xsl:when>
      <xsl:when test="@val='ltYellow'">#ffffe0</xsl:when>
      <xsl:when test="@val='magenta'">#ff00ff</xsl:when>
      <xsl:when test="@val='maroon'">#800000</xsl:when>
      <xsl:when test="@val='medAquamarine'">#66cdaa</xsl:when>
      <xsl:when test="@val='medBlue'">#0000cd</xsl:when>
      <xsl:when test="@val='medOrchid'">#ba55d3</xsl:when>
      <xsl:when test="@val='medPurple'">#9370db</xsl:when>
      <xsl:when test="@val='medSeaGreen'">#3cb371</xsl:when>
      <xsl:when test="@val='medSlateBlue'">#7b68ee</xsl:when>
      <xsl:when test="@val='medSpringGreen'">#00fa9a</xsl:when>
      <xsl:when test="@val='medTurquoise'">#48d1cc</xsl:when>
      <xsl:when test="@val='medVioletRed'">#c71585</xsl:when>
      <xsl:when test="@val='midnightBlue'">#191970</xsl:when>
      <xsl:when test="@val='mintCream'">#f5fffa</xsl:when>
      <xsl:when test="@val='mistyRose'">#ffe4e1</xsl:when>
      <xsl:when test="@val='moccasin'">#ffe4b5</xsl:when>
      <xsl:when test="@val='navajoWhite'">#ffdead</xsl:when>
      <xsl:when test="@val='navy'">#000080</xsl:when>
      <xsl:when test="@val='oldLace'">#fdf5e6</xsl:when>
      <xsl:when test="@val='olive'">#808000</xsl:when>
      <xsl:when test="@val='oliveDrab'">#6b8e23</xsl:when>
      <xsl:when test="@val='orange'">#ffa500</xsl:when>
      <xsl:when test="@val='orangeRed'">#ff4500</xsl:when>
      <xsl:when test="@val='orchid'">#da70d6</xsl:when>
      <xsl:when test="@val='paleGoldenrod'">#eee8aa</xsl:when>
      <xsl:when test="@val='paleGreen'">#98fb98</xsl:when>
      <xsl:when test="@val='paleTurquoise'">#afeeee</xsl:when>
      <xsl:when test="@val='paleVioletRed'">#db7093</xsl:when>
      <xsl:when test="@val='papayaWhip'">#ffefd5</xsl:when>
      <xsl:when test="@val='peachPuff'">#ffdab9</xsl:when>
      <xsl:when test="@val='peru'">#cd853f</xsl:when>
      <xsl:when test="@val='pink'">#ffc0cb</xsl:when>
      <xsl:when test="@val='plum'">#dda0dd</xsl:when>
      <xsl:when test="@val='powderBlue'">#b0e0e6</xsl:when>
      <xsl:when test="@val='purple'">#800080</xsl:when>
      <xsl:when test="@val='red'">#ff0000</xsl:when>
      <xsl:when test="@val='rosyBrown'">#bc8f8f</xsl:when>
      <xsl:when test="@val='royalBlue'">#4169e1</xsl:when>
      <xsl:when test="@val='saddleBrown'">#8b4513</xsl:when>
      <xsl:when test="@val='salmon'">#fa8072</xsl:when>
      <xsl:when test="@val='sandyBrown'">#f4a460</xsl:when>
      <xsl:when test="@val='seaGreen'">#2e8b57</xsl:when>
      <xsl:when test="@val='seaShell'">#fff5ee</xsl:when>
      <xsl:when test="@val='sienna'">#a0522d</xsl:when>
      <xsl:when test="@val='silver'">#c0c0c0</xsl:when>
      <xsl:when test="@val='skyBlue'">#87ceeb</xsl:when>
      <xsl:when test="@val='slateBlue'">#6a5acd</xsl:when>
      <xsl:when test="@val='slateGray'">#708090</xsl:when>
      <xsl:when test="@val='snow'">#fffafa</xsl:when>
      <xsl:when test="@val='springGreen'">#00ff7f</xsl:when>
      <xsl:when test="@val='steelBlue'">#4682b4</xsl:when>
      <xsl:when test="@val='tan'">#d2b48c</xsl:when>
      <xsl:when test="@val='teal'">#008080</xsl:when>
      <xsl:when test="@val='thistle'">#d8bfd8</xsl:when>
      <xsl:when test="@val='tomato'">#ff6347</xsl:when>
      <xsl:when test="@val='turquoise'">#40e0d0</xsl:when>
      <xsl:when test="@val='violet'">#ee82ee</xsl:when>
      <xsl:when test="@val='wheat'">#f5deb3</xsl:when>
      <xsl:when test="@val='white'">#ffffff</xsl:when>
      <xsl:when test="@val='whiteSmoke'">#f5f5f5</xsl:when>
      <xsl:when test="@val='yellow'">#ffff00</xsl:when>
      <xsl:when test="@val='yellowGreen'">#9acd32</xsl:when>
    </xsl:choose>
  </xsl:template>-->
  <!-- zhaobj主题色-->
  <!--<xsl:template match="a:schemeClr">
    <xsl:variable name ="color">
    <xsl:choose>   
      <xsl:when test="@val = 'bg1'">
        <xsl:value-of select="document('word/settings.xml')/w:settings/w:clrSchemeMapping/@w:bg1"/>
      </xsl:when>
      <xsl:when test="@val = 'bg2'">
        <xsl:value-of select="document('word/settings.xml')/w:settings/w:clrSchemeMapping/@w:bg2"/>
      </xsl:when>
      <xsl:when test="@val = 'tx1'">
        <xsl:value-of select="document('word/settings.xml')/w:settings/w:clrSchemeMapping/@w:t1"/>
      </xsl:when>
      <xsl:when test="@val = 'tx2'">
        <xsl:value-of select="document('word/settings.xml')/w:settings/w:clrSchemeMapping/@w:t2"/>
      </xsl:when>
      <xsl:otherwise>
        --><!--xsl:value-of select="'#4F81BD'"/--><!--
        <xsl:value-of select ="@val"/>     
      </xsl:otherwise>
    </xsl:choose>    
    </xsl:variable>
    <xsl:call-template name ="Findcolor">
          <xsl:with-param name="colorvalue" select="$color" />
        </xsl:call-template>
  </xsl:template>
  
  <xsl:template name ="Findcolor">
    <xsl:param name ="colorvalue"/>
    <xsl:choose>
      <xsl:when test="$colorvalue = 'dk1'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:dk1/a:sysClr/@lastClr)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'dark1'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:dk1/a:sysClr/@lastClr)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'lt1'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:lt1/a:sysClr/@lastClr)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'light1'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:lt1/a:sysClr/@lastClr)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'dk2'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:dk2/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'dark2'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:dk2/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'lt2'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:lt2/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'light2'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:lt2/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'accent1'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'accent2'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent2/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'accent3'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent3/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'accent4'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent4/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'accent5'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent5/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'accent6'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'hlink'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:hlink/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:when test="$colorvalue = 'folHlink'">
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:folHlink/a:srgbClr/@val)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val)"/>
      </xsl:otherwise>
      
    </xsl:choose>
    
    
  </xsl:template>-->

  <!--<xsl:template match ="wps:txbx" mode ="文本框">
    <图:文本内容 uof:locID="g0002" uof:attrList="文本框 左边距 右边距 上边距 下边距 水平对齐 垂直对齐 文字排列方向 自动换行 大小适应文字 前一链接 后一链接" >
      --><!--<xsl:attribute name ="文本框">
        <xsl:value-of select ="'true'"/>
      </xsl:attribute>--><!--
      <xsl:attribute name ="图:左边距">
        <xsl:value-of select ="following-sibling::wps:bodyPr/@lIns div 9525"/>
      </xsl:attribute>
      <xsl:attribute name ="图:右边距">
        <xsl:value-of select ="following-sibling::wps:bodyPr/@rIns div 9525"/>
      </xsl:attribute>
      <xsl:attribute name ="图:上边距">
        <xsl:value-of select ="following-sibling::wps:bodyPr/@tIns div 9525"/>
      </xsl:attribute>
      <xsl:attribute name ="图:下边距">
        <xsl:value-of select ="following-sibling::wps:bodyPr/@bIns div 9525"/>
      </xsl:attribute>
      <xsl:attribute name ="图:水平对齐">
        <xsl:value-of select ="following-sibling::wps:bodyPr/@lIns div 9525"/>
      </xsl:attribute>
      <xsl:attribute name ="图:水平对齐">
        <xsl:if test ="following-sibling::wps:bodyPr/@vert='eaVert'">
          <xsl:value-of select ="'left'"/>         
        </xsl:if>
        <xsl:if test ="following-sibling::wps:bodyPr/@vert='horn'">
          <xsl:value-of select ="'right'"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name ="图:垂直对齐">
        <xsl:choose>
          <xsl:when test ="following-sibling::wps:bodyPr/@anchor='t'">
            <xsl:value-of select ="'top'"/>
          </xsl:when>
          <xsl:when test ="following-sibling::wps:bodyPr/@anchor='ctr'">
            <xsl:value-of select ="'middle'"/>
          </xsl:when>
          <xsl:when test ="following-sibling::wps:bodyPr/@anchor='b'">
            <xsl:value-of select ="'bottom'"/>
          </xsl:when>
        </xsl:choose>        
      </xsl:attribute>
      <xsl:attribute name="图:文字排列方向">
        <xsl:if test ="following-sibling::wps:bodyPr/@vert='horz'">
          <xsl:choose>
            <xsl:when test ="parent::wps:wsp/@normalEastAsianFlow='1'">
               <xsl:value-of select ="'t2b-r2l-0e-0w'"/>
            </xsl:when>
            <xsl:otherwise>
               <xsl:value-of select ="'t2b-l2r-0e-0w'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <xsl:if test ="following-sibling::wps:bodyPr/@vert='eaVert'">
          <xsl:value-of select ="'r2l-t2b-0e-90w'"/>
        </xsl:if>
        <xsl:if test ="following-sibling::wps:bodyPr/@vert='vert'">
          <xsl:value-of select="'r2l-t2b-90e-90w'"/>
        </xsl:if>
        <xsl:if test ="following-sibling::wps:bodyPr/@vert='vert270'">
          <xsl:value-of select ="'l2r-b2t-270e-270w'"/>
        </xsl:if>
        --><!--第五种方向--><!--
      </xsl:attribute>
      <xsl:attribute name ="图:自动换行">
        <xsl:if test ="following-sibling::wps:bodyPr/@wrap='none'">
          <xsl:value-of select ="'false'"/>
        </xsl:if>
        <xsl:if test ="following-sibling::wps:bodyPr/@wrap='square'">
          <xsl:value-of select ="'true'"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name ="图:大小适应文字">
        <xsl:if test ="following-sibling::wps:bodyPr/a:spAutoFit">
          <xsl:value-of select ="true"/>
        </xsl:if>
        <xsl:if test ="following-sibling::wps:bodyPr/a:noAutofit">
          <xsl:value-of select ="false"/>
        </xsl:if>
        
      </xsl:attribute>
      <xsl:for-each select =".//w:txbxContent/w:p">
       --><!-- <xsl:call-template name ="paragraph"/> --><!--      
      </xsl:for-each>                 
    </图:文本内容>    
  </xsl:template>-->
                
                
                
               
</xsl:stylesheet>
