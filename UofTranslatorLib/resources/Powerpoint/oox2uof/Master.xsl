<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
        xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
				xmlns:fo="http://www.w3.org/1999/XSL/Format"
				xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
				xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
				xmlns:dc="http://purl.org/dc/elements/1.1/"
				xmlns:dcterms="http://purl.org/dc/terms/"
				xmlns:dcmitype="http://purl.org/dc/dcmitype/"
				xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
				xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
				xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
				xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                
				xmlns:uof="http://schemas.uof.org/cn/2009/uof"
				xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
				xmlns:演="http://schemas.uof.org/cn/2009/presentation"
				xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
				xmlns:图="http://schemas.uof.org/cn/2009/graph"
				xmlns:规则="http://schemas.uof.org/cn/2009/rules">
	<xsl:import href="ph.xsl"/>
  <xsl:import href ="fill.xsl"/>
  <!--2010-11-23 罗文甜-->
  <xsl:import href="animation.xsl"/>
	<xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="p:sldMaster" mode="ph">
		<演:母版_6C0D>
		<!--<演:母版 uof:locID="p0036" uof:attrList="标识符 名称 类型 页面设置引用 配色方案引用 页面版式引用 文本式样引用" 演:类型="slide" 演:页面设置引用="IDymsz">-->
			<xsl:variable name="masterID" select="@id"/>
			<xsl:attribute name="标识符_6BE8">
				<xsl:value-of select="substring-before(@id,'.xml')"/>
			</xsl:attribute>
			<!--<xsl:if test="p:cSld/@name">
				<xsl:attribute name="名称_6BE9">
					<xsl:value-of select="p:cSld/@name"/>
				</xsl:attribute>
			</xsl:if>-->
			<xsl:attribute name="名称_6BE9">默认设计模板</xsl:attribute>
			<xsl:attribute name="类型_6BEA">slide</xsl:attribute>
			<xsl:attribute name="文本式样引用_6BED">
				<xsl:value-of select="concat('txStyle-',substring-before(@id,'.xml'))"/>
			</xsl:attribute>
			
			<xsl:if test="p:clrMap/@*">
				<xsl:attribute name="配色方案引用_6BEB">
					<!--xsl:value-of select="substring-before(substring-after(following-sibling::rel:Relationships[@id=concat($masterID,'.rels')]/rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' or @Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride']/@Target,'theme/'),'.xml')"/-->
					<xsl:value-of select="concat('clr_',substring-before(@id,'.xml'))"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:attribute name="页面设置引用_6C18">IDymsz</xsl:attribute>
			<!--李娟 修改 页眉页脚引用 11.12.07 @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
			<xsl:if test="p:hf">
			<xsl:attribute name="页眉页脚引用_6C16">
				<xsl:value-of select="generate-id(//p:hf)"/>
			</xsl:attribute>
			</xsl:if>
			<!--<xsl:if test="p:sldLayout | //p:sld">
			<xsl:variable name="sldLayoutID" select="p:sldLayout/@id"/>
			<xsl:variable name="sldSlideID" select="substring-before(//p:sld[@sldLayoutID=$sldLayoutID]/@sldLayoutID,'.xml')"/>
			<xsl:attribute name="页面版式引用_6B27">
				<xsl:value-of select="$sldSlideID"/>
			</xsl:attribute>
			</xsl:if>-->
			<!--<xsl:apply-templates select="p:clrMap"/>-->

			<!--<xsl:apply-templates select="p:cSld/p:spTree/p:sp" mode="ph"/>--><!--····························@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
			<xsl:for-each select="p:cSld/p:spTree/p:sp|p:cSld/p:spTree/p:cxnSp|p:cSld/p:spTree/p:grpSp|p:cSld/p:spTree/p:pic|p:cSld/p:spTree/p:graphicFrame">
        <xsl:apply-templates select="." mode="ph"/>
      </xsl:for-each>
		
			
      <!--2010-11-23 罗文甜：增加动画-->    
      <xsl:apply-templates select="p:timing"/>
			<!--<演:背景_6B2C>
				<xsl:apply-templates select="../p:sldMaster/p:cSld/p:bg"/>
			</演:背景_6B2C>-->
      <xsl:call-template name="backgroundFill"/>
		<!--</演:母版>-->
		</演:母版_6C0D>
	</xsl:template>	
	
	<xsl:template match="p:notesMaster" mode="ph">
		<演:母版_6C0D>
		<!--<演:母版 uof:locID="p0036" uof:attrList="标识符 名称 类型 页面设置引用 配色方案引用 页面版式引用 文本式样引用" 演:类型="notes" 演:页面设置引用="IDymsz">-->
			<!--xsl:variable name="masterID" select="@id"/-->
			<xsl:attribute name="标识符_6BE8">
				<xsl:value-of select="substring-before(@id,'.xml')"/>
			</xsl:attribute>
			<!--add by linyh-->
			<xsl:attribute name="类型_6BEA">notes</xsl:attribute>
			<xsl:attribute name="页面设置引用_6C18">IDymsz</xsl:attribute>
			<xsl:if test="p:clrMap/@*">
				<xsl:attribute name="配色方案引用_6BEB">
					<xsl:value-of select="concat('clr_',substring-before(@id,'.xml'))"/>
				</xsl:attribute>
			</xsl:if>
			<!--<xsl:if test="p:hf">
				<xsl:attribute name="页眉页脚引用_6C16">
					<xsl:value-of select=""/>
				</xsl:attribute>
			</xsl:if>-->
      <!--<xsl:apply-templates select="p:clrMap"/>-->
      <!--xsl:call-template name="ColorSide"/-->
			
			<xsl:apply-templates select="p:cSld/p:spTree/p:sp" mode="ph"/>
			<!--add by linyh-->
			<xsl:apply-templates select=".//p:spTree/p:pic" mode="ph"/>
			<xsl:call-template name="backgroundFill"/>
		</演:母版_6C0D>
		<!--</演:母版>-->
	</xsl:template>	
		
	<xsl:template match="p:handoutMaster" mode="ph">
		<演:母版_6C0D>
		<xsl:attribute name="标识符_6BE8">
				<xsl:value-of select="substring-before(@id,'.xml')"/>
			</xsl:attribute>
			<xsl:attribute name="类型_6BEA">handout</xsl:attribute>
			<xsl:attribute name="页面设置引用_6C18">IDymsz</xsl:attribute>
			<!--add by linyh-->
			<xsl:if test="p:clrMap/@*">
				<xsl:attribute name="配色方案引用_6BEB">
					<xsl:value-of select="concat('clr_',substring-before(@id,'.xml'))"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:variable name="sldLayoutID" select="@sldLayoutID"/>
			<xsl:attribute name="页面版式引用_6BEC">
				<xsl:value-of select="substring-before($sldLayoutID,'.xml')"/>
			</xsl:attribute>
      <!--<xsl:apply-templates select="p:clrMap"/>-->
      <!--xsl:call-template name="ColorSide"/-->
			<xsl:apply-templates select="p:cSld/p:spTree/p:sp" mode="ph"/>
			<!--add by linyh-->
			<xsl:apply-templates select=".//p:spTree/p:pic" mode="ph"/>
      <!-- 09.11.12 马有旭 添加 -->
      <xsl:apply-templates select="p:cSld/p:spTree/p:cxnSp" mode="ph"/>
      <xsl:call-template name="backgroundFill"/>
			<!-- 李娟 11.12.07 添加页面版式引用-->
			
		</演:母版_6C0D>
		<!--</演:母版>-->
	</xsl:template>	
	
  <!--2010-11-15罗文甜：配色方案集-->
  <xsl:template match="p:clrMap">
	  <!--<规则:配色方案集_6C11>-->
  <!--<演:配色方案集 uof:locID="p0007">-->
    <!--p:clrMap bg1="FFFFFF" tx1="000000" bg2="EEECE1" tx2="000000" accent1="4F81BD" accent2="C0504D" accent3="9BBB59" accent4="8064A2" accent5="4BACC6" accent6="F79646" hlink="0000FF" folHlink="800080"/-->
     <!--<演:配色方案 uof:locID="p0008" uof:attrList="标识符 名称 类型" 演:类型="standard">-->
		  <规则:配色方案_6BE4>
		  <xsl:attribute name="标识符_6B0A">
          <xsl:value-of select="concat('clr_',substring-before(../@id,'.xml'))"/>
        </xsl:attribute>
			  <规则:背景色_6B02>
          <xsl:value-of select="@bg1"/>
        </规则:背景色_6B02>
			<规则:文本和线条_6B03>
          <xsl:value-of select="@tx1"/>
          </规则:文本和线条_6B03>
		 <规则:阴影_6B04>
          <xsl:value-of select="@accent1"/>
        </规则:阴影_6B04>
			  <规则:标题文本_6B05>
          <xsl:value-of select="@tx2"/>
        </规则:标题文本_6B05>
		<规则:填充_6B06>
          <xsl:value-of select="@bg2"/>
        </规则:填充_6B06>
			  <规则:强调_6B07>
          <xsl:value-of select="@accent2"/>
        </规则:强调_6B07>
			  <规则:强调和超级链接_6B08>
          <xsl:value-of select="@hlink"/>
        </规则:强调和超级链接_6B08>
			  <规则:强调和尾随超级链接_6B09>
          <xsl:value-of select="@folHlink"/>
        </规则:强调和尾随超级链接_6B09>
      <!--</演:配色方案>-->
		  
	  
	  
	  </规则:配色方案_6BE4>
  <!--</演:配色方案集>-->
	  <!--</规则:配色方案集_6C11>-->
  </xsl:template>
</xsl:stylesheet>
