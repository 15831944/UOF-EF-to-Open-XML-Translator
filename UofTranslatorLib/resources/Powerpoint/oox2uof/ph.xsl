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
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"  
				xmlns:uof="http://schemas.uof.org/cn/2009/uof"
				xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
				xmlns:演="http://schemas.uof.org/cn/2009/presentation"
				xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
        xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
				xmlns:图="http://schemas.uof.org/cn/2009/graph"
        xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
	<xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="p:sp" mode="ph">
		<!--11.11.10 李娟 锚点的属性:标识符_C62C 随动方式_C62F 是否显示缩略图_C630 有待添加-->
	
		<uof:锚点_C644>
			<xsl:attribute name="是否显示缩略图_C630">
				<xsl:choose>
					<xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='sldImg'">
						<xsl:value-of select="'true'"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of  select="'false'"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
			<xsl:choose>
				<xsl:when test="p:spPr/a:xfrm">
					<xsl:attribute name="图形引用_C62E">
						<xsl:value-of select="translate(@id,':','')"/>
					</xsl:attribute>
					<xsl:apply-templates select="p:spPr/a:xfrm" mode="ph"/>
					<xsl:apply-templates select="p:nvSpPr/p:nvPr/p:ph"/>
				</xsl:when>
				<xsl:when test="name(p:spPr/a:xfrm)=''">
					<xsl:call-template name="rel"/>
					<xsl:attribute name="图形引用_C62E">
						<xsl:value-of select="concat('obj_',generate-id(.))"/>
					</xsl:attribute>
					<xsl:apply-templates select="p:nvSpPr/p:nvPr/p:ph"/>
				</xsl:when>
			</xsl:choose>
			<!--<xsl:choose>
				<xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='sldImg'">
					<xsl:attribute name="是否显示缩略图_C630">true</xsl:attribute>
				</xsl:when>
			</xsl:choose>-->
		</uof:锚点_C644>
	</xsl:template>
	<xsl:template match="p:cxnSp" mode="ph">
		<uof:锚点_C644>
			<!--<xsl:apply-templates select="p:spPr/a:xfrm" mode="ph"/>-->
			<xsl:choose>
				<xsl:when test="p:spPr/a:xfrm">
					<xsl:attribute name="图形引用_C62E">
						<xsl:value-of select="translate(@id,':','')"/>
					</xsl:attribute>
					<xsl:apply-templates select="p:spPr/a:xfrm" mode="ph"/>
					<xsl:apply-templates select="p:nvSpPr/p:nvPr/p:ph"/>
				</xsl:when>
				<xsl:when test="name(p:spPr/a:xfrm)=''">

					<xsl:call-template name="rel"/>
					<xsl:attribute name="图形引用_C62E">
						<xsl:value-of select="concat('obj_',generate-id(.))"/>
					</xsl:attribute>

					<!--重写占位符节点-->
					<xsl:apply-templates select="p:nvSpPr/p:nvPr/p:ph"/>


				</xsl:when>
			</xsl:choose>
			
			<!--<xsl:apply-templates select="p:nvSpPr/p:nvPr/p:ph"/-->>
		</uof:锚点_C644>
	</xsl:template>
	<xsl:template match="p:grpSp" mode="ph">
		<uof:锚点_C644>
     			<xsl:attribute name="图形引用_C62E">
				<xsl:value-of select="translate(@id,':','')"/>
			</xsl:attribute>
			<xsl:apply-templates select="p:grpSpPr/a:xfrm" mode="ph"/>
			<xsl:apply-templates select="p:nvGrpSpPr/p:nvPr/p:ph"/>
		</uof:锚点_C644>
	</xsl:template>
  
	
  
  
  
	<xsl:template match="p:pic" mode="ph">
		<uof:锚点_C644>
			<xsl:choose>
				<xsl:when test="p:spPr/a:xfrm">
					<xsl:attribute name="图形引用_C62E">
						<xsl:value-of select="translate(@id,':','')"/>
						<!--<xsl:value-of select="concat('obj_',generate-id(.))"/>-->
					</xsl:attribute>
					<xsl:apply-templates select="p:spPr/a:xfrm" mode="ph"/>
				</xsl:when>
				<xsl:when test="name(p:spPr/a:xfrm)=''">

					<xsl:call-template name="rel"/>
					<xsl:attribute name="图形引用_C62E">
						<xsl:value-of select="concat('obj_',generate-id(.))"/>
					</xsl:attribute>

					<!--重写占位符节点-->
					<xsl:apply-templates select="p:nvSpPr/p:nvPr/p:ph"/>


				</xsl:when>
			</xsl:choose>
			<!--<xsl:apply-templates select="p:spPr/a:xfrm" mode="ph"/>
			<xsl:attribute name="图形引用_C62E">
				<xsl:value-of select="translate(@id,':','')"/>
				--><!--<xsl:value-of select="concat('obj_',generate-id(.))"/>--><!--
			</xsl:attribute>-->
		</uof:锚点_C644>
	</xsl:template>
	
	<xsl:template match="p:pic" mode="sldph">
		
		<xsl:param name="layoutref"/>
		<uof:锚点_C644>
			<xsl:attribute name="图形引用_C62E">
				<xsl:value-of select="translate(@id,':','')"/>
				<!--<xsl:value-of select="concat('obj_',generate-id(.))"/>-->
			</xsl:attribute>
			<xsl:choose>
				<xsl:when test="p:spPr/a:xfrm">
					<xsl:apply-templates select="p:spPr/a:xfrm" mode="ph"/>
				</xsl:when>
				<xsl:otherwise>
					<!--
             否则则从其引用的版式中去找位置
             -->
					<xsl:for-each select="//p:sldLayout[./@id=$layoutref]/p:cSld/p:spTree/p:sp[./p:nvSpPr/p:nvPr/p:ph/@type='pic']">
						<xsl:apply-templates select="p:spPr/a:xfrm" mode="ph"/>
					</xsl:for-each>
				</xsl:otherwise>
			</xsl:choose>
		</uof:锚点_C644>
	</xsl:template>
  <!--liuyangyang 2015-03-06 添加对公式锚点的转换-->
  <xsl:template match="mc:AlternateContent" mode="ph">
    <uof:锚点_C644>
      <xsl:attribute name="是否显示缩略图_C630">
        <xsl:choose>
          <xsl:when test="./mc:Fallback//p:nvSpPr/p:nvPr/p:ph/@type='sldImg'">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of  select="'false'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="图形引用_C62E">
        <xsl:value-of select="translate(@id,':','')"/>
      </xsl:attribute>
      <xsl:choose>
        <xsl:when test="./mc:Fallback//p:spPr/a:xfrm">
          <xsl:apply-templates select="./mc:Fallback//p:spPr/a:xfrm" mode="ph"/>
          <xsl:apply-templates select="./mc:Fallback//p:nvSpPr/p:nvPr/p:ph"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="./mc:Fallback/p:sp">
            <xsl:call-template name="rel"/>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </uof:锚点_C644>
  </xsl:template>

  <!--end-->

  <!--liuyangyang 2015-03-06 添加对smartart锚点的转换-->
  <xsl:template match="dsp:sp" mode="ph">
    <!--11.11.10 李娟 锚点的属性:标识符_C62C 随动方式_C62F 是否显示缩略图_C630 有待添加-->
    <xsl:param name="x"/>
    <xsl:param name="y"/>
    <uof:锚点_C644>
      <xsl:attribute name="是否显示缩略图_C630">
        <xsl:choose>
          <xsl:when test="dsp:nvSpPr/dsp:nvPr/p:ph/@type='sldImg'">
            <xsl:value-of select="'true'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of  select="'false'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:choose>
        <xsl:when test="dsp:spPr/a:xfrm">
          <xsl:attribute name="图形引用_C62E">
            <xsl:value-of select="translate(@id,':','')"/>
          </xsl:attribute>
          <xsl:if test="dsp:spPr/a:xfrm/a:off">
            <uof:位置_C620>
              <uof:水平_4106>
                <uof:绝对_4107>
                  <xsl:attribute name="值_4108">
                    <xsl:variable name="tempX" select="dsp:spPr/a:xfrm/a:off/@x"/>
                    <xsl:value-of select="($tempX +$x) div 12700"/>
                  </xsl:attribute>
                </uof:绝对_4107>
              </uof:水平_4106>
              <!--<xsl:attribute name="uof:y坐标">
        <xsl:value-of select="a:off/@y div 12700"/>
      </xsl:attribute>-->
              <uof:垂直_410D>
                <uof:绝对_4107>
                  <xsl:attribute name="值_4108">
                    <xsl:variable name="tempY" select="dsp:spPr/a:xfrm/a:off/@y"/>
                    <xsl:value-of select="($tempY + $y) div 12700"/>
                  </xsl:attribute>
                </uof:绝对_4107>
              </uof:垂直_410D>
            </uof:位置_C620>
          </xsl:if>
          <xsl:if test="dsp:spPr/a:xfrm/a:ext">
            <uof:大小_C621>
              <xsl:attribute name="宽_C605">
                <xsl:value-of select="dsp:spPr/a:xfrm/a:ext/@cx div 12700"/>
              </xsl:attribute>
              <xsl:attribute name="长_C604">
                <xsl:value-of select="dsp:spPr/a:xfrm/a:ext/@cy div 12700"/>
              </xsl:attribute>
            </uof:大小_C621>
          </xsl:if>
          <xsl:apply-templates select="dsp:nvSpPr/dsp:nvPr/p:ph"/>

        </xsl:when>
        <!-- <xsl:when test="name(dsp:spPr/a:xfrm)=''">
          <xsl:call-template name="rel"/>
          <xsl:attribute name="图形引用_C62E">
            <xsl:value-of select="concat('obj_',generate-id(.))"/>
          </xsl:attribute>
          <xsl:apply-templates select="dsp:nvSpPr/dsp:nvPr/p:ph"/>
        </xsl:when>-->
      </xsl:choose>
      <!--<xsl:choose>
				<xsl:when test="p:nvSpPr/p:nvPr/p:ph/@type='sldImg'">
					<xsl:attribute name="是否显示缩略图_C630">true</xsl:attribute>
				</xsl:when>
			</xsl:choose>-->
    </uof:锚点_C644>
  </xsl:template>
  <!--end-->
	
	
	<xsl:template match="p:graphicFrame" mode="ph">
		<uof:锚点_C644>
			<xsl:attribute name="图形引用_C62E">
				<xsl:value-of select="translate(@id,':','')"/>
			</xsl:attribute>
			<xsl:apply-templates select="p:xfrm" mode="ph"/>
	</uof:锚点_C644>
    <!--liuyangyang 2015-03-06 对smart中的图形添加锚点-->
    <xsl:variable name="x" select="p:xfrm/a:off/@x"/>
    <xsl:variable name="y" select="p:xfrm/a:off/@y"/>
    <xsl:variable name="graphicdataType">
      <xsl:value-of select="descendant::a:graphicData/@uri"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="contains($graphicdataType,'diagram')">
        <xsl:variable name="refslide" select ="ancestor::p:sld/@id"/>
        <xsl:variable name="relationshipId" select="concat($refslide,'.rels')"/>
        <xsl:variable name="drawingId" select="concat('rId', substring-after(descendant::a:graphicData/dgm:relIds/@r:cs,'rId') + 1)"/>
        <xsl:variable name="drawingFileId" select="substring-after(ancestor::p:presentation//rel:Relationships[@id=$relationshipId]//rel:Relationship[@Id=$drawingId]/@Target, '../diagrams/')"/>
        <xsl:for-each select="//dsp:drawing[@refBy=$refslide and @id=$drawingFileId]//dsp:sp">
          <xsl:apply-templates select="." mode="ph">
            <xsl:with-param name="x" select="$x"/>
            <xsl:with-param name="y" select="$y"/>
          </xsl:apply-templates>
        </xsl:for-each>
      </xsl:when>
      <xsl:when test="contains($graphicdataType,'oleObject')">
        <xsl:apply-templates select="./a:graphic/a:graphicData/mc:AlternateContent/mc:Fallback/p:oleObj/p:pic" mode="ph"/>
      </xsl:when>
    </xsl:choose>
    <!--end-->
	</xsl:template>
	
	<xsl:template match="a:xfrm|p:xfrm" mode="ph">

		<!--x和y水平的属性 和子元素 相对值有待添加 李娟 11.11.10-->
		<xsl:if test="a:off">
			<uof:位置_C620>
				<uof:水平_4106>
					<uof:绝对_4107>
						<xsl:attribute name="值_4108">
							<xsl:value-of select="a:off/@x div 12700"/>
						</xsl:attribute>
					</uof:绝对_4107>
				</uof:水平_4106>
				<!--<xsl:attribute name="uof:y坐标">
        <xsl:value-of select="a:off/@y div 12700"/>
      </xsl:attribute>-->
				<uof:垂直_410D>
					<uof:绝对_4107>
						<xsl:attribute name="值_4108">
							<xsl:value-of select="a:off/@y div 12700"/>
						</xsl:attribute>
					</uof:绝对_4107>
				</uof:垂直_410D>
			</uof:位置_C620>
		</xsl:if>


		<xsl:if test="a:ext">
			<uof:大小_C621>
				<xsl:attribute name="宽_C605">
					<xsl:value-of select="a:ext/@cx div 12700"/>
				</xsl:attribute>
				<xsl:attribute name="长_C604">
					<xsl:value-of select="a:ext/@cy div 12700"/>
				</xsl:attribute>
			</uof:大小_C621>
		</xsl:if>

	</xsl:template>
	<xsl:template match="p:ph">
		<xsl:variable name="phType" select="@type"/>

		<!--<xsl:if test="$phType='sldImg'">
			<xsl:attribute name="是否显示缩略图_C630">true</xsl:attribute>
		</xsl:if>-->
		<uof:占位符_C626>
		<xsl:attribute name="类型_C627">
			<xsl:choose>
				<xsl:when test="name($phType)='' or $phType='obj'or not(name($phType))">
					<xsl:value-of select="'object'"/>
				</xsl:when>
				<xsl:when test="$phType='title' and @orient='vert'">
					<xsl:value-of select="'vertical_title'"/>
				</xsl:when>
				<xsl:when test="$phType='title' and name(@orient)=''">
					<xsl:value-of select="'title'"/>
				</xsl:when>
				<xsl:when test="$phType='body'">
					<xsl:choose>
						<xsl:when test="name(ancestor::p:cSld/..)='p:notesMaster' or name(ancestor::p:cSld/..)='p:notes'">notes</xsl:when>
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="@orient='vert'">
									<xsl:value-of select="'vertical_text'"/>
								</xsl:when>
								<xsl:when test="@orient=''or not(@orient)">
									<xsl:value-of select="'text'"/>
								</xsl:when>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$phType='ctrTitle'">
					<xsl:value-of select="'centertitle'"/>
				</xsl:when>
				<xsl:when test="$phType='subTitle' and @orient='vert'">
					<xsl:value-of select="'vertical_subtitle'"/>
				</xsl:when>
				<xsl:when test="$phType='subTitle' and name(@orient)=''">
					<xsl:value-of select="'subtitle'"/>
				</xsl:when>
				<xsl:when test="$phType='dt'">
					<xsl:value-of select="'date'"/>
				</xsl:when>
				<xsl:when test="$phType='sldNum'">
					<xsl:value-of select="'number'"/>
				</xsl:when>
				<xsl:when test="$phType='ftr'">
					<xsl:value-of select="'footer'"/>
				</xsl:when>
				<xsl:when test="$phType='hdr'">
					<xsl:value-of select="'header'"/>
				</xsl:when>
				<xsl:when test="$phType='chart'">
					<xsl:value-of select="'chart'"/>
				</xsl:when>
				<xsl:when test="$phType='tbl'">
					<xsl:value-of select="'table'"/>
				</xsl:when>
				<xsl:when test="$phType='clipArt'">
					<xsl:value-of select="'clipart'"/>
				</xsl:when>
				<xsl:when test="$phType='media'">
					<xsl:value-of select="'media_clip'"/>
				</xsl:when>
				<xsl:when test="$phType='dgm'">
					<xsl:value-of select="'object'"/>
				</xsl:when>
				<!--xsl:when test="$phType='sldImg'">
					<xsl:value-of select="'object'"/>
				</xsl:when-->
				<xsl:when test="$phType='pic'">
					<xsl:value-of select="'object'"/>
				</xsl:when>
				<xsl:when test="$phType='sldImg'">
					<xsl:value-of select="'notes'"/>
				</xsl:when>
			</xsl:choose>
		</xsl:attribute>
		</uof:占位符_C626>
		<!-- 09.10.23 马有旭 添加 拷贝其他属性 -->
		<!--<xsl:if test="@sz">
      <xsl:attribute name="sz">
        <xsl:value-of select="@sz"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="@idx">
      <xsl:attribute name="idx">
        <xsl:value-of select="@idx"/>
      </xsl:attribute>
    </xsl:if>-->

	</xsl:template>
	<xsl:template name="rel">
		<xsl:variable name="cSld" select="ancestor::p:cSld"/>
		<xsl:variable name="ParentcSld" select="$cSld/.."/>
		<xsl:variable name="phType" select="p:nvSpPr/p:nvPr/p:ph/@type"/>
		<xsl:variable name="phIdx" select="p:nvSpPr/p:nvPr/p:ph/@idx"/>
		<xsl:variable name="sldRel" select="concat($ParentcSld/@id,'.rels')"/>
		<xsl:if test="name($ParentcSld)='p:sld'">
			<xsl:variable name="rootTar" select="//rel:Relationships[@id=$sldRel]/rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout']/@Target"/>
			<xsl:variable name="rootID" select="substring-after($rootTar,'slideLayouts/')"/>
			<xsl:variable name="rootEleSp" select="//p:sldLayout[@id=$rootID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx = $phIdx or p:nvSpPr/p:nvPr/p:ph/@type = $phType]"/>
			<!--xsl:if test="$rootEleSp/p:spPr/a:xfrm">
        <xsl:apply-templates select="$rootEleSp/p:spPr/a:xfrm" mode="ph"/>
      </xsl:if>
      <xsl:if test="name($rootEleSp/p:spPr/a:xfrm)=''">
        <xsl:apply-templates select="$rootEleSp" mode="rel"/>
      </xsl:if-->
			<!--罗文甜：10.12.02修改-->
			<xsl:choose>
				<xsl:when test="$rootEleSp/p:spPr/a:xfrm">
					<xsl:apply-templates select="$rootEleSp/p:spPr/a:xfrm" mode="ph"/>
				</xsl:when>
				<xsl:when test="not($rootEleSp)">
					<xsl:variable name="rootID1" select="concat($rootID,'.rels')"/>
					<xsl:variable name="rootTar1" select="//rel:Relationships[@id=$rootID1]/rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster']/@Target"/>
					<xsl:variable name="rootIDm" select="substring-after($rootTar1,'s/')"/>
					<xsl:variable name="rootEleSpm" select="//p:sldMaster[@id=$rootIDm]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type = $phType]"/>
					<xsl:apply-templates select="$rootEleSpm/p:spPr/a:xfrm" mode="ph"/>
				</xsl:when>
			</xsl:choose>
		</xsl:if>
		<xsl:if test="name($ParentcSld)='p:sldLayout'">
			<xsl:variable name="rootTar" select="//rel:Relationships[@id=$sldRel]/rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster']/@Target"/>
			<xsl:variable name="rootID" select="substring-after($rootTar,'s/')"/>
			<!--xsl:variable name="rootEle" select="//p:sldMaster[@id=$rootID]"/-->
			<xsl:variable name="rootEleSp" select="//p:sldMaster[@id=$rootID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx = $phIdx or p:nvSpPr/p:nvPr/p:ph/@type = $phType]"/>
			<xsl:if test="$rootEleSp/p:spPr/a:xfrm">
				<xsl:apply-templates select="$rootEleSp/p:spPr/a:xfrm" mode="ph"/>
			</xsl:if>
			<xsl:if test="name($rootEleSp/p:spPr/a:xfrm)=''">
				<xsl:apply-templates select="$rootEleSp" mode="rel"/>
			</xsl:if>
		</xsl:if>
	</xsl:template>
	<xsl:template match="p:sp" mode="rel">
		<xsl:call-template name="rel"/>
	</xsl:template>
</xsl:stylesheet>
