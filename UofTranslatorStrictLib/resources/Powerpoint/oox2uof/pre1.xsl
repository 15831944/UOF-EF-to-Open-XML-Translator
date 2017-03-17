<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
				xmlns:fo="http://www.w3.org/1999/XSL/Format"
				xmlns:dc="http://purl.org/dc/elements/1.1/"
				xmlns:dcterms="http://purl.org/dc/terms/"
				xmlns:dcmitype="http://purl.org/dc/dcmitype/"
				xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
				xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
        
				xmlns:cus="http://purl.oclc.org/ooxml/officeDocument/customProperties"
        xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
				xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
				xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
				xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
        xmlns:dgm="http://purl.oclc.org/ooxml/drawingml/diagram"
                
				xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
        xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
				xmlns:uof="http://schemas.uof.org/cn/2009/uof"
				xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
				xmlns:演="http://schemas.uof.org/cn/2009/presentation"
				xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
				xmlns:图="http://schemas.uof.org/cn/2009/graph">

	<xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
	<!--presentation模板-->
	<xsl:template name="presentation" match="/">
		<xsl:for-each select="p:presentation">
			<p:presentation>
				<!-- 把presenation所有的引用都添加过来 李娟 11.12.26-->
				<xsl:variable name="rels" select="document('../_rels/.rels')"/>
				<xsl:variable name="pptrels" select="document('_rels/presentation.xml.rels')"/>
				<xsl:copy-of select="@*"/>
				<xsl:copy-of select="$pptrels"/>
				<xsl:copy-of select="document('viewProps.xml')"/>
				<!--把../_rels/.rels引用的东西复制过来 李娟 11.12.26-->
				<xsl:if test="$rels/rel:Relationships/rel:Relationship[@Target='docProps/app.xml']">
					<xsl:copy-of select="document('../docProps/app.xml')"/>
				</xsl:if>
				<xsl:if test="$rels/rel:Relationships/rel:Relationship[@Target='docProps/core.xml']">
					<xsl:copy-of select="document('../docProps/core.xml')"/>
				</xsl:if>
				<xsl:if test="$rels/rel:Relationships/rel:Relationship[@Target='docProps/custom.xml']">
					<xsl:copy-of select="document('../docProps/custom.xml')"/>
				</xsl:if>
				<xsl:if test="$pptrels/rel:Relationships/rel:Relationship[@Target='commentAuthors.xml']">
					<xsl:copy-of select="document('../ppt/commentAuthors.xml')"/>
					
				</xsl:if>
				<!--把preprops 下面的东西引用过来 李娟11.12.26-->
				<xsl:copy-of select="document('presProps.xml')"/>
				<xsl:for-each select="p:sldMasterIdLst/p:sldMasterId|p:notesMasterIdLst/p:notesMasterId |p:handoutMasterIdLst/p:handoutMasterId | p:sldIdLst/p:sldId">
					<xsl:variable name="Type" select="$pptrels/rel:Relationships/*[@Id=current()/@r:id]/@Type"/>
					
					<xsl:variable name="Target" select="$pptrels/rel:Relationships/*[@Id=current()/@r:id]/@Target"/>
					<!--sldMaster列表 各sldMater所包含的sldLayout-->
					
						
					
					<xsl:for-each select="document(string($Target))/*">
						<xsl:call-template name="sldMaster-sld">
							<xsl:with-param name="ID" select="substring-after($Target,'/')"/>
						</xsl:call-template>
					</xsl:for-each>
					<!--sldMaster_rels列表-->
					<!--<xsl:variable name="sldTarget" select="document(concat(substring-before($Target,'/'),'/_rels/',substring-after($Target,'/'),'.rels'))/rel:Relationships">-->
					<!--<xsl:for-each select="document(concat(substring-before($Target,'/'),'/_rels/',substring-after($Target,'/'),'.rels'))/rel:Relationships">-->
					<xsl:apply-templates select="document(concat(substring-before($Target,'/'),'/_rels/',substring-after($Target,'/'),'.rels'))/rel:Relationships">
						<xsl:with-param name="relsID" select="concat(substring-after($Target,'/'),'.rels')"/>
					</xsl:apply-templates>
					<!--</xsl:for-each>-->
				</xsl:for-each>
				<xsl:apply-templates select="p:defaultTextStyle"/>
				
				<!--
2010.1.13 黎美秀修改 修改综合案例放映方式
<xsl:copy-of select="*[name()!='p:sldMasterIdLst' and name()!='p:notesMasterIdLst'  and name()!='p:handoutMasterIdLst' and name()!='p:sldIdLst' and name()!='p:defaultTextStyle']"/>
-->
				<xsl:copy-of select="*[name()!='p:sldMasterIdLst' and name()!='p:notesMasterIdLst'  and name()!='p:handoutMasterIdLst' and name()!='p:defaultTextStyle']"/>
				
			</p:presentation>
		</xsl:for-each>
	</xsl:template>
	<!--缺省文本样式模板-->
	<xsl:template match="p:defaultTextStyle">
		<xsl:call-template name="styles"/>
	</xsl:template>
	<!--sldMaster-sld-->
	<xsl:template name="sldMaster-sld">
		<xsl:param name="ID"/>
		<xsl:copy>
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:attribute name="id">
				<xsl:value-of select="$ID"/>
			</xsl:attribute>
			<xsl:apply-templates select="p:cSld"/>
			<xsl:copy-of select="*[name()!='p:cSld' and name()!='p:txStyles' and name()!='p:notesStyle']"/>
			<xsl:for-each select="p:txStyles">
				<xsl:copy>
					<xsl:for-each select="*">
						<xsl:call-template name="styles"/>
					</xsl:for-each>
				</xsl:copy>
			</xsl:for-each>
			<xsl:for-each select="p:notesStyle">
				<xsl:call-template name="styles"/>
			</xsl:for-each>
			<xsl:apply-templates select="p:sldLayoutIdLst">
				<xsl:with-param name="sldMasterID" select="$ID"/>
			</xsl:apply-templates>
		</xsl:copy>
	</xsl:template>
	<!--xsl:template match="p:txStyles" mode="addID">
		<xsl:copy>
			<xsl:for-each select="*">
				<xsl:call-template name="styles"/>
			</xsl:for-each>
		</xsl:copy>
	</xsl:template-->
	<!--样式模板-->
	<xsl:template name="styles">
		<xsl:copy>
			<xsl:for-each select="*">
				<xsl:copy>
					<xsl:attribute name="id">
						<xsl:value-of select="translate(generate-id(.),':','')"/>
					</xsl:attribute>
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
					<xsl:for-each select="*">
						<xsl:copy-of select="."/>
					</xsl:for-each>
				</xsl:copy>
			</xsl:for-each>
		</xsl:copy>
	</xsl:template>
	<!--p:cSld模板-->
	<xsl:template match="p:cSld">
		<p:cSld>
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:if test="name(.)!='p:spTree'">
					<xsl:copy-of select="."/>
				</xsl:if>
				<xsl:if test="name(.)='p:spTree'">
					<xsl:apply-templates select="."/>
				</xsl:if>
			</xsl:for-each>
		</p:cSld>
	</xsl:template>
	<xsl:template match="p:spTree">
		<p:spTree>
			<xsl:call-template name="sps"/>
		</p:spTree>
	</xsl:template>
	<!--
      2010.3.14  黎美秀修改 增加文字表的处理
      <xsl:template name="sps">
    <xsl:attribute name="id">
      <xsl:value-of select="generate-id()"/>
    </xsl:attribute>
    <xsl:for-each select="*">
      <xsl:if test="name(.)='p:nvGrpSpPr'">
        <xsl:copy-of select="."/>
      </xsl:if>
      <xsl:if test="name(.)='p:grpSpPr'">
        <xsl:copy-of select="."/>
      </xsl:if>
      <xsl:if test="name(.)='p:sp'">
        <xsl:apply-templates select="."/>
      </xsl:if>
      <xsl:if test="name(.)='p:grpSp'">
        <xsl:apply-templates select="."/>
      </xsl:if>    
     <xsl:if test="name(.)!='p:nvGrpSpPr' and name(.)!='p:grpSpPr'and name(.)!='p:sp' and name(.)!='p:grpSp'">
        <xsl:copy>
          <xsl:attribute name="id">
            <xsl:value-of select="generate-id()"/>
          </xsl:attribute>
          <xsl:copy-of select="./*"/>
        </xsl:copy>
      </xsl:if>
    </xsl:for-each>
  </xsl:template> 
      -->
	<!--文字表的处理-->
  <!--新增功能点smartart liqiuling 2012.12.11 start-->
	<xsl:template name="sps">
		<xsl:attribute name="id">
			<xsl:value-of select="generate-id()"/>
		</xsl:attribute>
		<xsl:for-each select="*">
			<xsl:choose>
				<xsl:when test="name(.)='p:nvGrpSpPr'or name(.)='dsp:nvGrpSpPr'">
					<xsl:copy-of select="."/>
				</xsl:when>
				<xsl:when test="name(.)='p:grpSpPr'">
					<xsl:copy-of select="."/>
				</xsl:when>
				<xsl:when test="name(.)='p:sp' or name(.)='dsp:sp'">
					<xsl:apply-templates select="."/>
				</xsl:when>
				<xsl:when test="name(.)='p:grpSp'">
					<xsl:apply-templates select="."/>
				</xsl:when>
				<xsl:when test="name(.)='p:graphicFrame'">
					<xsl:call-template name="graphicFrame"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:copy>
						<xsl:attribute name="id">
							<xsl:value-of select="generate-id()"/>
						</xsl:attribute>
						<xsl:copy-of select="./*"/>
					</xsl:copy>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:for-each>
	</xsl:template>
	<!--graphicFrame模板-->
	<xsl:template name="graphicFrame">
		<xsl:copy>
			<xsl:attribute name="id">
				<xsl:value-of select="generate-id()"/>
			</xsl:attribute>
			<xsl:choose>
				<xsl:when test="not(.//a:tbl)">
					<xsl:copy-of select="./*"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
					<xsl:for-each select="*">
						<xsl:choose>
							<xsl:when test="name(.)!='a:graphic'">
								<xsl:copy-of select="."/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:apply-templates select=".">

								</xsl:apply-templates>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:for-each>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:copy>
		<xsl:if test=".//a:tbl">
			<xsl:copy-of select="document('tableStyles.xml')/a:tblStyleLst"/>
		</xsl:if>
	</xsl:template>
	<!--end-->

	<xsl:template match="a:graphic">
		<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
			<a:tbl>
				<xsl:apply-templates select=".//a:tbl"/>
			</a:tbl>
		</a:graphicData>
	</xsl:template>

	<xsl:template match="a:tbl">
		<xsl:for-each select="@*">
			<xsl:copy/>
		</xsl:for-each>
		<xsl:for-each select="*">
			<xsl:if test="name(.)!='a:tr'">
				<xsl:copy-of select="."/>
			</xsl:if>
			<xsl:if test="name(.)='a:tr'">
				<xsl:apply-templates select=".">
				</xsl:apply-templates>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="a:tr">
		<a:tr>
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:choose>
					<xsl:when test="name(.)!='a:tc'">
						<xsl:copy-of select="."/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:apply-templates select="."/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
		</a:tr>
	</xsl:template>
	<xsl:template match="a:tc">
		<a:tc>
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:choose>
					<xsl:when test="name(.)!='a:txBody'">
						<xsl:copy-of select="."/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:apply-templates select="."/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
		</a:tc>
	</xsl:template>
	<xsl:template match="a:txBody">
		<a:txBody>
			<xsl:for-each select="*">
				<xsl:choose>
					<xsl:when test="name(.)!='a:p'">
						<xsl:choose>
							<xsl:when test="name(.)!='a:lstStyle'">
								<xsl:copy-of select="."/>
							</xsl:when>
							<xsl:when test="name(.)='a:lstStyle'">
								<xsl:call-template name="styles"/>
							</xsl:when>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="tblpara"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
		</a:txBody>
	</xsl:template>
	<!--tblpara模板-->
	<xsl:template name="tblpara">
		<a:p>
			<xsl:attribute name="id">
				<xsl:value-of select="generate-id()"/>
			</xsl:attribute>
			<xsl:copy-of select="./*"/>
		</a:p>
	</xsl:template>
	<xsl:template match="p:grpSp">
		<p:grpSp>
			<xsl:call-template name="sps"/>
		</p:grpSp>
	</xsl:template>
  
	<xsl:template match="p:sp">
		<p:sp>
			
			<xsl:attribute name="id">
				
				<xsl:value-of select="generate-id()"/>
			</xsl:attribute>
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:if test="name(.)!='p:txBody'">
					<xsl:copy-of select="."/>
				</xsl:if>
				<xsl:if test="name(.)='p:txBody'">
					<xsl:apply-templates select=".">
						<xsl:with-param name="pht" select="../p:nvSpPr/p:nvPr/p:ph/@type"/>
					</xsl:apply-templates>
				</xsl:if>
			</xsl:for-each>
		</p:sp>
	</xsl:template>
	<xsl:template match="p:txBody">
		<xsl:param name="pht"/>
		<p:txBody>
			<xsl:for-each select="*">
				<xsl:if test="name(.)!='a:p'">
					<xsl:if test="name(.)!='a:lstStyle'">
						<xsl:copy-of select="."/>
					</xsl:if>
					<xsl:if test="name(.)='a:lstStyle'">
						<xsl:call-template name="styles"/>
					</xsl:if>
				</xsl:if>
        
        <!--2014-06-05, tangjiang, 将只有段落结束符的段落预处理为含有一个空格符的段落，修复文字位置转换差异 start -->
        <xsl:choose>
          <xsl:when test="name(.)='a:p' and ./a:endParaRPr/@sz and not(./a:r)">
            <a:p>
              <a:pPr>
                <a:buNone/>
              </a:pPr>
              <a:r>
                <a:rPr lang="en-US" dirty="0" smtClean="0">
                  <xsl:attribute name="sz">
                    <xsl:value-of select="./a:endParaRPr/@sz"/>
                  </xsl:attribute>
                </a:rPr>
                <a:t xml:space="preserve"><xsl:value-of select="' '"/></a:t>
              </a:r>
              <a:endParaRPr lang="en-US" dirty="0" smtClean="0">
                <xsl:attribute name="sz">
                  <xsl:value-of select="./a:endParaRPr/@sz"/>
                </xsl:attribute>
              </a:endParaRPr>
            </a:p>
          </xsl:when>
          <xsl:when test="name(.)='a:p'">
            <xsl:apply-templates select=".">
              <xsl:with-param name="pht" select="$pht"/>
            </xsl:apply-templates>
          </xsl:when>
        </xsl:choose>
        <!--2014-06-05, tangjiang, 将只有段落结束符的段落预处理为含有一个空格符的段落，修复文字位置转换差异 end -->
        
			</xsl:for-each>
		</p:txBody>
	</xsl:template>
	<xsl:template match="a:p">
		<xsl:param name="pht"/>
		<a:p>
			<xsl:attribute name="id">
				<xsl:value-of select="generate-id()"/>
			</xsl:attribute>
			<xsl:if test="ancestor::p:sldMaster and not(ancestor::p:sldLayout)">
				<xsl:attribute name="styleID">
					<!--  10.22 黎美秀修改 增加fld的情况 此时格式为 p/fld/ppr 而不是普通的 p/ppr-->
					<xsl:choose>
						<xsl:when test="$pht='title'">
							<xsl:if test="preceding-sibling::a:lstStyle/*">
								<xsl:variable name="pPrID">
									<xsl:call-template name="lstStyle"/>
								</xsl:variable>
								<!--09.10.10 马有旭 添加 <xsl:value-of select="$pPrID"/> -->
								<xsl:value-of select="$pPrID"/>
							</xsl:if>
							<xsl:if test="not(preceding-sibling::a:lstStyle/*)">
								<xsl:if test="not(a:pPr/@lvl)">
									<xsl:value-of select="translate(generate-id(ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr),':','')"/>
								</xsl:if>
							</xsl:if>
						</xsl:when>
						<xsl:when test="$pht='body'">
							<xsl:if test="preceding-sibling::a:lstStyle/*">
								<xsl:call-template name="lstStyle"/>
							</xsl:if>
							<xsl:if test="not(preceding-sibling::a:lstStyle/*)">
								<xsl:call-template name="bodyStyle"/>
							</xsl:if>
						</xsl:when>
						<xsl:otherwise>
							<xsl:if test="preceding-sibling::a:lstStyle/*">
								<xsl:call-template name="lstStyle"/>
							</xsl:if>
							<xsl:if test="not(preceding-sibling::a:lstStyle/*)">
								<xsl:call-template name="otherStyles"/>
							</xsl:if>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test="ancestor::p:notesMaster">
				<xsl:attribute name="styleID">
					<xsl:if test="preceding-sibling::a:lstStyle/*">
						<xsl:call-template name="lstStyle"/>
					</xsl:if>
					<xsl:if test="not(preceding-sibling::a:lstStyle/*)">
						<xsl:call-template name="notesStyle"/>
					</xsl:if>
				</xsl:attribute>
			</xsl:if>
			<!--
      2010.4.25 黎美秀修改
      <xsl:copy-of select="./*"/>
    </a:p>
  </xsl:template>        
       -->
			<xsl:for-each select ="*">
				<xsl:choose>
					<xsl:when test ="not(name(.)='a:r')">
						<xsl:copy-of select ="."/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name ="run"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
		</a:p>
	</xsl:template>
	<xsl:template name ="run">
		<a:r>
			<xsl:copy-of select ="@*"/>
			<xsl:for-each select ="*">
				<xsl:choose>
					<xsl:when test ="not(name(.)='a:t')">
						<xsl:copy-of select ="."/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:apply-templates select ="."/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
		</a:r>
	</xsl:template>
	<xsl:template match ="a:t">
		<a:t xml:space="preserve"><xsl:value-of select ="."/></a:t>
	</xsl:template>

	<!--lstStyle模板-->
	<xsl:template name="lstStyle">
		<!-- 10.22 黎美秀修改 增加判断增加fld的情况 此时格式为 p/fld/ppr 而不是普通的 p/ppr
    
     <xsl:variable name="pPrID">
      <xsl:choose>
        <xsl:when test="a:pPr/@lvl='0'">
          <xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl1pPr)"/>
        </xsl:when>
        <xsl:when test="a:pPr/@lvl='1'">
          <xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl2pPr)"/>
        </xsl:when>
        <xsl:when test="a:pPr/@lvl='2'">
          <xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl3pPr)"/>
        </xsl:when>
        <xsl:when test="a:pPr/@lvl='3'">
          <xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl4pPr)"/>
        </xsl:when>
        <xsl:when test="a:pPr/@lvl='4'">
          <xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl5pPr)"/>
        </xsl:when>
        <xsl:when test="a:pPr/@lvl='5'">
          <xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl6pPr)"/>
        </xsl:when>
        <xsl:when test="a:pPr/@lvl='6'">
          <xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl7pPr)"/>
        </xsl:when>
        <xsl:when test="a:pPr/@lvl='7'">
          <xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl8pPr)"/>
        </xsl:when>
        <xsl:when test="a:pPr/@lvl='8'">
          <xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl9pPr)"/>
        </xsl:when>
        <xsl:when test="not(a:pPr/@lvl)">
          <xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl1pPr)"/>
        </xsl:when>
      </xsl:choose>
    </xsl:variable>
    
    -->
		<xsl:variable name="pPrID">
			<xsl:choose>
				<xsl:when test="a:pPr/@lvl='0' or a:fld/a:pPr/@lvl='0'">
					<xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl1pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='1' or a:fld/a:pPr/@lvl='1'">
					<xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl2pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='2' or a:fld/a:pPr/@lvl='2'">
					<xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl3pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='3' or a:fld/a:pPr/@lvl='3'">
					<xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl4pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='4' or a:fld/a:pPr/@lvl='4'">
					<xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl5pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='5' or a:fld/a:pPr/@lvl='5'">
					<xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl6pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='6' or a:fld/a:pPr/@lvl='6'">
					<xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl7pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='7'or a:fld/a:pPr/@lvl='7'">
					<xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl8pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='8' or a:fld/a:pPr/@lvl='8'">
					<xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl9pPr)"/>
				</xsl:when>
				<xsl:when test="not(a:pPr/@lvl) or not(a:fld/a:pPr/@lvl)">
					<xsl:value-of select="generate-id(preceding-sibling::a:lstStyle/a:lvl1pPr)"/>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<xsl:value-of select="translate($pPrID,':','')"/>
	</xsl:template>

	<!--otherStyles模板-->
	<xsl:template name="otherStyles">
		<xsl:variable name="pPrID">
			<xsl:choose>
				<xsl:when test="a:pPr/@lvl='0'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl1pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='1'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl2pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='2'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl3pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='3'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl4pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='4'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl5pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='5'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl6pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='6'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl7pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='7'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl8pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='8'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl9pPr)"/>
				</xsl:when>
				<xsl:when test="not(a:pPr/@lvl)">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl1pPr)"/>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<xsl:value-of select="translate($pPrID,':','')"/>
	</xsl:template>
	<!--bodyStyle模板-->
	<xsl:template name="bodyStyle">
		<xsl:variable name="pPrID">
			<xsl:choose>
				<xsl:when test="a:pPr/@lvl='0'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='1'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl2pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='2'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl3pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='3'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl4pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='4'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl5pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='5'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl6pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='6'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl7pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='7'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl8pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='8'">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl9pPr)"/>
				</xsl:when>
				<xsl:when test="not(a:pPr/@lvl)">
					<xsl:value-of select="generate-id(ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr)"/>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<xsl:value-of select="translate($pPrID,':','')"/>
	</xsl:template>
	<!--notesStyle模板-->
	<xsl:template name="notesStyle">
		<xsl:variable name="pPrID">
			<xsl:choose>
				<xsl:when test="a:pPr/@lvl='0'">
					<xsl:value-of select="generate-id(ancestor::p:notesMaster/p:notesStyle/a:lvl1pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='1'">
					<xsl:value-of select="generate-id(ancestor::p:notesMaster/p:notesStyle/a:lvl2pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='2'">
					<xsl:value-of select="generate-id(ancestor::p:notesMaster/p:notesStyle/a:lvl3pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='3'">
					<xsl:value-of select="generate-id(ancestor::p:notesMaster/p:notesStyle/a:lvl4pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='4'">
					<xsl:value-of select="generate-id(ancestor::p:notesMaster/p:notesStyle/a:lvl5pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='5'">
					<xsl:value-of select="generate-id(ancestor::p:notesMaster/p:notesStyle/a:lvl6pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='6'">
					<xsl:value-of select="generate-id(ancestor::p:notesMaster/p:notesStyle/a:lvl7pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='7'">
					<xsl:value-of select="generate-id(ancestor::p:notesMaster/p:notesStyle/a:lvl8pPr)"/>
				</xsl:when>
				<xsl:when test="a:pPr/@lvl='8'">
					<xsl:value-of select="generate-id(ancestor::p:notesMaster/p:notesStyle/a:lvl9pPr)"/>
				</xsl:when>
				<xsl:when test="not(a:pPr/@lvl)">
					<xsl:value-of select="generate-id(ancestor::p:notesMaster/p:notesStyle/a:lvl1pPr)"/>
				</xsl:when>
			</xsl:choose>
		</xsl:variable>
		<xsl:value-of select="translate($pPrID,':','')"/>
	</xsl:template>
	<xsl:template match="p:sldLayoutIdLst">
		<xsl:param name="sldMasterID"/>
		<xsl:for-each select="p:sldLayoutId">
			<!--条件中的属性值限定时需[@Id=$rid]，因为xsl不能由[@Id=@r:id]判定到底是哪个的属性-->
			<xsl:variable name="rid" select="@r:id"/>
			<xsl:variable name="relShip" select="document(concat('slideMasters/_rels/',$sldMasterID,'.rels'))/rel:Relationships/rel:Relationship[@Id=$rid]"/>
			<xsl:variable name="Target" select="$relShip/@Target"/>
			<xsl:for-each select="document(substring-after($Target,'../'))/p:sldLayout">
				<xsl:call-template name="sldMaster-sld">
					<xsl:with-param name="ID" select="substring-after($Target,'slideLayouts/')"/>
					<xsl:with-param name="type" select="'sldLayout'"/>
				</xsl:call-template>
			</xsl:for-each>
			<!--sldLayout的rel列表-->
			<xsl:apply-templates select="document(concat('slideLayouts/_rels/',substring-after($Target,'slideLayouts/'),'.rels'))/rel:Relationships">
				<xsl:with-param name="relsID" select="concat(substring-after($Target,'slideLayouts/'),'.rels')"/>
			</xsl:apply-templates>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="rel:Relationships">
		
		<xsl:param name="relsID"/>
		<xsl:copy>
			<xsl:attribute name="id">
				<xsl:value-of select="$relsID"/>
			</xsl:attribute>
			<xsl:copy-of select="*"/>
		</xsl:copy>
		
		<xsl:for-each select="rel:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/notesSlide']">
			
			<xsl:variable name="Target" select="@Target"/>
			<xsl:for-each select="document(substring-after($Target,'../'))/*">
				<xsl:call-template name="sldMaster-sld">
					<xsl:with-param name="ID" select="substring-after($Target,'notesSlides/')"/>
				</xsl:call-template>
			</xsl:for-each>
			<xsl:for-each select="document(concat('notesSlides/_rels/',substring-after($Target,'notesSlides/'),'.rels'))/*">
				<xsl:copy>
					<xsl:attribute name="id">
						<xsl:value-of select="concat(substring-after($Target,'notesSlides/'),'.rels')"/>
					</xsl:attribute>
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
					<xsl:for-each select="*">
						<xsl:copy-of select="."/>
					</xsl:for-each>
				</xsl:copy>
			</xsl:for-each>
		</xsl:for-each>
	<!--添加讲义母版 李娟 2012.04.20-->
		<xsl:for-each select="rel:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/handoutMaster']">
			
			<xsl:variable name="Target" select="@Target"/>
			<xsl:for-each select="document(substring-after($Target,'../'))/*">
				<xsl:call-template name="sldMaster-sld">
					<xsl:with-param name="ID" select="substring-after($Target,'handoutMasters/')"/>
				</xsl:call-template>
			</xsl:for-each>
			<xsl:for-each select="document(concat('notesSlides/_rels/',substring-after($Target,'handoutMasters/'),'.rels'))/*">
				<xsl:copy>
					<xsl:attribute name="id">
						<xsl:value-of select="concat(substring-after($Target,'handoutMasters/'),'.rels')"/>
					</xsl:attribute>
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
					<xsl:for-each select="*">
						<xsl:copy-of select="."/>
					</xsl:for-each>
				</xsl:copy>
			</xsl:for-each>
		</xsl:for-each>
		
		<xsl:for-each select="rel:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/theme' or @Type='http://purl.oclc.org/ooxml/officeDocument/relationships/themeOverride']">
			<xsl:variable name="fileName" select="@Target"/>
			<xsl:variable name="refBy" select="substring-before($relsID,'.rels')"/>
			<xsl:for-each select="document(substring-after($fileName,'../'))/*">
				<xsl:copy>
					<xsl:attribute name="id">
						<xsl:value-of select="substring-after($fileName,'theme/')"/>
					</xsl:attribute>
					<xsl:attribute name="refBy">
						<xsl:value-of select="$refBy"/>
					</xsl:attribute>
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
					<xsl:for-each select="*">
						<xsl:copy-of select="."/>
					</xsl:for-each>
				</xsl:copy>
			</xsl:for-each>
		</xsl:for-each>
		<xsl:for-each select="rel:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/comments']">
			<xsl:variable name="commentName" select="@Target"/>
			<xsl:variable name="refBy" select="substring-before($relsID,'.rels')"/>
			<xsl:for-each select="document(substring-after($commentName,'../'))/*">
				<xsl:copy>
					<xsl:attribute name="id">
						<xsl:value-of select="substring-after($commentName,'comments/')"/>
					</xsl:attribute>
					<xsl:attribute name="refBy">
						<xsl:value-of select="$refBy"/>
					</xsl:attribute>
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
					<xsl:for-each select="*">
						<xsl:copy-of select="."/>
					</xsl:for-each>
				</xsl:copy>
			</xsl:for-each>
		</xsl:for-each>
    <!--2012-12-20, liqiuling, 解决OOXML到UOF新增功能点SmartArt  start -->
    <xsl:for-each select="rel:Relationship[@Type='http://schemas.microsoft.com/office/2007/relationships/diagramDrawing']">
      <xsl:variable name="diagramDrawingName" select="@Target"/>
      <xsl:variable name="refBy" select="substring-before($relsID,'.rels')"/>
      <xsl:for-each select="document(substring-after($diagramDrawingName,'../'))/*">
        <xsl:copy>
          <xsl:attribute name="id">
            <xsl:value-of select="substring-after($diagramDrawingName,'diagrams/')"/>
          </xsl:attribute>
          <xsl:attribute name="refBy">
            <xsl:value-of select="$refBy"/>
          </xsl:attribute>
          <xsl:for-each select="@*">
            <xsl:copy/>
          </xsl:for-each>
           <xsl:for-each select="*">
            <xsl:if test="name(.)!='dsp:spTree'">
              <xsl:copy-of select="."/>
            </xsl:if>
            <xsl:if test="name(.)='dsp:spTree'">
              <xsl:apply-templates select="."/>
            </xsl:if>
          </xsl:for-each>
        </xsl:copy>
      </xsl:for-each>
    </xsl:for-each>
 
	</xsl:template>
  
  <xsl:template match="dsp:spTree">
		<dsp:spTree>
			<xsl:call-template name="sps"/>
		</dsp:spTree>
	</xsl:template>

  <xsl:template match="dsp:sp">
    <dsp:sp>

      <xsl:attribute name="id">

        <xsl:value-of select="generate-id()"/>
      </xsl:attribute>
      <xsl:for-each select="@*">
        <xsl:copy/>
      </xsl:for-each>
      <xsl:for-each select="*">
        <xsl:if test="name(.)!='dsp:txBody'">
          <xsl:copy-of select="."/>
        </xsl:if>
        <xsl:if test="name(.)='dsp:txBody'">
          <xsl:apply-templates select=".">
            <xsl:with-param name="pht" select="../dsp:nvSpPr/dsp:nvPr/dsp:ph/@type"/>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:for-each>
    </dsp:sp>
  </xsl:template>

  <xsl:template match="dsp:txBody">
    <xsl:param name="pht"/>
    <dsp:txBody>
      <xsl:for-each select="*">
        <xsl:if test="name(.)!='a:p'">
          <xsl:if test="name(.)!='a:lstStyle'">
            <xsl:copy-of select="."/>
          </xsl:if>
          <xsl:if test="name(.)='a:lstStyle'">
            <xsl:call-template name="styles"/>
          </xsl:if>
        </xsl:if>
        <xsl:if test="name(.)='a:p'">
          <xsl:apply-templates select=".">
            <xsl:with-param name="pht" select="$pht"/>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:for-each>
    </dsp:txBody>
  </xsl:template>
  <!--end-->
	
	<!--批注模板 李娟 2012.03.15-->
	
</xsl:stylesheet>
