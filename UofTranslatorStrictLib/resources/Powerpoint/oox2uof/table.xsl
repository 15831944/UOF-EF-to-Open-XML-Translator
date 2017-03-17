<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
				xmlns:fo="http://www.w3.org/1999/XSL/Format"
				xmlns:dc="http://purl.org/dc/elements/1.1/" 
				xmlns:dcterms="http://purl.org/dc/terms/"
				xmlns:dcmitype="http://purl.org/dc/dcmitype/"
				xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
				xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
        xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
        xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
				xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
				xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" 
				xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
                
				xmlns:uof="http://schemas.uof.org/cn/2009/uof"
				xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
				xmlns:演="http://schemas.uof.org/cn/2009/presentation"
				xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
				xmlns:图="http://schemas.uof.org/cn/2009/graph">
  
	<!--11.11.14 李娟-->
	<xsl:import href="fill.xsl"/>
	<!--<xsl:import href="shape.xsl"/>-->
	<xsl:output encoding="UTF-8" indent="yes" method="xml" version="2.0"/>
  <xsl:template  name="table">
    <xsl:param name="type"/>
	  <图:图形_8062> 
    <!--<图:图形 uof:locID="g0000" uof:attrList="层次 标识符 组合列表 其他对象">-->
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select="translate(@id,':','')"/>
      </xsl:attribute>
      <xsl:if test =".//a:tbl">      
        <xsl:apply-templates select=".//a:tbl"/>
      </xsl:if>
	  </图:图形_8062>
    <!--</图:图形>-->
  </xsl:template>
	<!--添加文本属性 李娟 11.12.19-->

<xsl:template match="a:tbl">
	  <图:文本_803C 是否为文本框_8046="true" >
		  <xsl:attribute name="是否自动换行_8047">
			  <xsl:if test="./@wrap='none'">
				  <xsl:value-of select="'false'"/>
			  </xsl:if>
			  <xsl:if test="./@wrap='square' or name(./@wrap)=''">
				  <xsl:value-of select="'true'"/>
			  </xsl:if>
		  </xsl:attribute>
		  <xsl:attribute name="是否大小适应文字_8048">
			  <xsl:if test="not(name(./a:spAutoFit)) or a:noAutofit">
				  <xsl:value-of select="'false'"/>
			  </xsl:if>
			  <xsl:if test="./a:spAutoFit">
				  <xsl:value-of select="'true'"/>
			  </xsl:if>
		  </xsl:attribute>
		  <图:边距_803D>
			  <xsl:if test="./@lIns">
				  <xsl:attribute name="左_C608">
					  <xsl:value-of select="./@lIns div 12700"/>
				  </xsl:attribute>
			  </xsl:if>
			  <xsl:if test="./@rIns">
				  <xsl:attribute name="右_C60A">
					  <xsl:value-of select="./@rIns div 12700"/>
				  </xsl:attribute>
			  </xsl:if>
			  <xsl:if test="./@tIns">
				  <xsl:attribute name="上_C609">
					  <xsl:value-of select="./@tIns div 12700"/>
				  </xsl:attribute>
			  </xsl:if>
			  <xsl:if test="./@bIns">
				  <xsl:attribute name="下_C60B">
					  <xsl:value-of select="./@bIns div 12700"/>
				  </xsl:attribute>
			  </xsl:if>
		  </图:边距_803D>
		  <图:对齐_803E>
			  <xsl:attribute name="水平对齐_421D">
				  <xsl:if test="not(./@anchorCtr) or ./@anchorCtr='0'">
					  <xsl:value-of select="'left'"/>
				  </xsl:if>
				  <xsl:if test="./@anchorCtr='1'">
					  <xsl:value-of select="'center'"/>
				  </xsl:if>
			  </xsl:attribute>
			  <xsl:if test="./@anchor">
				  <xsl:attribute name="文字对齐_421E">
					  <xsl:choose>
						  <xsl:when test="./@anchor='b'">
							  <xsl:value-of select="'bottom'"/>
						  </xsl:when>
						  <xsl:when test="./@anchor='ctr' or ./@anchor='dist'">
							  <xsl:value-of select="'center'"/>
						  </xsl:when>
						  <xsl:when test="./@anchor='t'">
							  <xsl:value-of select="'top'"/>
						  </xsl:when>
						  <xsl:when test="./@anchor='just'">
							  <xsl:value-of select="'auto'"/>
						  </xsl:when>
					  </xsl:choose>
				  </xsl:attribute>
			  </xsl:if>
		  </图:对齐_803E>
		  <xsl:if test=".//a:tcPr/@vert='vert'">
			  <图:文字排列方向_8042> 
				   <!--2010.10.25 罗文甜 修改文字排列方向--> 
				  <xsl:choose>
					  <xsl:when test=".//@vert='horz' or not(./@vert)">
						  <xsl:value-of select ="'t2b-l2r-0e-0w'"/>
					  </xsl:when>
					  <xsl:when test=".//@vert='eaVert'">
						  <xsl:value-of select="'r2l-t2b-0e-90w'"/>
					  </xsl:when>
					  <xsl:when test=".//a:tcPr/@vert='vert'">
						  <xsl:value-of select="'r2l-t2b-90e-90w'"/>
					  </xsl:when>
					 
					  <xsl:when test=".//@vert='vert270' or ./@vert='wordArtVert'">
						  <xsl:value-of select="'l2r-b2t-270e-270w'"/>
					  </xsl:when>
					  <xsl:when test=".//@vert='wordArtVertRtl'">
						  <xsl:value-of select="'r2l-t2b-0e-90w'"/>
					  </xsl:when>
					  <xsl:otherwise>
						  <xsl:value-of select ="'t2b-l2r-0e-0w'"/>
					  </xsl:otherwise>
				  </xsl:choose>
			 </图:文字排列方向_8042>
		  </xsl:if>
		  <图:内容_8043>
		<字:文字表_416C>
    <!--<字:文字表 xmlns:uof="http://schemas.uof.org/cn/2003/uof" uof:locID="t0128" uof:attrList="类型">-->
      <!--liuyangyang,2014-11-06，修改表格自动生成样式颜色填充不正确 start-->
        <!--<xsl:for-each select="node()">-->
      <xsl:variable name="tableStyleId">
        <xsl:variable name="tmpID" select="./a:tblPr/a:tableStyleId"/>
        <xsl:choose>
          <xsl:when test="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tmpID]">
            <xsl:value-of select="$tmpID"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="ancestor::p:spTree/a:tblStyleLst/@def"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
    <xsl:variable name="bgClr">
        <xsl:choose>
          <xsl:when test="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]/a:tblBg/a:fillRef">
            <xsl:variable name="idx" select="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]/a:tblBg/a:fillRef/@idx"/>
            <xsl:variable name="rgb">
              <xsl:apply-templates select="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]/a:tblBg/a:fillRef/a:schemeClr"/>
            </xsl:variable>
            <xsl:variable name="sldLayoutID">
              <xsl:choose>
                <xsl:when test="ancestor::dsp:drawing/@refBy!=''">
                  <xsl:variable name="slideID" select="ancestor::dsp:drawing/@refBy"/>
                  <xsl:value-of select="//p:sld[@id=$slideID]/@sldLayoutID"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="ancestor::p:sld/@sldLayoutID"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="sldLayoutRel" select="concat($sldLayoutID,'.rels')"/>
            <xsl:variable name="sldMasterID" select="ancestor::p:sldMaster/@id"/>
            <xsl:variable name="slideMasterID">
              <xsl:choose>
                <xsl:when test="$sldLayoutID!=''">
                  <xsl:value-of select="substring-after(//rel:Relationships[@id=$sldLayoutRel]//rel:Relationship[contains(@Type,'slideMaster')]/@Target,'../slideMasters/')"/>
                </xsl:when>
                <xsl:when test="$sldMasterID">
                  <xsl:value-of select="$sldMasterID"/>
                </xsl:when>
                <xsl:otherwise>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="themeID" select="//a:theme[@refBy=$slideMasterID]/@id"/>
            <xsl:choose>
              <xsl:when test="$idx &gt; 0 or $idx &lt; 1000">
                <xsl:choose>
                  <xsl:when test="name(//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:fmtScheme/a:fillStyleLst/*[position()=$idx]) = 'a:solidFill'">
                    <xsl:apply-templates select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:fmtScheme/a:fillStyleLst/*[position()=$idx]/a:schemeClr">
                      <xsl:with-param name="phClr" select="$rgb"/>
                    </xsl:apply-templates>
                  </xsl:when>
                  <xsl:when test="name(//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:fmtScheme/a:fillStyleLst/*[position()=$idx]) = 'a:gradFill'">
                    <xsl:apply-templates select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:fmtScheme/a:fillStyleLst/*[position()=$idx]/a:gsLst/a:gs[position()=2]/a:schemeClr">
                      <xsl:with-param name="phClr" select="$rgb"/>
                    </xsl:apply-templates>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="$idx &gt; 1000">
                <xsl:choose>
                  <xsl:when test="name(//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:fmtScheme/a:bgFillStyleLst/*[position()=$idx - 1000]) = 'a:solidFill'">
                    <xsl:apply-templates select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:fmtScheme/a:bgFillStyleLst/*[position()=$idx - 1000]/a:schemeClr">
                      <xsl:with-param name="phClr" select="$rgb"/>
                    </xsl:apply-templates>
                  </xsl:when>
                  <xsl:when test="name(//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:fmtScheme/a:bgFillStyleLst/*[position()= $idx - 1000]) = 'a:gradFill'">
                    <xsl:apply-templates select="//p:presentation/a:theme[@id=$themeID]/a:themeElements/a:fmtScheme/a:bgFillStyleLst/*[position()= $idx - 1000]/a:gsLst/a:gs[position()=2]/a:schemeClr">
                      <xsl:with-param name="phClr" select="$rgb"/>
                    </xsl:apply-templates>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
            </xsl:choose>
         <!-- <xsl:choose>
              <xsl:when test="name($childFill[$idxx]) = 'a:solidFill'">
                <xsl:call-template name="calClrFromScheme">
                  <xsl:with-param name="colorAttri" select="$childFill[$idxx]/a:schemeClr/child::*"/>
                  <xsl:with-param name="colorAttriCount" select="count($childFill[$idxx]/a:schemeClr/child::*)"/>
                  <xsl:with-param name="schemeColor" select="$rgb"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:when test="name($childFill[$idxx]) = 'a:gradFill'">
                <xsl:call-template name="calClrFromScheme">
                  <xsl:with-param name="colorAttri" select="$childFill[$idxx]/a:gsLst/a:gs[position()=2]/a:schemeClr/child::*"/>
                  <xsl:with-param name="colorAttriCount" select="count($childFill[$idxx]/a:gsLst/a:gs[position()=2]/a:schemeClr/child::*)"/>
                  <xsl:with-param name="schemeColor" select="$rgb"/>
                </xsl:call-template>
              </xsl:when>
            </xsl:choose>-->
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'#ffffff'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:for-each select="./child::*">
      <!--end liuyangyang，2014-11-06-->
          <xsl:choose>
            <xsl:when test="name(.)='a:tblPr'">
              <xsl:call-template name="tblPr"/>
            </xsl:when>
            <xsl:when test="name(.)='a:tr'">
              <xsl:call-template name="tblRow">
                <xsl:with-param name="tableStyleId" select="$tableStyleId"/>
                <xsl:with-param name="totalRowCount" select="count(../a:tr)"/><!--总行数-->
                <xsl:with-param name="totalColumnCount" select="count(./a:tc)"/><!--总列数-->
                <xsl:with-param  name="rowCounter" select="position()-2"/><!--行号-->
                <xsl:with-param name="bgClr" select ="$bgClr"/>
              </xsl:call-template>
            </xsl:when>
          </xsl:choose>
        </xsl:for-each>
    <!--</字:文字表>-->
	  </字:文字表_416C>
			</图:内容_8043>
	  </图:文本_803C>
  </xsl:template>
  <xsl:template name="tblPr"> 
	  <字:文字表属性_41CC>
    <!--<字:文字表属性 uof:locID="t0129" uof:attrList="式样引用">-->
		  <!--<xsl:attribute name="式样引用_419C"></xsl:attribute>-->
      
      <!--2013-1-12 修改列宽和表行属性 liqiuling -->
      <xsl:if test="following-sibling::a:tblGrid">
		  <字:列宽集_41C1>
          <xsl:for-each select="following-sibling::a:tblGrid/a:gridCol">
			  <字:列宽_41C2>
              <xsl:value-of select="@w div 12700"/>
            </字:列宽_41C2>
          </xsl:for-each>
        </字:列宽集_41C1>
      </xsl:if>
      <!--end-->

      <!--2014-01-08, tangjiang, 添加文字表式样转换 start -->
      <xsl:if test="./a:tableStyleId">
        <xsl:variable name="tableStyleId" select="./a:tableStyleId"/>
        <xsl:variable name="tblStyle" select="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]"/>
        <xsl:if test="$tblStyle/a:wholeTbl/a:tcStyle/a:tcBdr">
          <xsl:variable name="tcBdr" select="$tblStyle/a:wholeTbl/a:tcStyle/a:tcBdr"/>
          <字:文字表边框_4227>
            <xsl:if test="$tcBdr/a:left">
              <uof:左_C613>
                <xsl:attribute name="线型_C60D">single</xsl:attribute>
                <xsl:attribute name="宽度_C60F">1.5</xsl:attribute>
                <xsl:attribute name="颜色_C611">
                  <xsl:variable name="leftBdrColor" select="$tcBdr/a:left/a:lnRef/a:schemeClr/@val"/>
                  <xsl:choose>
                    <xsl:when test="$leftBdrColor">
                      <xsl:for-each select="ancestor::p:presentation/p:sldMaster/p:clrMap/attribute::*">
                        <xsl:if test="name(.)=$leftBdrColor">
                          <xsl:value-of select="."/>
                        </xsl:if>
                      </xsl:for-each>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'#ffffff'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </uof:左_C613>
            </xsl:if>
            <xsl:if test="$tcBdr/a:top">
              <uof:上_C614>
                <xsl:attribute name="线型_C60D">single</xsl:attribute>
                <xsl:attribute name="宽度_C60F">1.5</xsl:attribute>
                  <xsl:attribute name="颜色_C611">
                <xsl:variable name="topBdrColor" select="$tcBdr/a:top/a:lnRef/a:schemeClr/@val"/>
                    <xsl:choose>
                      <xsl:when test="$topBdrColor">
                        <xsl:for-each select="ancestor::p:presentation/p:sldMaster/p:clrMap/attribute::*">
                          <xsl:if test="name(.)=$topBdrColor">
                            <xsl:value-of select="."/>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'#ffffff'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
              </uof:上_C614>
            </xsl:if>
            <xsl:if test="$tcBdr/a:right">
              <uof:右_C615>
                <xsl:attribute name="线型_C60D">single</xsl:attribute>
                <xsl:attribute name="宽度_C60F">1.5</xsl:attribute>
                  <xsl:attribute name="颜色_C611">
                <xsl:variable name="rightBdrColor" select="$tcBdr/a:right/a:lnRef/a:schemeClr/@val"/>
                    <xsl:choose>
                      <xsl:when test="$rightBdrColor">
                        <xsl:for-each select="ancestor::p:presentation/p:sldMaster/p:clrMap/attribute::*">
                          <xsl:if test="name(.)=$rightBdrColor">
                            <xsl:value-of select="."/>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'#ffffff'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
              </uof:右_C615>
            </xsl:if>
            <xsl:if test="$tcBdr/a:bottom">
              <uof:下_C616>
                <xsl:attribute name="线型_C60D">single</xsl:attribute>
                <xsl:attribute name="宽度_C60F">1.5</xsl:attribute>
                  <xsl:attribute name="颜色_C611">
                <xsl:variable name="bottomBdrColor" select="$tcBdr/a:bottom/a:lnRef/a:schemeClr/@val"/>
                    <xsl:choose>
                      <xsl:when test="$bottomBdrColor">
                        <xsl:for-each select="ancestor::p:presentation/p:sldMaster/p:clrMap/attribute::*">
                          <xsl:if test="name(.)=$bottomBdrColor">
                            <xsl:value-of select="."/>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'#ffffff'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
              </uof:下_C616>
            </xsl:if>
            <xsl:if test="$tcBdr/a:insideH">
              <uof:内部横线_C619>
                <xsl:attribute name="线型_C60D">single</xsl:attribute>
                <xsl:attribute name="宽度_C60F">1.5</xsl:attribute>
                  <xsl:attribute name="颜色_C611">
                <xsl:variable name="insideHBdrColor" select="$tcBdr/a:insideH/a:lnRef/a:schemeClr/@val"/>
                    <xsl:choose>
                      <xsl:when test="$insideHBdrColor">
                        <xsl:for-each select="ancestor::p:presentation/p:sldMaster/p:clrMap/attribute::*">
                          <xsl:if test="name(.)=$insideHBdrColor">
                            <xsl:value-of select="."/>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'#ffffff'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
              </uof:内部横线_C619>
            </xsl:if>
            <xsl:if test="$tcBdr/a:insideV">
              <uof:内部竖线_C61A>
                <xsl:attribute name="线型_C60D">single</xsl:attribute>
                <xsl:attribute name="宽度_C60F">1.5</xsl:attribute>
                  <xsl:attribute name="颜色_C611">
                <xsl:variable name="insideVBdrColor" select="$tcBdr/a:insideV/a:lnRef/a:schemeClr/@val"/>
                    <xsl:choose>
                      <xsl:when test="$insideVBdrColor">
                        <xsl:for-each select="ancestor::p:presentation/p:sldMaster/p:clrMap/attribute::*">
                          <xsl:if test="name(.)=$insideVBdrColor">
                            <xsl:value-of select="."/>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'#ffffff'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
              </uof:内部竖线_C61A>
            </xsl:if>
          </字:文字表边框_4227>
        </xsl:if>
        <xsl:if test="$tblStyle/a:wholeTbl/a:tcStyle/a:fill">
          <xsl:call-template name="tableStyle">
            <xsl:with-param name="cellStyleValue" select="$tblStyle/a:wholeTbl/a:tcStyle"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:if>
      <!--end 2014-01-08, tangjiang, 添加文字表式样转换-->
      
      <!--
      2010.3.14 黎美秀  
      OOXML中表格的对齐方式通过表格偏移位置来表示，没有专门元素表示；
        暂时不转
      <字:对齐 uof:locID="t0133">left</字:对齐>
      -->
		  <字:左缩进_41C4>0.0</字:左缩进_41C4>
		 <xsl:if test="a:solidFill or a:blipFill">
		  <字:填充_4134>
          <xsl:apply-templates select="a:solidFill | a:blipFill"/>
        </字:填充_4134>
      </xsl:if>
		  <字:是否自动调整大小_41C8>true</字:是否自动调整大小_41C8>
      <!--<字:自动调整大小 uof:locID="t0140" uof:attrList="值" 字:值="true"/>-->
      <!--
      Office 2007中默认单元格边距的值      
      -->
		  <字:默认单元格边距_41CA>
			  <xsl:attribute name="左_C608">0.25</xsl:attribute>
			  <xsl:attribute name="上_C609">0.13</xsl:attribute>
			  <xsl:attribute name="下_C60B">0.13</xsl:attribute>
			  <xsl:attribute name="右_C60A">0.25</xsl:attribute>
		  </字:默认单元格边距_41CA>
      <!--<字:默认单元格边距 uof:locID="t0141" uof:attrList="上 左 右 下" 字:上="0.13" 字:左="0.25" 字:右="0.25" 字:下="0.13"/>-->
    <!--</字:文字表属性>-->
	  </字:文字表属性_41CC>
  </xsl:template>

  <!--2014-01-13, tangjiang,添加Strict OOXML到UOF文字表式样中填充的转换 start -->
  <xsl:template name="tblRow">
    <xsl:param name="tableStyleId" select="''"/>
    <xsl:param name="totalRowCount" select="1"/>
    <xsl:param name="totalColumnCount" select="1"/>
    <xsl:param name="rowCounter" select="1"/>
    <xsl:param name="bgClr"/>
	  <字:行_41CD>
      <xsl:call-template name="rowPr"/>
      <xsl:for-each select="a:tc">
        <xsl:call-template name="tblcell">
          <xsl:with-param name="tableStyleId" select="$tableStyleId"/>
          <xsl:with-param name="totalRowCount" select="$totalRowCount"/>
          <xsl:with-param name="totalColumnCount" select="$totalColumnCount"/>
          <xsl:with-param name="rowCounter" select="$rowCounter"/>
          <xsl:with-param name="cellCounter" select="position()"/>
          <xsl:with-param name="bgClr" select="$bgClr"/>
        </xsl:call-template>
      </xsl:for-each>
    </字:行_41CD>
  </xsl:template>
  <!-- end 2014-01-13, tangjiang,添加Strict OOXML到UOF文字表式样中填充的转换-->
  
  <xsl:template name="rowPr">
	  <字:表行属性_41BD>
      <!--    
       此处对应关系还需要计算
    -->
		  <字:高度_41B8>
        <xsl:attribute name="最小值_41BA">
          <xsl:value-of select="@h div 12700"/>
        </xsl:attribute>
			  <!--<xsl:attribute name="固定值_41B9">
				  <xsl:value-of select="@h div 20"/>
			  </xsl:attribute>-->
      </字:高度_41B8>
      <!--
         还未找到对应
        <字:跨页 uof:locID="t0146" uof:attrList="值" 字:值="true"/>
        -->
    </字:表行属性_41BD>
  </xsl:template>


  <!--2014-01-13, tangjiang,添加Strict OOXML到UOF文字表式样中填充的转换 start -->
  <xsl:template name="tblcell">
    <xsl:param name="tableStyleId" select="''"/>
    <xsl:param name="totalRowCount" select="1"/>
    <xsl:param name="totalColumnCount" select="1"/>
    <xsl:param name="rowCounter" select="1"/>
    <xsl:param name="cellCounter" select="1"/>
    <xsl:param name="bgClr"/>
	  <!--为实现纵向单元格合并 修改 李娟 2012 04.11-->
	  <xsl:if test="not(@vMerge or @hMerge)">
		<字:单元格_41BE>
        <xsl:apply-templates select="a:tcPr">
          <xsl:with-param name="tableStyleId" select="$tableStyleId"/>
          <xsl:with-param name="totalRowCount" select="$totalRowCount"/>
          <xsl:with-param name="totalColumnCount" select="$totalColumnCount"/>
          <xsl:with-param name="rowCounter" select="$rowCounter"/>
          <xsl:with-param name="cellCounter" select="$cellCounter"/>
          <xsl:with-param name="position" select="position()"/>
          <xsl:with-param name="bgClr" select="$bgClr"/>
        </xsl:apply-templates>
        <xsl:apply-templates select="a:txBody"/>
      </字:单元格_41BE>
</xsl:if>
  </xsl:template>
  <!-- end 2014-01-13, tangjiang,添加Strict OOXML到UOF文字表式样中填充的转换-->
  
  <!--
   调用递归算法求出合并列之后的最终单元格宽度
    -->
  <xsl:template name="tcwidth">
    <xsl:param name="gridSpan"/>
    <xsl:param name="position"/>
	  <xsl:param name="rowSpan"/>
	  
    <xsl:choose>
      <xsl:when test="$gridSpan = 1">
        <xsl:value-of select="ancestor::a:tbl/a:tblGrid/a:gridCol[$position]/@w"/>
      </xsl:when>
      <xsl:when test="$gridSpan > 1">
        <xsl:variable name="temp">
          <xsl:call-template name="tcwidth">
            <xsl:with-param name="gridSpan" select="$gridSpan -1"/>
            <xsl:with-param name="position" select="$position"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:value-of select="number($temp + ancestor::a:tbl/a:tblGrid/a:gridCol[number($position + $gridSpan)-1]/@w) div 20"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>


  <!--2014-01-13, tangjiang,添加Strict OOXML到UOF文字表式样中填充的转换 start -->
  <xsl:template match="a:tcPr">
    <xsl:param name="tableStyleId" select="''"/>
    <xsl:param name="totalRowCount" select="1"/>
    <xsl:param name="totalColumnCount" select="1"/>
    <xsl:param name="rowCounter" select="1"/>
    <xsl:param name="cellCounter" select="1"/>
    <xsl:param name="position"/>
    <xsl:param name="bgClr"/>
    <!--<xsl:comment>
      <xsl:value-of select="$tableStyleId"/>
    </xsl:comment>
    <xsl:comment>
      <xsl:value-of select="$totalRowCount"/>
    </xsl:comment>
    <xsl:comment>
      <xsl:value-of select="$totalColumnCount"/>
    </xsl:comment>
    <xsl:comment>
      <xsl:value-of select="$rowCounter"/>
    </xsl:comment>
    <xsl:comment>
      <xsl:value-of select="$cellCounter"/>
    </xsl:comment>-->
    <!-- end 2014-01-13, tangjiang,添加Strict OOXML到UOF文字表式样中填充的转换-->
    
    <字:单元格属性_41B7>
      <!--
      通过列宽集计算单元格宽度
      绝对值还是相对值还没有判断
      -->
		  <!--相对值没有添加 李娟11.11.14-->
		  <字:宽度_41A1>
        <xsl:attribute name="绝对值_41A2">
          <!--
          黎美秀 不能简单计算第几个值，若是存在跨列的情况，则需要几个a:gridCol相加          
              -->
          <xsl:choose>
            <xsl:when test="not(../@gridSpan)">
              <xsl:value-of select="ancestor::a:tbl/a:tblGrid/a:gridCol[$position]/@w div 12700"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="tcwidth">
                <xsl:with-param name="gridSpan" select="../@gridSpan"/>
                <xsl:with-param name="position" select="$position"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </字:宽度_41A1>

      <!--2014-01-13, tangjiang,添加Strict OOXML到UOF文字表式样中填充的转换 start -->
      <!--2014-11-25, liuyangyang,修改表格转换样式转换 start -->
      <xsl:choose>
        <xsl:when test="$rowCounter='1'">
          <xsl:variable name="firstRowStyle" select="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]//a:firstRow/a:tcStyle"/>
          <xsl:call-template name="tableStyle">
            <xsl:with-param name="cellStyleValue" select="$firstRowStyle"/>
            <xsl:with-param name="bgClr" select="$bgClr"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:when test="not($rowCounter mod 2 = 0)">
            <xsl:choose>
              <xsl:when test="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]//a:band2H/a:tcStyle/a:fill">
                <xsl:call-template name="tableStyle">
                <xsl:with-param name="cellStyleValue" select="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]//a:band2H/a:tcStyle"/>
                  <xsl:with-param name="bgClr" select="$bgClr"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <xsl:call-template name="tableStyle">
                  <xsl:with-param name="cellStyleValue" select="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]//a:wholeTbl/a:tcStyle"/>
                  <xsl:with-param name="bgClr" select="$bgClr"/>
                </xsl:call-template>
                  </xsl:otherwise>
            </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]//a:band1H/a:tcStyle/a:fill">
              <xsl:call-template name="tableStyle">
                <xsl:with-param name="cellStyleValue" select="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]//a:band1H/a:tcStyle"/>
                <xsl:with-param name="bgClr" select="$bgClr"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:call-template name="tableStyle">
                <xsl:with-param name="cellStyleValue" select="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]//a:wholeTbl/a:tcStyle"/>
                <xsl:with-param name="bgClr" select="$bgClr"/>
              </xsl:call-template>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
      <!--end 2014-11-25,liuyangyang,修改表格样式转换-->
      <!-- end 2014-01-13, tangjiang,添加Strict OOXML到UOF文字表式样中填充的转换-->

      <字:单元格边距_41A4>
        <xsl:attribute name="上_C609">
          <xsl:choose>
            <xsl:when test="not(@marT)">
              <xsl:value-of select="'3.7'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="@marT div 12698"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="左_C608">
          <xsl:choose>
            <xsl:when test="not(@marL)">
              <xsl:value-of select="'7.1'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="@marL div 12698"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="右_C60A">
          <xsl:choose>
            <xsl:when test="not(@marR)">
              <xsl:value-of select="'7.1'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="@marR div 12698"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="下_C60B">
          <xsl:choose>
            <xsl:when test="not(@marB)">
              <xsl:value-of select="'7.1'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="@marB div 12698"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </字:单元格边距_41A4>
      <xsl:if test="a:lnL|a:lnR|a:lnT|a:lnB|a:lnTlToBr|a:lnBlToTr">
        <xsl:call-template name="border"/>
      </xsl:if>
      <!-- 增加两列合并 -->

      <!--<xsl:comment>
        <xsl:value-of select="name(.)"/>
        <xsl:value-of select="name(..)"/>
        <xsl:value-of select="name(../..)"/>
        <xsl:value-of select="name(../../..)"/>
        <xsl:value-of select="name(../../../..)"/>
      </xsl:comment>-->
      <xsl:if test="./a:solidFill|./a:blipFill|./a:gradFill|./a:pattFill|./a:noFill">
        <字:填充_4134>
          <xsl:choose>
            <xsl:when test="a:solidFill">
              <xsl:apply-templates select="a:solidFill"/>
            </xsl:when>
            <xsl:when test="a:blipFill">
              <xsl:apply-templates select="a:blipFill"/>
            </xsl:when>
            <xsl:when test="a:gradFill">
              <xsl:apply-templates select="a:gradFill"/>
            </xsl:when>
            <xsl:when test="a:pattFill">
              <xsl:apply-templates select="a:pattFill"/>
            </xsl:when>
            <xsl:when test="a:noFill">
              <图:颜色_8004>auto</图:颜色_8004>
            </xsl:when>
          </xsl:choose>
        </字:填充_4134>
      </xsl:if>
      
      <xsl:if test="../@gridSpan">
		  <字:跨列_41A7>
          <!--<xsl:attribute name="字:值">-->
            <xsl:value-of select="../@gridSpan"/>
          <!--</xsl:attribute>-->
			  </字:跨列_41A7>
      </xsl:if>
		  <xsl:if test="../@rowSpan">
			  <字:跨行_41A6>
			  
				  <xsl:value-of select="../@rowSpan"/>
			 
			  </字:跨行_41A6>
		  </xsl:if>
		  <!--字:跨行_41A6-->
		  <字:垂直对齐方式_41A5>
        <xsl:choose>
          <xsl:when test="not(@anchor)">
            <xsl:value-of select="'top'"/>
          </xsl:when>
          <xsl:when test="@anchor='ctr'">
            <xsl:value-of select="'center'"/>
          </xsl:when>
          <xsl:when test="@anchor='b'">
            <xsl:value-of select="'bottom'"/>
          </xsl:when>
        </xsl:choose>
      </字:垂直对齐方式_41A5>
		  <字:是否自动换行_41A8>false</字:是否自动换行_41A8>
		  <字:是否适应文字_41A9>false</字:是否适应文字_41A9>
		  <xsl:if test="@vert">
			  <字:文字排列方向_41AA>
				  <xsl:choose>
					  <xsl:when test="@vert='wordArtVertRtl'">
						  <xsl:value-of select="'r2l-t2b-0e-90w'"/>
					  </xsl:when>
					  <xsl:when test="@vert='horz'">
						  <xsl:value-of select ="'t2b-l2r-0e-0w'"/>
					  </xsl:when>
					  <xsl:when test="@vert='eaVert'">
						  <xsl:value-of select="'r2l-t2b-0e-90w'"/>
					  </xsl:when>
					  <xsl:when test="@vert='vert'">
						  <xsl:value-of select="'r2l-t2b-90e-90w'"/>
					  </xsl:when>

					  <xsl:when test="@vert='vert270' or @vert='wordArtVert'">
						  <xsl:value-of select="'l2r-b2t-270e-270w'"/>
					  </xsl:when>
					  <xsl:otherwise>
						  <xsl:value-of select ="'t2b-l2r-0e-0w'"/>
					  </xsl:otherwise>
				  </xsl:choose>
				  
			  </字:文字排列方向_41AA>
		  </xsl:if>
      <!--<字:适应文字 uof:locID="t0158" uof:attrList="值" 字:值="false"/>-->
    </字:单元格属性_41B7>
  </xsl:template>
	<xsl:template match="a:gradFill" mode="tabFill">
		
	</xsl:template>
	
  <xsl:template match="a:txBody">
    <xsl:apply-templates select="a:p"/>
  </xsl:template>
  <xsl:template name="border">
    <!--
    ooxml中cap属性没有对应    
    -->
	  <字:边框_4133>
      
      <xsl:if test="a:lnL">
        <xsl:for-each select="a:lnL">
			<uof:左_C613>
          <!--<uof:左 uof:locID="u0057" uof:attrList="类型 宽度 边距 颜色 阴影">-->
            <xsl:call-template name="borderline">
              <xsl:with-param name="CompoundLine" select="@cmpd"/>
              <xsl:with-param name="Preset-Dash" select="a:prstDash/@val"/>
            </xsl:call-template>
          <!--</uof:左>-->
			</uof:左_C613>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="a:lnT">
        <xsl:for-each select="a:lnT">
			<uof:上_C614>
          <!--<uof:上 uof:locID="u0058" uof:attrList="类型 宽度 边距 颜色 阴影">-->
            <xsl:call-template name="borderline">
              <xsl:with-param name="CompoundLine" select="@cmpd"/>
              <xsl:with-param name="Preset-Dash" select="a:prstDash/@val"/>
            </xsl:call-template>
          <!--</uof:上>-->
			</uof:上_C614>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="a:lnR">
        <xsl:for-each select="a:lnR">
			<uof:右_C615>
          <!--<uof:右 uof:locID="u0059" uof:attrList="类型 宽度 边距 颜色 阴影">-->
            <xsl:call-template name="borderline">
              <xsl:with-param name="CompoundLine" select="@cmpd"/>
              <xsl:with-param name="Preset-Dash" select="a:prstDash/@val"/>
            </xsl:call-template>
          <!--</uof:右>-->
			</uof:右_C615>
        </xsl:for-each>
      </xsl:if>
      <xsl:if test="a:lnB">
        <xsl:for-each select="a:lnB">
			<uof:下_C616>			
            <xsl:call-template name="borderline">
              <xsl:with-param name="CompoundLine" select="@cmpd"/>
              <xsl:with-param name="Preset-Dash" select="a:prstDash/@val"/>
            </xsl:call-template>
          </uof:下_C616>
        </xsl:for-each>
      </xsl:if>
		  <!--2012 05 10 lijuan-->
      <xsl:if test="a:lnTlToBr">
		  <xsl:choose>
			  <xsl:when test="a:lnTlToBr/a:noFill">
			  </xsl:when>
			  <xsl:when test="a:lnTlToBr/a:solidFill">
				 <xsl:for-each select="a:lnTlToBr">
			     <uof:对角线1_C617>
						<xsl:call-template name="borderline">
							<xsl:with-param name="CompoundLine" select="@cmpd"/>
							<xsl:with-param name="Preset-Dash" select="a:prstDash/@val"/>
						</xsl:call-template>
				 </uof:对角线1_C617>
				</xsl:for-each> 
			  </xsl:when>
		  </xsl:choose>
         </xsl:if>
      <xsl:if test="a:lnBlToTr">
		  <xsl:choose>
			  <xsl:when test="a:lnBlToTr/a:noFill">
				  
			  </xsl:when>
			  <xsl:when test="a:lnBlToTr/a:solidFill">
				  <xsl:for-each select="a:lnBlToTr">
			<uof:对角线2_C618>
            <xsl:call-template name="borderline">
              <xsl:with-param name="CompoundLine" select="@cmpd"/>
              <xsl:with-param name="Preset-Dash" select="a:prstDash/@val"/>
            </xsl:call-template>
          </uof:对角线2_C618>
        </xsl:for-each>
			  </xsl:when>
		  </xsl:choose>
        
      </xsl:if>
    </字:边框_4133>
  </xsl:template>
  <xsl:template name="borderline">
    <xsl:param name="CompoundLine"/>
    <xsl:param name="Preset-Dash"/>
	<xsl:if test="a:solidFill">
		<xsl:attribute name="颜色_C611">
			<xsl:apply-templates select="a:solidFill" mode="tabln"/>
		</xsl:attribute>
	</xsl:if>
    <xsl:attribute name="线型_C60D">
      <xsl:choose>
        <xsl:when test="$CompoundLine='sng'">
          <xsl:choose>
            <xsl:when test="$Preset-Dash='solid'">
              <xsl:value-of select="'single'"/>
            </xsl:when>
            <xsl:when test="$Preset-Dash='sysDash' or $Preset-Dash='dash' or a:prstDash/@val='lgDash'">
              <xsl:value-of select="'dash'"/>
            </xsl:when>
            <xsl:when test="a:prstDash/@val='sysDot'">
              <xsl:value-of select="'dotted-heavy'"/>
            </xsl:when>
            <xsl:when test="a:prstDash/@val='dot'">
              <xsl:value-of select="'dotted'"/>
            </xsl:when>
            <xsl:when test="a:prstDash/@val='sysDashDot' or a:prstDash/@val='dashDot' or a:prstDash/@val='lgDashDot'">
              <xsl:value-of select="'dot-dash'"/>
            </xsl:when>
            <xsl:when test="a:prstDash/@val='sysDashDotDot'">
              <xsl:value-of select="'dot-dot-dash'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'single'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:when test="$CompoundLine='dbl'">
          <xsl:value-of select="'double'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'single'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    
    <xsl:attribute name="宽度_C60F">
      <xsl:choose>
        <xsl:when test ="@w">
            <xsl:value-of select="@w div 12700"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select ="'1'"/>
        </xsl:otherwise>  
        
      </xsl:choose>     
    
    </xsl:attribute>
    
  </xsl:template>
	<xsl:template match="p:nvGraphicFramePr">
		<图:预定义图形_8018>
			
			<图:类别_8019>11</图:类别_8019>
			<图:名称_801A>Rectangle</图:名称_801A>
			<图:生成软件_801B>Yozo Office</图:生成软件_801B>
				
			<!--<图:属性_801D>
				
				<xsl:if test="p:cNvGraphicFramePr">
					<图:缩放是否锁定纵横比_8055>
						<xsl:choose>
							<xsl:when test="@noGrp='1'">
								<xsl:value-of select="'true'"/>
							</xsl:when>
							
						</xsl:choose>
					</图:缩放是否锁定纵横比_8055>
				</xsl:if>
			</图:属性_801D>-->
			
		</图:预定义图形_8018>
	</xsl:template>
  
  <xsl:template name="dots">
    <xsl:param name="count" select="1"/>
    <xsl:if test="$count > 0">
      <xsl:text>.</xsl:text>
      <xsl:call-template name="dots">
        <xsl:with-param name="count" select="$count - 1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

  <!--2014-01-13, tangjiang,添加Strict OOXML到UOF文字表式样中填充的转换 start -->
  <xsl:template name="tableStyle" >
    <xsl:param name="cellStyleValue" select="''"/>
    <xsl:param name="bgClr"/>
    <xsl:if test="not(./a:solidFill|./a:blipFill|./a:gradFill|./a:pattFill|./a:noFill)">
      <字:填充_4134>
        <图:颜色_8004>
        <xsl:choose>
          <xsl:when test="$cellStyleValue/a:fill/a:solidFill">
            <xsl:choose>
              <xsl:when test ="$cellStyleValue/a:fill/a:solidFill/a:schemeClr">
                <xsl:apply-templates select="$cellStyleValue/a:fill/a:solidFill/a:schemeClr">
                  <xsl:with-param name="bgClr" select="$bgClr"/>
                </xsl:apply-templates>
              </xsl:when>
              <xsl:otherwise>
                <xsl:apply-templates select="$cellStyleValue/a:fill/a:solidFill"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:when test="$cellStyleValue/a:fill/a:blipFill">
            <xsl:apply-templates select="$cellStyleValue/a:fill/a:blipFill"/>
          </xsl:when>
          <xsl:when test="$cellStyleValue/a:fill/a:gradFill">
            <xsl:apply-templates select="$cellStyleValue/a:fill/a:gradFill"/>
          </xsl:when>
          <xsl:when test="$cellStyleValue/a:fill/a:pattFill">
            <xsl:apply-templates select="$cellStyleValue/a:fill/a:pattFill"/>
          </xsl:when>
          <xsl:when test="$cellStyleValue/a:fill/a:noFill">
            
              <xsl:value-of select="$bgClr"/>
          </xsl:when>
          <!--liuyangyang 2014-11-17 添加缺省表格填充颜色 start-->
          <xsl:otherwise>
               <xsl:value-of select="$bgClr"/>
          </xsl:otherwise>
          <!--end liuyangyang 2014-11-17 添加缺省表格填充颜色-->
        </xsl:choose>
        </图:颜色_8004>
      </字:填充_4134>
    </xsl:if>
  </xsl:template>
  <!--end 2014-01-13, tangjiang,添加Strict OOXML到UOF文字表式样中填充的转换-->
  
</xsl:stylesheet>
