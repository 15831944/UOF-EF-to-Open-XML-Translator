<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
  xmlns:pzip="urn:u2o:xmlns:post-processings:special"
  xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fo="http://www.w3.org/1999/XSL/Format"
  xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
  xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
  xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
  xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:规则="http://schemas.uof.org/cn/2009/rules"
  xmlns:元="http://schemas.uof.org/cn/2009/metadata"
  xmlns:图形="http://schemas.uof.org/cn/2009/graphics"
  xmlns:图表="http://schemas.uof.org/cn/2009/chart"
  xmlns:对象="http://schemas.uof.org/cn/2009/objects"
  xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
  xmlns:式样="http://schemas.uof.org/cn/2009/styles"
  xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing">
  <xsl:import href ="pPr.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

  <!--txBody.xsl文件已修改 liyang2012-2-17-->
  <xsl:template name ="TxBody">
      <xdr:txBody>
        <xsl:call-template name ="bodyPr"/>
        <a:lstStyle/>
        <xsl:for-each select="./图:内容_8043/字:段落_416B">
          <a:p>
            <xsl:for-each select="./字:段落属性_419B">
              <a:pPr>
                <xsl:call-template name ="pPr"/>
              </a:pPr>
			  
              <!--zl 20150504-->
              <!--yanghaojie-start-->
              <!--<a:endParaRPr/>-->
              <!--yanghaojie-end-->
            </xsl:for-each>
            <xsl:for-each select="./字:句_419D">
      
                
              
              <!--<xsl:if test ="字:文本串_415B or 字:空格符_4161">-->
                <a:r>
                  <!--<xsl:apply-templates select ="./字:句属性_4158" mode ="obj"/>-->
                  <!--修改，李杨2012-4-16-->
                  
                  <!--2014-3-29, update by Qihy, 修复bug3159中文本框中文本多一空行， start-->
                  <xsl:if test="字:文本串_415B and not(字:文本串_415B/text() != '')">
                    <a:t/>
                  </xsl:if>
                  <xsl:if test="字:文本串_415B/text() != ''">
                  
                   <!--zl 20150504-->
                    <!--yanghaojie-start-解决SmartArt字体颜色丢失问题-->
                    <a:rPr lang="zh-CN" altLang="en-US">
                      <xsl:if test="字:句属性_4158/字:字体_4128/@颜色_412F">
                        <a:solidFill>
                          <a:srgbClr>
                            <xsl:attribute name="val">
                              <xsl:variable name="col">
                                <xsl:value-of select="./字:句属性_4158/字:字体_4128/@颜色_412F"/>
                              </xsl:variable>
                              <xsl:value-of select="substring-after($col,'#')"/>
                            </xsl:attribute>  
                          </a:srgbClr>
                        </a:solidFill>                         
                      </xsl:if>
                    </a:rPr> 
                    <!--yanghaojie-end-->
                    
                      <a:t>
                        <xsl:value-of select="./字:文本串_415B"/>
                      </a:t>
                   <!--zl 20150504-->
                    
                  <!--zl 20150504 把这个a:t注释掉，有问题-->  
                  <!--<a:t>
                    <xsl:variable name ="v">
                      <xsl:for-each select="./字:文本串_415B">
                        <xsl:value-of select="concat(.,'^')"/>
                      </xsl:for-each>
                    </xsl:variable>
                    <xsl:value-of select ="translate($v,'^',' ')"/>
                    --><!--<xsl:if test="字:空格符_4161">
                      <xsl:value-of select ="' '"/>
                    </xsl:if>
                    <xsl:if test ="字:文本串_415B">
                      <xsl:value-of select ="字:文本串_415B"/>
                    </xsl:if>--><!--
                  </a:t>-->
                  <!--zl 20150504-->

                  </xsl:if>
                  <!--2014-3-29 end-->
                  
                </a:r>
              
              <!--zl 20150504-->
                  
              <!-- 20130318 add by xuzhenwei BUG_2757 oo-uof 功能测试 文本框文件转换失败 
                                   gaoyuwei  Bug_2751 uof-oo 功能测试 中英文的换行丢失 中英文的换行丢失 start -->
              <xsl:if test="./字:换行符_415F">
                <xsl:for-each select="./字:换行符_415F">
                  <a:br>
                    <a:rPr lang="en-US" sz="1100"/>
                  </a:br>
                </xsl:for-each>
              </xsl:if>
              <!-- end -->
              <!--</xsl:if>-->
       
            </xsl:for-each>
            <xsl:if test ="not(./字:句_419D)">
              <a:endParaRPr/>
            </xsl:if>
            <!--what does this means?-->
            <!--
				<a:fld>
					<xsl:choose>
						<xsl:when test="ancestor::uof:锚点/@uof:占位符='date'">
							<xsl:attribute name="id">{530820CF-B880-4189-942D-D702A7CBA730}</xsl:attribute>
							<xsl:attribute name ="type">datetimeFigureOut</xsl:attribute>
						</xsl:when>
						<xsl:when test="ancestor::uof:锚点/@uof:占位符='number'">
							<xsl:attribute name="id">{0C913308-F349-4B6D-A68A-DD1791B4A57B}</xsl:attribute>
							<xsl:attribute name="type">slidenum</xsl:attribute>
						</xsl:when>
					</xsl:choose>
					<xsl:for-each select="字:段落属性">
					<a:pPr>
						<xsl:call-template name ="pPr"/>
					</a:pPr>
					</xsl:for-each>
					<xsl:for-each select="字:句">		
						<xsl:if test="字:文本串">
							<a:t>
								<xsl:value-of select ="字:文本串"/>
							</a:t>
						</xsl:if>
					</xsl:for-each>
				</a:fld>
        -->
          </a:p>
        </xsl:for-each>
        
        <xsl:if test ="not(./图:内容_8043/字:段落_416B)">
          <a:p>
            <a:pPr algn="l"/>
            <a:endParaRPr lang="zh-CN" altLang="en-US" sz="1100"/>
          </a:p>
        </xsl:if>
      </xdr:txBody>
  </xsl:template>

  <xsl:template name="bodyPr">
    <a:bodyPr>
      <xsl:if test="./图:边距_803D/@左_C608">
        <xsl:attribute name="lIns">
          <xsl:value-of select="floor(./图:边距_803D/@左_C608 * 12700)"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./图:边距_803D/@右_C60A">
        <xsl:attribute name="rIns">
          <xsl:value-of select="floor(./图:边距_803D/@右_C60A * 12700)"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./图:边距_803D/@上_C609">
        <xsl:attribute name="tIns">
          <xsl:value-of select="floor(./图:边距_803D/@上_C609 * 12700)"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./图:边距_803D/@下_C60B">
        <xsl:attribute name="bIns">
          <xsl:value-of select="floor(./图:边距_803D/@下_C60B * 12700)"/>
        </xsl:attribute>
      </xsl:if>

      <!--@anchorCtr不存在时默认值为0；uof中的center对应过来，其他指定值及未指定值转为非center-->
      <xsl:if test ="./图:对齐_803E/@水平对齐_421D='center'">
        <xsl:attribute name ="anchorCtr">1</xsl:attribute>
      </xsl:if>

      <!--UOF2.0中文字对齐的值没有justify，还有一个auto值。李杨2012-2-17-->
      <xsl:for-each select ="./图:对齐_803E/@文字对齐_421E">
        <xsl:attribute name ="anchor">
          <xsl:choose>
            <xsl:when test =".='bottom'">b</xsl:when>
          </xsl:choose>
          <xsl:choose>
            <!--<xsl:when test =".='middle'">ctr</xsl:when>-->
            <xsl:when test =".='center'">ctr</xsl:when>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test =".='top'">t</xsl:when>
          </xsl:choose>
          <xsl:choose>
            <!--<xsl:when test =".='justify'">just</xsl:when>-->
            <xsl:when test =".='base'">just</xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:for-each>

      <!--文字排列方向-->
      <!--Modified by LDM in 2011/01/18-->
        <xsl:if test="./图:文字排列方向_8042">
        <xsl:variable name="textDirection">
          <xsl:value-of select="./图:文字排列方向_8042"/>
        </xsl:variable>
        <xsl:attribute name="vert">
        <xsl:choose>
          <xsl:when test ="$textDirection = 't2b-l2r-0e-ow'">
            <xsl:value-of select="'horz'"/>
          </xsl:when>
          <xsl:when test ="$textDirection = 't2b-r2l-0e-0w'">
            <xsl:value-of select="'vert'"/>
          </xsl:when>
          <xsl:when test ="$textDirection = 'r2l-t2b-90e-90w'">
            <xsl:value-of select="'vert'"/>
          </xsl:when>
          <xsl:when test ="$textDirection = 'r2l-t2b-0e-90w'">
            <xsl:value-of select="'eaVert'"/>
          </xsl:when>
          <!-- add xuzhenwei 2012-11-21 解决UOF到OOXML垂直文本框变成横排文本框的BUG start -->
          <xsl:when test ="$textDirection = 'l2r-t2b-0e-90w'">
            <xsl:value-of select="'wordArtVert'"/>
          </xsl:when>
          <!-- end -->
          <xsl:when test ="$textDirection = 'l2r-b2t-270e-270w'">
            <xsl:value-of select="'vert270'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'horz'"/>
          </xsl:otherwise>
        </xsl:choose>
        </xsl:attribute>
      </xsl:if>

      <xsl:for-each select ="./@是否自动换行_8047">
        <xsl:attribute name ="wrap">
          <xsl:choose>
            <xsl:when test =".='false' or .='0' or .='off'">none</xsl:when>
            <xsl:otherwise>square</xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:for-each>

      <xsl:for-each select ="./@是否大小适应文字_8048">
        <xsl:choose>
          <xsl:when test =".='true' or .='1'or .='on'">
            <a:spAutoFit/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </a:bodyPr>
  </xsl:template>

  <xsl:template match ="演:段落式样">
    <xsl:call-template name ="pPr"/>
  </xsl:template>
  <xsl:template match ="uof:段落式样">
    <xsl:call-template name ="pPr"/>
  </xsl:template>
  <xsl:template match ="字:段落属性">
    <xsl:call-template name ="pPr"/>
  </xsl:template>

  <xsl:template match ="字:句属性_4158" mode="obj">
    <a:rPr>
      <xsl:call-template name="rPr"/>
    </a:rPr>
  </xsl:template>

</xsl:stylesheet>

