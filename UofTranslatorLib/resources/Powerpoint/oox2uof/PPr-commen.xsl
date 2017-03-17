<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
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
        xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
                
				xmlns:uof="http://schemas.uof.org/cn/2009/uof"
				xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
				xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
				xmlns:演="http://schemas.uof.org/cn/2009/presentation"
				xmlns:图="http://schemas.uof.org/cn/2009/graph"
				xmlns:式样="http://schemas.uof.org/cn/2009/styles">
  <xsl:import href="fill.xsl"/>
  <xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
  <xsl:template name="lstStylefont">
    <xsl:param name="lvl"/>
    <xsl:param name="phtype"/>
    <xsl:param name="hIDref"/>
    <xsl:param name="sp"/>
    <xsl:param name="idx"/>

    <xsl:variable name="lstStylefont">

      <xsl:for-each select="p:txBody/a:lstStyle">

        <xsl:choose>
          <xsl:when test="$lvl='0'">
            <xsl:value-of select="a:lvl1pPr/a:defRPr/@sz"/>
          </xsl:when>
          <xsl:when test="$lvl='1'">
            <xsl:value-of select="a:lvl2pPr/a:defRPr/@sz"/>
          </xsl:when>
          <xsl:when test="$lvl='2'">
            <xsl:value-of select="a:lvl3pPr/a:defRPr/@sz"/>
          </xsl:when>
          <xsl:when test="$lvl='3'">
            <xsl:value-of select="a:lvl4pPr/a:defRPr/@sz"/>
          </xsl:when>
          <xsl:when test="$lvl='4'">
            <xsl:value-of select="a:lvl5pPr/a:defRPr/@sz"/>
          </xsl:when>
          <xsl:when test="$lvl='5'">
            <xsl:value-of select="a:lvl6pPr/a:defRPr/@sz"/>
          </xsl:when>
          <xsl:when test="$lvl='6'">
            <xsl:value-of select="a:lvl7pPr/a:defRPr/@sz"/>
          </xsl:when>
          <xsl:when test="$lvl='7'">
            <xsl:value-of select="a:lvl8pPr/a:defRPr/@sz"/>
          </xsl:when>
          <xsl:when test="$lvl='8'">
            <xsl:value-of select="a:lvl9pPr/a:defRPr/@sz"/>
          </xsl:when>
        </xsl:choose>
      </xsl:for-each>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$lstStylefont!=''">
        <xsl:value-of select="$lstStylefont"/>
      </xsl:when>
      <xsl:otherwise>

        <xsl:choose>
          <!-- 当有继承的图形时，转到继承图形部分去找 -->
          <xsl:when test="$hIDref!=''">
            <xsl:for-each select="//p:sp[@id=$hIDref]">
              <xsl:call-template name="find_fontsize">
                <xsl:with-param name="sp" select="."/>
                <xsl:with-param name="phtype" select="p:nvSpPr/p:nvPr/p:ph/@type"/>
                <xsl:with-param name="hIDref" select="@hID"/>
                <xsl:with-param name="lvl" select="$lvl"/>
                <xsl:with-param name="idx" select="p:nvSpPr/p:nvPr/p:ph/@idx"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:when>
          <!-- 若是没有继承任何图形，则根据文本内容调用相应的默认式样（title,或者body或者others） -->
          <xsl:otherwise>
            <xsl:variable name="Mphtype">
              <xsl:choose>
                <xsl:when test="$phtype!=''">
                  <xsl:value-of select="$phtype"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="//p:sldMaster[1]/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@idx=$idx]/@type"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:choose>
              <xsl:when test="$Mphtype='title'">
                <xsl:choose>
                  <xsl:when test="$lvl='0'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='1'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl2pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='2'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl3pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='3'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl4pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='4'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl5pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='5'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl6pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='6'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl7pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='7'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl8pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='8'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl9pPr/a:defRPr/@sz"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="$Mphtype='body'">
                <xsl:choose>
                  <xsl:when test="$lvl='0'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='1'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl2pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='2'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl3pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='3'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl4pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='4'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl5pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='5'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl6pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='6'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl7pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='7'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl8pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='8'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl9pPr/a:defRPr/@sz"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:otherwise>
                <xsl:choose>
                  <xsl:when test="$lvl='0'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='1'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl2pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='2'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl3pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='3'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl4pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='4'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl5pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='5'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl6pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='6'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl7pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='7'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl8pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='8'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl9pPr/a:defRPr/@sz"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="find_fontsize">
    <xsl:param name="phtype"/>
    <xsl:param name="hIDref"/>
    <xsl:param name="lvl"/>
    <xsl:param name="sp"/>
    <xsl:param name="idx"/>
    <!-- <xsl:param name="rpr"/>
    先直接从本层sp的rpr中找，若无，试图找lststyle，
    若无，看是否有引用图形，从引用图形那进行新一轮查找，
    若无引用图形，则直接找相对应的占位符类型的字体
    -->
    <xsl:choose>
      <xsl:when test="p:txBody/a:lstStyle/node()">

        <xsl:call-template name="lstStylefont">
          <xsl:with-param name="lvl" select="$lvl"/>
          <xsl:with-param name="sp" select="$sp"/>
          <xsl:with-param name="phtype" select="$phtype"/>
          <xsl:with-param name="hIDref" select="$hIDref"/>
          <xsl:with-param name="idx" select="$idx"/>
        </xsl:call-template>
      </xsl:when>
      <!-- 当本层sp的rpr中没字体大小且前驱a:lstStyle没有节点,则试图去找继承的图形中是否有相应字体定义-->
      <xsl:otherwise>

        <xsl:choose>
          <!-- 当有继承的图形时，转到继承图形部分去找 -->
          <xsl:when test="$hIDref!=''">

            <xsl:for-each select="//p:sp[@id=$hIDref]">
              <xsl:call-template name="find_fontsize">
                <xsl:with-param name="sp" select="."/>
                <xsl:with-param name="phtype" select="p:nvSpPr/p:nvPr/p:ph/@type"/>
                <xsl:with-param name="hIDref" select="@hID"/>
                <xsl:with-param name="lvl" select="$lvl"/>
                <xsl:with-param name="idx" select="p:nvSpPr/p:nvPr/p:ph/@idx"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:when>
          <!-- 若是没有继承任何图形，则根据文本内容调用相应的默认式样（title,或者body或者others） -->
          <xsl:otherwise>

            <xsl:variable name="Mphtype">
              <xsl:choose>
                <xsl:when test="$phtype!=''">
                  <xsl:value-of select="$phtype"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="//p:sldMaster[1]/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@idx=$idx]/@type"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:choose>
              <xsl:when test="$Mphtype='title'">
                <xsl:choose>
                  <xsl:when test="$lvl='0'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='1'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl2pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='2'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl3pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='3'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl4pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='4'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl5pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='6'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl7pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='7'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl8pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='8'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl9pPr/a:defRPr/@sz"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="$Mphtype='body'">
                <xsl:choose>
                  <xsl:when test="$lvl='0'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='1'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl2pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='2'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl3pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='3'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl4pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='4'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl5pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='5'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl6pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='6'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl7pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='7'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl8pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='8'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl9pPr/a:defRPr/@sz"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:otherwise>
                <xsl:choose>
                  <xsl:when test="$lvl='0'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='1'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl2pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='2'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl3pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='3'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl4pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='4'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl5pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='5'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl6pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='6'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl7pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='7'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl8pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$lvl='8'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl9pPr/a:defRPr/@sz"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--修改文字对齐方式不正确 liqiuling 2013-4-25 start-->
  <xsl:template name="PPr-commen">
	  <!--句属性-->
    <!--对齐-->
    <xsl:if test="@algn|@fontAlgn | ancestor::p:txBody/a:lstStyle/a:lvl1pPr/@algn">
      <字:对齐_417D >
        <xsl:if test="@algn|ancestor::p:txBody/a:lstStyle/a:lvl1pPr/@algn">
          <xsl:attribute name="水平对齐_421D">
            <xsl:choose>
              <xsl:when test="@algn">
                <xsl:choose>
                  <xsl:when test="@algn='l'">left</xsl:when>
                  <xsl:when test="@algn='r'">right</xsl:when>
                  <xsl:when test="@algn='ctr'">center</xsl:when>
                  <xsl:when test="@algn='just' or @algn='justLow'">justified</xsl:when>
                  <xsl:when test="@algn='dist'or @algn='thaiDist'">distributed</xsl:when>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="ancestor::p:txBody/a:lstStyle/a:lvl1pPr/@algn">
                  <xsl:choose>
                    <xsl:when test="ancestor::p:txBody/a:lstStyle/a:lvl1pPr/@algn = 'l'">left</xsl:when>
                    <xsl:when test="ancestor::p:txBody/a:lstStyle/a:lvl1pPr/@algn = 'r'">right</xsl:when>
                    <xsl:when test="ancestor::p:txBody/a:lstStyle/a:lvl1pPr/@algn = 'ctr'">center</xsl:when>
                    <xsl:when test="ancestor::p:txBody/a:lstStyle/a:lvl1pPr/@algn='just' or ancestor::p:txBody/a:lstStyle/a:lvl1pPr/@algn='justLow'">justified</xsl:when>
                    <xsl:when test="ancestor::p:txBody/a:lstStyle/a:lvl1pPr/@algn='dist'or ancestor::p:txBody/a:lstStyle/a:lvl1pPr/@algn='thaiDist'">distributed</xsl:when>
                  </xsl:choose>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
        <!--修改文字对齐方式不正确 liqiuling 2013-4-25 start-->
      <xsl:if test="@fontAlgn">
           <xsl:attribute name="文字对齐_421E">
             <xsl:choose>
              <xsl:when test="@fontAlgn='auto'">auto</xsl:when>
              <xsl:when test="@fontAlgn='t'">top</xsl:when>
              <xsl:when test="@fontAlgn='ctr'">center</xsl:when>
              <xsl:when test="@fontAlgn='base'">base</xsl:when>
              <xsl:when test="@fontAlgn='b'">bottom</xsl:when>
        
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
        <!--luowentian-->
        <!--<xsl:if test="not(@fontAlgn)">
          <xsl:attribute name="水平对齐_421D">auto</xsl:attribute>
          </xsl:if>-->
      </字:对齐_417D>
    </xsl:if>
    <!--缩进-->
    <xsl:if test="@indent|@marL|@marR">
      <字:缩进_411D>
       <xsl:if test="@marL">
          <字:左_410E>
            <字:绝对_4107 >
              <xsl:attribute name="值_410F">
                <!--金山中-->
                <xsl:value-of select="@marL div 12700"/>
                <!--<xsl:choose>
                  <xsl:when test="number(@marL) > number(@indent)">
                    <xsl:value-of select="(@marL+@indent) div 12700"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="@marL div 12700"/>
                  </xsl:otherwise>
                </xsl:choose>-->
              </xsl:attribute>
            </字:绝对_4107>
			  <!--字:相对_4109-->
          </字:左_410E>
        </xsl:if>
        <xsl:if test="@marR">
          <字:右_4110>
			  <字:绝对_4107>
              <xsl:attribute name="值_410F">
                <xsl:value-of select="@marR div 12700"/>
              </xsl:attribute>
            </字:绝对_4107>
          
	  </字:右_4110>
        </xsl:if>
        <xsl:if test="@indent">
          <字:首行_4111>
            <字:绝对_4107 >
              <xsl:attribute name="值_410F">
                <xsl:value-of select="@indent div 12700"/>
              </xsl:attribute>
            </字:绝对_4107>
          </字:首行_4111>
        </xsl:if>
      </字:缩进_411D>
      
    </xsl:if>
    <!--行距-->
    
    <xsl:choose>
      <xsl:when test="ancestor::p:txBody/a:bodyPr/a:normAutofit/@lnSpcReduction">
        <xsl:variable name="scalelnspc" select="ancestor::p:txBody/a:bodyPr/a:normAutofit/@lnSpcReduction"/>
        <xsl:variable name="lnSpc">
          <xsl:choose>
            <xsl:when test="a:lnSpc">
              <xsl:if test="a:lnSpc/a:spcPct">
                <xsl:value-of select="a:lnSpc/a:spcPct/@val"/>
              </xsl:if>
              <xsl:if test="a:lnSpc/a:spcPts">
                <xsl:value-of select="a:lnSpc/a:spcPts/@val * 1000"/>
              </xsl:if>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="number(100000)"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="a:lnSpc/a:spcPct">
            <字:行距_417E>
              <xsl:attribute name="类型_417F">multi-lines</xsl:attribute>
              <xsl:attribute name="值_4108">
                <xsl:value-of select="($lnSpc -$scalelnspc) div 100000"/>
              </xsl:attribute>
            </字:行距_417E>
          </xsl:when>
          <xsl:otherwise>
            <字:行距_417E>
              <xsl:attribute name="类型_417F">multi-lines</xsl:attribute>
              <xsl:attribute name="值_4108">
                <xsl:value-of select="($lnSpc -$scalelnspc) div 100000"/>
              </xsl:attribute>
            </字:行距_417E>
          </xsl:otherwise>
        </xsl:choose>

      </xsl:when>
      <xsl:when test="a:lnSpc">
        <字:行距_417E>
            <xsl:if test="a:lnSpc/a:spcPct">
              <xsl:attribute name="类型_417F">multi-lines</xsl:attribute>
              <xsl:attribute name="值_4108">
                <xsl:value-of select="a:lnSpc/a:spcPct/@val div 100000"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:lnSpc/a:spcPts">
              <xsl:attribute name="类型_417F">fixed</xsl:attribute>
              <xsl:attribute name="值_4108">
                <xsl:value-of select="a:lnSpc/a:spcPts/@val div 100"/>
              </xsl:attribute>
            </xsl:if>
            <!--最小值、行间距 无对应-->
          </字:行距_417E>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx &lt;14">
          <字:行距_417E>
            <xsl:attribute name="类型_417F">multi-lines</xsl:attribute>
            <xsl:attribute name="值_4108">
              <xsl:value-of select="'0.9'"/>
            </xsl:attribute>
          </字:行距_417E>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <!--段间距-->
    <xsl:if test="a:spcBef|a:spcAft">
      <字:段间距_4180>
        <xsl:if test="a:spcBef">
          <字:段前距_4181>
			  <!--字:自动_4182-->
            <xsl:if test="a:spcBef/a:spcPct">
				<字:相对值_4184>
                <!--<xsl:attribute name="字:值">-->
                  <xsl:value-of select="a:spcBef/a:spcPct/@val div 100000"/>
                <!--</xsl:attribute>-->
              </字:相对值_4184>
            </xsl:if>
            <xsl:if test="a:spcBef/a:spcPts">
				<字:绝对值_4183>
                <!--<xsl:attribute name="字:值">-->
                  <xsl:value-of select="a:spcBef/a:spcPts/@val div 100"/>
                <!--</xsl:attribute>-->
              </字:绝对值_4183>
            </xsl:if>
            <!--自动 无对应-->
          </字:段前距_4181>
        </xsl:if>
        <xsl:if test="a:spcAft">
			<字:段后距_4185>
            <xsl:if test="a:spcAft/a:spcPct">
				<字:相对值_4184>
                <!--<xsl:attribute name="字:值">-->
                  <xsl:value-of select="a:spcAft/a:spcPct/@val div 100000"/>
                <!--</xsl:attribute>-->
              </字:相对值_4184>
            </xsl:if>
            <xsl:if test="a:spcAft/a:spcPts">
				<字:绝对值_4183>
                <!--<xsl:attribute name="字:值">-->
                  <xsl:value-of select="a:spcAft/a:spcPts/@val div 100"/>
                <!--</xsl:attribute>-->
              </字:绝对值_4183>
            </xsl:if>
            <!--自动 无对应-->
          </字:段后距_4185>
        </xsl:if>
      </字:段间距_4180>
    </xsl:if>
    <!--编号引用-->
	  <xsl:choose>
		  <xsl:when test="a:buChar|a:buAutoNum|a:buBlip">
		<字:自动编号信息_4186>
				<!--<xsl:attribute name="标识符_4100"></xsl:attribute>-->
				<!--<xsl:attribute name="名称_4122"></xsl:attribute>-->
				<!--这两个属性都有待扩充李娟2011-11-08-->
        <xsl:attribute name="编号引用_4187">
          <xsl:if test="@id">
            <xsl:value-of select="concat('bn',@id)"/>
          </xsl:if>
          <!--p/pPr/bu...的情况 -->
          <!-- 修复bug：图形内项目编号转换不正确  liqiuling 2013-03-05 tangjiang 2013-4-19 start-->
          <xsl:if test="not(@id)">
            <!--<xsl:value-of select="concat('bn',translate(ancestor::p:txBody/a:p[a:pPr/a:buAutoNum]/@id|../@id,':',''))"/>-->
            <xsl:variable name="buAutoNumType" select="./a:buAutoNum/@type"/>
            <xsl:value-of select="concat('bn',translate(ancestor::p:txBody/a:p[a:pPr/a:buAutoNum/@type = $buAutoNumType][1]/@id|../@id,':',''))"/>
          </xsl:if>
        </xsl:attribute>
      <!-- 修复bug：图形内项目编号转换不正确  liqiuling 2013-03-05 end-->
			<xsl:attribute name="编号级别_4188">1</xsl:attribute>
				<!--<字:级别_4112>
					<xsl:attribute name="级别值_4121">1</xsl:attribute>
					
				</字:级别_4112>-->
				<!--此部分需调用字处理部分，李娟2011-11-08-->
			
      </字:自动编号信息_4186>
    </xsl:when>
		  <xsl:when test="a:buNone">
			  <字:自动编号信息_4186>
				  <xsl:attribute name="编号引用_4187">bn0</xsl:attribute>
				  <xsl:attribute name="编号级别_4188">0</xsl:attribute>
			  </字:自动编号信息_4186>
		  </xsl:when>
     
		  <xsl:otherwise>
        <xsl:variable name="type">
          <xsl:choose>
            <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph">
              <xsl:value-of select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'0'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="idx">
          <xsl:choose>
            <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx">
              <xsl:value-of select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'0'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
			  <xsl:choose>
				  <xsl:when test="$type='body' and $idx &lt; 10">
					  <xsl:variable name="para">
						<xsl:choose>
						  <xsl:when test="@lvl">
							  <xsl:value-of select="./@lvl"/>
						  </xsl:when>
						  <xsl:when test="not(@lvl)">
							  <xsl:value-of select="'0'"/>
						</xsl:when>
					  </xsl:choose>
					  </xsl:variable>
            <!-- 母板符号  李秋玲 2012.11.16 -->
            <字:自动编号信息_4186>
              <xsl:variable name="sldMasterID" select="ancestor::p:sldMaster/@id"/>
              <xsl:choose>
                <xsl:when test="$para='0' and ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl1pPr/a:buChar/@char">
                  <xsl:variable name="sldLayoutId" select="ancestor::p:sld/@sldLayoutID"/>
                  <xsl:variable name="sldLayout" select="//p:sldLayout[@id=$sldLayoutId]"/>
                  <xsl:variable name="sp" select="$sldLayout//p:sp[./p:nvSpPr/p:nvPr/p:ph/@idx=$idx]"/>
                  <xsl:variable name="buNone" select="$sp/p:txBody/a:lstStyle/a:lvl1pPr/a:buNone"/>
                  <xsl:if test="not($buNone)">
                    <xsl:attribute name="编号引用_4187">
                      <xsl:choose>
                        <xsl:when test="ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl1pPr/@id">
                          <xsl:value-of select="concat('bn',ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl1pPr/@id)"/>
                        </xsl:when>
                        <xsl:when test="not(ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl1pPr/@id)">

                        </xsl:when>
                      </xsl:choose>
                    </xsl:attribute>
                    <xsl:attribute name="编号级别_4188">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:when>
                <xsl:when test="$para='1' and ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl2pPr/a:buChar/@char">
                  <xsl:variable name="sldLayoutId" select="ancestor::p:sld/@sldLayoutID"/>
                  <xsl:variable name="sldLayout" select="//p:sldLayout[@id=$sldLayoutId]"/>
                  <xsl:variable name="sp" select="$sldLayout//p:sp[./p:nvSpPr/p:nvPr/p:ph/@idx=$idx]"/>
                  <xsl:variable name="buNone" select="$sp/p:txBody/a:lstStyle/a:lvl2pPr/a:buNone"/>
                  <xsl:if test="not($buNone)">
                    <xsl:attribute name="编号引用_4187">
                      <xsl:choose>
                        <xsl:when test="ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl2pPr/@id">
                          <xsl:value-of select="concat('bn',ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl2pPr/@id)"/>
                        </xsl:when>
                        <xsl:when test="not(ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl2pPr/@id)">

                        </xsl:when>
                      </xsl:choose>
                    </xsl:attribute>
                    <xsl:attribute name="编号级别_4188">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:when>
                <xsl:when test="$para='2' and ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl3pPr/a:buChar/@char">
                  <xsl:variable name="sldLayoutId" select="ancestor::p:sld/@sldLayoutID"/>
                  <xsl:variable name="sldLayout" select="//p:sldLayout[@id=$sldLayoutId]"/>
                  <xsl:variable name="sp" select="$sldLayout//p:sp[./p:nvSpPr/p:nvPr/p:ph/@idx=$idx]"/>
                  <xsl:variable name="buNone" select="$sp/p:txBody/a:lstStyle/a:lvl3pPr/a:buNone"/>
                  <xsl:if test="not($buNone)">
                    <xsl:attribute name="编号引用_4187">
                      <xsl:choose>
                        <xsl:when test="ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl3pPr/@id">
                          <xsl:value-of select="concat('bn',ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl3pPr/@id)"/>
                        </xsl:when>
                        <xsl:when test="not(ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl3pPr/@id)">

                        </xsl:when>
                      </xsl:choose>
                    </xsl:attribute>
                    <xsl:attribute name="编号级别_4188">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:when>
                <xsl:when test="$para='3' and ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl4pPr/a:buChar/@char">
                  <xsl:variable name="sldLayoutId" select="ancestor::p:sld/@sldLayoutID"/>
                  <xsl:variable name="sldLayout" select="//p:sldLayout[@id=$sldLayoutId]"/>
                  <xsl:variable name="sp" select="$sldLayout//p:sp[./p:nvSpPr/p:nvPr/p:ph/@idx=$idx]"/>
                  <xsl:variable name="buNone" select="$sp/p:txBody/a:lstStyle/a:lvl4pPr/a:buNone"/>
                  <xsl:if test="not($buNone)">
                    <xsl:attribute name="编号引用_4187">
                      <xsl:choose>
                        <xsl:when test="ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl4pPr/@id">
                          <xsl:value-of select="concat('bn',ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl4pPr/@id)"/>
                        </xsl:when>
                        <xsl:when test="not(ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl4pPr/@id)">

                        </xsl:when>
                      </xsl:choose>
                    </xsl:attribute>
                    <xsl:attribute name="编号级别_4188">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:when>
                <xsl:when test="$para='4' and ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl5pPr/a:buChar/@char">
                  <xsl:variable name="sldLayoutId" select="ancestor::p:sld/@sldLayoutID"/>
                  <xsl:variable name="sldLayout" select="//p:sldLayout[@id=$sldLayoutId]"/>
                  <xsl:variable name="sp" select="$sldLayout//p:sp[./p:nvSpPr/p:nvPr/p:ph/@idx=$idx]"/>
                  <xsl:variable name="buNone" select="$sp/p:txBody/a:lstStyle/a:lvl5pPr/a:buNone"/>
                  <xsl:if test="not($buNone)">
                    <xsl:attribute name="编号引用_4187">
                      <xsl:choose>
                        <xsl:when test="ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl5pPr/@id">
                          <xsl:value-of select="concat('bn',ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl5pPr/@id)"/>
                        </xsl:when>
                        <xsl:when test="not(ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl5pPr/@id)">

                        </xsl:when>
                      </xsl:choose>
                    </xsl:attribute>
                    <xsl:attribute name="编号级别_4188">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:when>
                <xsl:when test="$para='5' and ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl6pPr/a:buChar/@char">
                  <xsl:variable name="sldLayoutId" select="ancestor::p:sld/@sldLayoutID"/>
                  <xsl:variable name="sldLayout" select="//p:sldLayout[@id=$sldLayoutId]"/>
                  <xsl:variable name="sp" select="$sldLayout//p:sp[./p:nvSpPr/p:nvPr/p:ph/@idx=$idx]"/>
                  <xsl:variable name="buNone" select="$sp/p:txBody/a:lstStyle/a:lvl6pPr/a:buNone"/>
                  <xsl:if test="not($buNone)">
                    <xsl:attribute name="编号引用_4187">
                      <xsl:choose>
                        <xsl:when test="ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl6pPr/@id">
                          <xsl:value-of select="concat('bn',ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl6pPr/@id)"/>
                        </xsl:when>
                        <xsl:when test="not(ancestor::p:presentation//p:sldMaster[@id=$sldMasterID]/p:txStyles/p:bodyStyle/a:lvl6pPr/@id)">

                        </xsl:when>
                      </xsl:choose>
                    </xsl:attribute>
                    <xsl:attribute name="编号级别_4188">
                      <xsl:value-of select="'1'"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:when>
              </xsl:choose>
            </字:自动编号信息_4186>
            
				  </xsl:when>
			  </xsl:choose>
		  </xsl:otherwise>
	  </xsl:choose>  
    
   
   
    <xsl:if test="a:tabLst">
		<字:制表位设置_418F>
		<xsl:for-each select="a:tabLst/a:tab">
          <字:制表位_4171>
            <xsl:attribute name="位置_4172">
              <xsl:value-of select="@pos div 12698.4"/>
            </xsl:attribute>
            <xsl:attribute name="类型_4173">
              <xsl:choose>
                <xsl:when test="@algn='l'">
                  <xsl:value-of select="'left'"/>
                </xsl:when>
                <xsl:when test="@algn='ctr'">
                  <xsl:value-of select="'center'"/>
                </xsl:when>
                <xsl:when test="@algn='r'">
                  <xsl:value-of select="'right'"/>
                </xsl:when>
                <xsl:when test="@algn='dec'">
                  <xsl:value-of select="'decimal'"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
			  <!--<xsl:attribute name="前导符_4174"></xsl:attribute>
			  <xsl:attribute name="制表位字符_4175"></xsl:attribute>
			  <xsl:attribute name="取消制表位_4245"></xsl:attribute>-->
		  </字:制表位_4171>
        </xsl:for-each>
		</字:制表位设置_418F>
    </xsl:if>  
    <!--允许单词断字，即西文在单词中间换行-->
    <xsl:if test="@latinLnBrk='1' or @latinLnBrk='true'">
		<字:是否允许单词断字_4194>true</字:是否允许单词断字_4194>
    </xsl:if>
    <xsl:if test="@latinLnBrk = '0' or @latinLnBrk='false' or not(@latinLnBrk)">
		<字:是否允许单词断字_4194>false</字:是否允许单词断字_4194>
    </xsl:if>
    <!--行首尾标点控制-->
    <xsl:if test="@hangingPunct='0' or @hangingPunct='false'">
      <!--标点在下一行显示==允许标点溢出边界为false-->
		<字:是否行首尾标点控制_4195>false</字:是否行首尾标点控制_4195>
    </xsl:if>
    <xsl:if test="@hangingPunct='1' or @hangingPunct='true' or not(@hangingPunct)">
      <!--标点在本行显示==允许标点溢出边界-->
		<字:是否行首尾标点控制_4195>true</字:是否行首尾标点控制_4195>
    </xsl:if>
    <!--中文习惯首尾字符-->
    <xsl:if test="@eaLnBrk='0' or @eaLnBrk='false'">
		<字:是否采用中文习惯首尾字符_4197>false</字:是否采用中文习惯首尾字符_4197>
    </xsl:if>
    <xsl:if test="@eaLnBrk='1' or @eaLnBrk='true' or not(@eaLnBrk)">
		<字:是否采用中文习惯首尾字符_4197>true</字:是否采用中文习惯首尾字符_4197>
    </xsl:if>    
    <!--add by linyh项目符号是图片-->
    <xsl:if test="a:buBlip">
      <xsl:apply-templates select=".//..//a:buBlip" mode="pictureText"/>
    </xsl:if>
  </xsl:template>
  <xsl:template match="a:defRPr" mode="pPr-child">
	  <xsl:param name="ID" select="p:presentation/p:sldMaster/@id"/>
    <xsl:call-template name="rPr"/>
  </xsl:template>
  <xsl:template name="rPr">
    
    <!--2014-04-16, tangjiang, 修复标题字体转换错误 start -->
    <!--2010.4.10 黎美秀修改 增加   -->
    <!--母版字体 2012.11.15 李秋玲-->
    <xsl:param name="ID" select="//p:presentation/p:sldMaster/@id"/>
    <!--<xsl:variable name="ID">
      <xsl:value-of select="ancestor::p:sldMaster/@id"/>
    </xsl:variable>-->
    <!--end 2014-04-16, tangjiang, 修复标题字体转换错误 -->
    
	  <xsl:param name="layoutID" select="ancestor::p:sld/@sldLayoutID"/>
  
		<xsl:variable name="testlvl">
			
      <xsl:choose>
        <xsl:when test="ancestor::a:p/a:pPr/@lvl">
          <xsl:value-of select="ancestor::a:p/a:pPr/@lvl"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'0'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
	 
	  <字:句属性_4158>
      <!--字体-->

      <!--2014-04-10, tangjiang, 添加默认字号 start -->
      <xsl:variable name="defaultFontSize" select="'18.0'"/>
      <!--2014-04-10, tangjiang, 添加默认字号 -->

      <xsl:choose>
        <!--修改字体 如果是默认字体 则调用母版中的字体 liqiuling 2013-3-25  start-->
		  <xsl:when test="not(a:latin or a:ea or a:cs)  and  not(ancestor::p:sldLayout) and not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='sldNum' or ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='hdr'  or ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='ftr' or ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='dt')">
			  <xsl:variable name="mytype">
				  <xsl:choose>
					  <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type">
						  <xsl:value-of select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
					  </xsl:when>
					  <xsl:when test="not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph)">
						  <xsl:value-of select="0"/>
					  </xsl:when>
				  </xsl:choose>
			  </xsl:variable>
			  <字:字体_4128>          
				  <xsl:if test="./@lang='zh-CN'">
					  
					  <xsl:attribute name="中文字体引用_412A">
						  <xsl:choose>
							  <xsl:when test="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$mytype]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:ea/@typeface">
								  <xsl:value-of select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$mytype]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:ea/@typeface"/>
							  </xsl:when>
							  <xsl:otherwise>
                  <xsl:choose>
                    <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='title'or ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='ctrTitle'">
                    <xsl:choose>
                        <xsl:when test="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:ea/@typeface='+mj-ea'">
                          <xsl:variable name="fontCIntheme">
                            <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:majorFont/a:font[@script='Hans']/@typeface"/>
                          </xsl:variable>
                          <xsl:choose>
                            <xsl:when test="//a:theme/a:themeElements/a:fontScheme/a:majorFont/a:ea/@typeface=''">
                              <xsl:value-of select="$fontCIntheme"/>
                            </xsl:when>
                            <xsl:when test="$fontCIntheme=''">
                              <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:majorFont/a:ea/@typeface"/>
                            </xsl:when>
                          </xsl:choose>
                          
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:if test="not(//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:ea/@typeface='')">
                            <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:ea/@typeface"/>
                          </xsl:if>
                        </xsl:otherwise>
                        </xsl:choose>
                    </xsl:when>

                    <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='subTitle'">
                      <xsl:choose>
                        <xsl:when test="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:ea/@typeface='+mj-ea'">
                          <xsl:variable name="fontCIntheme">
                            <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:font[@script='Hans']/@typeface"/>
                          </xsl:variable>
                          <xsl:choose>
                            <xsl:when test="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:ea/@typeface=''">
                              <xsl:value-of select="$fontCIntheme"/>
                            </xsl:when>
                            <xsl:when test="$fontCIntheme=''">
                              <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:ea/@typeface"/>
                            </xsl:when>
                          </xsl:choose>

                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:if test="not(//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:ea/@typeface='')">
                            <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl2pPr/a:defRPr/a:ea/@typeface"/>
                          </xsl:if>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>

                    <!--<xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='body'">-->
                    <xsl:otherwise>
                      <xsl:choose>
                        <xsl:when test="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:ea/@typeface='+mn-ea' 
                                              or //p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:ea/@typeface='+mj-ea'">
                          <xsl:variable name="fontCIntheme">
                            <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:font[@script='Hans']/@typeface"/>
                          </xsl:variable>
                          <xsl:choose>
                            <xsl:when test="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:ea/@typeface=''">
                              <xsl:value-of select="$fontCIntheme"/>
                            </xsl:when>
                            <xsl:when test="$fontCIntheme=''">
                              <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:ea/@typeface"/>
                            </xsl:when>
                          </xsl:choose>

                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:if test="not(//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:ea/@typeface='')">
                            <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl2pPr/a:defRPr/a:ea/@typeface"/>
                          </xsl:if>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                    <!--</xsl:when>-->

                  </xsl:choose>
							<!--  <xsl:choose>
									<xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='subTitle'  or not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph)">
										  <xsl:variable name="fontCIntheme">
											  <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:minorFont/a:font[@script='Hans']/@typeface"/>
										  </xsl:variable>
										  <xsl:choose>
											  <xsl:when test="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:minorFont/a:ea/@typeface=''">
												  <xsl:value-of select="$fontCIntheme"/>
											  </xsl:when>
											  <xsl:when test="$fontCIntheme=''">
												  <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:minorFont/a:ea/@typeface"/>
											  </xsl:when>
										  </xsl:choose>
									  </xsl:when>
									  <xsl:otherwise>
										  <xsl:variable name="fontCIntheme">
											  <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:majorFont/a:font[@script='Hans']/@typeface"/>
										  </xsl:variable>
										  <xsl:choose>
											  <xsl:when test="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:majorFont/a:ea/@typeface=''">
												  <xsl:value-of select="$fontCIntheme"/>
											  </xsl:when>
											  <xsl:when test="$fontCIntheme=''">
												  <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:majorFont/a:ea/@typeface"/>
											  </xsl:when>
										  </xsl:choose>
									  </xsl:otherwise>
								  </xsl:choose>-->
							  </xsl:otherwise>
						  </xsl:choose>
						  
						 
						
					  </xsl:attribute>
				  </xsl:if>
         
				  <xsl:if test="./@lang='en-US'">
            
            <xsl:attribute name="西文字体引用_4129">
						   <xsl:variable name="fontCIntheme">
						  <xsl:choose>
                <!--母版字体  李秋玲 2012.11.13-->
							<!-- <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='subTitle' or not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph)">
									  <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
						        </xsl:when>
							     <xsl:otherwise>
								 <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
							  </xsl:otherwise>-->
                <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='title'or ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='ctrTitle'">

                  <!--2014-04-17, tangjiang, 修复标题字体转换错误 start -->
                  <xsl:variable name="type" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
                  <xsl:variable name="sldLayoutID" select="ancestor::p:sld/@sldLayoutID"/>
                  <xsl:variable name="sldLayoutTypeface" select="//p:sldLayout[@id=$sldLayoutID]//p:sp[./p:nvSpPr/p:nvPr/p:ph/@type=$type]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:latin/@typeface"/>
                  <xsl:variable name="sldMasterTypeface" select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:latin/@typeface"/>
                  <xsl:choose>
                    <xsl:when test="$sldLayoutTypeface">
                      <xsl:choose>
                        <xsl:when test="$sldLayoutTypeface ='+mj-lt'">
                          <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
                        </xsl:when>
                        <xsl:when test="$sldLayoutTypeface ='+mn-lt'">
                          <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$sldLayoutTypeface"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:choose>
                        <xsl:when test="$sldMasterTypeface ='+mj-lt'">
                          <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
                        </xsl:when>
                        <xsl:when test="$sldMasterTypeface ='+mn-lt'">
                          <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$sldMasterTypeface"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                  </xsl:choose>
                  <!-- end 2014-04-17, tangjiang, 修复标题字体转换错误 -->

                </xsl:when>

                <!--2014-04-17, tangjiang, 修复正文字体转换错误 start -->
                <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx ='1'">
                  <xsl:variable name="idx" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx"/>
                  <xsl:variable name="sldLayoutID" select="ancestor::p:sld/@sldLayoutID"/>
                  <xsl:variable name="sldLayoutTypeface" select="//p:sldLayout[@id=$sldLayoutID]//p:sp[./p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:latin/@typeface"/>
                  <xsl:variable name="sldMasterTypeface" select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/a:latin/@typeface"/>
                  <xsl:choose>
                    <xsl:when test="$sldLayoutTypeface">
                      <xsl:choose>
                        <xsl:when test="$sldLayoutTypeface ='+mj-lt'">
                          <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
                        </xsl:when>
                        <xsl:when test="$sldLayoutTypeface ='+mn-lt'">
                          <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$sldLayoutTypeface"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:choose>
                        <xsl:when test="$sldMasterTypeface ='+mj-lt'">
                          <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
                        </xsl:when>
                        <xsl:when test="$sldMasterTypeface ='+mn-lt'">
                          <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$sldMasterTypeface"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <!-- end 2014-04-17, tangjiang, 修复正文字体转换错误 -->
                <xsl:when test="ancestor::p:sp/p:spPr/a:prstGeom/@prst='rect'
                          and /p:presentation/p:defaultTextStyle/a:lvl1pPr/a:defRPr/a:latin/@typeface!='+mj-lt'
                          and /p:presentation/p:defaultTextStyle/a:lvl1pPr/a:defRPr/a:latin/@typeface!='+mn-lt' ">
                  <xsl:value-of select="/p:presentation/p:defaultTextStyle/a:lvl1pPr/a:defRPr/a:latin/@typeface"/>
                </xsl:when>


                <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type='body'
                              and ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@sz='quarter'
                              and ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx='14'">
                  <xsl:variable name="idx" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx"/>
                  <xsl:variable name="sldLayoutID" select="ancestor::p:sld/@sldLayoutID"/>
                  <xsl:variable name="sldLayoutTypeface" select="//p:sldLayout[@id=$sldLayoutID]//p:sp[./p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:latin/@typeface"/>
                  <xsl:variable name="sldMasterTypeface" select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:latin/@typeface"/>
                  <xsl:choose>
                    <xsl:when test="$sldLayoutTypeface">
                      <xsl:choose>
                        <xsl:when test="$sldLayoutTypeface ='+mj-lt'">
                          <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
                        </xsl:when>
                        <xsl:when test="$sldLayoutTypeface ='+mn-lt'">
                          <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$sldLayoutTypeface"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:choose>
                        <xsl:when test="$sldMasterTypeface ='+mj-lt'">
                          <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
                        </xsl:when>
                        <xsl:when test="$sldMasterTypeface ='+mn-lt'">
                          <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="$sldMasterTypeface"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>


                <xsl:otherwise>
                  <xsl:value-of select="//a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
                </xsl:otherwise>

              </xsl:choose>
						   </xsl:variable>
						  <xsl:value-of select="$fontCIntheme"/>
					  </xsl:attribute>
					  
					  <!--<xsl:if test="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface!=''">

						  <xsl:variable name="fontEIntheme">
							  <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
						  </xsl:variable>
						  <xsl:attribute name="西文字体引用_4129">
							  <xsl:value-of select="$fontEIntheme"/>
						  </xsl:attribute>
					  </xsl:if>-->
				  </xsl:if>
				 
				  <xsl:if test="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:majorFont/a:cs/@typeface!=''">
					  <xsl:variable name="fontTtheme">
						  <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:majorFont/a:cs/@typeface"/>
					  </xsl:variable>
					  <xsl:attribute name="特殊字体引用_412B">
						  <xsl:value-of select="$fontTtheme"/>
					  </xsl:attribute>
				  </xsl:if>
          <!--修改字体 如果是默认字体 则调用母版中的字体 liqiuling 2013-3-25 end-->
				  <!--字体大小 李娟2012 05 26-->
				  <xsl:choose>
					  <xsl:when test="ancestor::p:txBody/a:bodyPr/a:normAutofit/@fontScale">
						  <xsl:variable name="fontScale" select="ancestor::p:txBody/a:bodyPr/a:normAutofit/@fontScale div 100000"/>
						  <xsl:variable name="fontsize">
							  <xsl:choose>
								  <xsl:when test="@sz">
									  <xsl:value-of select="@sz"/>
								  </xsl:when>
								  <xsl:otherwise>
									  <xsl:for-each select="ancestor::p:sp">
										  <xsl:call-template name="find_fontsize">
											  <xsl:with-param name="sp" select="."/>
											  <xsl:with-param name="phtype" select="p:nvSpPr/p:nvPr/p:ph/@type"/>
											  <xsl:with-param name="hIDref" select="@hID"/>
											  <xsl:with-param name="lvl" select="$testlvl"/>
											  <xsl:with-param name="idx" select="p:nvSpPr/p:nvPr/p:ph/@idx"/>
										  </xsl:call-template>
									  </xsl:for-each>
								  </xsl:otherwise>
							  </xsl:choose>
						  </xsl:variable>
              <xsl:choose>
                <xsl:when test="$fontScale != '' and $fontsize != ''">
                  <xsl:attribute name="字号_412D">
                    <xsl:value-of select="format-number(round(($fontScale * $fontsize div 100)+0.5),'0.0')"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="字号_412D">
                    <xsl:value-of select="$defaultFontSize"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
					  </xsl:when>
					  <xsl:when test="./@sz">
						   <xsl:attribute name="字号_412D">
								  <xsl:value-of select="format-number(./@sz div 100,'0.0')"/>
							  </xsl:attribute>
						</xsl:when>
					  <xsl:when test="not(./@sz) and not(ancestor::p:txBody/a:bodyPr/a:normAutofit/@fontScale) and ancestor::p:txBody/a:lstStyle/a:lvl1pPr">
						  <xsl:variable name="fontType" select="ancestor::p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz">
						  </xsl:variable>
              <xsl:choose>
                <xsl:when test="$fontType != ''">
                  <xsl:attribute name="字号_412D">
                    <xsl:value-of select="format-number($fontType div 100,'0.0')"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="字号_412D">
                    <xsl:value-of select="$defaultFontSize"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
					  </xsl:when>
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type!='' and not(ancestor::p:sldMaster or ancestor::p:notesMaster)">
                  <xsl:variable name="typeRef">
                    <xsl:choose>
                      <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type">
                        <xsl:value-of select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
                      </xsl:when>
                      <xsl:when test="not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type)">
                        <xsl:value-of select="'0'"/>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:variable>
                  <xsl:variable name="idx" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx"/>
                  <xsl:variable name="halfSize" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@sz"/>
                  <!--<xsl:choose>
										  <xsl:when test="$typeRef='ctrTitle'or $typeRef='title'">-->
                  <xsl:choose>
                    <xsl:when test="$halfSize='half' and $idx='1'">
                      <xsl:variable name="fontNum">
                        <xsl:value-of select="//p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx and p:nvSpPr/p:nvPr/p:ph/@sz=$halfSize]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz"/>
                      </xsl:variable>
                      <xsl:choose>
                        <xsl:when test="$fontNum!=''">
                          <xsl:attribute name="字号_412D">
                            <xsl:value-of select="format-number($fontNum div 100,'0.0')"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:attribute name="字号_412D">
                            <xsl:value-of select="$defaultFontSize"/>
                          </xsl:attribute>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <xsl:when test="not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph)">
                      <xsl:if test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx">
                        <xsl:variable name="type" select="//p:sldMaster/p:cSld/p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph[@idx=$idx]/@type"/>
                        <xsl:variable name="fontNum">
                          <xsl:choose>
                            <xsl:when test="$type='body'">
                              <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz"/>
                            </xsl:when>
                          </xsl:choose>
                        </xsl:variable>
                        <xsl:choose>
                          <xsl:when test="$fontNum!=''">
                            <xsl:attribute name="字号_412D">
                              <xsl:value-of select="format-number($fontNum div 100,'0.0')"/>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:attribute name="字号_412D">
                              <xsl:value-of select="$defaultFontSize"/>
                            </xsl:attribute>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:if>
                    </xsl:when>
                    <!--修改bug字号转换不正确 liqiuling 2013-03-29 2019-5-15 start  -->
                    
                    <!--2014-03-25, tangjiang, 修复备注默认字号的转换（即在备注母板中找对应的字号）-->
                    <xsl:when test="ancestor::p:notes">
                      <xsl:variable name="notesSlideId" select="ancestor::p:notes/@id"/>
                      <xsl:variable name="notesRelTarget" select="//rel:Relationships[@id=concat($notesSlideId,'.rels')]//rel:Relationship[contains(@Type,'notesMaster')]/@Target"/>
                      <xsl:variable name="notesMasterId" select="substring-after($notesRelTarget,'notesMasters/')"/>
                      <xsl:variable name="fontSize" select="//p:notesMaster[@id=$notesMasterId]//p:notesStyle/a:lvl1pPr/a:defRPr/@sz"/>
                      <xsl:choose>
                        <xsl:when test="$fontSize!=''">
                          <xsl:attribute name="字号_412D">
                            <xsl:value-of select="format-number($fontSize div 100,'0.0')"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:attribute name="字号_412D">
                            <xsl:value-of select="$defaultFontSize"/>
                          </xsl:attribute>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:when>
                    <!--end 2014-03-25, tangjiang, 修复备注默认字号的转换（即在备注母板中找对应的字号）-->
                    
                    <xsl:otherwise>
                      <xsl:variable name="lvl" select="ancestor::a:p/a:pPr/@lvl"/>
                      <xsl:variable name="layoutTextSize">
                        <xsl:choose>
                          <xsl:when test="$idx!='' and $lvl='0'">
                            <xsl:value-of  select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$idx!='' and $lvl='1'">
                            <xsl:value-of  select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$idx!='' and $lvl='2'">
                            <xsl:value-of  select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$idx!='' and $lvl='3'">
                            <xsl:value-of  select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$idx!='' and $lvl='4'">
                            <xsl:value-of  select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$idx!='' and $lvl=''">
                            <xsl:value-of  select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of  select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:variable>
                      <xsl:variable name="fontType">
                        <xsl:choose>
                          <xsl:when test="$layoutTextSize!='' and $idx='1'">
                            <xsl:value-of select="$layoutTextSize"/>
                          </xsl:when>
                          <xsl:when test="$layoutTextSize!='' and $idx!='1' and  $lvl='0'">
                            <xsl:value-of select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef and p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$layoutTextSize!='' and $idx!='1' and  $lvl='1'">
                            <xsl:value-of select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef and p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$layoutTextSize!='' and $idx!='1' and  $lvl='2'">
                            <xsl:value-of select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef and p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl3pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$layoutTextSize!='' and $idx!='1' and  $lvl='3'">
                            <xsl:value-of select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef and p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl4pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$layoutTextSize!='' and $idx!='1' and  $lvl='4'">
                            <xsl:value-of select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef and p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl5pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$layoutTextSize!='' and $idx!='1' and  $lvl=''">
                            <xsl:value-of select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef and p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$layoutTextSize!=''">
                            <xsl:value-of select="$layoutTextSize"/>
                          </xsl:when>
                          <xsl:when test="not($layoutTextSize!='')">
                            <!--<xsl:variable name="SLHid">
                              <xsl:value-of select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef and p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/@hID"/>
                            </xsl:variable>
                            <xsl:variable name="styleID">
                              <xsl:value-of select="//p:sldMaster/p:cSld/p:spTree/p:sp[@id = $SLHid]/p:txBody/a:p[a:pPr/@lvl = $lvl]/@styleID"/>
                            </xsl:variable>
                            <xsl:value-of select="//p:sldMaster/p:txStyles/*/*[@id = $styleID]/a:defRPr/@sz"/>-->
                            <xsl:choose>
                              <xsl:when test="$typeRef='title' or $typeRef='ctrTitle'">
                                <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/@sz"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz"/>
                              </xsl:otherwise>
                            </xsl:choose>
                          </xsl:when>
                        </xsl:choose>
                      </xsl:variable>
                      <xsl:choose>
                        <xsl:when test=" $fontType!='' ">
                          <xsl:attribute name="字号_412D">
                            <xsl:value-of select="format-number($fontType div 100,'0.0')"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:attribute name="字号_412D">
                            <xsl:value-of select="$defaultFontSize"/>
                          </xsl:attribute>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                    <!--修改bug字号转换不正确 liqiuling 2013-03-29 end  -->
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type!='' and (ancestor::p:sldMaster)">
                  <xsl:variable name="typeRef" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type">
                  </xsl:variable>
                  <xsl:variable name="lvl" select="ancestor::a:p/a:pPr/@lvl"/>
                  <xsl:variable name="fontType">
                    <xsl:choose>
                      <xsl:when test="$typeRef='title'">
                        <xsl:choose>
                          <xsl:when test="$lvl='0' or not(ancestor::a:p/a:pPr/@lvl)">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$lvl='1'">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl2pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$lvl='2'">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl3pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$lvl='3'">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl4pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$lvl='4'">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl5pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$lvl='5'">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl6pPr/a:defRPr/@sz"/>
                          </xsl:when>
                        </xsl:choose>
                      </xsl:when>
                      <xsl:when test="$typeRef='body'">
                        <xsl:choose>
                          <xsl:when test="$lvl='0' or $lvl=''">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$lvl='1'">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl2pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$lvl='2'">
                            <xsl:choose>
                              <xsl:when test="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl3pPr/a:defRPr/@sz">
                                <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl3pPr/a:defRPr/@sz"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl3pPr/a:defRPr/@sz"/>
                              </xsl:otherwise>
                            </xsl:choose>

                          </xsl:when>
                          <xsl:when test="$lvl='3'">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl4pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$lvl='4'">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl5pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$lvl='5'">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl6pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$lvl='6'">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl7pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$lvl='7'">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl8pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:when test="$lvl='8'">
                            <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl9pPr/a:defRPr/@sz"/>
                          </xsl:when>
                        </xsl:choose>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:variable>
                  <xsl:choose>
                    <xsl:when test=" $fontType!='' ">
                      <xsl:attribute name="字号_412D">
                        <xsl:value-of select="format-number($fontType div 100,'0.0')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:attribute name="字号_412D">
                        <xsl:value-of select="$defaultFontSize"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                
                <!--2014-05-29, tangjiang, 添加文本框字体转换,这里有很多特化处理,理论上不应这么转 start -->
                <xsl:when test="ancestor::p:sp/p:spPr/a:prstGeom/@prst='rect'">
                  <xsl:variable name="fontNum">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz"/>
                  </xsl:variable>
                  <xsl:choose>
                    <xsl:when test="$fontNum='2500'">
                      <xsl:attribute name="字号_412D">
                        <xsl:value-of select="format-number($fontNum div 100,'0.0')"/>
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:when test="$fontNum='2400'">
                      <xsl:attribute name="字号_412D">
                        <xsl:value-of select="format-number($fontNum div 100,'0.0')"/>
                      </xsl:attribute>
                    </xsl:when>

                    <!--start 2015.04.08 guoyongbin 修改图中的文字变大bug-->
                    <xsl:when test="(ancestor::p:sp/p:nvSpPr/p:cNvPr/@id='2' and ancestor::p:sp/p:spPr/a:xfrm/a:off/@x='1158618')
                                or (ancestor::p:sp/p:nvSpPr/p:cNvPr/@id='3' and  ancestor::p:sp/p:spPr/a:xfrm/a:off/@x!='1088976') ">
                      <xsl:attribute name="字号_412D">
                        <xsl:value-of select="$defaultFontSize"/>
                      </xsl:attribute>
                    </xsl:when>
                    <!--end  2015.04.08 guoyongbin 修改图中的文字变大bug-->
                    <xsl:otherwise>
                      <xsl:attribute name="字号_412D">
                        <xsl:value-of select="$defaultFontSize"/>
                      </xsl:attribute>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <!--2014-05-29, tangjiang, 添加文本框字体转换 end -->

                <!--2014-03-09, pengxin, 修改OOXML到UOF字号转换错误 start -->
                <xsl:otherwise>
                  <xsl:attribute name="字号_412D">
                    <xsl:value-of select="$defaultFontSize"/>
                  </xsl:attribute>
                </xsl:otherwise>
                <!--2014-03-09, pengxin, 修改OOXML到UOF字号转换错误 end  -->
                
              </xsl:choose>
					  </xsl:otherwise>
					</xsl:choose>

          <!--2014-04-06, tangjiang, 将字体颜色转换换成模板调用 start -->
          <!--字体颜色丽娟 2012 05 30-->
          <xsl:call-template name="fontColor">
            <xsl:with-param name="layoutID" select="$layoutID"/>
          </xsl:call-template>
          <!--end 2014-04-06, tangjiang, 将字体颜色转换换成模板调用 -->


        </字:字体_4128>
		  </xsl:when>
        <!--修改字体liqiuling 2013-4-19  start-->
          <xsl:when test ="@lang|@sz|a:latin|a:ea|a:cs|a:solidFill|ancestor::p:txBody/a:bodyPr/a:normAutofit/@fontScale|ancestor::p:txBody/a:lstStyle//@sz">
          <字:字体_4128>
			  <xsl:if test="a:latin or ./@lang='en-US'">
				  <xsl:attribute name="西文字体引用_4129">
					  <xsl:variable name="type">
						  <xsl:choose>
							  <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type">
								  <xsl:value-of select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
							  </xsl:when>
							  <xsl:when test="not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph)">
								  <xsl:value-of select="0"/>
							  </xsl:when>
						  </xsl:choose>
					  </xsl:variable>
					  <xsl:choose>
						  <xsl:when test="not(a:latin)">
							  <xsl:choose>
								  <xsl:when test="$type='subTitle' or $type='0'">
									  <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
								  </xsl:when>
								  <xsl:otherwise>
									  <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
								  </xsl:otherwise>
							  </xsl:choose>
						  </xsl:when>
              <xsl:when test="a:latin/@typeface='+mj-lt'">
                <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:majorFont/a:latin/@typeface"/>
              </xsl:when>
              <xsl:when test="a:latin/@typeface='+mn-lt'">
                <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface"/>
              </xsl:when>
              <xsl:when test="not(a:latin/@typeface)">
                <xsl:value-of select="'Calibri'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="a:latin/@typeface"/>
              </xsl:otherwise>
					  </xsl:choose>
				  </xsl:attribute>
			  </xsl:if>
			  <xsl:if test="a:ea  or ./@lang='zh-CN'">
				  <xsl:attribute name="中文字体引用_412A">
					  <xsl:variable name="type">
						  <xsl:choose>
							  <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type">
								  <xsl:value-of select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
							  </xsl:when>
							  <xsl:when test="not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph)">
								  <xsl:value-of select="0"/>
							  </xsl:when>
						  </xsl:choose>
					  </xsl:variable>
					  <xsl:choose>
						  <xsl:when test="not(a:ea)">
							  <xsl:choose>
								  <xsl:when test="$type='subTitle' or $type='0'">
									  <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:minorFont/a:font[@script='Hans']/@typeface"/>
								  </xsl:when>
								  <xsl:otherwise>
									  <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:majorFont/a:font[@script='Hans']/@typeface"/>
								  </xsl:otherwise>
							  </xsl:choose>
						  </xsl:when>
              <xsl:when test="a:ea/@typeface='+mj-ea'">
                <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:majorFont/a:ea/@typeface"/>
              </xsl:when>
              <xsl:when test="a:ea/@typeface='+mn-ea'">
                <xsl:value-of select="//a:theme[@refBy=$ID]/a:themeElements/a:fontScheme/a:minorFont/a:ea/@typeface"/>
              </xsl:when>
              <xsl:when test="not(a:ea/@typeface)">
                <xsl:value-of select="'宋体'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="a:ea/@typeface"/>
              </xsl:otherwise>
					  </xsl:choose>
				  </xsl:attribute>
          <!--修改字体liqiuling 2013-4-19  end-->
			  </xsl:if>
			  <xsl:if test="a:cs">
				  <xsl:attribute name="特殊字体引用_412B">
					  <xsl:choose>

						  <xsl:when test="a:cs/@typeface='+mn-cs' or a:cs/@typeface='+mj-cs'">
							  <xsl:value-of select="'宋体'"/>
						  </xsl:when>
						  <xsl:otherwise>
							  <xsl:value-of select="a:cs/@typeface"/>
						  </xsl:otherwise>
					  </xsl:choose>
				  </xsl:attribute>
			  </xsl:if>
			<xsl:choose>
        <xsl:when test="ancestor::p:txBody/a:bodyPr/a:normAutofit/@fontScale">
          <xsl:variable name="fontScale" select="ancestor::p:txBody/a:bodyPr/a:normAutofit/@fontScale div 100000"/>
          <xsl:variable name="fontsize">
            <xsl:choose>
              <xsl:when test="@sz">
                <xsl:value-of select="@sz"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:for-each select="ancestor::p:sp">
                  <xsl:call-template name="find_fontsize">
                    <xsl:with-param name="sp" select="."/>
                    <xsl:with-param name="phtype" select="p:nvSpPr/p:nvPr/p:ph/@type"/>
                    <xsl:with-param name="hIDref" select="@hID"/>
                    <xsl:with-param name="lvl" select="$testlvl"/>
                    <xsl:with-param name="idx" select="p:nvSpPr/p:nvPr/p:ph/@idx"/>
                  </xsl:call-template>
                </xsl:for-each>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="$fontScale!='' and $fontsize!=''">
              <xsl:attribute name="字号_412D">
                <xsl:value-of select="format-number(round($fontScale * $fontsize div 100),'0.0')"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="字号_412D">
                <xsl:value-of select="$defaultFontSize"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
				  <!--李娟2011-11-08-->
				  <!--<xsl:attribute name="是否西文绘制_412C"></xsl:attribute>-->
				  <!--<xsl:attribute name="相对字号_412E"></xsl:attribute>-->
                <!--<xsl:attribute name="字:lvl">
                  <xsl:value-of select="$testlvl"/>
                </xsl:attribute>
                <xsl:attribute name="字:测试">
                  <xsl:value-of select="$fontsize"/>
                </xsl:attribute>
                <xsl:attribute name="字:比例">
                  <xsl:value-of select="$fontScale"/>
                </xsl:attribute>-->
              </xsl:when>
        <!--修改字号转换不正确  liqiuling 2013-4-9 start-->
         <xsl:otherwise>
				  <xsl:choose>
					  <xsl:when test="@sz">
						  <xsl:attribute name="字号_412D">
							 <xsl:value-of select="format-number(@sz div 100,'0.0')"/>
						  </xsl:attribute>
					  </xsl:when>
            <xsl:when test="ancestor::p:txBody/a:lstStyle//@sz">
              <xsl:variable name="fontSize" select="ancestor::p:txBody/a:lstStyle//@sz">
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="$fontSize!=''">
                  <xsl:attribute name="字号_412D">
                    <xsl:value-of select="format-number($fontSize div 100,'0.0')"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="字号_412D">
                    <xsl:value-of select="$defaultFontSize"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <!--修改字号转换不正确  liqiuling 2013-4-9 end--> 
					  <xsl:when test="not(@sz) and (ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx or ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type or ancestor::a:p/a:pPr/@lvl)">
						  <!--<xsl:variable name="typeRef" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>-->
						  <xsl:variable name="idx" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx"/>
						  <xsl:variable name="typeRef" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
						  <xsl:variable name="lvl" select="ancestor::a:p/a:pPr/@lvl"/>
              <xsl:variable name="fontNum">
                <xsl:choose>
                  <xsl:when test="$typeRef='title' or $typeRef='ctrTitle' or $typeRef='subTitle' and ($lvl='0' or not(ancestor::a:p/a:pPr/@lvl))">
                    <xsl:choose>
                      <xsl:when test="//p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz">
                        <xsl:value-of select="//p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:choose>
                          <xsl:when test="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/@sz">
                            <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/@sz"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/@sz"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:when test="$typeRef='title' and $lvl='1'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl2pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$typeRef='title' and $lvl='2'">
                    <xsl:choose>
                      <xsl:when test="//p:sldMaster/p:txStyles/p:titleStyle/a:lvl3pPr/a:defRPr/@sz">
                        <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl3pPr/a:defRPr/@sz"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl3pPr/a:defRPr/@sz"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:when test="$typeRef='body' and $lvl='0'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$typeRef='body' and $lvl='1'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl2pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$typeRef='body' and $lvl='2'">
                    <xsl:choose>
                      <xsl:when test="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl3pPr/a:defRPr/@sz">
                        <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl3pPr/a:defRPr/@sz"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl3pPr/a:defRPr/@sz"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:when test="$typeRef='body' and $lvl='3'">
                    <xsl:choose>
                      <xsl:when test="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl4pPr/a:defRPr/@sz">
                        <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl4pPr/a:defRPr/@sz"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="//p:sldMaster/p:txStyles/p:otherStyle/a:lvl4pPr/a:defRPr/@sz"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                  <xsl:when test="$typeRef='body' and $lvl='4'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl5pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$typeRef='body' and $lvl='5'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl6pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$typeRef='body' and $lvl='6'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl7pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$typeRef='body' and $lvl='7'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl8pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$typeRef='body' and $lvl='8'">
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl9pPr/a:defRPr/@sz"/>
                  </xsl:when>
                  <xsl:when test="$typeRef='body' and  not(ancestor::a:p/a:pPr/@lvl) and $idx">
                    <xsl:variable name="layoutFontSize" select="//p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz"/>
                    <xsl:choose>
                      <xsl:when test="$layoutFontSize!=''">
                        <xsl:value-of select="$layoutFontSize"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:when>
                  <!-- 默认按body的lvl1pPr处理 added by tangjiang 2014-03-27 -->
                  <xsl:otherwise>
                    <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz"/>
                  </xsl:otherwise>
                </xsl:choose>
						  </xsl:variable>
              <xsl:choose>
                <xsl:when test="$fontNum!=''">
                  <xsl:attribute name="字号_412D">
                    <xsl:value-of select="format-number($fontNum div 100,'0.0')"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="字号_412D">
                    <xsl:value-of select="$defaultFontSize"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
					  </xsl:when>
            
            <!--2013-1-13 增加默认字号设置 李秋玲 start-->
            <!--2014-05-29, tangjiang, 添加文本框字体转换 start -->
            <xsl:when test="ancestor::p:sp/p:spPr/a:prstGeom/@prst='rect'">
              <xsl:variable name="fontNum">
                <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz"/>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="$fontNum='2500'">
                  <xsl:attribute name="字号_412D">
                    <xsl:value-of select="format-number($fontNum div 100,'0.0')"/>
                  </xsl:attribute>
                </xsl:when>
                    <xsl:when test="$fontNum='2400'">
                      <xsl:attribute name="字号_412D">
                        <xsl:value-of select="format-number($fontNum div 100,'0.0')"/>
                      </xsl:attribute>
                    </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="字号_412D">
                    <xsl:value-of select="$defaultFontSize"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <!--2014-05-29, tangjiang, 添加文本框字体转换 end -->
            
            <xsl:otherwise>
              <xsl:attribute name="字号_412D">
                <xsl:value-of select="$defaultFontSize"/>
              </xsl:attribute>
            </xsl:otherwise>
            <!--2013-1-13 增加默认字号设置 李秋玲 end-->
				  </xsl:choose>
         </xsl:otherwise>
         </xsl:choose>

            <!--2014-04-06, tangjiang, 将字体颜色转换换成模板调用 start -->
            <!--字体颜色只考虑纯色填充-->
            <!--<xsl:choose>
              <xsl:when test="a:solidFill/a:srgbClr">
                <xsl:attribute name="颜色_412F">
                  <xsl:apply-templates select="a:solidFill/a:srgbClr"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:when test="a:solidFill/a:schemeClr">
                <xsl:attribute name="颜色_412F">
                  <xsl:apply-templates select="a:solidFill/a:schemeClr"/>
                </xsl:attribute>
              </xsl:when>
              -->
            <!--2011-5-26 luowentian-->
            <!--
              <xsl:otherwise>
                <xsl:attribute name="颜色_412F">
                  <xsl:value-of select="'#000000'"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>-->

            <xsl:call-template name="fontColor">
              <xsl:with-param name="layoutID" select="$layoutID"/>
            </xsl:call-template>
            <!--end 2014-04-06, tangjiang, 将字体颜色转换换成模板调用 -->
          </字:字体_4128>
        </xsl:when>
		  <!--<xsl:when test="not(a:latin or a:ea or a:cs) and ancestor::p:sldLayout">
			  <xsl:if test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type">
				  <xsl:variable name="typeRef" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
				  <xsl:variable name="lvl" select="ancestor::a:p/a:pPr/@lvl"/>
				   <xsl:variable name="fontNum">
				  <xsl:choose>
					  <xsl:when test="$typeRef='body' and ($lvl='0' or not(ancestor::a:p/a:pPr/@lvl))">
						 
						  <xsl:choose>
							  <xsl:when test="ancestor::p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz">
								  <xsl:value-of select="ancestor::p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz"/>
							  </xsl:when>
							  <xsl:when test="not(ancestor::p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@sz)">
								  <xsl:value-of select="//p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz"/>
							  </xsl:when>
						  </xsl:choose>
						 
					  </xsl:when>
				  </xsl:choose>
				   </xsl:variable>
				  <xsl:attribute name="字号66_412D">
					  <xsl:value-of select="format-number($fontNum div 100,'0.0')"/>
				  </xsl:attribute>
			  </xsl:if>
		  </xsl:when>-->
         <xsl:when test ="not(ancestor::p:txBody//a:solidFill)">          
          <xsl:if test="ancestor::p:sp/p:style">
			   <字:字体_4128>
              <xsl:attribute name="颜色_412F">
                <xsl:value-of select="'#ffffff'"/>
              </xsl:attribute>
            </字:字体_4128>
          </xsl:if>
        </xsl:when>
		 
      </xsl:choose>
      <!--修改粗体丢失 liqiuling 2013-5-15 start-->
      <!--粗体-->
      <xsl:variable name="type">
        <xsl:choose>
          <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type">
            <xsl:value-of select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
          </xsl:when>
          <xsl:when test="not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph)">
            <xsl:value-of select="0"/>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name="idx">
        <xsl:choose>
          <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx">
            <xsl:value-of select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx"/>
          </xsl:when>
          <xsl:when test="not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph)">
            <xsl:value-of select="0"/>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="@b">
          <字:是否粗体_4130>
            <!--<xsl:attribute name="字:值">-->
            <xsl:choose>
              <xsl:when test="(@b='on')or(@b='1')or(@b='true')">
                <xsl:value-of select="'true'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'false'"/>
              </xsl:otherwise>
            </xsl:choose>
            <!--</xsl:attribute>-->
          </字:是否粗体_4130>
        </xsl:when>
        <xsl:when test="//p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@b">
          <字:是否粗体_4130>
            <!--<xsl:attribute name="字:值">-->
            <xsl:choose>
              <xsl:when test="//p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/@b='1'">
                <xsl:value-of select="'true'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'false'"/>
              </xsl:otherwise>
            </xsl:choose>
            <!--</xsl:attribute>-->
          </字:是否粗体_4130>

        </xsl:when>

      </xsl:choose>
      <!--修改粗体丢失 liqiuling 2013-5-15 end-->
      <!--斜体-->
      <xsl:if test="@i or ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx !=''">
        <字:是否斜体_4131>
          <!--<xsl:attribute name="字:值">-->
          <xsl:choose>
            <xsl:when test="(@i='on')or(@i='1')or(@i='true')">
              <xsl:value-of select="'true'"/>
            </xsl:when>

            <!--2014-04-07, tangjiang, 修复副标题斜体转换丢失 Start -->
            <xsl:when test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx !=''">
              <xsl:variable name="hID" select="ancestor::p:sp/@hID"/>
              <xsl:variable name="sldLayoutSp" select="//p:sp[@id=$hID]"/>
              <xsl:choose>
                <xsl:when test="$sldLayoutSp/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr[(@i='on')or(@i='1')or(@i='true')]">
                  <xsl:value-of select="'true'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:variable name="hID2" select="$sldLayoutSp/@hID"/>
                  <xsl:variable name="sldLayoutSp2" select="//p:sp[@id=$hID2]"/>
                  <xsl:choose>
                    <xsl:when test="$sldLayoutSp2/p:txBody/a:p[a:pPr/@lvl='0']/a:r/a:rPr[(@i='on')or(@i='1')or(@i='true')] or $sldLayoutSp2/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr[(@i='on')or(@i='1')or(@i='true')]">
                      <xsl:value-of select="'true'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:variable name="styleID" select ="$sldLayoutSp2/p:txBody/a:p[a:pPr/@lvl='0']/@styleID"/>
                      <xsl:choose>
                        <xsl:when test="//p:sldMaster/p:txStyles/*/*[@id=$styleID]/a:defRPr[(@i='on')or(@i='1')or(@i='true')]">
                          <xsl:value-of select="'true'"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:value-of select="'false'"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <!--end 2014-04-07, tangjiang, 修复副标题斜体转换丢失 -->

            <xsl:otherwise>
              <xsl:value-of select="'false'"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--</xsl:attribute>-->
        </字:是否斜体_4131>
      </xsl:if>
      <!--突出显示-->
      <xsl:if test="a:highlight">
		  <字:突出显示颜色_4132>
          <xsl:if test="a:highlight/a:srgbClr">
            <!--<xsl:attribute name="字:颜色">-->
              <xsl:apply-templates select="a:highlight/a:srgbClr"/>
            <!--</xsl:attribute>-->
          </xsl:if>
        </字:突出显示颜色_4132>
      </xsl:if>
      <!--填充-->
      <!--xsl:call-template name="solidFill"/>
			<xsl:apply-templates select="a:solidFill" mode="tx"/-->
      <!--删除线-->
      <xsl:if test="@strike">
		  <字:删除线_4135>
          <!--<xsl:attribute name="字:类型">-->
            <xsl:choose>
              <xsl:when test="@strike='sngStrike'">single</xsl:when>
              <xsl:when test="@strike='dblStrike'">double</xsl:when>
              <xsl:when test="@strike='noStrike'">none</xsl:when>
            </xsl:choose>
          <!--</xsl:attribute>-->
        </字:删除线_4135>
      </xsl:if>
      <!--2011-3-30罗文甜，修改下划线对应-->
      <!--2013-1-13  增加粗线型和波浪线的对应 liqiuling start-->
      <xsl:if test="@u">
		  <字:下划线_4136>
          <xsl:if test="@u">
            <xsl:choose>
              <!--heavy无对应 转为single类型-->
              <xsl:when test="@u='sng' or @u='heavy'">
                <xsl:attribute name="线型_4137">single</xsl:attribute>
              </xsl:when>
              <xsl:when test="@u='dbl'">
                <xsl:attribute name="线型_4137">double</xsl:attribute>
              </xsl:when>
              <xsl:when test="@u='dotted'">
                <xsl:attribute name="线型_4137">single</xsl:attribute>
                <xsl:attribute name="虚实_4138">square-dot</xsl:attribute>
              </xsl:when>
              <xsl:when test="@u='dottedHeavy'">
                <xsl:attribute name="线型_4137">thick</xsl:attribute>
                <xsl:attribute name="虚实_4138">square-dot</xsl:attribute>
              </xsl:when>
              <xsl:when test="@u='dash'">
                <xsl:attribute name="线型_4137">single</xsl:attribute>
                <xsl:attribute name="虚实_4138">dash</xsl:attribute>
              </xsl:when>
              <xsl:when test="@u='dashHeavy'">
                <xsl:attribute name="线型_4137">thick</xsl:attribute>
                <xsl:attribute name="虚实_4138">dash</xsl:attribute>
              </xsl:when>
              <xsl:when test="@u='dashLong'">
                <xsl:attribute name="线型_4137">single</xsl:attribute>
                <xsl:attribute name="虚实_4138">long-dash</xsl:attribute>
              </xsl:when>
              <xsl:when test="@u='dashLongHeavy'">
                <xsl:attribute name="线型_4137">thick</xsl:attribute>
                <xsl:attribute name="虚实_4138">long-dash</xsl:attribute>
              </xsl:when>
              <xsl:when test="@u='dotDash'">
                <xsl:attribute name="线型_4137">single</xsl:attribute>
                <xsl:attribute name="虚实_4138">dash-dot</xsl:attribute>
              </xsl:when>
              <xsl:when test="@u='dotDashHeavy'">
                <xsl:attribute name="线型_4137">thick</xsl:attribute>
                <xsl:attribute name="虚实_4138">dash-dot</xsl:attribute>
              </xsl:when>
              <xsl:when test="@u='dotDotDash'">
                <xsl:attribute name="线型_4137">single</xsl:attribute>
                <xsl:attribute name="虚实_4138">dash-dot-dot</xsl:attribute>
              </xsl:when>
              <xsl:when test="@u='dotDotDashHeavy'">
                <xsl:attribute name="线型_4137">thick</xsl:attribute>
                <xsl:attribute name="虚实_4138">dash-dot-dot</xsl:attribute>
              </xsl:when>
				<xsl:when test="@u='wavyDbl'">
					<xsl:attribute name="线型_4137">double</xsl:attribute>
          <xsl:attribute name="虚实_4138">wave</xsl:attribute>
				</xsl:when>
				<xsl:when test="@u='wavy'">
					<xsl:attribute name="线型_4137">single</xsl:attribute>
          <xsl:attribute name="虚实_4138">wave</xsl:attribute>
				</xsl:when>
              <xsl:when test="@u='wavyHeavy'">
                <xsl:attribute name="线型_4137">thick</xsl:attribute>
                <xsl:attribute name="虚实_4138">wave</xsl:attribute>
              </xsl:when>
              
              <xsl:when test="@u='wavyDbl'">
                <xsl:attribute name="线型_4137">double</xsl:attribute>
                <xsl:attribute name="虚实_4138">wave</xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:attribute name="线型_4137">
                  <xsl:value-of select="@u"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:if>
        <!--2013-1-13  增加粗线型和波浪线的对应 liqiuling end-->
          <!--oox中 表示类型的@u属性不一定出现，不出现表示类型为none-->
          <xsl:if test="not(@u)">
            <xsl:attribute name="是否字下划线_4139">false</xsl:attribute>
			  <!--原来是下划线 none 李娟 2011-11-08-->
          </xsl:if>
          <xsl:if test="a:uFill">
            <!--2011-6-24 罗文甜 增加颜色-->           
            <xsl:attribute name="颜色_412F">
              <xsl:choose>
                <xsl:when test="a:uFill/a:solidFill/a:srgbClr">
                  <xsl:apply-templates select="a:uFill/a:solidFill/a:srgbClr"/>
                </xsl:when>
                <xsl:when test="a:uFill/a:solidFill/a:schemeClr">
                  <xsl:apply-templates select="a:uFill/a:solidFill/a:schemeClr"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
          </xsl:if>
        </字:下划线_4136>
      </xsl:if>
      <!--阴影-->
      <xsl:if test="a:effectLst/a:outerShdw|a:effectLst/a:innerShdw|a:effectLst/a:prstShdw">
		  <字:是否阴影_4140>true</字:是否阴影_4140>
      </xsl:if>
      <!--醒目字体-->
      <xsl:if test="@cap">
		  <字:醒目字体类型_4141>
          <!--<xsl:attribute name="字:类型">-->
            <xsl:choose>
              <xsl:when test="@cap='none'">none</xsl:when>
              <xsl:when test="@cap='small'">small-caps</xsl:when>
              <xsl:when test="@cap='all'">uppercase</xsl:when>
            </xsl:choose>
          <!--</xsl:attribute>-->
        </字:醒目字体类型_4141>
      </xsl:if>
      <!--上下标-->
      <xsl:if test="@baseline">
		  <字:上下标类型_4143>
          <!--<xsl:attribute name="字:值">-->
            <xsl:if test="@baseline&gt;0">sup</xsl:if>
            <xsl:if test="@baseline=0">none</xsl:if>
            <xsl:if test="@baseline&lt;0">sub</xsl:if>
          <!--</xsl:attribute>-->
        </字:上下标类型_4143>
      </xsl:if>
      <!--字符间距-->
      <xsl:if test="@spc">
		  <字:字符间距_4145>
          <xsl:value-of select="@spc div 100"/>
        </字:字符间距_4145>
        
        
        
      </xsl:if>
      <!--xsl:if test="not(@spc)">
					<xsl:value-of select="'0'"/>
				</xsl:if-->
      <!--调整字间距-->
      <xsl:if test="@kern">
		  <字:调整字间距_4146>
          <xsl:value-of select="format-number(@kern div 100,'0.0')"/>
        </字:调整字间距_4146>
      </xsl:if>
      <!--xsl:if test="not(@kern)">
					<xsl:value-of select="'0'"/>
				</xsl:if-->
      <!--xsl:apply-templates select ="a:highlight" mode="defRPrChildren"/-->
    </字:句属性_4158>
  
</xsl:template>
	<!--<xsl:template name="lvlRef">
		<xsl:param name="ID" select="//p:presentation/p:sldMaster/@id"/>
		<xsl:variable name="clrlvl"/>
		<xsl:variable name="clrMap">
			<xsl:comment>999</xsl:comment>
				<xsl:choose>
					<xsl:when test="$clrlvl='tx2'">
						<xsl:value-of select="123434"/>
						--><!--<xsl:value-of select="//p:sldMaster/p:clrMap/@tx2"/>--><!--
					</xsl:when>
					<xsl:when test="$clrlvl='tx1'">
						<xsl:value-of select="//p:sldMaster/p:clrMap/@tx1"/>
					</xsl:when>
					<xsl:when test="$clrlvl='bg1'">
						<xsl:value-of select="//p:sldMaster/p:clrMap/@bg1"/>
					</xsl:when>
					<xsl:when test="$clrlvl='bg2'">
						<xsl:value-of select="//p:sldMaster/p:clrMap/@bg2"/>
					</xsl:when>
					<xsl:when test="$clrlvl='accent1'">
						<xsl:value-of select="//p:sldMaster/p:clrMap/@accent1"/>
					</xsl:when>
					<xsl:when test="$clrlvl='accent2'">
						<xsl:value-of select="//p:sldMaster/p:clrMap/@accent2"/>
					</xsl:when>
					<xsl:when test="$clrlvl='accent3'">
						<xsl:value-of select="//p:sldMaster/p:clrMap/@accent3"/>
					</xsl:when>
					<xsl:when test="$clrlvl='accent4'">
						<xsl:value-of select="//p:sldMaster/p:clrMap/@accent4"/>
					</xsl:when>
					<xsl:when test="$clrlvl='accent5'">
						<xsl:value-of select="//p:sldMaster/p:clrMap/@accent5"/>
					</xsl:when>
					<xsl:when test="$clrlvl='accent6'">
						<xsl:value-of select="//p:sldMaster/p:clrMap/@accent6"/>
					</xsl:when>
					<xsl:when test="$clrlvl='hlink'">
						<xsl:value-of select="//p:sldMaster/p:clrMap/@hlink"/>
					</xsl:when>
					<xsl:when test="$clrlvl='folHlink'">
						<xsl:value-of select="//p:sldMaster/p:clrMap/@folHlink"/>
					</xsl:when>
				</xsl:choose>
			
		</xsl:variable>
	<xsl:choose>
			<xsl:when  test="$clrMap='dk2'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:dk2/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='dk1'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:dk1/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='tx2'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:tx2/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='tx1'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:tx1/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='lt2'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:lt2/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='lt1'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:lt1/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='accent2'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:accent2/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='accent1'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='accent4'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:accent4/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='accent3'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:accent3/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='accent6'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='accent5'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:accent5/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='hlink'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:hlink/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when  test="$clrMap='folHlink'">
				<xsl:value-of select="concat('#',//a:theme[@refBy=$ID]/a:themeElements/a:clrScheme/a:folHlink/a:srgbClr/@val)"/>
			</xsl:when>
		</xsl:choose>
		
	</xsl:template>-->
  <xsl:template match="a:solidFill" mode="tx">
	  <字:填充_4134>
      <!--add by linyh-->
      <xsl:call-template name="sixFill"/>
    </字:填充_4134>
  </xsl:template>

  <!--2014-03-04, tangjiang, 修改OOXML到UOF字体颜色的转换 start -->
  <xsl:template name="clrSchemeMap">
    <xsl:param name="clrVal" select="'tx2'"/>
    <xsl:choose>
      <xsl:when test="$clrVal='tx2'">
        <xsl:value-of select="//p:sldMaster/p:clrMap/@tx2"/>
      </xsl:when>
      <xsl:when test="$clrVal='tx1'">
        <xsl:value-of select="//p:sldMaster/p:clrMap/@tx1"/>
      </xsl:when>
      <xsl:when test="$clrVal='bg1'">
        <xsl:value-of select="//p:sldMaster/p:clrMap/@bg1"/>
      </xsl:when>
      <xsl:when test="$clrVal='bg2'">
        <xsl:value-of select="//p:sldMaster/p:clrMap/@bg2"/>
      </xsl:when>
      <xsl:when test="$clrVal='accent1'">
        <xsl:value-of select="//p:sldMaster/p:clrMap/@accent1"/>
      </xsl:when>
      <xsl:when test="$clrVal='accent2'">
        <xsl:value-of select="//p:sldMaster/p:clrMap/@accent2"/>
      </xsl:when>
      <xsl:when test="$clrVal='accent3'">
        <xsl:value-of select="//p:sldMaster/p:clrMap/@accent3"/>
      </xsl:when>
      <xsl:when test="$clrVal='accent4'">
        <xsl:value-of select="//p:sldMaster/p:clrMap/@accent4"/>
      </xsl:when>
      <xsl:when test="$clrVal='accent5'">
        <xsl:value-of select="//p:sldMaster/p:clrMap/@accent5"/>
      </xsl:when>
      <xsl:when test="$clrVal='accent6'">
        <xsl:value-of select="//p:sldMaster/p:clrMap/@accent6"/>
      </xsl:when>
      <xsl:when test="$clrVal='hlink'">
        <xsl:value-of select="//p:sldMaster/p:clrMap/@hlink"/>
      </xsl:when>
      <xsl:when test="$clrVal='folHlink'">
        <xsl:value-of select="//p:sldMaster/p:clrMap/@folHlink"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!-- end 2014-03-04, tangjiang, 修改OOXML到UOF字体颜色的转换-->
  <!--liuyangyang 2015-03-06 添加模版用以获取表格所在行数 start-->
  <xsl:template match="a:tr" mode ="getRowCounter">
    <xsl:choose>
      <xsl:when test="preceding-sibling::a:tr">
        <xsl:value-of select="2"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="1"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!-- end liuyangyang 2015-03-06 添加模版用以获取表格所在行数 -->


  <!--2014-04-06, tangjiang, 由于字体颜色涉及多处相同的调用，故将字体颜色转换做成模板 start -->
  <xsl:template name="fontColor">
    <xsl:param name="layoutID"></xsl:param>
    <xsl:choose>
      <xsl:when test="a:solidFill/a:srgbClr">
        <xsl:attribute name="颜色_412F">
          <xsl:apply-templates select="a:solidFill/a:srgbClr"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="a:solidFill/a:schemeClr">
        <xsl:attribute name="颜色_412F">
          <xsl:apply-templates select="a:solidFill/a:schemeClr"/>
        </xsl:attribute>
      </xsl:when>
      <!--2014-04-20, tangjiang, 添加prstClr颜色参数控制，有待完善 start -->
      <xsl:when test="a:solidFill/a:prstClr">
        <xsl:attribute name="颜色_412F">
          <xsl:variable name="clrVal">
            <xsl:apply-templates select="a:solidFill/a:prstClr"/>
          </xsl:variable>
          <xsl:variable name="lumModVal" select="a:solidFill/a:prstClr/a:lumMod/@val"/>
          <xsl:variable name="lumOffVal" select="a:solidFill/a:prstClr/a:lumOff/@val"/>
          <xsl:choose>
            <xsl:when test="$lumOffVal">
              <xsl:call-template name="schemeClrTransparency">
                <xsl:with-param name="schemeClr" select="substring-after($clrVal,'#')"/>
                <xsl:with-param name="trans" select="$lumOffVal"/>
              </xsl:call-template>
            </xsl:when>
            <!-- 白色的变淡，特殊处理一下 -->
            <xsl:when test="$lumModVal and $clrVal='#ffffff'">
              <xsl:call-template name="schemeClrTransparency">
                <xsl:with-param name="schemeClr" select="'000000'"/>
                <xsl:with-param name="trans" select="$lumModVal"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:when test="$lumModVal and $clrVal!='#ffffff'">
              <xsl:call-template name="schemeClrTransparency">
                <xsl:with-param name="schemeClr" select="substring-after($clrVal,'#')"/>
                <xsl:with-param name="trans" select="$lumModVal"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$clrVal"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
      <!--2014-04-20, tangjiang, 添加prstClr颜色参数控制，有待完善 -->

      <!--2014-03-04, tangjiang, 修改Strict OOXML到UOF艺术字字体颜色的转换 start -->
      <xsl:when test="a:gradFill//a:schemeClr">
        <xsl:attribute name="颜色_412F">
          <xsl:apply-templates select="a:gradFill/a:gsLst/a:gs[1]/a:schemeClr"/>
        </xsl:attribute>
      </xsl:when>
      <xsl:when test="a:gradFill/a:gsLst">
        <xsl:attribute name="颜色_412F">
          <xsl:value-of select="concat('#',a:gradFill/a:gsLst/a:gs/a:srgbClr/@val)"/>
        </xsl:attribute>
      </xsl:when>
      <!-- end 2014-03-04, tangjiang, 修改Strict OOXML到UOF艺术字字体颜色的转换-->

      <!--修改自定义图形中字体颜色转换不正确，增加默认字体颜色的转换  liqiuling 2013-1-21  增加smartart字体的转换 2013-3-19 start-->
      <xsl:when test ="ancestor::p:sp/p:style/a:fontRef|ancestor::dsp:sp/dsp:style/a:fontRef">
        <xsl:choose>
          <xsl:when test="ancestor::p:sp/p:style/a:fontRef/a:srgbClr |ancestor::dsp:sp/dsp:style/a:fontRef/a:srgbClr">
            <xsl:attribute name="颜色_412F">
              <xsl:apply-templates select="ancestor::p:sp/p:style/a:fontRef/a:srgbClr |ancestor::dsp:sp/dsp:style/a:fontRef/a:srgbClr"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="ancestor::p:sp/p:style/a:fontRef/a:schemeClr |ancestor::dsp:sp/dsp:style/a:fontRef/a:schemeClr">
            <xsl:attribute name="颜色_412F">
              <xsl:apply-templates select="ancestor::p:sp/p:style/a:fontRef/a:schemeClr  | ancestor::dsp:sp/dsp:style/a:fontRef/a:schemeClr"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
      <!--修改自定义图形中字体颜色转换不正确，增加默认字体颜色的转换  liqiuling 2013-1-21 end-->

      <xsl:when  test="not(a:gradFill or a:solidFill or ancestor::p:sp/p:style/a:fontRef )">
        <xsl:attribute name="颜色_412F">
          <xsl:variable name="typeRef" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@type"/>
          <xsl:variable name="idx" select="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph/@idx"/>
          <xsl:choose>
            <!--liuyangyang 2014-11-13 添加默认表格样式字体颜色转换 start-->
            <xsl:when test="name(../../../..)='a:tc'">
              <xsl:variable name="tableStyleId" select="ancestor::a:tbl/a:tblPr/a:tableStyleId"/>
              <xsl:variable name="rowTh">
                <xsl:apply-templates select="../../../../.." mode="getRowCounter"/>
              </xsl:variable>
              <xsl:choose>
                <xsl:when test="$rowTh=1">
                  <xsl:choose>
                    <xsl:when test="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]/a:firstRow/a:tcTxStyle/a:schemeClr">
                      <xsl:apply-templates select="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]/a:firstRow/a:tcTxStyle/a:schemeClr"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:apply-templates select="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]/a:wholeTbl/a:tcTxStyle/a:schemeClr"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:apply-templates select="ancestor::p:spTree/a:tblStyleLst/a:tblStyle[@styleId=$tableStyleId]/a:wholeTbl/a:tcTxStyle/a:schemeClr"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <!--end liuyangyang 2014-11-13-->
            <xsl:when test="$typeRef='ctrTitle' or $typeRef='title'">
              <xsl:choose>
                <xsl:when test="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill">
                  <xsl:variable name="solidFillColor" select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill"/>
                  <xsl:choose>
                    <xsl:when test="$solidFillColor/a:schemeClr">
                      <xsl:apply-templates select="$solidFillColor/a:schemeClr"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="concat('#',//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
                      <!--<xsl:value-of select="concat('#',//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>-->
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:gradFill">
                  <xsl:variable name="gradFillVal" select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:gradFill"/>
                  <xsl:choose>
                    <xsl:when test="$gradFillVal/a:gsLst/a:gs[1]/a:schemeClr">
                      <xsl:apply-templates select="$gradFillVal/a:gsLst/a:gs[1]/a:schemeClr"/>
                    </xsl:when>
                    <xsl:when test="$gradFillVal/a:gsLst/a:gs[1]/a:srgbClr">
                      <xsl:value-of select="concat('#',$gradFillVal/a:gsLst/a:gs[1]/a:srgbClr/@val)"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="concat('#',$gradFillVal/a:gsLst/a:gs[1]/a:srgbClr/@val)"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="not(//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill)">
                  <xsl:variable name="lvl" select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='title']/p:txBody/a:p/a:pPr/@lvl"/>
                  <xsl:choose>
                    <xsl:when test="$lvl='0' or not(//p:sldMaster[p:sldLayout/@id = $layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='title']/p:txBody/a:p/a:pPr)">

                      <!--2014-03-26, tangjiang, 修复标题字体颜色转换不正确 start -->

                      <xsl:choose>
                        <xsl:when test="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:ln/a:solidFill/a:schemeClr">
                          <xsl:apply-templates select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:ln/a:solidFill/a:schemeClr"/>
                        </xsl:when>
                        <xsl:when test="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:gradFill/a:gsLst/a:gs[1]/a:schemeClr">
                          <xsl:apply-templates select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:gradFill/a:gsLst/a:gs[1]/a:schemeClr"/>
                        </xsl:when>
                        <xsl:when test="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr">
                          <xsl:apply-templates select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr"/>
                        </xsl:when>
                        <xsl:otherwise>
                          <xsl:variable name="tempcolor">
                            <xsl:choose>
                              <xsl:when test="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr">
                                <xsl:value-of select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val"/>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:value-of select="'000000'"/>
                              </xsl:otherwise>
                            </xsl:choose>
                          </xsl:variable>
                          <xsl:value-of select="concat('#',$tempcolor)"/>
                        </xsl:otherwise>
                      </xsl:choose>
                      <!-- end 2014-03-26, tangjiang, 修复标题字体颜色转换不正确 -->

                    </xsl:when>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select=" '#000000' "/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="$typeRef='body'">
              <xsl:variable name="lvl" select="ancestor::a:p/a:pPr/@lvl"/>
              <xsl:choose>
                <xsl:when test="not(ancestor::a:p/a:pPr/@lvl) ">
                  <xsl:variable name="solidFillColor" select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:bodyStyle/a:lvl2pPr/a:defRPr/a:solidFill"/>
                  <xsl:variable name="solidgradColor" select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:bodyStyle/a:lvl2pPr/a:defRPr/a:gradFill"/>
                  <xsl:choose>
                    <xsl:when test="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:gradFill/a:gsLst/a:gs[1]/a:srgbClr/@val">
                      <xsl:value-of select="concat('#',//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:gradFill/a:gsLst/a:gs[1]/a:srgbClr/@val)"/>
                    </xsl:when>
                    <xsl:when test="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:gradFill/a:gsLst/a:gs[1]/a:schemeClr">

                      <xsl:apply-templates select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$idx]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:gradFill/a:gsLst/a:gs[1]/a:schemeClr"/>

                    </xsl:when>
                    <xsl:when test="$solidFillColor/a:schemeClr">
                      <xsl:apply-templates  select="$solidFillColor/a:schemeClr"/>
                    </xsl:when>
                    <xsl:when test="$solidFillColor/a:srgbClr/@val">
                      <xsl:value-of select="concat('#',$solidFillColor/a:srgbClr/@val)"/>
                    </xsl:when>
                    <xsl:when test="$solidgradColor/a:gsLst/a:gs[1]/a:schemeClr">
                      <xsl:apply-templates  select="$solidgradColor/a:gsLst/a:gs[1]/a:schemeClr"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select=" '#000000' "/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:when test="$lvl='0' or $lvl=''">
                  <!--2014-04-27, tangjiang, 修复字体颜色错误 Start -->
                  <xsl:variable name="solidFillColor" select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/a:solidFill"/>
                  <xsl:variable name="solidgradColor" select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:bodyStyle/a:lvl2pPr/a:defRPr/a:gradFill"/>
                  <xsl:choose>
                    <xsl:when test="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr">
                      <xsl:apply-templates  select="//p:sldMaster/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef][1]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr"/>
                    </xsl:when>
                    <xsl:when test="$solidFillColor/a:schemeClr">
                      <xsl:apply-templates  select="$solidFillColor/a:schemeClr"/>
                    </xsl:when>
                    <xsl:when test="$solidFillColor/a:srgbClr/@val">
                      <xsl:value-of select="concat('#',$solidFillColor/a:srgbClr/@val)"/>
                    </xsl:when>
                    <xsl:when test="$solidgradColor/a:gsLst/a:gs[1]/a:schemeClr">
                      <xsl:apply-templates  select="$solidgradColor/a:gsLst/a:gs[1]/a:schemeClr"/>
                    </xsl:when>
                  </xsl:choose>
                  <!--end 2014-04-27, tangjiang, 修复字体颜色错误  -->

                </xsl:when>
                <xsl:when test="$lvl='1'">
                  <!--2014-04-27, tangjiang, 修复字体颜色错误 Start -->
                  <xsl:variable name="solidFillColor" select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:bodyStyle/a:lvl2pPr/a:defRPr/a:solidFill"/>
                  <xsl:variable name="solidgradColor" select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:bodyStyle/a:lvl2pPr/a:defRPr/a:gradFill"/>
                  <xsl:choose>
                    <xsl:when test="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr">
                      <xsl:apply-templates select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$typeRef]/p:txBody/a:lstStyle/a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr"/>
                    </xsl:when>
                    <xsl:when test="$solidFillColor/a:schemeClr">
                      <xsl:apply-templates select="$solidFillColor/a:schemeClr"/>
                    </xsl:when>
                    <xsl:when test="$solidFillColor/a:srgbClr/@val">
                      <xsl:value-of select="concat('#',$solidFillColor/a:srgbClr/@val)"/>
                    </xsl:when>
                    <xsl:when test="$solidgradColor/a:gsLst/a:gs[1]/a:schemeClr">
                      <xsl:apply-templates select="$solidgradColor/a:gsLst/a:gs[1]/a:schemeClr"/>
                    </xsl:when>
                  </xsl:choose>
                  <!--end 2014-04-27, tangjiang, 修复字体颜色错误  -->

                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select=" '#000000' "/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="$typeRef='subTitle'">
              <xsl:choose>
                <xsl:when test="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type = $typeRef]/p:txBody/a:lstStyle/a:lvl1pPr">
                  <xsl:choose>
                    <xsl:when test="not(//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type = $typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill)">
                      <xsl:apply-templates select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr"/>
                    </xsl:when>
                    <xsl:when test="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type = $typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill">
                      <xsl:apply-templates select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type = $typeRef]/p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/*"/>
                    </xsl:when>
                  </xsl:choose>

                </xsl:when>
                <xsl:when test="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:sldLayout[@id=$layoutID]/p:cSld/p:spTree/p:sp/p:txBody/a:lstStyle/a:lvl2pPr">
                  <xsl:apply-templates  select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:otherStyle/a:lvl2pPr/a:defRPr/a:solidFill/a:schemeClr"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select=" '#000000' "/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="not(ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph)">
              <xsl:choose>
                <xsl:when test="ancestor::p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill">
                  <xsl:apply-templates select="ancestor::p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr"/>
                </xsl:when>
                <xsl:when test="not(ancestor::p:txBody/a:lstStyle/a:lvl1pPr)">
                  <xsl:apply-templates  select="//p:sldMaster[p:sldLayout/@id = $layoutID]/p:txStyles/p:otherStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:schemeClr"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select=" '#000000' "/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="$typeRef='ftr' or $typeRef='dt' or $typeRef='hdr' or $typeRef='sldImg' or $typeRef='sldNum'">
              <xsl:choose>
                <xsl:when test="ancestor::p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val">
                  <xsl:value-of select="concat('#',ancestor::p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill/a:srgbClr/@val)"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'#808080'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--end 2014-04-06, tangjiang, 由于字体颜色涉及多处相同的调用，故将字体颜色转换做成模板-->
  
</xsl:stylesheet>
