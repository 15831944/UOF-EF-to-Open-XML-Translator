<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:fo="http://www.w3.org/1999/XSL/Format"
xmlns:dc="http://purl.org/dc/elements/1.1/"
xmlns:dcterms="http://purl.org/dc/terms/"
xmlns:dcmitype="http://purl.org/dc/dcmitype/"
xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    
xmlns:app="http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
                
xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph">
  
  <xsl:output method="xml" version="2.0" encoding="UTF-8" indent="yes"/>
  <xsl:template name="prestep2" match="/">
    <p:presentation>
      <!-- 2010.03.25 马有旭 -->
      <xsl:copy-of select="p:presentation/@*"/>
      <xsl:for-each select="p:presentation/*">
        <xsl:choose>
          <xsl:when test="name()='p:sldMaster' or name()='p:notesMaster' or name()='p:handoutMaster'">
            <xsl:call-template name="master"/>
          </xsl:when>
          <xsl:when test="name()='p:sld' or  name()='p:notes'">
            <xsl:apply-templates select="."/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:copy-of select="."/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
    </p:presentation>
  </xsl:template>

  <xsl:template name="master">
    <xsl:copy>
      <xsl:for-each select="@*">
        <xsl:copy/>
      </xsl:for-each>
      <xsl:for-each select="*">
        <xsl:if test="name(.)!='p:sldLayout' and name(.)!='p:cSld' and name(.)!='p:clrMap'">
          <xsl:copy-of select="."/>
        </xsl:if>
        <xsl:if test="name(.)='p:cSld'">
          <xsl:apply-templates select=".">
            <xsl:with-param name="type" select="'master'"/>
          </xsl:apply-templates>
        </xsl:if>
        <xsl:if test="name(.)='p:sldLayout'">
          <xsl:apply-templates select="."/>
        </xsl:if>
        <xsl:if test="name(.)='p:clrMap'">
          <xsl:apply-templates select="."/>
        </xsl:if>
      </xsl:for-each>
    </xsl:copy>
  </xsl:template>
  <!--p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/-->
  <xsl:template match="p:clrMap">
    <xsl:copy>
      <xsl:attribute name="bg1">
        <xsl:call-template name="clrMapVal">
          <xsl:with-param name="attrVal" select="@bg1"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="tx1">
        <xsl:call-template name="clrMapVal">
          <xsl:with-param name="attrVal" select="@tx1"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="bg2">
        <xsl:call-template name="clrMapVal">
          <xsl:with-param name="attrVal" select="@bg2"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="tx2">
        <xsl:call-template name="clrMapVal">
          <xsl:with-param name="attrVal" select="@tx2"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="accent1">
        <xsl:call-template name="clrMapVal">
          <xsl:with-param name="attrVal" select="@accent1"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="accent2">
        <xsl:call-template name="clrMapVal">
          <xsl:with-param name="attrVal" select="@accent2"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="accent3">
        <xsl:call-template name="clrMapVal">
          <xsl:with-param name="attrVal" select="@accent3"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="accent4">
        <xsl:call-template name="clrMapVal">
          <xsl:with-param name="attrVal" select="@accent4"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="accent5">
        <xsl:call-template name="clrMapVal">
          <xsl:with-param name="attrVal" select="@accent5"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="accent6">
        <xsl:call-template name="clrMapVal">
          <xsl:with-param name="attrVal" select="@accent6"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="hlink">
        <xsl:call-template name="clrMapVal">
          <xsl:with-param name="attrVal" select="@hlink"/>
        </xsl:call-template>
      </xsl:attribute>
      <xsl:attribute name="folHlink">
        <xsl:call-template name="clrMapVal">
          <xsl:with-param name="attrVal" select="@folHlink"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:copy>
  </xsl:template>
  <xsl:template name="clrMapVal">
    <xsl:param name="attrVal"/>
    <xsl:for-each select="../following-sibling::a:theme[1]/a:themeElements/a:clrScheme/*[name()=concat('a:',$attrVal)]">
      <xsl:choose>
        <xsl:when test="a:sysClr">
          <xsl:if test="a:sysClr/@lastClr">
            <xsl:value-of select="concat('#',a:sysClr/@lastClr)"/>
          </xsl:if>
          <xsl:if test="not(a:sysClr/@lastClr)">auto</xsl:if>
        </xsl:when>
        <xsl:when test="a:srgbClr">
          <xsl:value-of select="concat('#',a:srgbClr/@val)"/>
        </xsl:when>
        <xsl:otherwise>auto</xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <xsl:template match="p:sldLayout">
    <xsl:copy>
      <xsl:for-each select="@*">
        <xsl:copy/>
      </xsl:for-each>
      <xsl:apply-templates select="p:cSld">
        <xsl:with-param name="type" select="'sldLayout'"/>
      </xsl:apply-templates>
      <xsl:copy-of select="*[name()!='p:cSld']"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="p:notes">
    <xsl:variable name="notesMaster" select="substring-after(following-sibling::rel:Relationships[1]/rel:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/notesMaster']/@Target,'notesMasters/')"/>
    <xsl:copy>
      <xsl:for-each select="@*">
        <xsl:copy/>
      </xsl:for-each>
      <xsl:apply-templates select="p:cSld">
        <xsl:with-param name="type" select="'notes'"/>
        <xsl:with-param name="refFrom" select="$notesMaster"/>
      </xsl:apply-templates>
      <xsl:copy-of select="*[name()!='p:cSld']"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="p:sld">
    <p:sld>
      <xsl:variable name="layout" select="substring-after(following-sibling::rel:Relationships[1]/rel:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout']/@Target,'slideLayouts/')"/>
      <xsl:attribute name="sldLayoutID">
        <xsl:value-of select="$layout"/>
      </xsl:attribute>
		<xsl:if test="following-sibling::rel:Relationships[1]/rel:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/comments']">
			
		<xsl:variable name="comment" select="substring-after(following-sibling::rel:Relationships[1]/rel:Relationship[@Type='http://purl.oclc.org/ooxml/officeDocument/relationships/comments']/@Target,'comments/')"/>
		<xsl:attribute name="commentID">
			<xsl:value-of select="$comment"/>
		</xsl:attribute>
		</xsl:if>
      <xsl:for-each select="@*">
        <xsl:copy/>
      </xsl:for-each>
      <xsl:apply-templates select="p:cSld">
        <xsl:with-param name="type" select="'sld'"/>
        <xsl:with-param name="refFrom" select="$layout"/>
      </xsl:apply-templates>
      <xsl:copy-of select="*[name()!='p:cSld']"/>
    </p:sld>
  </xsl:template>


  <xsl:template match="p:cSld">
    <xsl:param name="type"/>
    <xsl:param name="refFrom"/>
    <p:cSld>
      <xsl:for-each select="@*">
        <xsl:copy/>
      </xsl:for-each>
      <xsl:for-each select="*">
        <xsl:if test="name(.)!='p:spTree'">
          <xsl:copy-of select="."/>
        </xsl:if>
        <xsl:if test="name(.)='p:spTree'">
          <p:spTree>
            <xsl:call-template name="sps">
              <xsl:with-param name="type" select="$type"/>
              <xsl:with-param name="refFrom" select="$refFrom"/>
            </xsl:call-template>
          </p:spTree>
        </xsl:if>
      </xsl:for-each>
    </p:cSld>
  </xsl:template>

  <xsl:template name="sps">
    <xsl:param name="type"/>
    <xsl:param name="refFrom"/>
    <xsl:for-each select="@*">
      <xsl:copy/>
    </xsl:for-each>
    <xsl:for-each select="*">
      <xsl:choose>
        <xsl:when test="name(.)='p:sp' or name(.)='p:grpSp'">
          <xsl:apply-templates select=".">
            <xsl:with-param name="type" select="$type"/>
            <xsl:with-param name="refFrom" select="$refFrom"/>
          </xsl:apply-templates>
        </xsl:when>
        <xsl:otherwise>
          <xsl:copy-of select="."/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
    <!--2010.04.28-->
    <xsl:if test="not(p:sp[p:nvSpPr/p:nvPr/p:ph/@type='ftr']) and $type='master'">
      <xsl:apply-templates select="//p:sld[//p:ph/@type='ftr']//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='ftr']"/>
    </xsl:if>
    <xsl:if test="not(p:sp[p:nvSpPr/p:nvPr/p:ph/@type='dt']) and $type='master'">
      <xsl:apply-templates select="//p:sld[//p:ph/@type='dt']//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='dt']"/>
    </xsl:if>
    <xsl:if test="not(p:sp[p:nvSpPr/p:nvPr/p:ph/@type='sldNum']) and $type='master'">
      <xsl:apply-templates select="//p:sld[//p:ph/@type='sldNum']//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='sldNum']"/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="p:grpSp">
    <xsl:param name="type"/>
    <xsl:param name="refFrom"/>
    <xsl:variable name="pht" select="p:nvSpPr/p:nvPr/p:ph/@type"/>
    <p:grpSp>
      <xsl:for-each select="@*">
        <xsl:copy/>
      </xsl:for-each>
      <xsl:if test="$pht">
        <xsl:attribute name="hID">
          <xsl:if test="$type='sldLayout'">
            <xsl:value-of select="ancestor::p:sldMaster/p:cSld/p:spTree//p:grpSp[p:nvSpPr/p:nvPr/p:ph/@type=$pht]/@id"/>
          </xsl:if>
          <xsl:if test="$type='sld'">
            <xsl:value-of select="//p:sldMaster/p:sldLayout[@id = $refFrom]/p:cSld/p:spTree//p:grpSp[p:nvSpPr/p:nvPr/p:ph/@type=$pht]/@id"/>
          </xsl:if>
        </xsl:attribute>
      </xsl:if>
      <xsl:call-template name="sps"/>
    </p:grpSp>
  </xsl:template>

  <xsl:template match="p:sp">
    <xsl:param name="type"/>
    <xsl:param name="refFrom"/>
    <xsl:variable name="pht" select="p:nvSpPr/p:nvPr/p:ph/@type"/>
    <xsl:variable name="phi" select="p:nvSpPr/p:nvPr/p:ph/@idx"/>
    <p:sp>
      <xsl:for-each select="@*">
        <xsl:copy/>
      </xsl:for-each>
      <xsl:variable name="localID" select="@id"/>

      <xsl:variable name="hrID">
        <xsl:choose>
          <xsl:when test="p:nvSpPr/p:nvPr/p:ph">
            <xsl:choose>
              <xsl:when test="$type='sldLayout'">
                <xsl:choose>
                  <!--$pht='' is wrong-->
                  <!--原来的<xsl:when test="not($pht) or $pht='obj' or $pht='body' or $pht='subTitle'">
									<xsl:value-of select="ancestor::p:sldMaster/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='body']/@id"/>
									</xsl:when>
									<xsl:when test="$pht='ctrTitle' or $pht='title'">
									<xsl:value-of select="ancestor::p:sldMaster/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='title']/@id"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="ancestor::p:sldMaster/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$pht]/@id"/>
									</xsl:otherwise>-->
                  <xsl:when test="$pht='obj' or $pht='body' or $pht='subTitle'">
                    <xsl:if test ="@idx">
                      <xsl:value-of select="ancestor::p:sldMaster/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='body' and p:nvSpPr/p:nvPr/p:ph/@idx=current()/@idx]/@id"/>
                    </xsl:if>
                    <xsl:if test ="not(@idx)">
                      <xsl:value-of select="ancestor::p:sldMaster/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='body']/@id"/>
                    </xsl:if>
                  </xsl:when>
                  <xsl:when test="not($pht)">
                    <xsl:if test ="@idx">
                      <xsl:value-of select="ancestor::p:sldMaster/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=current()/@idx]/@id"/>
                    </xsl:if>
                    <xsl:if test ="not(@idx)">
                      <xsl:value-of select="ancestor::p:sldMaster/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='body']/@id"/>
                    </xsl:if>
                  </xsl:when>
                  <xsl:when test="$pht='ctrTitle' or $pht='title'">
                    <xsl:value-of select="ancestor::p:sldMaster/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@type='title']/@id"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:if test ="@idx">
                      <xsl:value-of select="ancestor::p:sldMaster/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$pht and p:nvSpPr/p:nvPr/p:ph/@idx=current()/@idx]/@id"/>
                    </xsl:if>
                    <xsl:if test ="not(@idx)">
                      <xsl:value-of select="ancestor::p:sldMaster/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$pht]/@id"/>
                    </xsl:if>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <xsl:when test="$type='notes'">
                <xsl:value-of select="ancestor::p:presentation/p:notesMaster[@id = $refFrom]/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$pht]/@id"/>
              </xsl:when>
              <xsl:when test="$type='sld'">
                <xsl:choose>
                  <xsl:when test="not($phi)and $pht!=''">
                    <xsl:value-of select="ancestor::p:presentation/p:sldMaster/p:sldLayout[@id = $refFrom]/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@type=$pht]/@id"/>
                  </xsl:when>
                  <xsl:when test ="$phi!=''">
                    <xsl:value-of select="ancestor::p:presentation/p:sldMaster/p:sldLayout[@id = $refFrom]/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$phi]/@id"/>
                  </xsl:when>
                  <!--xsl:when test ="$phi='' and $pht=''">
                    <xsl:value-of select="ancestor::p:presentation/p:sldMaster/p:sldLayout[@id = $refFrom]/p:cSld/p:spTree//p:sp[p:nvSpPr/p:nvPr/p:ph/@idx=$phi]/@id"/>
                  </xsl:when-->
                </xsl:choose>

              </xsl:when>
              <xsl:otherwise>0</xsl:otherwise >
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise >
        </xsl:choose>
      </xsl:variable>
      <xsl:if test="$hrID != '0'">
        <xsl:attribute name="hID">
          <xsl:value-of select="$hrID"/>
        </xsl:attribute>
      </xsl:if>

      <xsl:for-each select="*">

        <xsl:if test="name(.)!='p:txBody'">
          <xsl:copy-of select="."/>
        </xsl:if>
        <xsl:if test="name(.)='p:txBody'">
          <xsl:apply-templates select=".">
            <xsl:with-param name="pht" select="$pht"/>
            <xsl:with-param name="phi" select="$phi"/>
            <xsl:with-param name="type" select="$type"/>
            <xsl:with-param name="refFrom" select="$refFrom"/>
            <xsl:with-param name="hID" select="$hrID"/>
          </xsl:apply-templates>
        </xsl:if>
      </xsl:for-each>
    </p:sp>
  </xsl:template>

  <xsl:template match="p:txBody">
    <xsl:param name="pht"/>
    <xsl:param name="phi"/>
    <xsl:param name="type"/>
    <xsl:param name="refFrom"/>
    <xsl:param name="hID"/>
    <p:txBody>
      <xsl:for-each select="*">
        <xsl:if test="name(.)!='a:p'">
          <xsl:copy-of select="."/>
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
              <xsl:with-param name="phi" select="$phi"/>
              <xsl:with-param name="type" select="$type"/>
              <xsl:with-param name="refFrom" select="$refFrom"/>
              <xsl:with-param name="hID" select="$hID"/>
            </xsl:apply-templates>
          </xsl:when>
        </xsl:choose>
        <!--2014-06-05, tangjiang, 将只有段落结束符的段落预处理为含有一个空格符的段落，修复文字位置转换差异 end -->
        
      </xsl:for-each>
    </p:txBody>
  </xsl:template>
  <xsl:template match="a:p">
    <xsl:param name="pht"/>
    <xsl:param name="phi"/>
    <xsl:param name="type"/>
    <xsl:param name="refFrom"/>
    <xsl:param name="hID"/>
    <a:p>
      <xsl:for-each select="@*">
        <xsl:copy/>
      </xsl:for-each>
      <xsl:variable name="lvl" select="a:pPr/@lvl"/>
      <!-- 
      10.23 黎美秀修改 在版式中增加图形时，式样引用信息丢失 此处需要增加判断

      
      -->
      <xsl:if test="$type='sldLayout'">
        <xsl:attribute name="styleID">
          <xsl:variable name="pPrID">
            <xsl:if test="preceding-sibling::a:lstStyle/*">
              <xsl:call-template name="lstStyle">
                <xsl:with-param name="lvl" select="$lvl"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:variable>
          <xsl:if test="$pPrID">
            <xsl:value-of select="$pPrID"/>
          </xsl:if>
          <!-- ='' and not() aren't the same-->

          <xsl:if test="$pPrID=''">
            <!--find the right style in right Master/sp/p/@styleID by the lvl of p-->
            <xsl:variable name="sID">
              <!--@lvl exist-->
              <xsl:if test="a:pPr/@lvl">
                <xsl:value-of select="ancestor::p:sldMaster//p:sp[@id = $hID]/p:txBody/a:p[a:pPr/@lvl=$lvl]/@styleID"/>
              </xsl:if>
              <xsl:if test="not(a:pPr/@lvl)">
                <xsl:value-of select="ancestor::p:sldMaster//p:sp[@id = $hID]/p:txBody/a:p[a:pPr/@lvl='0' or not(a:pPr/@lvl)]/@styleID"/>
              </xsl:if>
            </xsl:variable>

            <xsl:if test="$sID !=''">
              <xsl:value-of select="$sID"/>
            </xsl:if>
            <!--Master p with this lvl not exist;look for the style in the txStyles of Master-->
            <!-- 10.23 2 黎美秀 问题主要出在此处，因为图形并非来自母版继承-->
            <xsl:if test="$sID=''">
              <xsl:if test="ancestor::p:sp/p:nvSpPr/p:nvPr/p:ph">
                <xsl:choose>
                  <xsl:when test="$pht='title'">
                    <xsl:call-template name="titleStyle">
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:when test="$pht='body' or $pht='obj' or not($pht)">
                    <xsl:call-template name="bodyStyle">
                    </xsl:call-template>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:call-template name="otherStyles"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:if>
              <!-- 10.23 黎美秀增加 不知道调用的式样对不对

              一个是判断自己添加图形的，一个用于判断母版中继承的图形

              -->
              <xsl:if test ="ancestor::p:sp/p:nvSpPr/p:nvPr[@userDrawn='1'] or not(ancestor::p:sp/p:nvSpPr/p:nvPr/*)">

                <xsl:call-template name="bodyStyle"/>

              </xsl:if>

            </xsl:if>
          </xsl:if>
        </xsl:attribute>
      </xsl:if>

      <xsl:if test="$type='notes'">
        <xsl:attribute name="styleID">
          <xsl:variable name="pPrID">
            <xsl:if test="preceding-sibling::a:lstStyle/*">
              <xsl:call-template name="lstStyle">
                <xsl:with-param name="lvl" select="$lvl"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:variable>
          <xsl:if test="$pPrID">
            <xsl:value-of select="$pPrID"/>
          </xsl:if>
          <xsl:if test="$pPrID=''">
            <!--find the right style in right Master/sp/p/@styleID by the lvl of p-->
            <xsl:variable name="masterID">
              <!--@lvl exist-->
              <xsl:if test="a:pPr/@lvl">
                <xsl:value-of select="//p:notesMaster//p:sp[@id = $hID]/p:txBody/a:p[a:pPr/@lvl=$lvl]/@styleID"/>
              </xsl:if>
              <xsl:if test="not(a:pPr/@lvl)">
                <xsl:value-of select="//p:notesMaster//p:sp[@id = $hID]/p:txBody/a:p[a:pPr/@lvl='0' or not(a:pPr/@lvl)]/@styleID"/>
              </xsl:if>
            </xsl:variable>
            <xsl:if test="$masterID !=''">
              <xsl:value-of select="$masterID"/>
            </xsl:if>
            <!--Master p with this lvl not exist;look for the style in the txStyles of Master-->
            <xsl:if test="$masterID=''">
              <xsl:call-template name="notesStyle">
                <xsl:with-param name="refFrom" select="$refFrom"/>
              </xsl:call-template>
            </xsl:if>
          </xsl:if>
        </xsl:attribute>
      </xsl:if>
      <xsl:copy-of select="./*"/>
    </a:p>
  </xsl:template>

  <xsl:template name="notesStyle">
    <xsl:param name="refFrom"/>
    <xsl:choose>
      <xsl:when test="a:pPr/@lvl='0'">
        <xsl:value-of select="//p:notesMaster[@id=$refFrom]/p:notesStyle/a:lvl1pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='1'">
        <xsl:value-of select="//p:notesMaster[@id=$refFrom]/p:notesStyle/a:lvl2pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='2'">
        <xsl:value-of select="//p:notesMaster[@id=$refFrom]/p:notesStyle/a:lvl3pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='3'">
        <xsl:value-of select="//p:notesMaster[@id=$refFrom]/p:notesStyle/a:lvl4pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='4'">
        <xsl:value-of select="//p:notesMaster[@id=$refFrom]/p:notesStyle/a:lvl5pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='5'">
        <xsl:value-of select="//p:notesMaster[@id=$refFrom]/p:notesStyle/a:lvl6pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='6'">
        <xsl:value-of select="//p:notesMaster[@id=$refFrom]/p:notesStyle/a:lvl7pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='7'">
        <xsl:value-of select="//p:notesMaster[@id=$refFrom]/p:notesStyle/a:lvl8pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='8'">
        <xsl:value-of select="//p:notesMaster[@id=$refFrom]/p:notesStyle/a:lvl9pPr/@id"/>
      </xsl:when>
      <xsl:when test="not(a:pPr/@lvl)">
        <xsl:value-of select="//p:notesMaster[@id=$refFrom]/p:notesStyle/a:lvl1pPr/@id"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="lstStyle">
    <xsl:param name="lvl"/>
    <xsl:choose>
      <xsl:when test="$lvl='0' or $lvl='' or not($lvl)">
        <xsl:value-of select="../a:lstStyle/a:lvl1pPr/@id"/>
      </xsl:when>
      <xsl:when test="$lvl='1'">
        <xsl:value-of select="../a:lstStyle/a:lvl2pPr/@id"/>
      </xsl:when>
      <xsl:when test="$lvl='2'">
        <xsl:value-of select="../a:lstStyle/a:lvl3pPr/@id"/>
      </xsl:when>
      <xsl:when test="$lvl='3'">
        <xsl:value-of select="../a:lstStyle/a:lvl4pPr/@id"/>
      </xsl:when>
      <xsl:when test="$lvl='4'">
        <xsl:value-of select="../a:lstStyle/a:lvl5pPr/@id"/>
      </xsl:when>
      <xsl:when test="$lvl='5'">
        <xsl:value-of select="../a:lstStyle/a:lvl6pPr/@id"/>
      </xsl:when>
      <xsl:when test="$lvl='6'">
        <xsl:value-of select="../a:lstStyle/a:lvl7pPr/@id"/>
      </xsl:when>
      <xsl:when test="$lvl='7'">
        <xsl:value-of select="../a:lstStyle/a:lvl8pPr/@id"/>
      </xsl:when>
      <xsl:when test="$lvl='8'">
        <xsl:value-of select="../a:lstStyle/a:lvl9pPr/@id"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="otherStyles">
    <xsl:choose>
      <xsl:when test="a:pPr/@lvl='0'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl1pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='1'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl2pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='2'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl3pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='3'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl4pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='4'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl5pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='5'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl6pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='6'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl7pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='7'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl8pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='8'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl9pPr/@id"/>
      </xsl:when>
      <xsl:when test="not(a:pPr/@lvl)">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:otherStyle/a:lvl1pPr/@id"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="titleStyle">
    <xsl:choose>
      <xsl:when test="a:pPr/@lvl='0'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='1'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl2pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='2'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl3pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='3'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl4pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='4'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl5pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='5'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl6pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='6'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl7pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='7'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl8pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='8'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl9pPr/@id"/>
      </xsl:when>
      <xsl:when test="not(a:pPr/@lvl)">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:titleStyle/a:lvl1pPr/@id"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="bodyStyle">
    <xsl:choose>
      <xsl:when test="a:pPr/@lvl='0'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='1'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl2pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='2'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl3pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='3'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl4pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='4'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl5pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='5'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl6pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='6'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl7pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='7'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl8pPr/@id"/>
      </xsl:when>
      <xsl:when test="a:pPr/@lvl='8'">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl9pPr/@id"/>
      </xsl:when>
      <xsl:when test="not(a:pPr/@lvl)">
        <xsl:value-of select="ancestor::p:sldMaster/p:txStyles/p:bodyStyle/a:lvl1pPr/@id"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>
