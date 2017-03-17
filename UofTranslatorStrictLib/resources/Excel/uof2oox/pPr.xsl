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
  <!--pPr.xsl已修改。李杨2012-2-17-->
  <xsl:template name="pPr">
    <xsl:for-each select ="字:大纲级别_417C">
      <xsl:attribute name ="lvl">
        <xsl:value-of select ="."/>
      </xsl:attribute>
    </xsl:for-each>
    <xsl:for-each select ="字:对齐_417D">
      <xsl:if test ="@水平对齐_421D">
        <xsl:attribute name ="algn">
          <xsl:choose>
            <xsl:when test ="@水平对齐_421D='left'">l</xsl:when>
            <xsl:when test ="@水平对齐_421D='right'">r</xsl:when>
            <xsl:when test ="@水平对齐_421D='center'">ctr</xsl:when>
            <xsl:when test ="@水平对齐_421D='justified'">just</xsl:when>
            <xsl:when test ="@水平对齐_421D='distributed'">dist</xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test ="@文字对齐_421E">
        <xsl:attribute name ="fontAlgn">
          <xsl:choose>
            <xsl:when test ="@文字对齐_421E='auto'">auto</xsl:when>
            <xsl:when test ="@文字对齐_421E='top'">t</xsl:when>
            <xsl:when test ="@文字对齐_421E='center'">ctr</xsl:when>
            <xsl:when test ="@文字对齐_421E='base'">base</xsl:when>
            <xsl:when test ="@文字对齐_421E='bottom'">b</xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
    </xsl:for-each>
    <xsl:for-each select ="字:缩进_411D">
      <xsl:for-each select ="字:首行_4111/字:绝对_4107">
        <xsl:attribute name ="indent">
          <xsl:value-of select ="round(number(@值_410F) * 12700)"/>
        </xsl:attribute>
      </xsl:for-each>
      <xsl:for-each select ="字:左_410E/字:绝对_4107">
        <xsl:attribute name ="marL">
          <xsl:value-of select ="round(number(@值_410F) * 12700)"/>
        </xsl:attribute>
      </xsl:for-each>
      <xsl:for-each select ="字:右_4110/字:绝对_4107">
        <xsl:attribute name ="marR">
          <xsl:value-of select ="round(number(@值_410F) * 12700)"/>
        </xsl:attribute>
      </xsl:for-each>
    </xsl:for-each>
    <!--中文习惯首尾字符 默认为“1”-->
    <xsl:if test ="字:是否采用中文习惯首尾字符_4197='false' or 字:是否采用中文习惯首尾字符_4197='0'">
      <xsl:attribute name ="latinLnBrk">0</xsl:attribute>
    </xsl:if>
    <!--允许单词断字，即西文在单词中间换行-->
    <xsl:if test ="字:是否允许单词断字_4194='false' or 字:是否允许单词断字_4194='0'">
      <xsl:attribute name ="eaLnBrk">0</xsl:attribute>
    </xsl:if>
    <!--行首尾标点控制-->
    <xsl:if test ="字:是否行首尾标点控制_4195='true' or 字:是否行首尾标点控制_4195='1'">
      <xsl:attribute name ="hangingPunct">1</xsl:attribute>
    </xsl:if>

    <!--字:行距_417E中的类型  还有一个line-space值。李杨 2012-2-17-->
    <xsl:for-each select ="字:行距_417E">
      <xsl:choose>
        <xsl:when test ="@类型_417F='multi-lines'">
          <a:lnSpc>
            <a:spcPct>
              <xsl:attribute name ="val">
                <xsl:value-of select ="concat(number(@值_4108) * 100,'%')"/>
              </xsl:attribute>
            </a:spcPct>
          </a:lnSpc>
        </xsl:when>
        <xsl:when test ="@类型_417F='fixed'">
          <a:lnSpc>
            <a:spcPts>
              <xsl:attribute name ="val">
                <xsl:value-of select ="number(@值_4108) * 100"/>
              </xsl:attribute>
            </a:spcPts>
          </a:lnSpc>
        </xsl:when>
        <!--最小值、行间距 无对应-->
        <!--最小值不转的话出问题，默认转为固定值-->
        <xsl:when test ="@类型_417F='at-least'">
          <a:lnSpc>
            <a:spcPts>
              <xsl:attribute name ="val">
                <xsl:value-of select ="number(@值_4108) * 100"/>
              </xsl:attribute>
            </a:spcPts>
          </a:lnSpc>
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>

    <xsl:for-each select ="字:段间距_4180">
      <xsl:for-each select ="字:段前距_4181">
        <a:spcBef>
          <xsl:if test ="name(*)='字:相对值_4184'">
            <a:spcPct>
              <xsl:attribute name ="val">
                <xsl:value-of select ="concat(number(字:相对值_4184) * 100,'%')"/>
              </xsl:attribute>
            </a:spcPct>
          </xsl:if>
          <xsl:if test ="name(*)='字:绝对值_4183'">
            <a:spcPts>
              <xsl:attribute name ="val">
                <xsl:value-of select ="number(字:绝对值_4183) * 100"/>
              </xsl:attribute>
            </a:spcPts>
          </xsl:if>
          <!--自动 无对应-->
        </a:spcBef>
      </xsl:for-each>
      <xsl:for-each select ="字:段后距_4185">
        <a:spcAft>
          <xsl:if test ="name(*)='字:相对值_4184'">
            <a:spcPct>
              <xsl:attribute name ="val">
                <xsl:value-of select ="concat(number(字:相对值_4184) * 100,'%')"/>
              </xsl:attribute>
            </a:spcPct>
          </xsl:if>
          <xsl:if test ="name(*)='字:绝对值_4183'">
            <a:spcPts>
              <xsl:attribute name ="val">
                <xsl:value-of select ="number(字:绝对值_4183) * 100"/>
              </xsl:attribute>
            </a:spcPts>
          </xsl:if>
          <!--自动 无对应-->
        </a:spcAft>
      </xsl:for-each>
    </xsl:for-each> 
  <xsl:apply-templates select ="字:句属性_4158" mode="pPr"/>
  </xsl:template>
  
  <!--Modified by LDM in 2011/01/17-->
  <xsl:template match="字:句属性_4158" mode ="pPr">
    <xsl:element name ="a:defRPr">
      <xsl:call-template name="rPr"/>
    </xsl:element>
  </xsl:template>
  <xsl:template name="rPr">

    <!--2014-3-29, update by Qihy, 文本框中字体不正确，start-->
    <!--<xsl:attribute name ="lang">zh-CN</xsl:attribute>
    <xsl:attribute name ="altLang">en-US</xsl:attribute>-->
    <xsl:for-each select ="字:字体_4128[@是否西文绘制_412C]">
      <xsl:choose>
        <xsl:when test="@是否西文绘制_412C = 'true'">
          <xsl:attribute name ="lang">
            <xsl:value-of select="'en-US'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name ="lang">
            <xsl:value-of select="'zh-CN'"/>
          </xsl:attribute>
          <xsl:attribute name ="altLang">
            <xsl:value-of select="'en-US'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
    <!--2014-3-29 end-->
    
    <xsl:for-each select ="字:字体_4128/@字号_412D">
      <xsl:attribute name ="sz">
        <xsl:value-of select="round(number(current()) * 100)"/>
      </xsl:attribute>
    </xsl:for-each>
    <xsl:if test ="字:是否粗体_4130='true' or 字:是否粗体_4130='1'">
      <xsl:attribute name ="b">1</xsl:attribute>
    </xsl:if>
    <xsl:if test ="字:是否斜体_4131='true' or 字:是否斜体_4131='1'">
      <xsl:attribute name ="i">1</xsl:attribute>
    </xsl:if>
    <!--下划线-->
    <xsl:if test="./字:下划线_4136/@线型_4137!='none' and ./字:下划线_4136/@线型_4137!='0'">
      <xsl:attribute name="u">
        <xsl:choose>
          <xsl:when test="./字:下划线_4136/@线型_4137 = 'single'">
            <xsl:choose>
              <xsl:when test="not(./字:下划线_4136/@虚实_4138)">
                <xsl:value-of select="'sng'"/>
              </xsl:when>
              <xsl:when test="./字:下划线_4136/@虚实_4138 = 'square-dot'">
                <xsl:value-of select="'dottedHeavy'"/>
              </xsl:when>
              <xsl:when test="./字:下划线_4136/@虚实_4138 = 'dash'">
                <xsl:value-of select="'dashLong'"/>
              </xsl:when>
              <xsl:when test="./字:下划线_4136/@虚实_4138 = 'long-dash'">
                <xsl:value-of select="'dashLong'"/>
              </xsl:when>
              <xsl:when test="./字:下划线_4136/@虚实_4138 = 'dash-dot'">
                <xsl:value-of select="'dotDash'"/>
              </xsl:when>
              <xsl:when test="./字:下划线_4136/@虚实_4138 = 'dash-dot-dot'">
                <xsl:value-of select="'dotDotDash'"/>
              </xsl:when>
            </xsl:choose>
          </xsl:when>
          <xsl:when test="./字:下划线_4136/@线型_4137 = 'double'">
            <xsl:value-of select="'dbl'"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
    <!--删除线-->
    <xsl:if test="字:删除线_4135">
      <xsl:attribute name="strike">
        <xsl:choose>
          <xsl:when test="字:删除线_4135='single'">
            <xsl:value-of select="'sngStrike'"/>
            </xsl:when>
          <xsl:when test="字:删除线_4135='double'">
            <xsl:value-of select="'dblStrike'"/>
          </xsl:when>
          <xsl:when test="字:删除线_4135='none'">
            <xsl:value-of select="'noStrike'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'noStrike'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
    <!--
    <xsl:attribute name ="dirty">0</xsl:attribute>
    <xsl:attribute name ="smtClean">0</xsl:attribute>
    -->
    <!--调整字间距-->
    <xsl:if test="字:调整字间距_4146">
      <xsl:attribute name ="kern">
        <xsl:value-of select ="number(./字:调整字间距_4146) *100"/>
      </xsl:attribute>
    </xsl:if>
    <!--醒目字体-->
    <xsl:if test="字:醒目字体类型_4141">
      <xsl:attribute name="cap">
        <xsl:choose>
          <xsl:when test="字:醒目字体类型_4141='none'">none</xsl:when>
          <xsl:when test="字:醒目字体类型_4141='small-caps'">small</xsl:when>
          <xsl:when test="字:醒目字体类型_4141='uppercase'">all</xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if> 
    <!--字符间距-->
    <xsl:if test="字:字符间距_4145">
      <xsl:attribute name ="spc">
        <xsl:value-of select ="number(./字:字符间距_4145) *100"/>
      </xsl:attribute>
    </xsl:if>
    <!--上下标-->
    <xsl:if test="字:上下标类型_4143">
      <xsl:attribute name ="baseline">
        <xsl:choose>
          <xsl:when test ="字:上下标类型_4143='sup'">30%</xsl:when>
          <xsl:when test ="字:上下标类型_4143='none'">0%</xsl:when>
          <xsl:when test ="字:上下标类型_4143='sub'">-25%</xsl:when>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="./字:字体_4128/@颜色_412F">
        <a:solidFill>
          <a:srgbClr>
            <xsl:attribute name ="val">
              <xsl:value-of select ="substring-after(./字:字体_4128/@颜色_412F,'#')"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:solidFill>
    </xsl:if>
    <xsl:if test="./字:下划线_4136/@颜色_412F">
        <a:uFill>
          <a:solidFill>
            <a:srgbClr>
              <xsl:attribute name ="val">
                <xsl:value-of select="substring-after(./字:下划线_4136/@颜色_412F,'#')"/>
              </xsl:attribute>
            </a:srgbClr>
          </a:solidFill>
        </a:uFill>
    </xsl:if>

    <!--阴影-->
    <xsl:if test="字:是否阴影_4140='true' or 字:是否阴影_4140='1'">
      <a:effectLst>
        <a:prstShdw prst="shdw2">
          <a:srgbClr val="000000"/>
        </a:prstShdw>
      </a:effectLst>
    </xsl:if>
    <xsl:if test="字:突出显示颜色_4132">
      <a:highlight>
        <a:srgbClr>
          <xsl:value-of select="substring-after(字:突出显示颜色_4132,'#')"/>
        </a:srgbClr>
      </a:highlight>
    </xsl:if>

    <xsl:if test ="字:字体_4128">
      <xsl:if test="字:字体_4128/@西文字体引用_4129">
        <xsl:variable name="latin">
          <xsl:value-of select="字:字体_4128/@西文字体引用_4129"/>
        </xsl:variable>
        <a:latin>
          <xsl:attribute name ="typeface">
            <!--20121217,gaoyuwei，解决BUG"文本框中艺术字大小不正确"UOF-OOXML（实际是字体丢失） start--> 
			
            <!--<xsl:value-of select ="/uof:UOF/式样:式样集_990B/式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$latin]/式样:字体族_9900"/>-->
			  <xsl:value-of select ="/uof:UOF/式样:式样集_990B/式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$latin]/@名称_9903"/>
			  			  
			  <!--end-->
          </xsl:attribute>
        </a:latin>
      </xsl:if>
      <xsl:if test="字:字体_4128/@中文字体引用_412A">
        <xsl:variable name="ea">
          <xsl:value-of select="字:字体_4128/@中文字体引用_412A"/>
        </xsl:variable>
        <a:ea>
          <xsl:attribute name ="typeface">
            		  <!--20121217,gaoyuwei，解决BUG"文本框中艺术字大小不正确"UOF-OOXML（实际是字体丢失） start-->
			  <!--<xsl:value-of select ="/uof:UOF/式样:式样集_990B/式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$ea]/式样:字体族_9900"/>-->
			  <xsl:value-of select ="/uof:UOF/式样:式样集_990B/式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$ea]/@名称_9903"/>
			  <!--end-->
			 
          </xsl:attribute>
        </a:ea>
      </xsl:if>
      <xsl:if test="字:字体_4128/@特殊字体引用_412B">
        <xsl:variable name="cs">
          <xsl:value-of select="字:字体_4128/@特殊字体引用_412B"/>
        </xsl:variable>
        <a:cs>
          <xsl:attribute name ="typeface">
            <xsl:value-of select ="/uof:UOF/式样:式样集_990B/式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$cs]/式样:字体族_9900"/>
          </xsl:attribute>
        </a:cs>
      </xsl:if>
    </xsl:if>
  </xsl:template>

  <!--obsolute-->
  <!--Modified by LDM in 2010/12/16-->
  <!--default paragraph properties-->
  <!--Is it necessary?-->
  <xsl:template name="defaultParaProp">
    <a:pPr>
      <a:defRPr>
        <xsl:choose>
          <xsl:when test="表:字体/字:字体[@字:西文字体引用]">
            <xsl:element name="a:latin">
              <xsl:variable name="latinFont" select="表:字体/字:字体/@字:西文字体引用"/>
              <xsl:attribute name="typeface">
                <xsl:value-of select="/uof:UOF/uof:式样集/uof:字体集/uof:字体声明[@uof:标识符 = $latinFont]/@uof:名称"/>
              </xsl:attribute>
            </xsl:element>
          </xsl:when>
          <xsl:when test="not(表:字体/字:字体[@字:西文字体引用])">
            <xsl:element name="a:latin">
              <xsl:attribute name="typeface">
                <xsl:value-of select="'Times New Roman'"/>
              </xsl:attribute>
            </xsl:element>
          </xsl:when>
          <xsl:when test="表:字体/字:字体[@字:中文字体引用]">
            <xsl:element name="a:ea">
              <xsl:variable name="eaFont" select="表:字体/字:字体/@字:中文字体引用"/>
              <xsl:attribute name="typeface">
                <xsl:value-of select="/uof:UOF/uof:式样集/uof:字体集/uof:字体声明[@uof:标识符=$eaFont]/@uof:名称"/>
              </xsl:attribute>
            </xsl:element>
          </xsl:when>
          <xsl:when test="not(表:字体/字:字体[@字:中文字体引用])">
            <xsl:element name="a:latin">
              <xsl:attribute name="typeface">
                <xsl:value-of select="'永中宋体'"/>
              </xsl:attribute>
            </xsl:element>
          </xsl:when>
        </xsl:choose>
      </a:defRPr>
    </a:pPr>
  </xsl:template>

</xsl:stylesheet>

<!--
<xsl:for-each select="字:下划线/@字:类型">
  <xsl:attribute name="u">
    <xsl:choose>
      <xsl:when test="not(../@字:字下划线)">
        <xsl:choose>
          <xsl:when test="current()='dotted-heavy'">dottedHeavy</xsl:when>
          <xsl:when test="current()='dashed-heavy'">dashHeavy</xsl:when>
          <xsl:when test="current()='dash-long'">dashLong</xsl:when>
          <xsl:when test="current()='dash-long-heavy'">dashLongHeavy</xsl:when>
          <xsl:when test="current()='dot-dash'">dotDash</xsl:when>
          <xsl:when test="current()='dash-dot-heavy'">dotDashHeavy</xsl:when>
          <xsl:when test="current()='dot-dot-dash'">dotDotDash</xsl:when>
          <xsl:when test="current()='dash-dot-dot-heavy'">dotDotDashHeavy</xsl:when>
          <xsl:when test="current()='wave-heavy'">wavyHeavy</xsl:when>
          <xsl:when test="current()='wavy-double'">wavyDbl</xsl:when>
          <xsl:when test="current()='double'">dbl</xsl:when>
          <xsl:when test="current()='single'">sng</xsl:when>
          <xsl:when test="current()='none'">none</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'sng'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test =".!='none'">words</xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:attribute>
</xsl:for-each>
-->