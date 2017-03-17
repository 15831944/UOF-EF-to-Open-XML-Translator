<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
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
                xmlns:扩展="http://schemas.uof.org/cn/2009/extend"
                xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
                xmlns:公式="http://schemas.uof.org/cn/2009/equations"
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:pzip="urn:u2o:xmlns:post-processings:special"
                xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                xmlns:dc="http://purl.org/dc/elements/1.1/"
                xmlns:dcterms="http://purl.org/dc/terms/"
                xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:pc="http://schemas.openxmlformats.org/package/2006/content-types"
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:ori="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                xmlns:u2opic="urn:u2opic:xmlns:post-processings:special"
                xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
	<xsl:import href="metadata.xsl"/>
	<xsl:import href="sheetprop.xsl"/>
	<xsl:import href="sheets.xsl"/>
	<xsl:import href="sheetfilter.xsl"/>
	<xsl:import href="sheetbreak.xsl"/>
	<xsl:import href="sheetcommon.xsl"/>
	<xsl:import href="hyperlinks.xsl"/>
	<xsl:import href="metadata2.xsl"/>
	<xsl:import href="style.xsl"/>
	<xsl:import href="table.xsl"/>
  <!--<xsl:import href ="table2.xsl"/>-->
	<xsl:import href="object.xsl"/>
  <xsl:param name="outputFile"/>
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
  <xsl:template name="oox2uof" match ="/">
    <pzip:archive pzip:target="{$outputFile}">

      <xsl:if test="//w:media">
        <xsl:copy-of select="//w:media/*"/>
      </xsl:if>
      <!-- uof.xml-->
      <pzip:entry pzip:target="uof.xml">
        <uof:UOF_0000 mimetype_0001="vnd.uof.spreadsheet" language_0002="cn" version_0003="2.0">
          <uof:概览_0006>
            <uof:工作表统计_0007>
              <uof:工作表摘要_0010 标识符_0008="SHEET_1" 名称_0009="工作表1"/>
              <uof:工作表摘要_0010 标识符_0008="SHEET_2" 名称_0009="工作表2"/>
              <uof:工作表摘要_0010 标识符_0008="SHEET_3" 名称_0009="工作表3"/>
            </uof:工作表统计_0007>
          </uof:概览_0006>
        </uof:UOF_0000>
      </pzip:entry>
      
      <!-- _meta/meta.xml-->
      <pzip:entry pzip:target="_meta/meta.xml">
        <!--元数据-->
        <xsl:call-template name="metadata"/>
      </pzip:entry>

      <!-- content.xml-->
      <pzip:entry pzip:target="content.xml">
        <表:电子表格文档_E826>
          <xsl:for-each select="ws:spreadsheets/ws:spreadsheet">
            <xsl:choose>
              <xsl:when test="ws:worksheet">
                <表:工作表_E825>
                  <xsl:attribute name="标识符_E7AC">
                    <xsl:variable name="pos" select="concat('SHEET_',position( ))"/>
                    <xsl:value-of select="$pos"/>
                  </xsl:attribute>
                  <xsl:attribute name="名称_E822">
                    <xsl:value-of select="./@sheetName"/>
                  </xsl:attribute>
                  <xsl:variable name="sheetName">
                    <xsl:value-of select="@sheetName"/>
                  </xsl:variable>
                  <!--隐藏-->
                  <xsl:if test="/ws:spreadsheets/ws:workbook/ws:sheets/ws:sheet[@name=$sheetName]/@state='hidden'">
                    <xsl:attribute name="是否隐藏_E73C">
                      <xsl:value-of select="'true'"/>
                    </xsl:attribute>
                  </xsl:if>
                  <xsl:attribute name="式样引用_E824">
                    <xsl:value-of select="'DEFSTYLE'"/>
                  </xsl:attribute>
                  <!--工作表属性-->
                  <xsl:call-template name="sheetprop"/>
                  <!--工作表内容-->
                  <xsl:call-template name="worksheet"/>
                  <!--筛选-->
                  <xsl:call-template name="sheetfilter"/>
                  <!--分页符集-->
                  <xsl:call-template name="sheetbreak"/>
                </表:工作表_E825 >
              </xsl:when>
              <xsl:when test="ws:chartsheet">
                <表:工作表_E825>
                  <xsl:attribute name="标识符_E7AC">
                    <xsl:variable name="pos" select="position()"/>
                    <xsl:variable name="poss" select="concat('SHEET_',$pos)"/>
                    <xsl:value-of select="$poss"/>
                  </xsl:attribute>
                  <xsl:attribute name="名称_E822">
                    <xsl:value-of select="./@sheetName"/>
                  </xsl:attribute>
                  <xsl:attribute name="式样引用_E824">
                    <xsl:value-of select="'DEFSTYLE'"/>
                  </xsl:attribute>
                  <xsl:call-template name="sheetprop"/>
                  <!--工作表属性-->
                  <xsl:call-template name="worksheet"/>
                  <!--工作表内容-->
                  <!--<xsl:call-template name="table"/>-->
                  <!--这里是新图，做为一张表存在的-->
                  <xsl:call-template name="sheetbreak"/>
                  <!--分页符集-->
                </表:工作表_E825 >
              </xsl:when>
            </xsl:choose>
          </xsl:for-each>
        </表:电子表格文档_E826>
      </pzip:entry>

      <!-- rules.xml-->
      <pzip:entry pzip:target="rules.xml">
        <规则:公用处理规则_B665>
          <规则:长度单位_B666>pt</规则:长度单位_B666>
          <规则:电子表格_B66C>
            <xsl:apply-templates select="ws:spreadsheets/ws:workbook" mode="common"/>
            <!--数据有效性集合-->
            <xsl:if test="//ws:worksheet[ws:dataValidations]">
              <规则:数据有效性集_B618>
                <xsl:for-each select="ws:spreadsheets/ws:spreadsheet">
                  <xsl:if test="ws:worksheet[ws:dataValidations]">
                    <xsl:for-each select="ws:worksheet/ws:dataValidations/ws:dataValidation">
                      <xsl:variable name="p2" select="ancestor::ws:spreadsheet/@sheetName"/>
                      <xsl:call-template name="DataValidation">
                        <xsl:with-param name="sheetid" select="$p2"/>
                      </xsl:call-template>
                    </xsl:for-each>
                  </xsl:if>
                </xsl:for-each>
              </规则:数据有效性集_B618 >
            </xsl:if>
            <!--条件格式化集-->
            <xsl:if test="//ws:worksheet[ws:conditionalFormatting]">
              <规则:条件格式化集_B628>
                <xsl:for-each select="ws:spreadsheets/ws:spreadsheet">
                  <xsl:if test="ws:worksheet[ws:conditionalFormatting]">
                    <xsl:for-each select="ws:worksheet/ws:conditionalFormatting">
                      <xsl:if test="./ws:cfRule/@type='cellIs' or ./ws:cfRule/@type='expression'">
                        <xsl:variable name="p1" select="ancestor::ws:spreadsheet/@sheetName"/>
                        <xsl:call-template name="conFormat">
                          <xsl:with-param name="sheetId" select="$p1"/>
                        </xsl:call-template>
                      </xsl:if>
                    </xsl:for-each>
                  </xsl:if>
                </xsl:for-each>
              </规则:条件格式化集_B628 >
            </xsl:if>
          </规则:电子表格_B66C>
        </规则:公用处理规则_B665>
      </pzip:entry>

      <!--2014-3-11, add by Qihy, 用户自定义数据集 start-->
      <!-- extend.xml -->
      <pzip:entry pzip:target="extend.xml">
        <扩展:扩展区_B200>
          <xsl:if test="//w:document/w:customXML">
            <xsl:copy-of select="//w:document/w:customXML/*"/>
          </xsl:if>
        </扩展:扩展区_B200>
      </pzip:entry>
      <!--2014-3-11 end-->

      <!--euqations.xml-->
      <xsl:if test ="ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/xdr:oneCellAnchor//m:oMathPara">
        <pzip:entry pzip:target="equations.xml">
          <公式:公式集_C200>
            <xsl:for-each select="ws:spreadsheets/ws:spreadsheet">
              <xsl:variable name="sheetName" select="@sheetName"/>
              <xsl:for-each select="ws:Drawings/xdr:wsDr/xdr:oneCellAnchor">
                <xsl:variable name="pos" select="position()"/>
                <xsl:if test =".//m:oMathPara">
                  <公式:数学公式_C201>
                    <xsl:attribute name="标识符_C202">
                      <xsl:value-of select="concat($sheetName,'_equ_',$pos)"/>
                    </xsl:attribute>
                    <xsl:copy-of select=".//m:oMathPara"/>
                  </公式:数学公式_C201>
                </xsl:if>
              </xsl:for-each>
            </xsl:for-each>
          </公式:公式集_C200>
        </pzip:entry>
      </xsl:if>


      <!-- styles.xml-->
      <pzip:entry pzip:target="styles.xml">
        <!-- 式样集-->
        <xsl:call-template name="style"/>
      </pzip:entry>

      <!-- hyperlinks.xml-->
      <xsl:if test="./ws:spreadsheets/ws:spreadsheet/ws:worksheet/ws:hyperlinks">
        <pzip:entry pzip:target="hyperlinks.xml">
          <!--链接集-->
          <xsl:call-template name="Hyperlinks"/>
        </pzip:entry>
      </xsl:if>

      <!-- objectdata.xml-->
      <pzip:entry pzip:target="objectdata.xml">
        <对象:对象数据集_D700>
            <xsl:call-template name="object_root"/>
        </对象:对象数据集_D700>
      </pzip:entry>

      <!-- graphics.xml-->
      <pzip:entry pzip:target="graphics.xml">
        <图形:图形集_7C00>
          <xsl:if test ="./ws:spreadsheets/ws:spreadsheet/ws:comments">
            <xsl:for-each select="./ws:spreadsheets/ws:spreadsheet/ws:comments">
              <xsl:call-template name="pizhuji-root">
                <xsl:with-param name="comment-name" select="."/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>

          <xsl:if test ="./ws:spreadsheets/ws:spreadsheet/ws:Drawings">
 
                <xsl:call-template name ="drawings"/>

          </xsl:if>
        </图形:图形集_7C00>
      </pzip:entry>

      <!-- chart.xml-->
      <pzip:entry pzip:target="chart.xml">
        <图表:图表集_E836>
          <!--<xsl:call-template name ="object_chart"/>-->
          
          <!--图表(零到无穷个)-->
          <xsl:call-template name="table"/>
        </图表:图表集_E836>
      </pzip:entry>

      <!--外挂多媒体对象文件夹 data/xxxxx.jpg或.wav或.mpg等-->
      <!--<pzip:entry pzip:target="data">
        
      </pzip:entry>-->
    </pzip:archive>
  </xsl:template>
	<xsl:template match="ws:workbook" mode="common">
		<!--表:公用处理规则-->
    
      <xsl:if test="ws:workbookPr[@date1904 and @date1904 = 1]">
        <规则:日期系统_B614>1904</规则:日期系统_B614>
      </xsl:if>

      <xsl:if test=".//@showSheetTabs = '0'">
        <规则:是否显示工作表标签_B635>false</规则:是否显示工作表标签_B635>
      </xsl:if>
      <xsl:if test=".//@showHorizontalScroll = '0'">
        <规则:是否显示水平滚动条_B636>false</规则:是否显示水平滚动条_B636>
      </xsl:if>

      <xsl:if test=".//@showVerticalScroll ='0'">
        <规则:是否显示垂直滚动条_B637>false</规则:是否显示垂直滚动条_B637>
      </xsl:if>
      <!--计算设置-->
      <xsl:apply-templates select="ws:calcPr"/>
      <!--数据有效性-->
    
		<!--/表:公用处理规则-->
	</xsl:template>
	<xsl:template match="ws:calcPr">
		<xsl:if test="@iterate = 1 or @iterate = 'true'">
      <规则:计算设置_B615>
        <xsl:attribute name="迭代次数_B616">
          <xsl:if test="@iterateCount">
            <xsl:value-of select="@iterateCount"/>
          </xsl:if>
          <xsl:if test="not(@iterateCount)">
            <xsl:value-of select="100"/>
          </xsl:if>
        </xsl:attribute>
        <xsl:attribute name="偏差值_B617">
          <xsl:if test="@iterateDelta">
            <xsl:value-of select="@iterateDelta"/>
          </xsl:if>
          <xsl:if test="not(@iterateDelta)">
            <xsl:value-of select="0.001"/>
          </xsl:if>
        </xsl:attribute>
      </规则:计算设置_B615>
		</xsl:if>
		<xsl:if test="@fullPrecision">
      <规则:精确度是否以显示值为准_B613>true</规则:精确度是否以显示值为准_B613>
		</xsl:if>
      <xsl:choose>
        <xsl:when test="@refMode='R1C1'">
          <规则:是否RC引用_B634>true</规则:是否RC引用_B634>
        </xsl:when>
        <xsl:otherwise>
          <规则:是否RC引用_B634>false</规则:是否RC引用_B634>
        </xsl:otherwise>
      </xsl:choose>
	</xsl:template>
  <!--单元格引用值转换:OOXML2UOF-->
  <xsl:template name="SqrefTrans">
  </xsl:template>
	<xsl:template match="sqref">
		<xsl:variable name="cell" select="."/>
		<!--B10:C11-->
		<xsl:choose>
			<!--是否要加条件呢-->
			<xsl:when test="contains($cell,':')">
				<xsl:variable name="b" select="substring-before($cell,':')"/>
				<xsl:variable name="a" select="substring-after($cell,':')"/>
				<xsl:variable name="bc" select="translate($b,'0123456789','')"/>
				<xsl:variable name="bn" select="translate($b,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
				<xsl:variable name="ac" select="translate($a,'0123456789','')"/>
				<xsl:variable name="an" select="translate($a,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
				<xsl:value-of select="concat('$',$bc,'$',$bn,':','$',$ac,'$',$an)"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="f" select="translate($cell,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
				<xsl:variable name="s" select="translate($cell,'0123456789','')"/>
				<xsl:variable name="ff" select="concat('$',$s)"/>
				<xsl:variable name="ss" select="concat('$',$f)"/>
				<xsl:value-of select="concat($ff,$ss)"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

  <!--2014-3-18，update by Qihy，增加数据有效性-选中R1C1引用样式时的转换， start-->
  <xsl:template name="sqreftrans1">
    <xsl:param name="sqref"/>
    <!--B10:C11-->
    <xsl:choose>
      <xsl:when test="contains($sqref,' ')">
        <xsl:call-template name="multisqref1">
          <xsl:with-param name="multisqref">
            <xsl:value-of select="$sqref"/>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:when test="contains($sqref,':')">
        <xsl:variable name="b" select="substring-before($sqref,':')"/>
        <xsl:variable name="a" select="substring-after($sqref,':')"/>
        <xsl:variable name="bc" select="translate($b,'0123456789','')"/>
        <xsl:variable name="col1">
          <xsl:call-template name="colseq">
            <xsl:with-param name ="s" select ="$bc"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="bn" select="translate($b,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
        <xsl:variable name="ac" select="translate($a,'0123456789','')"/>
        <xsl:variable name="col2">
          <xsl:call-template name="colseq">
            <xsl:with-param name ="s" select ="$ac"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="an" select="translate($a,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
        <xsl:value-of select="concat('R',$bn,'C',$col1,':','R',$an,'C',$col2)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="f" select="translate($sqref,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
        <xsl:variable name="s" select="translate($sqref,'0123456789','')"/>
        <xsl:variable name="col">
          <xsl:call-template name="colseq">
            <xsl:with-param name ="s" select ="$s"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="ff" select="concat('C',$col)"/>
        <xsl:variable name="ss" select="concat('R',$f)"/>
        <xsl:value-of select="concat($ss,$ff)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="multisqref1">
    <xsl:param name="multisqref"/>
    <xsl:variable name="sqreffirst" select="substring-before($multisqref,' ')"/>
    <xsl:variable name="sqrefleft" select="substring-after($multisqref,' ')"/>
    <xsl:variable name="sqref1">
      <xsl:call-template name="sqreftrans1">
        <xsl:with-param name="sqref">
          <xsl:value-of select="$sqreffirst"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="sqref2">
      <xsl:call-template name="sqreftrans1">
        <xsl:with-param name="sqref">
          <xsl:value-of select="$sqrefleft"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="concat($sqref1,' ',$sqref2)"/>
  </xsl:template>
  <xsl:template name ="colseq">
    <xsl:param name ="s"/>
    <xsl:choose>
      <xsl:when test ="contains($s,'A') or contains($s,'B')  or contains($s,'C')  or contains($s,'D')  or contains($s,'E') or contains($s,'F') or contains($s,'G')  or contains($s,'H')  or contains($s,'I')">
        <xsl:value-of select ="translate($s, 'A,B,C,D,E,F,G,H,I','1,2,3,4,5,6,7,8,9')"/>
      </xsl:when>
      <xsl:when test ="contains($s,'J')">
        <xsl:value-of select ="10 + number(translate($s, 'J','0'))"/>
      </xsl:when>
      <xsl:when test ="contains($s,'K') or contains($s,'L')  or contains($s,'M')  or contains($s,'N')  or contains($s,'O') or contains($s,'P') or contains($s,'Q')  or contains($s,'R')  or contains($s,'S')">
        <xsl:value-of select ="10 + number(translate($s, 'K,L,M,N,O,P,Q,R,S','1,2,3,4,5,6,7,8,9'))"/>
      </xsl:when>
      <xsl:when test ="contains($s,'T')">
        <xsl:value-of select ="20 + number(translate($s, 'T','0'))"/>
      </xsl:when>
      <xsl:when test ="contains($s,'U')  or contains($s,'V')  or contains($s,'W')  or contains($s,'X') or contains($s,'Y') or contains($s,'Z')">
        <xsl:value-of select ="10 + number(translate($s, 'U,V,W,X,Y,Z','1,2,3,4,5,6'))"/>
      </xsl:when>
      <xsl:otherwise/>
    </xsl:choose>
  </xsl:template>
  <!--2014-3-18 end-->
  
	<xsl:template name="sqreftrans">
		<xsl:param name="sqref"/>
		<!--B10:C11-->
		<xsl:choose>
			<xsl:when test="contains($sqref,' ')">
				<xsl:call-template name="multisqref">
					<xsl:with-param name="multisqref">
						<xsl:value-of select="$sqref"/>
					</xsl:with-param>
				</xsl:call-template>
			</xsl:when>
      <xsl:when test="contains($sqref,':')">
        <xsl:variable name="b" select="substring-before($sqref,':')"/>
        <xsl:variable name="a" select="substring-after($sqref,':')"/>
        <xsl:variable name="bc" select="translate($b,'0123456789','')"/>
        <xsl:variable name="bn" select="translate($b,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
        <xsl:variable name="ac" select="translate($a,'0123456789','')"/>
        <xsl:variable name="an" select="translate($a,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
        <xsl:value-of select="concat('$',$bc,'$',$bn,':','$',$ac,'$',$an)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="f" select="translate($sqref,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
        <xsl:variable name="s" select="translate($sqref,'0123456789','')"/>
        <xsl:variable name="ff" select="concat('$',$s)"/>
        <xsl:variable name="ss" select="concat('$',$f)"/>
        <xsl:value-of select="concat($ff,$ss)"/>
      </xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="multisqref">
		<xsl:param name="multisqref"/>
		<xsl:variable name="sqreffirst" select="substring-before($multisqref,' ')"/>
		<xsl:variable name="sqrefleft" select="substring-after($multisqref,' ')"/>
		<xsl:variable name="sqref1">
			<xsl:call-template name="sqreftrans">
				<xsl:with-param name="sqref">
					<xsl:value-of select="$sqreffirst"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:variable>
		<xsl:variable name="sqref2">
			<xsl:call-template name="sqreftrans">
				<xsl:with-param name="sqref">
					<xsl:value-of select="$sqrefleft"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:variable>
		<xsl:value-of select="concat($sqref1,' ',$sqref2)"/>
	</xsl:template>
  
  <!--数据有效性-->
	<xsl:template name="DataValidation">
		<xsl:param name="sheetid"/>
		<!--sheetId:Sheet1-->
		<xsl:variable name="sheetname" select='concat("&apos;",$sheetid,"&apos;")'/>
		<!--sheetname:%工作表1%!-->
		<xsl:variable name="sqref">
			<!--<xsl:apply-templates select="@sqref"/>-->

      <!--2014-3-18,add by Qihy, 增加数据有效性-选中R1C1引用样式时的转换，start-->
      <!--<xsl:call-template name="sqreftrans">
        <xsl:with-param name="sqref" select="./@sqref"/>
      </xsl:call-template>-->
      <xsl:choose>
        <xsl:when test ="ancestor::ws:spreadsheets/ws:workbook/ws:calcPr/@refMode='R1C1'">
          <xsl:call-template name ="sqreftrans1">
            <xsl:with-param name="sqref" select="./@sqref"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="sqreftrans">
            <xsl:with-param name="sqref" select="./@sqref"/>
          </xsl:call-template>
        </xsl:otherwise>  
      </xsl:choose>
     <!--2014-3-18 end-->
            
		</xsl:variable>
    <规则:数据有效性_B619>
      <规则:区域集_B61A>
        <规则:区域_B61B>
          <xsl:value-of select="concat($sheetname,'!',$sqref)"/>
        </规则:区域_B61B>
      </规则:区域集_B61A>
      <规则:校验类型_B61C>
        <xsl:choose>
          <xsl:when test="@type='none'">
            <xsl:value-of select="'any-value'"/>
          </xsl:when>
          <xsl:when test="@type='whole'">
            <xsl:value-of select="'whole-number'"/>
          </xsl:when>
          <xsl:when test="@type='decimal'">
            <xsl:value-of select="'decimal'"/>
          </xsl:when>
          <xsl:when test="@type='list'">
            <xsl:value-of select="'list'"/>
          </xsl:when>
          <xsl:when test="@type='date'">
            <xsl:value-of select="'date'"/>
          </xsl:when>
          <xsl:when test="@type='time'">
            <xsl:value-of select="'time'"/>
          </xsl:when>
          <xsl:when test="@type='textLength'">
            <xsl:value-of select="'text-length'"/>
          </xsl:when>
          <xsl:when test="@type='custom'">
            <xsl:value-of select="'custom'"/>
          </xsl:when>
          <xsl:otherwise >
            <xsl:value-of select ="'any-value'"/>
          </xsl:otherwise>
        </xsl:choose>
      </规则:校验类型_B61C>
      
        <xsl:if test="@operator">
          <规则:操作码_B61D>
          <xsl:choose>
            <xsl:when test="@operator='between'">
              <xsl:value-of select="'between'"/>
            </xsl:when>
            <xsl:when test="@operator='equal'">
              <xsl:value-of select="'equal-to'"/>
            </xsl:when>
            <xsl:when test="@operator='greaterThan'">
              <xsl:value-of select="'greater-than'"/>
            </xsl:when>
            <xsl:when test="@operator='greaterThanOrEqual'">
              <xsl:value-of select="'greater-than-or-equal-to'"/>
            </xsl:when>
            <xsl:when test="@operator='lessThan'">
              <xsl:value-of select="'less-than'"/>
            </xsl:when>
            <xsl:when test="@operator='lessThanOrEqual'">
              <xsl:value-of select="'less-than-or-equal-to'"/>
            </xsl:when>
            <xsl:when test="@operator='notBetween'">
              <xsl:value-of select="'not-between'"/>
            </xsl:when>
            <xsl:when test="@operator='notEqual'">
              <xsl:value-of select="'not-equal-to'"/>
            </xsl:when>
          </xsl:choose>
            </规则:操作码_B61D>
        </xsl:if>
      
        <!--add 2014-1-14 Qihy 修复BUG2939-设置最大值丢失 start-->
        <xsl:if test="not(@operator)">
          <规则:操作码_B61D>
            <xsl:value-of select="'between'"/>
          </规则:操作码_B61D>
        </xsl:if>
        <!--add 2014-1-14 Qihy 修复BUG2939-设置最大值丢失 end-->
      
      <xsl:if test="ws:formula1">
        <规则:第一操作数_B61E>
          <xsl:variable name="ff" select="translate(ws:formula1,'&quot;','')"/>
          <xsl:if test="contains($ff,'$')">
            <xsl:value-of select="concat('=',$ff)"/>
          </xsl:if>
          <xsl:if test="not(contains($ff,'$'))">
            <xsl:value-of select="$ff"/>
          </xsl:if>
        </规则:第一操作数_B61E>
      </xsl:if>
      <xsl:if test="ws:formula2">
        <规则:第二操作数_B61F>
          <xsl:variable name="ff" select="ws:formula2"/>
          <xsl:if test="contains($ff,'$')">
            <xsl:value-of select="concat('=',ws:formula1)"/>
          </xsl:if>
          <xsl:if test="not(contains($ff,'$'))">
            <xsl:value-of select="ws:formula2"/>
          </xsl:if>
        </规则:第二操作数_B61F>
      </xsl:if>
      <xsl:if test ="not(ws:formula1) and not(ws:formula2)">
        <规则:第一操作数_B61E/>
      </xsl:if>
      <xsl:if test="@allowBlank and @allowBlank='1'">
        <规则:是否忽略空格_B620>true</规则:是否忽略空格_B620>
      </xsl:if>
      <xsl:if test="not(@allowBlank) or @allowBlank='0' ">
        <规则:是否忽略空格_B620>false</规则:是否忽略空格_B620>
      </xsl:if>
      <xsl:if test="not(@showDropDown) or @showDropDown='0'">
        <规则:是否显示下拉箭头_B621>true</规则:是否显示下拉箭头_B621>
      </xsl:if>
      <xsl:if test="@showDropDown and @showDropDown='1'">
        <规则:是否显示下拉箭头_B621>false</规则:是否显示下拉箭头_B621>
      </xsl:if>
      <规则:输入提示_B622>
        <xsl:attribute name="是否显示_B623">
          <xsl:choose>
            <xsl:when test="@showInputMessage and @showInputMessage=1">
              <xsl:value-of select="'true'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'false'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="标题_B624">
          <xsl:value-of select="@promptTitle"/>
        </xsl:attribute>
        <xsl:attribute name="内容_B625">
          <xsl:value-of select="@prompt"/>
        </xsl:attribute>
      </规则:输入提示_B622>
      <!-- update by xuzhenwei BUG_2469:数据有效性错误信息提示错误 2013-01-21 start -->
      <规则:错误提示_B626>
        <xsl:attribute name="是否显示_B623">
          <xsl:choose>
            <xsl:when test="@showErrorMessage and @showErrorMessage=1">
              <xsl:value-of select="'true'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'false'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="类型_B627">
          <xsl:choose>
            <xsl:when test="not(@errorStyle)">
              <xsl:value-of select="'stop'"/>
            </xsl:when>
            <xsl:when test="@errorStyle">
              <xsl:value-of select="@errorStyle"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
        <!-- end -->
        <xsl:attribute name="标题_B624">
          <xsl:value-of select="@errorTitle"/>
        </xsl:attribute>
        <xsl:attribute name="内容_B625">
          <xsl:value-of select="@error"/>
        </xsl:attribute>
      </规则:错误提示_B626>
     

    </规则:数据有效性_B619>
	</xsl:template>
  
  <!--条件格式化 20130513 update by xuzhenwei 该部分有待进一步完善 -->
	<xsl:template name="conFormat">
		<xsl:param name="sheetId"/>
		<!--sheetId:Sheet1-->
		<xsl:variable name="sheetname" select='concat("&apos;",$sheetId,"&apos;")'/>
		<!--sheetname:%工作表1%!-->
		<xsl:variable name="sqref">
			<xsl:call-template name="sqreftrans">
				<xsl:with-param name="sqref">
					<xsl:value-of select="@sqref"/>
				</xsl:with-param>
			</xsl:call-template>
		</xsl:variable>

    <规则:条件格式化_B629>
       <规则:区域集_B61A>
        <规则:区域_B61B>
          <xsl:value-of select="concat($sheetname,'!',$sqref)"/>
        </规则:区域_B61B>
      </规则:区域集_B61A>
      
      <xsl:for-each select="ws:cfRule">
        <!--[@type='cellIs']-->
        <规则:条件_B62B>
          <xsl:attribute name="类型_B673">
            <xsl:choose>
              <xsl:when test="./@type='cellIs'">
                <xsl:value-of select="'cell-value'"/>
              </xsl:when>
              <xsl:when test="./@type='expression'">
                <xsl:value-of select="'formula'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'cell value'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:choose>
            <xsl:when test="./@type='cellIs'">
              <规则:操作码_B62C>
                <xsl:choose>
                  <xsl:when test="@operator='between'">
                    <xsl:value-of select="'between'"/>
                  </xsl:when>
                  <xsl:when test="@operator='equal'">
                    <xsl:value-of select="'equal-to'"/>
                  </xsl:when>
                  <xsl:when test="@operator='greaterThan'">
                    <xsl:value-of select="'greater-than'"/>
                  </xsl:when>
                  <xsl:when test="@operator='greaterThanOrEqual'">
                    <xsl:value-of select="'greater-than-or-equal-to'"/>
                  </xsl:when>
                  <xsl:when test="@operator='lessThan'">
                    <xsl:value-of select="'less-than'"/>
                  </xsl:when>
                  <xsl:when test="@operator='lessThanOrEqual'">
                    <xsl:value-of select="'less-than-or-equal-to'"/>
                  </xsl:when>
                  <xsl:when test="@operator='notBetween'">
                    <xsl:value-of select="'not-between'"/>
                  </xsl:when>
                  <xsl:when test="@operator='notEqual'">
                    <xsl:value-of select="'not-equal-to'"/>
                  </xsl:when>
                  <xsl:when test="@operator='beginsWith'">
                    <xsl:value-of select="'start-with'"/>
                  </xsl:when>
                  <xsl:when test="@operator='containsText'">
                    <xsl:value-of select="'contain'"/>
                  </xsl:when>
                  <xsl:when test="@operator='endsWith'">
                    <xsl:value-of select="'end-with'"/>
                  </xsl:when>
                </xsl:choose>
              </规则:操作码_B62C>
            </xsl:when>
            <xsl:when test="./@type='duplicateValues'">
              <规则:操作码_B62C>
                 <xsl:value-of select="'equal-to'"/>
              </规则:操作码_B62C>
            </xsl:when>
            <xsl:when test="./@type='containsText'">
              <规则:操作码_B62C>
              <xsl:choose>
                <xsl:when test="@operator='containsText'">
                  <xsl:value-of select="'contain'"/>
                </xsl:when>
                <!-- 有待进一步完善 -->
              </xsl:choose>
              </规则:操作码_B62C>
            </xsl:when>
            <xsl:when test="./@type='aboveAverage'">
              <!-- 有待进一步完善 -->
            </xsl:when>
            <xsl:when test="./@type='dataBar'">
              <!-- UOF标准更新后，该条件格式化部分功能有待完善 -->
            </xsl:when>
            <xsl:when test="./@type='colorScale'">
              <!-- UOF标准更新后，该条件格式化部分功能有待完善 -->
            </xsl:when>
            <xsl:when test="./@type='iconSet'">
              <!-- UOF标准更新后，该条件格式化部分功能有待完善 -->
            </xsl:when>
            <xsl:otherwise>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:variable name="bishu">
            <xsl:value-of select="count(ws:formula)"/>
          </xsl:variable>
          
          <xsl:if test="not(./@type='expression')">
            <xsl:if test="$bishu='1'">
              <规则:第一操作数_B61E>
                <xsl:value-of select="ws:formula"/>
              </规则:第一操作数_B61E>
            </xsl:if>
            <xsl:if test="$bishu='2'">
              <规则:第一操作数_B61E>
                <xsl:value-of select="ws:formula"/>
              </规则:第一操作数_B61E>
              <xsl:if test="ws:formula[position()=2]">
                <规则:第二操作数_B61F>
                  <xsl:value-of select="ws:formula[last()]"/>
                </规则:第二操作数_B61F>
              </xsl:if>
            </xsl:if>
            <xsl:if test ="$bishu!='1' and $bishu!='2'">
              <规则:第一操作数_B61E/>
            </xsl:if>
          </xsl:if>
          <xsl:if test="./@type='expression'">
            <规则:第一操作数_B61E>
              <xsl:variable name="form">
                <xsl:value-of select="./ws:formula"/>
              </xsl:variable>
              <xsl:value-of select="concat('=',$form)"/>
            </规则:第一操作数_B61E>
          </xsl:if>
          <规则:格式_B62D>
            <xsl:attribute name="式样引用_B62E">
              <xsl:if test="./@dxfId">
                <xsl:value-of select="concat('tjgsh_',@dxfId+1)"/>
              </xsl:if>
              <xsl:if test="not(./@dxfId)">
                <xsl:value-of select="'tjgsh_1'"/>
              </xsl:if>
            </xsl:attribute>
          </规则:格式_B62D>
        </规则:条件_B62B>
      </xsl:for-each>
    </规则:条件格式化_B629>
    <!--
    <xsl:if test="
            ws:cfRule[@operator='between' or @operator='equal' or @operator='greaterThan' 
            or @operator='greaterThanOrEqual' or @operator='lessThan' or @operator='lessThanOrEqual' 
            or @operator='notBetween' or @operator='notEqual' or @operator='beginsWith' 
            or @operator='containsText' or @operator='endsWith']"></xsl:if>
            -->
	</xsl:template>
</xsl:stylesheet>
