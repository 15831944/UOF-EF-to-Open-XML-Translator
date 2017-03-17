<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
                xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
                xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:图表="http://schemas.uof.org/cn/2009/chart"
                xmlns:式样="http://schemas.uof.org/cn/2009/styles"
                xmlns:pzip="urn:u2o:xmlns:post-processings:special"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

  <!--Modified by LDM in 2011/01/03-->
  <!--Need Consideration-->
  <xsl:template name="ValAxisUnit">
    <xsl:if test="./图表:刻度_E71D/图表:主单位_E721">
      <c:majorUnit>
        <xsl:attribute name="val">
          <xsl:value-of select="./图表:刻度_E71D/图表:主单位_E721"/>
        </xsl:attribute>
      </c:majorUnit>
    </xsl:if>
    <xsl:if test="./图表:刻度_E71D/图表:次单位_E722">
      <c:minorUnit>
        <xsl:attribute name="val">
          <xsl:value-of select="./图表:刻度_E71D/图表:次单位_E722"/>
        </xsl:attribute>
      </c:minorUnit>
    </xsl:if>
    <c:dispUnits>
      <xsl:if test="./图表:刻度_E71D/图表:显示单位_E724">
        <c:builtInUnit>
          <xsl:attribute name="val">
            <xsl:choose>
              <xsl:when test="./图表:刻度_E71D/图表:显示单位_E724/@类型_E728='hundreds'">
                <xsl:value-of select="'hundreds'"/>
              </xsl:when>
              <xsl:when test="./图表:刻度_E71D/图表:显示单位_E724/@类型_E728='thousands'">
                <xsl:value-of select="'thousands'"/>
              </xsl:when>
              <xsl:when test="./图表:刻度_E71D/图表:显示单位_E724/@类型_E728='ten thousands'">
                <xsl:value-of select="'tenThousands'"/>
              </xsl:when>
              <xsl:when test="./图表:刻度_E71D/图表:显示单位_E724/@类型_E728='one hundred thousands'">
                <xsl:value-of select="'hundredThousands'"/>
              </xsl:when>
              <xsl:when test="./图表:刻度_E71D/图表:显示单位_E724/@类型_E728/text()='millons'">
                <xsl:value-of select="'millions'"/>
              </xsl:when>
              <xsl:when test="./图表:刻度_E71D/图表:显示单位_E724/@类型_E728='ten millons'">
                <xsl:value-of select="'tenMillions'"></xsl:value-of>
              </xsl:when>
              <xsl:when test="./图表:刻度_E71D/图表:显示单位_E724/@类型_E728='one hundred millons'">
                <xsl:value-of select="'hundredMillions'"></xsl:value-of>
              </xsl:when>
              <xsl:when test="./图表:刻度_E71D/图表:显示单位_E724/@类型_E728='trillions'">
                <xsl:value-of select="'trillions'"></xsl:value-of>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </c:builtInUnit>
      </xsl:if>
		<!--20130507 update by gaoyuwei_qhy bug_2880 图表单位转换后丢失 即工作表“刻度”中，"显示单位：百，图表中包含显示单位"转换不正确-->
		<xsl:if test="./图表:刻度_E71D/图表:显示单位_E724[@是否显示_E727='true']">
			<!--end-->
        <c:dispUnitsLbl>
          <c:layout/>
        </c:dispUnitsLbl>
      </xsl:if>
    </c:dispUnits>
  </xsl:template>
  <xsl:template name="shuzhizhou_kedu2" match="表:数值轴/表:刻度2">
    <xsl:if test="表:数值轴/表:刻度/表:主单位">
      <c:majorUnit>
        <xsl:attribute name="val">
          <xsl:value-of select="表:数值轴/表:刻度/表:主单位"/>
        </xsl:attribute>
      </c:majorUnit>
    </xsl:if>
    <xsl:if test="表:数值轴/表:刻度/表:次单位">
      <c:minorUnit>
        <xsl:attribute name="val">
          <xsl:value-of select="表:数值轴/表:刻度/表:次单位"/>
        </xsl:attribute>
      </c:minorUnit>
    </xsl:if>
    <c:dispUnits>
      <xsl:if test="表:数值轴/表:刻度/表:单位">
        <c:builtInUnit>
          <xsl:attribute name="val">
            <xsl:choose>
              <xsl:when test="表:数值轴/表:刻度/表:单位='hundreds'">
                <xsl:value-of select="'hundreds'"/>
              </xsl:when>
              <xsl:when test="表:数值轴/表:刻度/表:单位='thousands'">
                <xsl:value-of select="'thousands'"/>
              </xsl:when>
              <xsl:when test="表:数值轴/表:刻度/表:单位='ten thousands'">
                <xsl:value-of select="'tenThousands'"/>
              </xsl:when>
              <xsl:when test="表:数值轴/表:刻度/表:单位='one hundred thousands'">
                <xsl:value-of select="'hundredThousands'"/>
              </xsl:when>
              <xsl:when test="表:数值轴/表:刻度/表:单位/text()='millons'">
                <xsl:value-of select="'millions'"/>
              </xsl:when>
              <xsl:when test="表:数值轴/表:刻度/表:单位='ten millons'">
                <xsl:value-of select="'tenMillions'"></xsl:value-of>
              </xsl:when>
              <xsl:when test="表:数值轴/表:刻度/表:单位='one hundred millons'">
                <xsl:value-of select="'hundredMillions'"></xsl:value-of>
              </xsl:when>
              <xsl:when test="表:数值轴/表:刻度/表:单位='trillions'">
                <xsl:value-of select="'trillions'"></xsl:value-of>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </c:builtInUnit>
      </xsl:if>
      <xsl:if test="表:数值轴/表:刻度/表:显示单位[@表:值='true']">
        <!--如果uof显示单位flase,ooxml就没有这项-->
        <c:dispUnitsLbl>
          <c:layout/>
        </c:dispUnitsLbl>
      </xsl:if>
    </c:dispUnits>
  </xsl:template>

  <!--计算值-->
  <xsl:template name="caculateValue">
    <xsl:param name="valValue"/>
    <xsl:variable name="resultValue">
      <xsl:choose>
        <!--系列名是引用的一系列单元格-->
        <xsl:when test="contains($valValue,'!') and contains($valValue,':')">
          <xsl:call-template name="GetSeriesName">
            <xsl:with-param name="refValue" select="$valValue"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$valValue"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:value-of select="$resultValue"/>
  </xsl:template>
  <!--数据源的分类名和系列值-->
  <xsl:template name="CatSer">
    
    <!--<xsl:if test="./@分类名_E776">
      <c:cat>
        <c:strRef>
          <c:f>
            <xsl:variable name="catName">
              <xsl:value-of select="./@分类名_E776"/>
            </xsl:variable>
            <xsl:choose>
              -->
    <!--系列名是引用的一系列单元格-->
    <!--
              <xsl:when test="contains($catName,'!') and contains($catName,':')">
                <xsl:call-template name="GetSeriesName">
                  <xsl:with-param name="refValue" select="$catName"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$catName"/>
              </xsl:otherwise>
            </xsl:choose>
          </c:f>
        </c:strRef>
      </c:cat>
    </xsl:if>-->
    <xsl:if test="./@分类名_E776">
      <c:cat>
        <c:strRef>
          <c:f>
            <xsl:call-template name="caculateValue">
              <xsl:with-param name="valValue" select="./@分类名_E776"/>
            </xsl:call-template>
          </c:f>
        </c:strRef>
      </c:cat>
    </xsl:if>
    <xsl:if test="./@值_E775">      
      <xsl:choose>
        <!--散点图-->
        <xsl:when test="./@类型_E75D = 'scatter' or ./@类型_E75D = 'bubble'">
          <xsl:if test="./@分类名_E776">
            <c:xVal>
              <c:numRef>
                <c:f>
                  <xsl:call-template name="caculateValue">
                    <xsl:with-param name="valValue" select="./@分类名_E776"/>
                  </xsl:call-template>
                </c:f>
              </c:numRef>
            </c:xVal>
          </xsl:if>
          
          <c:yVal>
            <c:numRef>
              <c:f>
                <xsl:call-template name="caculateValue">
                  <xsl:with-param name="valValue" select="./@值_E775"/>
                </xsl:call-template>
              </c:f>
            </c:numRef>
          </c:yVal>
          <xsl:if test="./@类型_E75D = 'bubble'">
            <c:bubbleSize>
              <c:numRef>
                <c:f>
                  <xsl:call-template name="caculateValue">
                    <xsl:with-param name="valValue" select="./图表:气泡大小区域_E83F"/>
                  </xsl:call-template>
                </c:f>
              </c:numRef>
            </c:bubbleSize>
            <c:bubble3D val="0"/>
          </xsl:if>
        </xsl:when>
        <xsl:otherwise>
          <c:val>
            <c:numRef>
              <c:f>
                <xsl:call-template name="caculateValue">
                  <xsl:with-param name="valValue" select="./@值_E775"/>
                </xsl:call-template>
              </c:f>
            </c:numRef>
          </c:val>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="contains(./@子类型_E777,'smooth')">
          <c:smooth val="1"/>
        </xsl:when>
        <xsl:otherwise>
          <c:smooth val="0"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    
  </xsl:template>

  <!--Modified by LDM in 2010/12/18-->
  <!--网格线模板-->
  <xsl:template match="图表:网格线_E734">
    <xsl:if test="./@位置_E735='major'">
      <c:majorGridlines>
        <c:spPr>
          <xsl:call-template name="Line"/>
        </c:spPr>
      </c:majorGridlines>
    </xsl:if>
    <xsl:if test="./@位置_E735='minor'">
      <c:minorGridlines>
        <c:spPr>
          <xsl:call-template name="Line"/>
        </c:spPr>
      </c:minorGridlines>
    </xsl:if>
  </xsl:template>

  <!--Modified by LDM in 2010/12/18-->
  <!--标题模板-->
  <xsl:template match="图表:标题_E736">
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr>
            <xsl:if test="图表:对齐_E726">
              <xsl:call-template name="Alignment"/>
            </xsl:if>
          </a:bodyPr>
          <a:p>
            <a:pPr>
              <a:defRPr/>
            </a:pPr>
            <a:r>
              <a:rPr>
                <xsl:call-template name="Font"/>
              </a:rPr>
              <xsl:if test="./@名称_E742">
                <a:t>
                  <xsl:value-of select="./@名称_E742"/>
                </a:t>
              </xsl:if>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
     
        <xsl:if test="./图表:位置_E70A/@x_C606">
          <xsl:variable name="plotareaPos_X">
            <xsl:value-of select="./ancestor::图表:图表_E837/图表:绘图区_E747/图表:位置_E70A/@x_C606"/>
          </xsl:variable>
          <xsl:variable name="plotareaPos_Y">
            <xsl:value-of select="./ancestor::图表:图表_E837/图表:绘图区_E747/图表:位置_E70A/@y_C607"/>
          </xsl:variable>
          <c:layout>
                <xsl:call-template name="Layout">
                  <xsl:with-param name="off_X" select="$plotareaPos_X"/>
                  <xsl:with-param name="off_Y" select="$plotareaPos_Y"/>
                </xsl:call-template>
            </c:layout> 
        </xsl:if>
      <c:spPr>
        <xsl:if test="图表:填充_E746">
          <xsl:call-template name="Fill"/>
        </xsl:if>
        <xsl:if test="图表:边框线_4226">
          <xsl:call-template name="Border"/>
        </xsl:if>
      </c:spPr>
      <c:overlay val="0"/>
    </c:title>
  </xsl:template>

  <!--Modified by LDM in 2010/12/15-->
  <!--文字排列方向 uof to ooxml-->
  <xsl:template name ="directionMap">
    <xsl:variable name="uofDir">
      <xsl:value-of select="."/>
    </xsl:variable>
    <xsl:choose>
      <!--横排-->
      <xsl:when test="$uofDir = 't2b-l2r-0e-0w'">
        <xsl:value-of select="'horz'"/>
      </xsl:when>
      <!--竖排-->
      <xsl:when test="$uofDir = 'r2l-t2b-0e-90w'">
        <xsl:value-of select="'eaVert'"/>
      </xsl:when>
      <!--所有文字旋转270度-->
      <xsl:when test="$uofDir = 'l2r-b2t-270e-270w'">
        <xsl:value-of select="'vert270'"/>
      </xsl:when>
      <!--所有文字旋转90度-->
      <xsl:when test="$uofDir = 'r2l-t2b-90e-90w'">
        <xsl:value-of select="'vert'"/>
      </xsl:when>
      <!--缺省为横排-->
      <xsl:otherwise>
        <xsl:value-of select="'horz'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--Modified by LDM in 2010/12/15-->
  <!--对齐模板-->
  <!--Alignment-->
  <xsl:template name="Alignment">
    <!--文字排列方向-->
    <xsl:variable name="textDir">
      <xsl:choose>
        <xsl:when test="./图表:对齐_E730/图表:文字排列方向_E703">
          <xsl:for-each select="./图表:对齐_E730/图表:文字排列方向_E703">
            <xsl:call-template name="directionMap"/>
          </xsl:for-each>
        </xsl:when>
        <xsl:when test="./图表:对齐_E726/表:文字排列方向_E703">
          <xsl:for-each select="./图表:对齐_E726/表:文字排列方向_E703">
            <xsl:call-template name="directionMap"/>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'horz'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <!--水平对齐方式-->
    <xsl:variable name="alignHorz">
      <xsl:value-of select="./图表:对齐_E726/表:水平对齐方式_E700"/>
    </xsl:variable>
    <!--垂直对齐方式-->
    <xsl:variable name="alignVert">
      <xsl:value-of select="./图表:对齐_E726/表:垂直对齐方式_E701"/>
    </xsl:variable>
    <!--文字旋转角度-->
    <!--OOXML中只有文字排列方向为横排时，才可以设置旋转角度。竖排或者其它方向排列时旋转角度rot只能为0-->
    
    <xsl:if test="./图表:对齐_E726/表:文字旋转角度_E704 or ./图表:对齐_E730/图表:文字旋转角度_E704">
      <xsl:attribute name="rot">
        <xsl:choose>
          <xsl:when test="$textDir = 'horz'">
            <xsl:choose>
              <xsl:when test="./图表:对齐_E726/表:文字旋转角度_E704">
                <xsl:value-of select="(./图表:对齐_E726/表:文字旋转角度_E704)*(-60000) "/>
              </xsl:when>
              <xsl:when test="./图表:对齐_E730/图表:文字旋转角度_E704">
                <xsl:value-of select="(./图表:对齐_E730/图表:文字旋转角度_E704)*(-60000) "/>
              </xsl:when>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'0'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </xsl:if>
    
    <!--文字排列方向-->
   
      <xsl:attribute name="vert">
        <xsl:value-of select="$textDir"/>
      </xsl:attribute>

   
    <!--文字对齐方式-->
    <xsl:choose>
      <!--文字方向为横排的时候，相应的垂直对齐方式-->
      <xsl:when test="$textDir = 'horz'">
        <xsl:choose>
          <xsl:when test="$alignVert='top'">
            <xsl:attribute name="anchor">
              <xsl:value-of select="'t'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="$alignVert='center'">
            <xsl:attribute name="anchor">
              <xsl:value-of select="'ctr'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="$alignVert='bottom'">
            <xsl:attribute name="anchor">
              <xsl:value-of select="'b'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="$alignVert='justify'">
            <xsl:attribute name="anchor">
              <xsl:value-of select="'just'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="$alignVert='distributed'">
            <xsl:attribute name="anchor">
              <xsl:value-of select="'dist'"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </xsl:when>
      <!--文字方向非横排的时候，相应的水平对齐方式-->
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$alignHorz='left'">
            <xsl:attribute name="anchor">
              <xsl:value-of select="'b'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="$alignHorz='center'">
            <xsl:attribute name="anchor">
              <xsl:value-of select="'ctr'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="$alignHorz='right'">
            <xsl:attribute name="anchor">
              <xsl:value-of select="'t'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="$alignHorz='justify'">
            <xsl:attribute name="anchor">
              <xsl:value-of select="'just'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:when test="$alignHorz='distributed'">
            <xsl:attribute name="anchor">
              <xsl:value-of select="'dist'"/>
            </xsl:attribute>
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--Modified by LDM in 2011/01/03-->
  <!--偏移量-->
  <!--Offset-->
  <xsl:template name="Offset">
    <c:lblOffset>
      <xsl:attribute name="val">
        <xsl:value-of select="./图表:对齐_E730/图表:偏移量_E732"/>
      </xsl:attribute>
    </c:lblOffset>
  </xsl:template>

  <!--Modified by LDM in 2010/12/17-->
  <!--Layout-->
  <xsl:template name="Layout">
    <xsl:param name="off_X"/>
    <xsl:param name="off_Y"/>

    <!--<xsl:variable name="chartPos_x">
      <xsl:value-of select="./ancestor::图表:图表_E837/图表:绘图区_E747/图表:位置_E70A/@x_C606"/>
    </xsl:variable>
    <xsl:variable name="chartPos_y">
      <xsl:value-of select="./ancestor::图表:图表_E837/图表:绘图区_E747/图表:位置_E70A/@y_C607"/>
    </xsl:variable>
    <xsl:variable name="chartPos_h">
      <xsl:value-of select="./ancestor::图表:图表_E837/图表:绘图区_E747/图表:大小_E748/@长_C604"/>
    </xsl:variable>
    <xsl:variable name="chartPos_w">
      <xsl:value-of select="./ancestor::图表:图表_E837/图表:绘图区_E747/图表:大小_E748/@宽_C605"/>
    </xsl:variable>-->
    <xsl:variable name="chartPos_x">
      <xsl:value-of select="//uof:锚点_C644[@图形引用_C62E=//图:图形_8062[图:图表数据引用_8065=./ancestor::图表:图表_E837/@标识符_E828]/@标识符_804B]/uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108"/>
    </xsl:variable>
    <xsl:variable name="chartPos_y">
      <xsl:value-of select="//uof:锚点_C644[@图形引用_C62E=//图:图形_8062[图:图表数据引用_8065=./ancestor::图表:图表_E837/@标识符_E828]/@标识符_804B]/uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108"/>
    </xsl:variable>
    <xsl:variable name="chartPos_h">
      <xsl:value-of select="//uof:锚点_C644[@图形引用_C62E=//图:图形_8062[图:图表数据引用_8065=./ancestor::图表:图表_E837/@标识符_E828]/@标识符_804B]/uof:大小_C621/@长_C604"/>
    </xsl:variable>
    <xsl:variable name="chartPos_w">
      <xsl:value-of select="//uof:锚点_C644[@图形引用_C62E=//图:图形_8062[图:图表数据引用_8065=./ancestor::图表:图表_E837/@标识符_E828]/@标识符_804B]/uof:大小_C621/@宽_C605"/>
    </xsl:variable>
    
    <c:manualLayout>
      <xsl:if test="./图表:位置_E70A/@x_C606">
        <xsl:if test ="not(name()='图表:图例_E794') and not(name()='图表:标题_E736')">
          <c:layoutTarget val="inner"/>
        </xsl:if>
        <c:xMode val="edge"/>
        <c:yMode val="edge"/>
        <c:x>
          <xsl:attribute name="val">
            <xsl:variable name="xCoord">
              <xsl:value-of select="./图表:位置_E70A/@x_C606"/>
            </xsl:variable>
            <!-- 20130416 update by xuzhenwei BUG_2830 互操作 oo-uof（编辑）-oo 024实用文档-损益表(1).xlsx 文档需要修复 start -->
            <xsl:choose>
              <xsl:when test="$chartPos_w">
                <xsl:value-of select="$xCoord"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="($xCoord + $off_X) div $chartPos_w"/>
              </xsl:otherwise>
            </xsl:choose>
            <!-- end -->
          </xsl:attribute>
        </c:x>
      </xsl:if>
      <xsl:if test="./图表:位置_E70A/@y_C607">
        <c:y>
          <xsl:attribute name="val">
            <xsl:variable name="yCoord">
              <xsl:value-of select="./图表:位置_E70A/@y_C607"/>
            </xsl:variable>
            <!-- 20130416 update by xuzhenwei BUG_2830 互操作 oo-uof（编辑）-oo 024实用文档-损益表(1).xlsx 文档需要修复 start -->
            <xsl:choose>
              <xsl:when test="$chartPos_h">
                <xsl:value-of select="$yCoord"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="(($yCoord + $off_Y) div $chartPos_h)+0.1"/>
              </xsl:otherwise>
            </xsl:choose>
            <!-- end -->
          </xsl:attribute>
        </c:y>
      </xsl:if>
      <xsl:if test="./图表:大小_E748/@宽_C605">
        <c:w>
          <xsl:attribute name="val">
            <xsl:variable name="width">
              <xsl:value-of select="./图表:大小_E748/@宽_C605"/>
            </xsl:variable>
            <!-- 20130416 update by xuzhenwei BUG_2830 互操作 oo-uof（编辑）-oo 024实用文档-损益表(1).xlsx 文档需要修复 start -->
            <xsl:choose>
              <xsl:when test="$chartPos_w">
                <xsl:value-of select="$width"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$width div $chartPos_w"/>
              </xsl:otherwise>
            </xsl:choose>
            <!-- end -->
          </xsl:attribute>
        </c:w>
      </xsl:if>
      <xsl:if test="./图表:大小_E748/@长_C604">
        <c:h>
          <xsl:attribute name="val">
            <xsl:variable name="height">
              <xsl:value-of select="./图表:大小_E748/@长_C604"/>
            </xsl:variable>
            <!-- 20130416 update by xuzhenwei BUG_2830 互操作 oo-uof（编辑）-oo 024实用文档-损益表(1).xlsx 文档需要修复 start -->
            <xsl:choose>
              <xsl:when test="$chartPos_h">
                <xsl:value-of select="$height"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$height div $chartPos_h"/>
              </xsl:otherwise>
            </xsl:choose>
            <!-- end --> 
          </xsl:attribute>
        </c:h>
      </xsl:if>
    </c:manualLayout>
  </xsl:template>


  <!--Modified by LDM in 2011/01/03-->
  <!--图表类型-->
  <!--ChartType-->
  <xsl:template name="ChartType">
    <xsl:choose>
      <!-- add by xuzhenwei 判断三维图表类型方式的修改:通过uof中的<图表:透视深度_E783>标签进行判断是否是三维效果图，这里取默认的三维效果设置 start -->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:透视深度_E783">
        <!--zl 20150107 start-->
        <c:view3D>
          <c:rotX val="15"/>
          <c:rotY val="20"/>
          <c:rAngAx val="1"/>
        </c:view3D>
        <!--zl 20150107 end-->
      </xsl:when>
      <!-- end 可能以下的代码就没什么作用了，先保留着 -->
      <!--三维簇状柱形图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='clustered-3d']">
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--三维堆积柱形图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='stacked-3d']">
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--三维百分比堆积柱形图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='percent-3d']">
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--三维柱形图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='3d']">
        <!--
        <c:view3D>
          <c:rAngAx val="1"/>
        </c:view3D>
        -->
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--柱形簇状圆柱图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='clustered-cylinder']">
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
        <!--
        <c:sideWall>
          <c:spPr>
            <a:solidFill>
              <a:schemeClr val="tx2">
                <a:lumMod val="20000"/>
                <a:lumOff val="80000"/>
              </a:schemeClr>
            </a:solidFill>
            <a:ln w="22225" cap="flat" cmpd="dbl">
              <a:solidFill>
                <a:srgbClr val="7030A0"/>
              </a:solidFill>
              <a:prstDash val="dash"/>
            </a:ln>
          </c:spPr>
        </c:sideWall>
        <c:backWall>
          <c:spPr>
            <a:solidFill>
              <a:schemeClr val="tx2">
                <a:lumMod val="20000"/>
                <a:lumOff val="80000"/>
              </a:schemeClr>
            </a:solidFill>
            <a:ln w="22225" cap="flat" cmpd="dbl">
              <a:solidFill>
                <a:srgbClr val="7030A0"/>
              </a:solidFill>
              <a:prstDash val="dash"/>
            </a:ln>
          </c:spPr>
        </c:backWall>
        -->
      </xsl:when>

      <!--堆积柱形圆柱图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='stacked-cylinder']">
        <c:view3D>
          <c:rAngAx val="1"/>
        </c:view3D>
      </xsl:when>

      <!--百分比堆积柱形圆柱图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='percent-cylinder']">
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--三维柱形圆柱图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='3D-cylinder']">
        <!--
        <c:view3D>
          <c:rAngAx val="1"/>
        </c:view3D>
        -->
        <c:view3D>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--柱形圆锥图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='clustered-cone']">
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--堆积柱形圆锥图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='stacked-cone']">
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--百分比堆积柱形圆锥图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='percent-cone']">
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--三维柱形圆锥图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='3D-cone']">
        <!--
        <c:view3D>
          <c:rAngAx val="1"/>
        </c:view3D>
        -->
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--柱形棱锥图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='clustered-pyramid']">
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--堆积柱形棱锥图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='stacked-pyramid']">
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--百分比堆积柱形棱锥图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='percent-pyramid']">
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--三维柱形棱锥图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='3D-pyramid']">
        <!--
        <c:view3D>
          <c:rAngAx val="1"/>
        </c:view3D>
        -->
        <c:view3D>
          <c:rAngAx val="0"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>

      <!--三维饼图-->
      <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='standard-3d']">
        <c:view3D>
          <c:rotX val="30"/>
          <c:perspective val="30"/>
        </c:view3D>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--Modified by LDM in 2011/01/03-->
  <!--图表区转换模板-->
  <!--PlotArea-->
  <xsl:template name ="PlotArea">
    <xsl:variable name ="seriesNo">
      <xsl:value-of select="position()"/>
    </xsl:variable>
    <c:plotArea>
      <xsl:if test="图表:绘图区_E747">
        <xsl:apply-templates select="图表:绘图区_E747"/>
      </xsl:if>
        <!--柱形图-->
        <!--簇状柱形图-->
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='clustered']">
          <c:barChart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='clustered']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:barChart>
        </xsl:if>
        <!--堆积柱形图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='stacked']">
          <c:barChart>
            <c:barDir val="col"/>
            <c:grouping val="stacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='stacked']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:overlap val="100"/>
            <!-- 20130325 add by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 start -->
            <c:gapWidth val="150"/>
            <!-- end -->
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:barChart>
        </xsl:if>
        <!--百分比堆积柱形图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='percent-stacked']">
          <c:barChart>
            <c:barDir val="col"/>
            <c:grouping val="percentStacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='percent-stacked']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:overlap val="100"/>
            <!-- 20130325 add by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 start -->
            <c:gapWidth val="150"/>
            <!-- end -->
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:barChart>
        </xsl:if>
        <!--散点图造成的bug-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='marker']">
          <c:barChart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='marker']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:barChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='line-marker']">
          <c:barChart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='line-marker']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:barChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='standard']">
          <c:barChart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='standard']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:barChart>
        </xsl:if>
        <!--三维簇状柱形图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='clustered-3d']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='clustered-3d']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:gapWidth val="150"/>
            <c:shape val="box"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:bar3DChart>
        </xsl:if>
        <!--三维堆积柱形图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='exploded-3d']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="stacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='exploded-3d']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:overlap val="100"/>
            <c:gapWidth val="150"/>
            <c:shape val="box"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:bar3DChart>
        </xsl:if>
        <!--三维百分比堆积柱形图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='percent-stacked-marker']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="percentStacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='percent-stacked-marker']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:overlap val="100"/>
            <c:gapWidth val="150"/>
            <c:shape val="box"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:bar3DChart>
        </xsl:if>
        <!--三维柱形图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='standard-3d']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="standard"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='standard-3d']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="box"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:bar3DChart>
        </xsl:if>
        <!--柱形簇状圆柱图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='clustered']">
          <c:barChart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='clustered']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="cylinder"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:barChart>
        </xsl:if>
        <!--堆积柱形圆柱图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='stacked']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="stacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='stacked']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:overlap val="100"/>
            <!-- 20130325 add by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 start -->
            <c:gapWidth val="150"/>
            <!-- end -->
            <c:shape val="cylinder"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:bar3DChart>
        </xsl:if>
        <!--百分比堆积柱形圆柱图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='percent-stacked']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="percentStacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='percent-stacked']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="cylinder"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:bar3DChart>
        </xsl:if>
        <!--三维柱形圆柱图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='standard-3d']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="standard"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='standard-3d']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="cylinder"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:bar3DChart>
        </xsl:if>
        <!--柱形圆锥图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='clustered-cone']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='clustered-cone']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="cone"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:bar3DChart>
        </xsl:if>
        <!--堆积柱形圆锥图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='stacked-cone']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="stacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='stacked-cone']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="cone"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:bar3DChart>
        </xsl:if>
        <!--百分比堆积柱形圆锥图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='percent-cone']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="percentStacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='percent-cone']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="cone"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:bar3DChart>
        </xsl:if>
        <!--三维柱形圆锥图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='3D-cone']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="standard"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='3D-cone']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="cone"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:bar3DChart>
        </xsl:if>
        <!--柱形棱锥图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='clustered-pyramid']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='clustered-pyramid']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="pyramid"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:bar3DChart>
        </xsl:if>
        <!--堆积柱形圆锥图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='stacked-pyramid']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="stacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='stacked-pyramid']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="pyramid"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:bar3DChart>
        </xsl:if>
        <!--百分比堆积柱形圆锥图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='percent-pyramid']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="percentStacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='percent-pyramid']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="pyramid"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:bar3DChart>
        </xsl:if>
        <!--三维柱形圆锥图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='3D-pyramid']">
          <c:bar3DChart>
            <c:barDir val="col"/>
            <c:grouping val="standard"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='column' and @子类型_E777='3D-pyramid']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="pyramid"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
            <c:axId val="222222222"/>
          </c:bar3DChart>
        </xsl:if>
        <!--条形图-->
        <!--簇状条形图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='clustered']">
          <c:barChart>
            <c:barDir val="bar"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='clustered']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:barChart>
        </xsl:if>
		
	<!--zhaolin 2015-4-14-->
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='clustered-cone']">
        <c:barChart>
          <c:barDir val="bar"/>
          <c:grouping val="clustered"/>
          <c:varyColors val="0"/>
          <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='clustered-cone']">
            <xsl:call-template name="SeriesTrans"/>
          </xsl:for-each>
          <c:axId val="000000000"/>
          <c:axId val="111111111"/>
        </c:barChart>
      </xsl:if>
      <!--zhaolin 2015-4-14-->
		
        <!--堆积条形图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='stacked']">
          <c:barChart>
            <c:barDir val="bar"/>
            <c:grouping val="stacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='stacked']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:overlap val="100"/>
            <!-- 20130325 add by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 start -->
            <c:gapWidth val="150"/>
            <!-- end -->
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:barChart>
        </xsl:if>
		
	<!--zhaolin 2015-4-13-->
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='stacked-cone']">
        <c:barChart>
          <c:barDir val="bar"/>
          <c:grouping val="stacked"/>
          <c:varyColors val="0"/>
          <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='stacked-cone']">
            <xsl:call-template name="SeriesTrans"/>
          </xsl:for-each>
          <c:axId val="000000000"/>
          <c:axId val="111111111"/>
        </c:barChart>
      </xsl:if>
      <!--zhaolin 2015-4-13-->
		
        <!--百分比堆积条形图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='percent-stacked']">
          <c:barChart>
            <c:barDir val="bar"/>
            <c:grouping val="percentStacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='percent-stacked']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:overlap val="100"/>
            <!-- 20130325 add by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 start -->
            <c:gapWidth val="150"/>
            <!-- end -->
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:barChart>
        </xsl:if>
		
	<!--zhaolin 2015-4-13-->
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='percent-cone']">
        <c:barChart>
          <c:barDir val="bar"/>
          <c:grouping val="percentStacked"/>
          <c:varyColors val="0"/>
          <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='percent-cone']">
            <xsl:call-template name="SeriesTrans"/>
          </xsl:for-each>
          <c:overlap val="100"/>
          <!-- 20130325 add by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 start -->
          <c:gapWidth val="150"/>
          <!-- end -->
          <c:axId val="000000000"/>
          <c:axId val="111111111"/>
        </c:barChart>
      </xsl:if>
      <!--zhaolin 2015-4-13-->
		
        <!--三维簇状条形图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='clustered-3d']">
          <c:bar3DChart>
            <c:barDir val="bar"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='clustered-3d']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="box"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:bar3DChart>
        </xsl:if>
		
	<!--zhaolin 2015-4-13-->
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='clustered-pyramid']">
        <c:bar3DChart>
          <c:barDir val="bar"/>
          <c:grouping val="clustered"/>
          <c:varyColors val="0"/>
          <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='clustered-pyramid']">
            <xsl:call-template name="SeriesTrans"/>
          </xsl:for-each>
          <c:shape val="box"/>
          <c:axId val="000000000"/>
          <c:axId val="111111111"/>
        </c:bar3DChart>
      </xsl:if>
      <!--zhaolin 2015-4-13-->
		
        <!--三维堆积条形图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='exploded-3d']">
          <c:bar3DChart>
            <c:barDir val="bar"/>
            <c:grouping val="stacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='exploded-3d']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:overlap val="100"/>
            <!-- 20130325 add by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 start -->
            <c:gapWidth val="150"/>
            <!-- end -->
            <c:shape val="box"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:bar3DChart>
        </xsl:if>
		
	<!--zhaolin 2015-4-13-->
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='stacked-pyramid']">
        <c:bar3DChart>
          <c:barDir val="bar"/>
          <c:grouping val="stacked"/>
          <c:varyColors val="0"/>
          <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='stacked-pyramid']">
            <xsl:call-template name="SeriesTrans"/>
          </xsl:for-each>
          <c:overlap val="100"/>
          <!-- 20130325 add by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 start -->
          <c:gapWidth val="150"/>
          <!-- end -->
          <c:shape val="box"/>
          <c:axId val="000000000"/>
          <c:axId val="111111111"/>
        </c:bar3DChart>
      </xsl:if>
      <!--zhaolin 2015-4-13-->
		
        <!--三维百分比堆积条形图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='percent-stacked-marker']">
          <c:bar3DChart>
            <c:barDir val="bar"/>
            <c:grouping val="percentStacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='percent-stacked-marker']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:overlap val="100"/>
            <!-- 20130325 add by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 start -->
            <c:gapWidth val="150"/>
            <!-- end -->
            <c:shape val="box"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:bar3DChart>
        </xsl:if>
		
	<!--zhaolin 2015-4-13-->
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='percent-pyramid']">
        <c:bar3DChart>
          <c:barDir val="bar"/>
          <c:grouping val="percentStacked"/>
          <c:varyColors val="0"/>
          <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bar' and @子类型_E777='percent-pyramid']">
            <xsl:call-template name="SeriesTrans"/>
          </xsl:for-each>
          <c:overlap val="100"/>
          <!-- 20130325 add by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 start -->
          <c:gapWidth val="150"/>
          <!-- end -->
          <c:shape val="box"/>
          <c:axId val="000000000"/>
          <c:axId val="111111111"/>
        </c:bar3DChart>
      </xsl:if>
      <!--zhaolin 2015-4-13-->
		
        <!--簇状水平圆柱图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='clustered-3d']">
          <c:bar3DChart>
            <c:barDir val="bar"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='clustered-3d']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="cylinder"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:bar3DChart>
        </xsl:if>
        <!--堆积水平圆柱图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='exploded-3d']">
          <c:bar3DChart>
            <c:barDir val="bar"/>
            <c:grouping val="stacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='exploded-3d']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:shape val="cylinder"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:bar3DChart>
        </xsl:if>
        <!--百分比堆积水平圆柱图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='percent-stacked-marker']">
          <c:bar3DChart>
            <c:barDir val="bar"/>
            <c:grouping val="percentStacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='surface' and @子类型_E777='percent-stacked-marker']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:overlap val="100"/>
            <!-- 20130325 add by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 start -->
            <c:gapWidth val="150"/>
            <!-- end -->
            <c:shape val="cylinder"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:bar3DChart>
        </xsl:if>
        <!--散点图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='scatter' and @子类型_E777='marker']">
          <c:scatterChart>
            <c:scatterStyle val="lineMarker"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='scatter' and @子类型_E777='marker']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:scatterChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='scatter' and @子类型_E777='smooth-marker']">
          <c:scatterChart>
            <c:scatterStyle val="smoothMarker"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='scatter' and @子类型_E777='smooth-marker']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:scatterChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='scatter' and @子类型_E777='smooth']">
          <c:scatterChart>
            <c:scatterStyle val="smoothMarker"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='scatter' and @子类型_E777='smooth']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:scatterChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='scatter' and @子类型_E777='line-marker']">
          <c:scatterChart>
            <c:scatterStyle val="lineMarker"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='scatter' and @子类型_E777='line-marker']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:scatterChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='scatter' and @子类型_E777='line']">
          <c:scatterChart>
            <c:scatterStyle val="lineMarker"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='scatter' and @子类型_E777='line']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:scatterChart>
        </xsl:if>
        <!--圆环图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='doughnut' and (@子类型_E777='standard' or @子类型_E777='exploded')]">
          <c:doughnutChart>
            <c:varyColors val="1"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='doughnut' and (@子类型_E777='standard' or @子类型_E777='exploded')]">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            
            <!--20130123 gaoyuwei bug_2638_1 圆环图转换后类型不明-->
			  <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:第一扇区起始角_E78A">
				  <c:firstSliceAng>
					  <xsl:attribute name="val">
						  <xsl:value-of select ="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:第一扇区起始角_E78A"/>
					  </xsl:attribute>
				  </c:firstSliceAng>			  
			  </xsl:if>		 
			  <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:圆环图内径大小_E788">
				  <c:holeSize>
				  <xsl:attribute name="val">
					  <xsl:value-of select ="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:圆环图内径大小_E788"/>
				  </xsl:attribute>  
				  </c:holeSize>				  
			  </xsl:if>
			<!--end-->
            
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:doughnutChart>
        </xsl:if>
        <!--雷达图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='radar' and (@子类型_E777='standard' or @子类型_E777='standard-marker')]">
          <c:radarChart>
            <c:radarStyle val="marker"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='radar' and (@子类型_E777='standard' or @子类型_E777='standard-marker')]">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:radarChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='radar' and @子类型_E777='fill']">
          <c:radarChart>
            <c:radarStyle val="filled"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='radar' and @子类型_E777='fill']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:radarChart>
        </xsl:if>
        <!--气泡图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bubble' and @子类型_E777='standard']">
          <c:bubbleChart>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bubble' and @子类型_E777='standard']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
			
			<!--2015-3-25 zhaolin-->
            <c:bubbleScale val="100"/>
            <c:showNegBubbles val="1"/>
            <!--2015-3-25 zhaolin-->
			
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:bubbleChart>
        </xsl:if>
        <!--面积图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='area' and @子类型_E777='standard']">
          <c:areaChart>
            <c:grouping val="standard"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='area' and @子类型_E777='standard']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:areaChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='area' and @子类型_E777='stacked']">
          <c:areaChart>
            <c:grouping val="stacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='area' and @子类型_E777='stacked']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:areaChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='area' and @子类型_E777='percent-stacked']">
          <c:areaChart>
            <c:grouping val="percentStacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='area' and @子类型_E777='percent-stacked']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:areaChart>
        </xsl:if>
        <!--折线图-->
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and @子类型_E777='clustered']">
          <c:lineChart>
            <c:grouping val="standard"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and @子类型_E777='clustered']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:marker val="1"/>
            <!-- 20130419 update by xuzhenwei BUG_2830 互操作 oo-uof（编辑）-oo 024实用文档-损益表(1).xlsx 文档需要修复 start -->
            <xsl:variable name="formatN" select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791/图表:数值_E70D/@格式码_E73F"/>
            <xsl:choose>
              <xsl:when test="contains($formatN,'yyyy')">
                <c:smooth val="0"/>
                <c:axId val="223499392"/>
                <c:axId val="223502720"/>
              </xsl:when>
              <xsl:otherwise>
                <c:axId val="000000000"/>
                <c:axId val="111111111"/>
              </xsl:otherwise>
            </xsl:choose>
            <!-- end -->
          </c:lineChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and @子类型_E777='stacked']">
          <c:lineChart>
            <c:grouping val="stacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and @子类型_E777='stacked']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:marker val="1"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:lineChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and @子类型_E777='percent-stacked']">
          <c:lineChart>
            <c:grouping val="percentStacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and @子类型_E777='percent-stacked']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:marker val="1"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:lineChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and @子类型_E777='standard-marker']">
          <c:lineChart>
            <c:grouping val="standard"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and @子类型_E777='standard-marker']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:marker val="1"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:lineChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and @子类型_E777='stacked-marker']">
          <c:lineChart>
            <c:grouping val="stacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and @子类型_E777='stacked-marker']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:marker val="1"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:lineChart>
        </xsl:if>
        <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and @子类型_E777='percent-stacked-marker']">
          <c:lineChart>
            <c:grouping val="percentStacked"/>
            <c:varyColors val="0"/>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and @子类型_E777='percent-stacked-marker']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:marker val="1"/>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:lineChart>
        </xsl:if>
        <!--圆饼图-->
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='standard']">
          <c:pieChart>

            <!--2014-6-8 update by Qihy, 饼图按扇区着色不正确， start-->
            <!--<c:varyColors val="1"/>-->
            <c:varyColors>
              <xsl:attribute name="val">
                <xsl:choose>
                  <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:是否依数据点着色_E77A = 'false'">
                    <xsl:value-of select="0"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </c:varyColors>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='standard']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:firstSliceAng val="0"/>
          </c:pieChart>
        </xsl:if>
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='standard-3d']">
          <c:pie3DChart>
            <!--<c:varyColors val="1"/>-->
            <c:varyColors>
              <xsl:attribute name="val">
                <xsl:choose>
                  <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:是否依数据点着色_E77A = 'false'">
                    <xsl:value-of select="0"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </c:varyColors>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='standard-3d']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
          </c:pie3DChart>
        </xsl:if>
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='of-pie']">
          <c:ofPieChart>
            <c:ofPieType val="pie"/>
            <!--<c:varyColors val="1"/>-->
            <c:varyColors>
              <xsl:attribute name="val">
                <xsl:choose>
                  <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:是否依数据点着色_E77A = 'false'">
                    <xsl:value-of select="0"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </c:varyColors>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='of-pie']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:gapWidth val="100"/>
            <c:secondPieSize val="75"/>
            <c:serLines/>
          </c:ofPieChart>
        </xsl:if>
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='exploded']">
          <c:pieChart>
            <!--<c:varyColors val="1"/>-->
            <c:varyColors>
              <xsl:attribute name="val">
                <xsl:choose>
                  <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:是否依数据点着色_E77A = 'false'">
                    <xsl:value-of select="0"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </c:varyColors>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='exploded']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:firstSliceAng val="0"/>
          </c:pieChart>
        </xsl:if>
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='exploded-3d']">
          <c:pie3DChart>
            <!--<c:varyColors val="1"/>-->
            <c:varyColors>
              <xsl:attribute name="val">
                <xsl:choose>
                  <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:是否依数据点着色_E77A = 'false'">
                    <xsl:value-of select="0"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </c:varyColors>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='exploded-3d']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
          </c:pie3DChart>
        </xsl:if>
      <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='of-bar']">
          <c:ofPieChart>
            <c:ofPieType val="bar"/>
            <!--<c:varyColors val="1"/>-->
            <c:varyColors>
              <xsl:attribute name="val">
                <xsl:choose>
                  <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:是否依数据点着色_E77A = 'false'">
                    <xsl:value-of select="0"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="1"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </c:varyColors>
            <!--2014-6-8 end-->
            
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='pie' and @子类型_E777='of-bar']">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:gapWidth val="100"/>
            <c:secondPieSize val="75"/>
            <c:serLines/>
          </c:ofPieChart>
        </xsl:if>
               
       <!--20130122 gaoyuwei bug_2640 转换后文件打开有提示信息 实现方式不同 算是差异 start--> 
        <!--股价图?暂不能转-->
      
        <!--<xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and ancestor-or-self::图表:组_E74D/图表:高低线_E77D]">
          <c:stockChart>
            <xsl:for-each select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line' and ancestor-or-self::图表:组_E74D/图表:高低线_E77D]">
              <xsl:call-template name="SeriesTrans"/>
            </xsl:for-each>
            <c:hiLowLines>
              <c:spPr>
                <xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:高低线_E77D/图表:边框线_4226">
                  <xsl:apply-templates select="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:高低线_E77D/图表:边框线_4226"/>
                </xsl:if>
              </c:spPr>
            </c:hiLowLines>
            <c:axId val="000000000"/>
            <c:axId val="111111111"/>
          </c:stockChart>
        </xsl:if>-->
         <!--end-->

      <!--饼图没有网格线、坐标轴和数据表-->
      <!--Modified by LDM in 2010/12/18-->
      
     <!--20130124 gaoyuwei bug 2638_2 圆环图转换后类型不明 start-->
		<!--圆环没有网格线、坐标轴和数据表-->
		<xsl:if test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D!='pie' and @类型_E75D!='doughnut']">
        <!--end-->
        
        <xsl:choose>
          <xsl:when test="./图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='bubble' or @类型_E75D='scatter']">
            <c:valAx>
              <c:axId val="000000000"/>
              <c:delete val="0"/>
              <xsl:if test ="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']">
                <xsl:for-each select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[1]">
                  <xsl:call-template name="数值轴"/>
                </xsl:for-each>
              </xsl:if>
              
              <!--2014-4-23, update by Qihy, 修复bug3080中的uof-oo横轴标题丢失问题， start-->
              <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category' or @子类型_E793='value']/图表:标题_E736">
              <!--2014-4-23 end-->
                
                <xsl:apply-templates select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[1]/图表:标题_E736"/>
              </xsl:if>
              <!-- 20130322 add by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 start -->
              <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:字体_E70B">
                <c:txPr>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:pPr>
                      <a:defRPr>
                        <xsl:attribute name="sz">
                          <xsl:choose>
                            <xsl:when test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:字体_E70B/字:字体_4128[@字号_412D]">
                              <xsl:value-of select="floor((./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:字体_E70B/字:字体_4128/@字号_412D) * 100)"/>
                            </xsl:when>
                            <!--默认字号大小为10-->
                            <xsl:otherwise>
                              <xsl:value-of select="'1000'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <!--bold-->
                        <!--<xsl:value-of select="./图表:字体_E70B/字:字体_4128/@颜色_412F"/>-->
                        <xsl:attribute name="b">
                          <xsl:choose>
                            <xsl:when test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:字体_E70B[字:是否粗体_4130='true']">
                              <xsl:value-of select="'1'"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'0'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <!--italic-->
                        <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:字体_E70B[字:是否斜体_4131='true']">
                          <xsl:attribute name="i">
                            <xsl:value-of select="'1'"/>
                          </xsl:attribute>
                        </xsl:if>

                        <!--文本框中的字符串下划线类型-->
                        <!--Not Finished-->
                        <!--Modified by LDM in 2010/12/29-->
                        <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:字体_E70B/字:下划线_4136[@线型_4137!='none']">
                          <xsl:attribute name="u">
                            <xsl:variable name="lineType">
                              <xsl:value-of select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:字体_E70B/字:下划线_4136/@线型_4137"/>
                            </xsl:variable>
                            <xsl:variable name="dashType">
                              <xsl:value-of select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:字体_E70B/字:下划线_4136/@虚实_4138"/>
                            </xsl:variable>
                            <xsl:call-template name="UnderlineTypeMapping">
                              <xsl:with-param name="lineType" select="$lineType"/>
                              <xsl:with-param name="dashType" select="$dashType"/>
                            </xsl:call-template>
                          </xsl:attribute>
                        </xsl:if>

                        <!--Need modify of the rest of this template-->
                        <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:字体_E70B/字:删除线_4135">
                          <xsl:choose>
                            <xsl:when test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:字体_E70B[字:删除线_4135='none']">
                              <xsl:attribute name="strike">
                                <xsl:value-of select="'noStrike'"/>
                              </xsl:attribute>
                            </xsl:when>
                            <xsl:when test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:字体_E70B[字:删除线_4135='single']">
                              <xsl:attribute name="strike">
                                <xsl:value-of select="'sngStrike'"/>
                              </xsl:attribute>
                            </xsl:when>
                            <xsl:when test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:字体_E70B[字:删除线_4135='double']">
                              <xsl:attribute name="strike">
                                <xsl:value-of select="'dblStrike'"/>
                              </xsl:attribute>
                            </xsl:when>
                          </xsl:choose>
                        </xsl:if>
                      </a:defRPr>
                    </a:pPr>
                    <a:endParaRPr lang="zh-CN"/>
                  </a:p>
                </c:txPr>
              </xsl:if>
              <!-- end -->
              <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:网格线集_E733/图表:网格线_E734">
                <xsl:apply-templates select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[1]/图表:网格线集_E733/图表:网格线_E734"/>
              </xsl:if>
            </c:valAx>
            <c:valAx>
              <c:axId val="111111111"/>
              <c:delete val="0"/>
              <xsl:if test ="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']">
                <xsl:for-each select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[2]">
                  <xsl:call-template name="数值轴"/>
                </xsl:for-each>
              </xsl:if>
              <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:标题_E736">
                <xsl:apply-templates select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[2]/图表:标题_E736"/>
              </xsl:if>

              <!--2014-4-22, delete by Qihy, 修复bug3209，打开需要恢复，生成重复的c:txPr元素， start-->
              <!-- 20130322 add by xuzhenwei BUG_#2494:字体颜色，下划线效果丢失 start -->
              <!--<xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:字体_E70B">
                <c:txPr>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:pPr>
                      <a:defRPr>
                        <xsl:attribute name="sz">
                          <xsl:choose>
                            <xsl:when test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:字体_E70B/字:字体_4128[@字号_412D]">
                              <xsl:value-of select="floor((./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:字体_E70B/字:字体_4128/@字号_412D) * 100)"/>
                            </xsl:when>
                            --><!--默认字号大小为10--><!--
                            <xsl:otherwise>
                              <xsl:value-of select="'1000'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        --><!--bold--><!--
                        --><!--<xsl:value-of select="./图表:字体_E70B/字:字体_4128/@颜色_412F"/>--><!--
                        <xsl:attribute name="b">
                          <xsl:choose>
                            <xsl:when test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:字体_E70B[字:是否粗体_4130='true']">
                              <xsl:value-of select="'1'"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'0'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        --><!--italic--><!--
                        <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:字体_E70B[字:是否斜体_4131='true']">
                          <xsl:attribute name="i">
                            <xsl:value-of select="'1'"/>
                          </xsl:attribute>
                        </xsl:if>

                        --><!--文本框中的字符串下划线类型--><!--
                        --><!--Not Finished--><!--
                        --><!--Modified by LDM in 2010/12/29--><!--
                        <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:字体_E70B/字:下划线_4136[@线型_4137!='none']">
                          <xsl:attribute name="u">
                            <xsl:variable name="lineType">
                              <xsl:value-of select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:字体_E70B/字:下划线_4136/@线型_4137"/>
                            </xsl:variable>
                            <xsl:variable name="dashType">
                              <xsl:value-of select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:字体_E70B/字:下划线_4136/@虚实_4138"/>
                            </xsl:variable>
                            <xsl:call-template name="UnderlineTypeMapping">
                              <xsl:with-param name="lineType" select="$lineType"/>
                              <xsl:with-param name="dashType" select="$dashType"/>
                            </xsl:call-template>
                          </xsl:attribute>
                        </xsl:if>

                        --><!--Need modify of the rest of this template--><!--
                        <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:字体_E70B/字:删除线_4135">
                          <xsl:choose>
                            <xsl:when test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:字体_E70B[字:删除线_4135='none']">
                              <xsl:attribute name="strike">
                                <xsl:value-of select="'noStrike'"/>
                              </xsl:attribute>
                            </xsl:when>
                            <xsl:when test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:字体_E70B[字:删除线_4135='single']">
                              <xsl:attribute name="strike">
                                <xsl:value-of select="'sngStrike'"/>
                              </xsl:attribute>
                            </xsl:when>
                            <xsl:when test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:字体_E70B[字:删除线_4135='double']">
                              <xsl:attribute name="strike">
                                <xsl:value-of select="'dblStrike'"/>
                              </xsl:attribute>
                            </xsl:when>
                          </xsl:choose>
                        </xsl:if>
                      </a:defRPr>
                    </a:pPr>
                    <a:endParaRPr lang="zh-CN"/>
                  </a:p>
                </c:txPr>
              </xsl:if>-->
              <!-- end -->
              <!--2014-2-22 end-->
              
              <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:网格线集_E733/图表:网格线_E734">
                <xsl:apply-templates select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[2]/图表:网格线集_E733/图表:网格线_E734"/>
              </xsl:if>
            </c:valAx>
          </xsl:when>
          <xsl:otherwise>
            <c:catAx>
              <!-- 20130416 update by xuzhenwei BUG_2830 互操作 oo-uof（编辑）-oo 024实用文档-损益表(1).xlsx 文档需要修复 start -->
              <xsl:variable name="formatN" select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791/图表:数值_E70D/@格式码_E73F"/>
              <xsl:choose>
                <xsl:when test="contains($formatN,'yyyy')">
                  <c:axId val="223499392"/>
                </xsl:when>
                <xsl:otherwise>
                  <c:axId val="000000000"/>
                </xsl:otherwise>
              </xsl:choose>
              <!-- end -->
              <c:delete val="0"/>
              <xsl:if test ="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='category']">
                <xsl:for-each select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='category']">
                  <xsl:call-template name="分类轴"/>
                </xsl:for-each>
              </xsl:if>
              <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:标题_E736">
                <xsl:apply-templates select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:标题_E736"/>
              </xsl:if>
              <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:网格线集_E733/图表:网格线_E734">
                <xsl:apply-templates select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='series' or @子类型_E793='category']/图表:网格线集_E733/图表:网格线_E734"/>
              </xsl:if>
			  
			        <!--zl  2015-5-19-->
              <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791/图表:边框线_4226/@线型_C60D='single' and ./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791/图表:边框线_4226/@虚实_C60E='solid' and not(./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791/图表:边框线_4226/@宽度_C60F)">
                <c:majorGridlines/>
              </xsl:if>
              <!--zl  2015-5-19-->
			  
            </c:catAx>
            <c:valAx>
              <!-- 20130416 update by xuzhenwei BUG_2830 互操作 oo-uof（编辑）-oo 024实用文档-损益表(1).xlsx 文档需要修复 start -->
              <!-- <c:axId val="111111111"/> -->
              <xsl:variable name="formatN" select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791/图表:数值_E70D/@格式码_E73F"/>
              <xsl:choose>
                <xsl:when test="contains($formatN,'yyyy')">
                  <c:axId val="223502720"/>
                </xsl:when>
                <xsl:otherwise>
                  <c:axId val="111111111"/>
                </xsl:otherwise>
              </xsl:choose>
              <!-- end -->
              <c:delete val="0"/>
              <xsl:if test ="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']">
                <!-- 20130506 update by xuzhenwei BUG_2911 第二轮回归 uof-oo 功能 图表 文档打开需要修复 start -->
                <xsl:for-each select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[2]">
                  <xsl:call-template name="数值轴"/>
                </xsl:for-each>
               <!-- end -->
              </xsl:if>
              <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:标题_E736">
                <xsl:apply-templates select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:标题_E736"/>
              </xsl:if>
              <xsl:if test="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:网格线集_E733/图表:网格线_E734">
                <xsl:apply-templates select="./图表:绘图区_E747/图表:坐标轴集_E790/图表:坐标轴_E791[@子类型_E793='value']/图表:网格线集_E733/图表:网格线_E734"/>
              </xsl:if>
            </c:valAx>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>

      <!--数据表-->
      <!--Modified by LDM in 2010/01/03-->
      <xsl:if test="./图表:数据表_E79B">
        <xsl:apply-templates select="./图表:数据表_E79B"/>
      </xsl:if>
    </c:plotArea>
  </xsl:template>

  <xsl:template match="图表:边框线_4226">
    <xsl:call-template name="Border"/>
  </xsl:template>
  <!--Modified by LDM in 2011/01/03-->
  <!--SeriesTrans-->
  <!--图表系列数据源转换模板-->
  <xsl:template name="SeriesTrans">
    <xsl:variable name="seriesNo">
      <!--<xsl:number count="图表:数据系列_E74F" level="single"/>-->
      <xsl:value-of select="./@标识符_E778"/>
    </xsl:variable>
    <c:ser>
      <xsl:if test="./图表:趋势线集_E762">
        <c:trendline>
          <c:trendlineType val="exp"/>
          <c:dispRSqr val="0"/>
          <c:dispEq val="0"/>
        </c:trendline>
      </xsl:if>
      <c:idx>
        <xsl:attribute name="val">
          <xsl:value-of select="$seriesNo"/>
        </xsl:attribute>
      </c:idx>
      <c:order>
        <xsl:attribute name="val">
          <xsl:value-of select="$seriesNo"/>
        </xsl:attribute>
      </c:order>
      <!--分离型饼图-->

      <xsl:if test="contains(./@子类型_E777,'exploded')">
        <c:explosion val="25"/>
      </xsl:if>      
      <xsl:choose>
        <xsl:when test="./@类型_E75D = 'line' and (./@子类型_E777 = 'clustered' or ./@子类型_E777 = 'stacked' or ./@子类型_E777 = 'percent-stacked')">
          <c:marker>
            <c:symbol val="none"/>
          </c:marker>
        </xsl:when>
        <xsl:when test="./@类型_E75D = 'scatter' and (./@子类型_E777 = 'smooth' or ./@子类型_E777 = 'line' )">
          <c:marker>
            <c:symbol val="none"/>
          </c:marker>
        </xsl:when>
        <xsl:when test="./@类型_E75D = 'radar' and ./@子类型_E777 = 'standard'">
          <c:marker>
            <c:symbol val="none"/>
          </c:marker>
        </xsl:when>
      </xsl:choose>
      
      
      <xsl:variable name="seriesNameValue">
        <xsl:if test="./@名称_E774">
          <xsl:value-of select="./@名称_E774"/>
        </xsl:if>
        <xsl:if test="not(./@名称_E774)">
          <xsl:value-of select="concat('系列',$seriesNo)"/>
        </xsl:if>
      </xsl:variable>
      <xsl:choose>
        <!--系列名是引用的一系列单元格-->
        <xsl:when test="contains($seriesNameValue,'!') and contains($seriesNameValue,':')">
          <c:tx>
            <c:strRef>
              <c:f>
                <xsl:call-template name="GetSeriesName">
                  <xsl:with-param name="refValue" select="$seriesNameValue"/>
                </xsl:call-template>
              </c:f>
            </c:strRef>
          </c:tx>
        </xsl:when>
        <xsl:when test="contains($seriesNameValue,'!') and not(contains($seriesNameValue,':'))">
          <c:tx>
            <c:strRef>
              <c:f>
                <xsl:call-template name="GetSeriesName">
                  <xsl:with-param name="refValue" select="$seriesNameValue"/>
                </xsl:call-template>
              </c:f>
            </c:strRef>
          </c:tx>
        </xsl:when>
        <!--系列名是普通的字符串值-->
        <xsl:otherwise>
          <c:tx>
            <c:v>
              <xsl:value-of select="translate($seriesNameValue,'&quot;','')"/>
            </c:v>
          </c:tx>
        </xsl:otherwise>
      </xsl:choose>

		<!--20130121 gaoyuwei bug2645_2 工作表“刻度”中，分类轴Y轴刻度相关设置转换后不正确 start-->
		<c:invertIfNegative val="0"/>
		<!--end-->

      <!--  <xsl:call-template name="GetSeriesName"/>-->
      <xsl:call-template name="CatSer"/>
        <xsl:if test="./图表:边框线_4226 or ./图表:填充_E746">
          <!--填充-->
          <c:spPr>
            <xsl:if test="图表:填充_E746">
              <xsl:call-template name="Fill"/>
            </xsl:if>
            <!--边框-->
            <xsl:if test="图表:边框线_4226">
              <xsl:call-template name="Border"/>
            </xsl:if>

            <!--20130122 gaoyuwei BUG_2649&2639 散点图转换前后不一致  start 20130322再修正-->
            <xsl:if test="not(图表:边框线_4226) and not(//图表:绘图区_E747/图表:图表类型组集_E74C/图表:组_E74D/图表:数据系列集_E74E/图表:数据系列_E74F[@类型_E75D='line'])">
              <a:ln w="28575" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main">
                <a:noFill/>
              </a:ln>
            </xsl:if>
            <!--end-->
            
          </c:spPr>
          <!--数据系列显示标志-->
          <xsl:if test="./图表:数据标签_E752">
            <c:dLbls>
              <c:txPr>
                <a:bodyPr/>
                <a:p>
                  <a:pPr>
                    
                    <!--2014-4-23, update by Qihy, 默认字体大小不正确， start-->
                    <a:defRPr sz="1000">
                    <!--2014-4-23 end-->
                      
                      <a:latin typeface="Times New Roman"/>
                      <a:ea typeface="永中宋体"/>
                    </a:defRPr>
                  </a:pPr>
                </a:p>
              </c:txPr>
              <xsl:call-template name="SeriesDisplay"/>
            </c:dLbls>
          </xsl:if>
        </xsl:if>
      <!-- add by xuzhenwei BUG_2506:数据点集边框线渐进色转为单色线 2013-01-19 start -->
      <!-- 图表:数据点集 -->
      <xsl:if test="./图表:数据点集_E755">
        <xsl:variable name="id" select="-1"/>
        <xsl:for-each select="./图表:数据点集_E755//图表:数据点_E756">
          <c:dPt>
            <c:idx>
              <xsl:attribute name="val">
                <xsl:value-of select="$id+1"/>
              </xsl:attribute>
            </c:idx>
            <c:invertIfNegative val="0"/>
            <c:bubble3D val="0"/>
            <c:spPr>
                <xsl:if test="./图表:边框线_4226">
                  <xsl:call-template name="Border"/>
                </xsl:if>
            </c:spPr>
          </c:dPt>
        </xsl:for-each>
      </xsl:if>
      <!-- end -->
  
    </c:ser>
  </xsl:template>
  <!--数据点集显示标志模板-->
  <!--not finished-->
  <xsl:template name="datasetDisplay">
    <c:dLbl>
      <c:txPr>
        <a:bodyPr/>
        <a:p>
          <a:pPr>
            <a:defRPr sz="1200">
              <a:latin typeface="Times New Roman"/>
              <a:ea typeface="永中宋体"/>
            </a:defRPr>
          </a:pPr>
        </a:p>
      </c:txPr>
      <xsl:if test="@表:点">
        <c:idx>
          <xsl:attribute name="val">
            <xsl:value-of select="(@表:点)-1"/>
          </xsl:attribute>
        </c:idx>
      </xsl:if>
      <xsl:if test="表:显示标志/@表:图例标志='true'">
        <c:showLegendKey>
          <xsl:attribute name="val">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </c:showLegendKey>
      </xsl:if>
      <xsl:if test="表:显示标志/@表:数值='true'">
        <c:showVal>
          <xsl:attribute name="val">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </c:showVal>
      </xsl:if>
      <xsl:if test="表:显示标志/@表:类别名='true'">
        <c:showCatName>
          <xsl:attribute name="val">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </c:showCatName>
      </xsl:if>
      <xsl:if test="表:显示标志/@表:系列名='true'">
        <c:showSerName>
          <xsl:attribute name="val">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </c:showSerName>
      </xsl:if>
      <xsl:if test="表:显示标志/@表:分隔符">
        <!--按照枚举处理的,但实际不是枚举-->
        <c:separator>
          <xsl:choose>
            <xsl:when test="表:显示标志/@表:分隔符='0'">
              <xsl:value-of select="' '"/>
            </xsl:when>
            <xsl:when test="表:显示标志/@表:分隔符='1'">
              <xsl:value-of select="','"/>
            </xsl:when>
            <xsl:when test="表:显示标志/@表:分隔符='2'">
              <xsl:value-of select="';'"/>
            </xsl:when>
            <xsl:when test="表:显示标志/@表:分隔符='3'">
              <xsl:value-of select="'.'"/>
            </xsl:when>
          </xsl:choose>
        </c:separator>
      </xsl:if>
      <xsl:if test="表:显示标志/@表:百分数='true'">
        <!--没找到在哪里设置,所以不知道在ooxml中的位置是否正确-->
        <c:showPercent>
          <xsl:attribute name="val">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </c:showPercent>
      </xsl:if>
    </c:dLbl>
  </xsl:template>

  <!--Modified by LDM in 2010/12/18-->
  <!--数据系列显示模板-->
  <xsl:template name="SeriesDisplay">
    <c:showLegendKey>
      <xsl:choose>
        <xsl:when test ="图表:数据标签_E752/@是否显示图例标志_E719='true'">
          <xsl:attribute name="val">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="val">
            <xsl:value-of select="'0'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </c:showLegendKey>
    <c:showVal>
      <xsl:choose>
        <xsl:when test ="图表:数据标签_E752/@是否显示数值_E717='true'">
          <xsl:attribute name="val">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:when>
		
		        
        <!--zhaolin-->
        <xsl:when test ="图表:数据标签_E752/图表:数值_E753/@是否链接到源_E73E='true'">
          <xsl:attribute name="val">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:when>
        <!--zhaolin-->
		
        <xsl:otherwise>
          <xsl:attribute name="val">
            <xsl:value-of select="'0'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </c:showVal>
    <c:showCatName>
      <xsl:choose>
        <xsl:when test ="图表:数据标签_E752/@是否显示类别名_E716='true'">
          <xsl:attribute name="val">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="val">
            <xsl:value-of select="'0'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </c:showCatName>
    <c:showSerName>
      <xsl:choose>
        <xsl:when test ="图表:数据标签_E752/@是否显示系列名_E715='true'">
          <xsl:attribute name="val">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="val">
            <xsl:value-of select="'0'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </c:showSerName>

    <c:showPercent>
      <xsl:choose>
        <xsl:when test ="图表:数据标签_E752/@是否百分数图表_E718='true'">
          <xsl:attribute name="val">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="val">
            <xsl:value-of select="'0'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </c:showPercent>
    <c:showBubbleSize val="0">
      <xsl:choose>
        <xsl:when test ="图表:数据标签_E752/@气泡尺寸_E71B='true'">
          <xsl:attribute name="val">
            <xsl:value-of select="'1'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="val">
            <xsl:value-of select="'0'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
    </c:showBubbleSize>
    <xsl:if test="图表:数据标签_E752/@分隔符_E71A">
      <c:separator>
        <xsl:choose>
          <xsl:when test="图表:数据标签_E752/@分隔符_E71A='0'">
            <xsl:value-of select="' '"/>
          </xsl:when>
          <xsl:when test="图表:数据标签_E752/@分隔符_E71A='1'">
            <xsl:value-of select="','"/>
          </xsl:when>
          <xsl:when test="图表:数据标签_E752/@分隔符_E71A='2'">
            <xsl:value-of select="';'"/>
          </xsl:when>
          <xsl:when test="图表:数据标签_E752/@分隔符_E71A='3'">
            <xsl:value-of select="'.'"/>
          </xsl:when>
        </xsl:choose>
      </c:separator>
    </xsl:if>
    
  </xsl:template>

  <!--Modified by LDM in 2010/12/17-->
  <!--表填充-->
  <!--Fill-->
  <xsl:template name="Fill">
    <!--
    <xsl:if test ="not(表:填充) or 表:填充/图:颜色='auto'">
      <a:noFill/>
    </xsl:if>
    -->
    <xsl:if test="./图表:填充_E746/图:颜色_8004 and ./图表:填充_E746/图:颜色_8004!='auto'">
      <a:solidFill>
        <a:srgbClr>
          <xsl:attribute name="val">
            <xsl:value-of select="substring-after(./图表:填充_E746/图:颜色_8004,'#')"/>
          </xsl:attribute>
        </a:srgbClr>
      </a:solidFill>
    </xsl:if>
    <!--渐变-->
    <xsl:if test="图表:填充_E746/图:渐变_800D">
      <xsl:apply-templates select="图表:填充_E746/图:渐变_800D"/>
    </xsl:if>
    <!--图案填充-->
    <xsl:if test="图表:填充_E746/图:图案_800A">
      <xsl:call-template name="PatternFill"/>
    </xsl:if>

    <!--图片和纹理填充-->
    <xsl:if test="图表:填充_E746/图:图片_8005">
      <xsl:call-template name="PicFill"/>
    </xsl:if>
  </xsl:template>

  <!--Modified by LDM in 2010/12/17-->
  <!--PatternFill-->
  <!--图案填充模板-->
  <xsl:template name="PatternFill">
    <xsl:if test="图表:填充_E746/图:图案_800A">
      <a:pattFill>
        <xsl:attribute name="prst">
          <xsl:call-template name="PatternMapping"/>
        </xsl:attribute>
        <!--前景色-->
        <a:fgClr>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring-after(图表:填充_E746/图:图案_800A/@前景色_800B,'#')"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:fgClr>
        <!-- 20130508 updata by xuzhenwei 修改图表背景主题色 start -->
        <xsl:variable name="zhutise" select="图表:填充_E746/图:图案_800A/@背景色_800C"/>
        <!--背景色-->
        <a:bgClr>
          <xsl:choose>
            <xsl:when test="$zhutise='bg1'">
              <a:schemeClr val="bg1"/>
            </xsl:when>
            <xsl:otherwise>
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring-after(图表:填充_E746/图:图案_800A/@背景色_800C,'#')"/>
                </xsl:attribute>
              </a:srgbClr>
            </xsl:otherwise>
          </xsl:choose>
          <!-- end -->
        </a:bgClr>
      </a:pattFill>
    </xsl:if>
  </xsl:template>

  <!--Modified by LDM in 2010/12/17-->
  <!--PicFill-->
  <!--图片和纹理填充模板-->
  <xsl:template name="PicFill">
    <a:blipFill>
      <a:blip>
        <xsl:attribute name="r:embed">
          <xsl:value-of select="concat('rId',./图表:填充_E746/图:图片_8005/@图形引用_8007)"/>
        </xsl:attribute>
      </a:blip>
      <!--tile-->
      <xsl:if test="图表:填充_E746/图:图片_8005/位置_8006='tile'">
        <a:tile/>
      </xsl:if>
      <!--a:srcRect/-->
      <!--stretch-->
      <!--xsl:if test="图表:填充_E746/图:图片_8005/位置_8006='stretch'"-->
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      <!--/xsl:if-->
    </a:blipFill>
  </xsl:template>

  <!--Modified by LDM in 2010/12/17-->
  <!--表边框-->
  <!--Border-->
  <xsl:template name="Border">
    <!--表边框-->
    <xsl:if test="./图表:边框线_4226">
      <a:ln>
        <!--边框宽度-->
        <xsl:if test="./图表:边框线_4226/@宽度_C60F">
          <xsl:attribute name="w">
            <xsl:value-of select="(./图表:边框线_4226/@宽度_C60F)*12700"/>
          </xsl:attribute>
        </xsl:if>
        <!--默认值-->
        <xsl:attribute name="cap">
          <xsl:value-of select="'flat'"/>
        </xsl:attribute>
        <!--边框线型-->
        <xsl:if test="./图表:边框线_4226/@线型_C60D and not(./图表:边框线_4226/@线型_C60D = 'none')">
          <xsl:attribute name="cmpd">
            <xsl:call-template name="LineTypeMapping_Border">
              <xsl:with-param name="lineType" select="./图表:边框线_4226/@线型_C60D"/>
            </xsl:call-template>
          </xsl:attribute>
        </xsl:if>
        <!--需要注意各元素的出现顺序，在OOXML中和子元素是sequence类型的，有次序关系的，否则会有转换错误-->
        <!--边框颜色-->

        <!--zl 20150519-->
        <!--<xsl:if test="./图表:边框线_4226/@线型_C60D = 'none'">
          <a:noFill/>
        </xsl:if>-->
        <!--2014-4-23, update by Qihy, 修复bug3080 uof-oo转换中边框线无线条填充的转换， start-->
        <xsl:if test ="./图表:边框线_4226[@颜色_C611] and not(./图表:边框线_4226/@线型_C60D = 'none')">
        <!--zl 20150519-->
          
        <a:solidFill>
            <!--<a:schemeClr val="tx1"/>-->
        <xsl:if test="./图表:边框线_4226/@颜色_C611 !='auto'">
          
            <a:srgbClr>
              <xsl:attribute name="val">
                <xsl:value-of select="substring-after(./图表:边框线_4226/@颜色_C611,'#')"/>
              </xsl:attribute>
            </a:srgbClr>
          
        </xsl:if>
        </a:solidFill>
        </xsl:if>
        <xsl:if test="./图表:边框线_4226/@线型_C60D = 'none' and not(./图表:边框线_4226[@颜色_C611])">
          <a:noFill/>
        </xsl:if>

        <xsl:if test="not(./图表:边框线_4226[@颜色_C611]) and ./图表:边框线_4226/图:渐变_800D">
          <xsl:apply-templates select="./图表:边框线_4226/图:渐变_800D"/>
        </xsl:if>
        <!--2014-4-23 end-->
        
        <!--边框类型-->
        <xsl:if test="./图表:边框线_4226/@线型_C60D and not(./图表:边框线_4226/@线型_C60D = 'none')">
          <xsl:variable name="dashType">
            <xsl:if test="./图表:边框线_4226/@虚实_C60E">
              <xsl:value-of select="./图表:边框线_4226/@虚实_C60E"/>
            </xsl:if>
            <xsl:if test="not(./图表:边框线_4226/@虚实_C60E)">
              <xsl:value-of select="'noDash'"/>
            </xsl:if>
          </xsl:variable>
          <xsl:call-template name="LineDashMapping_Border">
            <xsl:with-param name="dashType" select="$dashType"/>
          </xsl:call-template>
        </xsl:if>
        
      </a:ln>
    </xsl:if>
    <!--阴影-->
    <!--<xsl:if test="表:边框/@uof:阴影='true'">
      <a:effectLst>
        <xsl:call-template name="shadow"/>
      </a:effectLst>
    </xsl:if>-->
  </xsl:template>

  <!--渐变模板-->
  <xsl:template match="图:渐变_800D">
    <xsl:variable name="angle" select="@渐变方向_8013"/>
    <a:gradFill>
      <a:gsLst>
        <xsl:choose>
          <xsl:when test="@种子类型_8010='radar'">
            <a:gs pos="0%">
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring(@起始色_800E,2,6)"/>
                </xsl:attribute>
              </a:srgbClr>
            </a:gs>
            <a:gs pos="50%">
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring(@终止色_800F,2,6)"/>
                </xsl:attribute>
              </a:srgbClr>
            </a:gs>
            <a:gs pos="100%">
              <a:srgbClr>
                <xsl:attribute name="val">
                  <xsl:value-of select="substring(@起始色_800E,2,6)"/>
                </xsl:attribute>
              </a:srgbClr>
            </a:gs>
          </xsl:when>
          <xsl:otherwise>
            <xsl:choose>
              <xsl:when test="$angle='135' or $angle='180' or $angle='225' or $angle='270'">
                <a:gs pos="100%">
                  <a:srgbClr>
                    <xsl:attribute name="val">
                      <xsl:value-of select="substring(@起始色_800E,2,6)"/>
                    </xsl:attribute>
                  </a:srgbClr>
                </a:gs>
                <a:gs pos="0%">
                  <a:srgbClr>
                    <xsl:attribute name="val">
                      <xsl:value-of select="substring(@终止色_800F,2,6)"/>
                    </xsl:attribute>
                  </a:srgbClr>
                </a:gs>
              </xsl:when>
              <xsl:otherwise>
                <a:gs pos="0">
                  <a:srgbClr>
                    <xsl:attribute name="val">
                      <xsl:value-of select="substring(@起始色_800E,2,6)"/>
                    </xsl:attribute>
                  </a:srgbClr>
                </a:gs>
                <a:gs pos="100%">
                  <a:srgbClr>
                    <xsl:attribute name="val">
                      <xsl:value-of select="substring(@终止色_800F,2,6)"/>
                    </xsl:attribute>
                  </a:srgbClr>
                </a:gs>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </a:gsLst>
      <xsl:choose>
        <xsl:when test="@种子类型_8010='square'">
          <a:path>
            <xsl:attribute name="path">rect</xsl:attribute>
            <xsl:variable name="x" select="@种子X位置_8015"/>
            <xsl:variable name="y" select="@种子Y位置_8016"/>
            <xsl:choose>
              <xsl:when test="$x='30' and $y='30'">
                <a:fillToRect r="100%" b="100%"/>
              </xsl:when>
              <xsl:when test="$x='30' and $y='60'">
                <a:fillToRect t="100%" r="100%"/>
              </xsl:when>
              <xsl:when test="$x='60' and $y='30'">
                <a:fillToRect l="100%" b="100%"/>
              </xsl:when>
              <xsl:when test="$x='60' and $y='60'">
                <a:fillToRect l="100%" t="100%"/>
              </xsl:when>
              
            </xsl:choose>
          </a:path>
        </xsl:when>
        
		 <!--20130319 gaoyuwei add bug_2749 uof-oo 功能测试 预定义图形“渐变”填充不正确-->
		  <xsl:when test="@种子类型_8010='rectangle'">
			  <a:path>
				  <xsl:attribute name="path">shape</xsl:attribute>
				  <xsl:variable name="x" select="@种子X位置_8015"/>
				  <xsl:variable name="y" select="@种子Y位置_8016"/>
				  <xsl:choose>
					  <xsl:when test="$x='50' and $y='50'">
						  <a:fillToRect l="50%" t="50%" r="50%" b="50%"/>
					  </xsl:when>
				  </xsl:choose>
			  </a:path>
		  </xsl:when>
		  <!--end-->
		  
        <xsl:otherwise>
          <a:lin>
            <xsl:attribute name="ang">
              <xsl:choose>
                <xsl:when test="$angle='0'">
                  <xsl:value-of select="90*60000"/>
                </xsl:when>
                <xsl:when test="$angle='45'">
                  <xsl:value-of select="45*60000"/>
                </xsl:when>
                <xsl:when test="$angle='90'">
                  <xsl:value-of select="0*60000"/>
                </xsl:when>
                <xsl:when test="$angle='135'">
                  <xsl:value-of select="315*60000"/>
                </xsl:when>
                <xsl:when test="$angle='180'">
                  <xsl:value-of select="270*60000"/>
                </xsl:when>
                <xsl:when test="$angle='225'">
                  <xsl:value-of select="225*60000"/>
                </xsl:when>
                <xsl:when test="$angle='270'">
                  <xsl:value-of select="180*60000"/>
                </xsl:when>
                <xsl:when test="$angle='315'">
                  <xsl:value-of select="135*60000"/>
                </xsl:when>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="scaled">1</xsl:attribute>
          </a:lin>
        </xsl:otherwise>
      </xsl:choose>
      <a:tileRect/>
    </a:gradFill>
  </xsl:template>

  <!--Modified by LDM in 2010/12/17-->
  <!--获取系列名模板-->
  <!--GetSeriesName-->
  <xsl:template name="GetSeriesName">
    <xsl:param name="refValue"/>
    <!--系列名是引用的一系列单元格-->
    <xsl:choose>
      <!--引用形式如A1:C1-->
      <xsl:when test="contains($refValue,':')">
        <xsl:variable name="sheetName">
          <xsl:value-of select="substring-before($refValue,'!')"/>
        </xsl:variable>
        <xsl:variable name="sheetNameTemp1">
          <!--<xsl:value-of select='translate($sheetName,"&apos;","")'/>-->
          <xsl:value-of select='$sheetName'/>
        </xsl:variable>
        <xsl:variable name="sheetNameTemp">
          <xsl:value-of select='translate($sheetNameTemp1,"=","")'/>
        </xsl:variable>
        <xsl:variable name="seriesValue">
          <xsl:value-of select="substring-after($refValue,'!')"/>
        </xsl:variable>
        <xsl:variable name="svStart">
          <xsl:value-of select="substring-before($seriesValue,':')"/>
        </xsl:variable>
        <xsl:variable name="svEnd">
          <xsl:value-of select="substring-after($seriesValue,':')"/>
        </xsl:variable>
        <xsl:variable name="svStartLength">
          <xsl:value-of select="string-length($svStart)"/>
        </xsl:variable>
        <xsl:variable name="svStartPart1">
          <xsl:value-of select="substring($svStart,1,1)"/>
        </xsl:variable>
        <xsl:variable name="svStartPart2">
          <xsl:value-of select="substring($svStart,2,$svStartLength - 1)"/>
        </xsl:variable>
        <xsl:variable name="svStartNew">
          <xsl:value-of select="concat('$',$svStartPart1,'$',$svStartPart2)"/>
        </xsl:variable>
        <xsl:variable name="svEndLength">
          <xsl:value-of select="string-length($svEnd)"/>
        </xsl:variable>
        <xsl:variable name="svEndPart1">
          <xsl:value-of select="substring($svEnd,1,1)"/>
        </xsl:variable>
        <xsl:variable name="svEndPart2">
          <xsl:value-of select="substring($svEnd,2,$svEndLength - 1)"/>
        </xsl:variable>
        <xsl:variable name="svEndNew">
          <xsl:value-of select="concat('$',$svEndPart1,'$',$svEndPart2)"/>
        </xsl:variable>
        <xsl:variable name="seriesNew">
          <xsl:value-of select="concat($sheetNameTemp,'!',$svStartNew,':',$svEndNew)"/>
        </xsl:variable>
        <xsl:value-of select="$seriesNew"/>
      </xsl:when>
      <!--引用形式如A1-->
      <xsl:otherwise>
        <xsl:variable name="sheetName">
          <xsl:value-of select="substring-before($refValue,'!')"/>
        </xsl:variable>
        <xsl:variable name="sheetNameTemp1">
          <xsl:value-of select='translate($sheetName,"&apos;","")'/>
        </xsl:variable>
        <xsl:variable name="sheetNameTemp">
          <xsl:value-of select='translate($sheetNameTemp1,"=","")'/>
        </xsl:variable>
        <xsl:variable name="seriesValue">
          <xsl:value-of select="substring-after($refValue,'!')"/>
        </xsl:variable>
        <xsl:variable name="svLength">
          <xsl:value-of select="string-length($seriesValue)"/>
        </xsl:variable>
        <xsl:variable name="svPart1">
          <xsl:value-of select="substring($seriesValue,1,1)"/>
        </xsl:variable>
        <xsl:variable name="svPart2">
          <xsl:value-of select="substring($seriesValue,2,$svLength - 1)"/>
        </xsl:variable>
        <xsl:variable name="svNew">
          <xsl:value-of select="concat('$',$svPart1,'$',$svPart2)"/>
        </xsl:variable>

        <xsl:value-of select="$refValue"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--Not Finished-->
  <!--Modified by LDM in 2010/12/29-->
  <xsl:template name="UnderlineTypeMapping">
    <xsl:param name="lineType"/>
    <xsl:param name="dashType"/>
    <xsl:choose>
      <xsl:when test="$lineType='single'">
        <xsl:choose>
          <xsl:when test="$dashType='square-dot'">
            <xsl:value-of select="'dotted'"/>
          </xsl:when>
          <xsl:when test="$dashType='dash'">
            <xsl:value-of select="'dash'"/>
          </xsl:when>
          <xsl:when test="$dashType='long-dash'">
            <xsl:value-of select="'dashLong'"/>
          </xsl:when>
          <xsl:when test="$dashType='dash-dot'">
            <xsl:value-of select="'dotDash'"/>
          </xsl:when>
          <xsl:when test="$dashType='dash-dot-dot'">
            <xsl:value-of select="'dotDotDash'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'sng'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$lineType='double'">
        <xsl:value-of select="'dbl'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--Modified by LDM in 2010/12/16-->
  <!--Font-->
  <!--字体设置模板-->
  <xsl:template name="Font">
    <!--font size-->
    <xsl:attribute name="sz">
      <xsl:choose>
        <xsl:when test="./图表:字体_E70B/字:字体_4128[@字号_412D]">
          <xsl:value-of select="floor((图表:字体_E70B/字:字体_4128/@字号_412D) * 100)"/>
        </xsl:when>
      <!--默认字号大小为10-->
        <xsl:otherwise>
          <xsl:value-of select="'1000'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <!--bold-->
    <!--<xsl:value-of select="./图表:字体_E70B/字:字体_4128/@颜色_412F"/>-->
    <xsl:attribute name="b">
      <xsl:choose>
        <xsl:when test="./图表:字体_E70B[字:是否粗体_4130='true']">
          <xsl:value-of select="'1'"/>
        </xsl:when>
        <xsl:when test="./图表:字体_E70B/字:是否粗体_4130='true'">
          <xsl:value-of select="'1'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'0'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <!--italic-->
    <xsl:if test="./图表:字体_E70B[字:是否斜体_4131='true']">
      <xsl:attribute name="i">
        <xsl:value-of select="'1'"/>
      </xsl:attribute>
    </xsl:if>
    <!--文本框中的字符串下划线类型-->
    <!--Not Finished-->
    <!--Modified by LDM in 2010/12/29-->
    <xsl:if test="./图表:字体_E70B/字:下划线_4136[@线型_4137!='none']">
      <xsl:attribute name="u">
        <xsl:variable name="lineType">
          <xsl:value-of select="./图表:字体_E70B/字:下划线_4136/@线型_4137"/>
        </xsl:variable>
        <xsl:variable name="dashType">
          <xsl:value-of select="./图表:字体_E70B/字:下划线_4136/@虚实_4138"/>
        </xsl:variable>
        <xsl:call-template name="UnderlineTypeMapping">
          <xsl:with-param name="lineType" select="$lineType"/>
          <xsl:with-param name="dashType" select="$dashType"/>
        </xsl:call-template>
      </xsl:attribute>
    </xsl:if>

    <!--Need modify of the rest of this template-->
    <xsl:if test="./图表:字体_E70B/字:删除线_4135">
      <xsl:choose>
        <xsl:when test="./图表:字体_E70B[字:删除线_4135='none']">
          <xsl:attribute name="strike">
            <xsl:value-of select="'noStrike'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="./图表:字体_E70B[字:删除线_4135='single']">
          <xsl:attribute name="strike">
            <xsl:value-of select="'sngStrike'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="./图表:字体_E70B[字:删除线_4135='double']">
          <xsl:attribute name="strike">
            <xsl:value-of select="'dblStrike'"/>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <xsl:if test="./图表:字体_E70B/字:上下标类型_4143">
      <xsl:choose>
        <xsl:when test="./图表:字体_E70B[字:上下标类型_4143='sup']">
          <xsl:attribute name="baseline">
            <xsl:value-of select="'30%'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="./图表:字体_E70B[字:上下标类型_4143='sub']">
          <xsl:attribute name="baseline">
            <xsl:value-of select="'-25%'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="./图表:字体_E70B[字:上下标类型_4143='none']">
          <xsl:attribute name="baseline">
            <xsl:value-of select="'0%'"/>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:if>
    <!--字符间距、调整字间距，显示不出来,而且没有办法验证是否正确-->
    <xsl:if test="./图表:字体_E70B/字:字符间距_4145">
      <xsl:attribute name="spc">
        <xsl:value-of select="(./图表:字体_E70B/字:字符间距_4145)*100"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="./图表:字体_E70B/字:调整字间距_4146">
      <xsl:attribute name="kern">
        <xsl:value-of select="(./图表:字体_E70B/字:调整字间距_4146) * 100 "/>
      </xsl:attribute>
    </xsl:if>

    <xsl:if test="./图表:字体_E70B/字:醒目字体类型_4141='uppercase'">
      <xsl:attribute name="cap">
        <xsl:value-of select="'all'"/>
      </xsl:attribute>
    </xsl:if>
    <xsl:if test="./图表:字体_E70B/字:醒目字体类型_4141='small-caps'">
      <xsl:attribute name="cap">
        <xsl:value-of select="'small'"/>
      </xsl:attribute>
    </xsl:if>

    <!--Modified by LDM in 2010/12/16-->
    <!--font color-->
    <a:solidFill>
      <a:srgbClr>
        <xsl:attribute name="val">
          <xsl:choose>
            <xsl:when test="./图表:字体_E70B/字:字体_4128/@颜色_412F!='auto'">
              <xsl:variable name="fontColor">
                <xsl:value-of select="./图表:字体_E70B/字:字体_4128/@颜色_412F"/>
              </xsl:variable>
              <xsl:value-of select="substring-after($fontColor,'#')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'000000'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </a:srgbClr>
    </a:solidFill>

    <xsl:if test="./图表:字体_E70B/字:下划线_4136/@颜色_412F!='auto'">
      <a:uFill>
        <a:solidFill>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring-after(./图表:字体_E70B/字:下划线_4136/@颜色_412F,'#')"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:solidFill>
      </a:uFill>
    </xsl:if>
    <!--Modified by LDM in 2010/12/16-->
    <!--font face-->
    <a:latin>
      <xsl:attribute name="typeface">
        <xsl:choose>
          <xsl:when test="./图表:字体_E70B/字:字体_4128/@西文字体引用_4129">
            <xsl:variable name="latinRef">
              <xsl:value-of select="./图表:字体_E70B/字:字体_4128/@西文字体引用_4129"/>
            </xsl:variable>
            <xsl:value-of select="/uof:UOF/式样:式样集_990B/式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$latinRef]/@名称_9903"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'Times New Roman'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="pitchFamily">
        <xsl:value-of select="'2'"/>
      </xsl:attribute>
      <xsl:attribute name="charset">
        <xsl:value-of select="'-122'"/>
      </xsl:attribute>
    </a:latin>
    <a:ea>
      <xsl:attribute name="typeface">
        <xsl:choose>
          <xsl:when test="./图表:字体_E70B/字:字体_4128/@中文字体引用_412A">
            <xsl:variable name="eaRef">
              <xsl:value-of select="./图表:字体_E70B/字:字体_4128/@中文字体引用_412A"/>
            </xsl:variable>
            <xsl:variable name="fontname">
              <xsl:value-of select="/uof:UOF/式样:式样集_990B/式样:字体集_990C/式样:字体声明_990D[@标识符_9902=$eaRef]/@名称_9903"/>
            </xsl:variable>
            <xsl:if test="$fontname=''">
              <xsl:value-of select="'宋体'"/>
            </xsl:if>
            <xsl:if test="not($fontname='')">
              <xsl:value-of select="$fontname"/>
            </xsl:if>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'宋体'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="pitchFamily">
        <xsl:value-of select="'2'"/>
      </xsl:attribute>
      <xsl:attribute name="charset">
        <xsl:value-of select="'-122'"/>
      </xsl:attribute>
    </a:ea>
  </xsl:template>

  <!--LegendPos-->
  <!--Legend position-->
  <xsl:template name="LegendPos">
    <c:legendPos>
      <xsl:choose>
        <xsl:when test="./图表:图例位置_E795='bottom'">
          <xsl:attribute name="val">
            <xsl:value-of select="'b'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="./图表:图例位置_E795='top'">
          <xsl:attribute name="val">
            <xsl:value-of select="'t'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="./图表:图例位置_E795='left'">
          <xsl:attribute name="val">
            <xsl:value-of select="'l'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="./图表:图例位置_E795='right'">
          <xsl:attribute name="val">
            <xsl:value-of select="'r'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="./图表:图例位置_E795='corner'">
          <xsl:attribute name="val">
            <xsl:value-of select="'tr'"/>
          </xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </c:legendPos>
  </xsl:template>

  <!--Modified by LDM in 2010/12/17-->
  <!--线型模板-->
  <xsl:template match="图表:边框线_4226">
    
    <!--2014-4-22, update by Qihy, 修复bug3209图表边框线转换不正确引起打开需要恢复问题， start-->
    <!--<xsl:call-template name="Line"/>-->
    <xsl:call-template name="Border"/>
    <!--2014-4-22 end-->
    
  </xsl:template>

  <!--Modified by LDM in 2010/12/17-->
  <!--图例模板-->
  <xsl:template match ="图表:图例_E794">
    <c:legend>
     
      <!--2014-6-8, update by lingfeng, 图例丢失， start-->
      <xsl:if test="./图表:图例位置_E795='bottom' and not(./图表:位置_E70A) and not(./图表:大小_E748) and not(./图表:填充_E746) and not(./图表:图例项_E765)">
        <c:legendPos val="b" />
        <c:layout />
        <c:overlay val="0" />
        <c:spPr>
          <a:ln cap="flat">
            <a:noFill />
          </a:ln>
        </c:spPr>
        <c:txPr>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:defRPr sz="1000" b="0">
                <a:ea typeface="宋体" pitchFamily="2" charset="-122" />
              </a:defRPr>
            </a:pPr>
          </a:p>
        </c:txPr>
      </xsl:if>

      <xsl:if test="not( ./图表:图例位置_E795='bottom' and not(./图表:位置_E70A) and not(./图表:大小_E748) and not(./图表:填充_E746) and not(./图表:图例项_E765) )">

        <xsl:if test="./图表:图例位置_E795">
          <xsl:call-template name="LegendPos"/>
        </xsl:if>
        <c:layout>
          <xsl:if test="./图表:位置_E70A/@x_C606">
            <xsl:call-template name="Layout">
              <xsl:with-param name="off_X" select="'0'"/>
              <xsl:with-param name="off_Y" select="'0'"/>
            </xsl:call-template>
          </xsl:if>
        </c:layout>
        <c:overlay val="0"/>
        <c:spPr>
          <xsl:if test="./图表:填充_E746">
            <xsl:call-template name="Fill"/>
          </xsl:if>
          <xsl:if test="./图表:边框线_4226">
            <xsl:call-template name="Border">
            </xsl:call-template>
          </xsl:if>
          <xsl:if test="not(./图表:边框线_4226)">
            <a:ln cap="flat">
              <a:solidFill>
                <a:srgbClr val="000000"/>
              </a:solidFill>
              <a:prstDash val="solid"/>
            </a:ln>
          </xsl:if>
        </c:spPr>
        <xsl:if test="./图表:字体_E70B">
          <c:txPr>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:pPr>
                <a:defRPr>
                  <xsl:call-template name="Font"/>
                </a:defRPr>
              </a:pPr>
            </a:p>
          </c:txPr>
        </xsl:if>
      </xsl:if>
      <!--2014-6-8 end-->

    </c:legend>
  </xsl:template>
  <!--Modified by LDM in 2010/12/17-->
  <!--绘图区模板-->
  <xsl:template match ="图表:绘图区_E747">
    <xsl:variable name="off_X">
      <xsl:value-of select ="'0'"/>
    </xsl:variable>
    <xsl:variable name="off_Y">
      <xsl:value-of select ="'0'"/>
    </xsl:variable>
    <c:layout>
      <xsl:if test="./图表:位置_E70A/@x_C606">
        <xsl:call-template name="Layout">
          <xsl:with-param name="off_X" select="$off_X"/>
          <xsl:with-param name="off_Y" select="$off_Y"/>
        </xsl:call-template>
      </xsl:if>
    </c:layout>
    <c:spPr>
      <xsl:if test="./图表:填充_E746">
        <xsl:call-template name="Fill">
        </xsl:call-template>
      </xsl:if>
      <xsl:if test="./图表:边框线_4226">
        <xsl:call-template name="Border">
        </xsl:call-template>
      </xsl:if>
    </c:spPr>
  </xsl:template>
  <!--Modified by LDM in 2010/12/17-->
  <!--图表区模板-->
  <xsl:template match ="图表:图表区_E743">
    <c:spPr>
      <xsl:if test="./图表:填充_E746">
        <xsl:call-template name="Fill">
        </xsl:call-template>
      </xsl:if>
      <xsl:if test="./图表:边框线_4226">
        <xsl:call-template name="Border">
        </xsl:call-template>
      </xsl:if>
    </c:spPr>
    <xsl:if test="./图表:字体_E70B">
      <c:txPr>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:pPr>
            <a:defRPr>
              <xsl:call-template name="Font"/>
            </a:defRPr>
          </a:pPr>
        </a:p>
      </c:txPr>
    </xsl:if>
  </xsl:template>
  <!--Modified by LDM in 2010/12/17-->
  <!--数据表模板-->
  <xsl:template match ="图表:数据表_E79B">
    <c:dTable>
     <!--20130115 gaoyuwei bug_2641数据表“水平、垂直、外边框”工作表，无边框转换后为默认边框  start-->
		<!--显示水平线-->
		<xsl:if test="./@是否水平显示_E79F='true'">
			 <c:showHorzBorder val="1"/>
			</xsl:if>
		<xsl:if test="./@是否水平显示_E79F='false'">
			<c:showHorzBorder val="0"/>
		</xsl:if>
		<!--显示垂直线-->
		<xsl:if test="./@是否垂直显示_E79E ='true'">
			   <c:showVertBorder val="1"/>
		</xsl:if>
		<xsl:if test="./@是否垂直显示_E79E ='false'">
			<c:showVertBorder val="0"/>
		</xsl:if>
		<!--显示垂直线-->
		<xsl:if test="./@是否显示外边框_E79D='true'">
				<c:showOutline val="1"/>
		</xsl:if>
		<xsl:if test="./@是否垂直显示_E79E ='false'">
				<c:showOutline val="0"/>
		</xsl:if>
			<!--显示图例-->
			<xsl:if test="./@是否显示图例_E79C='true'">
				 <c:showKeys val="1"/>
			</xsl:if>
			<xsl:if test="./@是否显示图例_E79C='false'">
				<c:showKeys val="0"/>
			</xsl:if>
		
		<!--end-->
      <!--<xsl:if test="./表:填充">
        <c:spPr>
          <xsl:call-template name="Fill"/>
        </c:spPr>
      </xsl:if>-->
      <xsl:if test="./图表:字体_E70B">
        <c:txPr>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:defRPr>
                <xsl:call-template name="Font"/>
              </a:defRPr>
            </a:pPr>
          </a:p>
        </c:txPr>
      </xsl:if>
      <!-- add by xuzhenwei bug_2503: 数据表边框转换后为无边框 2013-01-19 start -->
      <xsl:if test="./图表:边框线_4226">
        <c:spPr>
          <xsl:call-template name="Border"/>
        </c:spPr>
      </xsl:if>
      <!-- end -->
    </c:dTable>
  </xsl:template>

  <!--Modified by LDM in 2011/01/01-->
  <!--轴调用模板-->
  <!--AxisTemplate-->
  <xsl:template name="AxisTemplate">
    <xsl:param name="axisType"/>
    <!--
    <xsl:if test="$axisType='category axis'">
      <c:axId val="000000000"/>
    </xsl:if>
    <xsl:if test="$axisType='value axis'">
      <c:axId val="111111111"/>
    </xsl:if>
    -->
    <c:scaling>
      <!-- 20130426 update by xuzhenwei office软件版本从1000升级到5000，造成修复的原因 start -->
      <xsl:if test="./图表:刻度_E71D/图表:是否显示为对数刻度_E729='true'">
        <c:logBase>
          <xsl:attribute name="val">
            <xsl:value-of select="'10'"/>
          </xsl:attribute>
        </c:logBase>
      </xsl:if>
      <!-- end -->
      <c:orientation>
        <xsl:attribute name="val">
          <xsl:choose>
            <xsl:when test="./图表:刻度_E71D/图表:是否次序反转_E72B='true'">
              <xsl:value-of select="'maxMin'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'minMax'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </c:orientation>
      <xsl:if test="./图表:刻度_E71D/图表:最大值_E720">
        <c:max>
          <xsl:attribute name="val">
            <xsl:value-of select="./图表:刻度_E71D/图表:最大值_E720" />
          </xsl:attribute>
        </c:max>
      </xsl:if>
      <xsl:if test="./图表:刻度_E71D/图表:最小值_E71E">
        <c:min>
          <xsl:attribute name="val">
            <xsl:value-of select="./图表:刻度_E71D/图表:最小值_E71E" />
          </xsl:attribute>
        </c:min>
      </xsl:if>
    </c:scaling>
    <c:axPos>
      <xsl:attribute name="val">
        <xsl:if test="$axisType='category axis'">
          <xsl:value-of select="'b'"/>
        </xsl:if>
        <xsl:if test="$axisType='value axis'">
          <xsl:value-of select="'l'"/>
        </xsl:if>
      </xsl:attribute>
    </c:axPos>
	  <!--20130424 update by gaoyuwei_qhy bug_2876	图表刻度线类型和刻度线标致转换后不正确
	  bug_2877 回归 功能 uof-oo 坐标轴刻度线转换后丢失 -->
	  <!--主刻度类型-->
	  <xsl:if test="./@主刻度类型_E737">
		  <c:majorTickMark>
			  <xsl:attribute name="val">
				  <xsl:choose>
					  <xsl:when test="./@主刻度类型_E737='outside'">
						  <xsl:value-of select="'out'"/>
					  </xsl:when>
					  <xsl:when test="./@主刻度类型_E737='inside'">
						  <xsl:value-of select="'in'"/>
					  </xsl:when>
					  <xsl:when test="./@主刻度类型_E737='cross'">
						  <xsl:value-of select="'cross'"/>
					  </xsl:when>
					  <xsl:otherwise>
						  <xsl:value-of select="'none'"/>
					  </xsl:otherwise>
				  </xsl:choose>
			  </xsl:attribute>
		  </c:majorTickMark>
	  </xsl:if>
	  <!--次刻度类型-->
	  <xsl:if test="./@次刻度类型_E738">
		  <c:minorTickMark>
			  <xsl:attribute name="val">
				  <xsl:choose>
					  <xsl:when test="./@次刻度类型_E738='outside'">
						  <xsl:value-of select="'out'"/>
					  </xsl:when>
					  <xsl:when test="./@次刻度类型_E738='inside'">
						  <xsl:value-of select="'in'"/>
					  </xsl:when>
					  <xsl:when test="./@次刻度类型_E738='cross'">
						  <xsl:value-of select="'cross'"/>
					  </xsl:when>
					  <xsl:otherwise>
						  <xsl:value-of select="'none'"/>
					  </xsl:otherwise>
				  </xsl:choose>
			  </xsl:attribute>
		  </c:minorTickMark>
	  </xsl:if>
	  <!--end-->
    <!--刻度线标志-->
    <xsl:if test="./@刻度线标志_E739">
      <c:tickLblPos>
        <xsl:attribute name="val">
        
		 <!--20130121 gaoyuwei bug2645_3 工作表“刻度”中，分类轴Y轴刻度相关设置转换后不正确 start-->
			<xsl:if test="$axisType='category axis'">
				<xsl:choose>
					<xsl:when test="../图表:坐标轴_E791[@子类型_E793='category']/@刻度线标志_E739='none'">
						<xsl:value-of select="'none'"/>
					</xsl:when>
					<xsl:when test="../图表:坐标轴_E791[@子类型_E793='category']/@刻度线标志_E739='inside'">
						<xsl:value-of select="'high'"/>
					</xsl:when>
					<xsl:when test="../图表:坐标轴_E791[@子类型_E793='category']/@刻度线标志_E739='outside'">
						<xsl:value-of select="'low'"/>
					</xsl:when>
					<xsl:when test="../图表:坐标轴_E791[@子类型_E793='category']/@刻度线标志_E739='next-to-axis'">
						<xsl:value-of select="'nextTo'"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="'low'"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:if>
			<xsl:if test="$axisType='value axis'">
				<xsl:choose>
					<xsl:when test="../图表:坐标轴_E791[@子类型_E793='value']/@刻度线标志_E739='none'">
						<xsl:value-of select="'none'"/>
					</xsl:when>
					<xsl:when test="../图表:坐标轴_E791[@子类型_E793='value']/@刻度线标志_E739='inside'">
						<xsl:value-of select="'high'"/>
					</xsl:when>
					<xsl:when test="../图表:坐标轴_E791[@子类型_E793='value']/@刻度线标志_E739='outside'">
						<xsl:value-of select="'low'"/>
					</xsl:when>
					<xsl:when test="../图表:坐标轴_E791[@子类型_E793='value']/@刻度线标志_E739='next-to-axis'">
						<!--20130422 update by gaoyuwei_qhy bug_2878 工作表“刻度”中，"X轴刻度的设置：数值（Y）轴置于分类轴上的分类号-10" 原来为low-->
						<!--“轴旁”-->
						<xsl:value-of select="'nextTo'"/>
						<!--end-->
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="'low'"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:if>
			<!--end-->
			
        </xsl:attribute>
      </c:tickLblPos>
    </xsl:if>

    <!--分类刻度数-->
	  <!--20130422 update by gaoyuwei_qhy bug_2879 工作表“刻度”中，“X轴刻度的设置：分类轴刻度之间的分类数-2”,"分类标志之间的分类数-2"图表转换后不正确-->
	  <xsl:if test="./图表:刻度_E71D/图表:分类刻度数_E72D">
		  <c:tickMarkSkip>
			  <xsl:attribute name="val">
				  <xsl:value-of select="./图表:刻度_E71D/图表:分类刻度数_E72D"/>
			  </xsl:attribute>
		  </c:tickMarkSkip>
	  </xsl:if>
	  <!-- 分类标签数 -->
	  <xsl:if test="./图表:刻度_E71D/图表:分类标签数_E72C">
		  <c:tickLblSkip>
			  <xsl:attribute name="val">
				  <xsl:value-of select="./图表:刻度_E71D/图表:分类标签数_E72C"/>
			  </xsl:attribute>
		  </c:tickLblSkip>
	  </xsl:if>
	  <!--end-->

    <!--默认值-->

	  <!--20130422 update by gaoyuwei_qhy 002坐标轴集 工作表“刻度”中，"Y轴刻度的设置：选中“分类轴X轴交叉于最大值"转换不正确-->
	  <xsl:if test="$axisType='value axis'">
		  <xsl:choose>
			  <xsl:when test="../图表:坐标轴_E791[@子类型_E793='category']/图表:刻度_E71D/图表:是否交叉于最大值_E72E='false'">
				  <c:crosses val="autoZero"/>
			  </xsl:when>
			  <xsl:when test="../图表:坐标轴_E791[@子类型_E793='category']/图表:刻度_E71D/图表:是否交叉于最大值_E72E='true'">
				  <c:crosses val="max"/>
			  </xsl:when>
			  <xsl:otherwise/>
		  </xsl:choose>
	  </xsl:if>
	  <xsl:if test="$axisType='category axis'">
		  <xsl:choose>
			  <xsl:when test="../图表:坐标轴_E791[@子类型_E793='value']/图表:刻度_E71D/图表:是否交叉于最大值_E72E='false'">
				  <c:crosses val="autoZero"/>
			  </xsl:when>
			  <xsl:when test="../图表:坐标轴_E791[@子类型_E793='value']/图表:刻度_E71D/图表:是否交叉于最大值_E72E='true'">
				  <c:crosses val="max"/>
			  </xsl:when>
			  <xsl:otherwise/>
		  </xsl:choose>
	  </xsl:if>
	  <!--end-->
	  
	  <!--20130422 qihouying bug_2878 图表Y轴的位置不正确 即工作表“刻度”中，"X轴刻度的设置：数值（Y）轴置于分类轴上的分类号-10" Y轴位置不正确-->

	  <xsl:if test="$axisType='value axis'">
		  <xsl:if test="../图表:坐标轴_E791[@子类型_E793='category']/图表:刻度_E71D/图表:交叉点_E723">
			  <c:crossesAt>
				  <xsl:attribute name="val">
					  <xsl:value-of select="../图表:坐标轴_E791[@子类型_E793='category']/图表:刻度_E71D/图表:交叉点_E723"/>
				  </xsl:attribute>
			  </c:crossesAt>
		  </xsl:if>
	  </xsl:if>
	  <xsl:if test="$axisType='category axis'">
		  <xsl:if test="../图表:坐标轴_E791[@子类型_E793='value']/图表:刻度_E71D/图表:交叉点_E723">
			  <c:crossesAt>
				  <xsl:attribute name="val">
					  <xsl:value-of select="../图表:坐标轴_E791[@子类型_E793='value']/图表:刻度_E71D/图表:交叉点_E723"/>
				  </xsl:attribute>
			  </c:crossesAt>
		  </xsl:if>
	  </xsl:if>
	  <!--end-->
	  
    <!--线型-->
    <xsl:if test="./图表:边框线_4226">
      <c:spPr>
        <xsl:apply-templates select="./图表:边框线_4226"/>
      </c:spPr>
    </xsl:if>
    <!--对齐-->
    <xsl:if test="./图表:对齐_E730 or ./图表:字体_E70B">
      <c:txPr>
        <a:bodyPr>
          <xsl:if test="./图表:对齐_E730">
            <xsl:call-template name="Alignment"/>
          </xsl:if>
        </a:bodyPr>
        <a:lstStyle/>
        <a:p>
          <a:pPr>
            <xsl:choose>
              <xsl:when test="./图表:字体_E70B">
                <a:defRPr>
                  <xsl:call-template name="Font"/>
                </a:defRPr>
              </xsl:when>
              <xsl:otherwise>
                <a:defRPr/>
              </xsl:otherwise>
            </xsl:choose>
          </a:pPr>
          <a:endParaRPr lang="zh-CN"/>
        </a:p>
      </c:txPr>
    </xsl:if>
    <xsl:if test="$axisType='category axis'">
      <xsl:variable name="formatN" select="../图表:坐标轴_E791/图表:数值_E70D/@格式码_E73F"/>
      <xsl:choose>
        <xsl:when test="contains($formatN,'yyyy')">
          <c:crossAx val="223502720"/>
        </xsl:when>
        <xsl:otherwise>
          <c:crossAx val="111111111"/>
        </xsl:otherwise>
      </xsl:choose>
      <!--表示值的对齐方式-->
      <c:lblAlgn val="ctr"/>
      <!--
    <c:lblOffset val="100"/>
    -->
    </xsl:if>
    <xsl:if test="$axisType='value axis'">
      <!-- 20130416 updata by xuzhenwei BUG_2830:互操作 oo-uof（编辑）-oo 024实用文档-损益表(1).xlsx 文档需要修复 start -->
      <!--<c:crossAx val="000000000"/>-->
      <xsl:variable name="formatN" select="../图表:坐标轴_E791/图表:数值_E70D/@格式码_E73F"/>
      <xsl:choose>
        <xsl:when test="contains($formatN,'yyyy')">
          <c:crossAx val="223499392"/>
          <c:crosses val="autoZero"/>
        </xsl:when>
        <xsl:otherwise>
          <c:crossAx val="000000000"/>
        </xsl:otherwise>
      </xsl:choose>
      <!-- end -->
      <c:crossBetween val="between"/>
      <!--默认值-->
      <!--
      <c:numFmt formatCode="General" sourceLinked="1"/>
      -->
    </xsl:if>
    <!--格式码-->
    <xsl:call-template name="NumberFormat"/>
    <!--
    <c:auto val="1"/>
    -->
    <!--偏移量-->
    <xsl:if test="./图表:对齐_E730/图表:偏移量_E732">
      <xsl:call-template name="Offset"/>
    </xsl:if>
    <!--可设置旋转角度-->
    <!--Not Finished-->
    <!--Modified by LDM in 2010/12/29-->
  </xsl:template>

  <!--数值轴模板-->
  <!--Modified by LDM in 2010/12/18-->
  <xsl:template name="数值轴">
    <xsl:call-template name="AxisTemplate">
      <xsl:with-param name="axisType" select="'value axis'"/>
    </xsl:call-template>
    <xsl:if test="./图表:刻度_E71D">
      <xsl:call-template name="ValAxisUnit"/>
    </xsl:if>
  </xsl:template>

  <!--分类轴模板-->
  <!--Modified by LDM in 2010/12/18-->
  <xsl:template name="分类轴">
    <xsl:call-template name="AxisTemplate">
      <xsl:with-param name="axisType" select="'category axis'"/>
    </xsl:call-template>
    <!--what does these codes mean?-->
    <!--
    <xsl:if  test="not(表:分类轴) and @子类型_E777!='doughnut_exploded' and @子类型_E777!='doughnut_standard' and @子类型_E777!='radar_standard' and @子类型_E777!='radar_marker' and @子类型_E777!='radar_filled'">
      <c:dateAx>
        <c:axId val="00000000"/>
        <c:scaling>
          <c:orientation val="minMax"/>
        </c:scaling>
        <c:axPos val="'b'"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="11111111"/>
        <c:crosses val="autoZero"/>
        <c:lblOffset val="100"/>
        <c:baseTimeUnit val="days"/>
      </c:dateAx>
    </xsl:if>
    -->
  </xsl:template>

  <!--Modified by LDM in 2011/01/03-->
  <!--数字格式模板-->
  <!--决定提到预处理部分用C#代码解决格式码转换问题-->
  <!--此模板仅是简单默认版本-->
  <!--Marked by LDM in 2011/01/03-->
  <xsl:template name="NumberFormat">
    <c:numFmt>
      <xsl:attribute name="formatCode">
        <xsl:choose>
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='general'">
            <xsl:value-of select="'General'"/>
          </xsl:when>
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='number'">
            <xsl:value-of select="'#,##0.00_);[Red]\(#,##0.00\)'"/>
          </xsl:when>
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='currency'">
            <xsl:value-of select="'&quot;￥&quot;#,##0.00_);[Red]\(&quot;￥&quot;#,##0.00\)'"/>
          </xsl:when>
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='accounting'">
            <xsl:value-of select="'_ &quot;￥&quot;* #,##0.00_ ;_ &quot;￥&quot;* \-#,##0.00_ ;_ &quot;￥&quot;* &quot;-&quot;??_ ;_ @_ '"/>
          </xsl:when>
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='date'">
            <!-- 20130518 update by xuzhenwei BUG_2881:回归集成oo-uof工作表"折线图"分类轴X轴转换前为年月，转换后为数字 start -->
            <xsl:choose>
              <xsl:when test="contains(./图表:数值_E70D/@格式码_E73F,'yyyy&quot;年&quot;m&quot;月&quot;')">
                <xsl:value-of select="'yyyy&quot;年&quot;m&quot;月&quot;;@'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'yyyy-m-d'"/>
              </xsl:otherwise>
            </xsl:choose>
            <!-- end -->
          </xsl:when>
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='time'">
            <xsl:value-of select="'[$-F400]h:mm:ss AM/PM'"/>
          </xsl:when>
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='percentage'">
            <xsl:value-of select="'0.00%'"/>
          </xsl:when>
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='fraction'">
            <xsl:value-of select="'# ?/?'"/>
          </xsl:when>
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='scientific'">
            <xsl:value-of select="'0.00E+00'"/>
          </xsl:when>
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='text'">
            <xsl:value-of select="'@'"/>
          </xsl:when>
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='specialization'">
            <xsl:value-of select="'000000'"/>
          </xsl:when>
          
          <!--2014-4-23, update by Qihy, bug3080，转换后横轴坐标表轴显示坐标不正确， start-->
          <!--<xsl:when test="./图表:数值_E70D/@分类名称_E740='custom'">
            <xsl:value-of select="'_ ￥* #,##0.00_ ;_ ￥* -#,##0.00_ ;_ ￥* &quot;-&quot;??_ ;_ @_ '"/>
          </xsl:when>-->
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='custom' and ./图表:数值_E70D[@格式码_E73F]">
            <xsl:value-of select="./图表:数值_E70D/@格式码_E73F"/>
          </xsl:when>
          <xsl:when test="./图表:数值_E70D/@分类名称_E740='custom' and not(./图表:数值_E70D[@格式码_E73F])">
            <xsl:value-of select="'_ ￥* #,##0.00_ ;_ ￥* -#,##0.00_ ;_ ￥* &quot;-&quot;??_ ;_ @_ '"/>
          </xsl:when>
          <!--2014-4-23 end-->
          
          <xsl:otherwise>
            <xsl:value-of select="'General'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <!--格式码还没转换-->
      <xsl:attribute name="sourceLinked">
        <xsl:choose>
          <xsl:when test="./图表:数值_E70D/@是否链接到源_E73E='true'">
            <xsl:value-of select="'1'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'0'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </c:numFmt>
  </xsl:template>

  <!--Common Codes======================================================================================================================-->

  <!--Modified by LDM in 2010/12/17-->
  <!--边框线阴影模板|common-->
  <!--设置为默认值，UOF中的阴影无法与OOXML完全对应-->
  <!--shadow-->
  <xsl:template name="shadow">
    <a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">
      <a:prstClr val="black">
        <a:alpha val="40%"/>
      </a:prstClr>
    </a:outerShdw>
  </xsl:template>
  <!--Modified by LDM in 2010/12/17-->
  <!--图案填充对应模板u2o|common-->
  <!--PatternMapping-->
  <xsl:template name="PatternMapping">
    <xsl:choose>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn001'">pct5</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn002'">pct10</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn003'">pct20</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn004'">pct25</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn005'">pct30</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn006'">pct40</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn007'">pct50</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn008'">pct60</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn009'">pct70</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn010'">pct75</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn011'">pct80</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn012'">pct90</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn013'">ltDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn014'">ltUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn015'">dkDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn016'">dkUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn017'">wdDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn018'">wdUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn019'">ltVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn020'">ltHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn021'">narVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn022'">narHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn023'">dkVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn024'">dkHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn025'">dashDnDiag</xsl:when>

      <xsl:when test=".//图:图案_800A/@类型_8008='ptn026'">dashUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn027'">dashHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn028'">dashVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn029'">smConfetti</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn030'">lgConfetti</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn031'">zigZag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn032'">wave</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn033'">diagBrick</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn034'">horzBrick</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn035'">weave</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn036'">plaid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn037'">divot</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn038'">dotGrid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn039'">dotDmnd</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn040'">shingle</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn041'">trellis</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn042'">sphere</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn043'">smGrid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn044'">lgGrid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn045'">smCheck</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn046'">lgCheck</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn047'">openDmnd</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn048'">solidDmnd</xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="PatternMapping2">
    <xsl:choose>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn001'">pct5</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn002'">pct50</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn003'">ltDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn004'">ltVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn005'">wave</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn006'">zigZag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn007'">divot</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn008'">smGrid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn009'">pct10</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn010'">pct60</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn011'">ltUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn012'">ltHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn013'">upDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn014'">wave</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn015'">dotGrid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn016'">lgGrid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn017'">pct20</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn018'">pct70</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn019'">dkDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn020'">narVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn021'">horz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn022'">diagBrick</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn023'">dotDmnd</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn024'">smCheck</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn025'">pct25</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn026'">pct75</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn027'">dkUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn028'">narHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn029'">vert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn030'">horzBrick</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn031'">dkDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn032'">lgCheck</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn033'">pct30</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn034'">pct80</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn035'">wdDnDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn036'">dkVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn037'">smConfetti</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn038'">weave</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn039'">trellis</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn040'">openDmnd</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn041'">pct40</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn042'">pct90</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn043'">wdUpDiag</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn044'">dkHorz</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn045'">lgConfetti</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn046'">plaid</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn047'">sphere</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn048'">solidDmnd</xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--Modified by ldm in 2010/12/18-->
  <!--Line-->
  <!--线型模板，与边框类型模板不同-->
  <xsl:template name="Line">
    <a:ln>
      <xsl:if test="./@宽度_C60F">
        <xsl:attribute name="w">
          <xsl:variable name="va">
            <xsl:value-of select="(./@宽度_C60F)*12700"/>
          </xsl:variable>
          <xsl:choose>
            <xsl:when test="contains($va,'-')">
              <xsl:value-of select="substring-after($va,'-')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="$va"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./@颜色_C611!='auto'">
        <a:solidFill>
          <a:srgbClr>
            <xsl:attribute name="val">
              <xsl:value-of select="substring-after(./@颜色_C611,'#')"/>
            </xsl:attribute>
          </a:srgbClr>
        </a:solidFill>
        </xsl:if>
        
      <!--uof命名空间中的线型，如边框等-->
      <xsl:if test="./@线型_C60D and ./@线型_C60D!='none'">
        <xsl:variable name="lineType">
          <xsl:value-of select="./@线型_C60D"/>
        </xsl:variable>
        <xsl:variable name="dashType">
          <xsl:value-of select="./@虚实_C60E"/>
        </xsl:variable>
        <xsl:call-template name="LineDashMapping_Border">
          <xsl:with-param name="lineType" select="$lineType"/>
          <xsl:with-param name="dashType" select="$dashType"/>
        </xsl:call-template>

      </xsl:if>
      <!--字命名空间中的线型，如下划线等-->
      <!--<xsl:if test="./@线型_C60D">
        <xsl:variable name="lineType">
          <xsl:value-of select="./@线型_C60D"/>
        </xsl:variable>
        <xsl:variable name="dashType">
          <xsl:value-of select="./@虚实_C60E"/>
        </xsl:variable>
        <xsl:call-template name="LineDashMapping_Border">
          <xsl:with-param name="dashType" select="$dashType"/>
        </xsl:call-template>
      </xsl:if>-->
    </a:ln>
    <!--<xsl:if test="./@uof:阴影='true'">
      <a:effectLst>
        <a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">
          <a:prstClr val="black">
            <a:alpha val="40000"/>
          </a:prstClr>
        </a:outerShdw>
      </a:effectLst>
    </xsl:if>-->
  </xsl:template>

  <xsl:template name="LineTypeMapping_Border">
    <xsl:param name ="lineType"/>
    <xsl:choose>
      <xsl:when test="$lineType = 'single'">
        <xsl:value-of select="'sng'"/>
      </xsl:when>
      <xsl:when test="$lineType = 'double'">
        <xsl:value-of select="'dbl'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--Modified by LDM in 2010/12/29-->
  <!--线形模板-->
  <!--LineTypeMapping_Border-->
  <xsl:template name="LineDashMapping_Border">
    <xsl:param name="dashType"/>
    <xsl:choose>
      <!--单线-->
      <xsl:when test ="not($dashType = 'none')">
        <a:prstDash>
          <xsl:attribute name ="val">
            <xsl:choose>
              <!--实线-->
              <xsl:when test="$dashType = 'noDash' or $dashType = 'solid'">
                <xsl:value-of select="'solid'"/>
              </xsl:when>
              <!--虚线-->
              <xsl:when test="$dashType = 'square-dot'">
                <xsl:value-of select="'dot'"/>
              </xsl:when>
              <!--划线-->
              <xsl:when test="$dashType = 'dash'">
                <xsl:value-of select="'dash'"/>
              </xsl:when>
              <!--长划线-->
              <xsl:when test="$dashType = 'long-dash'">
                <xsl:value-of select="'lgDash'"/>
              </xsl:when>
              <!--点划线-->
              <xsl:when test="$dashType = 'dash-dot'">
                <xsl:value-of select="'lgDashDot'"/>
              </xsl:when>
              <!--双点划线-->
              <xsl:when test="$dashType = 'dash-dot-dot'">
                <xsl:value-of select="'lgDashDotDot'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'solid'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </a:prstDash>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>