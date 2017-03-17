<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof" xmlns:v="urn:schemas-microsoft-com:vml"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram">
  <xsl:import href="table.xsl"/>
  <!--<xsl:import href ="table2.xsl"/>-->
  <xsl:import href="sheetfilter.xsl"/>
  <xsl:template name="worksheet">
    <xsl:choose>
      <xsl:when test="ws:worksheet">
        <xsl:variable name="s_name">
          <xsl:value-of select="ancestor::ws:sname[position()=1]"/>
        </xsl:variable>
        <表:工作表内容_E80E>
          <!--最大行-->
          <表:最大行列_E7E6>
            <xsl:attribute name="最大行_E7E7">
              <xsl:choose>
                <xsl:when test="ws:worksheet/ws:sheetData/ws:row">
                  <xsl:value-of select="ws:worksheet/ws:sheetData/ws:row[last()]/@r"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="1"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <!--最大列没有进行转换-->
            <!--xsl:attribute name="最大列_E7E8"></xsl:attribute-->
          </表:最大行列_E7E6>

          <!--单位全部用毫米-->
          <表:缺省行高列宽_E7E9>
            <xsl:if test="ws:worksheet/ws:sheetFormatPr[@defaultRowHeight]">
              <xsl:attribute name="缺省行高_E7EA">
                <xsl:value-of select="ws:worksheet/ws:sheetFormatPr/@defaultRowHeight"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="ws:worksheet/ws:sheetFormatPr[@defaultColWidth]">
              <xsl:variable name="x1">
                <xsl:value-of select="ws:worksheet/ws:sheetFormatPr/@defaultColWidth"/>
              </xsl:variable>
              <xsl:attribute name="缺省列宽_E7EB">
                <xsl:value-of select="$x1  div 8.38 * 54"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="ws:worksheet/ws:sheetFormatPr[not(@defaultColWidth)]">
              <xsl:attribute name="缺省列宽_E7EB">
                <xsl:value-of select="54.0"/>
              </xsl:attribute>
            </xsl:if>
          </表:缺省行高列宽_E7E9>


          <xsl:apply-templates select="ws:worksheet/ws:cols"/>
          <xsl:apply-templates select="ws:worksheet/ws:sheetData" mode="row"/>
          <!-- add  -->
          <xsl:apply-templates select="ws:worksheet/ws:scenarios"/>
          <!-- end -->
          
          <!--uof：锚点(零到无穷个)[not(xdr:graphicFrame)]-->
          <xsl:if test="ws:Drawings">
            <xsl:for-each select=".//xdr:twoCellAnchor">
              <xsl:choose >
                <xsl:when test ="contains(./xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name,'Diagram')">
                  <xsl:if test="ancestor::xdr:wsDr/dsp:drawing/dsp:spTree/dsp:sp">
                    <xsl:call-template name="graphiclocation"/>
                  </xsl:if>
                </xsl:when>
                <xsl:otherwise >
                  <xsl:call-template name="graphiclocation"/>
                </xsl:otherwise>
              </xsl:choose>

            </xsl:for-each>

            <xsl:for-each select="ws:Drawings/xdr:wsDr/xdr:oneCellAnchor">
              <xsl:variable name="pos" select="position()"/>
              <xsl:choose>
                <xsl:when test="./mc:AlternateContent//m:oMathPara">
                  <xsl:call-template name="equLocation">
                    <xsl:with-param name="pos" select="$pos"/>
                  </xsl:call-template>
                </xsl:when>
                <xsl:when test="not(xdr:graphicFrame) and not(mc:AlternateContent)">
                  <xsl:call-template name="graphiclocation"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="graphiclocation"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each>
          </xsl:if>

          <!--
					<xsl:if test="ws:worksheet/*[@outlineLevel]">
						<xsl:call-template name="DataGrouping"/>
					</xsl:if>
          -->
          <!--分组集 修改：李杨 2011-11-24-->
          <!--<xsl:if test=".//GroupSet">-->
          <xsl:if test ="ws:worksheet/ws:cols/ws:col[@outlineLevel] or ws:worksheet/ws:sheetData/ws:row[@outlineLevel]">
            <xsl:call-template name="GroupSet"/>
          </xsl:if>
          <!--</xsl:if>-->
          <!--<xsl:if test=".//ws:autoFilter">
            <xsl:call-template name="sheetfilter"/>
          </xsl:if>-->
        </表:工作表内容_E80E>
      </xsl:when>
      <!---->
      <xsl:when test ="ws:chartsheet">
        <xsl:if test ="ws:Drawings/xdr:wsDr/pr:Relationships/pr:Relationship">
          <xsl:call-template name ="graphiclocation2"/>
        </xsl:if>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="GroupSet">
    <表:分组集_E7F6 标记是否在数据下方_E838="true" 标记是否在数据右侧_E839="true">
      <!--分组集 添加：李杨 2011-11-24-->
      <!--<xsl:if test =".//ws:row[@outlineLevel=1]">
        <表:列_E841>
          <xsl:for-each select =".//ws:col[@outlineLevel=1]">
            <xsl:if test ="./@outlineLevel='1' and position()='1'">
              <xsl:attribute name="起始_E73A">
                <xsl:value-of select="./@min -1"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="./@outlineLevel='1' and position()=last()">
              <xsl:attribute name="终止_E73B">
                <xsl:value-of select="./@max -1"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
          
          <xsl:attribute name="是否隐藏_E73C">
            <xsl:value-of select="'false'"/>
          </xsl:attribute>
        </表:列_E841>
      </xsl:if>

      <xsl:if test =".//ws:row[@outlineLevel=2]">
        <表:列_E841>
          <xsl:for-each select =".//ws:col[@outlineLevel=2]">
            <xsl:if test ="./@outlineLevel='2' and position()='1'">
              <xsl:attribute name="起始_E73A">
                <xsl:value-of select="./@min -1"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="./@outlineLevel='2' and position()=last()">
              <xsl:attribute name="终止_E73B">
                <xsl:value-of select="./@max -1"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>

          <xsl:attribute name="是否隐藏_E73C">
            <xsl:value-of select="'false'"/>
          </xsl:attribute>
        </表:列_E841>
      </xsl:if>

      <xsl:if test =".//ws:row[@outlineLevel=1]">
        <表:行_E842>
          <xsl:for-each select =".//ws:row[@outlineLevel='1']">
            <xsl:variable name ="last">
              <xsl:if test ="position()=last()">
                <xsl:value-of select ="@r"/>
              </xsl:if>
            </xsl:variable>
            <xsl:if test ="./@outlineLevel=1 and position() = '1'">
              <xsl:attribute name="起始_E73A">
                <xsl:value-of select ="@r -1"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:attribute name="终止_E73B">
              <xsl:value-of select="$last -1"/>
            </xsl:attribute>
          </xsl:for-each>
          
          <xsl:attribute name="是否隐藏_E73C">
            <xsl:value-of select="'false'"/>
          </xsl:attribute>
        </表:行_E842>
      </xsl:if>

      <xsl:if test =".//ws:row[@outlineLevel=2]">
        <表:行_E842>
          <xsl:for-each select =".//ws:row[@outlineLevel='2']">
            <xsl:variable name ="last">
              <xsl:if test ="position()=last()">
                <xsl:value-of select ="@r"/>
              </xsl:if>
            </xsl:variable>
            <xsl:if test ="./@outlineLevel=2 and position() = '1'">
              <xsl:attribute name="起始_E73A">
                <xsl:value-of select ="@r -1"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:attribute name="终止_E73B">
              <xsl:value-of select="$last -1"/>
            </xsl:attribute>
          </xsl:for-each>

          <xsl:attribute name="是否隐藏_E73C">
            <xsl:value-of select="'false'"/>
          </xsl:attribute>
        </表:行_E842>
      </xsl:if>-->

      <!--以下原有代码转换后有时候不对  李杨2011-11-25-->
      <xsl:for-each select=".//GroupSet/Group">
        <xsl:if test="./@Position='Col'">
          <表:列_E841>
            <xsl:attribute name="起始_E73A">
              <xsl:value-of select="./@Start"/>
            </xsl:attribute>
            <xsl:attribute name="终止_E73B">
              <xsl:value-of select="./@End"/>
            </xsl:attribute>
            <!-- update by xuzhenwei BUG_2517,2518：分组设置，分组位置错误 20130115 -->
            <xsl:if test="ancestor::ws:spreadsheet/ws:worksheet/ws:sheetData/col[@c &gt;= @Start and @c &lt;= @End]">
              <xsl:attribute name="是否隐藏_E73C">
                <xsl:value-of select="'false'"/>
              </xsl:attribute>
            </xsl:if>
            <!-- end -->
          </表:列_E841>
        </xsl:if>
        <xsl:if test="./@Position='Row'">
          <表:行_E842>
            <xsl:attribute name="起始_E73A">
              <xsl:value-of select="./@Start"/>
            </xsl:attribute>
            <xsl:attribute name="终止_E73B">
              <xsl:value-of select="./@End"/>
            </xsl:attribute>
            <!-- update by xuzhenwei BUG_2517,2518：分组设置，分组位置错误 20130115 -->
            <xsl:if test="ancestor::ws:spreadsheet/ws:worksheet/ws:sheetData/row[@r &gt;= @Start and @r &lt;= @End]">
              <xsl:attribute name="是否隐藏_E73C">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
            <!-- end -->
          </表:行_E842>
        </xsl:if>
      </xsl:for-each>
    </表:分组集_E7F6>
    <!--<表:组及分级显示 uof:locID="s0137" uof:attrList="标记是否数据下方 标记是否数据右侧" 表:标记是否数据右侧="true" 表:标记是否数据下方="true"/>-->
  </xsl:template>
  <xsl:template name="graphiclocation2">
    <uof:锚点_C644>
      <xsl:variable name="s_namenn">
        <xsl:value-of select="./@sheetName"/>
      </xsl:variable>
      <xsl:if test ="ws:Drawings/xdr:wsDr/xdr:absoluteAnchor/xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr">
        <xsl:variable name ="ppid">
          <xsl:value-of select ="ws:Drawings/xdr:wsDr/xdr:absoluteAnchor/xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@id"/>
        </xsl:variable>
        <xsl:attribute name="图形引用_C62E">
          <xsl:value-of select ="concat($s_namenn,'_OBJ0000',$ppid)"/>
        </xsl:attribute>
      </xsl:if>

      <uof:大小_C621 长_C604="301.0" 宽_C605="542.0"/>
    </uof:锚点_C644>
  </xsl:template>
  <xsl:template name="graphiclocation">
    <uof:锚点_C644>
      <!--uof:attrList="x坐标 y坐标 宽度 高度 图形引用 随动方式 缩略图 占位符">-->
      <xsl:variable name="s_namenn">
        <xsl:value-of select="ancestor::ws:spreadsheet/@sheetName"/>
      </xsl:variable>
      <xsl:if test="xdr:pic">
        <xsl:variable name="ppid">
          <xsl:value-of select="xdr:pic/xdr:nvPicPr/xdr:cNvPr/@id"/>
        </xsl:variable>
        <xsl:attribute name="图形引用_C62E">
          <!--3yue27ri 夏艳霞checked-->
          <xsl:value-of select="concat($s_namenn,'_OBJ0000',$ppid)"/>
          <!--xsl:value-of select="concat('OBJ',$ppid)"/-->
        </xsl:attribute>
        <!--yx,error图形引用:Sheet1_OBJ1只不过是一个图形引用只要找到就行，两个软件不一定要一一对应 10.17 15:52 -->
      </xsl:if>
      <xsl:if test="xdr:sp">
        <xsl:variable name="ppid">
          <xsl:value-of select="xdr:sp/xdr:nvSpPr/xdr:cNvPr/@id"/>
        </xsl:variable>
        <!--3yue28ri 夏艳霞checked-->
        <xsl:attribute name="图形引用_C62E">
          <xsl:value-of select="concat($s_namenn,'_OBJ0000',$ppid)"/>
        </xsl:attribute>
        <!--图形引用:Sheet1_OBJ1-->
      </xsl:if>
      <xsl:if test="xdr:cxnSp">
        <xsl:variable name="ppid">
          <xsl:value-of select="xdr:cxnSp/xdr:nvCxnSpPr/xdr:cNvPr/@id"/>
        </xsl:variable>
        <!--3yue28ri 夏艳霞checked-->
        <xsl:attribute name="图形引用_C62E">
          <xsl:value-of select="concat($s_namenn,'_OBJ0000',$ppid)"/>
        </xsl:attribute>
        <!--图形引用:Sheet1_OBJ1-->
      </xsl:if>
      <xsl:if test="./xdr:grpSp">
        <xsl:variable name="ppid">
          <xsl:value-of select="xdr:grpSp/xdr:nvGrpSpPr/xdr:cNvPr/@id"/>
        </xsl:variable>
        <!--3yue28ri 夏艳霞checked-->
        <xsl:attribute name="图形引用_C62E">
          <xsl:value-of select="concat($s_namenn,'_OBJ0000',$ppid)"/>
        </xsl:attribute>
        <!--图形引用:Sheet1_OBJ1-->
      </xsl:if>
      <!--添加图表的锚点属性：图形引用  李杨2011-12-02-->
      <xsl:if test ="./xdr:graphicFrame">
        <xsl:variable name ="ppid">
          <!-- 2013-01-08 xuzhenwei 新增功能点smartArt start -->
          <xsl:variable name="rcs" select="./xdr:graphicFrame/a:graphic/a:graphicData/dgm:relIds/@r:cs"/>
          <xsl:variable name="rels" select="substring($rcs,1,3)"/>
          <xsl:variable name="nums" select="substring-after($rcs,'rId')"/>
          <xsl:variable name="dsmRelid" select="concat($rels,$nums + 1)"/>
          <xsl:variable name="drawingId">
            <xsl:for-each select="parent::xdr:wsDr/pr:Relationships//pr:Relationship">
              <xsl:if test="@Id=$dsmRelid">
                <xsl:variable name="target" select="@Target"/>
                <xsl:value-of select="substring-after($target,'../diagrams/')"/>
              </xsl:if>
            </xsl:for-each>
          </xsl:variable>
          <!-- 20130506 update by xuzhenwei BUG_2909 第二轮回归 oo-uof SmartArt丢失 start -->
          <xsl:variable name="rcsname" select="./xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name"/>
          <xsl:choose>
            <xsl:when test="contains($rcsname,'图示') or contains($rcsname,'Diagram')">
              <xsl:value-of select="substring-before($drawingId,'.xml')"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select ="xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@id"/>
            </xsl:otherwise>
          </xsl:choose>
          <!-- end -->
          <!-- end -->
        </xsl:variable>
        <xsl:attribute name ="图形引用_C62E">
          <xsl:value-of select ="concat($s_namenn,'_OBJ0000',$ppid)"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:attribute name="随动方式_C62F">
        <!-- add by xuzhenwei BUG_2472:转换后文本框属性发生变化 2013-01-19  start -->
        <!-- <xsl:value-of select="'movesize'"/>-->
        <xsl:choose>
          <xsl:when test="@editAs='twoCell'">
            <xsl:value-of select="'move'"/>
          </xsl:when>
          <xsl:when test="@editAs='oneCell'">
            <xsl:value-of select="'movesize'"/>
          </xsl:when>
          <xsl:when test="@editAs='absolute'">
            <xsl:value-of select="'none'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'move'"/>
          </xsl:otherwise>
        </xsl:choose>
        <!-- end -->
      </xsl:attribute>

      <!--uof:x坐标uof:y坐标uof:宽度uof:高度-->
      <xsl:variable name="fc">
        <xsl:value-of select="xdr:from/xdr:col"/>
      </xsl:variable>
      <xsl:variable name="fr">
        <xsl:value-of select="xdr:from/xdr:row"/>
      </xsl:variable>
      <xsl:variable name="tc">
        <xsl:value-of select="xdr:to/xdr:col"/>
      </xsl:variable>
      <xsl:variable name="tr">
        <xsl:value-of select="xdr:to/xdr:row"/>
      </xsl:variable>
      <xsl:variable name="fco">
        <xsl:value-of select="xdr:from/xdr:colOff"/>
      </xsl:variable>
      <xsl:variable name="fro">
        <xsl:value-of select="xdr:from/xdr:rowOff"/>
      </xsl:variable>
      <xsl:variable name="tco">
        <xsl:value-of select="xdr:to/xdr:colOff"/>
      </xsl:variable>
      <xsl:variable name="tro">
        <xsl:value-of select="xdr:to/xdr:rowOff"/>
      </xsl:variable>
      <xsl:variable name="c">
        <xsl:value-of select="$tc - $fc"/>
      </xsl:variable>
      <xsl:variable name="r">
        <xsl:value-of select="$tr - $fr"/>
      </xsl:variable>
      <xsl:variable name="co">
        <xsl:value-of select="$tco - $fco"/>
      </xsl:variable>
      <xsl:variable name="ro">
        <xsl:value-of select="$tro - $fro"/>
      </xsl:variable>

      <uof:位置_C620>
        <uof:水平_4106>
          <uof:绝对_4107>
            <xsl:attribute name ="值_4108">
              <!-- 20130426 add by xuzhenwei BUG_2704:图片错位 start -->
              <!--<xsl:choose>
                  <xsl:when test="ancestor::ws:spreadsheet/ws:worksheet/ws:cols">
                    <xsl:variable name="cc" select="count(ancestor::ws:spreadsheet/ws:worksheet/ws:cols/ws:col[@min=@max and @min &lt;= ( $fc+1 )])"/>
                <xsl:variable name="sumdistant" select="sum(ancestor::ws:spreadsheet/ws:worksheet/ws:cols/ws:col[@min=@max and @min &lt;= ( $fc+1 )]/@width)"/>
                  <xsl:value-of select="floor(($fc + 1 - $cc) * 54  + $sumdistant * 54 div 8.38  + $fco * 28.3  div 3600000)"/> 
                  </xsl:when>
                  <xsl:otherwise>-->
              <!--xsl:value-of select="floor($fco)"/-->
              <xsl:value-of select="$fco"/>
              <!--</xsl:otherwise>
                </xsl:choose>-->
              <!-- end -->
            </xsl:attribute>
          </uof:绝对_4107>
        </uof:水平_4106>
        <uof:垂直_410D>
          <uof:绝对_4107>
            <xsl:attribute name ="值_4108">
              <!-- 20140415 add by lingfeng 锚点位置不准 start -->
              <xsl:value-of select ="$fro"/>
              <!--end-->
            </xsl:attribute>
          </uof:绝对_4107>
        </uof:垂直_410D>
      </uof:位置_C620>

      <!--<xsl:attribute name="uof:x坐标">
        <xsl:value-of select="floor($fc * 54 + $fco * 28.3 div 360000)"/>
      </xsl:attribute>
      <xsl:attribute name="uof:y坐标">
        <xsl:value-of select="floor($fr * 13.5 + $fro * 28.3 div 360000)"/>
      </xsl:attribute>-->

      <xsl:variable name="wit">
        <xsl:value-of select="$co"/>
      </xsl:variable>
      <xsl:variable name="hit">
        <xsl:value-of select="$ro"/>
      </xsl:variable>

      <uof:大小_C621>
        <xsl:attribute name="长_C604">
          <xsl:choose>
            <xsl:when test="xdr:sp/xdr:spPr/a:xfrm/a:ext/@cy">
              <xsl:value-of select="xdr:sp/xdr:spPr/a:xfrm/a:ext/@cy div 12700"/>
            </xsl:when>
            <xsl:when test="contains($hit,'-')">
              <xsl:value-of select="translate($hit,'-','')"/>
            </xsl:when>
            <!--yx,add,2010.5.6-->
            <xsl:otherwise>
              <xsl:value-of select="$hit"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <!--yx,all to long datatype,2010.4.5-->

        <xsl:attribute name="宽_C605">
          <xsl:choose>
            <xsl:when test="xdr:sp/xdr:spPr/a:xfrm/a:ext/@cx">
              <xsl:value-of select="round(xdr:sp/xdr:spPr/a:xfrm/a:ext/@cx div 12700)"/>
            </xsl:when>
            <!-- update by xuzhenwei uof_大小_C621 宽_C605的值修订 start -->
            <xsl:when test="xdr:pic/xdr:spPr/a:xfrm/a:ext/@cx">
              <xsl:value-of select="round(xdr:pic/xdr:spPr/a:xfrm/a:ext/@cx div 12700)"/>
            </xsl:when>
            <!-- end -->
            <xsl:when test="contains($wit,'-')">
              <xsl:value-of select="translate($wit,'-','')"/>
            </xsl:when>
            <!--yx,add,2010.5.6-->
            <xsl:otherwise>
              <xsl:value-of select="$wit"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </uof:大小_C621>

      <!--uof:x坐标uof:y坐标uof:宽度uof:高度-->
    </uof:锚点_C644>
  </xsl:template>

  <!--add by linyh equation-->
  <xsl:template name="equLocation">
    <xsl:param name="pos"/>
    <uof:锚点_C644>
      <!--uof:attrList="x坐标 y坐标 宽度 高度 图形引用 随动方式 缩略图 占位符">-->
      <xsl:variable name="s_namenn">
        <xsl:value-of select="ancestor::ws:spreadsheet/@sheetName"/>
      </xsl:variable>

      <xsl:attribute name="图形引用_C62E">
        <xsl:value-of select="concat($s_namenn,'_equ_',$pos)"/>
      </xsl:attribute>
      <xsl:if test="xdr:cxnSp">
        <xsl:attribute name="图形引用_C62E">
          <xsl:value-of select="concat($s_namenn,'_equ_',$pos)"/>
        </xsl:attribute>
        <!--图形引用:Sheet1_OBJ1-->
      </xsl:if>

      <!--uof:x坐标uof:y坐标uof:宽度uof:高度-->
      <xsl:variable name="fc">
        <xsl:value-of select="xdr:from/xdr:col"/>
      </xsl:variable>
      <xsl:variable name="fr">
        <xsl:value-of select="xdr:from/xdr:row"/>
      </xsl:variable>
      <xsl:variable name="tc">
        <xsl:value-of select="xdr:to/xdr:col"/>
      </xsl:variable>
      <xsl:variable name="tr">
        <xsl:value-of select="xdr:to/xdr:row"/>
      </xsl:variable>
      <xsl:variable name="fco">
        <xsl:value-of select="xdr:from/xdr:colOff"/>
      </xsl:variable>
      <xsl:variable name="fro">
        <xsl:value-of select="xdr:from/xdr:rowOff"/>
      </xsl:variable>
      <xsl:variable name="tco">
        <xsl:value-of select="xdr:to/xdr:colOff"/>
      </xsl:variable>
      <xsl:variable name="tro">
        <xsl:value-of select="xdr:to/xdr:rowOff"/>
      </xsl:variable>
      <xsl:variable name="c">
        <xsl:value-of select="$tc - $fc"/>
      </xsl:variable>
      <xsl:variable name="r">
        <xsl:value-of select="$tr - $fr"/>
      </xsl:variable>
      <xsl:variable name="co">
        <xsl:value-of select="$tco - $fco"/>
      </xsl:variable>
      <xsl:variable name="ro">
        <xsl:value-of select="$tro - $fro"/>
      </xsl:variable>

      <uof:位置_C620>
        <uof:水平_4106>
          <uof:绝对_4107>
            <xsl:attribute name ="值_4108">
              <!-- 20130426 add by xuzhenwei BUG_2704:图片错位 start -->
              <!--<xsl:choose>
                  <xsl:when test="ancestor::ws:spreadsheet/ws:worksheet/ws:cols">
                    <xsl:variable name="cc" select="count(ancestor::ws:spreadsheet/ws:worksheet/ws:cols/ws:col[@min=@max and @min &lt;= ( $fc+1 )])"/>
                <xsl:variable name="sumdistant" select="sum(ancestor::ws:spreadsheet/ws:worksheet/ws:cols/ws:col[@min=@max and @min &lt;= ( $fc+1 )]/@width)"/>
                  <xsl:value-of select="floor(($fc + 1 - $cc) * 54  + $sumdistant * 54 div 8.38  + $fco * 28.3  div 3600000)"/> 
                  </xsl:when>
                  <xsl:otherwise>-->
              <!--xsl:value-of select="floor($fco)"/-->
              <!--<xsl:value-of select="$fco"/>-->

              <xsl:value-of select=".//a:off/@x div 12700"/>
              <!--</xsl:otherwise>
                </xsl:choose>-->
              <!-- end -->
            </xsl:attribute>
          </uof:绝对_4107>
        </uof:水平_4106>
        <uof:垂直_410D>
          <uof:绝对_4107>
            <xsl:attribute name ="值_4108">
              <!-- 20140415 add by lingfeng 锚点位置不准 start -->
              <!--<xsl:value-of select ="$fro"/>-->
              <xsl:value-of select=".//a:off/@y div 12700"/>
              <!--end-->
            </xsl:attribute>
          </uof:绝对_4107>
        </uof:垂直_410D>
      </uof:位置_C620>

      <!--<xsl:attribute name="uof:x坐标">
        <xsl:value-of select="floor($fc * 54 + $fco * 28.3 div 360000)"/>
      </xsl:attribute>
      <xsl:attribute name="uof:y坐标">
        <xsl:value-of select="floor($fr * 13.5 + $fro * 28.3 div 360000)"/>
      </xsl:attribute>-->

      <xsl:variable name="wit">
        <xsl:value-of select="$co"/>
      </xsl:variable>
      <xsl:variable name="hit">
        <xsl:value-of select="$ro"/>
      </xsl:variable>

      <uof:大小_C621>
        <xsl:attribute name="长_C604">
          <xsl:value-of select=".//a:ext/@cx div 12700"/>
        </xsl:attribute>
        <!--yx,all to long datatype,2010.4.5-->

        <xsl:attribute name="宽_C605">
          <xsl:value-of select=".//a:ext/@cy div 12700"/>
        </xsl:attribute>
      </uof:大小_C621>

      <!--uof:x坐标uof:y坐标uof:宽度uof:高度-->
    </uof:锚点_C644>
  </xsl:template>


  <xsl:template match="ws:cols">
    <xsl:if test="ws:col">
      <xsl:for-each select="ws:col">
        <!--行列中的 式样引用，是否自适应列宽上期没转换  李杨2011.11.7-->
        <表:列_E7EC>

          <xsl:if test="@min">
            <xsl:attribute name="列号_E7ED">
              <xsl:value-of select="@min"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@hidden">
            <xsl:attribute name="是否隐藏_E73C">
              <xsl:if test="@hidden='1'">
                <xsl:value-of select="'true'"/>
              </xsl:if>
              <xsl:if test="@hidden='0'">
                <xsl:value-of select="'false'"/>
              </xsl:if>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="not(@hidden)">
            <xsl:attribute name="是否隐藏_E73C">
              <xsl:value-of select="'false'"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test ="not(@min and @max) or (@min = @max)">
            <xsl:if test="@width">
              <xsl:attribute name="列宽_E7EE">

                <!--2012-5-2, update by 凌峰 列宽计算错误， start-->
                <xsl:value-of select="@width * 6"/>
                <!--2012-5-2 end-->

              </xsl:attribute>
            </xsl:if>
            <xsl:if test ="@bestFit and @bestFit='1'">
              <xsl:attribute name ="是否自适应列宽_E7F0">
                <xsl:value-of select ="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:if>
          <xsl:if test="@min and @max and (@min != @max)">
            <!--<xsl:attribute name ="式样引用_E7BD">DEFSTYLE</xsl:attribute>-->

            <!--2014-3-26, add by Qihy, 修复bug3152中单元格宽度不正确， start-->
            <xsl:if test="@width">
              <xsl:attribute name="列宽_E7EE">
                <xsl:value-of select="@width div 8.38 * 54"/>
              </xsl:attribute>
            </xsl:if>
            <!--2014-3-26 end-->

            <xsl:attribute name="跨度_E7EF">
              <xsl:value-of select="@max - @min"/>
            </xsl:attribute>
            <xsl:attribute name ="是否自适应列宽_E7F0">true</xsl:attribute>
          </xsl:if>

          <!--2014-6-9, add by Qihy, 修复单元格拉伸时单元格式样不正确导致单元格宽度变化和图表图形变化等问题， start-->
          <xsl:if test="@style">
            <xsl:attribute name="式样引用_E7BD">
              <xsl:value-of select="concat('CELLSTYLE_', number(@style)+1)"/>
            </xsl:attribute>
          </xsl:if>
          <!--2014-6-9 end-->

        </表:列_E7EC>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
  <xsl:template match="ws:sheetData" mode="row">
    <xsl:if test="ws:row">
      <xsl:for-each select="ws:row">
        <!-- add by xuzhenwei BUG_2468:转换后一部分批注丢失 2013-01-20 start -->
        <xsl:variable name="cc">
          <xsl:for-each select="ws:c">
            <xsl:value-of select="concat(' ',@r)"/>
          </xsl:for-each>
        </xsl:variable>
        <!-- end -->
        <!-- 统计有多少列 即有多少个cols/col节点-->
        <xsl:variable name="rownumber" select="@r"/>
        <表:行_E7F1>
          <!--行属性-->
          <xsl:if test="@r">
            <xsl:attribute name="行号_E7F3">
              <xsl:value-of select="@r"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@hidden">
            <xsl:attribute name="是否隐藏_E73C">
              <xsl:if test="@hidden=1">
                <xsl:value-of select="'true'"/>
              </xsl:if>
              <xsl:if test="@hidden=0">
                <xsl:value-of select="'false'"/>
              </xsl:if>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="@ht">
            <xsl:attribute name="行高_E7F4">
              <xsl:value-of select="@ht"/>
            </xsl:attribute>
          </xsl:if>
          <!--UOF的单元格-->
          <!-- add by xuzhenwei BUG_2462 单元格列合并效果丢失 2013-01-16 start 20130308再修复-->
          <xsl:if test="ws:c">            
            <xsl:for-each select="ws:c">
              <xsl:variable name="cell_id" select="@r"/>
              <表:单元格_E7F2>
                <!--设置单元格属性-->
                <xsl:variable name="id" select="@r"/>
                <xsl:variable name="iid" select="translate($id,'1234567890','')"/>
                <xsl:attribute name="列号_E7BC">
                  <!-- add by xuzhenwei BUG_2462 单元格列合并效果丢失 2013-01-16 -->
                  <xsl:call-template name="Ts2Dec">
                    <xsl:with-param name="tsSource" select="$iid"/>
                  </xsl:call-template>
                  <!-- end -->
                </xsl:attribute>
                <xsl:if test="name(@s)">
                  <!--3yue28ri 夏艳霞checked-->
                  <xsl:attribute name="式样引用_E7BD">
                    <xsl:variable name="s" select="@s"/>
                    <xsl:variable name="ss" select="$s + 1"/>
                    <xsl:value-of select="concat('CELLSTYLE_',$ss)"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="not(@s)">
                  <xsl:attribute name="式样引用_E7BD">

                    <!--2014-5-6, update by Qihy, 修复bug3270字体不正确， start-->
                    <!--<xsl:value-of select="'DEFSTYLE'"/>-->
                    <xsl:value-of select="'CELLSTYLE_1'"/>
                    <!--2014-5-6 end-->

                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="ancestor::ws:worksheet[ws:hyperlinks]">
                  <xsl:for-each select="ancestor::ws:worksheet/ws:hyperlinks/ws:hyperlink">
                    <xsl:variable name="xtix">
                      <xsl:value-of select="@ref"/>
                    </xsl:variable>
                    <xsl:variable name="ppos12">
                      <xsl:value-of select="position()"/>
                    </xsl:variable>
                    <xsl:variable name ="sheetName">
                      <xsl:value-of select ="../../../@sheetName"/>
                    </xsl:variable>
                    <xsl:if test="$xtix=$cell_id">
                      <xsl:attribute name="超链接引用_E7BE">
                        <xsl:value-of select="concat($sheetName,'_Hyperlink_',$ppos12)"/>
                      </xsl:attribute>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:if>
                <!-- update by xuzhenwei 20130304 BUG_2462:单元格列合并效果丢失 start -->
                <xsl:for-each select="ancestor::ws:worksheet/ws:mergeCells/ws:mergeCell">
                  <xsl:variable name="mergeref" select="@ref"/>
                  <xsl:variable name="mss" select="substring-before($mergeref,':')"/>
                  <xsl:if test="$mss=$id">
                    <表:合并_E7AF>
                      <xsl:call-template name="hbdyg">
                        <xsl:with-param name="cid" select="$rownumber"/>
                        <xsl:with-param name="mergeref" select="$mergeref"/>
                      </xsl:call-template>
                    </表:合并_E7AF>
                  </xsl:if>
                </xsl:for-each>
                <!-- end -->
                <xsl:variable name="style">
                  <xsl:choose>
                    <xsl:when test="@t and @t = 'b'">
                      <xsl:value-of select="'boolean'"/>
                    </xsl:when>
                    <xsl:when test="@t and @t = 'e'">
                      <xsl:value-of select="'error'"/>
                    </xsl:when>
                    <xsl:when test="@t and @t = 'n'">
                      <xsl:value-of select="'number'"/>
                    </xsl:when>
                    <xsl:when test="@t = 's' or @t = 'str' or @t = 'inlineStr'">
                      <xsl:value-of select="'text'"/>
                    </xsl:when>
                  </xsl:choose>
                </xsl:variable>
                <xsl:if test="ws:v">
                  <表:数据_E7B3>
                    <xsl:if test="name(@t)">
                      <xsl:attribute name="类型_E7B6">
                        <xsl:value-of select="$style"/>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:if test="not(name(@t))">
                      <xsl:attribute name="类型_E7B6">
                        <xsl:value-of select="'number'"/>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:if test="@t = 's' or @t = 'str' or @t = 'inlineStr'">
					
					  <!--zl 2015-4-20-->
                      <xsl:if test="ancestor::ws:spreadsheet/ws:worksheet/ws:sheetData/ws:row/ws:c/ws:f[@t = 'array']">
                        <字:句_419D >
                          <字:文本串_415B>
                            <xsl:value-of select="ws:v"/>
                          </字:文本串_415B>
                        </字:句_419D>
                      </xsl:if>
                      <!--zl 2015-4-20-->
					
                      <xsl:variable name="snumber">
                        <xsl:value-of select="ws:v +1"/>
                      </xsl:variable>
                      <xsl:if test="/ws:spreadsheets/ws:sst/ws:si[position()=$snumber]">
                        <xsl:for-each select="/ws:spreadsheets/ws:sst/ws:si[position()=$snumber]">
                          <xsl:choose>
                            <xsl:when test="ws:t">
                              <!--single row-->
                              <xsl:variable name="string">
                                <xsl:value-of select="ws:t"/>
                              </xsl:variable>
                              <字:句_419D>
                                <字:文本串_415B >
                                  <xsl:value-of select="$string"/>
                                </字:文本串_415B>
                              </字:句_419D>
                            </xsl:when>
                            <xsl:when test="ws:r">
                              <xsl:variable name="sloc">

                                <!--2014-6-9, update by Qihy, 增加sharedString.xml中字体的转换， start-->
                                <xsl:value-of select="$snumber"/>
                                <!--2014-6-9 end-->

                              </xsl:variable>
                              <xsl:for-each select="ws:r">

                                <!--2014-4-26, delete by Qihy, 凌峰增加部分不正确，not(./ws:rPr)情况下，字体样式应该从样式表中获取，而不是从下一个句中获取 start-->
                                <!-- update by 凌峰 BUG_3108:集成测试中,部分单元格内容丢失(D3)，单元格背景色丢失  20140314 start -->
                                <!--<xsl:variable name="pos">
                                          <xsl:value-of select="position()"/>
                                      </xsl:variable>-->
                                <字:句_419D>
                                  <!--<xsl:if test="not(./ws:rPr)">
                                            <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr">
                                                <字:句属性_4158>
                                                    <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:rFont">
                                                        <字:字体_4128>
                                                            <xsl:variable name="fid">
                                                                <xsl:value-of select="position()"/>
                                                            </xsl:variable>
                                                            <xsl:variable name="ffid">
                                                                <xsl:value-of select="position()"/>
                                                            </xsl:variable>
                                                            <xsl:attribute name="西文字体引用_4129">
                                                                <xsl:value-of select="concat('font:',$sloc,$fid)"/>
                                                            </xsl:attribute>
                                                            <xsl:attribute name="中文字体引用_412A">
                                                                <xsl:value-of select="concat('font:',$sloc,$fid)"/>
                                                            </xsl:attribute>
                                                            -->
                                  <!--字号 -->
                                  <!--
                                                            <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:sz">
                                                                <xsl:attribute name="字号_412D">
                                                                    <xsl:value-of select="../ws:r[position() = $pos+1]/ws:rPr/ws:sz/@val"/>
                                                                </xsl:attribute>
                                                            </xsl:if>
                                                            -->
                                  <!--color-->
                                  <!--
                                                            <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:color">
                                                                <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:color[@theme]">
                                                                    <xsl:variable name="the" select="ws:rPr/ws:color/@theme"/>
                                                                    <xsl:variable name="them" select="ws:spreadsheets/a:theme/a:themeElements/a:clrScheme/*[position()=$the]/a:sysClr/@val"/>
                                                                    <xsl:variable name="them2" select="ws:spreadsheets/a:theme/a:themeElements/a:clrScheme/*[position()=$the]/a:srgbClr/@val"/>
                                                                    <xsl:attribute name="颜色_412F">
                                                                        <xsl:choose>
                                                                            <xsl:when test="$them='windowText'">
                                                                                <xsl:value-of select="'#000000'"/>
                                                                            </xsl:when>
                                                                            <xsl:when test="$them='window'">
                                                                                <xsl:value-of select="'#ffffff'"/>
                                                                            </xsl:when>
                                                                            <xsl:otherwise>
                                                                                <xsl:variable name="colll">
                                                                                    <xsl:value-of select="translate($them2,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                                                                                </xsl:variable>
                                                                                -->
                                  <!--有问题！取不出来！李杨2012-5-10-->
                                  <!--
                                                                                <xsl:value-of select="concat('#','$them')"/>
                                                                            </xsl:otherwise>
                                                                        </xsl:choose>
                                                                    </xsl:attribute>
                                                                </xsl:if>
                                                                <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:color[@rgb]">
                                                                    <xsl:attribute name="颜色_412F">
                                                                        <xsl:variable name="c">
                                                                            <xsl:value-of select="../ws:r[position() = $pos+1]/ws:rPr/ws:color/@rgb"/>
                                                                        </xsl:variable>
                                                                        <xsl:variable name="co">
                                                                            <xsl:value-of select="substring($c,3,8)"/>
                                                                        </xsl:variable>
                                                                        <xsl:variable name="coll">
                                                                            <xsl:value-of select="translate($co,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                                                                        </xsl:variable>
                                                                        <xsl:value-of select="concat('#',$coll)"/>
                                                                    </xsl:attribute>
                                                                </xsl:if>
                                                            </xsl:if>
                                                        </字:字体_4128>
                                                    </xsl:if>
                                                    <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:b">
                                                        <字:是否粗体_4130>true</字:是否粗体_4130>
                                                    </xsl:if>
                                                    <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:i">
                                                        <字:是否斜体_4131>true</字:是否斜体_4131>
                                                    </xsl:if>
                                                    <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:strike">
                                                        <字:删除线_4135>single</字:删除线_4135>
                                                    </xsl:if>
                                                    <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr[ws:u]">
                                                        <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:u[@val!='none']">
                                                            -->
                                  <!--下划线中属性：线型，虚实，颜色，是否字下划线  李杨2011-11-9-->
                                  <!--
                                                            <字:下划线_4136>
                                                                -->
                                  <!--uof:attrList="类型 颜色 字下划线"-->
                                  <!--
                                                                -->
                                  <!--<xsl:attribute name="字:类型">-->
                                  <!--
                                                                <xsl:attribute name="线型_4137">
                                                                    <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:u[@val='double' or @val='doubleAccounting' ]">
                                                                        <xsl:value-of select="'double'"/>
                                                                    </xsl:if>
                                                                    <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:u[@val='single' or @val='singleAccounting' ]">
                                                                        <xsl:value-of select="'single'"/>
                                                                    </xsl:if>
                                                                </xsl:attribute>
                                                            </字:下划线_4136>
                                                        </xsl:if>
                                                    </xsl:if>
                                                    <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:outline">
                                                        <字:是否空心_413E>true</字:是否空心_413E>
                                                    </xsl:if>
                                                    <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:shadow">
                                                        <字:是否阴影_4140>true</字:是否阴影_4140>
                                                    </xsl:if>
                                                    <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:vertAlign">
                                                        <字:上下标类型_4143>
                                                            <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:vertAlign/@val='superscript'">
                                                                <xsl:value-of select="'sup'"/>
                                                            </xsl:if>
                                                            <xsl:if test="../ws:r[position() = $pos+1]/ws:rPr/ws:vertAlign/@val='subscript'">
                                                                <xsl:value-of select="'sub'"/>
                                                            </xsl:if>
                                                        </字:上下标类型_4143>
                                                    </xsl:if>
                                                </字:句属性_4158>
                                            </xsl:if>
                                        </xsl:if>-->
                                  <!--end-->
                                  <!--2014-4-26 end-->

                                  <xsl:if test="ws:rPr">
                                    <!--2014-3-12, update by Qihy, 修复BUG3061 ooxml-uof集成测试中,区域选择丢失(20102013)，G2内容丢失， start-->
                                    <!--<字:句属性_4158 式样引用_417B="">-->
                                    <字:句属性_4158>
                                      <!--2014-3-12 end-->
                                      <xsl:if test="ws:rPr/ws:rFont">
                                        <字:字体_4128>
                                          <xsl:variable name="fid">
                                            <xsl:value-of select="position()"/>
                                          </xsl:variable>
                                          <xsl:variable name="ffid">
                                            <xsl:value-of select="position()"/>
                                          </xsl:variable>

                                          <!--2014-6-9, update by Qihy, 增加sharedString.xml中字体的转换， start-->
                                          <xsl:attribute name="西文字体引用_4129">
                                            <xsl:value-of select="concat('font_',$sloc,'_',$fid)"/>
                                          </xsl:attribute>
                                          <xsl:attribute name="中文字体引用_412A">
                                            <xsl:value-of select="concat('font_',$sloc,'_',$fid)"/>
                                          </xsl:attribute>
                                          <!--2014-6-9 end-->

                                          <!--字号 -->
                                          <xsl:if test="ws:rPr/ws:sz">
                                            <xsl:attribute name="字号_412D">
                                              <xsl:value-of select="ws:rPr/ws:sz/@val"/>
                                            </xsl:attribute>
                                          </xsl:if>
                                          <!--color-->
                                          <xsl:if test="ws:rPr/ws:color">
                                            <xsl:if test="ws:rPr/ws:color[@theme]">
                                              <xsl:variable name="the" select="ws:rPr/ws:color/@theme"/>

                                              <!--2014-6-4, update by Qihy, 字体颜色取值不正确， start-->
                                              <xsl:variable name="tint" select="ws:rPr/ws:color/@tint"/>
                                              <!--2014-6-3, update by Qihy, 字体颜色取值不正确， start-->
                                              <!--<xsl:variable name="them" select="ws:spreadsheets/a:theme/a:themeElements/a:clrScheme/*[position()=$the]/a:sysClr/@val"/>
                                                  <xsl:variable name="them2" select="ws:spreadsheets/a:theme/a:themeElements/a:clrScheme/*[position()=$the]/a:srgbClr/@val"/>
                                                  <xsl:attribute name="颜色_412F">
                                                    <xsl:choose>
                                                      <xsl:when test="$them='windowText'">
                                                        <xsl:value-of select="'#000000'"/>
                                                      </xsl:when>
                                                      <xsl:when test="$them='window'">
                                                        <xsl:value-of select="'#FFFFFF'"/>
                                                      </xsl:when>
                                                      <xsl:otherwise>
                                                        <xsl:variable name="colll">
                                                          <xsl:value-of select="translate($them2,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                                                        </xsl:variable>
                                                        -->
                                              <!--有问题！取不出来！李杨2012-5-10-->
                                              <!--
                                                        <xsl:value-of select="concat('#',$them)"/>
                                                      </xsl:otherwise>
                                                    </xsl:choose>-->
                                              <xsl:attribute name="颜色_412F">
                                                <xsl:choose>
                                                  <xsl:when test="$the = '0' and $tint='-0.249977111117893'">
                                                    <xsl:value-of select="'#c0c0c0'"/>
                                                  </xsl:when>
                                                  <xsl:when test="$the = '0' and $tint != '-0.249977111117893'">
                                                    <xsl:value-of select="'#000000'"/>
                                                  </xsl:when>
                                                  <!--2014-6-4 end-->

                                                  <xsl:when test="$the = '1'">
                                                    <xsl:value-of select="'#000000'"/>
                                                  </xsl:when>
                                                  <xsl:when test="$the = '2'">
                                                    <xsl:value-of select="'#f5f5f8'"/>
                                                  </xsl:when>
                                                  <xsl:when test ="$the = '3'">
                                                    <xsl:value-of select ="'#082e54'"/>
                                                  </xsl:when>
                                                  <xsl:when test ="$the = '7'">
                                                    <xsl:value-of select ="'#e8e0ff'"/>
                                                  </xsl:when>
                                                  <xsl:when test ="$the= '9'">
                                                    <xsl:value-of select ="'#F6E7DC'"/>
                                                  </xsl:when>
                                                </xsl:choose>
                                              </xsl:attribute>
                                              <!--2014-6-3 end-->

                                            </xsl:if>
                                            <xsl:if test="ws:rPr/ws:color[@rgb]">
                                              <xsl:attribute name="颜色_412F">
                                                <xsl:variable name="c">
                                                  <xsl:value-of select="ws:rPr/ws:color/@rgb"/>
                                                </xsl:variable>
                                                <xsl:variable name="co">
                                                  <xsl:value-of select="substring($c,3,8)"/>
                                                </xsl:variable>
                                                <xsl:variable name="coll">
                                                  <xsl:value-of select="translate($co,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"/>
                                                </xsl:variable>
                                                <xsl:value-of select="concat('#',$coll)"/>
                                              </xsl:attribute>
                                            </xsl:if>

                                            <!--2014-6-9, add by Qihy, 字体颜色转换不正确问题， start-->
                                            <xsl:if test="ws:rPr/ws:color[@indexed]">
                                              <xsl:attribute name="颜色_412F">
                                                <xsl:call-template name="indexColor">
                                                  <xsl:with-param name="indexed" select="ws:rPr/ws:color/@indexed"/>
                                                </xsl:call-template>
                                              </xsl:attribute>
                                            </xsl:if>
                                            <!--2014-6-9 end-->

                                          </xsl:if>
                                        </字:字体_4128>
                                      </xsl:if>
                                      <xsl:if test="ws:rPr/ws:b">
                                        <字:是否粗体_4130>true</字:是否粗体_4130>
                                      </xsl:if>
                                      <xsl:if test="ws:rPr/ws:i">
                                        <字:是否斜体_4131>true</字:是否斜体_4131>
                                      </xsl:if>
                                      <xsl:if test="ws:rPr/ws:strike">
                                        <字:删除线_4135>single</字:删除线_4135>
                                      </xsl:if>
                                      <xsl:if test="ws:rPr[ws:u]">
                                        <xsl:if test="ws:rPr/ws:u[@val!='none']">
                                          <!--下划线中属性：线型，虚实，颜色，是否字下划线  李杨2011-11-9-->
                                          <字:下划线_4136>
                                            <!--uof:attrList="类型 颜色 字下划线"-->
                                            <!--<xsl:attribute name="字:类型">-->
                                            <xsl:attribute name="线型_4137">
                                              <xsl:if test="ws:rPr/ws:u[@val='double' or @val='doubleAccounting' ]">
                                                <xsl:value-of select="'double'"/>
                                              </xsl:if>
                                              <xsl:if test="ws:rPr/ws:u[@val='single' or @val='singleAccounting' ]">
                                                <xsl:value-of select="'single'"/>
                                              </xsl:if>
                                            </xsl:attribute>
                                          </字:下划线_4136>
                                        </xsl:if>
                                      </xsl:if>
                                      <xsl:if test="ws:rPr/ws:outline">
                                        <字:是否空心_413E>true</字:是否空心_413E>
                                      </xsl:if>
                                      <xsl:if test="ws:rPr/ws:shadow">
                                        <字:是否阴影_4140>true</字:是否阴影_4140>
                                      </xsl:if>
                                      <xsl:if test="ws:rPr/ws:vertAlign">
                                        <字:上下标类型_4143>
                                          <xsl:if test="ws:rPr/ws:vertAlign/@val='superscript'">
                                            <xsl:value-of select="'sup'"/>
                                          </xsl:if>
                                          <xsl:if test="ws:rPr/ws:vertAlign/@val='subscript'">
                                            <xsl:value-of select="'sub'"/>
                                          </xsl:if>
                                        </字:上下标类型_4143>
                                      </xsl:if>
                                    </字:句属性_4158>
                                  </xsl:if>
                                  <字:文本串_415B >
                                    <xsl:value-of select="ws:t"/>
                                  </字:文本串_415B>
                                </字:句_419D>
                              </xsl:for-each>
                            </xsl:when>
                          </xsl:choose>
                        </xsl:for-each>
                      </xsl:if>
                    </xsl:if>
                    <xsl:if test="@t = 'b' or @t = 'e' or @t = 'n'">
                      <字:句_419D>
                        <字:文本串_415B>
                          <xsl:value-of select="ws:v"/>
                        </字:文本串_415B>
                      </字:句_419D>
                    </xsl:if>
                    <xsl:if test="not(@t)">
                      <字:句_419D>
                        <字:文本串_415B>
                          <xsl:value-of select="ws:v"/>
                        </字:文本串_415B>
                      </字:句_419D>
                    </xsl:if>

                    <!--2014-5-6, uodate by Qihy, 修复bug3297互操作测试转换值出错，uof不支持公式SUM(IF(ISTEXT(C6:K6),1,0)，因此直接转换为数值， start-->
                    <!--<xsl:if test="ws:f and ws:f!=''and not(contains(ws:f,'['))">-->

                    <!--2014-5-31, update by Qihy, 修复bug3335中的026数据有效性-条件格式化-区域公式集-筛选-2013.xlsx中公式丢失问题（RC引用C#处理后在公式中增加了'['）， start-->
                    <!--<xsl:if test="ws:f and ws:f!=''and not(contains(ws:f,'[')) and not(ws:f[@t ='array'])">-->

                    <!--2014-6-4, update by Qihy, 修复bug3335公式丢失问题，原来直接转换为数值，现在把公式还原 start-->
                    <!--<xsl:if test="ws:f and ws:f!=''and (contains(ws:f,'R[') or contains(ws:f,'C[') or not(contains(ws:f,'['))) and not(ws:f[@t ='array'])">-->
                    <xsl:if test="ws:f and ws:f!=''and (contains(ws:f,'R[') or contains(ws:f,'C[') or not(contains(ws:f,'[')))">
                      <!--2014-6-4 end-->

                      <!--2014-5-31 end-->

                      <!--2014-5-6 end-->

                      <xsl:if test="ws:f[not(@ca)]">
                        <!--添加。李杨2012-3-19-->
                        <表:公式_E7B5>
                          <xsl:choose >
                            <xsl:when test ="contains(ws:f,'_xlfn.')">
                              <xsl:variable name ="ff">
                                <xsl:value-of select ="translate(ws:f,'_xlfn.','')"/>
                              </xsl:variable>
                              <xsl:if test ="contains(substring-before($ff,'('),'.')">
                                <xsl:value-of select ="concat('=',translate($ff,'.',''))"/>
                              </xsl:if>
                              <xsl:if test ="not(contains($ff,'.'))">
                                <xsl:value-of select ="concat('=',$ff)"/>
                              </xsl:if>
                            </xsl:when>
                            <!--  update by xuzhenwei BUG_2660:公式不正确， 2013-01-21 start 备注：暂时先注释掉 影响了FLOOR(3.14，1)的计算结果 
                            <xsl:when test ="not(contains(ws:f,'_xlfn.')) and contains(ws:f,'.')">
                              <xsl:value-of select ="concat('=',translate(ws:f,'.',''))"/>
                            </xsl:when>-->
                            <!-- end -->
                            <!--文本函数：DOLLAR-->
                            <xsl:when test ="contains(substring-before(ws:f,'('),'USDOLLAR')">
                              <xsl:value-of select ="concat('=',translate(ws:f,'US',''))"/>
                            </xsl:when>
                            <!--文本函数：WIDECHAR-->
                            <xsl:when test ="substring-before(ws:f,'(')='DBCS'">
                              <xsl:variable name ="v" select ="substring-after(ws:f,'(')"/>
                              <xsl:value-of select ="concat('=WIDECHAR(',$v)"/>
                            </xsl:when>
							
							<!--zl 2015-4-15-->
                            <xsl:when test ="contains(ws:f,'totalExpenseProjected')">
                              <xsl:value-of select="concat('=',ws:v)"/>
                            </xsl:when>
                            <xsl:when test ="contains(ws:f,'totalExpenseActual')">
                              <xsl:value-of select="concat('=',ws:v)"/>
                            </xsl:when>
                            <!--zl 2015-4-15-->

                            <!--zl 2015-4-15-->
                            <xsl:when test ="contains(ws:f,'开始体重-最终体重')">
                              <xsl:value-of select="concat('=',ws:v)"/>
                            </xsl:when>
                            <xsl:when test ="contains(ws:f,'结束日期-开始日期')">
                              <xsl:value-of select="concat('=',ws:v)"/>
                            </xsl:when>
                            <xsl:when test ="contains(ws:f,'目标体重/A30')">
                              <xsl:value-of select="concat('=',substring(ws:v,0,5))"/>
                            </xsl:when>
                            <!--zl 2015-4-15-->
							
                            <xsl:otherwise >
                              <xsl:value-of select="concat('=',ws:f)"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </表:公式_E7B5>
                      </xsl:if>

                      <!--2014-3-11, add by Qihy, 修复BUGBUG2940-区域公式集多处错误中的now和today公式转换出错 start-->
                      <xsl:if test="ws:f[@ca = 1]">
                        <表:公式_E7B5>
                          <xsl:value-of select ="concat('=', ws:f)"/>
                        </表:公式_E7B5>
                      </xsl:if>
                      <!--2014-3-11 end -->

					  <!--zl 2015-4-20-->
                      <xsl:if test="ws:f[@ca = '1']">
						  <表:公式_E7B5>
							<xsl:value-of select ="concat('=', ws:f)"/>
						  </表:公式_E7B5>
                      </xsl:if>
                      <xsl:if test="ws:f[@t = 'array']">
						  <表:公式_E7B5>
							<xsl:value-of select ="concat('=', ws:f)"/>
						  </表:公式_E7B5>
                      </xsl:if>
                      <!--zl 2015-4-20-->
					  
                    </xsl:if>
                    <xsl:if test="ws:f and ws:f='' and ws:f/@t and not(ws:f/@ref)">

                      <!--2014-3-23 update by Qihy, 修复BUG3149 (Transition)OOX->UOF->OOX 第二轮转换失败， start-->
                      <!--<xsl:if test="parent::ws:row/ws:c/ws:f[@ref]">
                            -->
                      <!--
												<xsl:variable name="ref" select="translate(parent::ws:row/ws:c/ws:f/@ref,1234567890,'')"/>
												<xsl:variable name="gongshi" select="parent::ws:row/ws:c/ws:f[contains(@ref,$iid)]"/>
												<xsl:variable name="gongshif" select="substring-before($gongshi,'(')"/>
												<xsl:variable name="gongshit" select="substring-after($gongshi,'(')"/>
												<xsl:variable name="content" select="translate($gongshit,$iid,'ABCDEFGHIJKLMNOPQRSTUVWXYZ')"/>
												<xsl:variable name="gongs" select="concat($gongshif,'(',$content)"/>
												<表:公式 uof:locID="s0052">
													<xsl:value-of select="concat('=',$gongs)"/>$curRefTemp
												</表:公式>
                        -->
                      <!--
                            <xsl:variable name="siValue">
                              <xsl:value-of select="./ws:f/@si"/>
                            </xsl:variable>
                            <xsl:variable name="curRef">
                              <xsl:value-of select="./@r"/>
                            </xsl:variable>
                            <xsl:variable name="curRefTemp">
                              <xsl:value-of select="translate($curRef,'1234567890','')"/>
                            </xsl:variable>
                            <xsl:variable name="func">
                              <xsl:value-of select="parent::ws:row/ws:c/ws:f[@ref and @si=$siValue]"/>
                            </xsl:variable>
                            <xsl:variable name="funcName">
                              <xsl:value-of select="substring-before($func,'(')"/>
                            </xsl:variable>
                            <xsl:variable name="funcContent">
                              <xsl:value-of select="substring-after($func,$funcName)"/>
                            </xsl:variable>
                            <xsl:variable name="funcTemp">
                              <xsl:value-of select="translate($funcContent,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',concat($curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp,$curRefTemp))"/>
                            </xsl:variable>
                            <表:公式_E7B5>
                              <xsl:value-of select="concat('=',$funcName,$funcTemp)"/>
                            </表:公式_E7B5>
                          </xsl:if>-->

                      <xsl:variable name ="ref" select ="parent::ws:c/@r"/>
                      <xsl:variable name ="si" select ="parent::ws:c/@si"/>
                      <xsl:for-each select ="parent::ws:row/ws:c/ws:f[@ref]">
                        <xsl:if test ="@ref = $ref and @si = $si">
                          <表:公式_E7B5>
                            <xsl:value-of select="concat('=',ws:f)"/>
                          </表:公式_E7B5>
                        </xsl:if>
                      </xsl:for-each>
                      <!--2014-3-23 end-->

                    </xsl:if>
                  </表:数据_E7B3>
                </xsl:if>
                <!--有批注的情况下-->
                <xsl:if test="ws:comment">
                  <表:批注_E7B7>
                    <xsl:choose>
                      <xsl:when test="ws:comment/@visibility">
                        <xsl:attribute name="是否显示_E7B9">
                          <xsl:value-of select="'false'"/>
                        </xsl:attribute>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:attribute name="是否显示_E7B9">
                          <xsl:value-of select="'true'"/>
                        </xsl:attribute>
                      </xsl:otherwise>
                    </xsl:choose>
                    <uof:锚点_C644>
                      <xsl:attribute name="图形引用_C62E">
                        <xsl:value-of select="ws:comment/@phNoRef"/>
                      </xsl:attribute>
                      <xsl:attribute name="随动方式_C62F">none</xsl:attribute>
                      <!--添加 位置元素，李杨2011-11-30-->
                      <uof:位置_C620>
                        <xsl:variable name ="row">
                          <xsl:value-of select ="translate(ws:comment/@ref,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
                        </xsl:variable>
                        <xsl:variable name ="col">
                          <xsl:value-of select ="translate(ws:comment/@ref,'01234546789','')"/>
                        </xsl:variable>
                        <xsl:variable name ="colvalue">
                          <xsl:call-template name="Ts2Dec">
                            <xsl:with-param name="tsSource" select="$col"/>
                          </xsl:call-template>
                          <!--2014-4-27 end-->
                        </xsl:variable>
                        <uof:水平_4106>
                          <uof:绝对_4107>
                            <xsl:attribute name ="值_4108">
                              <xsl:value-of select ="$colvalue * 54"/>
                            </xsl:attribute>
                          </uof:绝对_4107>
                        </uof:水平_4106>
                        <uof:垂直_410D>
                          <uof:绝对_4107>
                            <xsl:attribute name ="值_4108">
                              <xsl:value-of select ="$row * 13.5"/>
                            </xsl:attribute>
                          </uof:绝对_4107>
                        </uof:垂直_410D>
                      </uof:位置_C620>

                      <uof:大小_C621 长_C604="59.999084" 宽_C605="97.49851"/>
                    </uof:锚点_C644>
                  </表:批注_E7B7>
                </xsl:if>
                
                <!--有批注的情况下-->
              </表:单元格_E7F2>
            </xsl:for-each>
          </xsl:if>

          <xsl:if test="not(ws:c)">
            <xsl:for-each select="ancestor::ws:worksheet/ws:mergeCells/ws:mergeCell">
              <xsl:variable name="mergeref" select="@ref"/>
              <xsl:variable name="mss" select="substring-before($mergeref,':')"/>
              <xsl:variable name="rrrid" select="translate($mss,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
              <xsl:if test="$rrrid=$rownumber">
                <表:单元格_E7F2>
                  <xsl:variable name="iid" select="translate($mss,'1234567890','')"/>
                  <xsl:attribute name="列号_E7BC">
                    <!-- add by xuzhenwei BUG_2462 单元格列合并效果丢失 2013-01-16 -->
                    <xsl:call-template name="Ts2Dec">
                      <xsl:with-param name="tsSource" select="$iid"/>
                    </xsl:call-template>
                    <!-- end -->
                  </xsl:attribute>
                  <表:合并_E7AF>
                    <xsl:call-template name="hbdyg">
                      <xsl:with-param name="cid" select="$rownumber"/>
                      <xsl:with-param name="mergeref" select="$mergeref"/>
                    </xsl:call-template>
                  </表:合并_E7AF>
                </表:单元格_E7F2>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
        </表:行_E7F1>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
  
  
  <!-- add xuzhenwei 2012-11-28 新增功能点：用户方案集 start -->
  <xsl:template match="ws:scenarios">
    <表:方案集_E7F8>
      <xsl:if test="ws:scenario">
        <xsl:for-each select="ws:scenario">
          <表:方案_E7F9>
            <!-- 方案集 -->
            <xsl:if test="@name">
              <xsl:attribute name="名称_E774">
                <xsl:value-of select="@name"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="@locked">
              <xsl:attribute name="是否防止更改_E800">
                <xsl:if test="@locked=1">
                  <xsl:value-of select="'true'"/>
                </xsl:if>
                <xsl:if test="@locked=0">
                  <xsl:value-of select="'false'"/>
                </xsl:if>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="@hidden">
              <xsl:attribute name="是否隐藏_E801">
                <xsl:if test="@hidden=1">
                  <xsl:value-of select="'true'"/>
                </xsl:if>
                <xsl:if test="@hidden=0">
                  <xsl:value-of select="'false'"/>
                </xsl:if>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="@comment">
              <xsl:attribute name="备注_E7FE">
                <xsl:value-of select="@comment"/>
              </xsl:attribute>
            </xsl:if>
            <!-- UOF的单元格内容 -->
            <xsl:if test="ws:inputCells">
              <xsl:for-each select="ws:inputCells">
                <xsl:variable name="cell_id" select="@r"/>
                <表:可变单元格_E7FA>
                  <!--设置单元格属性-->
                  <xsl:attribute name="值_E7FD">
                    <xsl:value-of select="@val"/>
                  </xsl:attribute>
                  <xsl:attribute name="行号_E7FB">
                    <xsl:variable name="id" select="@r"/>
                    <xsl:variable name="iid" select="translate($id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
                    <xsl:variable name="resultRow" select="substring-after($iid,'')"/>
                    <xsl:value-of select="$resultRow -1"/>
                  </xsl:attribute>
                  <xsl:attribute name="列号_E7FC">
                    <xsl:variable name="id" select="@r"/>
                    <xsl:variable name="iid" select="translate($id,'1234567890','')"/>
                    <xsl:variable name="z" select="'000000'"/>
                    <xsl:variable name="RowID" select="concat(substring($z,1,(string-length($z)-string-length($iid))),$iid)"/>
                    <xsl:variable name="Char26" select="'0ABCDEFGHIJKLMNOPQRSTUVWXYZ'"/>
                    <xsl:variable name="dec1" select="string-length(substring-before($Char26,substring($RowID,1,1)))"/>
                    <xsl:variable name="dec2" select="$dec1*26+string-length(substring-before($Char26,substring($RowID,2,1)))"/>
                    <xsl:variable name="dec3" select="$dec2*26+string-length(substring-before($Char26,substring($RowID,3,1)))"/>
                    <xsl:variable name="dec4" select="$dec3*26+string-length(substring-before($Char26,substring($RowID,4,1)))"/>
                    <xsl:variable name="dec5" select="$dec4*26+string-length(substring-before($Char26,substring($RowID,5,1)))"/>
                    <xsl:variable name="Result" select="$dec5*26+string-length(substring-before($Char26,substring($RowID,6,1)))"/>
                    <xsl:value-of select="$Result -1"/>
                  </xsl:attribute>
                  <!--有批注的情况下-->
                  <xsl:if test="following::pr:Relationships[1]/*[contains(@Target,'comments')]">
                    <xsl:for-each select="following::pr:Relationships[1]/*[contains(@Target,'comments')]">
                      <xsl:variable name="tar" select="@Target"/>
                      <xsl:variable name="tarId" select="concat('xlsx/xl/',substring-after($tar,'/'))"/>
                      <xsl:for-each select="./ancestor-or-self::ws:spreadsheet/ws:comments/ws:commentList/ws:comment">
                        <xsl:variable name="cmmpos">
                          <xsl:number count="ws:comment" level="single"/>
                        </xsl:variable>
                        <xsl:if test="@ref=$cell_id">
                          <表:批注_E7B7>
                            <xsl:attribute name="是否显示_E7B9">
                              <xsl:value-of select="'true'"/>
                            </xsl:attribute>
                            <uof:锚点_C644>
                              <!--uof:attrList="x坐标 y坐标 宽度 高度 图形引用 随动方式 缩略图 占位符" uof:宽度="97.49851" uof:高度="59.999084">-->
                              <xsl:variable name="xml">
                                <xsl:value-of select="substring-after($tar,'/')"/>
                              </xsl:variable>
                              <xsl:variable name="xxml">
                                <xsl:value-of select="substring-before($xml,'.')"/>
                              </xsl:variable>
                              <xsl:variable name ="phNo">
                                <xsl:value-of select ="ancestor::ws:spreadsheet/@sheetId"/>
                              </xsl:variable>
                              <xsl:attribute name="图形引用_C62E">
                                <xsl:value-of select="concat('comments',$phNo,'.',$cmmpos)"/>
                              </xsl:attribute>
                              <xsl:attribute name="随动方式_C62F">none</xsl:attribute>
                              <uof:位置_C620>
                                <xsl:variable name ="row">
                                  <xsl:value-of select ="translate(@ref,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
                                </xsl:variable>
                                <xsl:variable name ="col">
                                  <xsl:value-of select ="translate(@ref,'01234546789','')"/>
                                </xsl:variable>
                                <xsl:variable name ="colvalue">

                                  <!--2014-4-27, update by Qihy, 修复带有批注时，因取值为NaN导致第二轮转换失败， start-->
                                  <!--<xsl:variable name ="colv">
                                    <xsl:choose >
                                      <xsl:when test ="contains($col,'J')">
                                        <xsl:value-of select ="translate($col,'J','9')+1"/>
                                      </xsl:when>
                                      <xsl:when test ="contains($col,'K')">
                                        <xsl:value-of select ="translate($col,'K','9')+2"/>
                                      </xsl:when>
                                      <xsl:when test ="contains($col,'L')">
                                        <xsl:value-of select ="translate($col,'L','97')+2"/>
                                      </xsl:when>
                                      -->
                                  <!--取ABCDEF等等的列宽度-->
                                  <!--
                                      <xsl:otherwise>
                                        <xsl:value-of select ="translate($col,'ABCDEFGHI','123456789')"/>
                                      </xsl:otherwise>
                                    </xsl:choose>
                                  </xsl:variable>

                                  <xsl:value-of select ="$colv"/>-->
                                  <xsl:call-template name="Ts2Dec">
                                    <xsl:with-param name="tsSource" select="$col"/>
                                  </xsl:call-template>
                                  <!--2014-4-27 end-->

                                </xsl:variable>
                                <uof:水平_4106>
                                  <uof:绝对_4107>
                                    <xsl:attribute name ="值_4108">
                                      <xsl:value-of select ="$colvalue * 54"/>
                                    </xsl:attribute>
                                  </uof:绝对_4107>
                                </uof:水平_4106>
                                <uof:垂直_410D>
                                  <uof:绝对_4107>
                                    <xsl:attribute name ="值_4108">
                                      <xsl:value-of select ="$row * 13.5"/>
                                    </xsl:attribute>
                                  </uof:绝对_4107>
                                </uof:垂直_410D>
                              </uof:位置_C620>

                              <uof:大小_C621 长_C604="59.999084" 宽_C605="97.49851"/>
                            </uof:锚点_C644>
                          </表:批注_E7B7>
                        </xsl:if>
                      </xsl:for-each>
                    </xsl:for-each>
                  </xsl:if>
                  <!--有批注的情况下-->
                </表:可变单元格_E7FA>
              </xsl:for-each>
            </xsl:if>

            <xsl:if test="not(ws:c)">
              <xsl:variable name="hang" select="@r"/>
              <xsl:if test="following::pr:Relationships[1]/*[contains(@Target,'comments')]">
                <xsl:for-each select="following::pr:Relationships[1]/*[contains(@Target,'comments')]">
                  <xsl:variable name="tar" select="@Target"/>
                  <!--tar:../comments1.xml-->
                  <xsl:variable name="tarId" select="concat('xlsx/xl/',substring-after($tar,'/'))"/>
                  <xsl:for-each select="./ancestor-or-self::ws:spreadsheet/ws:comments/ws:commentList/ws:comment">
                    <xsl:variable name="cmmpos">
                      <xsl:number count="ws:comment" level="single"/>
                    </xsl:variable>
                    <xsl:variable name="reff" select="translate(@ref,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
                    <xsl:if test="$reff = $hang">
                      <表:单元格_E7F2>
                        <!--uof:attrList="列号 式样引用 超链接引用 合并列数 合并行数">-->
                        <xsl:attribute name="列号_E7BC">
                          <xsl:variable name="reff1" select="translate(@ref,'0123456789','')"/>
                          <xsl:call-template name="Ts2Dec">
                            <xsl:with-param name="tsSource" select="$reff1"/>
                          </xsl:call-template>
                        </xsl:attribute>
                        <表:批注_E7B7 是否显示_E7B9="false">
                          <uof:锚点_C644>
                            <!--uof:attrList="x坐标 y坐标 宽度 高度 图形引用 随动方式 缩略图 占位符" uof:宽度="97.49851" uof:高度="59.999084">-->
                            <xsl:variable name="xml">
                              <xsl:value-of select="substring-after($tar,'/')"/>
                            </xsl:variable>
                            <xsl:variable name="xxml">
                              <xsl:value-of select="substring-before($xml,'.')"/>
                            </xsl:variable>
                            <xsl:variable name ="phNo">
                              <xsl:value-of select ="ancestor::ws:spreadsheet/@sheetId"/>
                            </xsl:variable>
                            <xsl:attribute name="图形引用_C62E">
                              <xsl:value-of select="concat('comments',$phNo,'.',$cmmpos)"/>
                            </xsl:attribute>
                            <xsl:attribute name="随动方式_C62F">none</xsl:attribute>

                            <uof:大小_C621 长_C604="59.999084" 宽_C605="97.49851"/>
                          </uof:锚点_C644>
                        </表:批注_E7B7>
                      </表:单元格_E7F2>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:for-each>
              </xsl:if>
            </xsl:if>
          </表:方案_E7F9>
        </xsl:for-each>
      </xsl:if>
    </表:方案集_E7F8>
  </xsl:template>
  <!-- end -->
  <xsl:template name="hbdyg">
    <xsl:param name="cid"/>
    <xsl:param name="mergeref"/>

    <xsl:attribute name="列数_E7B0">
      <xsl:call-template name="hbls">
        <xsl:with-param name="lh" select="$mergeref"/>
      </xsl:call-template>
    </xsl:attribute>
    <xsl:attribute name="行数_E7B1">
      <xsl:call-template name="hbhs">
        <xsl:with-param name="hh" select="$mergeref"/>
      </xsl:call-template>
    </xsl:attribute>

  </xsl:template>
  <xsl:template name="hbls">
    <xsl:param name="lh"/>
    <xsl:variable name="b" select="substring-before($lh,':')"/>
    <xsl:variable name="a" select="substring-after($lh,':')"/>
    <xsl:variable name="bb" select="translate($b,'0123456789','')"/>
    <xsl:variable name="aa" select="translate($a,'0123456789','')"/>
    <xsl:variable name="bbb">
      <xsl:call-template name="Ts2Dec">
        <xsl:with-param name="tsSource" select="$bb"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:variable name="aaa">
      <xsl:call-template name="Ts2Dec">
        <xsl:with-param name="tsSource" select="$aa"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:value-of select="number($aaa)-number($bbb)"/>
  </xsl:template>
  <xsl:template name="hbhs">
    <xsl:param name="hh"/>
    <xsl:variable name="b" select="substring-before($hh,':')"/>
    <xsl:variable name="a" select="substring-after($hh,':')"/>
    <xsl:variable name="bb" select="translate($b,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
    <xsl:variable name="aa" select="translate($a,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
    <xsl:value-of select="number($aa)-number($bb)"/>
  </xsl:template>
  <xsl:template name="Ts2Dec">
    <xsl:param name="tsSource"/>
    <xsl:variable name="z" select="'000000'"/>
    <xsl:variable name="RowID" select="concat(substring($z,1,(string-length($z)-string-length($tsSource))),$tsSource)"/>
    <xsl:variable name="Char26" select="'0ABCDEFGHIJKLMNOPQRSTUVWXYZ'"/>
    <xsl:variable name="dec1" select="string-length(substring-before($Char26,substring($RowID,1,1)))">
    </xsl:variable>
    <xsl:variable name="dec2" select="$dec1*26+string-length(substring-before($Char26,substring($RowID,2,1)))">
    </xsl:variable>
    <xsl:variable name="dec3" select="$dec2*26+string-length(substring-before($Char26,substring($RowID,3,1)))">
    </xsl:variable>
    <xsl:variable name="dec4" select="$dec3*26+string-length(substring-before($Char26,substring($RowID,4,1)))">
    </xsl:variable>
    <xsl:variable name="dec5" select="$dec4*26+string-length(substring-before($Char26,substring($RowID,5,1)))">
    </xsl:variable>
    <xsl:variable name="Result" select="$dec5*26+string-length(substring-before($Char26,substring($RowID,6,1)))">
    </xsl:variable>
    <xsl:value-of select="$Result"/>
  </xsl:template>
  <xsl:template match="ws:si">
    <xsl:value-of select="position()"/>
  </xsl:template>
  <xsl:template name="Filter">

  </xsl:template>
  <!--
	<xsl:template name="DataGrouping">
		<表:分组集 uof:locID="s0098">
      <xsl:if test=".//cols[@outlineLevel]">
        
      </xsl:if>
		</表:分组集>
	</xsl:template>
  -->
</xsl:stylesheet>
