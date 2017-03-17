<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:图表="http://schemas.uof.org/cn/2009/chart"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:pc="http://schemas.openxmlformats.org/package/2006/content-types"
                xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:ori="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
  <xsl:import href="style.xsl"/>

  <xsl:template name="object_chart">
    <xsl:for-each select="ws:spreadsheets/ws:spreadsheet">
      <xsl:if test="ws:Drawings/xdr:wsDr/pr:Relationships/pr:Relationship">
        <xsl:for-each select="ws:Drawings/xdr:wsDr/pr:Relationships/pr:Relationship">
          <xsl:if test="contains(@Target,'charts')">
            <xsl:call-template name="chart-picture">
              <xsl:with-param name="rid" select="@Id"/>
              <!--Id="rId1"-->
              <xsl:with-param name="target" select="@Target"/>
              <!--Target="../charts/chart1.xml"-->
            </xsl:call-template>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="chart-picture">
    <xsl:param name="rid"/>
    <xsl:param name="target"/>
    <xsl:variable name="pic_id">
      <!--chart1.xml-->
      <xsl:value-of select="substring-after($target,'charts/')"/>
    </xsl:variable>
    <xsl:variable name="pic_id2">
      <!--xlsx/xl/charts/_rels/chart1.xml.rels文件-->
      <xsl:value-of select="concat('xlsx/xl/charts/_rels/',$pic_id,'.rels')"/>
    </xsl:variable>
    <xsl:variable name="path">
      <xsl:value-of select="concat('xlsx/xl/charts/',$pic_id)"/>
    </xsl:variable>
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace/ws:picture">
      <!--存在指定的元素,说明对应的chart1.xml.rels文件存在4.16-->
      <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=$pic_id]/pr:Relationships/pr:Relationship">
        <!--Target="../media/image1.gif-->
        <xsl:variable name="tar">
          <xsl:if test="@Target">
            <xsl:value-of select="@Target"/>
          </xsl:if>
        </xsl:variable>
        <xsl:variable name="id">
          <xsl:value-of select="@Id"/>
        </xsl:variable>
        <xsl:if test="contains($tar,'media')">
          <xsl:for-each select="@Target">

            <!--用C# 转media中的图片-->
            <uof:其他对象 uof:locID="u0036" uof:attrList="标识符 内嵌 公共类型 私有类型">
              <xsl:attribute name="uof:标识符">
                <!--标识符:../media/image1.jpeg-->
                <xsl:variable name="regid1">
                  <xsl:value-of select="substring-after(.,'media/')"/>
                </xsl:variable>
                <xsl:value-of select="substring-before($regid1,'.')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:内嵌">false</xsl:attribute>
              <xsl:choose>
                <xsl:when test="contains(.,'emf')">
                  <xsl:attribute name="uof:私有类型">
                    <xsl:variable name="type">
                      <xsl:value-of select="substring-after(.,'media/')"/>
                    </xsl:variable>
                    <xsl:value-of select="substring-after($type,'.')"/>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="uof:公共类型">
                    <xsl:variable name="type">
                      <xsl:value-of select="substring-after(.,'media/')"/>
                    </xsl:variable>
                    <xsl:value-of select="substring-after($type,'.')"/>
                  </xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
              <uof:数据 uof:locID="u0037">
                <!--用C# 转media中的图片，转成二进制代码-->
              </uof:数据>
            </uof:其他对象>
          </xsl:for-each>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
  <xsl:template name="table">
    <xsl:call-template name="chartset"/>
  </xsl:template>
  <xsl:template name="chartset">
    <xsl:if test="ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/pr:Relationships/pr:Relationship">
      <xsl:for-each select="ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/pr:Relationships/pr:Relationship">
        <xsl:if test="contains(@Target,'charts')">
          <xsl:call-template name="chart-tubiao">
            <xsl:with-param name="rid" select="@Id"/>
            <xsl:with-param name="target" select="@Target"/>
          </xsl:call-template>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
  <!--图表-->
  <xsl:template name="chart-tubiao">
    <xsl:param name="rid"/>
    <xsl:param name="target"/>
    <图表:图表_E837>
      <xsl:attribute name="标识符_E828">
        <xsl:value-of select="concat(ancestor::ws:spreadsheet/@sheetName,'_chart_',$rid)"/>
      </xsl:attribute>
      <xsl:variable name="aim">
        <xsl:value-of select="@Target"/>
      </xsl:variable>
      <xsl:variable name="aim2">
        <xsl:value-of select="substring-after(@Target,'../')"/>
      </xsl:variable>
      <xsl:variable name="aim3">
        <xsl:value-of select="concat('xlsx/xl/',$aim2)"/>
      </xsl:variable>
      <xsl:variable name="t">
        <xsl:value-of select="substring-after($target,'../')"/>
      </xsl:variable>
      <xsl:variable name="tt">
        <xsl:value-of select="concat('xlsx/xl/',$t)"/>
      </xsl:variable>
      <!--图表区-->
      <xsl:call-template name="chart-tubiaoqu">
        <xsl:with-param name="rid" select="@Id"/>
        <xsl:with-param name="target" select="@Target"/>
      </xsl:call-template>
      <!--绘图区-->
      <xsl:call-template name="chart-huituqu">
        <xsl:with-param name="tar" select="$tt"/>
        <xsl:with-param name="rid" select="@Id"/>
        <xsl:with-param name="target" select="@Target"/>
      </xsl:call-template>
      <!--图例-->
      <xsl:call-template name="chart-tuli">
        <xsl:with-param name="tar" select="$tt"/>
        <xsl:with-param name="rid" select="@Id"/>
        <xsl:with-param name="target" select="@Target"/>
      </xsl:call-template>
      <!--数据表-->
      <xsl:call-template name="chart-sjb">
        <xsl:with-param name="tar" select="$tt"/>
        <xsl:with-param name="rid" select="@Id"/>
        <xsl:with-param name="target" select="@Target"/>
      </xsl:call-template>
      <!--标题-->
      <xsl:call-template name="chart-btj">
        <xsl:with-param name="tar" select="$tt"/>
        <xsl:with-param name="rid" select="@Id"/>
        <xsl:with-param name="target" select="@Target"/>
      </xsl:call-template>
      <!--背景墙-->
      <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:backWall/c:spPr]">
        <图表:背景墙_E7A1>
          <xsl:apply-templates select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:backWall/c:spPr"/>
        </图表:背景墙_E7A1>
      </xsl:if>
      <!--基底-->
      <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:floor/c:spPr]">
        <图表:基底_E7A4>
          <xsl:apply-templates select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:floor/c:spPr"/>
        </图表:基底_E7A4>
      </xsl:if>
      <!--空白单元格绘制方式-->
      <图表:空白单元格绘制方式_E7A5>none</图表:空白单元格绘制方式_E7A5>
      <!--是否显示隐藏单元格-->
      <图表:是否显示隐藏单元格_E7A6>false</图表:是否显示隐藏单元格_E7A6>
    </图表:图表_E837>
  </xsl:template>
  <!--图表区-->
  <xsl:template name="chart-tubiaoqu">
    <xsl:param name="target"/>
    <xsl:param name="rid"/>
    <xsl:variable name="starget">
      <xsl:value-of select="substring-after($target,'charts/')"/>
    </xsl:variable>
    <xsl:variable name="sstarget">
      <xsl:value-of select="substring-before($starget,'.')"/>
    </xsl:variable>
    <图表:图表区_E743>
      <xsl:variable name="t">
        <xsl:value-of select="substring-after($target,'../')"/>
      </xsl:variable>
      <xsl:variable name="tt">
        <xsl:value-of select="concat('xlsx/xl/',$t)"/>
      </xsl:variable>
      <!--边框线和填充-->
      <xsl:choose>
        <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/') and c:spPr]">
          <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:spPr">
            <xsl:call-template name="chart-biankuang">
              <xsl:with-param name="target2" select="$target"/>
            </xsl:call-template>
            <xsl:call-template name="chart-tianchong">
              <xsl:with-param name="target2" select="$target"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          
          <!--2014-4-23, delete by Qihy, bug3154互操作后图表边框错误， start-->
          <!--<图表:边框线_4226 线型_C60D="single" 虚实_C60E="solid" 宽度_C60F="2" 颜色_C611="auto"/>-->
          <!--2014-4-23 end-->
          
          <!--2014-5-4, update by 凌峰, 边框线默认情况， start-->
          <图表:边框线_4226 线型_C60D="single" 虚实_C60E="solid" 宽度_C60F="1" 颜色_C611="#808080"/>
          <!--end-->
          
          <图表:填充_E746>
            <图:颜色_8004>auto</图:颜色_8004>
          </图表:填充_E746>
        </xsl:otherwise>
      </xsl:choose>

      <!--字体-->
      <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/') and c:txPr/a:p/a:pPr]">
        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:txPr/a:p/a:pPr">
          <图表:字体_E70B>
            <xsl:if test="a:defRPr/a:ea or a:defRPr/a:latin or a:defRPr/@sz or a:defRPr/a:solidFill/a:schemeClr or a:defRPr/a:solidFill/a:srgbClr">
              <!-- 20130329 update by xuzhenwei BUG_2494 字体颜色，下划线效果丢失 start -->
              <字:字体_4128>
                <xsl:attribute name="西文字体引用_4129">
                  <xsl:value-of select="concat('chartSpace_la_',$sstarget)"/>
                </xsl:attribute>
                <xsl:attribute name="中文字体引用_412A">
                  <xsl:value-of select="concat('chartSpace_ea_',$sstarget)"/>
                </xsl:attribute>
                <xsl:if test="a:defRPr[@sz]">
                  <xsl:attribute name="字号_412D">
                    <xsl:variable name="sz">
                      <xsl:value-of select="a:defRPr/@sz"/>
                    </xsl:variable>
                    <xsl:value-of select="$sz div 100"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="a:defRPr/a:solidFill/a:schemeClr or a:defRPr/a:solidFill/a:srgbClr">
                  <xsl:attribute name="颜色_412F">
                    <xsl:call-template name="color_ziti"/>
                  </xsl:attribute>
                </xsl:if>
              </字:字体_4128>
              <!-- end -->
              <xsl:choose>
                <xsl:when test="not(ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:txPr/a:p/a:r)">
                  <字:是否粗体_4130>false</字:是否粗体_4130>
                </xsl:when>
                <xsl:otherwise>
                  <字:是否粗体_4130>true</字:是否粗体_4130>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
            
            <xsl:call-template name="table-tbziti"/>
          </图表:字体_E70B>
        </xsl:for-each>

      </xsl:if>
    </图表:图表区_E743>
  </xsl:template>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
  <!--图表区-边框线-->
  <xsl:template name="chart-biankuang">
    <xsl:param name="target2"/>
    <xsl:if test="not(a:ln)">
      <图表:边框线_4226 线型_C60D="single" 虚实_C60E="solid" 宽度_C60F="2" 颜色_C611="#808080"/>
    </xsl:if>
    <xsl:for-each select="a:ln">
      <图表:边框线_4226>
        <!--线型-->
        <xsl:choose>
          <xsl:when test="./a:noFill">
            <xsl:attribute name="线型_C60D">
              <xsl:value-of select="'none'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="lineType">
              <xsl:choose>
                <xsl:when test="./@cmpd">
                  <xsl:value-of select="./@cmpd"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'none'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="dashType">
              <xsl:choose>
                <xsl:when test="./a:prstDash/@val">
                  <xsl:value-of select="./a:prstDash/@val"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'none'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:call-template name="LineTypeTransfer">
              <!--线型，虚实-->
              <xsl:with-param name="lineType" select="$lineType"/>
              <xsl:with-param name="dashType" select="$dashType"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
        <!--宽度-->
        <xsl:if test="@w">
          <xsl:variable name="w">
            <xsl:value-of select="@w"/>
          </xsl:variable>
          <xsl:attribute name="宽度_C60F">
            <!-- 20130416 update by xuzhenwei BUG_2830:互操作 oo-uof（编辑）-oo 024实用文档-损益表(1).xlsx 文档需要修复 start -->
            <xsl:choose>
              <xsl:when test="number(@w div 12700) &gt; 1">
                <xsl:value-of select="number(@w div 12700) - 1"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="1 - number(@w div 12700)"/>
              </xsl:otherwise>
            </xsl:choose>
            <!-- end -->
          </xsl:attribute>
        </xsl:if>
        <!--颜色-->
        <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
          <xsl:attribute name="颜色_C611">
            <xsl:call-template name="color"/>
          </xsl:attribute>
        </xsl:if>
      </图表:边框线_4226>
    </xsl:for-each>
  </xsl:template>
  <!--填充-->
  <xsl:template name="chart-tianchong">
    <xsl:param name="target2"/>
    <xsl:if test="a:solidFill or a:blipFill or a:gradFill or a:pattFill">
      <图表:填充_E746>
        <xsl:attribute name="是否填充随图形旋转_8067">false</xsl:attribute>
        <!--纯色填充-->
        <xsl:if test="a:solidFill">
          <图:颜色_8004>
            <xsl:call-template name="color"/>
          </图:颜色_8004>
        </xsl:if>
        <!--图片填充-->
        <xsl:if test="a:blipFill">
          <图:图片_8005 位置_8006="stretch">
            <xsl:attribute name="图形引用_8007">
              <xsl:variable name="rid">
                <xsl:value-of select="a:blipFill/a:blip/@r:embed"/>
              </xsl:variable>
              <xsl:variable name="t">
                <!--chart1.xml-->
                <xsl:value-of select="substring-after($target2,'charts/')"/>
              </xsl:variable>
              <xsl:variable name="tt">
                <xsl:value-of select="concat('xlsx/xl/charts/_rels/',$t,'.rels')"/>
                <!--xlsx/xl/charts/_rels/chart1.xml.rels-->
              </xsl:variable>
              <xsl:variable name="pictureobj">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target2,'charts/')]/pr:Relationships/pr:Relationship[@Id=$rid]/@Target"/>
                <!--图:图形引用="../media/image1.jpeg"-->
              </xsl:variable>
              <xsl:variable name="pictureobj1">
                <xsl:value-of select="substring-after($pictureobj,'media/')"/>
                <!--图:图形引用="../media/image1.jpeg"-->
              </xsl:variable>
              <xsl:value-of select="substring-before($pictureobj1,'.')"/>
            </xsl:attribute>
          </图:图片_8005>
        </xsl:if>
        <!--图案-->
        <xsl:if test="a:pattFill">
          <图:图案_800A>
            <xsl:if test="a:pattFill/@prst">
              <xsl:variable name="type">
                <xsl:value-of select="a:pattFill/@prst"/>
              </xsl:variable>
              <xsl:attribute name="类型_8008">
                <xsl:call-template name="ttype">
                  <xsl:with-param name="tttype" select="$type"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:pattFill/a:fgClr or a:pattFill/a:bgClr">
              <xsl:for-each select="a:pattFill/a:fgClr">
                <xsl:attribute name="前景色_800B">
                  <xsl:call-template name="pattClr"/>
                </xsl:attribute>
              </xsl:for-each>
              <xsl:for-each select="a:pattFill/a:bgClr">
                <!-- 20130329 update by xuzhenwei BUG_2495 图案填充：背景色蓝色转为白色 start -->
                <xsl:attribute name="背景色_800C">
                  <xsl:call-template name="pattClr"/>
                </xsl:attribute>
                <!-- end -->
              </xsl:for-each>
            </xsl:if>
          </图:图案_800A>
        </xsl:if>
        <!--渐变-->
        <xsl:if test="a:gradFill">
          <xsl:apply-templates select="a:gradFill"/>
        </xsl:if>
      </图表:填充_E746>
    </xsl:if>
  </xsl:template>
  <!--渐变-->
  <!--<xsl:template match="a:gradFill">
    <图:渐变_800D>
      <xsl:variable name="angle">
        <xsl:choose>
          <xsl:when test="a:lin">
            <xsl:value-of select="(360 - round(a:lin/@ang div 60000 div 45)*45+90) mod 360"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:attribute name="种子类型_8010">
        <xsl:choose>
          <xsl:when test="a:path">
            <xsl:choose>
              <xsl:when test="a:path/@path='rect'">square</xsl:when>
              <xsl:when test="a:path/@path='circle'">oval</xsl:when>
              <xsl:when test="a:path/@path='shape'">square</xsl:when>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>linear</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <xsl:attribute name="起始浓度_8011">1.0</xsl:attribute>
      <xsl:attribute name="终止浓度_8012">1.0</xsl:attribute>
      <xsl:attribute name="渐变方向_8013">
        <xsl:value-of select="$angle"/>
      </xsl:attribute>
      -->
  <!--起始色，终止色--><!--
      <xsl:choose>
        <xsl:when test="$angle='135' or $angle='180' or $angle='225' or $angle='270'">
          <xsl:for-each select="a:gsLst/a:gs">
            <xsl:sort select="@pos" data-type="number"/>
            <xsl:if test="position()=1">
              <xsl:attribute name="终止色_800F">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="position()=last()">
              <xsl:attribute name="起始色_800E">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="a:gsLst/a:gs">
            <xsl:sort select="@pos" data-type="number"/>
            <xsl:if test="position()=1">
              <xsl:attribute name="起始色_800E">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="position()=last()">
              <xsl:attribute name="终止色_800F">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:attribute name="边界_8014">50</xsl:attribute>
      --><!--种子X位置，种子Y位置-->
  <!--
      <xsl:choose>
        <xsl:when test="not(a:path)">
          <xsl:attribute name="种子X位置_8015">100</xsl:attribute>
          <xsl:attribute name="种子Y位置_8016">100</xsl:attribute>
        </xsl:when>
        <xsl:when test="a:path/@path='rect'">
          <xsl:choose>
            <xsl:when test="not(a:path/a:fillToRect/@l) and not(a:path/a:fillToRect/@t)">
              <xsl:attribute name="种子X位置_8015">30</xsl:attribute>
              <xsl:attribute name="种子Y位置_8016">30</xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:path/a:fillToRect/@l) and not(a:path/a:fillToRect/@b)">
              <xsl:attribute name="种子X位置_8015">30</xsl:attribute>
              <xsl:attribute name="种子Y位置_8016">60</xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:path/a:fillToRect/@r) and not(a:path/a:fillToRect/@t)">
              <xsl:attribute name="种子X位置_8015">60</xsl:attribute>
              <xsl:attribute name="种子Y位置_8016">30</xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:path/a:fillToRect/@r) and not(a:path/a:fillToRect/@b)">
              <xsl:attribute name="种子X位置_8015">60</xsl:attribute>
              <xsl:attribute name="种子Y位置_8016">60</xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
      </xsl:choose>
    </图:渐变_800D>
  </xsl:template>-->
  <!--图表字体-->
  <xsl:template name="table-tbziti">
    <!--是否粗体-->
    <xsl:choose>
      <xsl:when test="a:defRPr[@b] and a:defRPr[@b='1']">
        <字:是否粗体_4130>true</字:是否粗体_4130>
      </xsl:when>
      <xsl:when test="not(./ancestor::c:chartSpace/c:chart/c:title/c:txPr/a:p/a:r)">
        <字:是否粗体_4130>true</字:是否粗体_4130>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@b='1']">
          <字:是否粗体_4130>true</字:是否粗体_4130>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <!--是否斜体-->
    <xsl:choose>
      <xsl:when test="a:defRPr[@i] and a:defRPr[@i='1']">
        <字:是否斜体_4131>true</字:是否斜体_4131>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@i='1']">
          <字:是否斜体_4131>true</字:是否斜体_4131>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <!--删除线-->
    <xsl:choose>
      <xsl:when test="a:defRPr[@strike]">
        <字:删除线_4135>
          <xsl:choose>
            <xsl:when test="a:defRPr[@strike='dblStrike']">
              <xsl:value-of select="'double'"/>
            </xsl:when>
            <xsl:when test="a:defRPr[@strike='noStrike']">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test="a:defRPr[@strike='sngStrike']">
              <xsl:value-of select="'single'"/>
            </xsl:when>
          </xsl:choose>
        </字:删除线_4135>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@strike]">
          <字:删除线_4135>
            <xsl:choose>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@strike='dblStrike']">
                <xsl:value-of select="'double'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@strike='noStrike']">
                <xsl:value-of select="'none'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@strike='sngStrike']">
                <xsl:value-of select="'single'"/>
              </xsl:when>
            </xsl:choose>
          </字:删除线_4135>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <!--下划线-->
    <xsl:choose>
      <xsl:when test="a:defRPr[@u]">
        <xsl:call-template name="UnderlineTypeTransfer">
          <xsl:with-param name="lineType" select="a:defRPr/@u"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u]">
          <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u]">
            <xsl:call-template name="UnderlineTypeTransfer">
              <xsl:with-param name="lineType" select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@u"/>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <!--上下标-->
    <xsl:choose>
      <xsl:when test="a:defRPr[@baseline]">
        <字:上下标类型_4143>
          <xsl:choose>
            <xsl:when test="a:defRPr[@baseline = 0]">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test="a:defRPr[@baseline &gt;0]">
              <xsl:value-of select="'sup'"/>
            </xsl:when>
            <xsl:when test="a:defRPr[@baseline &lt;0]">
              <xsl:value-of select="'sub'"/>
            </xsl:when>
          </xsl:choose>
        </字:上下标类型_4143>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@baseline]">
          <字:上下标类型_4143>
            <xsl:choose>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@baseline = 0]">
                <xsl:value-of select="'none'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@baseline &gt;0]">
                <xsl:value-of select="'sup'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@baseline &lt;0]">
                <xsl:value-of select="'sub'"/>
              </xsl:when>
            </xsl:choose>
          </字:上下标类型_4143>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <!--字符间距-->
    <xsl:choose>
      <xsl:when test="a:defRPr[@spc]">
        <字:字符间距_4145>
          <xsl:variable name="d">
            <xsl:value-of select="a:defRPr/@spc"/>
          </xsl:variable>
          <xsl:value-of select="$d div 100"/>
        </字:字符间距_4145>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@spc]">
          <字:字符间距_4145>
            <xsl:variable name="d">
              <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@spc"/>
            </xsl:variable>
            <xsl:value-of select="$d div 100"/>
          </字:字符间距_4145>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <!--醒目字体类型-->
    <xsl:choose>
      <xsl:when test="a:defRPr[@cap='all']">
        <字:醒目字体类型_4141>uppercase</字:醒目字体类型_4141>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@cap='all']">
          <字:醒目字体类型_4141>uppercase</字:醒目字体类型_4141>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="a:defRPr[@cap='small']">
        <字:醒目字体类型_4141>small-caps</字:醒目字体类型_4141>
      </xsl:when>
      <xsl:otherwise>
        <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@cap='small']">
          <字:醒目字体类型_4141>small-caps</字:醒目字体类型_4141>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--下划线-->
  <xsl:template name="UnderlineTypeTransfer">
    <xsl:param name="lineType"/>
    <字:下划线_4136>
      <xsl:choose>
        <xsl:when test="$lineType='dbl'">
          <xsl:attribute name="线型_4137">
            <xsl:value-of select="'double'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="$lineType='none'">
          <xsl:attribute name="线型_4137">
            <xsl:value-of select="'none'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="$lineType='words'">
          <xsl:attribute name="是否字下划线_4139">
            <xsl:value-of select="'true'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="线型_4137">
            <xsl:value-of select="'single'"/>
          </xsl:attribute>
          <xsl:attribute name="虚实_4138">
            <xsl:choose>
              <xsl:when test="$lineType='dash'">
                <xsl:value-of select="'dash'"/>
              </xsl:when>
              <xsl:when test="$lineType='dashHeavy'">
                <xsl:value-of select="'long-dash'"/>
              </xsl:when>
              <xsl:when test="$lineType='dashLong'">
                <xsl:value-of select="'long-dash'"/>
              </xsl:when>
              <xsl:when test="$lineType='dashLongHeavy'">
                <xsl:value-of select="'long-dash'"/>
              </xsl:when>
              <xsl:when test="$lineType='dotDash'">
                <xsl:value-of select="'dash-dot'"/>
              </xsl:when>
              <xsl:when test="$lineType='dotDashHeavy'">
                <xsl:value-of select="'dash-dot'"/>
              </xsl:when>
              <xsl:when test="$lineType='dotDotDash'">
                <xsl:value-of select="'dash-dot-dot'"/>
              </xsl:when>
              <xsl:when test="$lineType='dotDotDashHeavy'">
                <xsl:value-of select="'dash-dot-dot'"/>
              </xsl:when>
              <xsl:when test="$lineType='dotted'">
                <xsl:value-of select="'square-dot'"/>
              </xsl:when>
              <xsl:when test="$lineType='dottedHeavy'">
                <xsl:value-of select="'square-dot'"/>
              </xsl:when>
              <xsl:when test="$lineType='heavy'">
                <xsl:value-of select="'solid'"/>
              </xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="./a:uFill">
          <xsl:for-each select="./a:uFill">
            <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
              <xsl:attribute name="颜色_412F">
                <xsl:call-template name="color"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:uFill">
            <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:uFill">
              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                <xsl:attribute name="颜色_412F">
                  <xsl:call-template name="color"/>
                </xsl:attribute>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </字:下划线_4136>
  </xsl:template>
  <!--颜色-->
  <xsl:template name="color">
    <xsl:choose>
      <xsl:when test="a:solidFill/a:schemeClr">
        <xsl:variable name="id">
          <xsl:value-of select="a:solidFill/a:schemeClr/@val"/>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="$id='dk1'">
            <xsl:value-of select="'#000000'"/>
          </xsl:when>
          <xsl:when test="$id='lt1'">
            <xsl:value-of select="'#ffffff'"/>
          </xsl:when>
          <xsl:when test="$id='dk2'">
            <xsl:value-of select="'#1f497d'"/>
          </xsl:when>
          <xsl:when test="$id='lt2'">
            <xsl:value-of select="'#eeece1'"/>
          </xsl:when>
          <xsl:when test="$id='accent1'">
            <xsl:value-of select="'#4f81bd'"/>
          </xsl:when>
          <xsl:when test="$id='accent2'">
            <xsl:value-of select="'#c0504d'"/>
          </xsl:when>
          <xsl:when test="$id='accent3'">
            <xsl:value-of select="'#9bbb59'"/>
          </xsl:when>
          <xsl:when test="$id='accent4'">
            <xsl:value-of select="'#8064a2'"/>
          </xsl:when>
          <xsl:when test="$id='accent5'">
            <xsl:value-of select="'#4bacc6'"/>
          </xsl:when>
          <xsl:when test="$id='accent6'">
            <xsl:value-of select="'#f79646'"/>
          </xsl:when>
          <xsl:when test="$id='hlink'">
            <xsl:value-of select="'#0000ff'"/>
          </xsl:when>
          <xsl:when test="$id='folHlink'">
            <xsl:value-of select="'#800080'"/>
          </xsl:when>
          <xsl:when test="$id='tx2'">
            <xsl:value-of select="'#558ed5'"/>
          </xsl:when>
          <xsl:when test="$id='tx1'">
            <xsl:value-of select="'#000000'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'auto'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="a:solidFill/a:srgbClr">
        <xsl:value-of select="concat('#',translate(a:solidFill/a:srgbClr/@val,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'auto'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="color_ziti">
    <xsl:choose>
      <xsl:when test="a:defRPr/a:solidFill/a:schemeClr">
        <xsl:variable name="id">
          <xsl:value-of select="a:defRPr/a:solidFill/a:schemeClr/@val"/>
        </xsl:variable>
        <xsl:variable name="aid">
          <xsl:value-of select="'a:accent11'"/>
        </xsl:variable>
        <!--有5个主题色微软自己就没设置，在这里转为auto-->
        <!--/xsl:if-->
        <xsl:choose>
          <xsl:when test="$id='dk1'">
            <xsl:value-of select="'#000000'"/>
          </xsl:when>
          <xsl:when test="$id='lt1'">
            <xsl:value-of select="'#ffffff'"/>
          </xsl:when>
          <xsl:when test="$id='dk2'">
            <xsl:value-of select="'#1f497d'"/>
          </xsl:when>
          <xsl:when test="$id='lt2'">
            <xsl:value-of select="'#eeece1'"/>
          </xsl:when>
          <xsl:when test="$id='accent1'">
            <xsl:value-of select="'#4f81bd'"/>
          </xsl:when>
          <xsl:when test="$id='accent2'">
            <xsl:value-of select="'#c0504d'"/>
          </xsl:when>
          <xsl:when test="$id='accent3'">
            <xsl:value-of select="'#9bbb59'"/>
          </xsl:when>
          <xsl:when test="$id='accent4'">
            <xsl:value-of select="'#8064a2'"/>
          </xsl:when>
          <xsl:when test="$id='accent5'">
            <xsl:value-of select="'#4bacc6'"/>
          </xsl:when>
          <xsl:when test="$id='accent6'">
            <xsl:value-of select="'#f79646'"/>
          </xsl:when>
          <xsl:when test="$id='hlink'">
            <xsl:value-of select="'#0000ff'"/>
          </xsl:when>
          <xsl:when test="$id='folHlink'">
            <xsl:value-of select="'#800080'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'auto'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="a:defRPr/a:solidfill/a:schemeclr">
        <xsl:variable name="id">
          <xsl:value-of select="a:defRPr/a:solidfill/a:schemeclr/@val"/>
        </xsl:variable>
        <xsl:variable name="aid">
          <xsl:value-of select="'a:accent11'"/>
        </xsl:variable>
        <!--有5个主题色微软自己就没设置，在这里转为auto-->
        <!--/xsl:if-->
        <xsl:choose>
          <xsl:when test="$id='dk1'">
            <xsl:value-of select="'#000000'"/>
          </xsl:when>
          <xsl:when test="$id='lt1'">
            <xsl:value-of select="'#ffffff'"/>
          </xsl:when>
          <xsl:when test="$id='dk2'">
            <xsl:value-of select="'#1f497d'"/>
          </xsl:when>
          <xsl:when test="$id='lt2'">
            <xsl:value-of select="'#eeece1'"/>
          </xsl:when>
          <xsl:when test="$id='accent1'">
            <xsl:value-of select="'#4f81bd'"/>
          </xsl:when>
          <xsl:when test="$id='accent2'">
            <xsl:value-of select="'#c0504d'"/>
          </xsl:when>
          <xsl:when test="$id='accent3'">
            <xsl:value-of select="'#9bbb59'"/>
          </xsl:when>
          <xsl:when test="$id='accent4'">
            <xsl:value-of select="'#8064a2'"/>
          </xsl:when>
          <xsl:when test="$id='accent5'">
            <xsl:value-of select="'#4bacc6'"/>
          </xsl:when>
          <xsl:when test="$id='accent6'">
            <xsl:value-of select="'#f79646'"/>
          </xsl:when>
          <xsl:when test="$id='hlink'">
            <xsl:value-of select="'#0000ff'"/>
          </xsl:when>
          <xsl:when test="$id='folHlink'">
            <xsl:value-of select="'#800080'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'auto'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="a:defRPr/a:solidFill/a:srgbClr">
        <xsl:value-of select="concat('#',translate(a:defRPr/a:solidFill/a:srgbClr/@val,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'auto'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="ttype">
    <xsl:param name="tttype"/>
    <xsl:choose>
      <xsl:when test="$tttype='pct5'">ptn001</xsl:when>
      <xsl:when test="$tttype='pct10'">ptn002</xsl:when>
      <xsl:when test="$tttype='pct20'">ptn003</xsl:when>
      <xsl:when test="$tttype='pct25'">ptn004</xsl:when>
      <xsl:when test="$tttype='pct30'">ptn005</xsl:when>
      <xsl:when test="$tttype='pct40'">ptn006</xsl:when>
      <xsl:when test="$tttype='pct50'">ptn007</xsl:when>
      <xsl:when test="$tttype='pct60'">ptn008</xsl:when>
      <xsl:when test="$tttype='pct70'">ptn009</xsl:when>
      <xsl:when test="$tttype='pct75'">ptn010</xsl:when>
      <xsl:when test="$tttype='pct80'">ptn011</xsl:when>
      <xsl:when test="$tttype='pct90'">ptn012</xsl:when>
      <xsl:when test="$tttype='ltDnDiag'">ptn013</xsl:when>
      <xsl:when test="$tttype='ltUpDiag'">ptn014</xsl:when>
      <xsl:when test="$tttype='dkDnDiag'">ptn015</xsl:when>
      <xsl:when test="$tttype='dkUpDiag'">ptn016</xsl:when>
      <xsl:when test="$tttype='wdDnDiag'">ptn017</xsl:when>
      <xsl:when test="$tttype='wdUpDiag'">ptn018</xsl:when>
      <xsl:when test="$tttype='ltVert'">ptn019</xsl:when>
      <xsl:when test="$tttype='ltHorz'">ptn020</xsl:when>
      <xsl:when test="$tttype='narVert'">ptn021</xsl:when>
      <xsl:when test="$tttype='narHorz'">ptn022</xsl:when>
      <xsl:when test="$tttype='dkVert'">ptn023</xsl:when>
      <xsl:when test="$tttype='dkHorz'">ptn024</xsl:when>
      <xsl:when test="$tttype='dashDnDiag'">ptn025</xsl:when>
      <xsl:when test="$tttype='dashUpDiag'">ptn026</xsl:when>
      <xsl:when test="$tttype='dashHorz'">ptn027</xsl:when>
      <xsl:when test="$tttype='dashVert'">ptn028</xsl:when>
      <xsl:when test="$tttype='smConfetti'">ptn029</xsl:when>
      <xsl:when test="$tttype='lgConfetti'">ptn030</xsl:when>
      <xsl:when test="$tttype='zigZag'">ptn031</xsl:when>
      <xsl:when test="$tttype='wave'">ptn032</xsl:when>
      <xsl:when test="$tttype='diagBrick'">ptn033</xsl:when>
      <xsl:when test="$tttype='horzBrick'">ptn034</xsl:when>
      <xsl:when test="$tttype='weave'">ptn035</xsl:when>
      <xsl:when test="$tttype='plaid'">ptn036</xsl:when>
      <xsl:when test="$tttype='divot'">ptn037</xsl:when>
      <xsl:when test="$tttype='dotGrid'">ptn038</xsl:when>
      <xsl:when test="$tttype='dotDmnd'">ptn039</xsl:when>
      <xsl:when test="$tttype='shingle'">ptn040</xsl:when>
      <xsl:when test="$tttype='trellis'">ptn041</xsl:when>
      <xsl:when test="$tttype='sphere'">ptn042</xsl:when>
      <xsl:when test="$tttype='smGrid'">ptn043</xsl:when>
      <xsl:when test="$tttype='lgGrid'">ptn044</xsl:when>
      <xsl:when test="$tttype='smCheck'">ptn045</xsl:when>
      <xsl:when test="$tttype='lgCheck'">ptn046</xsl:when>
      <xsl:when test="$tttype='openDmnd'">ptn047</xsl:when>
      <xsl:when test="$tttype='solidDmnd'">ptn048</xsl:when>
      <xsl:when test="$tttype='lightGray'">
        <xsl:value-of select="'ptn002'"/>
      </xsl:when>
      <xsl:when test="$tttype='gray125'">
        <xsl:value-of select="'ptn003'"/>
      </xsl:when>
      <xsl:when test="$tttype='gray0625'">
        <xsl:value-of select="'ptn002'"/>
      </xsl:when>
      <xsl:when test="$tttype='darkGray'">
        <xsl:value-of select="'ptn010'"/>
      </xsl:when>
      <xsl:when test="$tttype='mediumGray'">
        <xsl:value-of select="'ptn007'"/>
      </xsl:when>
      <xsl:when test="$tttype='lightGrid'">
        <xsl:value-of select="'ptn043'"/>
      </xsl:when>
      <xsl:when test="$tttype='lightTrellis'">
        <xsl:value-of select="'ptn005'"/>
      </xsl:when>
      <xsl:when test="$tttype='darkHorizontal'">
        <xsl:value-of select="'ptn024'"/>
      </xsl:when>
      <xsl:when test="$tttype='darkVertical'">
        <xsl:value-of select="'ptn023'"/>
      </xsl:when>
      <xsl:when test="$tttype='darkDown'">
        <xsl:value-of select="'ptn015'"/>
      </xsl:when>
      <xsl:when test="$tttype='darkTrellis'">
        <xsl:value-of select="'ptn006'"/>
      </xsl:when>
      <xsl:when test="$tttype='darkUp'">
        <xsl:value-of select="'ptn016'"/>
      </xsl:when>
      <xsl:when test="$tttype='shingle'">
        <xsl:value-of select="'ptn031'"/>
      </xsl:when>
      <xsl:when test="$tttype='darkGrid'">
        <xsl:value-of select="'ptn006'"/>
      </xsl:when>
      <xsl:when test="$tttype='lightHorizontal'">
        <xsl:value-of select="'ptn020'"/>
      </xsl:when>
      <xsl:when test="$tttype='lightVertical'">
        <xsl:value-of select="'ptn019'"/>
      </xsl:when>
      <xsl:when test="$tttype='lightDown'">
        <xsl:value-of select="'ptn013'"/>
      </xsl:when>
      <xsl:when test="$tttype='lightUp'">
        <xsl:value-of select="'ptn014'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="pattClr">
    <xsl:choose>
      <xsl:when test="a:schemeClr">
        <xsl:variable name="id">
          <xsl:value-of select="a:schemeClr/@val"/>
        </xsl:variable>
        <!--有5个主题色微软自己就没设置，在这里转为auto-->
        <!--/xsl:if-->
        <xsl:choose>
          <xsl:when test="$id='dk1'">
            <xsl:value-of select="'#000000'"/>
          </xsl:when>
          <xsl:when test="$id='lt1'">
            <xsl:value-of select="'#ffffff'"/>
          </xsl:when>
          <xsl:when test="$id='dk2'">
            <xsl:value-of select="'#1f497d'"/>
          </xsl:when>
          <xsl:when test="$id='lt2'">
            <xsl:value-of select="'#eeece1'"/>
          </xsl:when>
          <xsl:when test="$id='accent1'">
            <xsl:value-of select="'#4f81bd'"/>
          </xsl:when>
          <xsl:when test="$id='accent2'">
            <xsl:value-of select="'#c0504d'"/>
          </xsl:when>
          <xsl:when test="$id='accent3'">
            <xsl:value-of select="'#9bbb59'"/>
          </xsl:when>
          <xsl:when test="$id='accent4'">
            <xsl:value-of select="'#8064a2'"/>
          </xsl:when>
          <xsl:when test="$id='accent5'">
            <xsl:value-of select="'#4bacc6'"/>
          </xsl:when>
          <xsl:when test="$id='accent6'">
            <xsl:value-of select="'#f79646'"/>
          </xsl:when>
          <xsl:when test="$id='hlink'">
            <xsl:value-of select="'#0000ff'"/>
          </xsl:when>
          <xsl:when test="$id='folHlink'">
            <xsl:value-of select="'#800080'"/>
          </xsl:when>
          <xsl:when test="$id='tx2'">
            <xsl:value-of select="'#558ed5'"/>
          </xsl:when>
          <xsl:when test="$id='tx1'">
            <xsl:value-of select="'#000000'"/>
          </xsl:when>
          <!-- 20130508 updata by xuzhenwei 修改图表背景主题色 start -->
          <xsl:when test="$id='bg1'">
            <xsl:value-of select="'bg1'"/>
          </xsl:when>
          <!-- end -->
          <xsl:otherwise>
            <xsl:value-of select="'auto'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="a:srgbClr">
        <xsl:value-of select="concat('#',translate(a:srgbClr/@val,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="'auto'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!--uof的渐变填充-->
  <xsl:template match="a:schemeClr">
    <xsl:variable name="color">
      <xsl:value-of select="@val"/>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$color='dk1'">
        <xsl:value-of select="'#000000'"/>
      </xsl:when>
      <xsl:when test="$color='lt1'">
        <xsl:value-of select="'#ffffff'"/>
      </xsl:when>
      <xsl:when test="$color='dk2'">
        <xsl:value-of select="'#1f497d'"/>
      </xsl:when>
      <xsl:when test="$color='lt2'">
        <xsl:value-of select="'#eeece1'"/>
      </xsl:when>
      <xsl:when test="$color='accent1'">
        <xsl:value-of select="'#4f81bd'"/>
      </xsl:when>
      <xsl:when test="$color='accent2'">
        <xsl:value-of select="'#c0504d'"/>
      </xsl:when>
      <xsl:when test="$color='accent3'">
        <xsl:value-of select="'#9bbb59'"/>
      </xsl:when>
      <xsl:when test="$color='accent4'">
        <xsl:value-of select="'#8064a2'"/>
      </xsl:when>
      <xsl:when test="$color='accent5'">
        <xsl:value-of select="'#4bacc6'"/>
      </xsl:when>
      <xsl:when test="$color='accent6'">
        <xsl:value-of select="'#f79646'"/>
      </xsl:when>
      <xsl:when test="$color='hlink'">
        <xsl:value-of select="'#0000ff'"/>
      </xsl:when>
      <xsl:when test="$color='folHlink'">
        <xsl:value-of select="'#800080'"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="colorChoice">
    <xsl:choose>
      <xsl:when test="a:schemeClr">
        <xsl:apply-templates select="a:schemeClr">
        </xsl:apply-templates>
      </xsl:when>
      <xsl:when test="a:srgbClr">
        <xsl:variable name="color">
          <xsl:value-of select="a:srgbClr/@val"/>
        </xsl:variable>
        <xsl:value-of select="concat('#',translate($color,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

  <!--绘图区-->
  <xsl:template name="chart-huituqu">
    <xsl:param name="tar"/>
    <xsl:param name="target"/>
    <xsl:param name="rid"/>
    <xsl:variable name="aim">
      <xsl:value-of select="@Target"/>
    </xsl:variable>
    <xsl:variable name="aim2">
      <xsl:value-of select="substring-after(@Target,'../')"/>
    </xsl:variable>
    <xsl:variable name="aim3">
      <xsl:value-of select="concat('xlsx/xl/',$aim2)"/>
    </xsl:variable>
    <图表:绘图区_E747>
      <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:layout[c:lastLayout]">
        <图表:位置_E70A>
          <xsl:attribute name="x_C606">
            <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:layout/c:lastLayout/c:x/@val"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:layout/c:lastLayout/c:y/@val"/>
          </xsl:attribute>
        </图表:位置_E70A>
        <图表:大小_E748>
          <xsl:attribute name="长_C604">
            <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:layout/c:lastLayout/c:h/@val"/>
          </xsl:attribute>
          <xsl:attribute name="宽_C605">
            <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:layout/c:lastLayout/c:w/@val"/>
          </xsl:attribute>
        </图表:大小_E748>
      </xsl:if>
      <!--边框和填充-->
      <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:spPr]">
        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:spPr">
          <xsl:apply-templates select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:spPr">
            <xsl:with-param name="target1" select="$target"/>
          </xsl:apply-templates>
        </xsl:for-each>
      </xsl:if>
      <!-- add by xuzhenwei BUG_2488:股价图 控制背景填充色 2013-01-12 代码暂时没用到
      <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:spPr]">
        <xsl:apply-templates select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:spPr">
          <xsl:with-param name="target1" select="$target"/>
        </xsl:apply-templates>
      </xsl:if>-->
      <!-- end -->
      <!--按行按列-->
      <!--数据区域-->
      <!--图表类型组集-->
      <图表:图表类型组集_E74C>
            <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:plotArea]">
                <!--barChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:barChart]">
                  <图表:组_E74D>
                    <!--数据系列集-->
                      <图表:数据系列集_E74E>
                          <xsl:variable name="serIndex" select="0"/>
                          <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:barChart">
                              <xsl:apply-templates select="c:ser">
                                  <xsl:with-param name="sid" select="$serIndex+position()"/>
                                  <xsl:with-param name="tar" select="$tar"/>
                                  <xsl:with-param name="rid" select="$rid"/>
                                  <xsl:with-param name="target" select="$target"/>
                                  <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                  <xsl:with-param name="chart_type" select="'c:barChart'"/>
                                  <!-- end -->
                              </xsl:apply-templates>
                          </xsl:for-each>
                      </图表:数据系列集_E74E>
                    <xsl:call-template name="behindgroupset">
                      <xsl:with-param name="tar" select="$tar"/>
                      <xsl:with-param name="rid" select="$rid"/>
                      <xsl:with-param name="target" select="$target"/>
                      <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:barChart/c:gapWidth/@val"/>
                      <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:barChart/c:overlap/@val"/>
                    </xsl:call-template>
                    <!--重叠比例-->
                    <图表:重叠比例_E781>
                      <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:barChart/c:overlap/@val"/>
                    </图表:重叠比例_E781>
                    <!--分类间距-->
                    <图表:分类间距_E782>
                      <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:barChart/c:gapWidth/@val"/>
                    </图表:分类间距_E782>
                  </图表:组_E74D>
                </xsl:if>
                <!--bar3DChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:bar3DChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:bar3DChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:bar3DChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:bar3DChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:bar3DChart/c:overlap/@val"/>
                  </xsl:call-template>
                  <!--重叠比例-->
                  <图表:重叠比例_E781>
                    <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:bar3DChart/c:overlap/@val"/>
                  </图表:重叠比例_E781>
                  <!--分类间距-->
                  <图表:分类间距_E782>
                    <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:bar3DChart/c:gapWidth/@val"/>
                  </图表:分类间距_E782>
                  <!-- 20130515 add by xuzhenwei 用来控制三维效果 start -->
                  <!-- 三维效果：上下仰角、左转角度和透视系数 有待完善，这里只实现了柱状图且取默认 -->
                  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:view3D">
                    <图表:透视深度_E783>100</图表:透视深度_E783>
                  </xsl:if>
                  <!-- end -->
                </图表:组_E74D>
                </xsl:if>
                <!--lineChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:lineChart]">
                  <图表:组_E74D>
                    <!--数据系列集-->
                      <图表:数据系列集_E74E>
                          <xsl:variable name="serIndex" select="0"/>
                          <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:lineChart">
                              <xsl:apply-templates select="c:ser">
                                  <xsl:with-param name="sid" select="$serIndex+position()"/>
                                  <xsl:with-param name="tar" select="$tar"/>
                                  <xsl:with-param name="rid" select="$rid"/>
                                  <xsl:with-param name="target" select="$target"/>
                                  <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                  <xsl:with-param name="chart_type" select="'c:lineChart'"/>
                                  <!-- end -->
                              </xsl:apply-templates>
                          </xsl:for-each>
                      </图表:数据系列集_E74E>
                    <xsl:call-template name="behindgroupset">
                      <xsl:with-param name="tar" select="$tar"/>
                      <xsl:with-param name="rid" select="$rid"/>
                      <xsl:with-param name="target" select="$target"/>
                      <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:lineChart/c:gapWidth/@val"/>
                      <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:lineChart/c:overlap/@val"/>
                    </xsl:call-template>
                    <!--重叠比例-->
                    <图表:重叠比例_E781>
                      <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:lineChart/c:overlap/@val"/>
                    </图表:重叠比例_E781>
                    <!--分类间距-->
                    <图表:分类间距_E782>
                      <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:lineChart/c:gapWidth/@val"/>
                    </图表:分类间距_E782>
                  </图表:组_E74D>
                </xsl:if>
                <!--line3DChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:line3DChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:line3DChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:line3DChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:line3DChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:line3DChart/c:overlap/@val"/>
                  </xsl:call-template>
                  <!--重叠比例-->
                  <图表:重叠比例_E781>
                    <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:line3DChart/c:overlap/@val"/>
                  </图表:重叠比例_E781>
                  <!--分类间距-->
                  <图表:分类间距_E782>
                    <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:line3DChart/c:gapWidth/@val"/>
                  </图表:分类间距_E782>
                </图表:组_E74D>
                </xsl:if>
                <!--pieChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:pieChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:pieChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:pieChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:pieChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:pieChart/c:overlap/@val"/>
                  </xsl:call-template>
                </图表:组_E74D>
                </xsl:if>
                <!--pie3DChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:pie3DChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:pie3DChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:pie3DChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:pie3DChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:pie3DChart/c:overlap/@val"/>
                  </xsl:call-template>
                </图表:组_E74D>
                </xsl:if>
                <!--ofPieChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:ofPieChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:ofPieChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:ofPieChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:ofPieChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:ofPieChart/c:overlap/@val"/>
                  </xsl:call-template>
                </图表:组_E74D>
                </xsl:if>
                <!--scatterChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:scatterChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:scatterChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:scatterChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:scatterChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:scatterChart/c:overlap/@val"/>
                  </xsl:call-template>
                </图表:组_E74D>
                </xsl:if>
                <!--areaChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:areaChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:areaChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:areaChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:areaChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:areaChart/c:overlap/@val"/>
                  </xsl:call-template>
                </图表:组_E74D>
                </xsl:if>
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:area3DChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:area3DChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:area3DChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:area3DChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:area3DChart/c:overlap/@val"/>
                  </xsl:call-template>
                </图表:组_E74D>
                </xsl:if>
                <!--doughnutChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:doughnutChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:doughnutChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:doughnutChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:doughnutChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:doughnutChart/c:overlap/@val"/>
                  </xsl:call-template>
                </图表:组_E74D>
                </xsl:if>
                <!--radarChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:radarChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:radarChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:radarChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:radarChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:radarChart/c:overlap/@val"/>
                  </xsl:call-template>
                </图表:组_E74D>
                </xsl:if>
                <!--stockChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:stockChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:stockChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:stockChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:stockChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:stockChart/c:overlap/@val"/>
                  </xsl:call-template>
                </图表:组_E74D>
                </xsl:if>
                <!--surfaceChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:surface3DChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:surface3DChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:surface3DChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:surface3DChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:surface3DChart/c:overlap/@val"/>
                  </xsl:call-template>
                </图表:组_E74D>
                </xsl:if>
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:surfaceChart]">
                  <图表:组_E74D>
                    <!--数据系列集-->
                      <图表:数据系列集_E74E>
                          <xsl:variable name="serIndex" select="0"/>
                          <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:surfaceChart">
                              <xsl:apply-templates select="c:ser">
                                  <xsl:with-param name="sid" select="$serIndex+position()"/>
                                  <xsl:with-param name="tar" select="$tar"/>
                                  <xsl:with-param name="rid" select="$rid"/>
                                  <xsl:with-param name="target" select="$target"/>
                                  <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                  <xsl:with-param name="chart_type" select="'c:surfaceChart'"/>
                                  <!-- end -->
                              </xsl:apply-templates>
                          </xsl:for-each>
                      </图表:数据系列集_E74E>
                    <xsl:call-template name="behindgroupset">
                      <xsl:with-param name="tar" select="$tar"/>
                      <xsl:with-param name="rid" select="$rid"/>
                      <xsl:with-param name="target" select="$target"/>
                      <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:surfaceChart/c:gapWidth/@val"/>
                      <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:surfaceChart/c:overlap/@val"/>
                    </xsl:call-template>
                  </图表:组_E74D>
                </xsl:if>
                <!--bubbleChart-->
                <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:bubbleChart]">
                <图表:组_E74D>
                  <!--数据系列集-->
                    <图表:数据系列集_E74E>
                        <xsl:variable name="serIndex" select="0"/>
                        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:bubbleChart">
                            <xsl:apply-templates select="c:ser">
                                <xsl:with-param name="sid" select="$serIndex+position()"/>
                                <xsl:with-param name="tar" select="$tar"/>
                                <xsl:with-param name="rid" select="$rid"/>
                                <xsl:with-param name="target" select="$target"/>
                                <!-- add by xuzhenwei BUG_2488: 股价图，控制股价图类型 2013-01-11 start -->
                                <xsl:with-param name="chart_type" select="'c:bubbleChart'"/>
                                <!-- end -->
                            </xsl:apply-templates>
                        </xsl:for-each>
                    </图表:数据系列集_E74E>
                  <xsl:call-template name="behindgroupset">
                    <xsl:with-param name="tar" select="$tar"/>
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                    <xsl:with-param name="gapWidth" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:bubbleChart/c:gapWidth/@val"/>
                    <xsl:with-param name="overlap" select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:bubbleChart/c:overlap/@val"/>
                  </xsl:call-template>
                </图表:组_E74D>
                </xsl:if>
              </xsl:if>
      </图表:图表类型组集_E74C>
      <!--坐标轴集-->
      <图表:坐标轴集_E790>
        <xsl:choose>

            <!-- update by 凌峰 BUG_3080：坐标轴标题丢失，坐标刻度值变化 20140308 start -->
          <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:scatterChart or c:bubbleChart]">
            <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx[position()=1]">
              <图表:坐标轴_E791>
                <xsl:call-template name="chart-shuzhizhou">
                    <xsl:with-param name="pos" select="1"/>
                  <xsl:with-param name="rid" select="$rid"/>
                  <xsl:with-param name="target" select="$target"/>
                </xsl:call-template>
              </图表:坐标轴_E791>
            </xsl:for-each>
            <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx[position()=2]">
              <图表:坐标轴_E791>
                <xsl:call-template name="chart-shuzhizhou">
                    <xsl:with-param name="pos" select="2"/>
                  <xsl:with-param name="rid" select="$rid"/>
                  <xsl:with-param name="target" select="$target"/>
                </xsl:call-template>
                  <!--end-->
                  
              </图表:坐标轴_E791>
            </xsl:for-each>
          </xsl:when>
          <xsl:otherwise>
            <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:catAx]">
              <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx">
                <图表:坐标轴_E791>
                  <xsl:call-template name="chart-fenleizhou">
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                  </xsl:call-template>
                </图表:坐标轴_E791>
              </xsl:for-each>
            </xsl:if>
            <!-- 20130416 update by xuzhenwei BUG_2830 互操作 oo-uof（编辑）-oo 024实用文档-损益表(1).xlsx 文档需要修复 start -->
            <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:dateAx]">
              <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:dateAx">
                <图表:坐标轴_E791>
                  <xsl:call-template name="chart-fenleizhou">
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                  </xsl:call-template>
                </图表:坐标轴_E791>
              </xsl:for-each>
            </xsl:if>
            <!-- end -->
            <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[c:valAx]">
              <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx">
                <图表:坐标轴_E791>
                  <xsl:call-template name="chart-shuzhizhou">
                    
                    <!--2014-4-22, add by Qihy， 凌峰在模板中增加了参数， 该处需要进行相应修改-->
                    <xsl:with-param name="pos" select="position()"/>
                    <!--2014-4-22 end-->
                    
                    <xsl:with-param name="rid" select="$rid"/>
                    <xsl:with-param name="target" select="$target"/>
                  </xsl:call-template>
                </图表:坐标轴_E791>
              </xsl:for-each>
            </xsl:if>
          </xsl:otherwise>
        </xsl:choose>
      </图表:坐标轴集_E790>
    </图表:绘图区_E747>
  </xsl:template>
  <xsl:template name="behindgroupset">
    <xsl:param name="tar"/>
    <xsl:param name="target"/>
    <xsl:param name="rid"/>
    <xsl:param name="gapWidth"/>
    <xsl:param name="overlap"/>
    <!--是否依据点着色-->
    <!--系列线-->
    <!--垂直线-->
    <!--高低线-->
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:stockChart[c:hiLowLines]">
      <图表:高低线_E77D>
        <xsl:choose>
          <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:stockChart/c:hiLowLines[c:spPr]">
            <xsl:apply-templates select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:stockChart/c:hiLowLines/c:spPr"/>
          </xsl:when>
          <xsl:otherwise>
            <图表:边框线_4226 线型_C60D="single" 虚实_C60E="solid" 宽度_C60F="2" 颜色_C611="#808080"/>
          </xsl:otherwise>
        </xsl:choose>
      </图表:高低线_E77D>
    </xsl:if>
    <!-- add by xuzhenwei BUG_2488:股价图 2013-01-11 -->
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:stockChart[c:upDownBars]">
      <!--跌柱线-->
      <图表:跌柱线_E77E>
        <xsl:choose>
          <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:stockChart/c:upDownBars/c:upBars[c:spPr]">
            <xsl:apply-templates select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:stockChart/c:upDownBars/c:upBars/c:spPr"/>
          </xsl:when>
          <xsl:otherwise>
            <图表:填充_E746>
              <图表:填充_E746>auto</图表:填充_E746>
            </图表:填充_E746>
          </xsl:otherwise>
        </xsl:choose>
      </图表:跌柱线_E77E>
      <!--涨柱线-->
      <图表:涨柱线_E780>
        <xsl:choose>
          <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:stockChart/c:upDownBars/c:downBars[c:spPr]">
            <xsl:apply-templates select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:stockChart/c:upDownBars/c:downBars/c:spPr"/>
          </xsl:when>
          <xsl:otherwise>
            <图表:填充_E746>
              <图表:填充_E746>auto</图表:填充_E746>
            </图表:填充_E746>
          </xsl:otherwise>
        </xsl:choose>
      </图表:涨柱线_E780>
    </xsl:if>
    <!-- end -->
    <!--透视深度-->
    <!--系列间距-->
    <!--饼图分割-->
    <!--圆环图内径大小-->
    <!--是否互补色代表负值-->
    <!--第一扇区起始角-->
    <!--第二绘图区大小-->
    <!--气泡缩放大小-->
    <!--分离度-->
    <!--是否显示负气泡-->
    <!--大小表示-->

    <!--2014-6-8, update by Qihy, 饼图填充色不正确， start-->
    <!--<xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea[//c:grouping/@val='stacked' or //c:grouping/@val='percentStacked']">-->
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea//c:grouping/@val='stacked' or ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea//c:grouping/@val='percentStacked'">
      <图表:是否依数据点着色_E77A>false</图表:是否依数据点着色_E77A>
      <图表:是否互补色代表负值_E789>false</图表:是否互补色代表负值_E789>
    </xsl:if>
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:pieChart/c:varyColors[@val='0']">
      <图表:是否依数据点着色_E77A>false</图表:是否依数据点着色_E77A>
    </xsl:if>
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:pieChart/c:varyColors[@val='1']">
      <图表:是否依数据点着色_E77A>true</图表:是否依数据点着色_E77A>
    </xsl:if>
    <!--2014-6-8 end-->

  </xsl:template>
  
  <xsl:template match="c:ser">
      <xsl:param name="sid"/>
    <xsl:param name="tar"/>
    <xsl:param name="target"/>
    <xsl:param name="rid"/>
    <xsl:param name="chart_type"/>
   
    <xsl:if test="name()='c:ser'">
      <!--<xsl:for-each select="c:ser">-->
      <!--数据系列集-->
      <xsl:variable name="pos">
        <xsl:number count="c:ser" level="any" format="1"/>
      </xsl:variable>
      <图表:数据系列_E74F>
        <!--名称-->
        <!--<xsl:if test="c:dLbls/c:dLbl">
          <xsl:attribute name="名称_E774">
            <xsl:if test="c:dLbls/c:dLbl/c:tx/c:rich/a:p/a:r/a:t">
              <xsl:value-of select="c:dLbls/c:dLbl/c:tx/c:rich/a:p/a:r/a:t"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:if>-->
        <xsl:if test="c:tx/c:strRef">
          <xsl:attribute name="名称_E774">
            <xsl:if test="c:tx/c:strRef/c:f">
              <xsl:value-of select="concat('=',c:tx/c:strRef/c:f)"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:if>
        <!--值-->
        <xsl:choose>
          <xsl:when test="./c:yVal">
            <!--值-->
            <xsl:attribute name="值_E775">
              <xsl:call-template name="caculateValue">
                <xsl:with-param name="gValue" select="./c:yVal/c:numRef/c:f"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <!--值-->
            <xsl:attribute name="值_E775">
              <xsl:call-template name="caculateValue">
                <xsl:with-param name="gValue" select="./c:val/c:numRef/c:f"/>
              </xsl:call-template>
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <!--分类名-->
        <!-- 20130518 update by xuzhenwei BUG_2881:回归集成oo-uof工作表"折线图"分类轴X轴转换前为年月，转换后为数字 start -->
        <xsl:if test="./c:xVal/c:strRef or ./c:cat/c:strRef or ./c:cat/c:numRef or ./c:xVal/c:numRef">
          <xsl:attribute name="分类名_E776">
            <xsl:choose>
              <xsl:when test="./c:xVal/c:strRef">
                <xsl:call-template name="caculateValue">
                  <xsl:with-param name="gValue" select="./c:xVal/c:strRef/c:f"/>
                </xsl:call-template>
              </xsl:when>
                <xsl:when test="./c:xVal/c:numRef">
                    <xsl:call-template name="caculateValue">
                        <xsl:with-param name="gValue" select="./c:xVal/c:numRef/c:f"/>
                    </xsl:call-template>
                </xsl:when>
              <xsl:when test="./c:cat/c:numRef">
                <xsl:call-template name="caculateValue">
                  <xsl:with-param name="gValue" select="./c:cat/c:numRef/c:f"/>
                </xsl:call-template>
              </xsl:when>
              <!-- end -->
              <xsl:otherwise>
                <xsl:call-template name="caculateValue">
                  <xsl:with-param name="gValue" select="./c:cat/c:strRef/c:f"/>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
        </xsl:if>
        <xsl:variable name="chartType">
          <xsl:choose>
            <!-- update by xuzhenwei BUG_2484:簇状圆柱图，圆锥图，棱形图 start -->
            <!--柱形图-->
            <xsl:when test="$chart_type='c:barChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:barChart/c:barDir[@val='col']">
              <xsl:value-of select="'column'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='col' and c:shape/@val='box']">
              <xsl:value-of select="'column'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='col' and c:shape/@val='cylinder']">
              <xsl:value-of select="'surface'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='col' and c:shape/@val='cone']">
              <xsl:value-of select="'column'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='col' and c:shape/@val='pyramid']">
              <xsl:value-of select="'column'"/>
            </xsl:when>
            <!--条形图-->
            <xsl:when test="$chart_type='c:barChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:barChart/c:barDir[@val='bar']">
              <xsl:value-of select="'bar'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='bar' and c:shape/@val='box']">
              <xsl:value-of select="'bar'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='bar' and c:shape/@val='cylinder']">
              <xsl:value-of select="'surface'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='bar' and c:shape/@val='cone']">
              <xsl:value-of select="'bar'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='bar' and c:shape/@val='pyramid']">
              <xsl:value-of select="'bar'"/>
            </xsl:when>
			
			<!--zl 2015-4-14-->
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='bar' and c:shape/@val='cylinder']">
              <xsl:value-of select="'bar'"/>
            </xsl:when>
            <!--zl 2015-4-14-->     
			
            <!--股价图-->
            <xsl:when test="$chart_type='c:stockChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:stockChart">
              <xsl:value-of select="'line'"/>
            </xsl:when>
            <!--圆饼图-->
            <xsl:when test="$chart_type='c:pieChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:pieChart">
              <xsl:value-of select="'pie'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:pie3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:pie3DChart">
              <xsl:value-of select="'pie'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:ofPieChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:ofPieChart">
              <xsl:value-of select="'pie'"/>
            </xsl:when>
            <!--散点图-->
            <xsl:when test="$chart_type='c:scatterChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:scatterChart">
              <xsl:value-of select="'scatter'"/>
            </xsl:when>
            <!--面积图-->
            <xsl:when test="$chart_type='c:areaChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:areaChart">
              <xsl:value-of select="'area'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:area3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:area3DChart">
              <xsl:value-of select="'area'"/>
            </xsl:when>
            <!--环形图-->
            <xsl:when test="$chart_type='c:doughnutChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:doughnutChart">
              <xsl:value-of select="'doughnut'"/>
            </xsl:when>
            <!--雷达图-->
            <xsl:when test="$chart_type='c:radarChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:radarChart">
              <xsl:value-of select="'radar'"/>
            </xsl:when>
            <!--曲面图-->
            <!--<xsl:when test="$chart_type='c:surface3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:surface3DChart">
              <xsl:value-of select="'surface'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:surfaceChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:surfaceChart">
              <xsl:value-of select="'surface'"/>
            </xsl:when>-->
            <!--气泡图-->
            <xsl:when test="$chart_type='c:bubbleChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bubbleChart">
              <xsl:value-of select="'bubble'"/>
            </xsl:when>
            <!--折线图-->
            <xsl:when test="$chart_type='c:lineChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart">
              <xsl:value-of select="'line'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:line3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:line3DChart">
              <xsl:value-of select="'line'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:line3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:line3DChart">
              <xsl:value-of select="'line'"/>
            </xsl:when>
            
            <!-- end -->
            <!--默认柱形图-->
            <xsl:otherwise>
              <xsl:value-of select="'column'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <!--子类型-->
        <xsl:variable name="childType">
          <xsl:choose>
            <!-- update by xuzhenwei BUG_2484,2486: 簇状圆柱图，圆锥图，棱锥图子类型 start -->
            <!--条形图及柱形图子类型-->

            <!--zl 20150107 start-->
            <!--圆柱-->
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart/c:grouping[@val='clustered']  and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart/c:shape[@val='bar']">
              <xsl:value-of select="'clustered-3d'"/>
            </xsl:when>

			
            <!--簇状圆柱 zl 2015-4-13-->
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart/c:grouping[@val='clustered']  and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart/c:shape[@val='cylinder']">
              <xsl:value-of select="'clustered-3d'"/>
            </xsl:when>
			
            <!--簇状圆锥图-->
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart/c:grouping[@val='clustered'] and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart/c:shape[@val='cone']">
              <xsl:value-of select="'clustered-cone'"/>
            </xsl:when>

            <!--堆积圆锥-->
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:shape/@val='cone' and c:grouping/@val='stacked']">
              <xsl:value-of select="'stacked-cone'"/>
            </xsl:when>

            <!--百分比堆积圆锥-->
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:shape/@val='cone' and c:grouping/@val='percentStacked']">
              <xsl:value-of select="'percent-cone'"/>
            </xsl:when>

            <!--三维圆锥-->
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:shape/@val='cone' and c:grouping/@val='standard']">
              <xsl:value-of select="'3D-cone'"/>
            </xsl:when>

            <!--簇状菱锥图,堆积菱锥,百分比堆积菱锥,三维菱锥-->
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:shape/@val='pyramid' and c:grouping/@val='clustered']">
              <xsl:value-of select="'clustered-pyramid'"/>
            </xsl:when>

            <!--堆积菱锥,百分比堆积菱锥,三维菱锥-->
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:shape/@val='pyramid' and c:grouping/@val='stacked']">
              <xsl:value-of select="'stacked-pyramid'"/>
            </xsl:when>

            <!--百分比堆积菱锥,三维菱锥-->
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:shape/@val='pyramid' and c:grouping/@val='percentStacked']">
              <xsl:value-of select="'percent-pyramid'"/>
            </xsl:when>

            <!--三维菱锥-->
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:shape/@val='pyramid' and c:grouping/@val='standard']">
              <xsl:value-of select="'3D-pyramid'"/>
            </xsl:when>
            <!--zl end-->
            
            <xsl:when test="$chart_type='c:barChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:barChart/c:grouping[@val='clustered']">
              <xsl:value-of select="'clustered'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:barChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:barChart/c:grouping[@val='stacked']">
              <xsl:value-of select="'stacked'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:barChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:barChart/c:grouping[@val='percentStacked']">
              <xsl:value-of select="'percent-stacked'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:shape/@val='box' and c:grouping/@val='clustered']">
              <xsl:value-of select="'clustered-3d'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:shape/@val='cylinder' and c:grouping/@val='clustered']">
              <xsl:value-of select="'clustered'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:shape/@val='box' and c:grouping/@val='stacked']">
              <xsl:value-of select="'exploded-3d'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:shape/@val='box' and c:grouping/@val='percentStacked']">
              <xsl:value-of select="'percent-stacked-marker'"/>
            </xsl:when> 
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='col' and c:grouping/@val='clustered']">
              <xsl:value-of select="'clustered'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='col' and c:grouping/@val='stacked']">
              <xsl:value-of select="'stacked'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='col' and c:grouping/@val='percentStacked']">
              <xsl:value-of select="'percent-stacked'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart/c:grouping[@val='standard']">
              <xsl:value-of select="'standard-3d'"/>
            </xsl:when>
			
			<!--zhaolin-->
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='bar' and c:grouping/@val='clustered']">
              <xsl:value-of select="'clustered-cone'"/>
            </xsl:when>
			
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='bar' and c:grouping/@val='stacked']">
              <xsl:value-of select="'exploded-3d'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:bar3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bar3DChart[c:barDir/@val='bar' and c:grouping/@val='percentStacked']">
              <xsl:value-of select="'percent-stacked-marker'"/>
            </xsl:when>
            <!--圆饼图子类型-->
            <xsl:when test="$chart_type='c:pieChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:pieChart/c:ser[not(c:explosion)]">
              <xsl:value-of select="'standard'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:pie3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:pie3DChart/c:ser[not(c:explosion)]">
              <xsl:value-of select="'standard-3d'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:ofPieChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:ofPieChart/c:ofPieType/@val='pie'">
              <xsl:value-of select="'of-pie'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:pieChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:pieChart/c:ser[c:explosion]">
              <xsl:value-of select="'exploded'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:pie3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:pie3DChart/c:ser[c:explosion]">
              <xsl:value-of select="'exploded-3d'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:ofPieChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:ofPieChart/c:ofPieType/@val='bar'">
              <xsl:value-of select="'of-bar'"/>
            </xsl:when>
            <!--面积图子类型-->
            <xsl:when test="$chart_type='c:areaChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:areaChart/c:grouping [@val='standard']">
              <xsl:value-of select="'standard'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:area3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:area3DChart/c:grouping [@val='standard']">
              <xsl:value-of select="'standard'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:areaChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:areaChart/c:grouping [@val='stacked'] or ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:area3DChart/c:grouping [@val='stacked']">
              <xsl:value-of select="'stacked'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:areaChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:areaChart/c:grouping [@val='percentStacked'] or ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:area3DChart/c:grouping [@val='percentStacked']">
              <xsl:value-of select="'percent-stacked'"/>
            </xsl:when>
            <!--圆环图子类型-->
            <xsl:when test="$chart_type='c:doughnutChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:doughnutChart/c:ser[not(c:explosion)]">
              <xsl:value-of select="'standard'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:doughnutChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:doughnutChart/c:ser[c:explosion]">
              <xsl:value-of select="'exploded'"/>
            </xsl:when>
            <!--雷达图子类型-->
            <xsl:when test="$chart_type='c:radarChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:radarChart[c:radarStyle/@val='marker' and c:ser[c:marker]]">
              <xsl:value-of select="'standard'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:radarChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:radarChart[c:radarStyle/@val='marker' and c:ser[not(c:marker)]]">
              <xsl:value-of select="'standard-marker'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:radarChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:radarChart/c:radarStyle/@val='filled'">
              <xsl:value-of select="'fill'"/>
            </xsl:when>
            <!--散点图子类型-->
            <xsl:when test="$chart_type='c:scatterChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:scatterChart[c:scatterStyle/@val='marker' and c:ser[c:spPr]]">
              <xsl:value-of select="'marker'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:scatterChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:scatterChart[c:scatterStyle/@val='smoothMarker' and c:ser[not(c:marker)]]">
              <xsl:value-of select="'smooth-marker'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:scatterChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:scatterChart[c:scatterStyle/@val='smoothMarker' and c:ser[c:marker]]">
              <xsl:value-of select="'smooth'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:scatterChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:scatterChart[c:scatterStyle/@val='lineMarker' and c:ser[not(c:marker)]]">
              <xsl:value-of select="'line-marker'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:scatterChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:scatterChart[c:scatterStyle/@val='lineMarker' and c:ser[c:marker]]">
              <xsl:value-of select="'line'"/>
            </xsl:when>
            <!--气泡图子类型-->
            <xsl:when test="$chart_type='c:bubbleChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:bubbleChart">
              <xsl:value-of select="'standard'"/>
            </xsl:when>
            <!--折线图子类型-->
            <xsl:when test="$chart_type='c:lineChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart/c:grouping[@val='standard'] and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart/c:ser/c:marker/c:symbol[@val='none']">
              <xsl:value-of select="'clustered'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:lineChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart/c:grouping[@val='stacked'] and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart/c:ser/c:marker/c:symbol[@val='none']">
              <xsl:value-of select="'stacked'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:lineChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart/c:grouping[@val='percentStacked'] and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart/c:ser/c:marker/c:symbol[@val='none']">
              <xsl:value-of select="'percent-stacked'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:lineChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart/c:grouping[@val='standard'] and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart/c:ser[not(c:marker)]">
              <xsl:value-of select="'standard-marker'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:lineChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart/c:grouping[@val='stacked'] and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart/c:ser[not(c:marker)]">
              <xsl:value-of select="'stacked-marker'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:lineChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart/c:grouping[@val='percentStacked'] and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:lineChart/c:ser[not(c:marker)]">
              <xsl:value-of select="'percent-stacked-marker'"/>
            </xsl:when>
            <xsl:when test="$chart_type='c:line3DChart' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea/c:line3DChart/c:grouping[@val='standard']">
              <xsl:value-of select="'clustered'"/>
            </xsl:when>
            <!-- end -->
            <xsl:otherwise>
              <xsl:value-of select="'clustered'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <!--类型-->
        <xsl:attribute name="类型_E75D">
          <xsl:value-of select="$chartType"/>
        </xsl:attribute>
        
        <xsl:attribute name="子类型_E777">
          <xsl:value-of select="$childType"/>
        </xsl:attribute>
        <!--标识符-->
        <xsl:variable name="idx" select="c:idx/@val"/>
        <xsl:attribute name="标识符_E778">
          <xsl:value-of select="$idx"/>
        </xsl:attribute>
        <!--系列坐标系-->
        <xsl:attribute name="系列坐标系_E779">
          <xsl:value-of select="'primary'"/>
        </xsl:attribute>
        <!--边框+填充-->
                <xsl:if test="c:spPr/*">
                    <xsl:apply-templates select="c:spPr">
                        <!--hwj-->
                        <xsl:with-param name="target1" select="$target"/>
                        <xsl:with-param name="chart_type" select="$chartType"/>
                        <xsl:with-param name="child_type" select="$childType"/>
                    </xsl:apply-templates>
            </xsl:if>
        <!--边框+填充-->
        <xsl:if test="not(c:spPr)">
          <图表:填充_E746>
            <图:颜色_8004>auto</图:颜色_8004>
          </图表:填充_E746>
        </xsl:if>
        <!--<xsl:if test="not(c:spPr/*)">
                        <图表:填充_E746>
                          <图:颜色_8004>auto</图:颜色_8004>
                        </图表:填充_E746>
                      </xsl:if>-->

        <!--2014-4-21, add by Qihy, 修复bug3079，增加非图片，固定，渐变填充之外的填充，由于UOF没有对应的填充，此处暂时转为auto-->
        <xsl:if test="c:spPr/a:scene3d">
          <图表:填充_E746>
            <xsl:attribute name="是否填充随图形旋转_8067">false</xsl:attribute>
            <图:颜色_8004>auto</图:颜色_8004>
          </图表:填充_E746>
        </xsl:if>
        <!--2014-4-21 end-->

        <xsl:if test="c:spPr">
          <xsl:if test="c:spPr/a:ln and not(c:spPr/a:solidFill) and not(c:spPr/a:blipFill) and not (c:spPr/a:gradFill)">
            <图表:填充_E746>
              <图:颜色_8004>auto</图:颜色_8004>
            </图表:填充_E746>
          </xsl:if>
        </xsl:if>
        <!--数据标签-->
        <xsl:if test="c:dLbls">
          <xsl:for-each select="./c:dLbls">
            <xsl:call-template name="DisplayLabel"/>
          </xsl:for-each>
        </xsl:if>

        <xsl:if test="not(c:dLbls)">
          <xsl:if test="parent::node()/c:dLbls">
            <xsl:for-each select="parent::node()/c:dLbls">
              <xsl:call-template name="DisplayLabel"/>
            </xsl:for-each>
          </xsl:if>
        </xsl:if>
        <!--数据标记-->
        <!-- add by xuzhenwei BUG_2488: 股价图 显示数据标签 2013-01-11 start -->
        <!-- 数据标记 -->
        <xsl:if test="c:marker">
          <图表:数据标记_E70E>
            <xsl:attribute name="类型_E70F">
              <xsl:choose>
                <xsl:when test="c:marker/c:symbol/@val='square'">
                  <xsl:value-of select="'square'"/>
                </xsl:when>
                <xsl:when test="c:marker/c:symbol/@val='dash'">
                  <xsl:value-of select="'dash'"/>
                </xsl:when>
                <xsl:when test="c:marker/c:symbol/@val='plus'">
                  <xsl:value-of select="'plus'"/>
                </xsl:when>
                <xsl:when test="c:marker/c:symbol/@val='circle'">
                  <xsl:value-of select="'circle'"/>
                </xsl:when>
                <xsl:when test="c:marker/c:symbol/@val='star'">
                  <xsl:value-of select="'star'"/>
                </xsl:when>
                <xsl:when test="c:marker/c:symbol/@val='triangle'">
                  <xsl:value-of select="'triangle'"/>
                </xsl:when>
                <xsl:when test="c:marker/c:symbol/@val='diamond'">
                  <xsl:value-of select="'diamond'"/>
                </xsl:when>
                <xsl:when test="c:marker/c:symbol/@val='x'">
                  <xsl:value-of select="'x'"/>
                </xsl:when>
                <xsl:when test="c:marker/c:symbol/@val='dot'">
                  <xsl:value-of select="'half-line'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'none'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="前景色_E710">
              <xsl:choose>
                <xsl:when test="c:marker/c:spPr/a:ln/a:solidFill">
                  <xsl:value-of select="c:marker/c:spPr/a:ln/a:solidFill/a:srgbClr/@val"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'auto'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="背景色_E711">
              <xsl:choose>
                <xsl:when test="c:marker/c:spPr/a:ln/a:solidFill">
                  <xsl:value-of select="./a:srgbClr/@val"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'auto'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <xsl:attribute name="大小_E712">
              <xsl:value-of select="c:marker/c:size/@val"/>
            </xsl:attribute>
          </图表:数据标记_E70E>
        </xsl:if>
        <!-- end -->
          <!-->
          <xsl:if test="not(c:marker)">
              <图表:数据标记_E70E>
                  <xsl:attribute name="类型_E70F">
                      <xsl:choose>
                          <xsl:when test="$sid=5">
                              <xsl:value-of select="'square'"/>
                          </xsl:when>
                          <xsl:when test="$sid=4">
                              <xsl:value-of select="'square-cross'"/>
                          </xsl:when>
                          <xsl:when test="$sid=6">
                              <xsl:value-of select="'dash'"/>
                          </xsl:when>
                          <xsl:when test="$sid=7">
                              <xsl:value-of select="'plus'"/>
                          </xsl:when>
                          <xsl:when test="$sid=8">
                              <xsl:value-of select="'circle'"/>
                          </xsl:when>
                          <xsl:when test="$sid=9">
                              <xsl:value-of select="'star'"/>
                          </xsl:when>
                          <xsl:when test="$sid=3">
                              <xsl:value-of select="'triangle'"/>
                          </xsl:when>
                          <xsl:when test="$sid=1">
                              <xsl:value-of select="'diamond'"/>
                          </xsl:when>
                          <xsl:when test="$sid=10">
                              <xsl:value-of select="'x'"/>
                          </xsl:when>
                          <xsl:when test="$sid=11">
                              <xsl:value-of select="'half-line'"/>
                          </xsl:when>
                          <xsl:otherwise>
                              <xsl:value-of select="'none'"/>
                          </xsl:otherwise>
                      </xsl:choose>
                  </xsl:attribute>
                  <xsl:attribute name="前景色_E710">
                      <xsl:value-of select="'auto'"/>
                  </xsl:attribute>
                  <xsl:attribute name="背景色_E711">
                      <xsl:value-of select="'auto'"/>
                  </xsl:attribute>
                  <xsl:attribute name="大小_E712">
                      <xsl:value-of select="5"/>
                  </xsl:attribute>
              </图表:数据标记_E70E>
          </xsl:if>
          <-->
        <!--数据点集-->
        <xsl:if test="./c:dPt">
          <图表:数据点集_E755>
            <xsl:apply-templates select="./c:dPt">
              <xsl:with-param name="target" select="$target"/>
            </xsl:apply-templates>
          </图表:数据点集_E755>
        </xsl:if>
        <xsl:if test="not(./c:dPt) and ($chartType = 'pie') and ($childType = 'exploded-3d') and ($idx = 0)">
          <图表:数据点集_E755>
            <图表:数据点_E756 点_E757="1">
              <图表:填充_E746 是否填充随图形旋转_8067="false">
                <图:颜色_8004>#3366ff</图:颜色_8004>
              </图表:填充_E746>
              <图表:是否补色代表负值_E713>false</图表:是否补色代表负值_E713>
            </图表:数据点_E756>
            <图表:数据点_E756 点_E757="2">
              <图表:填充_E746 是否填充随图形旋转_8067="false">
                <图:颜色_8004>#b93019</图:颜色_8004>
              </图表:填充_E746>
              <图表:是否补色代表负值_E713>false</图表:是否补色代表负值_E713>
            </图表:数据点_E756>
            <图表:数据点_E756 点_E757="3">
              <图表:填充_E746 是否填充随图形旋转_8067="false">
                <图:颜色_8004>#51941b</图:颜色_8004>
              </图表:填充_E746>
              <图表:是否补色代表负值_E713>false</图表:是否补色代表负值_E713>
            </图表:数据点_E756>
            <图表:数据点_E756 点_E757="4">
              <图表:填充_E746 是否填充随图形旋转_8067="false">
                <图:颜色_8004>#7c5f9c</图:颜色_8004>
              </图表:填充_E746>
              <图表:是否补色代表负值_E713>false</图表:是否补色代表负值_E713>
            </图表:数据点_E756>
          </图表:数据点集_E755>
        </xsl:if>
        <!--<xsl:call-template name="chart-sjdj">
          <xsl:with-param name="tar" select="$tar"/>
          <xsl:with-param name="rid" select="$rid"/>
          <xsl:with-param name="target" select="$target"/>
        </xsl:call-template>-->
        <!--引导线 不转-->
        <!--误差线集 不转-->
        <!--趋势线集-->
        <xsl:if test="./c:trendline">
          <图表:趋势线集_E762>
            <图表:趋势线_E763 类型_E76C="linear" 前推预测周期_E772="0.0" 倒推预测周期_E773="0.0">
              <图表:公式_E834 是否显示公式_E770="false" 是否显示R平方值_E771="false"/>
            </图表:趋势线_E763>
          </图表:趋势线集_E762>
        </xsl:if>
        
        <!--气泡大小区域-->
        <xsl:if test="$chartType = 'bubble'">
          <图表:气泡大小区域_E83F>
            <!-- update by xuzhenwei BUG_2464,2487：气泡图丢失 2013-01-12 start -->
            <xsl:call-template name="caculateValue">
              <xsl:with-param name="gValue" select="./c:yVal/c:numRef/c:f"/>
            </xsl:call-template>
            <!-- end -->
          </图表:气泡大小区域_E83F>
        </xsl:if>
        <!--是否平滑线-->
        <xsl:if test="contains($childType,'smooth')">
          <图表:是否平滑线_E840>true</图表:是否平滑线_E840>
        </xsl:if>
      
    </图表:数据系列_E74F>
    </xsl:if>
  </xsl:template>
  <!--计算值-->
  <xsl:template name="caculateValue">
    <xsl:param name="gValue"/>
    <xsl:variable name="name">
      <xsl:choose>
        <xsl:when test="contains($gValue,'=')">
          <xsl:value-of select="substring-after($gValue,'=')"/>
        </xsl:when>
        
        <!--2014-4-17, add by Qihy, 修复bug3079，取值不正确， start-->
        <xsl:when test="contains($gValue,'(')">
		
          <!--zl 2015-5-5-->
          <!--<xsl:value-of select="substring-before(substring-after($gValue,'('),')')"/>-->
          <xsl:value-of select="substring-before($gValue,'!')"/>
          <!--zl 2015-5-5-->
        
		</xsl:when>
        <!--2014-4-17, end-->
                <xsl:otherwise>
          <xsl:value-of select="$gValue"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
	
    <xsl:choose>
      <xsl:when test="contains($name,',')">
        <xsl:value-of select="concat('=',$name)"/>
      </xsl:when>
	  
	  <!--zl 2015-5-5-->
      <xsl:when test="contains($name, ')')">
        <xsl:variable name="sheet">
          <xsl:value-of select="concat('=',$name)"/>
        </xsl:variable>
        <xsl:variable name="range">
          <xsl:value-of select="substring-after($gValue,'!')"/>
        </xsl:variable>
        <xsl:variable name="row_column">
          <xsl:value-of select="translate($range,'$','')"/>
        </xsl:variable>
        <xsl:variable name="new1">
          <xsl:value-of select="concat($sheet,'!')"/>
        </xsl:variable>
        <xsl:variable name="new2">
          <xsl:value-of select="concat($new1,$row_column)"/>
        </xsl:variable>
        <xsl:value-of select="$new2"/>        
      </xsl:when>
      <!--zl 2015-5-5-->
	  
      <xsl:otherwise>
        <xsl:variable name="n2">
          <xsl:value-of select="substring-before($name,'!')"/>
        </xsl:variable>
        <xsl:variable name="n3">
          <xsl:value-of select="concat('=','',$n2,'')"/>
          <!--<xsl:value-of select="concat('=','',$n2,'%')"/>-->
        </xsl:variable>
        <xsl:variable name="n4">
          <xsl:value-of select="substring-after($name,'!')"/>
        </xsl:variable>
        <xsl:variable name="n5">
          <xsl:value-of select="translate($n4,'$','')"/>
        </xsl:variable>
        <xsl:variable name="new">
          <xsl:value-of select="concat($n3,'!',$n5)"/>
        </xsl:variable>
        <xsl:value-of select="$new"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--数据点集-->
  <xsl:template match="c:dPt">
    <xsl:param name="target"/>
    <图表:数据点_E756>
      <xsl:attribute name="点_E757">
        <xsl:value-of select="(./c:idx/@val)+1"/>
      </xsl:attribute>
      <!--边框，填充-->
      <xsl:if test="./c:spPr">
        <xsl:apply-templates select="./c:spPr">
          <xsl:with-param name="target1" select="$target"/>
        </xsl:apply-templates>
      </xsl:if>
      <xsl:for-each select="ancestor::c:ser/c:dLbls/c:spPr">
        <xsl:if test="a:ln and not(a:solidFill) and not(a:blipFill) and not (a:gradFill)">
            <xsl:choose>
              <xsl:when test="(./c:idx/@val)+1 = 1">
                <图表:填充_E746>
                  <图:颜色_8004>#3366ff</图:颜色_8004>
                </图表:填充_E746>
              </xsl:when>
              <xsl:when test="(./c:idx/@val)+1 = 2">
                <图表:填充_E746>
                  <图:颜色_8004>#b93019</图:颜色_8004>
                </图表:填充_E746>
              </xsl:when>
              <xsl:when test="(./c:idx/@val)+1 = 3">
                <图表:填充_E746>
                  <图:颜色_8004>#51941b</图:颜色_8004>
                </图表:填充_E746>
              </xsl:when>
              <xsl:when test="(./c:idx/@val)+1 = 4">
                <图表:填充_E746>
                  <图:颜色_8004>#7c5f9c</图:颜色_8004>
                </图表:填充_E746>
              </xsl:when>
            </xsl:choose>
          <!--图表:填充_E746>
            <图:颜色_8004>auto</图:颜色_8004>
          </图表:填充_E746-->
        </xsl:if>
      </xsl:for-each>
      <xsl:for-each select="ancestor::c:ser/c:dLbls">
        <xsl:if test="(c:spPr/a:solidFill)">
          <表:填充_E746>
            <图:颜色_8004>
              <xsl:value-of select="concat('#',translate(c:spPr/a:solidFill/a:srgbClr/@val,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
            </图:颜色_8004>
          </表:填充_E746>
        </xsl:if>
      </xsl:for-each>
      <!--<xsl:if test="ancestor::c:ser[c:dLbls/c:dLbl]">
        <xsl:for-each select="ancestor::c:ser/c:dLbls/c:dLbl">
          <xsl:variable name="id" select="ancestor::c:ser/c:dLbls/c:dLbl/c:idx/@val"/>
          <xsl:variable name="id2" select="c:idx/@val"/>
          <xsl:if test="$id=$id2">
            <图表:数据标签_E752>
              <xsl:if test="c:showSerName">
                <xsl:attribute name="是否显示系列名_E715">
                  <xsl:value-of select="'true'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="c:showCatName">
                <xsl:attribute name="是否显示类别名_E716">
                  <xsl:value-of select="'true'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="c:showVal">
                <xsl:attribute name="是否显示数值_E717">
                  <xsl:value-of select="'true'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="c:showPercent">
                <xsl:attribute name="是否百分数图表_E718">
                  <xsl:value-of select="'true'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="c:showVertBorders">
                <xsl:attribute name="分隔符_E71A">
                  <xsl:value-of select="'true'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="c:showLegendKey">
                <xsl:attribute name="是否显示图例标志_E719">
                  <xsl:value-of select="'true'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(c:showSerName)">
                <xsl:attribute name="是否显示系列名_E715">
                  <xsl:value-of select="'false'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(c:showCatName)">
                <xsl:attribute name="是否显示类别名_E716">
                  <xsl:value-of select="'false'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(c:showVal)">
                <xsl:attribute name="是否显示数值_E717">
                  <xsl:value-of select="'false'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(c:showPercent)">
                <xsl:attribute name="是否百分数图表_E718">
                  <xsl:value-of select="'false'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(c:showVertBorders)">
                <xsl:attribute name="分隔符_E71A">
                  <xsl:value-of select="'false'"/>
                </xsl:attribute>
              </xsl:if>
              <xsl:if test="not(c:showLegendKey)">
                <xsl:attribute name="是否显示图例标志_E719">
                  <xsl:value-of select="'false'"/>
                </xsl:attribute>
              </xsl:if>
            </图表:数据标签_E752>
          </xsl:if>
        </xsl:for-each>
      </xsl:if>-->
    </图表:数据点_E756>
  </xsl:template>
  <xsl:template name="chart-sjdj">
    <xsl:param name="tar"/>
    <xsl:param name="target"/>
    <xsl:param name="rid"/>
    
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:plotArea]">
      <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:barChart/c:ser/c:dPt or ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:barChart/c:ser/c:dLbls">
        <图表:数据点集_E755>
          <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:barChart/c:ser[c:dLbls]">
            <xsl:variable name="pos">
              <xsl:number count="c:ser" level="any" format="1"/>
            </xsl:variable>
            <xsl:for-each select="c:dLbls/c:dLbl">
              <图表:数据点_E756>
                <xsl:attribute name="点_E757">
                  <xsl:value-of select="(c:idx/@val)+1"/>
                </xsl:attribute>
                <!--边框，填充-->
                <xsl:if test="c:spPr">
                  <xsl:apply-templates select="c:spPr">
                    <xsl:with-param name="target1" select="$target"/>
                  </xsl:apply-templates>
                </xsl:if>
                <xsl:for-each select="ancestor::c:ser/c:dLbls/c:spPr">
                  <xsl:if test="a:ln and not(a:solidFill) and not(a:blipFill) and not (a:gradFill)">
                    <图表:填充_E746>
                      <图:颜色_8004>auto</图:颜色_8004>
                    </图表:填充_E746>
                  </xsl:if>
                </xsl:for-each>
                <xsl:for-each select="ancestor::c:ser/c:dLbls">
                  <xsl:if test="(c:spPr/a:solidFill)">
                    <表:填充_E746>
                      <图:颜色_8004>
                        <xsl:value-of select="concat('#',translate(c:spPr/a:solidFill/a:srgbClr/@val,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
                      </图:颜色_8004>
                    </表:填充_E746>
                  </xsl:if>
                </xsl:for-each>
                <xsl:if test="ancestor::c:ser[c:dLbls/c:dLbl]">
                  <xsl:for-each select="ancestor::c:ser/c:dLbls/c:dLbl">
                    <xsl:variable name="id" select="ancestor::c:ser/c:dLbls/c:dLbl/c:idx/@val"/>
                    <xsl:variable name="id2" select="c:idx/@val"/>
                    <xsl:if test="$id=$id2">
                      <图表:数据标签_E752>
                        <xsl:if test="c:showSerName">
                          <xsl:attribute name="是否显示系列名_E715">
                            <xsl:value-of select="'true'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="c:showCatName">
                          <xsl:attribute name="是否显示类别名_E716">
                            <xsl:value-of select="'true'"/>
                          </xsl:attribute>
                        </xsl:if>
                        
                        <!--2014-6-8, update by lingfeng, 数据标签丢失，start-->
                        <xsl:if test="c:showVal/@val='1'">
                          <xsl:attribute name="是否显示数值_E717">
                            <xsl:value-of select="'true'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="c:showPercent">
                          <xsl:attribute name="是否百分数图表_E718">
                            <xsl:value-of select="'true'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="c:showVertBorders">
                          <xsl:attribute name="分隔符_E71A">
                            <xsl:value-of select="'true'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="c:showLegendKey">
                          <xsl:attribute name="是否显示图例标志_E719">
                            <xsl:value-of select="'true'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(c:showSerName)">
                          <xsl:attribute name="是否显示系列名_E715">
                            <xsl:value-of select="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(c:showCatName)">
                          <xsl:attribute name="是否显示类别名_E716">
                            <xsl:value-of select="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(c:showVal)">
                          <xsl:attribute name="是否显示数值_E717">
                            <xsl:value-of select="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(c:showPercent)">
                          <xsl:attribute name="是否百分数图表_E718">
                            <xsl:value-of select="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(c:showVertBorders)">
                          <xsl:attribute name="分隔符_E71A">
                            <xsl:value-of select="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(c:showLegendKey)">
                          <xsl:attribute name="是否显示图例标志_E719">
                            <xsl:value-of select="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="(c:showVal/@val='0' or not(c:showVal)) and c:showPercent/@val='1'">
                          <图表:数值_E753 是否链接到源_E73E="false" 格式码_E73F="0%" 分类名称_E740="percentage"/>
                        </xsl:if>
                         <!--2014-6-8 end-->
                      </图表:数据标签_E752>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:if>
                <!--xsl:if test="c:tx">
                  <xsl:for-each select="c:tx">
                    <xsl:if test="child::a:r">
                      <xsl:if test="child::a:r/a:rPr[@lang='zh-CN']">
                        <表:系列名 uof:locID="s0089">
                          <xsl:value-of select="child::a:r/a:rPr/a:t"/>
                        </表:系列名>
                      </xsl:if>
                      <xsl:if test="child::a:r/a:rPr[@lang='en-US'] and child::a:r/a:rPr[not(@lang='zh-CN')]">
                        <xsl:variable name="t" select="child::a:r/a:rPr/a:t"/>
                        <xsl:variable name="tt">
                          <xsl:value-of select="substring($t,'123456789','')"/>
                        </xsl:variable>
                        <xsl:variable name="ttt">
                          <xsl:value-of select="substring($tt,',:;','')"/>
                        </xsl:variable>
                        <表:系列名 uof:locID="s0089">
                          <xsl:value-of select="$ttt"/>
                        </表:系列名>
                      </xsl:if>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:if-->
              </图表:数据点_E756>
            </xsl:for-each>
          </xsl:for-each>
          <!--yx，添加，注意数据标签的格式永中保存有问题11.18-->
          <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/node()/c:ser[c:dPt]">
            <xsl:variable name="pos">
              <xsl:number count="c:ser"/>
            </xsl:variable>
            <xsl:for-each select="c:dPt">
              <图表:数据点_E756>
                <!--<xsl:attribute name="系列">
                  <xsl:value-of select="$pos"/>
                </xsl:attribute>-->
                <xsl:attribute name="点_E757">
                  <xsl:value-of select="(c:idx/@val) + 1"/>
                </xsl:attribute>
                <xsl:if test="c:spPr">
                  <!--边框+填充-->
                  <xsl:apply-templates select="c:spPr">
                    <xsl:with-param name="target1" select="$target"/>
                    <!--hwj-->
                  </xsl:apply-templates>
                </xsl:if>
                <!--只有边框不存在填充的情况-->
                <xsl:for-each select="ancestor::c:ser/c:dPt/c:spPr">
                  <xsl:if test="a:ln and not(a:solidFill) and not(a:blipFill) and not (a:gradFill)">
                    <图表:填充_E746>
                      <图:颜色_8004>auto</图:颜色_8004>
                    </图表:填充_E746>
                  </xsl:if>
                </xsl:for-each>
                <xsl:if test="ancestor::c:ser[c:dLbls/c:dLbl]">
                  <xsl:for-each select="ancestor::c:ser/c:dLbls/c:dLbl">
                    <xsl:variable name="id" select="ancestor::c:ser/c:dLbls/c:dLbl/c:idx/@val"/>
                    <xsl:variable name="id2" select="c:idx/@val"/>
                    <xsl:if test="$id=$id2">
                      <图表:数据标签_E752>
                        <xsl:if test="c:showSerName">
                          <xsl:attribute name="是否显示系列名_E715">
                            <xsl:value-of select="'true'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="c:showCatName">
                          <xsl:attribute name="是否显示类别名_E716">
                            <xsl:value-of select="'true'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="c:showVal/@val='1'">
                          <xsl:attribute name="是否显示数值_E717">
                            <xsl:value-of select="'true'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="c:showPercent">
                          <xsl:attribute name="是否百分数图表_E718">
                            <xsl:value-of select="'true'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="c:showVertBorders">
                          <xsl:attribute name="分隔符_E71A">
                            <xsl:value-of select="'true'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="c:showLegendKey">
                          <xsl:attribute name="是否显示图例标志_E719">
                            <xsl:value-of select="'true'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(c:showSerName)">
                          <xsl:attribute name="是否显示系列名_E715">
                            <xsl:value-of select="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(c:showCatName)">
                          <xsl:attribute name="是否显示类别名_E716">
                            <xsl:value-of select="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(c:showVal)">
                          <xsl:attribute name="是否显示数值_E717">
                            <xsl:value-of select="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(c:showPercent)">
                          <xsl:attribute name="是否百分数图表_E718">
                            <xsl:value-of select="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(c:showVertBorders)">
                          <xsl:attribute name="分隔符_E71A">
                            <xsl:value-of select="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(c:showLegendKey)">
                          <xsl:attribute name="是否显示图例标志_E719">
                            <xsl:value-of select="'false'"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="(c:showVal/@val='0' or not(c:showVal)) and c:showPercent/@val='1'">
                          <图表:数值_E753 是否链接到源_E73E="false" 格式码_E73F="0%" 分类名称_E740="percentage"/>
                        </xsl:if>
                      </图表:数据标签_E752>

                    </xsl:if>
                  </xsl:for-each>
                </xsl:if>
                <!--<xsl:if test="c:tx">
                  <xsl:for-each select="c:tx">
                    <xsl:if test="child::a:r">
                      <xsl:if test="child::a:r/a:rPr[@lang='zh-CN']">
                        <表:系列名 uof:locID="s0089">
                          <xsl:value-of select="child::a:r/a:rPr/a:t"/>
                        </表:系列名>
                      </xsl:if>
                      <xsl:if test="child::a:r/a:rPr[@lang='en-US'] and child::a:r/a:rPr[not(@lang='zh-CN')]">
                        <xsl:variable name="t" select="child::a:r/a:rPr/a:t"/>
                        <xsl:variable name="tt">
                          <xsl:value-of select="substring($t,'123456789','')"/>
                        </xsl:variable>
                        <xsl:variable name="ttt">
                          <xsl:value-of select="substring($tt,',:;','')"/>
                        </xsl:variable>
                        <表:系列名 uof:locID="s0089">
                          <xsl:value-of select="$ttt"/>
                        </表:系列名>
                      </xsl:if>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:if>-->
              </图表:数据点_E756>
            </xsl:for-each>
          </xsl:for-each>
        </图表:数据点集_E755>
      </xsl:if>
    </xsl:if>
  </xsl:template>
  <xsl:template name="DisplayLabel">
    <图表:数据标签_E752>
      <xsl:if test="./c:showSerName[@val='1']">
        <xsl:attribute name="是否显示系列名_E715">
          <xsl:value-of select="'true'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./c:showCatName[@val='1']">
        <xsl:attribute name="是否显示类别名_E716">
          <xsl:value-of select="'true'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./c:showVal[@val='1']">
        <xsl:attribute name="是否显示数值_E717">
          <xsl:value-of select="'true'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./c:showPercent[@val='1']">
        <xsl:attribute name="是否百分数图表_E718">
          <xsl:value-of select="'true'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./c:showLegendKey[@val='1']">
        <xsl:attribute name="是否显示图例标志_E719">
          <xsl:value-of select="'true'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./c:showVertBorders[@val='1']">
        <xsl:attribute name="图表:分隔符_E71A">
          <xsl:value-of select="'true'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="not(c:showSerName) or c:showSerName[@val='0']">
        <xsl:attribute name="是否显示系列名_E715">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="not(c:showCatName) or c:showCatName[@val='0']">
        <xsl:attribute name="是否显示类别名_E716">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="not(c:showVal) or c:showVal[@val='0']">
        <xsl:attribute name="是否显示数值_E717">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="not(c:showPercent)  or c:showPercent[@val='0']">
        <xsl:attribute name="是否百分数图表_E718">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="not(c:showLegendKey)">
        <xsl:attribute name="是否显示图例标志_E719">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="c:showLegendKey[@val='0']">
        <xsl:attribute name="是否显示图例标志_E719">
          <xsl:value-of select="'true'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="not(c:showVertBorders) or c:showVertBorders[@val='0']">
        <xsl:attribute name="分隔符_E71A">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
      </xsl:if>
      <!-- add by xuzhenwei BUG_2505:数据标签位置不正确，分隔符不正确 2013-01-19 start -->
      <xsl:if test="./c:separator">
        <xsl:attribute name="分隔符_E71A">
          <xsl:variable name="separator" select="./c:separator"/>
          <xsl:choose>
            <xsl:when test="$separator =';'">
              <xsl:value-of select="'2'"/>
            </xsl:when>
            <xsl:when test="$separator =','">
              <xsl:value-of select="'1'"/>
            </xsl:when>
            <xsl:when test="$separator ='.'">
              <xsl:value-of select="'3'"/>
            </xsl:when>
            <xsl:otherwise>
              <!-- 除";"，","，"."以外的按默认转换，即转换成";"-->
              <xsl:value-of select="'2'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      <!-- end -->
      <xsl:if test="(c:showVal/@val='0' or not(c:showVal)) and c:showPercent/@val='1'">
        <图表:数值_E753 是否链接到源_E73E="false" 格式码_E73F="0%" 分类名称_E740="percentage"/>
      </xsl:if>
    </图表:数据标签_E752>
  </xsl:template>
  <!--边框线、填充-->
  <xsl:template match="c:spPr">
    <xsl:param name="target1"/>
    <xsl:param name="chart_type"/>
      <!-- update by 凌峰 BUG_3029：仅带数据标记的散点图转换错误 20140308 start -->
      <xsl:param name="child_type"/>
    <图表:边框线_4226>
        <xsl:if test="not($chart_type='scatter' and $child_type='marker')">
            <!--end-->
            <xsl:for-each select="./a:ln">
                <xsl:variable name="lineType">
                    <xsl:choose>
                        <xsl:when test="./@cmpd">
                            <xsl:value-of select="./@cmpd"/>
                        </xsl:when>

                        <xsl:otherwise>
                            <xsl:value-of select="'sng'"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:variable>
                <xsl:variable name="dashType">
                    <xsl:choose>
                        <xsl:when test="./a:prstDash/@val">
                            <xsl:value-of select="./a:prstDash/@val"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:value-of select="'none'"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:variable>
                <!--线型，虚实-->
                <!-- add by xuzhenwei BUG_2488：股价图 控制是否有线性 2013-01-11 -->
                <xsl:choose>
                    <xsl:when test="$chart_type='c:stockChart'">
                        <xsl:attribute name="线型_C60D">none</xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:call-template name="LineTypeTransfer">
                            <xsl:with-param name="lineType" select="$lineType"/>
                            <xsl:with-param name="dashType" select="$dashType"/>
                        </xsl:call-template>
                        <xsl:choose>
                            <xsl:when test="@w">
                                <xsl:variable name="w">
                                    <xsl:value-of select="@w"/>
                                </xsl:variable>
                                <xsl:attribute name="宽度_C60F">
                                    <!-- 20130416 update by xuzhenwei BUG_2830 互操作 oo-uof（编辑）-oo 024实用文档-损益表(1).xlsx 文档需要修复 start -->
                                    <xsl:choose>
                                        <xsl:when test="number(@w div 12700) &gt; 1">
                                            <xsl:value-of select="number(@w div 12700) - 1"/>
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:value-of select="1 - number(@w div 12700)"/>
                                        </xsl:otherwise>
                                    </xsl:choose>
                                    <!-- end -->
                                </xsl:attribute>
                            </xsl:when>
                            <!-- 默认情况下边框线宽度 -->
                            <xsl:otherwise>
                                <xsl:attribute name="宽度_C60F">
                                    <xsl:value-of select="'2'"/>
                                </xsl:attribute>
                            </xsl:otherwise>
                        </xsl:choose>
                        <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                            <xsl:attribute name="颜色_C611">
                                <xsl:call-template name="color"/>
                            </xsl:attribute>
                        </xsl:if>
                        <!-- add by xuzhenwei BUG_2503:数据表边框渐近线颜色 2013-01-19 start -->
                        <xsl:if test="a:gradFill">
                            <xsl:apply-templates select="a:gradFill"/>
                        </xsl:if>
                        <!-- end -->
                    </xsl:otherwise>
                </xsl:choose>
                <!-- end -->
            </xsl:for-each>
        </xsl:if>
    </图表:边框线_4226>
    <!--填充-->
    <xsl:if test="a:solidFill or a:blipFill or a:gradFill or a:pattFill">
      <图表:填充_E746>
        <xsl:if test="./a:pattFill">
          <图:图案_800A>
            <xsl:if test="a:pattFill/@prst">
              <xsl:variable name="type">
                <xsl:value-of select="a:pattFill/@prst"/>
              </xsl:variable>
              <xsl:attribute name="类型_8008">
                <xsl:call-template name="ttype">
                  <xsl:with-param name="tttype" select="$type"/>
                </xsl:call-template>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="a:pattFill/a:fgClr or a:pattFill/a:bgClr">
              <xsl:for-each select="a:pattFill/a:fgClr">
                <xsl:attribute name="前景色_800B">
                  <xsl:call-template name="pattClr"/>
                </xsl:attribute>
              </xsl:for-each>
              <xsl:for-each select="a:pattFill/a:bgClr">
                <xsl:attribute name="背景色_800C">
                  <xsl:call-template name="pattClr"/>
                </xsl:attribute>
              </xsl:for-each>
            </xsl:if>
          </图:图案_800A>
        </xsl:if>
        <xsl:if test="./a:solidFill">
          <图:颜色_8004>
            <xsl:call-template name="color"/>
          </图:颜色_8004>
        </xsl:if>
        <xsl:if test="./a:blipFill">
          <图:图片_8005 位置_8006="stretch">
            <xsl:attribute name="图形引用_8007">
              <xsl:variable name="rid">
                <xsl:value-of select="a:blipFill/a:blip/@r:embed"/>
              </xsl:variable>
              <xsl:variable name="t">
                <!--chart1.xml-->
                <xsl:value-of select="substring-after($target1,'charts/')"/>
              </xsl:variable>
              <xsl:variable name="tt">
                <xsl:value-of select="concat('xlsx/xl/charts/_rels/',$t,'.rels')"/>
              </xsl:variable>
              <xsl:variable name="pictureobj">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target1,'charts/')]/pr:Relationships/pr:Relationship[@Id=$rid]/@Target"/>
              </xsl:variable>
              <xsl:variable name="pictureobj1">
                <xsl:value-of select="substring-after($pictureobj,'media/')"/>
              </xsl:variable>
              <xsl:value-of select="substring-before($pictureobj1,'.')"/>
            </xsl:attribute>
          </图:图片_8005>
        </xsl:if>
        <xsl:if test="./a:gradFill">
          <xsl:apply-templates select="a:gradFill"/>
        </xsl:if>
      </图表:填充_E746>
    </xsl:if>
  </xsl:template>
  <!--图例-->
  <xsl:template name="chart-tuli">
    <xsl:param name="tar"/>
    <xsl:param name="target"/>
    <xsl:param name="rid"/>
    <xsl:variable name="starget">
      <xsl:value-of select="substring-after($target,'charts/')"/>
    </xsl:variable>
    <xsl:variable name="sstarget">
      <xsl:value-of select="substring-before($starget,'.')"/>
    </xsl:variable>

    <!-- update by 凌峰 BUG_3027：转换后图例消失  20140306 start -->
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:legend]">
      <!--end-->
      
      <图表:图例_E794>
      <!--位置，大小-->
        <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:legendPos/@val='r' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:legend/c:layout/c:manualLayout]">
          <图表:图例位置_E795>bottom</图表:图例位置_E795>
        </xsl:if>
        <xsl:if test="not(ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:legendPos/@val='r' and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:legend/c:layout/c:manualLayout])">

          <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:legend/c:layout/c:manualLayout]">
            <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:layout/c:manualLayout">
              <图表:位置_E70A>
                <xsl:if test="c:x">
                  <xsl:attribute name="x_C606">
                    <xsl:value-of select="c:x/@val"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="c:y">
                  <xsl:attribute name="y_C607">
                    <xsl:value-of select="c:y/@val"/>
                  </xsl:attribute>
                </xsl:if>
              </图表:位置_E70A>
              <图表:大小_E748>
                <xsl:if test="c:w">
                  <xsl:attribute name="宽_C605">
                    <xsl:value-of select="c:w/@val"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="c:h">
                  <xsl:attribute name="长_C604">
                    <xsl:value-of select="c:h/@val"/>
                  </xsl:attribute>
                </xsl:if>
              </图表:大小_E748>
            </xsl:for-each>
          </xsl:if>
          <!--图例位置-->
          <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:legend/c:legendPos]">
            <图表:图例位置_E795>
              <xsl:choose>
                <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:legendPos[@val='b']">
                  <xsl:value-of select="'bottom'"/>
                </xsl:when>
                <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:legendPos[@val='t']">
                  <xsl:value-of select="'top'"/>
                </xsl:when>
                <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:legendPos[@val='l']">
                  <xsl:value-of select="'left'"/>
                </xsl:when>
                <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:legendPos[@val='r']">
                  <xsl:value-of select="'right'"/>
                </xsl:when>
                <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:legendPos[@val='tr']">
                  <xsl:value-of select="'corner'"/>
                </xsl:when>
              </xsl:choose>
            </图表:图例位置_E795>
          </xsl:if>
        </xsl:if>
      <!--边框+填充-->
      <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend[not(c:spPr)]">
        <图表:边框线_4226>
          <xsl:attribute name="线型_C60D">
            <xsl:value-of select="'none'"/>
          </xsl:attribute>
        </图表:边框线_4226>
      </xsl:if>
      <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:legend/c:spPr]">
        <xsl:apply-templates select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:spPr">
          <xsl:with-param name="target1" select="$target"/>
        </xsl:apply-templates>
      </xsl:if>
      <xsl:apply-templates select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:txPr">
        <xsl:with-param name="type" select="'valAx'"/>
        <xsl:with-param name="sstarget" select="$sstarget"/>
      </xsl:apply-templates>
      <!--图例项-->
      <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:legend/c:legendEntry]">
        <!-- update by xuzhenwei BUG_2494:字体颜色，下划线效果丢失 2013-01-19 start -->
          <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend//c:legendEntry/c:txPr">
          <!-- end -->
          <图表:图例项_E765>
            <xsl:attribute name="系列标示符_E79A">
              <xsl:value-of select="preceding-sibling::c:idx/@val"/>
            </xsl:attribute>
            <xsl:for-each select="a:p/a:pPr">
              <图表:字体_E70B>
                <xsl:if test="a:defRPr/a:ea or a:defRPr/a:latin or a:defRPr/@sz or a:defRPr/a:solidFill/a:srgbClr or a:defRPr/a:solidFill/a:schemeClr">
                  <字:字体_4128>
                    <xsl:attribute name="西文字体引用_4129">
                      <xsl:value-of select="concat($sstarget,'.tuli_l_font')"/>
                    </xsl:attribute>
                    <xsl:attribute name="中文字体引用_412A">
                      <xsl:value-of select="concat($sstarget,'.tuli_e_font')"/>
                    </xsl:attribute>
                    <xsl:if test="a:defRPr[@sz]">
                      <xsl:attribute name="字号_412D">
                        <xsl:variable name="sz">
                          <xsl:value-of select="a:defRPr/@sz"/>
                        </xsl:variable>
                        <xsl:value-of select="$sz div 100"/>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:if test="a:defRPr">
                      <xsl:for-each select="a:defRPr">
                        <xsl:if test="a:solidFill">
                          <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                            <xsl:attribute name="颜色_412F">
                              <xsl:call-template name="color"/>
                            </xsl:attribute>
                          </xsl:if>
                        </xsl:if>
                      </xsl:for-each>
                    </xsl:if>
                  </字:字体_4128>
                </xsl:if>
                <xsl:call-template name="table-tbziti"/>
              </图表:字体_E70B>
            </xsl:for-each>
            <!--/xsl:if-->
          </图表:图例项_E765>
        </xsl:for-each>
      </xsl:if>
    </图表:图例_E794>
    </xsl:if>
      </xsl:template>
  <!--数据表-->
  <xsl:template name="chart-sjb">
    <xsl:param name="tar"/>
    <xsl:param name="rid"/>
    <xsl:param name="target"/>
    <xsl:variable name="starget">
      <xsl:value-of select="substring-after($target,'charts/')"/>
    </xsl:variable>
    <xsl:variable name="sstarget">
      <xsl:value-of select="substring-before($starget,'.')"/>
    </xsl:variable>
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:plotArea/c:dTable]">
      <图表:数据表_E79B>
        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:dTable">
          <xsl:attribute name="是否显示图例_E79C">
            <xsl:choose>
              <xsl:when test="./c:showKeys/@val='1'">
                <xsl:value-of select="'true'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'false'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <!-- update by xuzhenwei BUG_2503:数据表-边框颜色 2013-01-23 start -->
          <xsl:attribute name="是否显示外边框_E79D">
            <xsl:choose>
              <xsl:when test="./c:showOutline/@val='1'">
                <xsl:value-of select="'true'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'false'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name="是否垂直显示_E79E">
            <xsl:choose>
              <xsl:when test="./c:showVertBorder/@val='1'">
                <xsl:value-of select="'true'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'false'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name="是否水平显示_E79F">
            <xsl:choose>
              <xsl:when test="./c:showHorzBorder/@val='1'">
                <xsl:value-of select="'true'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'false'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <!-- end -->
          <!--边框线，填充-->
            <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:dTable/c:spPr]">
              <xsl:apply-templates select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:dTable/c:spPr">
                <xsl:with-param name="target1" select="$target"/>
              </xsl:apply-templates>
            </xsl:if>
          <!--字体-->
          <xsl:for-each select="c:txPr/a:p/a:pPr">
            <图表:字体_E70B>
              <字:字体_4128>
                <xsl:attribute name="西文字体引用_4129">
                  <xsl:value-of select="concat($sstarget,'.sjb_l_font')"/>
                </xsl:attribute>
                <xsl:attribute name="中文字体引用_412A">
                  <xsl:value-of select="concat($sstarget,'.sjb_e_font')"/>
                </xsl:attribute>
                <xsl:if test="a:defRPr/@sz">
                  <xsl:attribute name="字号_412D">
                    <xsl:variable name="zh">
                      <xsl:value-of select="a:defRPr/@sz"/>
                    </xsl:variable>
                    <xsl:value-of select="$zh div 100"/>
                  </xsl:attribute>
                </xsl:if>
                <xsl:if test="a:defRPr">
                  <xsl:for-each select="a:defRPr">
                    <xsl:if test="a:solidFill">
                      <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                        <xsl:attribute name="颜色_412F">
                          <xsl:call-template name="color"/>
                        </xsl:attribute>
                      </xsl:if>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:if>
              </字:字体_4128>
              <xsl:call-template name="table-tbziti"/>
            </图表:字体_E70B>
          </xsl:for-each>
          <xsl:if test="not(c:tx/c:rich/a:p) and not(c:txPr)">
            <xsl:call-template name="autofont2"/>
          </xsl:if>
        </xsl:for-each>
      </图表:数据表_E79B>
    </xsl:if>
  </xsl:template>
  <xsl:template name="autofont2">
    <图表:字体_E70B>
      <字:字体_4128>
        <xsl:attribute name="字号_412D">
          <xsl:value-of select="'10'"/>
        </xsl:attribute>
      </字:字体_4128>
    </图表:字体_E70B>
  </xsl:template>
  <!--图表标题-->
  <xsl:template name="table-btj-ziti">
    <xsl:choose>
      <xsl:when test="a:rPr[@cap='all']">
        <字:醒目字体类型_4141>uppercase</字:醒目字体类型_4141>
      </xsl:when>
      <xsl:when test="not(a:rPr[@cap='all']) and ./preceding-sibling::a:pPr/a:defRPr[@cap='all']">
        <字:醒目字体类型_4141>uppercase</字:醒目字体类型_4141>
      </xsl:when>
      <xsl:when test="not(a:rPr[@cap='all']) and not(./preceding-sibling::a:pPr/a:defRPr[@cap='all']) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@cap='all']">
        <字:醒目字体类型_4141>uppercase</字:醒目字体类型_4141>
      </xsl:when>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="a:rPr[@cap='small']">
        <字:醒目字体类型_4141>small-caps</字:醒目字体类型_4141>
      </xsl:when>
      <xsl:when test="not(a:rPr[@cap='small']) and ./preceding-sibling::a:pPr/a:defRPr[@cap='small']">
        <字:醒目字体类型_4141>small-caps</字:醒目字体类型_4141>
      </xsl:when>
      <xsl:when test="not(a:rPr[@cap='small']) and ./preceding-sibling::a:pPr/a:defRPr[@cap='small'] and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@cap='small']">
        <字:醒目字体类型_4141>small-caps</字:醒目字体类型_4141>
      </xsl:when>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="a:rPr[@b='0'] or  ./preceding-sibling::a:pPr/a:defRPr[@b='0'] or ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@b='0']">
        <字:是否粗体_4130>false</字:是否粗体_4130>
      </xsl:when>
      <xsl:otherwise>
        <字:是否粗体_4130>true</字:是否粗体_4130>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="a:rPr[@i='1']">
        <字:是否斜体_4131>true</字:是否斜体_4131>
      </xsl:when>
      <xsl:when test="not(a:rPr[@i='1']) and ./preceding-sibling::a:pPr/a:defRPr[@i='1']">
        <字:是否斜体_4131>true</字:是否斜体_4131>
      </xsl:when>
      <xsl:when test="not(a:rPr[@i='1']) and not(./preceding-sibling::a:pPr/a:defRPr[@i='1']) and  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@i='1']">
        <字:是否斜体_4131>true</字:是否斜体_4131>
      </xsl:when>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="a:rPr[@strike]">
        <字:删除线_4135>
          <xsl:choose>
            <xsl:when test="a:rPr[@strike='dblStrike']">

              <xsl:value-of select="'double'"/>
            </xsl:when>
            <xsl:when test="a:rPr[@strike='noStrike']">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test="a:rPr[@strike='sngStrike']">
              <xsl:value-of select="'single'"/>
            </xsl:when>
          </xsl:choose>
        </字:删除线_4135>
      </xsl:when>
      <xsl:when test="not(a:rPr[@strike]) and  ./preceding-sibling::a:pPr/a:defRPr[@strike]">
        <字:删除线_4135>
          <xsl:choose>
            <xsl:when test=" ./preceding-sibling::a:pPr/a:defRPr[@strike='dblStrike']">
              <xsl:value-of select="'double'"/>
            </xsl:when>
            <xsl:when test=" ./preceding-sibling::a:pPr/a:defRPr[@strike='noStrike']">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test=" ./preceding-sibling::a:pPr/a:defRPr[@strike='sngStrike']">
              <xsl:value-of select="'single'"/>
            </xsl:when>
          </xsl:choose>
        </字:删除线_4135>
      </xsl:when>
      <xsl:when test="not(a:rPr[@strike]) and  not(./preceding-sibling::a:pPr/a:defRPr[@strike]) and  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@strike]">
        <字:删除线_4135>
          <xsl:choose>
            <xsl:when test="  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@strike='dblStrike']">
              <xsl:value-of select="'double'"/>
            </xsl:when>
            <xsl:when test="  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@strike='noStrike']">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test=" ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@strike='sngStrike']">
              <xsl:value-of select="'single'"/>
            </xsl:when>
          </xsl:choose>
        </字:删除线_4135>
      </xsl:when>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="a:rPr[@u]">
        <字:下划线_4136>
          <xsl:attribute name="线型_4137">
            <xsl:choose>
              <xsl:when test="a:rPr[@u='dash']">
                <xsl:value-of select="'dash'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='dashHeavy']">
                <xsl:value-of select="'dashed-heavy'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='dashLong']">
                <xsl:value-of select="'dash-long'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='dashLongHeavy']">
                <xsl:value-of select="'dash-long-heavy'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='dbl']">
                <xsl:value-of select="'double'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='dotDash']">
                <xsl:value-of select="'dot-dash'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='dotDashHeavy']">
                <xsl:value-of select="'dash-dot-heavy'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='dotDotDash']">
                <xsl:value-of select="'dot-dot-dash'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='dotDotDashHeavy']">
                <xsl:value-of select="'dash-dot-dot-heavy'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='dotted']">
                <xsl:value-of select="'dotted'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='dottedHeavy']">
                <xsl:value-of select="'dotted-heavy'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='heavy']">
                <xsl:value-of select="'thick'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='none']">
                <xsl:value-of select="'none'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='sng']">
                <xsl:value-of select="'single'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='wavy']">
                <xsl:value-of select="'wave'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='wavyDbl']">
                <xsl:value-of select="'wavy-double'"/>
              </xsl:when>
              <xsl:when test="a:rPr[@u='wavyHeavy']">
                <xsl:value-of select="'wavy-heavy'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'single'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:choose>
            <xsl:when test="a:rPr[a:uFill]">
              <xsl:for-each select="a:rPr/a:uFill">
                <xsl:if test="a:solidFill">
                  <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                    <xsl:attribute name="颜色_412F">
                      <xsl:call-template name="color"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:if>
              </xsl:for-each>
            </xsl:when>
            <xsl:when test="not(a:rPr[a:uFill]) and ./preceding-sibling::a:pPr/a:defRPr[a:uFill]">
              <xsl:for-each select="./preceding-sibling::a:pPr/a:defRPr/a:uFill">
                <xsl:if test="a:solidFill">
                  <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                    <xsl:attribute name="颜色_412F">
                      <xsl:call-template name="color"/>
                    </xsl:attribute>
                  </xsl:if>
                </xsl:if>
              </xsl:for-each>
            </xsl:when>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="a:rPr[@u='words']">
              <xsl:attribute name="是否字下划线_4139">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:rPr[@u='words']) and ./preceding-sibling::a:pPr/a:defRPr[@u='words']">
              <xsl:attribute name="是否字下划线_4139">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </字:下划线_4136>
      </xsl:when>
      <xsl:when test="not(a:rPr[@u]) and ./preceding-sibling::a:pPr/a:defRPr[@u]">
        <字:下划线_4136>
          <xsl:attribute name="线型_4137">
            <xsl:choose>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='dash']">
                <xsl:value-of select="'dash'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='dashHeavy']">
                <xsl:value-of select="'dashed-heavy'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='dashLong']">
                <xsl:value-of select="'dash-long'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='dashLongHeavy']">
                <xsl:value-of select="'dash-long-heavy'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='dbl']">
                <xsl:value-of select="'double'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='dotDash']">
                <xsl:value-of select="'dot-dash'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='dotDashHeavy']">
                <xsl:value-of select="'dash-dot-heavy'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='dotDotDash']">
                <xsl:value-of select="'dot-dot-dash'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='dotDotDashHeavy']">
                <xsl:value-of select="'dash-dot-dot-heavy'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='dotted']">
                <xsl:value-of select="'dotted'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='dottedHeavy']">
                <xsl:value-of select="'dotted-heavy'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='heavy']">
                <xsl:value-of select="'thick'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='none']">
                <xsl:value-of select="'none'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='sng']">
                <xsl:value-of select="'single'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='wavy']">
                <xsl:value-of select="'wave'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='wavyDbl']">
                <xsl:value-of select="'wavy-double'"/>
              </xsl:when>
              <xsl:when test="./preceding-sibling::a:pPr/a:defRPr[@u='wavyHeavy']">
                <xsl:value-of select="'wavy-heavy'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'single'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:if test=" ./preceding-sibling::a:pPr/a:defRPr[a:uFill]">
            <xsl:for-each select=" ./preceding-sibling::a:pPr/a:defRPr/a:uFill">
              <xsl:if test="a:solidFill">
                <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                  <xsl:attribute name="颜色_412F">
                    <xsl:call-template name="color"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
          <xsl:if test=" ./preceding-sibling::a:pPr/a:defRPr[@u='words']">
            <xsl:attribute name="是否字下划线_4139">
              <xsl:value-of select="'true'"/>
            </xsl:attribute>
          </xsl:if>
        </字:下划线_4136>
      </xsl:when>
      <xsl:when test="not(a:rPr[@u]) and not(./preceding-sibling::a:pPr/a:defRPr[@u]) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u]">
        <字:下划线_4136>
          <xsl:attribute name="线型_4137">
            <xsl:choose>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='dash']">
                <xsl:value-of select="'dash'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='dashHeavy']">
                <xsl:value-of select="'dashed-heavy'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='dashLong']">
                <xsl:value-of select="'dash-long'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='dashLongHeavy']">
                <xsl:value-of select="'dash-long-heavy'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='dbl']">
                <xsl:value-of select="'double'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='dotDash']">
                <xsl:value-of select="'dot-dash'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='dotDashHeavy']">
                <xsl:value-of select="'dash-dot-heavy'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='dotDotDash']">
                <xsl:value-of select="'dot-dot-dash'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='dotDotDashHeavy']">
                <xsl:value-of select="'dash-dot-dot-heavy'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='dotted']">
                <xsl:value-of select="'dotted'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='dottedHeavy']">
                <xsl:value-of select="'dotted-heavy'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='heavy']">
                <xsl:value-of select="'thick'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='none']">
                <xsl:value-of select="'none'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='sng']">
                <xsl:value-of select="'single'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='wavy']">
                <xsl:value-of select="'wave'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='wavyDbl']">
                <xsl:value-of select="'wavy-double'"/>
              </xsl:when>
              <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='wavyHeavy']">
                <xsl:value-of select="'wavy-heavy'"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'single'"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:attribute>
          <xsl:if test=" ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[a:uFill]">
            <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:uFill">
              <xsl:if test="a:solidFill">
                <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                  <xsl:attribute name="颜色_412F">
                    <xsl:call-template name="color"/>
                  </xsl:attribute>
                </xsl:if>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
          <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@u='words']">
            <xsl:attribute name="是否字下划线_4139">
              <xsl:value-of select="'true'"/>
            </xsl:attribute>
          </xsl:if>
        </字:下划线_4136>
      </xsl:when>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="a:rPr[@baseline]">
        <字:上下标类型_4143>
          <xsl:choose>
            <xsl:when test="a:rPr[@baseline = 0]">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test="a:rPr[@baseline &gt;0]">
              <xsl:value-of select="'sup'"/>
            </xsl:when>
            <xsl:when test="a:rPr[@baseline &lt;0]">
              <xsl:value-of select="'sub'"/>
            </xsl:when>
          </xsl:choose>
        </字:上下标类型_4143>
      </xsl:when>
      <xsl:when test="not(a:rPr[@baseline]) and ./preceding-sibling::a:pPr/a:defRPr[@baseline]">
        <字:上下标类型_4143>
          <xsl:choose>
            <xsl:when test=" ./preceding-sibling::a:pPr/a:defRPr[@baseline = 0]">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test=" ./preceding-sibling::a:pPr/a:defRPr[@baseline &gt;0]">
              <xsl:value-of select="'sup'"/>
            </xsl:when>
            <xsl:when test=" ./preceding-sibling::a:pPr/a:defRPr[@baseline &lt;0]">
              <xsl:value-of select="'sub'"/>
            </xsl:when>
          </xsl:choose>
        </字:上下标类型_4143>
      </xsl:when>
      <xsl:when test="not(a:rPr[@baseline]) and not(./preceding-sibling::a:pPr/a:defRPr[@baseline]) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@baseline]">
        <字:上下标类型_4143>
          <xsl:choose>
            <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@baseline = 0]">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@baseline &gt;0]">
              <xsl:value-of select="'sup'"/>
            </xsl:when>
            <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@baseline &lt;0]">
              <xsl:value-of select="'sub'"/>
            </xsl:when>
          </xsl:choose>
        </字:上下标类型_4143>
      </xsl:when>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="a:rPr[@spc]">
        <字:字符间距_4145>
          <xsl:variable name="d">
            <xsl:value-of select="a:rPr/@spc"/>
          </xsl:variable>
          <xsl:value-of select="$d div 100"/>
        </字:字符间距_4145>
      </xsl:when>
      <xsl:when test="not(a:rPr[@spc]) and  ./preceding-sibling::a:pPr/a:defRPr[@spc]">
        <字:字符间距_4145>
          <xsl:variable name="d">
            <xsl:value-of select=" ./preceding-sibling::a:pPr/a:defRPr/@spc"/>
          </xsl:variable>
          <xsl:value-of select="$d div 100"/>
        </字:字符间距_4145>
      </xsl:when>
      <xsl:when test="not(a:rPr[@spc]) and  not(./preceding-sibling::a:pPr/a:defRPr[@spc]) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[@spc]">
        <字:字符间距_4145>
          <xsl:variable name="d">
            <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@spc"/>
          </xsl:variable>
          <xsl:value-of select="$d div 100"/>
        </字:字符间距_4145>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="chart-btj">
    <xsl:param name="tar"/>
    <xsl:param name="rid"/>
    <xsl:param name="target"/>
    <xsl:variable name="starget">
      <xsl:value-of select="substring-after($target,'charts/')"/>
    </xsl:variable>
    <xsl:variable name="sstarget">
      <xsl:value-of select="substring-before($starget,'.')"/>
    </xsl:variable>
    <xsl:variable name="piecharttitle">
      <xsl:choose>
        <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:pie3DChart/c:ser[position()=1]/c:tx/c:strRef/c:strCache/c:pt/c:v">
          <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:pie3DChart/c:ser[position()=1]/c:tx/c:strRef/c:strCache/c:pt/c:v"/>
        </xsl:when>
        <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:pieChart/c:ser[position()=1]/c:tx/c:strRef/c:strCache/c:pt/c:v">
          <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:pieChart/c:ser[position()=1]/c:tx/c:strRef/c:strCache/c:pt/c:v"/>
        </xsl:when>
        <xsl:when test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:ofPieChart/c:ser[position()=1]/c:tx/c:strRef/c:strCache/c:pt/c:v">
          <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:ofPieChart/c:ser[position()=1]/c:tx/c:strRef/c:strCache/c:pt/c:v"/>
        </xsl:when>
        <xsl:otherwise>
          
        </xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:title] or ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:plotArea/c:catAx/c:title] or ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:plotArea/c:valAx/c:title]">
      <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:title]">
        <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title">
          <图表:标题_E736>
            <xsl:choose>

                <!-- update by 凌峰 BUG_3080：坐标轴标题丢失，坐标刻度值变化 20140308 start -->
                <xsl:when test="./c:tx/c:strRef/c:strCache/c:pt/c:v">
                    <xsl:attribute name="名称_E742">
                        <xsl:value-of select="./c:tx/c:strRef/c:strCache/c:pt/c:v"/>
                    </xsl:attribute>
                </xsl:when>
                <!--end-->
                
              <xsl:when test="descendant::a:t">
                <xsl:attribute name="名称_E742">
                  <xsl:variable name="text">
                    <xsl:for-each select="descendant::a:t">
                      <xsl:value-of select="."/>
                    </xsl:for-each>
                  </xsl:variable>
                  <xsl:value-of select="$text"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:variable name="count" select="count(following::c:plotArea[1]/*/c:ser)"/>
                <xsl:attribute name="名称_E742">
                  <xsl:variable name="text">
                    <xsl:choose>
                      <xsl:when test="$piecharttitle!=''">
                        <xsl:value-of select="$piecharttitle"/>
                      </xsl:when>
                      <xsl:when test="$count='1'">
                        <xsl:value-of select="following::c:plotArea[1]/*/c:ser/c:tx/c:strRef/c:strCache/c:pt/c:v"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="'图表标题'"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:variable>
                  <xsl:value-of select="$text"/>
                </xsl:attribute>
              </xsl:otherwise>
            </xsl:choose>
            <!--边框+填充-->
            <xsl:if test="c:spPr">
              <xsl:for-each select="c:spPr">
                <xsl:call-template name="chart-biankuang">
                  <xsl:with-param name="target2" select="$target"/>
                </xsl:call-template>
                <xsl:call-template name="chart-tianchong">
                  <xsl:with-param name="target2" select="$target"/>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:if>
            <!--字体-->
            <xsl:if test="not(c:tx/c:rich/a:p) and not(c:txPr)">
              <图表:字体_E70B>
                <字:字体_4128>
                  <xsl:attribute name="字号_412D">
                      <!-- update by 凌峰 BUG_3154：互操作-ooxml->uof->oo（2010/​2013） ，2013图形丢失，2010需要修复，图形失真 20140420 start -->
                    <xsl:value-of select="'18'"/>
                      <!--end-->
                  </xsl:attribute>
                </字:字体_4128>
                <字:是否粗体_4130>true</字:是否粗体_4130>
              </图表:字体_E70B>
            </xsl:if>
            <xsl:choose>
              <xsl:when test="c:tx/c:rich/a:p[a:r]">
                <xsl:for-each select="c:tx/c:rich/a:p[a:r]">
                  <图表:字体_E70B>                  
                    <xsl:for-each select="a:r">
                      <xsl:choose>
                        <xsl:when test="./a:rPr">
                          <!--yanghaojie
                          <xsl:if test="./a:rPr[@i='1']">
                            <字:是否斜体_4131>true</字:是否斜体_4131>
                          </xsl:if>
                          -->
                          <字:字体_4128>
                            <xsl:if test="./a:rPr/a:latin">
                              <xsl:attribute name="西文字体引用_4129">
                                <xsl:value-of select="concat($sstarget,'.btj_bt_l_font')"/>
                              </xsl:attribute>
                            </xsl:if>
                            <xsl:if test="./a:rPr/a:ea">
                              <xsl:attribute name="中文字体引用_412A">
                                <xsl:value-of select="concat($sstarget,'.btj_bt_e_font')"/>
                              </xsl:attribute>
                            </xsl:if>
                            <!--字号-->
                            <xsl:attribute name="字号_412D">
                              <xsl:choose>
                                <xsl:when test="a:rPr/@sz">
                                  <xsl:variable name="sz">
                                    <xsl:value-of select="a:rPr/@sz"/>
                                  </xsl:variable>
                                  <xsl:value-of select="$sz div 100"/>
                                </xsl:when>
                                <xsl:when test="not(a:rPr/@sz) and ./preceding-sibling::a:pPr/a:defRPr/@sz">
                                  <xsl:variable name="sz">
                                    <xsl:value-of select="./preceding-sibling::a:pPr/a:defRPr/@sz"/>
                                  </xsl:variable>
                                  <xsl:value-of select="$sz div 100"/>
                                </xsl:when>
                                <xsl:when test="not(a:rPr/@sz) and not(./preceding-sibling::a:pPr/a:defRPr/@sz) and  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                                  <xsl:variable name="sz">
                                    <xsl:value-of select=" ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                                  </xsl:variable>
                                  <xsl:value-of select="$sz div 100"/>
                                </xsl:when>
                                <xsl:otherwise>
                                  <xsl:value-of select="'12'"/>
                                </xsl:otherwise>
                              </xsl:choose>
                            </xsl:attribute>
                            <!--字体颜色-->
                            <xsl:choose>
                              <xsl:when test="a:rPr/a:solidFill">
                                <xsl:for-each select="a:rPr">
                                  <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                    <xsl:attribute name="颜色_412F">
                                      <xsl:call-template name="color"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:for-each>
                              </xsl:when>
                              <xsl:when test="not(a:rPr/a:solidFill) and ./preceding-sibling::a:pPr/a:defRPr/a:solidFill">
                                <xsl:for-each select="./preceding-sibling::a:pPr/a:defRPr">
                                  <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                    <xsl:attribute name="颜色_412F">
                                      <xsl:call-template name="color"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:for-each>
                              </xsl:when>
                              <xsl:when test="not(a:rPr/a:solidFill) and not(./preceding-sibling::a:pPr/a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                                <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr">
                                  <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                    <xsl:attribute name="颜色_412F">
                                      <xsl:call-template name="color"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:for-each>
                              </xsl:when>
                            </xsl:choose>
                          </字:字体_4128>
                          
                        </xsl:when>
                        <xsl:when test="not(./a:rPr[a:ea or a:latin]) and ./preceding-sibling::a:pPr/a:defRPr[a:ea or a:latin]">
                          <字:字体_4128>
                            <xsl:attribute name="西文字体引用_4129">
                              <xsl:value-of select="concat($sstarget,'.btj_bt_l_font')"/>
                            </xsl:attribute>
                            <xsl:attribute name="中文字体引用_412A">
                              <xsl:value-of select="concat($sstarget,'.btj_bt_e_font')"/>
                            </xsl:attribute>
                            <xsl:attribute name="字号_412D">
                              <xsl:choose>
                                <xsl:when test="a:rPr/@sz">
                                  <xsl:variable name="sz">
                                    <xsl:value-of select="a:rPr/@sz"/>
                                  </xsl:variable>
                                  <xsl:value-of select="$sz div 100"/>
                                </xsl:when>
                                <xsl:when test="not(a:rPr/@sz) and ./preceding-sibling::a:pPr/a:defRPr/@sz">
                                  <xsl:variable name="sz">
                                    <xsl:value-of select="./preceding-sibling::a:pPr/a:defRPr/@sz"/>
                                  </xsl:variable>
                                  <xsl:value-of select="$sz div 100"/>
                                </xsl:when>
                                <xsl:when test="not(a:rPr/@sz) and not(./preceding-sibling::a:pPr/a:defRPr/@sz) and  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                                  <xsl:variable name="sz">
                                    <xsl:value-of select=" ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                                  </xsl:variable>
                                  <xsl:value-of select="$sz div 100"/>
                                </xsl:when>
                                <xsl:otherwise>
                                  <xsl:value-of select="'18'"/>
                                </xsl:otherwise>
                              </xsl:choose>
                            </xsl:attribute>
                            <xsl:choose>
                              <xsl:when test="a:rPr/a:solidFill">
                                <xsl:for-each select="a:rPr">
                                  <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                    <xsl:attribute name="颜色_412F">
                                      <xsl:call-template name="color"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:for-each>
                              </xsl:when>
                              <xsl:when test="not(a:rPr/a:solidFill) and ./preceding-sibling::a:pPr/a:defRPr/a:solidFill">
                                <xsl:for-each select="./preceding-sibling::a:pPr/a:defRPr">
                                  <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                    <xsl:attribute name="颜色_412F">
                                      <xsl:call-template name="color"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:for-each>
                              </xsl:when>
                              <xsl:when test="not(a:rPr/a:solidFill) and not(./preceding-sibling::a:pPr/a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                                <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr">
                                  <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                    <xsl:attribute name="颜色_412F">
                                      <xsl:call-template name="color"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:for-each>
                              </xsl:when>
                            </xsl:choose>
                          </字:字体_4128>
                        </xsl:when>
                        <xsl:when test="not(./a:rPr[a:ea or a:latin]) and not (./preceding-sibling::a:pPr/a:defRPr[a:ea or a:latin]) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[a:ea or a:latin]">
                          <字:字体_4128>
                            <xsl:attribute name="西文字体引用_4129">
                              <xsl:value-of select="concat($sstarget,'.tbq_l_font')"/>
                            </xsl:attribute>
                            <xsl:attribute name="中文字体引用_412A">
                              <xsl:value-of select="concat($sstarget,'.tbq_e_font')"/>
                            </xsl:attribute>
                            <xsl:attribute name="字号_412D">
                              <xsl:choose>
                                <xsl:when test="a:rPr/@sz">
                                  <xsl:variable name="sz">
                                    <xsl:value-of select="a:rPr/@sz"/>
                                  </xsl:variable>
                                  <xsl:value-of select="$sz div 100"/>
                                </xsl:when>
                                <xsl:when test="not(a:rPr/@sz) and ./preceding-sibling::a:pPr/a:defRPr/@sz">
                                  <xsl:variable name="sz">
                                    <xsl:value-of select="./preceding-sibling::a:pPr/a:defRPr/@sz"/>
                                  </xsl:variable>
                                  <xsl:value-of select="$sz div 100"/>
                                </xsl:when>
                                <xsl:when test="not(a:rPr/@sz) and not(./preceding-sibling::a:pPr/a:defRPr/@sz) and  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                                  <xsl:variable name="sz">
                                    <xsl:value-of select=" ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                                  </xsl:variable>
                                  <xsl:value-of select="$sz div 100"/>
                                </xsl:when>
                                <xsl:otherwise>
                                  <xsl:value-of select="'18'"/>
                                </xsl:otherwise>
                              </xsl:choose>
                            </xsl:attribute>
                            <xsl:choose>
                              <xsl:when test="a:rPr/a:solidFill">
                                <xsl:for-each select="a:rPr">
                                  <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                    <xsl:attribute name="颜色_412F">
                                      <xsl:call-template name="color"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:for-each>
                              </xsl:when>
                              <xsl:when test="not(a:rPr/a:solidFill) and ./preceding-sibling::a:pPr/a:defRPr/a:solidFill">
                                <xsl:for-each select="./preceding-sibling::a:pPr/a:defRPr">
                                  <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                    <xsl:attribute name="颜色_412F">
                                      <xsl:call-template name="color"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:for-each>
                              </xsl:when>
                              <xsl:when test="not(a:rPr/a:solidFill) and not(./preceding-sibling::a:pPr/a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                                <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr">
                                  <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                    <xsl:attribute name="颜色_412F">
                                      <xsl:call-template name="color"/>
                                    </xsl:attribute>
                                  </xsl:if>
                                </xsl:for-each>
                              </xsl:when>
                            </xsl:choose>
                          </字:字体_4128>
                        </xsl:when>
                      </xsl:choose>
                      <xsl:if test="./a:rPr">
                        <xsl:call-template name="table-btj-ziti"/>
                      </xsl:if>
                    </xsl:for-each>
                    
                    
                    
                  </图表:字体_E70B>
                  <xsl:for-each select="ancestor::c:chart/c:title/c:tx/c:rich">
                    <xsl:if test="a:bodyPr">
                      <xsl:call-template name="table-btj-dq"/>
                    </xsl:if>
                  </xsl:for-each>
                </xsl:for-each>
              </xsl:when>
              <xsl:when test="c:tx/c:rich/a:p[not(a:r)] and c:tx/c:rich/a:p/a:pPr/a:defRPr">
                <xsl:for-each select="c:tx/c:rich/a:p/a:pPr/a:defRPr">
                  <图表:字体_E70B>
                    <xsl:if test="a:ea or a:latin">
                      <字:字体_4128>
                        <xsl:attribute name="西文字体引用_4129">
                          <xsl:value-of select="concat($sstarget,'.btj_bt_l_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="中文字体引用_412A">
                          <xsl:value-of select="concat($sstarget,'.btj_bt_e_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="字号_412D">
                          <xsl:choose>
                            <xsl:when test="@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(@sz) and  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select=" ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'18'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <xsl:choose>
                          <xsl:when test="a:solidFill">
                            <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                              <xsl:attribute name="颜色_412F">
                                <xsl:call-template name="color"/>
                              </xsl:attribute>
                            </xsl:if>
                          </xsl:when>
                          <xsl:when test="not(a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                        </xsl:choose>
                      </字:字体_4128>
                    </xsl:if>
                    <xsl:for-each select="..">
                      <xsl:call-template name="table-tbziti"/>
                    </xsl:for-each>
                  </图表:字体_E70B>
                </xsl:for-each>
              </xsl:when>
            </xsl:choose>
            <xsl:for-each select="c:txPr">
              <xsl:for-each select="a:p/a:pPr">
                <xsl:choose>
                  <xsl:when test="a:defRPr[a:ea or a:latin]">
                    <图表:字体_E70B>
                      <字:字体_4128>
                        
                        <!--2014-4-23, update by Qihy,将属性前缀"字:"去掉, start-->
                        <xsl:attribute name="西文字体引用_4129">
                          <xsl:value-of select="concat($sstarget,'.btj_bt_l_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="中文字体引用_412A">
                          <xsl:value-of select="concat($sstarget,'.btj_bt_e_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="字号_412D">
                          <xsl:choose>
                            <xsl:when test="a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:defRPr/@sz) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'18'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <xsl:choose>
                          <xsl:when test="a:defRPr/a:solidFill/a:schemeClr or a:defRPr/a:solidFill/a:srgbClr">
                            <xsl:attribute name="颜色_412F">
                              <xsl:call-template name="color_ziti"/>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:when test="not(a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr">
                              <xsl:attribute name="颜色_412F">
                                <xsl:call-template name="color_ziti"/>
                              </xsl:attribute>
                            </xsl:for-each>
                          </xsl:when>
                        </xsl:choose>
                      </字:字体_4128>
                      <xsl:call-template name="table-tbziti"/>
                    </图表:字体_E70B>
                  </xsl:when>
                  <xsl:when test="not(a:defRPr[a:ea or a:latin]) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[a:ea or a:latin]">
                    <图表:字体_E70B>
                      <字:字体_4128 >
                        <xsl:attribute name="西文字体引用_4129">
                          <xsl:value-of select="concat($sstarget,'.tbq_l_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="中文字体引用_412A">
                          <xsl:value-of select="concat($sstarget,'.tbq_e_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="字号_412D">
                          <xsl:choose>
                            <xsl:when test="a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:defRPr/@sz) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'18'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <xsl:choose>
                          <xsl:when test="a:defRPr/a:solidFill/a:schemeClr or a:defRPr/a:solidFill/a:srgbClr">
                            <xsl:attribute name="颜色_412F">
                              <xsl:call-template name="color_ziti"/>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:when test="not(a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr">
                              <xsl:attribute name="颜色_412F">
                                <xsl:call-template name="color_ziti"/>
                              </xsl:attribute>
                            </xsl:for-each>
                          </xsl:when>
                        </xsl:choose>
                      </字:字体_4128>
                      <xsl:call-template name="table-tbziti"/>
                    </图表:字体_E70B>
                  </xsl:when>
                  <xsl:when test="a:defRPr[not(a:ea and a:latin)] and not(./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[a:ea or a:latin])">
                    <图表:字体_E70B>
                      <字:字体_4128>
                        <xsl:attribute name="西文字体引用_4129">
                          <xsl:value-of select="concat($sstarget,'.btj_bt_l_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="中文字体引用_412A">
                          <xsl:value-of select="concat($sstarget,'.btj_bt_e_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="字号_412D">
                          <xsl:choose>
                            <xsl:when test="a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:defRPr/@sz) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'18'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <xsl:choose>
                          <xsl:when test="a:defRPr/a:solidFill/a:schemeClr or a:defRPr/a:solidFill/a:srgbClr">
                            <xsl:attribute name="颜色_412F">
                              <xsl:call-template name="color_ziti"/>
                            </xsl:attribute>
                          </xsl:when>
                          <xsl:when test="not(a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr">
                              <xsl:attribute name="颜色_412F">
                                <xsl:call-template name="color_ziti"/>
                              </xsl:attribute>
                            </xsl:for-each>
                          </xsl:when>
                        </xsl:choose>
                      </字:字体_4128>
                      <xsl:call-template name="table-tbziti"/>
                    </图表:字体_E70B>
                  </xsl:when>
                </xsl:choose>
              </xsl:for-each>
              <xsl:call-template name="table-btj-dq-piechart">
                <xsl:with-param name="pos" select="1"/>
              </xsl:call-template>
            </xsl:for-each>
          </图表:标题_E736>
        </xsl:for-each>
      </xsl:if>
    </xsl:if>
  </xsl:template>
  <xsl:template name="table-btj-dq">
    <图表:对齐_E726>
      <!--水平对齐方式-->
      <xsl:if test="a:bodyPr[@vert='eaVert'] and a:bodyPr[@anchor]">
        <表:水平对齐方式_E700>
          <xsl:choose>
            <xsl:when test="a:bodyPr[@anchor='t']">
              <xsl:value-of select="'left'"/>
            </xsl:when>
            <xsl:when test="a:bodyPr[@anchor='ctr']">
              <xsl:value-of select="'center'"/>
            </xsl:when>
            <xsl:when test="a:bodyPr[@anchor='b']">
              <xsl:value-of select="'right'"/>
            </xsl:when>
            <xsl:when test="a:bodyPr[@anchor='just']">
              <xsl:value-of select="'justify'"/>
            </xsl:when>
            <xsl:when test="a:bodyPr[@anchor='dist']">
              <xsl:value-of select="'distributed'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'center'"/>
            </xsl:otherwise>
          </xsl:choose>
        </表:水平对齐方式_E700>
      </xsl:if>
      <!--垂直对齐方式-->
      <xsl:if test="a:bodyPr[not(@vert='eaVert')] and a:bodyPr[@anchor]">
        <表:垂直对齐方式_E701>
          <xsl:choose>
            <xsl:when test="a:bodyPr[@anchor='t']">
              <xsl:value-of select="'top'"/>
            </xsl:when>
            <xsl:when test="a:bodyPr[@anchor='ctr']">
              <xsl:value-of select="'center'"/>
            </xsl:when>
            <xsl:when test="a:bodyPr[@anchor='b']">
              <xsl:value-of select="'bottom'"/>
            </xsl:when>
            <xsl:when test="a:bodyPr[@anchor='just']">
              <xsl:value-of select="'justify'"/>
            </xsl:when>
            <xsl:when test="a:bodyPr[@anchor='dist']">
              <xsl:value-of select="'distributed'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'top'"/>
            </xsl:otherwise>
          </xsl:choose>
        </表:垂直对齐方式_E701>
      </xsl:if>
      <!--采用默认-->
      <xsl:if test="a:bodyPr[@vert='eaVert'] or a:bodyPr[@vert='wordArtVertRtl']">
        <表:文字排列方向_E703>
          <xsl:value-of select="'r2l-t2b-0e-90w'"/>
        </表:文字排列方向_E703>
      </xsl:if>
      
      <!-- update by xuzhenwei BUG_2513:图表“纵坐标轴标题 旋转过的标题”，转换后文字方向不正确 20130115 -->
      <xsl:if test="a:bodyPr/@rot">
      <xsl:choose>
        <xsl:when test="a:bodyPr[not(@rot='0')]">
        <表:水平对齐方式_E700>left</表:水平对齐方式_E700>
        <表:垂直对齐方式_E701>top</表:垂直对齐方式_E701>
        <表:文字排列方向_E703>t2b-l2r-0e-0w</表:文字排列方向_E703>
          <表:文字旋转角度_E704>
            <xsl:variable name="dre">
              <xsl:value-of select="a:bodyPr/@rot"/>
            </xsl:variable>
            <xsl:value-of select="($dre div 60000) * (-1)"/>
          </表:文字旋转角度_E704>
        </xsl:when>
        <xsl:otherwise>
            <表:水平对齐方式_E700>left</表:水平对齐方式_E700>
            <表:垂直对齐方式_E701>top</表:垂直对齐方式_E701>
            <表:文字排列方向_E703>t2b-l2r-0e-0w</表:文字排列方向_E703>
          <表:文字旋转角度_E704>0</表:文字旋转角度_E704>
        </xsl:otherwise>
        <!-- end -->
      </xsl:choose>
      </xsl:if>
    </图表:对齐_E726>
  </xsl:template>
  <xsl:template name="table-btj-dq-piechart">
    
    <!--2014-4-23 update by Qihy, 修复bug3080互操作转换后图表标题文字对齐方式不一致问题， start-->
    <xsl:param name ="pos"/>
    <!--<xsl:for-each select="ancestor::c:chart/c:title/c:txPr">-->
      <xsl:if test="a:bodyPr">
        <图表:对齐_E726>
          <!--水平对齐方式-->
          <xsl:if test="a:bodyPr[@vert='eaVert'] and a:bodyPr[@anchor]">
            <表:水平对齐方式_E700>
              <xsl:choose>
                <xsl:when test="a:bodyPr[@anchor='t']">
                  <xsl:value-of select="'left'"/>
                </xsl:when>
                <xsl:when test="a:bodyPr[@anchor='ctr']">
                  <xsl:value-of select="'center'"/>
                </xsl:when>
                <xsl:when test="a:bodyPr[@anchor='b']">
                  <xsl:value-of select="'right'"/>
                </xsl:when>
                <xsl:when test="a:bodyPr[@anchor='just']">
                  <xsl:value-of select="'justify'"/>
                </xsl:when>
                <xsl:when test="a:bodyPr[@anchor='dist']">
                  <xsl:value-of select="'distributed'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'center'"/>
                </xsl:otherwise>
              </xsl:choose>
            </表:水平对齐方式_E700>
          </xsl:if>
          <!--垂直对齐方式-->
          <xsl:if test="a:bodyPr[not(@vert='eaVert')] and a:bodyPr[@anchor]">
            <表:垂直对齐方式_E701>
              <xsl:choose>
                <xsl:when test="a:bodyPr[@anchor='t']">
                  <xsl:value-of select="'top'"/>
                </xsl:when>
                <xsl:when test="a:bodyPr[@anchor='ctr']">
                  <xsl:value-of select="'center'"/>
                </xsl:when>
                <xsl:when test="a:bodyPr[@anchor='b']">
                  <xsl:value-of select="'bottom'"/>
                </xsl:when>
                <xsl:when test="a:bodyPr[@anchor='just']">
                  <xsl:value-of select="'justify'"/>
                </xsl:when>
                <xsl:when test="a:bodyPr[@anchor='dist']">
                  <xsl:value-of select="'distributed'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'top'"/>
                </xsl:otherwise>
              </xsl:choose>
            </表:垂直对齐方式_E701>
          </xsl:if>

          <xsl:if test="a:bodyPr[@vert='horz'] or a:bodyPr[not(@vert)]">
            <表:文字排列方向_E703>
              <xsl:value-of select="'t2b-l2r-0e-0w'"/>
            </表:文字排列方向_E703>
          </xsl:if>
          <xsl:if test="a:bodyPr[@vert='eaVert']">
            <表:文字排列方向_E703>
              <xsl:value-of select="'r2l-t2b-0e-90w'"/>
            </表:文字排列方向_E703>
          </xsl:if>
          <xsl:if test="a:bodyPr[@rot]">
            <表:文字旋转角度_E704>
              <xsl:variable name="dre">
                <xsl:value-of select="a:bodyPr/@rot"/>
              </xsl:variable>
              <xsl:value-of select="($dre div 60000) * (-1)"/>
            </表:文字旋转角度_E704>
          </xsl:if>

          <xsl:if test="not(a:bodyPr[@rot]) and $pos = '1'">
            <表:文字旋转角度_E704>0</表:文字旋转角度_E704>
          </xsl:if>
          <xsl:if test="not(a:bodyPr[@rot]) and $pos = '2'">
            <表:文字旋转角度_E704>90</表:文字旋转角度_E704>
          </xsl:if>
        </图表:对齐_E726>
      </xsl:if>
      <!--</xsl:for-each>-->
      <!--2014-4-23 end-->
    
  </xsl:template>
  <xsl:template name="LineTypeTransfer">
    <xsl:param name="lineType"/>
    <xsl:param name="dashType"/>
    <xsl:attribute name="线型_C60D">
      <xsl:choose>
        <xsl:when test="$lineType='sng'">
          <xsl:value-of select="'single'"/>
        </xsl:when>
        <xsl:when test="$lineType='dbl'">
          <xsl:value-of select="'double'"/>
        </xsl:when>
        <xsl:when test="$lineType='thickThin'">
          <xsl:value-of select="'double'"/>
        </xsl:when>
        <xsl:when test="$lineType='thinThick'">
          <xsl:value-of select="'double'"/>
        </xsl:when>
        <xsl:when test="$lineType='tri'">
          <xsl:value-of select="'double'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'single'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <!-- 20130318 update by xuzhenwei BUG_2756:oo-uof 功能测试 样式不正确：字体，线形，数据表格式等 start-->
    <!--<xsl:if test="$lineType">-->
      <xsl:attribute name="虚实_C60E">
        <xsl:choose>
          <xsl:when test="$dashType='solid'">
            <xsl:value-of select="'solid'"/>
          </xsl:when>
          <xsl:when test="$dashType='dot'">
            <xsl:value-of select="'dash'"/>
          </xsl:when>
          <xsl:when test="$dashType='dash'">
            <xsl:value-of select="'dash'"/>
          </xsl:when>
          <xsl:when test="$dashType='lgDash'">
            <xsl:value-of select="'long-dash'"/>
          </xsl:when>
          <xsl:when test="$dashType='dashDot'">
            <xsl:value-of select="'dash-dot'"/>
          </xsl:when>
          <xsl:when test="$dashType='lgDashDot'">
            <xsl:value-of select="'dash-dot'"/>
          </xsl:when>
          <xsl:when test="$dashType='lgDashDotDot'">
            <xsl:value-of select="'dash-dot-dot'"/>
          </xsl:when>
          <xsl:when test="$dashType='sysDash'">
            <xsl:value-of select="'dash'"/>
          </xsl:when>
          <xsl:when test="$dashType='sysDot'">
            <xsl:value-of select="'dash'"/>
          </xsl:when>
          <xsl:when test="$dashType='sysDashDot'">
            <xsl:value-of select="'dash-dot'"/>
          </xsl:when>
          <xsl:when test="$dashType='sysDashDotDot'">
            <xsl:value-of select="'dash-dot-dot'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'solid'"/>
            <!--'dash'-->
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    <!--</xsl:if>-->
    <!-- end -->
    <!---->
  </xsl:template>
  <xsl:template match="c:txPr">
    <xsl:param name="type"/>
    <xsl:param name="sstarget"/>
      <xsl:for-each select="./a:p/a:pPr/a:defRPr">
        <图表:字体_E70B>
          <字:字体_4128>
            <!--中文字体-->
            <xsl:if test="./a:ea">
              <xsl:attribute name="中文字体引用_412A">
                <xsl:value-of select="concat($type,'_ea_',$sstarget)"/>
              </xsl:attribute>
            </xsl:if>
            <!--西文字体-->
            <xsl:if test="./a:latin">
              <xsl:attribute name="西文字体引用_4129">
                <xsl:value-of select="concat($type,'_la_',$sstarget)"/>
              </xsl:attribute>
            </xsl:if>
            <!--字号-->
            <xsl:if test="@sz or ./ancestor::c:chartSpace/c:txPr">
              <xsl:attribute name="字号_412D">
                <xsl:variable name="sz">
                  <xsl:choose>
                    
                    <!--2014-4-23, update by Qihy, bug3080，坐标轴字体大小不正确， start-->
                    <xsl:when test="@sz">
                      <xsl:value-of select="@sz"/>
                    </xsl:when>
                    <!--2014-4-23 end-->
                    
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                      <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'1200'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>
                <xsl:value-of select="$sz div 100"/>
              </xsl:attribute>
            </xsl:if>
            <!--颜色-->
            <xsl:choose>
              <xsl:when test="./a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                <xsl:attribute name="颜色_412F">
                  <xsl:call-template name="color"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[a:solidFill/a:srgbClr or a:solidFill/a:schemeClr]">
                  <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr">
                    <xsl:attribute name="颜色_412F">
                      <xsl:call-template name="color"/>
                    </xsl:attribute>
                  </xsl:for-each>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
          </字:字体_4128>
          <!--粗体-->
          <xsl:choose>
            <xsl:when test="./@b='1'">
              <字:是否粗体_4130>true</字:是否粗体_4130>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@b='1'">
                <字:是否粗体_4130>true</字:是否粗体_4130>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <!--斜体-->
          <xsl:choose>
            <xsl:when test="./@i='1'">
              <字:是否斜体_4131>true</字:是否斜体_4131>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@i='1'">
                <字:是否斜体_4131>true</字:是否斜体_4131>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <!--删除线-->
          <xsl:choose>
            <xsl:when test="./@strike">
              <字:删除线_4135>
                <xsl:choose>
                  <xsl:when test="./@strike='dblStrike'">
                    <xsl:value-of select="'double'"/>
                  </xsl:when>
                  <xsl:when test="./@strike='noStrike'">
                    <xsl:value-of select="'none'"/>
                  </xsl:when>
                  <xsl:when test="./@strike='sngStrike'">
                    <xsl:value-of select="'single'"/>
                  </xsl:when>
                </xsl:choose>
              </字:删除线_4135>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@strike">
                <字:删除线_4135>
                  <xsl:choose>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@strike='dblStrike'">
                      <xsl:value-of select="'double'"/>
                    </xsl:when>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@strike='noStrike'">
                      <xsl:value-of select="'none'"/>
                    </xsl:when>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@strike='sngStrike'">
                      <xsl:value-of select="'single'"/>
                    </xsl:when>
                  </xsl:choose>
                </字:删除线_4135>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <!--下划线-->
          <xsl:choose>
            <xsl:when test="./@u">
              <xsl:call-template name="UnderlineTypeTransfer">
                <xsl:with-param name="lineType" select="./@u"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@u">
                <xsl:call-template name="UnderlineTypeTransfer">
                  <xsl:with-param name="lineType" select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@u"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <!--上下标-->
          <xsl:choose>
            <xsl:when test="./@baseline">
              <字:上下标类型_4143>
                <xsl:choose>
                  <xsl:when test="./@baseline = '0'">
                    <xsl:value-of select="'none'"/>
                  </xsl:when>
                  <xsl:when test="./@baseline &gt; 0">
                    <xsl:value-of select="'sup'"/>
                  </xsl:when>
                  <xsl:when test="./@baseline &lt; 0">
                    <xsl:value-of select="'sub'"/>
                  </xsl:when>
                </xsl:choose>
              </字:上下标类型_4143>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@baseline">
                <字:上下标类型_4143>
                  <xsl:choose>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@baseline = '0'">
                      <xsl:value-of select="'none'"/>
                    </xsl:when>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@baseline &gt; '0'">
                      <xsl:value-of select="'sup'"/>
                    </xsl:when>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@baseline &lt; '0'">
                      <xsl:value-of select="'sub'"/>
                    </xsl:when>
                  </xsl:choose>
                </字:上下标类型_4143>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <!--字符间距-->
          <xsl:choose>
            <xsl:when test="./@spc">
              <字:字符间距_4145>
                <xsl:value-of select="./@spc div 100"/>
              </字:字符间距_4145>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@spc">
                <字:字符间距_4145>
                  <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@spc div 100"/>
                </字:字符间距_4145>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <!--醒目字体类型-->
          <xsl:choose>
            <xsl:when test="./@cap='all'">
              <字:醒目字体类型_4141>uppercase</字:醒目字体类型_4141>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@cap='all'">
                <字:醒目字体类型_4141>uppercase</字:醒目字体类型_4141>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="./@cap='small'">
              <字:醒目字体类型_4141>small-caps</字:醒目字体类型_4141>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@cap='small'">
                <字:醒目字体类型_4141>small-caps</字:醒目字体类型_4141>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </图表:字体_E70B>
      </xsl:for-each>
  </xsl:template>
  <xsl:template name="FontTransfer">
    <xsl:param name="type"/>
    <xsl:param name="sstarget"/>
    <xsl:param name="target"/>
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:txPr">
      <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:txPr/a:p/a:pPr/a:defRPr">
        <图表:字体_E70B>
          <字:字体_4128>
            <!--中文字体-->
            <xsl:if test="./a:ea">
              <xsl:attribute name="中文字体引用_412A">
                <xsl:value-of select="concat($type,'_ea_',$sstarget)"/>
              </xsl:attribute>
            </xsl:if>
            <!--西文字体-->
            <xsl:if test="./a:latin">
              <xsl:attribute name="西文字体引用_4129">
                <xsl:value-of select="concat($type,'_la_',$sstarget)"/>
              </xsl:attribute>
            </xsl:if>
            <!--字号-->
            <xsl:if test="@sz or ./ancestor::c:chartSpace/c:txPr">
              <xsl:attribute name="字号_412D">
                <xsl:variable name="sz">
                  <xsl:choose>
                    <xsl:when test="@sz">
                      <xsl:value-of select="@sz"/>
                    </xsl:when>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                      <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'1200'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>
                <xsl:value-of select="$sz div 100"/>
              </xsl:attribute>
            </xsl:if>
            <!--颜色-->
            <xsl:choose>
              <xsl:when test="./a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                <xsl:attribute name="颜色_412F">
                  <xsl:call-template name="color"/>
                </xsl:attribute>
              </xsl:when>
              <xsl:otherwise>
                <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[a:solidFill/a:srgbClr or a:solidFill/a:schemeClr]">
                  <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr">
                    <xsl:attribute name="颜色_412F">
                      <xsl:call-template name="color"/>
                    </xsl:attribute>
                  </xsl:for-each>
                </xsl:if>
              </xsl:otherwise>
            </xsl:choose>
          </字:字体_4128>
          <!--粗体-->
          <xsl:choose>
            <xsl:when test="./@b='1'">
              <字:是否粗体_4130>true</字:是否粗体_4130>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@b='1'">
                <字:是否粗体_4130>true</字:是否粗体_4130>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <!--斜体-->
          <xsl:choose>
            <xsl:when test="./@i='1'">
              <字:是否斜体_4131>true</字:是否斜体_4131>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@i='1'">
                <字:是否斜体_4131>true</字:是否斜体_4131>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <!--删除线-->
          <xsl:choose>
            <xsl:when test="./@strike">
              <字:删除线_4135>
                <xsl:choose>
                  <xsl:when test="./@strike='dblStrike'">
                    <xsl:value-of select="'double'"/>
                  </xsl:when>
                  <xsl:when test="./@strike='noStrike'">
                    <xsl:value-of select="'none'"/>
                  </xsl:when>
                  <xsl:when test="./@strike='sngStrike'">
                    <xsl:value-of select="'single'"/>
                  </xsl:when>
                </xsl:choose>
              </字:删除线_4135>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@strike">
                <字:删除线_4135>
                  <xsl:choose>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@strike='dblStrike'">
                      <xsl:value-of select="'double'"/>
                    </xsl:when>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@strike='noStrike'">
                      <xsl:value-of select="'none'"/>
                    </xsl:when>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@strike='sngStrike'">
                      <xsl:value-of select="'single'"/>
                    </xsl:when>
                  </xsl:choose>
                </字:删除线_4135>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <!--下划线-->
          <xsl:choose>
            <xsl:when test="./@u">
              <xsl:call-template name="UnderlineTypeTransfer">
                <xsl:with-param name="lineType" select="./@u"/>
              </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@u">
                <xsl:call-template name="UnderlineTypeTransfer">
                  <xsl:with-param name="lineType" select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@u"/>
                </xsl:call-template>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <!--上下标-->
          <xsl:choose>
            <xsl:when test="./@baseline">
              <字:上下标类型_4143>
                <xsl:choose>
                  <xsl:when test="./@baseline = '0'">
                    <xsl:value-of select="'none'"/>
                  </xsl:when>
                  <xsl:when test="./@baseline &gt; 0">
                    <xsl:value-of select="'sup'"/>
                  </xsl:when>
                  <xsl:when test="./@baseline &lt; 0">
                    <xsl:value-of select="'sub'"/>
                  </xsl:when>
                </xsl:choose>
              </字:上下标类型_4143>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@baseline">
                <字:上下标类型_4143>
                  <xsl:choose>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@baseline = '0'">
                      <xsl:value-of select="'none'"/>
                    </xsl:when>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@baseline &gt; '0'">
                      <xsl:value-of select="'sup'"/>
                    </xsl:when>
                    <xsl:when test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@baseline &lt; '0'">
                      <xsl:value-of select="'sub'"/>
                    </xsl:when>
                  </xsl:choose>
                </字:上下标类型_4143>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <!--字符间距-->
          <xsl:choose>
            <xsl:when test="./@spc">
              <字:字符间距_4145>
                <xsl:value-of select="./@spc div 100"/>
              </字:字符间距_4145>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@spc">
                <字:字符间距_4145>
                  <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@spc div 100"/>
                </字:字符间距_4145>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <!--醒目字体类型-->
          <xsl:choose>
            <xsl:when test="./@cap='all'">
              <字:醒目字体类型_4141>uppercase</字:醒目字体类型_4141>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@cap='all'">
                <字:醒目字体类型_4141>uppercase</字:醒目字体类型_4141>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="./@cap='small'">
              <字:醒目字体类型_4141>small-caps</字:醒目字体类型_4141>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@cap='small'">
                <字:醒目字体类型_4141>small-caps</字:醒目字体类型_4141>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </图表:字体_E70B>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
  <xsl:template name="chart-fenleizhou">
    <!--分类轴-->
    <xsl:param name="rid"/>
    <xsl:param name="target"/>
    <xsl:variable name="starget">
      <xsl:value-of select="substring-after($target,'charts/')"/>
    </xsl:variable>
    <xsl:variable name="sstarget">
      <xsl:value-of select="substring-before($starget,'.')"/>
    </xsl:variable>
    <xsl:variable name="t">
      <xsl:value-of select="substring-after($target,'../')"/>
    </xsl:variable>
    <xsl:variable name="tt">
      <xsl:value-of select="concat('xlsx/xl/',$t)"/>
    </xsl:variable>
    <xsl:attribute name="主刻度类型_E737">
      <xsl:choose>
        <xsl:when test="c:majorTickMark">
          <xsl:choose>
            <xsl:when test="c:majorTickMark[@val='none']">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test="c:majorTickMark[@val='out']">
              <xsl:value-of select="'outside'"/>
            </xsl:when>
            <xsl:when test="c:majorTickMark[@val='in']">
              <xsl:value-of select="'inside'"/>
            </xsl:when>
            <xsl:when test="c:majorTickMark[@val='cross']">
              <xsl:value-of select="'cross'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'inside'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="次刻度类型_E738">
      <xsl:choose>
        <xsl:when test="c:minorTickMark">
          <xsl:choose>
            <xsl:when test="c:minorTickMark[@val='none']">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test="c:minorTickMark[@val='out']">
              <xsl:value-of select="'outside'"/>
            </xsl:when>
            <xsl:when test="c:minorTickMark[@val='in']">
              <xsl:value-of select="'inside'"/>
            </xsl:when>
            <xsl:when test="c:minorTickMark[@val='cross']">
              <xsl:value-of select="'cross'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'none'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="刻度线标志_E739">
      <xsl:choose>
        <xsl:when test="c:tickLblPos">
          <xsl:choose>
            <xsl:when test="c:tickLblPos[@val='none']">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test="c:tickLblPos[@val='high']">
              <xsl:value-of select="'outside'"/>
            </xsl:when>
            <xsl:when test="c:tickLblPos[@val='low']">
              <xsl:value-of select="'next to axis'"/>
            </xsl:when>
            <xsl:when test="c:tickLblPos[@val='nextTo']">
              <xsl:value-of select="'next to axis'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'next to axis'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="主类型_E792">
      <xsl:value-of select="'primary'"/>
    </xsl:attribute>
    <xsl:attribute name="子类型_E793">
      <xsl:value-of select="'category'"/>
    </xsl:attribute>

    <!-- update by 凌峰 边框线默认情况 2014.5.4 start -->
    <xsl:if test="not(ancestor::c:plotArea/c:catAx/c:spPr/a:ln)">
      <图表:边框线_4226 线型_C60D="single" 虚实_C60E="solid" 宽度_C60F="2" 颜色_C611="#808080"/>
    </xsl:if>
      <!--end-->
      
    <xsl:for-each select="ancestor::c:plotArea/c:catAx/c:spPr/a:ln">
      <图表:边框线_4226>
        <xsl:choose>
          <xsl:when test="./a:noFill">
            <xsl:attribute name="线型_C60D">
              <xsl:value-of select="'none'"/>
            </xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="lineType">
              <xsl:choose>
                <xsl:when test="./@cmpd">
                  <xsl:value-of select="./@cmpd"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'none'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="dashType">
              <xsl:choose>
                <xsl:when test="./a:prstDash/@val">
                  <xsl:value-of select="./a:prstDash/@val"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'none'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:call-template name="LineTypeTransfer">
              <xsl:with-param name="lineType" select="$lineType"/>
              <xsl:with-param name="dashType" select="$dashType"/>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
        <!--share-->
        <xsl:if test="@w">
          <xsl:attribute name="宽度_C60F">
            <xsl:value-of select="'4'"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
          <xsl:attribute name="颜色_C611">
            <xsl:call-template name="color"/>
          </xsl:attribute>
        </xsl:if>
        <!--share-->
      </图表:边框线_4226>
    </xsl:for-each>
    <!--数值-->
    <xsl:if test="c:numFmt">
      <图表:数值_E70D>
        <xsl:attribute name="是否链接到源_E73E">
          <xsl:choose>
            <xsl:when test="c:numFmt[not(@sourceLinked)] or c:numFmt[@sourceLinked='1']">
              <xsl:value-of select="'true'"/>
            </xsl:when>
            <xsl:when test="c:numFmt[@sourceLinked='0']">
              <xsl:value-of select="'false'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
        <!-- 20130518 update by xuzhenwei BUG_2881:回归集成oo-uof工作表"折线图"分类轴X轴转换前为年月，转换后为数字 start -->
        <xsl:if test="contains(c:numFmt/@formatCode,'yyyy')">
          <xsl:attribute name="分类名称_E740">
            <xsl:value-of select="'date'"/>
          </xsl:attribute>
            <xsl:attribute name="格式码_E73F">
                <xsl:value-of select="c:numFmt/@formatCode"/>
            </xsl:attribute>
          <!-- end -->
        </xsl:if>
          
          <xsl:if test="contains(c:numFmt/@formatCode,'$') and contains(c:numFmt/@formatCode,'#,##0_')">
              <xsl:attribute name="分类名称_E740">
                  <xsl:value-of select="'custom'"/>
              </xsl:attribute>
              <xsl:attribute name="格式码_E73F">
                  <xsl:value-of select="'$#,##0_);($#,##0)'"/>
              </xsl:attribute>
          </xsl:if>
        
      </图表:数值_E70D>
    </xsl:if>
    <xsl:if test="not(c:numFmt)">
      <!--默认值-->
      <图表:数值_E70D 是否链接到源_E73E="true"/>
    </xsl:if>
    <xsl:apply-templates select="./c:txPr">
      <xsl:with-param name="type" select="'catAx'"/>
      <xsl:with-param name="sstarget" select="$sstarget"/>
    </xsl:apply-templates>
    <!--<xsl:call-template name="FontTransfer">
      <xsl:with-param name="type" select="'catAx'"/>
      <xsl:with-param name="sstarget" select="$sstarget"/>
      <xsl:with-param name="target" select="$target"/>
    </xsl:call-template>-->
    <!--刻度-->
    <xsl:if test="c:scaling">
      <图表:刻度_E71D>
        <xsl:for-each select="c:scaling">
          <xsl:if test="c:min">
            <图表:最小值_E71E 是否自动_E71F="false">
              <xsl:value-of select="c:min/@val"/>
            </图表:最小值_E71E>
          </xsl:if>
          <xsl:if test="c:max">
            <图表:最大值_E720 是否自动_E71F="false">
              <xsl:value-of select="c:max/@val"/>
            </图表:最大值_E720>
          </xsl:if>
        </xsl:for-each>
        <xsl:if test="c:majorUnit">
          <图表:主单位_E721 是否自动_E71F="false">
            <xsl:value-of select="c:majorUnit/@val"/>
          </图表:主单位_E721>
        </xsl:if>
        <xsl:if test="c:minorUnit">
          <图表:次单位_E722 是否自动_E71F="false">
            <xsl:value-of select="c:minorUnit/@val"/>
          </图表:次单位_E722>
        </xsl:if>
        <!--交叉点-->
        <xsl:if test="c:dispUnits">
          <显示单位_E724>
            <xsl:if test="c:dispUnits[c:builtInUnit]">
              <xsl:attribute name="类型_E728">
                <xsl:choose>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='hundreds']">
                    <xsl:value-of select="'hundreds'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='thousands']">
                    <xsl:value-of select="'thousands'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='tenThousands']">
                    <xsl:value-of select="'ten thousands'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='hundredThousands']">
                    <xsl:value-of select="'one hundred thousands'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='millons']">
                    <xsl:value-of select="'millons'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='tenMillons']">
                    <xsl:value-of select="'ten millons'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='hundredMillons']">
                    <xsl:value-of select="'hundred millons'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='trillions']">
                    <xsl:value-of select="'trillions'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='billions']">
                    <xsl:value-of select="'trillions'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="c:dispUnits[c:dispUnitsLbl/c:layout]">
              <xsl:attribute name="是否显示_E727">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
            </xsl:if>
          </显示单位_E724>
        </xsl:if>
        <xsl:if test="c:scaling">
          <xsl:if test="c:scaling[c:logBase]">
            <图表:是否显示为对数刻度_E729>
              <!--xsl:value-of select="c:scaling/c:logBase/@val"/-->
              <xsl:value-of select="'true'"/>
            </图表:是否显示为对数刻度_E729>
          </xsl:if>
        </xsl:if>
        <xsl:if test="c:axPos">
          <xsl:if test="c:axPos[@val='r']">
            <图表:是否次序反转_E72B>
              <xsl:value-of select="'true'"/>
            </图表:是否次序反转_E72B>
          </xsl:if>
          <xsl:if test="c:axPos[@val='b']">
            <图表:是否次序反转_E72B>
              <xsl:value-of select="'false'"/>
            </图表:是否次序反转_E72B>
          </xsl:if>
        </xsl:if>
        <!-- add by xuzhenwei bug_2502:横坐标轴交叉于20, 20130115 start 此次代码没问题，由于永中软件有问题，显示效果是错误的. -->
        <xsl:if test="c:crossesAt">
          <图表:交叉点_E723 是否自动_E71F="false">
            <xsl:value-of select="./c:crossesAt/@val"/>
          </图表:交叉点_E723>
        </xsl:if>
        <!-- end -->
        <!--分类刻度数 -->
        <!-- add by xuzhenwei bug_2500:图表 刻度线间隔:2 标签间隔：2设置丢失 2013-01-19 start-->
        <xsl:if test="c:tickMarkSkip">
          <图表:分类刻度数_E72D 是否自动_E71F="false">
            <xsl:value-of select="c:tickMarkSkip/@val"/>
          </图表:分类刻度数_E72D>
        </xsl:if>
        <!-- end -->
        <!--分类标签数 -->
        <xsl:if test="c:tickLblSkip">
          <图表:分类标签数_E72C 是否自动_E71F="false">
            <xsl:value-of select="c:tickLblSkip/@val"/>
          </图表:分类标签数_E72C>
        </xsl:if>
        <!--是否交叉于最大值-->
        <!--数值轴是否置于分类轴-->
      </图表:刻度_E71D>
    </xsl:if>
    <!--对齐-->
    <xsl:if test="c:txPr or c:lblOffset">
      <xsl:if test="c:txPr[a:bodyPr] or c:lblOffset">
        <xsl:if test="c:txPr/a:bodyPr[@vert] or c:txPr/a:bodyPr[@rot] or c:lblOffset">
          <图表:对齐_E730>
            <xsl:if test="c:txPr/a:bodyPr[@vert]">
              <图表:文字排列方向_E703>
                <xsl:choose>
                  <xsl:when test="c:txPr/a:bodyPr[@vert='horz']">
                    <xsl:value-of select="'t2b-l2r-0e-0w'"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'r2l-t2b-0e-90w'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </图表:文字排列方向_E703>
            </xsl:if>
            <xsl:if test="c:txPr/a:bodyPr[@rot]">
              <图表:文字旋转角度_E704>
                <xsl:variable name="x">
                  <xsl:value-of select="c:txPr/a:bodyPr/@rot"/>
                </xsl:variable>
                <xsl:value-of select="($x div 60000) * (-1)"/>
              </图表:文字旋转角度_E704>
            </xsl:if>
            <xsl:if test="c:lblOffset">
              <xsl:if test="c:lblOffset[@val]">
                <图表:偏移量_E732>
                  <xsl:value-of select="c:lblOffset/@val"/>
                </图表:偏移量_E732>
              </xsl:if>
            </xsl:if>
          </图表:对齐_E730>
        </xsl:if>
      </xsl:if>
    </xsl:if>
    <!--网络线集-->
    <xsl:call-template name="chart-wgxj-cat">
      <xsl:with-param name="tar" select="$tt"/>
    </xsl:call-template>
    <!--标题-->
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:plotArea/c:catAx/c:title]">
      <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title">
        <图表:标题_E736>
          <xsl:choose>
            <xsl:when test="descendant::a:t">
              <xsl:attribute name="名称_E742">
                <xsl:variable name="text">
                  <xsl:for-each select="descendant::a:t">
                    <xsl:value-of select="."/>
                  </xsl:for-each>
                </xsl:variable>
                <xsl:value-of select="$text"/>
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="名称_E742">
                <xsl:variable name="text">
                  <xsl:value-of select="'坐标轴标题'"/>
                </xsl:variable>
                <xsl:value-of select="$text"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <!--边框+填充-->
          <xsl:if test="c:spPr">
            <xsl:for-each select="c:spPr">
              <xsl:call-template name="chart-biankuang">
                <xsl:with-param name="target2" select="$target"/>
              </xsl:call-template>
              <xsl:call-template name="chart-tianchong">
                <xsl:with-param name="target2" select="$target"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>
          <xsl:if test="not(c:tx/c:rich/a:p) and not(c:txPr)">
            <图表:字体_E70B>
              <字:字体_4128>
                <xsl:attribute name="字号_412D">
                  <xsl:value-of select="'10'"/>
                </xsl:attribute>
              </字:字体_4128>
            </图表:字体_E70B>
          </xsl:if>
          <!--字体-->
          <xsl:if test="c:tx/c:rich/a:p">
            <xsl:for-each select="c:tx/c:rich/a:p">
              <图表:字体_E70B>
                <xsl:for-each select="a:r">
                  <xsl:choose>
                    <xsl:when test="./a:rPr">
                      <!--yanghaojie
                      <xsl:if test="./a:rPr[@i='1']">
                        <字:是否斜体_4131>true</字:是否斜体_4131>
                      </xsl:if>
                      -->
                      <字:字体_4128>
                        <xsl:if test="./a:rPr/a:latin">
                          <xsl:attribute name="西文字体引用_4129">
                            <xsl:value-of select="concat($sstarget,'.btj_flz_l_font')"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./a:rPr/a:ea">
                          <xsl:attribute name="中文字体引用_412A">
                            <xsl:value-of select="concat($sstarget,'.btj_flz_e_font')"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:attribute name="字号_412D">
                          <xsl:choose>
                            <xsl:when test="a:rPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="a:rPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:rPr/@sz) and ./preceding-sibling::a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="./preceding-sibling::a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:rPr/@sz) and not(./preceding-sibling::a:pPr/a:defRPr/@sz) and  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select=" ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'10'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <xsl:choose>
                          <xsl:when test="a:rPr/a:solidFill">
                            <xsl:for-each select="a:rPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:when test="not(a:rPr/a:solidFill) and ./preceding-sibling::a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./preceding-sibling::a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:when test="not(a:rPr/a:solidFill) and not(./preceding-sibling::a:pPr/a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                        </xsl:choose>
                      </字:字体_4128>
                    </xsl:when>
                    <xsl:when test="not(./a:rPr[a:ea or a:latin]) and ./preceding-sibling::a:pPr/a:defRPr[a:ea or a:latin]">
                      <字:字体_4128>
                        <xsl:attribute name="西文字体引用_4129">
                          <xsl:value-of select="concat($sstarget,'.btj_flz_l_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="中文字体引用_412A">
                          <xsl:value-of select="concat($sstarget,'.btj_flz_e_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="字号_412D">
                          <xsl:choose>
                            <xsl:when test="a:rPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="a:rPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:rPr/@sz) and ./preceding-sibling::a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="./preceding-sibling::a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:rPr/@sz) and not(./preceding-sibling::a:pPr/a:defRPr/@sz) and  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select=" ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'10'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <xsl:choose>
                          <xsl:when test="a:rPr/a:solidFill">
                            <xsl:for-each select="a:rPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:when test="not(a:rPr/a:solidFill) and ./preceding-sibling::a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./preceding-sibling::a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:when test="not(a:rPr/a:solidFill) and not(./preceding-sibling::a:pPr/a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                        </xsl:choose>
                      </字:字体_4128>
                    </xsl:when>
                    <xsl:when test="not(./a:rPr[a:ea or a:latin]) and not (./preceding-sibling::a:pPr/a:defRPr[a:ea or a:latin]) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[a:ea or a:latin]">
                      <字:字体_4128>
                        <xsl:attribute name="西文字体引用_4129">
                          <xsl:value-of select="concat($sstarget,'.tbq_l_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="中文字体引用_412A">
                          <xsl:value-of select="concat($sstarget,'.tbq_e_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="字号_412D">
                          <xsl:choose>
                            <xsl:when test="a:rPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="a:rPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:rPr/@sz) and ./preceding-sibling::a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="./preceding-sibling::a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:rPr/@sz) and not(./preceding-sibling::a:pPr/a:defRPr/@sz) and  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select=" ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'10'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <xsl:choose>
                          <xsl:when test="a:rPr/a:solidFill">
                            <xsl:for-each select="a:rPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:when test="not(a:rPr/a:solidFill) and ./preceding-sibling::a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./preceding-sibling::a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:when test="not(a:rPr/a:solidFill) and not(./preceding-sibling::a:pPr/a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                        </xsl:choose>
                      </字:字体_4128>
                    </xsl:when>
                  </xsl:choose>
                  <xsl:if test="./a:rPr">
                    <xsl:call-template name="table-btj-ziti"/>
                  </xsl:if>
                </xsl:for-each>
              </图表:字体_E70B>

                  <xsl:for-each select="ancestor::c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich">
                    <xsl:if test="a:bodyPr">
                      <xsl:call-template name="table-btj-dq"/>
                    </xsl:if>
                  </xsl:for-each>
            </xsl:for-each>
          </xsl:if>
          <xsl:for-each select="c:txPr">
            <xsl:for-each select="a:p/a:pPr">
              <xsl:choose>
                <xsl:when test="a:defRPr[a:ea or a:latin]">
                  <图表:字体_E70B>
                    <字:字体_4128>
                      <xsl:attribute name="西文字体引用_4129">
                        <xsl:value-of select="concat($sstarget,'.btj_flz_l_font')"/>
                      </xsl:attribute>
                      <xsl:attribute name="中文字体引用_412A">
                        <xsl:value-of select="concat($sstarget,'.btj_flz_e_font')"/>
                      </xsl:attribute>
                      <xsl:attribute name="字号_412D">
                        <xsl:choose>
                          <xsl:when test="a:defRPr/@sz">
                            <xsl:variable name="sz">
                              <xsl:value-of select="a:defRPr/@sz"/>
                            </xsl:variable>
                            <xsl:value-of select="$sz div 100"/>
                          </xsl:when>
                          <xsl:when test="not(a:defRPr/@sz) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                            <xsl:variable name="sz">
                              <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                            </xsl:variable>
                            <xsl:value-of select="$sz div 100"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="'10'"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:attribute>
                      <xsl:choose>
                        <xsl:when test="a:defRPr/a:solidFill/a:schemeClr or a:defRPr/a:solidFill/a:srgbClr">
                          <xsl:attribute name="颜色_412F">
                            <xsl:call-template name="color_ziti"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test="not(a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                          <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr">
                            <xsl:attribute name="颜色_412F">
                              <xsl:call-template name="color_ziti"/>
                            </xsl:attribute>
                          </xsl:for-each>
                        </xsl:when>
                      </xsl:choose>
                    </字:字体_4128>
                    <xsl:call-template name="table-tbziti"/>
                  </图表:字体_E70B>
                </xsl:when>
                <xsl:when test="not(a:defRPr[a:ea or a:latin]) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[a:ea or a:latin]">
                  <图表:字体_E70B>
                    <字:字体_4128>
                      <xsl:attribute name="西文字体引用_4129">
                        <xsl:value-of select="concat($sstarget,'.tbq_l_font')"/>
                      </xsl:attribute>
                      <xsl:attribute name="中文字体引用_412A">
                        <xsl:value-of select="concat($sstarget,'.tbq_e_font')"/>
                      </xsl:attribute>
                      <xsl:attribute name="字号_412D">
                        <xsl:choose>
                          <xsl:when test="a:defRPr/@sz">
                            <xsl:variable name="sz">
                              <xsl:value-of select="a:defRPr/@sz"/>
                            </xsl:variable>
                            <xsl:value-of select="$sz div 100"/>
                          </xsl:when>
                          <xsl:when test="not(a:defRPr/@sz) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                            <xsl:variable name="sz">
                              <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                            </xsl:variable>
                            <xsl:value-of select="$sz div 100"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="'10'"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:attribute>
                      <xsl:choose>
                        <xsl:when test="a:defRPr/a:solidFill/a:schemeClr or a:defRPr/a:solidFill/a:srgbClr">
                          <xsl:attribute name="颜色_412F">
                            <xsl:call-template name="color_ziti"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test="not(a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                          <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr">
                            <xsl:attribute name="颜色_412F">
                              <xsl:call-template name="color_ziti"/>
                            </xsl:attribute>
                          </xsl:for-each>
                        </xsl:when>
                      </xsl:choose>
                    </字:字体_4128>
                    <xsl:call-template name="table-tbziti"/>
                  </图表:字体_E70B>
                </xsl:when>
              </xsl:choose>
            </xsl:for-each>
            <xsl:call-template name="table-btj-dq-piechart">
              <xsl:with-param name="pos" select="1"/>
            </xsl:call-template>
          </xsl:for-each>
        </图表:标题_E736>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
    
  <xsl:template name="chart-shuzhizhou">
    <!--数值轴-->
      <!-- update by 凌峰 BUG_3080：坐标轴标题丢失，坐标刻度值变化 20140308 start -->
      <xsl:param name="pos"/>
      <!--end-->
      <xsl:param name="rid"/>
    <xsl:param name="target"/>
    <xsl:variable name="starget">
      <xsl:value-of select="substring-after($target,'charts/')"/>
    </xsl:variable>
    <xsl:variable name="sstarget">
      <xsl:value-of select="substring-before($starget,'.')"/>
    </xsl:variable>
    <xsl:variable name="t">
      <xsl:value-of select="substring-after($target,'../')"/>
    </xsl:variable>
    <xsl:variable name="tt">
      <xsl:value-of select="concat('xlsx/xl/',$t)"/>
    </xsl:variable>
    <xsl:attribute name="主刻度类型_E737">
      <xsl:choose>
        <xsl:when test="c:majorTickMark">
          <xsl:choose>
            <xsl:when test="c:majorTickMark[@val='none']">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test="c:majorTickMark[@val='out']">
              <xsl:value-of select="'outside'"/>
            </xsl:when>
            <xsl:when test="c:majorTickMark[@val='in']">
              <xsl:value-of select="'inside'"/>
            </xsl:when>
            <xsl:when test="c:majorTickMark[@val='cross']">
              <xsl:value-of select="'cross'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'outside'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="次刻度类型_E738">
      <xsl:choose>
        <xsl:when test="c:minorTickMark">
          <xsl:choose>
            <xsl:when test="c:minorTickMark[@val='none']">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test="c:minorTickMark[@val='out']">
              <xsl:value-of select="'outside'"/>
            </xsl:when>
            <xsl:when test="c:minorTickMark[@val='in']">
              <xsl:value-of select="'inside'"/>
            </xsl:when>
            <xsl:when test="c:minorTickMark[@val='cross']">
              <xsl:value-of select="'cross'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'none'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="刻度线标志_E739">
      <xsl:choose>
        <xsl:when test="c:tickLblPos">
          <xsl:choose>
            <xsl:when test="c:tickLblPos[@val='none']">
              <xsl:value-of select="'none'"/>
            </xsl:when>
            <xsl:when test="c:tickLblPos[@val='high']">
              <xsl:value-of select="'inside'"/>
            </xsl:when>
            <xsl:when test="c:tickLblPos[@val='low']">
              <xsl:value-of select="'outside'"/>
            </xsl:when>
            <xsl:when test="c:tickLblPos[@val='nextTo']">
              <xsl:value-of select="'next to axis'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <!--yx,delete 10.16 13:36-->
        <xsl:otherwise>
          <xsl:value-of select="'next to axis'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
    <xsl:attribute name="主类型_E792">
      <xsl:value-of select="'primary'"/>
    </xsl:attribute>
    <xsl:attribute name="子类型_E793">
      <xsl:value-of select="'value'"/>
    </xsl:attribute>
    <!--线型-->

    <!-- update by 凌峰 边框线默认情况 2014.5.4 start -->
    <xsl:if test="not(ancestor::c:plotArea/c:valAx/c:spPr/a:ln)">
      <图表:边框线_4226 线型_C60D="single" 虚实_C60E="solid" 宽度_C60F="2" 颜色_C611="#808080"/>
    </xsl:if>
    <!--end-->
    
    <xsl:for-each select="ancestor::c:plotArea/c:valAx/c:spPr/a:ln">
      <图表:边框线_4226>
        <xsl:variable name="lineType">
          <xsl:choose>
            <xsl:when test="./@cmpd">
              <xsl:value-of select="./@cmpd"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'sng'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="dashType">
          <xsl:choose>
            <xsl:when test="./a:prstDash/@val">
              <xsl:value-of select="./a:prstDash/@val"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'none'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:call-template name="LineTypeTransfer">
          <xsl:with-param name="lineType" select="$lineType"/>
          <xsl:with-param name="dashType" select="$dashType"/>
        </xsl:call-template>
        <xsl:if test="@w">
          <xsl:attribute name="宽度_C60F">
            <xsl:value-of select="@w div 12700"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
          <xsl:attribute name="颜色_C611">
            <xsl:call-template name="color"/>
          </xsl:attribute>
        </xsl:if>
        <!--share-->
      </图表:边框线_4226>
    </xsl:for-each>
    <!--数值-->
    <xsl:if test="c:numFmt">
      <图表:数值_E70D>
        <xsl:attribute name="是否链接到源_E73E">
          <xsl:choose>
            <xsl:when test="c:numFmt[not(@sourceLinked)] or c:numFmt[@sourceLinked='1']">
              <xsl:value-of select="'true'"/>
            </xsl:when>
            <xsl:when test="c:numFmt[@sourceLinked='0']">
              <xsl:value-of select="'false'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
          <xsl:attribute name="格式码_E73F">
              <xsl:choose>
                  <xsl:when test="contains(c:numFmt/@formatCode,'&quot;$&quot;#,##0_);\(&quot;$&quot;#,##0\)')">
                      <xsl:value-of select="'$#,##0_);($#,##0)'"/>
                  </xsl:when>
                  <xsl:otherwise>
                      <xsl:value-of select="c:numFmt/@formatCode"/>
                  </xsl:otherwise>
              </xsl:choose>
          </xsl:attribute>
          <xsl:attribute name="分类名称_E740">
              <xsl:choose>
                  <xsl:when test="contains(c:numFmt/@formatCode,'&quot;$&quot;#,##0_);\(&quot;$&quot;#,##0\)')">
                      <xsl:value-of select="'custom'"/>
                  </xsl:when>
                  <xsl:otherwise>
                      <xsl:value-of select="'custom'"/>
                  </xsl:otherwise>
              </xsl:choose>
          </xsl:attribute>          
      </图表:数值_E70D>
    </xsl:if>
    <!--字体-->
    <xsl:apply-templates select="./c:txPr">
      <xsl:with-param name="type" select="'valAx'"/>
      <xsl:with-param name="sstarget" select="$sstarget"/>
    </xsl:apply-templates>
    <!--<xsl:call-template name="FontTransfer">
      <xsl:with-param name="type" select="'valAx'"/>
      <xsl:with-param name="sstarget" select="$sstarget"/>
      <xsl:with-param name="target" select="$target"/>
    </xsl:call-template>-->
    <!--刻度-->
    <xsl:if test="c:scaling">
      <图表:刻度_E71D>
        <xsl:for-each select="c:scaling">
          <xsl:if test="c:min">
            <图表:最小值_E71E 是否自动_E71F="false">
              <xsl:value-of select="c:min/@val"/>
            </图表:最小值_E71E>
          </xsl:if>
          <xsl:if test="c:max">
            <图表:最大值_E720 是否自动_E71F="false">
              <xsl:value-of select="c:max/@val"/>
            </图表:最大值_E720>
          </xsl:if>
        </xsl:for-each>
        <xsl:if test="c:majorUnit">
          <图表:主单位_E721 是否自动_E71F="false">
            <xsl:value-of select="c:majorUnit/@val"/>
          </图表:主单位_E721>
        </xsl:if>
        <xsl:if test="c:minorUnit">
          <图表:次单位_E722 是否自动_E71F="false">
            <xsl:value-of select="c:minorUnit/@val"/>
          </图表:次单位_E722>
        </xsl:if>
        <!--交叉点-->
        <!--显示单位-->
        <xsl:if test="c:dispUnits">
          <图表:显示单位_E724>
            <xsl:if test="c:dispUnits[c:builtInUnit]">
              <xsl:attribute name="类型_E728">
                <xsl:choose>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='hundreds']">
                    <xsl:value-of select="'hundreds'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='thousands']">
                    <xsl:value-of select="'thousands'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='tenThousands']">
                    <xsl:value-of select="'ten thousands'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='hundredThousands']">
                    <xsl:value-of select="'one hundred thousands'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='millions']">
                    <xsl:value-of select="'millons'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='tenMillions']">
                    <xsl:value-of select="'ten millons'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='hundredMillions']">
                    <xsl:value-of select="'one hundred millons'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='trillions']">
                    <xsl:value-of select="'trillions'"/>
                  </xsl:when>
                  <xsl:when test="c:dispUnits/c:builtInUnit[@val='billions']">
                    <xsl:value-of select="'trillions'"/>
                  </xsl:when>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="c:dispUnits[c:dispUnitsLbl/c:layout]">
              <xsl:attribute name="是否显示_E727">
                <xsl:value-of select="'true'"/>
              </xsl:attribute>
              <!--表:显示单位 uof:locID="s0071" uof:attrList="值" 表:值="true"/-->
            </xsl:if>
          </图表:显示单位_E724>
        </xsl:if>
        <!--是否显示为对数刻度-->
        <xsl:if test="c:scaling">
          <xsl:if test="c:scaling[c:logBase]">
            <图表:是否显示为对数刻度_E729>
              <!--xsl:value-of select="c:scaling/c:logBase/@val"/-->
              <xsl:value-of select="'true'"/>
            </图表:是否显示为对数刻度_E729>
          </xsl:if>
        </xsl:if>

        <xsl:if test="c:scaling and c:scaling/c:orientation/@val='maxMin'">
          <图表:是否次序反转_E72B>
            <xsl:value-of select="'true'"/>
          </图表:是否次序反转_E72B>
        </xsl:if>
        <!--分类标签数-->
        <!--分类刻度数-->
        <!--是否交叉于最大值-->
        <!--数值轴是否置于分类轴之间-->
      </图表:刻度_E71D>
    </xsl:if>
    <!--对齐-->
    <xsl:if test="c:txPr or c:lblOffset">
      <xsl:if test="c:txPr[a:bodyPr] or c:lblOffset">
        <xsl:if test="c:txPr/a:bodyPr[@vert] or c:txPr/a:bodyPr[@rot] or c:lblOffset">
          <图表:对齐_E730>
            <xsl:if test="c:txPr/a:bodyPr[@vert]">
              <图表:文字排列方向_E703>
                <xsl:choose>
                  <xsl:when test="c:txPr/a:bodyPr[@vert='horz']">
                    <xsl:value-of select="'t2b-l2r-0e-0w'"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'r2l-t2b-0e-90w'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </图表:文字排列方向_E703>
            </xsl:if>
            <xsl:if test="c:txPr/a:bodyPr[@rot]">
              <图表:文字旋转角度_E704>
                <xsl:variable name="x">
                  <xsl:value-of select="c:txPr/a:bodyPr/@rot"/>
                </xsl:variable>
                <xsl:value-of select="($x div 60000) * (-1)"/>
              </图表:文字旋转角度_E704>
            </xsl:if>
            <xsl:if test="c:lblOffset">
              <xsl:if test="c:lblOffset[@val]">
                <图表:偏移量_E732>
                  <xsl:value-of select="c:lblOffset/@val"/>
                </图表:偏移量_E732>
              </xsl:if>
            </xsl:if>
          </图表:对齐_E730>
        </xsl:if>
      </xsl:if>
    </xsl:if>
    <!--网络线集-->
    <xsl:call-template name="chart-wgxj-val">
      <xsl:with-param name="tar" select="$tt"/>
    </xsl:call-template>
    <!--标题-->

      <!-- update by 凌峰 BUG_3080：坐标轴标题丢失，坐标刻度值变化 20140308 start -->
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart[c:plotArea/c:valAx[position()=$pos]/c:title]">
      <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx[position()=$pos]/c:title">
        <图表:标题_E736>
          <xsl:choose>
            <xsl:when test="descendant::a:t">
              <xsl:attribute name="名称_E742">
                <xsl:variable name="text">
                  <xsl:for-each select="descendant::a:t">
                    <xsl:value-of select="."/>
                  </xsl:for-each>
                </xsl:variable>
                <!--<xsl:value-of select="concat($text,'&#10;')"/>-->
                
                <!--zl 2015-4-28-->
                <xsl:value-of select="$text"/>
                <!--zl 2015-4-28-->
              </xsl:attribute>
            </xsl:when>

              <xsl:when test="./c:tx/c:strRef/c:strCache/c:pt/c:v">
                  <xsl:attribute name="名称_E742">
                      <xsl:value-of select="./c:tx/c:strRef/c:strCache/c:pt/c:v"/>
                  </xsl:attribute>
              </xsl:when>
              <!--end-->
              
            <xsl:otherwise>
              <xsl:attribute name="名称_E742">
                <xsl:variable name="text">
                  <xsl:value-of select="'坐标轴标题'"/>
                </xsl:variable>
                <xsl:value-of select="$text"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <!--边框+填充-->
          <xsl:if test="c:spPr">
            <xsl:for-each select="c:spPr">
              <xsl:call-template name="chart-biankuang">
                <xsl:with-param name="target2" select="$target"/>
              </xsl:call-template>
              <xsl:call-template name="chart-tianchong">
                <xsl:with-param name="target2" select="$target"/>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:if>
          <!--字体-->
          <xsl:if test="not(c:tx/c:rich/a:p) and not(c:txPr)">
            <图表:字体_E70B>
              <字:字体_4128>
                <xsl:attribute name="字号_412D">
                  <xsl:value-of select="'10'"/>
                </xsl:attribute>
              </字:字体_4128>
              <字:是否粗体_4130>true</字:是否粗体_4130>
            </图表:字体_E70B>
                <图表:对齐_E726>
                  <表:水平对齐方式_E700>center</表:水平对齐方式_E700>
                  <表:垂直对齐方式_E701>center</表:垂直对齐方式_E701>
                  <表:文字排列方向_E703>t2b-l2r-0e-0w</表:文字排列方向_E703>
                  <表:文字旋转角度_E704>90</表:文字旋转角度_E704>
                </图表:对齐_E726>
          </xsl:if>
          <xsl:if test="c:tx/c:rich/a:p">
            <xsl:for-each select="c:tx/c:rich/a:p">
              <图表:字体_E70B>
                <xsl:for-each select="a:r">
                  <xsl:choose>
                    <xsl:when test="./a:rPr">
                      <!--yanghaojie
                      <xsl:if test="./a:rPr[@i='1']">
                        <字:是否斜体_4131>true</字:是否斜体_4131>
                      </xsl:if>
                      -->
                      <字:字体_4128>
                        <xsl:if test="./a:rPr/a:latin">
                          <xsl:attribute name="西文字体引用_4129">
                            <xsl:value-of select="concat($sstarget,'.btj_szz_l_font')"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="./a:rPr/a:ea">
                          <xsl:attribute name="中文字体引用_412A">
                            <xsl:value-of select="concat($sstarget,'.btj_szz_e_font')"/>
                          </xsl:attribute>
                        </xsl:if>
                        <xsl:attribute name="字号_412D">
                          <xsl:choose>
                            <xsl:when test="a:rPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="a:rPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:rPr/@sz) and ./preceding-sibling::a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="./preceding-sibling::a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:rPr/@sz) and not(./preceding-sibling::a:pPr/a:defRPr/@sz) and  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select=" ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'10'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <xsl:choose>
                          <xsl:when test="a:rPr/a:solidFill">
                            <xsl:for-each select="a:rPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:when test="not(a:rPr/a:solidFill) and ./preceding-sibling::a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./preceding-sibling::a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:when test="not(a:rPr/a:solidFill) and not(./preceding-sibling::a:pPr/a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                        </xsl:choose>
                      </字:字体_4128>
                    </xsl:when>
                    <xsl:when test="not(./a:rPr[a:ea or a:latin]) and ./preceding-sibling::a:pPr/a:defRPr[a:ea or a:latin]">
                      <字:字体_4128>
                        <xsl:attribute name="西文字体引用_4129">
                          <xsl:value-of select="concat($sstarget,'.btj_szz_l_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="中文字体引用_412A">
                          <xsl:value-of select="concat($sstarget,'.btj_szz_e_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="字号_412D">
                          <xsl:choose>
                            <xsl:when test="a:rPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="a:rPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:rPr/@sz) and ./preceding-sibling::a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="./preceding-sibling::a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:rPr/@sz) and not(./preceding-sibling::a:pPr/a:defRPr/@sz) and  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select=" ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'10'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <xsl:choose>
                          <xsl:when test="a:rPr/a:solidFill">
                            <xsl:for-each select="a:rPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:when test="not(a:rPr/a:solidFill) and ./preceding-sibling::a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./preceding-sibling::a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:when test="not(a:rPr/a:solidFill) and not(./preceding-sibling::a:pPr/a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                        </xsl:choose>
                      </字:字体_4128>
                    </xsl:when>
                    <xsl:when test="not(./a:rPr[a:ea or a:latin]) and not (./preceding-sibling::a:pPr/a:defRPr[a:ea or a:latin]) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[a:ea or a:latin]">
                      <字:字体_4128>
                        <xsl:attribute name="西文字体引用_4129">
                          <xsl:value-of select="concat($sstarget,'.tbq_l_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="中文字体引用_412A">
                          <xsl:value-of select="concat($sstarget,'.tbq_e_font')"/>
                        </xsl:attribute>
                        <xsl:attribute name="字号_412D">
                          <xsl:choose>
                            <xsl:when test="a:rPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="a:rPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:rPr/@sz) and ./preceding-sibling::a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select="./preceding-sibling::a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:when test="not(a:rPr/@sz) and not(./preceding-sibling::a:pPr/a:defRPr/@sz) and  ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                              <xsl:variable name="sz">
                                <xsl:value-of select=" ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                              </xsl:variable>
                              <xsl:value-of select="$sz div 100"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="'10'"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <xsl:choose>
                          <xsl:when test="a:rPr/a:solidFill">
                            <xsl:for-each select="a:rPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:when test="not(a:rPr/a:solidFill) and ./preceding-sibling::a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./preceding-sibling::a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:when test="not(a:rPr/a:solidFill) and not(./preceding-sibling::a:pPr/a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                            <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr">
                              <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
                                <xsl:attribute name="颜色_412F">
                                  <xsl:call-template name="color"/>
                                </xsl:attribute>
                              </xsl:if>
                            </xsl:for-each>
                          </xsl:when>
                        </xsl:choose>
                      </字:字体_4128>
                    </xsl:when>
                  </xsl:choose>
                  <xsl:if test="./a:rPr">
                    <xsl:call-template name="table-btj-ziti"/>
                  </xsl:if>
                </xsl:for-each>
              </图表:字体_E70B>
              <xsl:for-each select="ancestor::c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich">
                <xsl:if test="a:bodyPr">
                  <xsl:call-template name="table-btj-dq"/>
                </xsl:if>
              </xsl:for-each>
            </xsl:for-each>
          </xsl:if>
          <xsl:for-each select="c:txPr">
            <xsl:for-each select="a:p/a:pPr">
              <xsl:choose>
                <xsl:when test="a:defRPr[a:ea or a:latin]">
                  <图表:字体_E70B>
                    <字:字体_4128>
                      <xsl:attribute name="西文字体引用_4129">
                        <xsl:value-of select="concat($sstarget,'.btj_szz_l_font')"/>
                      </xsl:attribute>
                      <xsl:attribute name="中文字体引用_412A">
                        <xsl:value-of select="concat($sstarget,'.btj_szz_e_font')"/>
                      </xsl:attribute>
                      <xsl:attribute name="字号_412D">
                        <xsl:choose>
                          <xsl:when test="a:defRPr/@sz">
                            <xsl:variable name="sz">
                              <xsl:value-of select="a:defRPr/@sz"/>
                            </xsl:variable>
                            <xsl:value-of select="$sz div 100"/>
                          </xsl:when>
                          <xsl:when test="not(a:defRPr/@sz) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                            <xsl:variable name="sz">
                              <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                            </xsl:variable>
                            <xsl:value-of select="$sz div 100"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="'10'"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:attribute>
                      <xsl:choose>
                        <xsl:when test="a:defRPr/a:solidFill/a:schemeClr or a:defRPr/a:solidFill/a:srgbClr">
                          <xsl:attribute name="颜色_412F">
                            <xsl:call-template name="color_ziti"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test="not(a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                          <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr">
                            <xsl:attribute name="颜色_412F">
                              <xsl:call-template name="color_ziti"/>
                            </xsl:attribute>
                          </xsl:for-each>
                        </xsl:when>
                      </xsl:choose>
                    </字:字体_4128>
                    <xsl:call-template name="table-tbziti"/>
                  </图表:字体_E70B>
                </xsl:when>
                <xsl:when test="not(a:defRPr[a:ea or a:latin]) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr[a:ea or a:latin]">
                  <图表:字体_E70B>
                    <字:字体_4128>
                      <xsl:attribute name="西文字体引用_4129">
                        <xsl:value-of select="concat($sstarget,'.tbq_l_font')"/>
                      </xsl:attribute>
                      <xsl:attribute name="中文字体引用_412A">
                        <xsl:value-of select="concat($sstarget,'.tbq_e_font')"/>
                      </xsl:attribute>
                      <xsl:attribute name="字号_412D">
                        <xsl:choose>
                          <xsl:when test="a:defRPr/@sz">
                            <xsl:variable name="sz">
                              <xsl:value-of select="a:defRPr/@sz"/>
                            </xsl:variable>
                            <xsl:value-of select="$sz div 100"/>
                          </xsl:when>
                          <xsl:when test="not(a:defRPr/@sz) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz">
                            <xsl:variable name="sz">
                              <xsl:value-of select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/@sz"/>
                            </xsl:variable>
                            <xsl:value-of select="$sz div 100"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select="'18'"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:attribute>
                      <xsl:choose>
                        <xsl:when test="a:defRPr/a:solidFill/a:schemeClr or a:defRPr/a:solidFill/a:srgbClr">
                          <xsl:attribute name="颜色_412F">
                            <xsl:call-template name="color_ziti"/>
                          </xsl:attribute>
                        </xsl:when>
                        <xsl:when test="not(a:defRPr/a:solidFill) and ./ancestor::c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill">
                          <xsl:for-each select="./ancestor::c:chartSpace/c:txPr/a:p/a:pPr">
                            <xsl:attribute name="颜色_412F">
                              <xsl:call-template name="color_ziti"/>
                            </xsl:attribute>
                          </xsl:for-each>
                        </xsl:when>
                      </xsl:choose>
                    </字:字体_4128>
                    <xsl:call-template name="table-tbziti"/>
                  </图表:字体_E70B>
                </xsl:when>
              </xsl:choose>
            </xsl:for-each>
            <xsl:call-template name="table-btj-dq-piechart">
              <xsl:with-param name="pos" select="$pos"/>
            </xsl:call-template>
          </xsl:for-each>
        </图表:标题_E736>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
  <!--网格线集-->
  <xsl:template name="chart-wgxj-val">
    <xsl:param name="tar"/>
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart[c:plotArea]">
      <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea">
        <xsl:if test="c:valAx[c:majorGridlines] or c:valAx[c:minorGridlines]">
          <图表:网格线集_E733>
            <xsl:if test="c:valAx[c:majorGridlines]">
              <图表:网格线_E734 位置_E735="major">

                <!-- update by 凌峰 网格线默认情况 2014.5.4 start -->
                <xsl:if test="not(c:valAx/c:majorGridlines[c:spPr])">
                  <xsl:attribute name="线型_C60D">
                    <xsl:value-of select="'single'"/>
                  </xsl:attribute>
                  <xsl:attribute name="虚实_C60E">
                    <xsl:value-of select="'solid'"/>
                  </xsl:attribute>
                  <xsl:attribute name="宽度_C60F">
                    <xsl:value-of select="'-1'"/>
                  </xsl:attribute>
                  <xsl:attribute name="颜色_C611">
                    <xsl:value-of select="'#808080'"/>
                  </xsl:attribute>
                </xsl:if>
                
                <xsl:if test="c:valAx/c:majorGridlines[c:spPr]">
                  <xsl:for-each select="c:valAx/c:majorGridlines/c:spPr">
                    <xsl:call-template name="wanggexian_border"/>
                  </xsl:for-each>
                  <!--<xsl:apply-templates select="c:valAx/c:majorGridlines/c:spPr"/>-->
                </xsl:if>
              </图表:网格线_E734>
            </xsl:if>
            <xsl:if test="c:valAx[c:minorGridlines]">
              <图表:网格线_E734 位置_E735="minor">
                
                <xsl:if test="not(c:valAx/c:minorGridlines[c:spPr])">
                  <xsl:attribute name="线型_C60D">
                    <xsl:value-of select="'single'"/>
                  </xsl:attribute>
                  <xsl:attribute name="虚实_C60E">
                    <xsl:value-of select="'solid'"/>
                  </xsl:attribute>
                  <xsl:attribute name="宽度_C60F">
                    <xsl:value-of select="'-1'"/>
                  </xsl:attribute>
                  <xsl:attribute name="颜色_C611">
                    <xsl:value-of select="'#808080'"/>
                  </xsl:attribute>
                </xsl:if>
                <!--end-->
                
                <xsl:if test="c:valAx/c:minorGridlines[c:spPr]">
                  <xsl:for-each select="c:valAx/c:minorGridlines/c:spPr">
                    <xsl:call-template name="wanggexian_border"/>
                  </xsl:for-each>
                </xsl:if>
              </图表:网格线_E734>
            </xsl:if>

          </图表:网格线集_E733>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
  <xsl:template name="chart-wgxj-cat">
    <xsl:param name="tar"/>
    <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart[c:plotArea]">
      <xsl:for-each select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($tar,'charts/')]/c:chart/c:plotArea">
        <xsl:if test="c:catAx[c:majorGridlines] or c:catAx[c:minorGridlines]">
          <图表:网格线集_E733>
            <xsl:if test="c:catAx[c:majorGridlines]">
              <图表:网格线_E734 位置_E735="major">
                
                <!-- update by 凌峰 网格线默认情况 2014.5.4 start -->
                <xsl:if test="not(c:catAx/c:majorGridlines[c:spPr])">
                  <xsl:attribute name="线型_C60D">
                    <xsl:value-of select="'single'"/>
                  </xsl:attribute>
                  <xsl:attribute name="虚实_C60E">
                    <xsl:value-of select="'solid'"/>
                  </xsl:attribute>
                  <xsl:attribute name="宽度_C60F">
                    <xsl:value-of select="'-1'"/>
                  </xsl:attribute>
                  <xsl:attribute name="颜色_C611">
                    <xsl:value-of select="'#808080'"/>
                  </xsl:attribute>
                </xsl:if>

                <xsl:if test="c:catAx/c:majorGridlines[c:spPr]">
                  <xsl:for-each select="c:catAx/c:majorGridlines/c:spPr">
                    <xsl:call-template name="wanggexian_border"/>
                  </xsl:for-each>
                  <!--<xsl:apply-templates select="c:catAx/c:majorGridlines/c:spPr"/>-->
                </xsl:if>
              </图表:网格线_E734>
            </xsl:if>
            <xsl:if test="c:catAx[c:minorGridlines]">
              <图表:网格线_E734 位置_E735="minor">
                
                <xsl:if test="not(c:catAx/c:majorGridlines[c:spPr])">
                  <xsl:attribute name="线型_C60D">
                    <xsl:value-of select="'single'"/>
                  </xsl:attribute>
                  <xsl:attribute name="虚实_C60E">
                    <xsl:value-of select="'solid'"/>
                  </xsl:attribute>
                  <xsl:attribute name="宽度_C60F">
                    <xsl:value-of select="'-1'"/>
                  </xsl:attribute>
                  <xsl:attribute name="颜色_C611">
                    <xsl:value-of select="'#808080'"/>
                  </xsl:attribute>
                </xsl:if>
                <!--end-->
                
                <xsl:if test="c:catAx/c:minorGridlines[c:spPr]">
                  <xsl:for-each select="c:catAx/c:minorGridlines/c:spPr">
                    <xsl:call-template name="wanggexian_border"/>
                  </xsl:for-each>
                  <!--<xsl:apply-templates select="c:catAx/c:majorGridlines/c:spPr"/>-->
                </xsl:if>
              </图表:网格线_E734>
            </xsl:if>
          </图表:网格线集_E733>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
  <xsl:template name="wanggexian_border">
    <!--a:ln部分的处理-->
    <xsl:for-each select="a:ln">
      <!--表:边框 uof:locID="s0057" uof:attrList="类型 宽度 边距 颜色 阴影"-->
      <!--线型-->
      <xsl:choose>
        <xsl:when test="@cmpd">
          <xsl:attribute name="线型_C60D">
            <xsl:value-of select="'single'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="biankuangsty">
            <xsl:with-param name="sty" select="a:prstDash/@val"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
      <!--宽度-->
      <xsl:if test="@w">
        <xsl:variable name="w">
          <xsl:value-of select="@w"/>
        </xsl:variable>
        <xsl:attribute name="宽度_C60F">
          <xsl:value-of select="@w div 12700"/>
        </xsl:attribute>
      </xsl:if>
      <!--颜色-->
      <xsl:if test="a:solidFill/a:srgbClr or a:solidFill/a:schemeClr">
        <xsl:attribute name="颜色_C611">
          <!--xsl:variable name="color">
	<xsl:value-of select="a:solidFill/a:srgbClr/@val"/>
	</xsl:variable>
	<xsl:value-of select="concat('#',$color)"/-->
          <!--<xsl:variable name="aaaaaaaaaaa">
            11111111111111111111111111
          </xsl:variable>-->
          <xsl:call-template name="color"/>
        </xsl:attribute>
      </xsl:if>
      <!--<xsl:if test="following-sibling::a:effectLst[a:innerShdw or a:outerShdw or a:prstShdw]">
        <xsl:attribute name="uof:阴影">
          <xsl:value-of select="'true'"/>
        </xsl:attribute>
      </xsl:if>-->
      <!--边距-->
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="biankuangsty">
    <xsl:param name="sty"/>
    <xsl:attribute name="线型_C60D">
      <xsl:choose>
        <xsl:when test="$sty='dash' or $sty='sysDash'">
          <xsl:value-of select="'dash'"/>
        </xsl:when>
        <xsl:when test="$sty='dashDot' or $sty='sysDashDot'">
          <xsl:value-of select="'dot-dash'"/>
        </xsl:when>
        <xsl:when test="$sty='dot' or $sty='sysDot'">
          <xsl:value-of select="'dash'"/>
        </xsl:when>
        <xsl:when test="$sty='lgDash'">
          <xsl:value-of select="'dash-long'"/>
        </xsl:when>
        <xsl:when test="$sty='lgDashDot'">
          <xsl:value-of select="'dot-dash'"/>
        </xsl:when>
        <xsl:when test="$sty='lgDashDotDot'">
          <xsl:value-of select="'dot-dash'"/>
        </xsl:when>
        <xsl:when test="$sty='solid'">
          <xsl:value-of select="'single'"/>
        </xsl:when>
        <xsl:when test="$sty='sysDashDotDot'">
          <xsl:value-of select="'dot-dash'"/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="'single'"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:attribute>
  </xsl:template>
  <!--渐变-->
  <xsl:template match="a:gradFill">
    <图:渐变_800D>
      <xsl:variable name="angle">
        <xsl:choose>
          <xsl:when test="a:lin">
            <xsl:value-of select="(360 - round(a:lin/@ang div 60000 div 45)*45+90) mod 360"/>
          </xsl:when>
          <xsl:when test ="a:tileRect">
            <xsl:value-of select ="315"/>
          </xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$angle='135' or $angle='180' or $angle='225' or $angle='270'">
          <xsl:for-each select="a:gsLst/a:gs">
            <xsl:sort select="@pos" data-type="number"/>
            <xsl:if test="position()=1">
              <xsl:attribute name="终止色_800F">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="position()=last()">
              <xsl:attribute name="起始色_800E">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select="a:gsLst/a:gs">
            <xsl:sort select="@pos" data-type="number"/>
            <xsl:if test="position()=1">
              <xsl:attribute name="起始色_800E">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="position()=last()">
              <xsl:attribute name="终止色_800F">
                <xsl:call-template name="colorChoice"/>
              </xsl:attribute>
            </xsl:if>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
      
      <xsl:attribute name="种子类型_8010">
        <xsl:choose>
          <xsl:when test="a:path">
            <xsl:choose>
              <xsl:when test="a:path/@path='rect'">rectangle</xsl:when>
              <xsl:when test="a:path/@path='circle'">oval</xsl:when>

              <xsl:when test="a:path/@path='shape'">rectangle</xsl:when>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>linear</xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      
      <xsl:attribute name="起始浓度_8011">100.0</xsl:attribute>
      <xsl:attribute name="终止浓度_8012">100.0</xsl:attribute>
      <xsl:attribute name="渐变方向_8013">
        <xsl:value-of select="$angle"/>
      </xsl:attribute>
      
      <xsl:attribute name="边界_8014">50</xsl:attribute>

      <xsl:choose>
        <xsl:when test="not(a:path)">
          <xsl:attribute name="种子X位置_8015">100</xsl:attribute>
          <xsl:attribute name="种子Y位置_8016">100</xsl:attribute>
        </xsl:when>
        <xsl:when test="a:path/@path='rect'">
          <xsl:choose>
            <xsl:when test="not(a:path/a:fillToRect/@l) and not(a:path/a:fillToRect/@t)">
              <xsl:attribute name="种子X位置_8015">30</xsl:attribute>
              <xsl:attribute name="种子Y位置_8016">30</xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:path/a:fillToRect/@l) and not(a:path/a:fillToRect/@b)">
              <xsl:attribute name="种子X位置_8015">30</xsl:attribute>
              <xsl:attribute name="种子Y位置_8016">60</xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:path/a:fillToRect/@r) and not(a:path/a:fillToRect/@t)">
              <xsl:attribute name="种子X位置_8015">60</xsl:attribute>
              <xsl:attribute name="种子Y位置_8016">30</xsl:attribute>
            </xsl:when>
            <xsl:when test="not(a:path/a:fillToRect/@r) and not(a:path/a:fillToRect/@b)">
              <xsl:attribute name="种子X位置_8015">100</xsl:attribute>
              <xsl:attribute name="种子Y位置_8016">100</xsl:attribute>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="种子X位置_8015">50</xsl:attribute>
          <xsl:attribute name="种子Y位置_8016">50</xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>

    </图:渐变_800D>
  </xsl:template>

</xsl:stylesheet>
