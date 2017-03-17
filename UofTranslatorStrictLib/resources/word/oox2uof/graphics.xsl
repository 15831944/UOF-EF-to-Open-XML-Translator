<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:dgm="http://purl.oclc.org/ooxml/drawingml/diagram"
                xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:fo="http://www.w3.org/1999/XSL/Format" 
                xmlns:xs="http://www.w3.org/2001/XMLSchema" 
                xmlns:fn="http://www.w3.org/2005/xpath-functions" 
                xmlns:xdt="http://www.w3.org/2005/xpath-datatypes"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:对象="http://schemas.uof.org/cn/2009/objects" 
                xmlns:图形="http://schemas.uof.org/cn/2009/graphics" 
                xmlns:图="http://schemas.uof.org/cn/2009/graph" 
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc" 
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet" 
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"  
                xmlns:扩展="http://schemas.uof.org/cn/2009/extend" 
                xmlns:规则="http://schemas.uof.org/cn/2009/rules" 
                xmlns:式样="http://schemas.uof.org/cn/2009/styles" 
                xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" 
                xmlns:o="urn:schemas-microsoft-com:office:office"
                xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" 
                xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math" 
                xmlns:v="urn:schemas-microsoft-com:vml"                
                xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing" 
                xmlns:w10="urn:schemas-microsoft-com:office:word" 
                xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" 
                xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" 
                xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships" 
                xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
                xmlns:a15="http://schemas.microsoft.com/office/drawing/2012/main"
                xmlns:pic="http://purl.oclc.org/ooxml/drawingml/picture" 
                xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" 
                xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<!--*****************cxl2011/11/14修改****************-->
	<xsl:template name="graphic">
		<xsl:param name="objFrom"/>
    <!--2013-04-19，wudi，增加对对象集-图片，图表对象的转换，增加对w:object节点的处理，start-->
		<xsl:for-each select="//w:pict|//w:drawing|//w:object">
      
      <!--2013-03-28，wudi，修复BUG#2745集成测试00X到UOF方向-模板1图片由一个变为二个，start-->
      <!--2013-04-08，wudi，修复图片转换的BUG，没有考虑节点为w:pict，子节点为v:shape的情况，所有number计数增加对v:shape节点的统计，start-->
      <xsl:if test ="not(ancestor::mc:Fallback) and not(ancestor::w:pict)">
        <xsl:choose>
          <xsl:when test="name(.)='w:pict'">

            <!--2014-03-26，wudi，模板pictObj增加一个参数pictObjFrom，start-->
            <xsl:call-template name="pictObj">
              <xsl:with-param name="pictObjFrom" select="$objFrom"/>
            </xsl:call-template>
            <!--end-->
            
          </xsl:when>
          <!--end-->
          <xsl:when test="name(.)='w:drawing'">
            <xsl:call-template name="drawing">
              <xsl:with-param name="drawingFrom" select="$objFrom"/>
            </xsl:call-template>
          </xsl:when>

          <!--2013-04-19，wudi，增加对对象集-图片，图表对象的转换，增加对w:object节点的处理，start-->
          <xsl:when test ="name(.)='w:object'">

            <!--2014-03-26，wudi，模板pictObj增加一个参数pictObjFrom，start-->
            <xsl:call-template name="pictObj">
              <xsl:with-param name="pictObjFrom" select="$objFrom"/>
            </xsl:call-template>
            <!--end-->
            
          </xsl:when>
          <!--end-->
          
        </xsl:choose>
      </xsl:if>
      <!--end-->
      
    </xsl:for-each>
	</xsl:template>

  <!--2013-04-08，wudi，修复图片转换的BUG，没有考虑节点为w:pict，子节点为v:shape的情况，start-->
  
  <!--2014-03-26，wudi，模板pictObj增加一个参数pictObjFrom，start-->
  <xsl:template name="pictObj">
    <xsl:param name="pictObjFrom"/>
    <xsl:for-each select="v:rect | v:shape">
      <xsl:choose>
        <xsl:when test="name(.)='v:rect'">
          
          <!--2014-03-26，wudi，模板rectObj增加一个参数rectObjFrom，start-->
          <xsl:call-template name="rectObj">
            <xsl:with-param name="type" select="'rect'"/>
            <xsl:with-param name="rectObjFrom" select="$pictObjFrom"/>
          </xsl:call-template>
          <!--end-->
          
        </xsl:when>
        <xsl:when test ="name(.)='v:shape'">

          <!--2014-03-26，wudi，模板rectObj增加一个参数rectObjFrom，start-->
          <xsl:call-template name="rectObj">
            <xsl:with-param name="type" select="'rect'"/>
            <xsl:with-param name="rectObjFrom" select="$pictObjFrom"/>
          </xsl:call-template>
          <!--end-->
          
        </xsl:when>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>
  <!--end-->

  <!--2014-03-26，wudi，模板rectObj增加一个参数rectObjFrom，start-->
	<xsl:template name="rectObj">
    <xsl:param name="type"/>
    <xsl:param name="rectObjFrom"/>
    <xsl:variable name="number">
      <xsl:number format="1" level="any" count="v:rect | wp:anchor| wp:inline | v:shape"/>
    </xsl:variable>
    <图:图形_8062 层次_8063="3">	 
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select ="concat($rectObjFrom,'Obj',$number * 2 +1)"/>
      </xsl:attribute>
      <xsl:variable name="style" select="@style"/>
      <图:预定义图形_8018>
        <图:类别_8019>
          <xsl:choose>
            <xsl:when test="$type='rect'">
              <xsl:value-of select="'11'"/>
            </xsl:when>
          </xsl:choose>
        </图:类别_8019>
        <图:名称_801A>
          <xsl:choose>
            <xsl:when test="$type='rect'">
              <xsl:value-of select="'Rectangle'"/>
            </xsl:when>
          </xsl:choose>
        </图:名称_801A>
        <图:生成软件_801B>UOFTranslator</图:生成软件_801B>
        <图:属性_801D>

          <!--2015-01-07，wudi，水印的转换，start-->
          <xsl:if test ="v:fill and name(.)='v:shape'">
            <图:透明度_8050>50</图:透明度_8050>
          </xsl:if>
          <!--end-->
          
          <!--2015-01-29，wudi，修改w-temp1和h-temp1的转换BUG，start-->
          <!--2014-03-26，wudi，修改w-temp1和h-temp1的取值方法，start-->
          <xsl:if test="@style">
            <xsl:if test="contains(@style,'width') or contains(@style,'WIDTH') or contains(@style,'height') or contains(@style,'HEIGHT')">
              <xsl:variable name="w-temp" select="substring-after($style,'width:')"/>

              <!--2014-03-24，wudi，考虑单位为in的情况，start-->
              <xsl:variable name="W-temp">
                <xsl:value-of select="substring($w-temp,1,15)"/>
              </xsl:variable>
              <xsl:variable name="w-temp1">
                <xsl:if test="contains($W-temp,'pt')">
                  <xsl:value-of select="substring-before($w-temp,'pt')"/>
                </xsl:if>
                <xsl:if test="contains($W-temp,'in')">
                  <xsl:value-of select="substring-before($w-temp,'in') * 72"/>
                </xsl:if>
              </xsl:variable>
              <!--<xsl:variable name="w-temp1" select="substring-before($w-temp,'pt;')"/>-->
              <!--end-->

              <xsl:variable name="h-temp" select="substring-after($style,'height:')"/>

              <!--2014-03-21，wudi，修改h-temp1的取值方式，考虑单位in和margin同时存在的情况，start-->
              <!--2013-04-19，wudi，针对对象集-图片，图表对象的转换，修改h-temp1的取值方式，start-->
              <xsl:variable name ="H-temp">
                <xsl:value-of select ="substring($h-temp,1,15)"/>
              </xsl:variable>
              <!--<xsl:variable name ="tmp">
                <xsl:value-of select ="string-length($H-temp)"/>
              </xsl:variable>
              <xsl:variable name ="tmp1">
                <xsl:value-of select ="substring($H-temp,$tmp - 1,2)"/>
              </xsl:variable>-->
              <xsl:variable name ="h-temp1">
                <!--2015-01-07，wudi，注释掉此部分代码，start-->
                <!--<xsl:if test ="$tmp1 ='pt'">
                  <xsl:value-of select ="substring-before($h-temp,'pt')"/>
                </xsl:if>-->
                <!--end-->

                <!--2014-03-13，wudi，存在以in为单位的情况，1in=72pt，start-->

                <!--2015-01-07，wudi，注释掉此部分代码，start-->
                <!--<xsl:if test ="$tmp1 ='in'">
                  <xsl:value-of select ="number(substring-before($h-temp,'in')) * 72"/>
                </xsl:if>-->
                <!--end-->

                <xsl:if test ="contains($H-temp,'pt')">
                  <xsl:value-of select ="substring-before($H-temp,'pt')"/>
                </xsl:if>

                <xsl:if test ="contains($H-temp,'in')">
                  <xsl:value-of select ="substring-before($H-temp,'in') * 72"/>
                </xsl:if>

                <!--end-->

              </xsl:variable>
              <!--end-->
              <!--end-->
              
              <图:大小_8060>
                <xsl:attribute name="宽_C605">
                  <xsl:value-of select="$w-temp1"/>
                </xsl:attribute>
                <xsl:attribute name="长_C604">
                  <xsl:value-of select="$h-temp1"/>
                </xsl:attribute>
              </图:大小_8060>                              
            </xsl:if>
            <xsl:if test="contains(@style,'rotation') or contains(@style,'ROTATION')">
              <xsl:variable name="r-temp" select="substring-after($style,'rotation:')"/>
              <xsl:variable name="r-temp1" select="substring-before($r-temp,';')"/>
              <图:旋转角度_804D>
                <xsl:value-of select="$r-temp1"/>
              </图:旋转角度_804D>
            </xsl:if>
          </xsl:if>
          <!--end-->
          <!--end-->
          <xsl:if test="@fillcolor">
            <图:填充 uof:locID="g0012">
              <图:颜色 uof:locID="g0034">
                <xsl:value-of select="@fillcolor"/>
              </图:颜色>
            </图:填充>
          </xsl:if>
          <图:线_8057>
            <xsl:if test="@strokecolor">
              <图:线颜色 uof:locID="g0013">
                <xsl:value-of select="@strokecolor"/>
              </图:线颜色>
            </xsl:if>
            <xsl:if test="@strokeweight">
              <图:线粗细_805C>
                <xsl:value-of select="@strokeweight"/>
              </图:线粗细_805C>
            </xsl:if>
            
            <!--2015-01-06，wudi，水印转换，start-->
            <xsl:if test="v:stroke">
              <图:线类型_8059>
                <xsl:attribute name="线型_805A">
                  <xsl:call-template name="lineType">
                    <xsl:with-param name="linetype" select="@dashstyle"/>
                  </xsl:call-template>
                </xsl:attribute>
                <xsl:attribute name="虚实_805B">
                  <xsl:value-of select="'solid'"/>
                </xsl:attribute>
              </图:线类型_8059>
            </xsl:if>
            <xsl:if test ="not(v:stroke and @strokeweight and @strokecolor)">
              <图:线类型_8059 线型_805A="single" 虚实_805B="solid"/>
              <图:线粗细_805C>0.75</图:线粗细_805C>
              <图:线颜色 uof:locID="g0013">#000000</图:线颜色>
            </xsl:if>
            <!--end-->
            
          </图:线_8057>
          <xsl:if test="o:lock[aspectratio='t']">
            <图:缩放是否锁定纵横比_8055>1</图:缩放是否锁定纵横比_8055>
          </xsl:if>
          <xsl:if test="@alt">
            <图:Web文字 uof:locID="g0033">
              <xsl:value-of select="@alt"/>
            </图:Web文字>
          </xsl:if>
          <图:是否打印对象_804E>true</图:是否打印对象_804E>
        </图:属性_801D>
      </图:预定义图形_8018>
      <xsl:if test="v:textbox and not(name(.)='v:shape')">
        <图:文本_803C>
          <图:对齐_803E>
            <xsl:if test="contains($style,'v-text-anchor')">
              <xsl:variable name="v-temp" select="substring-after($style,'v-text-anchor:')"/>
              <xsl:variable name="v-temp1" select="substring-before($v-temp,'pt;')"/>
              <xsl:attribute name="文字对齐_421E">
                <xsl:value-of select="$v-temp1"/>
              </xsl:attribute>
              <xsl:attribute name="水平对齐_421D">
                <xsl:value-of select="'left'"/>
              </xsl:attribute>
            </xsl:if>
          </图:对齐_803E>
          <图:边距_803D>
            <xsl:if test="v:textbox/@inset">
              <xsl:variable name="i-temp" select="v:textbox/@inset"/>
              <xsl:variable name="i-temp1" select="substring-before($i-temp,'pt;')"/>
              <xsl:variable name="i-tempL1" select="substring-after($i-temp,';')"/>
              <xsl:variable name="i-temp2" select="substring-before($i-tempL1,'pt;')"/>
              <xsl:variable name="i-tempL2" select="substring-after($i-tempL1,';')"/>
              <xsl:variable name="i-temp3" select="substring-before($i-tempL2,'pt;')"/>
              <xsl:variable name="i-tempL3" select="substring-after($i-tempL2,';')"/>
              <xsl:variable name="i-temp4" select="substring-before($i-tempL3,'pt;')"/>
              <xsl:variable name="i-tempL4" select="substring-after($i-tempL3,';')"/>
              <xsl:attribute name="左_C608">
                <xsl:value-of select="$i-temp1"/>
              </xsl:attribute>
              <xsl:attribute name="上_C609">
                <xsl:value-of select="$i-temp2"/>
              </xsl:attribute>
              <xsl:attribute name="右_C60A">
                <xsl:value-of select="$i-temp3"/>
              </xsl:attribute>
              <xsl:attribute name="下_C60B">
                <xsl:value-of select="$i-temp4"/>
              </xsl:attribute>
            </xsl:if>
          </图:边距_803D>
          <xsl:if test="v:textbox/@style">
            <xsl:variable name="textstyle" select="v:textbox/@style"/>
            <xsl:if test="contains($textstyle,'layout-flow')">
              <xsl:variable name="f-temp1" select="substring-after($textstyle,'layout-flow:')"/>
              <xsl:variable name="f-temp" select="substring-before($f-temp1,';')"/>
              <xsl:variable name="lf-temp">
                <xsl:if test="contains($textstyle,'mso-layout-flow-alt')">
                  <xsl:variable name="lf-temp1" select="substring-after($textstyle,'mso-layout-flow-alt:')"/>
                  <xsl:value-of select="substring-before($lf-temp1,';')"/>
                </xsl:if>
                <xsl:if test="not (contains($textstyle,'mso-layout-flow-alt'))">
                  <xsl:value-of select="'0'"/>
                </xsl:if>
              </xsl:variable>
              <图:文字排列方向_8042>
                <xsl:choose>
                  <xsl:when test="$f-temp='vertical' and $lf-temp='bottom-to-top'">
                    <xsl:value-of select="'vert-r2l'"/>
                  </xsl:when>
                  <xsl:when test="$f-temp='vertical-ideographic'">
                    <xsl:value-of select="'vert-l2r'"/>
                  </xsl:when>
                  <xsl:when test="$f-temp='horizontal-ideographic'">
                    <xsl:value-of select="'hori-r2l'"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="'hori-l2r'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </图:文字排列方向_8042>
            </xsl:if>
            <xsl:if test="not (contains($textstyle,'layout-flow'))">
              <图:文字排列方向_8042>
                <xsl:value-of select ="'hori-l2r'"/>
              </图:文字排列方向_8042>
            </xsl:if>
          </xsl:if>
          <xsl:for-each select="v:textbox/w:txbxContent/w:p">
            <xsl:call-template name="paragraph">
              <xsl:with-param name ="pPartFrom" select ="'document'"/>
            </xsl:call-template>
          </xsl:for-each>
        </图:文本_803C>
      </xsl:if>

      <xsl:if test ="v:textbox and (name(.)='v:shape')">
        <xsl:for-each select="v:textbox/w:txbxContent/w:p/w:r/w:pict/v:shape">
          <图:图片数据引用_8037>
            <!--2013-04-19，wudi，删除不必要的代码-->
            <xsl:value-of select="concat('document','Obj',($number + 1) * 2)"/>
          </图:图片数据引用_8037>
        </xsl:for-each>
      </xsl:if>

      <!--2015-01-06，wudi，水印转换，start-->
      <xsl:if test="v:textpath and name(.)='v:shape'">
        <图:文本_803C 是否自动换行_8047="false" 是否大小适应文字_8048="true" 是否文字随对象旋转_8049="true">
          <图:边距_803D 左_C608="7.2" 右_C60A="7.2" 上_C609="3.6" 下_C60B="3.6"/>
          <图:对齐_803E 水平对齐_421D="left" 文字对齐_421E="top"/>
          <图:文字排列方向_8042>t2b-l2r-0e-0w</图:文字排列方向_8042>
          <图:内容_8043>
            <字:段落_416B>
              <字:段落属性_419B>
                <字:孤行控制_418A>0</字:孤行控制_418A>
                <字:句属性_4158>
                  <字:字体_4128 是否西文绘制_412C="false">

                    <xsl:variable name="fc-temp1" select ="@fillcolor"/>
                    <xsl:variable name="fc-temp" select="substring-before($fc-temp1,' ')"/>
                    <xsl:attribute name="颜色_412F">
                      <xsl:value-of select="$fc-temp"/>
                    </xsl:attribute>

                    <xsl:variable name="textstyle" select="v:textpath/@style"/>
                    <xsl:if test ="contains($textstyle,'font-size')">
                      <xsl:variable name="fs-temp1" select="substring-after($textstyle,'font-size:')"/>
                      <xsl:variable name="fs-temp">
                        <xsl:if test="contains($fs-temp1,'in')">
                          <xsl:value-of select="number(substring-before($fs-temp1,'in')) * 72"/>
                        </xsl:if>
                        <xsl:if test="contains($fs-temp1,'pt')">
                          <xsl:value-of select="number(substring-before($fs-temp1,'pt'))"/>
                        </xsl:if>
                      </xsl:variable>
                      <xsl:attribute name="字号_412D">
                        <xsl:if test="$fs-temp1='1pt'">
                          <xsl:value-of select="110"/>
                        </xsl:if>
                        <xsl:if test="not($fs-temp1='1pt')">
                          <xsl:value-of select="$fs-temp"/>
                        </xsl:if>
                      </xsl:attribute>
                    </xsl:if>

                    <xsl:if test="contains($textstyle,'font-family')">
                      <xsl:variable name="ff-temp1" select="substring-after($textstyle,'font-family:&quot;')"/>
                      <xsl:variable name="ff-temp">
                        <xsl:value-of select="substring-before($ff-temp1,'&quot;')"/>
                      </xsl:variable>
                      <xsl:attribute name="西文字体引用_4129">
                        <xsl:value-of select="$ff-temp"/>
                      </xsl:attribute>
                      <xsl:attribute name="中文字体引用_412A">
                        <xsl:value-of select="$ff-temp"/>
                      </xsl:attribute>
                    </xsl:if>

                  </字:字体_4128>
                  <字:是否阴影_4140>false</字:是否阴影_4140>
                </字:句属性_4158>
              </字:段落属性_419B>
              <字:句_419D>
                <字:句属性_4158>
                  <字:字体_4128 是否西文绘制_412C="false">
                    <xsl:variable name="fc-temp1" select ="@fillcolor"/>
                    <xsl:variable name="fc-temp" select="substring-before($fc-temp1,' ')"/>
                    <xsl:attribute name="颜色_412F">
                      <xsl:value-of select="$fc-temp"/>
                    </xsl:attribute>

                    <xsl:variable name="textstyle" select="v:textpath/@style"/>
                    <xsl:if test ="contains($textstyle,'font-size')">
                      <xsl:variable name="fs-temp1" select="substring-after($textstyle,'font-size:')"/>
                      <xsl:variable name="fs-temp">
                        <xsl:if test="contains($fs-temp1,'in')">
                          <xsl:value-of select="number(substring-before($fs-temp1,'in')) * 72"/>
                        </xsl:if>
                        <xsl:if test="contains($fs-temp1,'pt')">
                          <xsl:value-of select="number(substring-before($fs-temp1,'pt'))"/>
                        </xsl:if>
                      </xsl:variable>
                      <xsl:attribute name="字号_412D">
                        <xsl:if test="$fs-temp1='1pt'">
                          <xsl:value-of select="110"/>
                        </xsl:if>
                        <xsl:if test="not($fs-temp1='1pt')">
                          <xsl:value-of select="$fs-temp"/>
                        </xsl:if>
                      </xsl:attribute>
                    </xsl:if>
                    <xsl:if test="not(contains($textstyle,'font-size'))">
                      <xsl:attribute name="字号_412D">
                        <xsl:value-of select="36"/>
                      </xsl:attribute>
                    </xsl:if>

                    <xsl:if test="contains($textstyle,'font-family')">
                      <xsl:variable name="ff-temp1" select="substring-after($textstyle,'font-family:&quot;')"/>
                      <xsl:variable name="ff-temp">
                        <xsl:value-of select="substring-before($ff-temp1,'&quot;')"/>
                      </xsl:variable>
                      <xsl:attribute name="西文字体引用_4129">
                        <xsl:value-of select="$ff-temp"/>
                      </xsl:attribute>
                      <xsl:attribute name="中文字体引用_412A">
                        <xsl:value-of select="$ff-temp"/>
                      </xsl:attribute>
                    </xsl:if>

                  </字:字体_4128>
                  <字:是否阴影_4140>false</字:是否阴影_4140>
                </字:句属性_4158>
                <字:文本串_415B>
                  <xsl:value-of select="v:textpath/@string"/>
                </字:文本串_415B>
              </字:句_419D>
            </字:段落_416B>
          </图:内容_8043>
        </图:文本_803C>
      </xsl:if>
      <!--end-->
      
      <!--2014-05-20，wudi，修改document为$rectObjFrom，start-->
      <!--2013-04-19，wudi，增加对对象集-图片，图表对象的处理，start-->
      <xsl:if test ="name(.)='v:shape' and (parent::w:object)">
        <图:图片数据引用_8037>
          <xsl:value-of select="concat($rectObjFrom,'Obj',$number * 2)"/>
        </图:图片数据引用_8037>
      </xsl:if>
      <!--end-->
      <!--end-->

      <!--2014-05-20，wudi，修复水印转成图片的问题，start-->
      <!--2014-03-26，wudi，图片数据引用：增加w:pict节点的情况，start-->
      <xsl:if test ="name(.)='v:shape' and (parent::w:pict) and not(ancestor::w:sdt/w:sdtPr/w:docPartObj/w:docPartGallery/@w:val = 'Watermarks')">
        <图:图片数据引用_8037>
          <xsl:value-of select="concat($rectObjFrom,'Obj',$number * 2)"/>
        </图:图片数据引用_8037>
      </xsl:if>
      <!--end-->
      <!--end-->
      
    </图:图形_8062>
  </xsl:template>
  <!--end-->
  
  <!--end-->
  
	<xsl:template name="drawing">
	  <xsl:param name="drawingFrom"/>
	  <xsl:for-each select="wp:anchor | wp:inline">
			<xsl:choose>
				<xsl:when test="name(.)='wp:anchor'">

          <!--2013-11-05，wudi，Strict标准下为wp前缀，添加此前缀，start-->
					<xsl:if test="./a:graphic/a:graphicData/pic:pic or ./a:graphic/a:graphicData/wp:wsp">
						<xsl:apply-templates select="a:graphic/a:graphicData/pic:pic|./a:graphic/a:graphicData/wp:wsp">
							<xsl:with-param name="picType" select="'anchor'"/>
							<xsl:with-param name="picFrom" select="$drawingFrom"/>
						</xsl:apply-templates>
					</xsl:if>
          <!--end-->
          
					<!-- 组合图形-->
          <!--2014-03-06，wudi，修改前缀wpg为wp，start-->
					<xsl:if test="./a:graphic/a:graphicData/wp:wgp">
						<xsl:apply-templates select="a:graphic/a:graphicData/wp:wgp">
							<xsl:with-param name="picType" select="'anchor'"/>
							<xsl:with-param name="picFrom" select="$drawingFrom"/>
						</xsl:apply-templates>
					</xsl:if>
          <!--end-->
          
          <!--2012-01-14，wudi，OOX到UOF方向SmartArt转换，start-->
          <xsl:if test ="./a:graphic/a:graphicData/dgm:relIds">
            <xsl:variable name ="rcs" select ="./a:graphic/a:graphicData/dgm:relIds/@r:cs"/>
            <xsl:variable name ="rId" select ="substring-after($rcs,'rId')"/>
            <xsl:variable name ="gId" select ="number($rId) + 1"/>
            <xsl:variable name ="number">
              <xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline | v:shape"/>
            </xsl:variable>
            <xsl:variable name ="ExtCx" select ="number(./wp:extent/@cx) div 12700"/>
            <xsl:variable name ="ExtCy" select ="number(./wp:extent/@cy) div 12700"/>
            <xsl:apply-templates select ="document('word/_rels/document.xml.rels')/rel:Relationships" mode ="SmartArt">
              <xsl:with-param name ="GId" select ="$gId"/>
              <xsl:with-param name="picType" select="'anchor'"/>
              <xsl:with-param name="picFrom" select="$drawingFrom"/>
              <xsl:with-param name ="number" select ="$number"/>
              <xsl:with-param name ="ExtCx" select ="$ExtCx"/>
              <xsl:with-param name ="ExtCy" select ="$ExtCy"/>
            </xsl:apply-templates>
          </xsl:if>
          <!--end-->
          
					<!--<xsl:call-template name="picture">
            <xsl:with-param name="picType" select="'anchor'"/>
            <xsl:with-param name="picFrom" select="$drawingFrom"/>
          </xsl:call-template>-->
				</xsl:when>
				<xsl:when test="name(.)='wp:inline'">

          <!--2013-11-06，wudi，Strict标准下为wp前缀，添加此前缀，start-->
					<xsl:if test="./a:graphic/a:graphicData/pic:pic or ./a:graphic/a:graphicData/wp:wsp">
						<xsl:apply-templates select="a:graphic/a:graphicData/pic:pic|./a:graphic/a:graphicData/wp:wsp">
							<xsl:with-param name="picType" select="'inline'"/>
							<xsl:with-param name="picFrom" select="$drawingFrom"/>
						</xsl:apply-templates>
					</xsl:if>
          <!--end-->
          
					<!-- 组合图形-->
          <!--2014-03-06，wudi，修改前缀wpg为wp，start-->
					<xsl:if test="./a:graphic/a:graphicData/wp:wgp">
					<xsl:apply-templates select="a:graphic/a:graphicData/wp:wgp">
						<xsl:with-param name="picType" select="'inline'"/>
						<xsl:with-param name="picFrom" select="$drawingFrom"/>
					</xsl:apply-templates>
					</xsl:if>
          <!--end-->
          
          <!--2012-01-14，wudi，OOX到UOF方向SmartArt转换，start-->
          <xsl:if test ="./a:graphic/a:graphicData/dgm:relIds">
            <xsl:variable name ="rcs" select ="./a:graphic/a:graphicData/dgm:relIds/@r:cs"/>
            <xsl:variable name ="rId" select ="substring-after($rcs,'rId')"/>
            <xsl:variable name ="gId" select ="number($rId) + 1"/>
            <xsl:variable name ="number">
              <xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline | v:shape"/>
            </xsl:variable>
            <xsl:variable name ="ExtCx" select ="number(./wp:extent/@cx) div 12700"/>
            <xsl:variable name ="ExtCy" select ="number(./wp:extent/@cy) div 12700"/>
            <xsl:apply-templates select ="document('word/_rels/document.xml.rels')/rel:Relationships" mode ="SmartArt">
              <xsl:with-param name ="GId" select ="$gId"/>
              <xsl:with-param name="picType" select="'inline'"/>
              <xsl:with-param name="picFrom" select="$drawingFrom"/>
              <xsl:with-param name ="number" select ="$number"/>
              <xsl:with-param name ="ExtCx" select ="$ExtCx"/>
              <xsl:with-param name ="ExtCy" select ="$ExtCy"/>
            </xsl:apply-templates>
          </xsl:if>
          <!--end-->

          <!--2013-04-19，wudi，增加对chart的转换，start-->
          <xsl:if test ="./a:graphic/a:graphicData/c:chart">
            <xsl:call-template name ="chartObj">
              <xsl:with-param name="picType" select="'inline'"/>
              <xsl:with-param name="picFrom" select="$drawingFrom"/>
            </xsl:call-template>
          </xsl:if>
          <!--end-->
          
					<!--<xsl:call-template name="picture">
            <xsl:with-param name="picType" select="'inline'"/>
            <xsl:with-param name="picFrom" select="$drawingFrom"/>
          </xsl:call-template>-->
				</xsl:when>
		  </xsl:choose>
	  </xsl:for-each>
	</xsl:template>

  <!--2013-04-19，wudi，增加对chart的转换，start-->
  <xsl:template name ="chartObj">
    <xsl:param name ="picType"/>
    <xsl:param name ="picFrom"/>
    <xsl:variable name="number">
      <xsl:number format="1" level="any" count="v:rect | wp:anchor| wp:inline | v:shape"/>
    </xsl:variable>
    <图:图形_8062 层次_8063="3">
      <xsl:attribute name="标识符_804B">
        <xsl:value-of select ="concat('document','Obj',$number * 2 +1)"/>
      </xsl:attribute>
      <xsl:variable name="style" select="@style"/>
      <图:预定义图形_8018>
        <图:类别_8019>
          <xsl:value-of select="'11'"/>
        </图:类别_8019>
        <图:名称_801A>
          <xsl:value-of select="'Rectangle'"/>
        </图:名称_801A>
        <图:生成软件_801B>UOFTranslator</图:生成软件_801B>
        <图:属性_801D>
          <图:大小_8060>
            <xsl:attribute name="宽_C605">
              <xsl:value-of select="wp:extent/@cx div 12700"/>
            </xsl:attribute>
            <xsl:attribute name="长_C604">
              <xsl:value-of select="wp:extent/@cy div 12700"/>
            </xsl:attribute>
          </图:大小_8060>
          <图:线_8057>
            <图:线类型_8059>single</图:线类型_8059>
            <图:线粗细_805C>0.75</图:线粗细_805C>
            <图:线颜色 uof:locID="g0013">#000000</图:线颜色>
          </图:线_8057>
          <图:是否打印对象_804E>true</图:是否打印对象_804E>
        </图:属性_801D>
      </图:预定义图形_8018>
      <图:图片数据引用_8037>
        <xsl:value-of select="concat('document','Obj',$number * 2)"/>
      </图:图片数据引用_8037>
    </图:图形_8062>
  </xsl:template>
  <!--end-->
  
  <!--2012-01-14，wudi，OOX到UOF方向SmartArt转换，start-->
  <xsl:template match ="rel:Relationships" mode ="SmartArt">
    <xsl:param name ="GId"/>
    <xsl:param name ="picType"/>
    <xsl:param name ="picFrom"/>
    <xsl:param name ="number"/>
    <xsl:param name ="ExtCx"/>
    <xsl:param name ="ExtCy"/>
    <xsl:for-each select ="rel:Relationship">
      <xsl:if test ="@Id = concat('rId',$GId)">
        <xsl:apply-templates select ="document(concat('word/',@Target))/dsp:drawing/dsp:spTree" mode ="SmartArt">
          <xsl:with-param name ="PicType" select ="$picType"/>
          <xsl:with-param name ="PicFrom" select ="$picFrom"/>
          <xsl:with-param name ="number" select ="$number"/>
          <xsl:with-param name ="ExtCx" select ="$ExtCx"/>
          <xsl:with-param name ="ExtCy" select ="$ExtCy"/>
        </xsl:apply-templates>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
  <xsl:template match ="dsp:spTree" mode ="SmartArt">
    <xsl:param name ="PicType"/>
    <xsl:param name ="PicFrom"/>
    <xsl:param name ="number"/>
    <xsl:param name ="ExtCx"/>
    <xsl:param name ="ExtCy"/>
    <xsl:for-each select ="dsp:sp">
      <xsl:variable name ="num">
        <xsl:value-of select ="count(preceding-sibling::dsp:sp)"/>
      </xsl:variable>
      
      <!--2013-05-03，wudi，修改PicFrom为SmartArt，区分SmartArt与组合图形，流程图等，start-->
      <xsl:apply-templates select=".">
        <xsl:with-param name="PicType" select="'anchor'"/>
        <xsl:with-param name="PicFrom" select="'SmartArt'"/>
        <xsl:with-param name ="number" select ="$number"/>
        <xsl:with-param name ="num" select ="$num"/>
      </xsl:apply-templates>
      <!--end-->
      
    </xsl:for-each>
    <图:图形_8062>
      <xsl:attribute name="层次_8063">
        <xsl:if test="ancestor::wp:anchor/@relativeHeight">
          <xsl:value-of select="ancestor::wp:anchor/@relativeHeight"/>
        </xsl:if>
        <xsl:if test="not(ancestor::wp:anchor/@relativeHeight)">
          <xsl:value-of select="'3'"/>
        </xsl:if>
      </xsl:attribute>
      
      <!--2013-05-03，wudi，修改grpspObj为SmartArtObj，区分SmartArt与组合图形，流程图等，start-->
      <xsl:attribute name="标识符_804B">
        <xsl:if test="$PicFrom ='grpsp'">
          <xsl:value-of select="concat('SmartArtObj',./wpg:cNvPr/@id)"/>
        </xsl:if>
        <xsl:if test="$PicFrom !='grpsp'">
          <xsl:value-of select="concat($PicFrom,'Obj',$number * 2 + 1)"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name="组合列表_8064">
        <xsl:for-each select="dsp:sp">
          <xsl:variable name ="num1">
            <xsl:value-of select ="count(preceding-sibling::dsp:sp)"/>
          </xsl:variable>
            <xsl:value-of select="concat(' ','SmartArtObj',($number - 1) * 9 + $num1)"/>
        </xsl:for-each>
      </xsl:attribute>
      <!--end-->
      
      <图:预定义图形_8018>
        <图:类别_8019>11</图:类别_8019>
        <图:名称_801A>Rectangle</图:名称_801A>
        <图:生成软件_801B>UOFTranslator</图:生成软件_801B>
        <图:属性_801D>
          <图:大小_8060>
            <xsl:attribute name="宽_C605">
              <xsl:value-of select="$ExtCx"/>
            </xsl:attribute>
            <xsl:attribute name="长_C604">
              <xsl:value-of select="$ExtCy"/>
            </xsl:attribute>
          </图:大小_8060>
          <xsl:choose>
            <xsl:when test="./dsp:spPr/a:xfrm/@rot">
              <图:旋转角度_804D>
                <xsl:value-of select="./dsp:spPr/a:xfrm/@rot div 60000"/>
              </图:旋转角度_804D>
            </xsl:when>
            <xsl:otherwise>
              <图:旋转角度_804D>0.0</图:旋转角度_804D>
            </xsl:otherwise>
          </xsl:choose>
          <图:缩放是否锁定纵横比_8055>
            <xsl:if test="./dsp:nvSpPr/a:grpSpLocks/@noChangeAspect='1'">
              <xsl:value-of select=" 'true'"/>
            </xsl:if>
            <xsl:if test="./dsp:nvSpPr/a:grpSpLocks/@noChangeAspect='0' or not(./dsp:nvSpPr/a:grpSpLocks/@noChangeAspect)">
              <xsl:value-of select="'false'"/>
            </xsl:if>
          </图:缩放是否锁定纵横比_8055>
          <图:是否打印对象_804E>true</图:是否打印对象_804E>
        </图:属性_801D>
      </图:预定义图形_8018>
    </图:图形_8062>
  </xsl:template>
  <xsl:template match ="dsp:sp">
    <xsl:param name="PicFrom"/>
    <xsl:param name ="number"/>
    <xsl:param name ="num"/>
    <xsl:variable name="style" select="@style"/>
    <图:图形_8062>
      <!--预定义图形处理-->
      <xsl:attribute name="标识符_804B">
          <xsl:value-of select="concat($PicFrom,'Obj',($number - 1) * 9 + $num)"/>
      </xsl:attribute>
      <xsl:attribute name="层次_8063">
        <!--cxl,2012.5.14图形叠放次序，暂时先这样处理，效果上没问题-->
        <xsl:if test="ancestor::wp:anchor/@relativeHeight">
          <xsl:value-of select="ancestor::wp:anchor/@relativeHeight"/>
        </xsl:if>
        <xsl:if test="not(ancestor::wp:anchor/@relativeHeight)">
          <xsl:value-of select="'3'"/>
        </xsl:if>
      </xsl:attribute>
      <图:预定义图形_8018>
        <图:类别_8019>
          <xsl:choose>
            <!--cxl2011/11/21增加了几种漏转的图形（包括七边形、乘号、六角星、七角星、10角星、16角星），与差异文档做了对应关系的修改（包括减号、L形）-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rect'">11</xsl:when>
            <!--矩形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='parallelogram'">12</xsl:when>
            <!--平行四边形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='trapezoid'">13</xsl:when>
            <!--梯形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='diamond'">14</xsl:when>
            <!--菱形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='roundRect'">15</xsl:when>
            <!--圆角矩形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='heptagon'">16</xsl:when>
            <!--七边形转换为八边形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='octagon'">16</xsl:when>
            <!--八边形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='triangle'">17</xsl:when>
            <!--三角形17代表等腰三角形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rtTriangle'">18</xsl:when>
            <!--右三角形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='ellipse'">19</xsl:when>
            <!--椭圆-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='diagStripe'">18</xsl:when>
            <!--斜纹-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='corner'">11</xsl:when>
            <!--L形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='halfFrame'">18</xsl:when>
            <!--半闭框-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='decagon'">19</xsl:when>
            <!--十边形转换为圆-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='dodecagon'">19</xsl:when>
            <!--十二边形转换为圆-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='pie'">19</xsl:when>
            <!--饼形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='chord'">19</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='teardrop'">19</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='cloud'">19</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rightArrow'">21</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftArrow'">22</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='upArrow'">23</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='downArrow'">24</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftRightArrow'">25</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='upDownArrow'">26</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='quadArrow'">27</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftRightUpArrow'">28</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bentArrow'">29</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartProcess'">31</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartAlternateProcess'">32</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartDecision'">33</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartInputOutput'">34</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartPredefinedProcess'">35</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartInternalStorage'">36</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartDocument'">37</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartMultidocument'">38</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartTerminator'">39</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star4'">43</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star5'">44</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star6'">45</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star7'">45</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star8'">45</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star10'">46</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star12'">46</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star16'">46</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star24'">47</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star32'">48</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='ribbon'">410</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='cloudCallout'">19</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='line'">61</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='straightConnector1'">71</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='hexagon'">110</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='pentagon'">112</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='can'">113</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='cube'">114</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bevel'">115</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='foldedCorner'">116</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='smileyFace'">117</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='donut'">118</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='blockArc'">120</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='heart'">121</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='lightningBolt'">122</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='sun'">123</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='moon'">124</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='arc'">125</xsl:when>
            <!--弧形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bracketPair'">126</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bracePair'">127</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='plaque'">128</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftBracket'">129</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rightBracket'">130</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftBrace'">131</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rightBrace'">132</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='uturnArrow'">210</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftUpArrow'">211</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bentUpArrow'">212</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='curvedRightArrow'">213</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='curvedLeftArrow'">214</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='curvedUpArrow'">215</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='curvedDownArrow'">216</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='stripedRightArrow'">217</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='notchedRightArrow'">218</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='chevron'">220</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rightArrowCallout'">221</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftArrowCallout'">222</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='upArrowCallout'">223</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='downArrowCallout'">224</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftRightArrowCallout'">225</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='upDownArrowCallout'">226</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='circularArrow'">228</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartPreparation'">310</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartManualInput'">311</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartManualOperation'">312</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartConnector'">313</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartOffpageConnector'">314</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartPunchedCard'">315</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartPunchedTape'">316</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartSummingJunction'">317</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartOr'">318</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartCollate'">319</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartSort'">320</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartExtract'">321</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartMerge'">322</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartDelay'">324</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartMagneticDisk'">326</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartDisplay'">328</xsl:when>
            <!--<xsl:when test="a:prstGeom/@prst='ribbon2'">410</xsl:when>-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='ribbon2'">49</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='ellipseRibbon'">412</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='ellipseRibbon2'">411</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='verticalScroll'">413</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='horizontalScroll'">414</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='wave'">415</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='doubleWave'">416</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='noSmoking'">119</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='plus'">111</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='mathPlus'">111</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='mathMinus'">11</xsl:when>
            <!--减号转成矩形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='mathMultiply'">111</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='mathDivide'">111</xsl:when>
            <!--xsl:when test="a:prstGeom/@prst='straightConnector1'">71</xsl:when-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='irregularSeal1'">41</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='irregularSeal2'">42</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='wedgeRectCallout'">51</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='wedgeRoundRectCallout'">52</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='wedgeEllipseCallout'">53</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='borderCallout1'">56</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='borderCallout2'">57</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='borderCallout3'">58</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='accentCallout1'">510</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='accentCallout2'">511</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='accentCallout3'">512</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='callout1'">514</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='callout2'">515</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='callout3'">516</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='accentBorderCallout1'">518</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='accentBorderCallout2'">519</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='accentBorderCallout3'">520</xsl:when>
            <!--xsl:when test="a:prstGeom/@prst='straightConnector1'">62</xsl:when-->
            <!--<xsl:when test="./wps:spPr/a:prstGeom/@prst='curvedConnector3'">64</xsl:when>-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bentConnector3'">74</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='curvedConnector3'">77</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='homePlate'">219</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='quadArrowCallout'">227</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartOnlineStorage'">323</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartMagneticTape'">325</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartMagneticDrum'">327</xsl:when>
            <xsl:otherwise>11</xsl:otherwise>
          </xsl:choose>
        </图:类别_8019>
        <图:名称_801A>
          <xsl:choose>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rect'">Rectangle</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='parallelogram'">Parallelogram</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='trapezoid'">Trapezoid</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='diamond'">Diamond</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='roundRect'">Rounded Rectangle</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='heptagon'">Octagon</xsl:when>
            <!--七边形转换为八边形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='octagon'">Octagon</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='triangle'">Isosceles Triangle</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rtTriangle'">Right Triangle</xsl:when>
            <!--斜纹，半闭框，L形转换为三角形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='diagStripe'">Right Triangle</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='corner'">Rectangle</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='halfFrame'">Right Triangle</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='ellipse'">Oval</xsl:when>
            <!--2011：十边形，十二边形，饼，泪滴，弦月，云，转换为圆-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='decagon'">Oval</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='dodecagon'">Oval</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='pie'">Oval</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='chord'">Oval</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='teardrop'">Oval</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='cloud'">Oval</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rightArrow'">Right Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftArrow'">Left Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='upArrow'">Up Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='downArrow'">Down Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftRightArrow'">Left-Right Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='upDownArrow'">Up-Down Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='quadArrow'">Quad Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftRightUpArrow'">Left-Right-Up Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bentArrow'">Bent Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartProcess'">Process</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartAlternateProcess'">Alternate Process</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartDecision'">Decision</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartProcess'">Process</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartPredefinedProcess'">Predefined Process</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartInputOutput'">Data</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartInternalStorage'">Internal Storage</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartDocument'">Document</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartMultidocument'">Multidocument</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartTerminator'">Terminator</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star4'">4-Point Star</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star5'">5-Point Star</xsl:when>
            <!--2011：6、7角形转换为8角形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star6'">8-Point Star</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star7'">8-Point Star</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star8'">8-Point Star</xsl:when>
            <!--2011-：10、12角形转换为16角形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star10'">16-Point Star</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star12'">16-Point Star</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star16'">16-Point Star</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star24'">24-Point Star</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='star32'">32-Point Star</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='ribbon'">Down Ribbon</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='cloudCallout'">Oval</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='line'">Line</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='straightConnector1'">Straight Connector</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='hexagon'">Hexagon</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='pentagon'">Regual Pentagon</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='can'">Can</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='cube'">Cube</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bevel'">Bevel</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='foldedCorner'">Folded Corner</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='smileyFace'">Smiley Face</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='donut'">Donut</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='blockArc'">Block Arc</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='heart'">Heart</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='lightningBolt'">Lightning</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='sun'">Sun</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='moon'">Moon</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='arc'">Arc</xsl:when>
            <!--弧形-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bracketPair'">Double Bracket</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bracePair'">Double Brace</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='plaque'">Plaque</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftBracket'">Left Bracket</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rightBracket'">Right Bracket</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftBrace'">Left Brace</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rightBrace'">Right Brace</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='uturnArrow'">U-Turn Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftUpArrow'">Left-Up Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bentUpArrow'">Bent-Up Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='curvedRightArrow'">Curved Right Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='curvedLeftArrow'">Curved Left Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='curvedUpArrow'">Curved Up Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='curvedDownArrow'">Curved Down Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='stripedRightArrow'">Striped Right Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='notchedRightArrow'">Notched Right Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='chevron'">Chevron Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rightArrowCallout'">Right Arrow Callout</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftArrowCallout'">Left Arrow Callout</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='upArrowCallout'">Up Arrow Callout</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='downArrowCallout'">Down Arrow Callout</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftRightArrowCallout'">Left-Right Arrow Callout</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='upDownArrowCallout'">Up-Down Arrow Callout</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='circularArrow'">Circular Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartPreparation'">Preparation</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartManualInput'">Manual Input</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartManualOperation'">Manual Operation</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartConnector'">Connector</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartOffpageConnector'">Off-page Connector</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartPunchedCard'">Card</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartPunchedTape'">Punched Tape</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartSummingJunction'">Summing Junction</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartOr'">Or</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartCollate'">Collate</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartSort'">Sort</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartExtract'">Extract</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartMerge'">Merge</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartDelay'">Delay</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartMagneticDisk'">Magnetic Disk</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartDisplay'">Display</xsl:when>
            <!--zhaobj-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='ribbon2'">Up Ribbon</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='ellipseRibbon'">Curved Down Ribbon</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='ellipseRibbon2'">Curved Up Ribbon</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='verticalScroll'">Vertical Scroll</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='horizontalScroll'">Horizontal Scroll</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='wave'">Wave</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='doubleWave'">Double Wave</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='noSmoking'">No Symbol</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='plus'">Cross</xsl:when>
            <!--2011：加号、减号、除号转换为十字-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='mathPlus'">Cross</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='mathMinus'">Rectangle</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='mathMultiply'">Cross</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='mathDivide'">Cross</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='irregularSeal1'">Explosion 1</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='irregularSeal2'">Explosion 2</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='wedgeRectCallout'">Rectangular Callout</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='wedgeRoundRectCallout'">Rounded Rectangular Callout</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='wedgeEllipseCallout'">Oval Callout</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='borderCallout1'">Line Callout2</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='borderCallout2'">Line Callout3</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='borderCallout3'">Line Callout4</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='accentCallout1'">Line Callout2(Accent Bar)</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='accentCallout2'">Line Callout3(Accent Bar)</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='accentCallout3'">Line Callout4(Accent Bar)</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='callout1'">Line Callout2(No Border)</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='callout2'">Line Callout3(No Border)</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='callout3'">Line Callout4(No Border)</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='accentBorderCallout1'">Line Callout2(Border and Accent Bar)</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='accentBorderCallout2'">Line Callout3(Border and Accent Bar)</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='accentBorderCallout3'">Line Callout4(Border and Accent Bar)</xsl:when>
            <!--xsl:when test="a:prstGeom/@prst='straightConnector1'">Arrow</xsl:when-->
            <!--xsl:when test="a:prstGeom/@prst='curvedConnector3'">Curve</xsl:when-->
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bentConnector3'">Elbow Connector</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='curvedConnector3'">Curved Connector</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='curvedConnector4'">Curved Connector</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='homePlate'">Pentagon Arrow</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='quadArrowCallout'">Quad Arrow Callout</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartOnlineStorage'">Stored Data</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartMagneticTape'">Sequential Access Storage</xsl:when>
            <xsl:when test="./dsp:spPr/a:prstGeom/@prst='flowChartMagneticDrum'">Direct Access Storage</xsl:when>
            <!---->
            <xsl:when test="./dsp:spPr/a:custGeom/*">Curve</xsl:when>
            <xsl:otherwise>Rectangle</xsl:otherwise>
          </xsl:choose>
        </图:名称_801A>
        <图:生成软件_801B>UOFTranslator</图:生成软件_801B>
        <图:属性_801D>
          <xsl:choose>
            <xsl:when test="./dsp:spPr/a:gradFill">
              <!--渐变填充-->
              <图:填充_804C>
                <xsl:apply-templates select="dsp:spPr/a:gradFill"/>
              </图:填充_804C>
            </xsl:when>
            <xsl:when test="./dsp:spPr/a:pattFill">
              <!--图案填充-->
              <图:填充_804C>
                <xsl:apply-templates select="dsp:spPr/a:pattFill"/>
              </图:填充_804C>
            </xsl:when>
            <xsl:when test="./dsp:spPr/a:solidFill">
              <!--纯色填充-->
              <图:填充_804C>
                <图:颜色_8004>
                  <xsl:apply-templates select="dsp:spPr/a:solidFill"/>
                </图:颜色_8004>
              </图:填充_804C>
            </xsl:when>
            <xsl:when test="./dsp:spPr/a:blipFill">
              <!--图片填充-->
              <图:填充_804C>
                
                <!--2014-04-10，wudi，SmartArt图片填充需特殊处理，start-->
                <xsl:apply-templates select="dsp:spPr/a:blipFill">
                  <xsl:with-param name="picturefrom" select="$PicFrom"/>
                  <xsl:with-param name ="Number" select ="$number"/>
                </xsl:apply-templates>
                <!--end-->
                
              </图:填充_804C>
            </xsl:when>
            <xsl:when test="./dsp:spPr/a:grpFill">
              <!--组合图形填充-->
              <xsl:apply-templates select="./dsp:spPr/a:grpFill"/>
            </xsl:when>
            <xsl:when test="./dsp:spPr/a:noFill"/>
            <!--有style的情况-->
            <!-- <xsl:when test="following-sibling::p:style/a:fillRef">-->
            <xsl:when test="./dsp:style/a:fillRef[@idx!='0' and @idx!='1000']">
              <图:填充_804C>
                <图:颜色_8004>
                  <xsl:apply-templates select="wps:style/a:fillRef"/>
                </图:颜色_8004>
              </图:填充_804C>
            </xsl:when>
            <xsl:otherwise>
              <图:填充_804C>
                <图:颜色_8004>auto</图:颜色_8004>
              </图:填充_804C>
            </xsl:otherwise>
          </xsl:choose>
          <图:大小_8060>
            <xsl:attribute name="宽_C605">
              <xsl:if test="./dsp:spPr/a:xfrm/a:ext/@cx">
                <xsl:value-of select="./dsp:spPr/a:xfrm/a:ext/@cx div 12700"/>
              </xsl:if>
              <xsl:if test="./pic:spPr/a:xfrm/a:ext/@cx">
                <xsl:value-of select="./pic:spPr/a:xfrm/a:ext/@cx div 12700"/>
              </xsl:if>
            </xsl:attribute>
            <xsl:attribute name="长_C604">
              <xsl:if test="./dsp:spPr/a:xfrm/a:ext/@cy">
                <xsl:value-of select="./dsp:spPr/a:xfrm/a:ext/@cy div 12700"/>
              </xsl:if>
              <xsl:if test="./pic:spPr/a:xfrm/a:ext/@cy">
                <xsl:value-of select="./pic:spPr/a:xfrm/a:ext/@cy div 12700"/>
              </xsl:if>
            </xsl:attribute>
          </图:大小_8060>
          <!-- </xsl:if>-->
          <!--CXL线（包括线颜色、线粗线、线类型三个子元素）-->
          <图:线_8057>
            <xsl:choose>
              <xsl:when test="./dsp:spPr/a:ln">
                <xsl:apply-templates select="dsp:spPr/a:ln" mode="linecolor"/>
              </xsl:when>
              <xsl:when test="./dsp:style/a:lnRef/*">
                <xsl:for-each select="./wps:style/a:lnRef">
                  <图:线颜色_8058>
                    <xsl:call-template name="lnRefcolor"/>
                  </图:线颜色_8058>
                </xsl:for-each>
              </xsl:when>
              <xsl:otherwise>
                <图:线颜色_8058>#000000</图:线颜色_8058>
              </xsl:otherwise>
            </xsl:choose>
            <图:线类型_8059>
              <xsl:attribute name="线型_805A">
                <xsl:choose>
                  <xsl:when test="./dsp:spPr/a:ln/@cmpd='sng'">single</xsl:when>
                  <xsl:when test="./dsp:spPr/a:ln/@cmpd='dbl'">double</xsl:when>
                  <xsl:when test="./dsp:spPr/a:ln/@cmpd='thickThin'">thin-thick</xsl:when>
                  <xsl:when test="./dsp:spPr/a:ln/@cmpd='thinThick'">thick-thin</xsl:when>
                  <xsl:when test="./dsp:spPr/a:ln/@cmpd='tri'">thick-between-thin</xsl:when>
                  <xsl:otherwise>single</xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <xsl:choose>
                <xsl:when test="./dsp:spPr/a:ln/a:prstDash/@val">
                  <xsl:attribute name="虚实_805B">
                    <xsl:choose>
                      <xsl:when test="./dsp:spPr/a:ln/a:prstDash/@val='solid'">
                        <xsl:value-of select="'solid'"/>
                      </xsl:when>
                      <!--圆点线-->
                      <xsl:when test="./dsp:spPr/a:ln/a:prstDash/@val='sysDot'">
                        <xsl:value-of select="'round-dot'"/>
                      </xsl:when>
                      <!--方点线-->
                      <xsl:when test="./dsp:spPr/a:ln/a:prstDash/@val='sysDash'">
                        <xsl:value-of select="'square-dot'"/>
                      </xsl:when>
                      <!--短划线->虚线-->
                      <xsl:when test="./dsp:spPr/a:ln/a:prstDash/@val='dash'">
                        <xsl:value-of select="'dash'"/>
                      </xsl:when>
                      <!--划线-点->点虚线-->
                      <xsl:when test="./dsp:spPr/a:ln/a:prstDash/@val='dashDot'">
                        <xsl:value-of select="'dash-dot'"/>
                      </xsl:when>
                      <!--长划线->长虚线-->
                      <xsl:when test="./dsp:spPr/a:ln/a:prstDash/@val='lgDash'">
                        <xsl:value-of select="'long-dash'"/>
                      </xsl:when>
                      <!--长划线-点-->
                      <xsl:when test="./dsp:spPr/a:ln/a:prstDash/@val='lgDashDot'">
                        <xsl:value-of select="'long-dash-dot'"/>
                      </xsl:when>
                      <!--长划线-点-点-->
                      <xsl:when test="./dsp:spPr/a:ln/a:prstDash/@val='lgDashDotDot'">
                        <xsl:value-of select="'dash-dot-dot'"/>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:attribute name="虚实_805B">solid</xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </图:线类型_8059>
            <xsl:if test="name(.)='dsp:sp'">
              <xsl:choose>
                <xsl:when test="./dsp:spPr/a:ln/@w">
                  <图:线粗细_805C>
                    <xsl:value-of select="./dsp:spPr/a:ln/@w div 12700"/>
                  </图:线粗细_805C>
                </xsl:when>
                <xsl:otherwise>

                  <!--2012-11-22，wudi，修正对预定义图形线粗细的不正确处理，start-->
                  <图:线粗细_805C>
                    <xsl:choose>
                      <xsl:when test="./dsp:spPr/a:prstGeom/@prst='line'">0.75</xsl:when>
                      <!--直接连接符 2-->
                      <xsl:when test="./dsp:spPr/a:prstGeom/@prst='straightConnector1'">0.75</xsl:when>
                      <!--直接箭头连接符 3，直接箭头连接符 4-->
                      <xsl:when test="./dsp:spPr/a:prstGeom/@prst='arc'">0.75</xsl:when>
                      <!--弧形-->
                      <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bracketPair'">0.75</xsl:when>
                      <!--双括号 61-->
                      <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bracePair'">0.75</xsl:when>
                      <!--双大括号 62-->
                      <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftBracket'">0.75</xsl:when>
                      <!--左中括号 66-->
                      <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rightBracket'">0.75</xsl:when>
                      <!--右中括号 63-->
                      <xsl:when test="./dsp:spPr/a:prstGeom/@prst='leftBrace'">0.75</xsl:when>
                      <!--左大括号 64-->
                      <xsl:when test="./dsp:spPr/a:prstGeom/@prst='rightBrace'">0.75</xsl:when>
                      <!--右大括号 65-->
                      <xsl:when test="./dsp:spPr/a:prstGeom/@prst='bentConnector3'">0.75</xsl:when>
                      <!--肘形连接符 5，肘形连接符6，肘形连接符 7-->
                      <xsl:when test="./dsp:spPr/a:prstGeom/@prst='curvedConnector3'">0.75</xsl:when>
                      <!--曲线连接符8，曲线连接符9，曲线连接符11-->
                      <xsl:otherwise>2.0</xsl:otherwise>
                    </xsl:choose>
                  </图:线粗细_805C>
                  <!--end-->

                  <!--<图:线粗细_805C>2.0</图:线粗细_805C>-->
                  <!--原来是0.75，但发现WORD里预设为2磅-->
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
            <xsl:if test="name(.)='pic:pic'">
              <xsl:choose>
                <xsl:when test="./pic:spPr/a:ln/@w">
                  <!--cxl2011/12/9添加图片边框粗细-->
                  <图:线粗细_805C>
                    <xsl:value-of select="./pic:spPr/a:ln/@w div 12700"/>
                  </图:线粗细_805C>
                </xsl:when>
                <xsl:otherwise>
                  <图:线粗细_805C>0.0</图:线粗细_805C>
                  <!--原来是0.75，但发现WORD里预设为0.0磅-->
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
          </图:线_8057>
          <xsl:choose>
            <xsl:when test="./dsp:spPr/a:xfrm/@rot">
              <图:旋转角度_804D>
                <!-- 旋转角度<xsl:value-of select="(21600000-a:xfrm/@rot) div 60000"/>-->
                <xsl:value-of select="./dsp:spPr/a:xfrm/@rot div 60000"/>
              </图:旋转角度_804D>
            </xsl:when>
            <xsl:otherwise>
              <图:旋转角度_804D>0.0</图:旋转角度_804D>
            </xsl:otherwise>
          </xsl:choose>
          <!--zhaobj 箭头-->
          <xsl:if test="./dsp:spPr/a:ln">
            <xsl:apply-templates select="dsp:spPr/a:ln" mode="Arrow"/>
          </xsl:if>
          <!--<xsl:if test="wp:cNvGraphicFramePr/a:graphicFrameLocks/@noChangeAspect='1'">
            <图:锁定纵横比 uof:locID="g0028">0</图:锁定纵横比>
          </xsl:if>-->

          <图:缩放是否锁定纵横比_8055>
            <!--cxl,2012.3.6修改BUG-->
            <xsl:if test="$PicFrom = 'grpsp'">
              <xsl:if test="./wps:cNvSpPr/a:spLocks/@noChangeAspect='1'">
                <xsl:value-of select="'true'"/>
              </xsl:if>
              <xsl:if test="./wps:cNvSpPr/a:spLocks/@noChangeAspect='0' or not(./wps:cNvSpPr/a:spLocks/@noChangeAspect)">
                <xsl:value-of select="'false'"/>
              </xsl:if>
            </xsl:if>
            <xsl:if test="$PicFrom != 'grpsp'">
              <xsl:choose>
                <xsl:when test="./wps:cNvSpPr/a:spLocks/@noChangeAspect='1'">
                  <!--/../@locked-->
                  <xsl:value-of select="'true'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'false'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
          </图:缩放是否锁定纵横比_8055>
          <图:是否打印对象_804E>true</图:是否打印对象_804E>
          <!--<图:Web文字 uof:locID="g0033"/>-->
          <!--zhaobj 透明度a:solidFill 还有 a:gradFill a:pattFill a:blipFill的情况-->
          
          <!--2013-11-27，wudi，Strict标准下，val取值有变，修正，start-->
          <xsl:if test="./dsp:spPr/a:ln/a:solidFill//a:alpha">
            <图:透明度_8050>
              <xsl:variable name="tmp">
                <xsl:value-of select="number(substring-before(./dsp:spPr/a:ln/a:solidFill//a:alpha/@val,'%'))"/>
              </xsl:variable>
              <xsl:value-of select="round(100 - $tmp)"/>
            </图:透明度_8050>
          </xsl:if>
          <!--end-->
          
          <xsl:if test="./dsp:spPr/a:effectLst/a:outerShdw">
            <!--外部阴影-->
            <xsl:apply-templates select="dsp:spPr/a:effectLst/a:outerShdw" mode="shadow"/>
          </xsl:if>
          <xsl:if test="./dsp:spPr/a:effectLst/a:innerShdw">
            <!--内部阴影-->
            <xsl:apply-templates select="dsp:spPr/a:effectLst/a:innerShdw" mode="shadow"/>
          </xsl:if>
          
          <!--cxl2011/11/20开始增加三维效果转化代码*****************************************************************************-->
          <xsl:if test="./dsp:spPr/a:scene3d">
            <图:三维效果_8061>
              <xsl:apply-templates select="./dsp:spPr/a:scene3d"/>
            </图:三维效果_8061>
          </xsl:if>
        </图:属性_801D>
      </图:预定义图形_8018>
      <!--cxl2011/11/19添加图片数据引用,其它数据引用按范围文档不转************************************************-->
      <xsl:if test="./pic:blipFill">
        <!--./dsp:spPr/a:blipFill|-->
        <图:图片数据引用_8037>
          <xsl:variable name="picId">
            <xsl:if test="./pic:blipFill/a:blip/@r:embed">
              <xsl:value-of select="./pic:blipFill/a:blip/@r:embed"/>
            </xsl:if>
            <!--<xsl:if test="./dsp:spPr/a:blipFill/a:blip/@r:embed">-->
            <!--待做案例验证-->
            <!--
              <xsl:value-of select="./dsp:spPr/a:blipFill/a:blip/@r:embed"/>
            </xsl:if>-->
          </xsl:variable>
          <xsl:variable name="picName">
            <xsl:choose>
              <xsl:when test="$PicFrom='document'">
                <xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="$PicFrom='comments'">
                <xsl:value-of select="document('word/_rels/comments.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="$PicFrom='endnotes'">
                <xsl:value-of select="document('word/_rels/endnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="$PicFrom='footnotes'">
                <xsl:value-of select="document('word/_rels/footnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="contains($PicFrom,'header')">
                <xsl:variable name="hn" select="concat('word/_rels/',$PicFrom,'.xml.rels')"/>
                <xsl:value-of select="document($hn)/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="contains($PicFrom,'footer')">
                <xsl:variable name="fn" select="concat('word/_rels/',$PicFrom,'.xml.rels')"/>
                <xsl:value-of select="document($fn)/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <!--暂时先这样-->
              <xsl:when test="$PicFrom='grpsp'">
                <xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="pictype">
            <xsl:value-of select="substring-after($picName,'.')"/>
          </xsl:variable>
          <xsl:if test="$PicFrom = 'grpsp'">
            <xsl:if test="dsp:cNvPr/@id">
              <xsl:value-of select="concat('grpsppicObj',dsp:cNvPr/@id)"/>
            </xsl:if>
            <xsl:if test="dsp:cNvPr/@id">
              <xsl:value-of select="concat('grpsppicObj',dsp:cNvPr/@id)"/>
            </xsl:if>
          </xsl:if>
          <xsl:if test="$PicFrom != 'grpsp'">
            <xsl:value-of select="concat($PicFrom,'Obj',$number * 2)"/>
            <!--substring-after($picName,'/')-->
          </xsl:if>
        </图:图片数据引用_8037>
      </xsl:if>
      <!--控制点-->
      <!--翻转，cxl，2012.3.9互换x,y位置-->
      <xsl:if test="./*/a:xfrm/@flipH">
        <图:翻转_803A>
          <xsl:value-of select="'x'"/>
        </图:翻转_803A>
      </xsl:if>
      <xsl:if test="./*/a:xfrm/@flipV">
        <图:翻转_803A>
          <xsl:value-of select="'y'"/>
        </图:翻转_803A>
      </xsl:if>
      <xsl:for-each select="./dsp:txBody">
        <xsl:apply-templates select="." mode="文本"/>
        <!--预定义图形中文本的转化-->
      </xsl:for-each>
      
      <!--2013-05-03，wudi，修改grpsp为SmartArt，区分SmartArt与组合图形，流程图等，start-->
      <xsl:if test="$PicFrom ='SmartArt'">
        <图:组合位置_803B>
          <xsl:attribute name="x_C606">
            <xsl:value-of select="./dsp:spPr/a:xfrm/a:off/@x div 12700"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="./dsp:spPr/a:xfrm/a:off/@y div 12700"/>
          </xsl:attribute>
        </图:组合位置_803B>
      </xsl:if>
      <!--end-->
      
    </图:图形_8062>
  </xsl:template>
  <xsl:template match="dsp:txBody" mode="文本">
    <图:文本_803C>
      <!--<xsl:attribute name ="是否为文本框_8046">
        <xsl:value-of select ="'false'"/>
      </xsl:attribute>-->
      <xsl:attribute name="是否自动换行_8047">
        <xsl:if test="./a:bodyPr/@wrap='none'">
          <xsl:value-of select="'false'"/>
        </xsl:if>
        <xsl:if test="./a:bodyPr/@wrap='square'">
          <xsl:value-of select="'true'"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name="是否大小适应文字_8048">
        <xsl:if test="./a:bodyPr/a:spAutoFit">
          <xsl:value-of select="'true'"/>
        </xsl:if>
        <xsl:if test="./a:bodyPr/a:noAutofit">
          <xsl:value-of select="'false'"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name="是否文字随对象旋转_8049">
        <xsl:choose>
          <xsl:when test="./a:bodyPr/@vert='horz'">
            <xsl:value-of select="'false'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'true'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <图:边距_803D>

        <!--2014-06-05，wudi，修改转换比例，有变化，start-->
        <xsl:attribute name="左_C608">
          <xsl:value-of select="./a:bodyPr/@lIns div 12700"/>
        </xsl:attribute>
        <xsl:attribute name="右_C60A">
          <xsl:value-of select="./a:bodyPr/@rIns div 12700"/>
        </xsl:attribute>
        <xsl:attribute name="上_C609">
          <xsl:value-of select="./a:bodyPr/@tIns div 12700"/>
        </xsl:attribute>
        <xsl:attribute name="下_C60B">
          <xsl:value-of select="./a:bodyPr/@bIns div 12700"/>
        </xsl:attribute>
        <!--end-->
        
      </图:边距_803D>
      <图:对齐_803E>
        <xsl:attribute name="水平对齐_421D">
          <xsl:if test="./a:bodyPr/@lIns">
            <xsl:value-of select="./a:bodyPr/@lIns div 9525"/>
          </xsl:if>
          <xsl:if test="not(./a:bodyPr/@lIns)">
            <xsl:value-of select="'left'"/>
          </xsl:if>
        </xsl:attribute>
        <xsl:attribute name="水平对齐_421D">
          <xsl:choose>
            <xsl:when test="./a:bodyPr/@vert='eaVert'">
              <xsl:value-of select="'left'"/>
            </xsl:when>
            <xsl:when test="./a:bodyPr/@vert='horz'">
              <xsl:value-of select="'right'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'left'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="文字对齐_421E">
          <xsl:choose>
            <xsl:when test="./a:bodyPr/@anchor='t'">
              <xsl:value-of select="'top'"/>
            </xsl:when>
            <xsl:when test="./a:bodyPr/@anchor='ctr'">
              <xsl:value-of select="'center'"/>
            </xsl:when>
            <xsl:when test="./a:bodyPr/@anchor='b'">
              <xsl:value-of select="'bottom'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </图:对齐_803E>
      <图:文字排列方向_8042>
        <xsl:if test="./a:bodyPr/@vert='horz'">
          <xsl:choose>
            <xsl:when test="parent::dsp:sp/@normalEastAsianFlow='1'">
              <xsl:value-of select="'t2b-l2r-270e-0w'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'t2b-l2r-0e-0w'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <xsl:if test="./a:bodyPr/@vert='eaVert'">
          <xsl:value-of select="'r2l-t2b-0e-90w'"/>
        </xsl:if>
        <xsl:if test="./a:bodyPr/@vert='vert'">
          <xsl:value-of select="'r2l-t2b-90e-90w'"/>
        </xsl:if>
        <xsl:if test="./a:bodyPr/@vert='vert270'">
          <xsl:value-of select="'l2r-b2t-270e-270w'"/>
        </xsl:if>
      </图:文字排列方向_8042>
      <图:内容_8043>
        <xsl:for-each select="./a:p">
          <字:段落_416B>
            <字:段落属性_419B>

              <!--2014-05-06，wudi，修复SmartArt中文本式样转换BUG，start-->
              <xsl:attribute name="式样引用_419C">
                <xsl:variable name="styleId"
    select="document('word/styles.xml')/w:styles/w:style[(@w:type='paragraph') and (@w:default='1'or @w:default='on' or @w:default='true')]/@w:styleId">
                </xsl:variable>
                <xsl:choose>
                  <xsl:when test="starts-with($styleId,'id_')">
                    <!--用于比较两个字符串，若第一个是以第二个开始，返回True-->
                    <xsl:value-of select ="$styleId"/>
                    <!--用于提取选定节点的值-->
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select ="concat('id_',$styleId)"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <!--end-->
              
              <字:对齐_417D>
                <xsl:attribute name ="水平对齐_421D">
                  <xsl:if test ="./a:pPr/@algn='t'">
                    <xsl:value-of select ="'top'"/>
                  </xsl:if>
                  <xsl:if test ="./a:pPr/@algn='ctr'">
                    <xsl:value-of select ="'center'"/>
                  </xsl:if>
                  <xsl:if test ="./a:pPr/@algn='b'">
                    <xsl:value-of select ="'bottom'"/>
                  </xsl:if>
                  <xsl:if test ="./a:pPr/@algn='r'">
                    <xsl:value-of select ="'right'"/>
                  </xsl:if>
                  
                  <!--2013-05-02，wudi，修复OOX到UOF方向SmartArt转换问题，修正取值原'light'应为'left'-->
                  <xsl:if test ="./a:pPr/@algn='l'">
                    <xsl:value-of select ="'left'"/>
                  </xsl:if>
                  <!--end-->
                  
                </xsl:attribute>
              </字:对齐_417D>
            </字:段落属性_419B>
            <xsl:for-each select ="./a:r">
                <字:句_419D>
                  <字:句属性_4158>
                    <字:字体_4128>
                      <xsl:attribute name="字号_412D">
                        <xsl:value-of select="format-number(number(./a:rPr/@sz) div 100,'0.0')"/>
                      </xsl:attribute>
                      
                      <!--2013-05-02，wudi，修复OOX到UOF方向SmartArt转换问题，增加对字体颜色的转换，start-->
                      <xsl:attribute name ="颜色_412F">
                        <xsl:if test ="./a:rPr/a:solidFill/a:schemeClr">
                          <xsl:variable name ="tmp">
                            <xsl:value-of select ="concat('a:',./a:rPr/a:solidFill/a:schemeClr/@val)"/>
                          </xsl:variable>
                          <xsl:for-each select ="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/node()">
                            <xsl:if test ="name(.) =$tmp">
                              <xsl:value-of select ="concat('#',./a:srgbClr/@val)"/>
                            </xsl:if>
                          </xsl:for-each>
                        </xsl:if>
                        <xsl:if test ="not(./a:rPr/a:solidFill/a:schemeClr)">
                          <xsl:value-of select ="'#FFFFFF'"/>
                        </xsl:if>
                      </xsl:attribute>
                      <!--end-->
                      
                    </字:字体_4128>
                    <字:是否粗体_4130>
                      <xsl:choose>
                        
                        <!--2014-04-10，wudi，去掉not，改条件为(./a:rPr/@b)，start-->
                        <xsl:when test="(./a:rPr/@b) or (./a:rPr/@b='on') or (./a:rPr/@b='1') or (./a:rPr/@b='true')">
                          <xsl:value-of select="'true'"/>
                        </xsl:when>
                        <!--end-->
                        
                        <xsl:otherwise>
                          <xsl:value-of select="'false'"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </字:是否粗体_4130>
                    <字:是否斜体_4131>
                      <xsl:choose>

                        <!--2014-04-10，wudi，去掉not，改条件为(./a:rPr/@i)，start-->
                        <xsl:when test="(./a:rPr/@i) or (./a:rPr/@i='on') or (./a:rPr/@i='1') or (./a:rPr/@i='true')">
                          <xsl:value-of select="'true'"/>
                        </xsl:when>
                        <!--end-->
                        
                        <xsl:otherwise>
                          <xsl:value-of select="'false'"/>
                        </xsl:otherwise>
                      </xsl:choose>
                    </字:是否斜体_4131>
                    <字:调整字间距_4146>
                      <xsl:value-of select="format-number(number(./a:rPr/@kern) div 100,'0.0')"/>
                    </字:调整字间距_4146>
                  </字:句属性_4158>
                  <字:文本串_415B>
                    <xsl:value-of select ="./a:t"/>
                  </字:文本串_415B>
                </字:句_419D>
              </xsl:for-each>
          </字:段落_416B>
        </xsl:for-each>
      </图:内容_8043>
    </图:文本_803C>
  </xsl:template>
  <!--end-->

  <!--2013-11-06，wudi，Strict标准下为wp前缀，增加对此前缀的处理，start-->
  <xsl:template match="wp:wgp | wp:grpSp">
    <!--组合图形转换-->
    <xsl:param name="picFrom"/>
    <xsl:variable name="number">
      <xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline | v:shape"/>
    </xsl:variable>
    <xsl:for-each select="wp:wsp|wp:grpSp">
      <xsl:if test="name(.)='wp:wsp'">
        <xsl:apply-templates select=".">
          <xsl:with-param name="picType" select="'anchor'"/>
          <xsl:with-param name="picFrom" select="'grpsp'"/>
        </xsl:apply-templates>
      </xsl:if>
      <xsl:if test="name(.)='wp:grpSp'">
        <xsl:apply-templates select=".">
          <xsl:with-param name="picFrom" select="'grpsp'"/>
        </xsl:apply-templates>
      </xsl:if>
    </xsl:for-each>
    <图:图形_8062>
      <xsl:attribute name="层次_8063">
        <xsl:if test="ancestor::wp:anchor/@relativeHeight">
          <xsl:value-of select="ancestor::wp:anchor/@relativeHeight"/>
        </xsl:if>
        <xsl:if test="not(ancestor::wp:anchor/@relativeHeight)">
          <xsl:value-of select="'3'"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name="标识符_804B">
        <xsl:if test="$picFrom ='grpsp'">

          <!--2013-05-03，wudi，修复组合图形，流程图，SmartArt互操作BUG，三种“组合图形”用不同的标识符，start-->
          <xsl:choose>
            <xsl:when test ="contains(ancestor::wp:anchor/wp:docPr/@name,'组合') or contains(ancestor::wp:inline/wp:docPr/@name,'组合')">
              <xsl:value-of select="concat('grpspObj',./wp:cNvPr/@id)"/>
            </xsl:when>
            <xsl:when test ="contains(ancestor::wp:anchor/wp:docPr/@name,'Group') or contains(ancestor::wp:inline/wp:docPr/@name,'Group')">
              <xsl:value-of select="concat('GrpspObj',./wp:cNvPr/@id)"/>
            </xsl:when>
            <xsl:when test ="contains(ancestor::wp:anchor/wp:docPr/@name,'图示') or contains(ancestor::wp:inline/wp:docPr/@name,'图示')">
              <xsl:value-of select="concat('SmartArtObj',./wp:cNvPr/@id)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="concat('grpspObj',./wp:cNvPr/@id)"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--end-->

        </xsl:if>
        <xsl:if test="$picFrom !='grpsp'">
          <xsl:value-of select="concat($picFrom,'Obj',$number * 2 + 1)"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name="组合列表_8064">
        <xsl:for-each select="wp:wsp | wp:grpSp">
          <xsl:if test="name(.)='wp:wsp'">

            <!--2013-05-03，wudi，修复组合图形，流程图，SmartArt互操作BUG，三种“组合图形”用不同的标识符，start-->
            <xsl:choose>
              <xsl:when test ="contains(ancestor::wp:anchor/wp:docPr/@name,'组合') or contains(ancestor::wp:inline/wp:docPr/@name,'组合')">
                <xsl:value-of select="concat(' ','grpspObj',./wp:cNvPr/@id)"/>
              </xsl:when>
              <xsl:when test ="contains(ancestor::wp:anchor/wp:docPr/@name,'Group') or contains(ancestor::wp:inline/wp:docPr/@name,'Group')">
                <xsl:value-of select="concat(' ','GrpspObj',./wp:cNvPr/@id)"/>
              </xsl:when>
              <xsl:when test ="contains(ancestor::wp:anchor/wp:docPr/@name,'图示') or contains(ancestor::wp:inline/wp:docPr/@name,'图示')">
                <xsl:value-of select="concat('SmartArtObj',./wp:cNvPr/@id)"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="concat(' ','grpspObj',./wp:cNvPr/@id)"/>
              </xsl:otherwise>
            </xsl:choose>
            <!--end-->

          </xsl:if>
          <xsl:if test="name(.)='wp:grpSp'">

            <!--2013-05-03，wudi，修复组合图形，流程图，SmartArt互操作BUG，三种“组合图形”用不同的标识符，start-->
            <xsl:choose>
              <xsl:when test ="contains(ancestor::wp:anchor/wp:docPr/@name,'组合') or contains(ancestor::wp:inline/wp:docPr/@name,'组合')">
                <xsl:value-of select="concat(' ','grpspObj',./wp:cNvPr/@id)"/>
              </xsl:when>
              <xsl:when test ="contains(ancestor::wp:anchor/wp:docPr/@name,'Group') or contains(ancestor::wp:inline/wp:docPr/@name,'Group')">
                <xsl:value-of select="concat(' ','GrpspObj',./wp:cNvPr/@id)"/>
              </xsl:when>
              <xsl:when test ="contains(ancestor::wp:anchor/wp:docPr/@name,'图示') or contains(ancestor::wp:inline/wp:docPr/@name,'图示')">
                <xsl:value-of select="concat('SmartArtObj',./wp:cNvPr/@id)"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="concat(' ','grpspObj',./wp:cNvPr/@id)"/>
              </xsl:otherwise>
            </xsl:choose>
            <!--end-->

          </xsl:if>
        </xsl:for-each>
      </xsl:attribute>
      <图:预定义图形_8018>
        <图:类别_8019>11</图:类别_8019>
        <图:名称_801A>Rectangle</图:名称_801A>
        <图:生成软件_801B>UOFTranslator</图:生成软件_801B>
        <图:属性_801D>
          <图:大小_8060>
            <xsl:attribute name="宽_C605">
              <xsl:value-of select="./wp:grpSpPr/a:xfrm/a:ext/@cx div 12700"/>
            </xsl:attribute>
            <xsl:attribute name="长_C604">
              <xsl:value-of select="./wp:grpSpPr/a:xfrm/a:ext/@cy div 12700"/>
            </xsl:attribute>
          </图:大小_8060>
          <xsl:choose>
            <xsl:when test="./wp:grpSpPr/a:xfrm/@rot">
              <图:旋转角度_804D>
                <xsl:value-of select="./wp:grpSpPr/a:xfrm/@rot div 60000"/>
              </图:旋转角度_804D>
            </xsl:when>
            <xsl:otherwise>
              <图:旋转角度_804D>0.0</图:旋转角度_804D>
            </xsl:otherwise>
          </xsl:choose>
          <图:缩放是否锁定纵横比_8055>
            <!--cxl,2012.3.6修改BUG-->
            <xsl:if test="./wp:cNvGrpSpPr/a:grpSpLocks/@noChangeAspect='1'">
              <xsl:value-of select=" 'true'"/>
            </xsl:if>
            <xsl:if test="./wp:cNvGrpSpPr/a:grpSpLocks/@noChangeAspect='0' or not(./wp:cNvGrpSpPr/a:grpSpLocks/@noChangeAspect)">
              <xsl:value-of select="'false'"/>
            </xsl:if>
          </图:缩放是否锁定纵横比_8055>
          <图:是否打印对象_804E>true</图:是否打印对象_804E>
          <!--<图:Web文字 uof:locID="g0033"/>-->
        </图:属性_801D>
      </图:预定义图形_8018>

      <!--2013-05-07，wudi，修复OOX到UOF方向组合图形转换BUG，针对组合图形嵌套组合图形的情况，start-->
      <xsl:if test ="$picFrom ='grpsp'">
        <xsl:if test ="./wp:grpSpPr/a:xfrm/a:off/@x != '0' and ./wp:grpSpPr/a:xfrm/a:off/@y != '0'">
          <图:组合位置_803B>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="./wp:grpSpPr/a:xfrm/a:off/@x div 12700"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="./wp:grpSpPr/a:xfrm/a:off/@y div 12700"/>
            </xsl:attribute>
          </图:组合位置_803B>
        </xsl:if>
      </xsl:if>
      <!--end-->

      <!--cxl2011.11.23删除，因为生成不必要的代码-->
      <!--<xsl:if test="$picFrom ='grpsp'">
				<图:组合位置_803B>
					<xsl:attribute name="x_C606">
						<xsl:value-of select="./wp:grpSpPr/a:xfrm/a:off/@x div 12700"/>
					</xsl:attribute>
					<xsl:attribute name="y_C607">
						<xsl:value-of select="./wp:grpSpPr/a:xfrm/a:off/@y div 12700"/>
					</xsl:attribute>
				</图:组合位置_803B>
			</xsl:if>-->
    </图:图形_8062>
  </xsl:template>
  <!--end-->
  
  <!--@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->

  <!--2013-11-05，wudi，Strict标准下为wp前缀，加以区分，start-->
  <xsl:template match="wp:wsp|pic:pic">
    <xsl:param name="picFrom"/>
    <xsl:variable name="number">
      <xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline | v:shape"/>
    </xsl:variable>
    <xsl:variable name="style" select="@style"/>
    <!--<xsl:if test="./wps:spPr/a:blipFill|./pic:blipFill">
			<uof:其他对象 uof:locID="u0036" uof:attrList="标识符 内嵌 公共类型 私有类型" uof:内嵌="false">
				<xsl:variable name="picId">
					<xsl:if test="./pic:blipFill/a:blip/@r:embed">
						<xsl:value-of select="./pic:blipFill/a:blip/@r:embed"/>
					</xsl:if>
					<xsl:if test="./wps:spPr/a:blipFill/a:blip/@r:embed">
						<xsl:value-of select="./wps:spPr/a:blipFill/a:blip/@r:embed"/>
					</xsl:if>
				</xsl:variable>
				<xsl:variable name="picName">
					<xsl:choose>
						<xsl:when test="$picFrom='document'">
							<xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
						<xsl:when test="$picFrom='comments'">
							<xsl:value-of select="document('word/_rels/comments.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
						<xsl:when test="$picFrom='endnotes'">
							<xsl:value-of select="document('word/_rels/endnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
						<xsl:when test="$picFrom='footnotes'">
							<xsl:value-of select="document('word/_rels/footnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
						<xsl:when test="contains($picFrom,'header')">
							<xsl:variable name="hn" select="concat('word/_rels/',$picFrom,'.xml.rels')"/>
							<xsl:value-of select="document($hn)/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
						<xsl:when test="contains($picFrom,'footer')">
							<xsl:variable name="fn" select="concat('word/_rels/',$picFrom,'.xml.rels')"/>
							<xsl:value-of select="document($fn)/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
						-->
    <!--暂时先这样-->
    <!--
						<xsl:when test="$picFrom='grpsp'">
							<xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
						</xsl:when>
					</xsl:choose>
				</xsl:variable>
				<xsl:variable name="pictype">
					<xsl:value-of select="substring-after($picName,'.')"/>
				</xsl:variable>
				<xsl:attribute name="uof:标识符">
					<xsl:if test="$picFrom = 'grpsp'">
						<xsl:if test="wps:cNvPr/@id">
							<xsl:value-of select="concat('grpsppicOBJ',wps:cNvPr/@id)"/>
						</xsl:if>
						<xsl:if test="wpg:cNvPr/@id">
							<xsl:value-of select="concat('grpsppicOBJ',wpg:cNvPr/@id)"/>
						</xsl:if>
					</xsl:if>
					<xsl:if test="$picFrom != 'grpsp'">
						<xsl:value-of select="concat($picFrom,'OBJ',$number * 2)"/>
					</xsl:if>
				</xsl:attribute>
				<xsl:attribute name="uof:公共类型">
					<xsl:call-template name="objtype">
						<xsl:with-param name="val" select="$pictype"/>
					</xsl:call-template>
				</xsl:attribute>
				<o2upic:picture>
					<xsl:attribute name="target">
						<xsl:value-of select="concat('word/',$picName)"/>
					</xsl:attribute>
				</o2upic:picture>
			</uof:其他对象>
		</xsl:if>-->
    <图:图形_8062>
      <!--预定义图形处理-->
      <xsl:attribute name="标识符_804B">
        <xsl:if test="$picFrom='grpsp'">

          <!--2013-05-03，wudi，修复组合图形，流程图，SmartArt互操作BUG，三种“组合图形”用不同的标识符，start-->
          <xsl:choose>
            <xsl:when test ="contains(ancestor::wp:anchor/wp:docPr/@name,'组合') or contains(ancestor::wp:inline/wp:docPr/@name,'组合')">
              <xsl:value-of select="concat('grpspObj',./wp:cNvPr/@id)"/>
            </xsl:when>
            <xsl:when test ="contains(ancestor::wp:anchor/wp:docPr/@name,'Group') or contains(ancestor::wp:inline/wp:docPr/@name,'Group')">
              <xsl:value-of select="concat('GrpspObj',./wp:cNvPr/@id)"/>
            </xsl:when>
            <xsl:when test ="contains(ancestor::wp:anchor/wp:docPr/@name,'图示') or contains(ancestor::wp:inline/wp:docPr/@name,'图示')">
              <xsl:value-of select="concat('SmartArtObj',./wp:cNvPr/@id)"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="concat('grpspObj',./wp:cNvPr/@id)"/>
            </xsl:otherwise>
          </xsl:choose>
          <!--end-->

        </xsl:if>
        <xsl:if test="$picFrom != 'grpsp'">
          <xsl:value-of select="concat($picFrom,'Obj',$number * 2 +1)"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name="层次_8063">
        <!--cxl,2012.5.14图形叠放次序，暂时先这样处理，效果上没问题-->
        <xsl:if test="ancestor::wp:anchor/@relativeHeight">
          <xsl:value-of select="ancestor::wp:anchor/@relativeHeight"/>
        </xsl:if>
        <xsl:if test="not(ancestor::wp:anchor/@relativeHeight)">
          <xsl:value-of select="'3'"/>
        </xsl:if>
      </xsl:attribute>
      <!--<xsl:if test ="name(.)='pic:pic'">
        <xsl:attribute name="图:其他对象">
          <xsl:value-of select ="concat($picFrom,'OBJ',$number * 2)"/>
        </xsl:attribute>
      </xsl:if>-->
      <图:预定义图形_8018>
        <图:类别_8019>
          <xsl:choose>
            <!--cxl2011/11/21增加了几种漏转的图形（包括七边形、乘号、六角星、七角星、10角星、16角星），与差异文档做了对应关系的修改（包括减号、L形）-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='rect'">11</xsl:when>
            <!--矩形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='parallelogram'">12</xsl:when>
            <!--平行四边形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='trapezoid'">13</xsl:when>
            <!--梯形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='diamond'">14</xsl:when>
            <!--菱形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='roundRect'">15</xsl:when>
            <!--圆角矩形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='heptagon'">16</xsl:when>
            <!--七边形转换为八边形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='octagon'">16</xsl:when>
            <!--八边形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='triangle'">17</xsl:when>
            <!--三角形17代表等腰三角形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='rtTriangle'">18</xsl:when>
            <!--右三角形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='ellipse'">19</xsl:when>
            <!--椭圆-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='diagStripe'">18</xsl:when>
            <!--斜纹-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='corner'">11</xsl:when>
            <!--L形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='halfFrame'">18</xsl:when>
            <!--半闭框-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='decagon'">19</xsl:when>
            <!--十边形转换为圆-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='dodecagon'">19</xsl:when>
            <!--十二边形转换为圆-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='pie'">19</xsl:when>
            <!--饼形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='chord'">19</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='teardrop'">19</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='cloud'">19</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='rightArrow'">21</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftArrow'">22</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='upArrow'">23</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='downArrow'">24</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftRightArrow'">25</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='upDownArrow'">26</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='quadArrow'">27</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftRightUpArrow'">28</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='bentArrow'">29</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartProcess'">31</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartAlternateProcess'">32</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartDecision'">33</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartInputOutput'">34</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartPredefinedProcess'">35</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartInternalStorage'">36</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartDocument'">37</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartMultidocument'">38</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartTerminator'">39</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star4'">43</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star5'">44</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star6'">45</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star7'">45</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star8'">45</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star10'">46</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star12'">46</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star16'">46</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star24'">47</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star32'">48</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='ribbon'">410</xsl:when>
            
            <!--2013-12-04，wudi，云形标注，start-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='cloudCallout'">54</xsl:when>
            <!--end-->
            
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='line'">61</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='straightConnector1'">71</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='hexagon'">110</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='pentagon'">112</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='can'">113</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='cube'">114</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='bevel'">115</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='foldedCorner'">116</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='smileyFace'">117</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='donut'">118</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='blockArc'">120</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='heart'">121</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='lightningBolt'">122</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='sun'">123</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='moon'">124</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='arc'">125</xsl:when>
            <!--弧形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='bracketPair'">126</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='bracePair'">127</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='plaque'">128</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftBracket'">129</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='rightBracket'">130</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftBrace'">131</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='rightBrace'">132</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='uturnArrow'">210</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftUpArrow'">211</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='bentUpArrow'">212</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedRightArrow'">213</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedLeftArrow'">214</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedUpArrow'">215</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedDownArrow'">216</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='stripedRightArrow'">217</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='notchedRightArrow'">218</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='chevron'">220</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='rightArrowCallout'">221</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftArrowCallout'">222</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='upArrowCallout'">223</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='downArrowCallout'">224</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftRightArrowCallout'">225</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='upDownArrowCallout'">226</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='circularArrow'">228</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartPreparation'">310</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartManualInput'">311</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartManualOperation'">312</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartConnector'">313</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartOffpageConnector'">314</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartPunchedCard'">315</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartPunchedTape'">316</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartSummingJunction'">317</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartOr'">318</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartCollate'">319</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartSort'">320</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartExtract'">321</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartMerge'">322</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartDelay'">324</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartMagneticDisk'">326</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartDisplay'">328</xsl:when>
            <!--<xsl:when test="a:prstGeom/@prst='ribbon2'">410</xsl:when>-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='ribbon2'">49</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='ellipseRibbon'">412</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='ellipseRibbon2'">411</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='verticalScroll'">413</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='horizontalScroll'">414</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='wave'">415</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='doubleWave'">416</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='noSmoking'">119</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='plus'">111</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='mathPlus'">111</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='mathMinus'">11</xsl:when>
            <!--减号转成矩形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='mathMultiply'">111</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='mathDivide'">111</xsl:when>
            <!--xsl:when test="a:prstGeom/@prst='straightConnector1'">71</xsl:when-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='irregularSeal1'">41</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='irregularSeal2'">42</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='wedgeRectCallout'">51</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='wedgeRoundRectCallout'">52</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='wedgeEllipseCallout'">53</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='borderCallout1'">56</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='borderCallout2'">57</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='borderCallout3'">58</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='accentCallout1'">510</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='accentCallout2'">511</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='accentCallout3'">512</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='callout1'">514</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='callout2'">515</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='callout3'">516</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='accentBorderCallout1'">518</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='accentBorderCallout2'">519</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='accentBorderCallout3'">520</xsl:when>
            <!--xsl:when test="a:prstGeom/@prst='straightConnector1'">62</xsl:when-->
            <!--<xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedConnector3'">64</xsl:when>-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='bentConnector3'">74</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedConnector3'">77</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='homePlate'">219</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='quadArrowCallout'">227</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartOnlineStorage'">323</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartMagneticTape'">325</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartMagneticDrum'">327</xsl:when>
            <xsl:otherwise>11</xsl:otherwise>
          </xsl:choose>
        </图:类别_8019>
        <图:名称_801A>
          <xsl:choose>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='rect'">Rectangle</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='parallelogram'">Parallelogram</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='trapezoid'">Trapezoid</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='diamond'">Diamond</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='roundRect'">Rounded Rectangle</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='heptagon'">Octagon</xsl:when>
            <!--七边形转换为八边形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='octagon'">Octagon</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='triangle'">Isosceles Triangle</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='rtTriangle'">Right Triangle</xsl:when>
            <!--斜纹，半闭框，L形转换为三角形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='diagStripe'">Right Triangle</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='corner'">Rectangle</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='halfFrame'">Right Triangle</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='ellipse'">Oval</xsl:when>
            <!--2011：十边形，十二边形，饼，泪滴，弦月，云，转换为圆-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='decagon'">Oval</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='dodecagon'">Oval</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='pie'">Oval</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='chord'">Oval</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='teardrop'">Oval</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='cloud'">Oval</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='rightArrow'">Right Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftArrow'">Left Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='upArrow'">Up Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='downArrow'">Down Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftRightArrow'">Left-Right Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='upDownArrow'">Up-Down Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='quadArrow'">Quad Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftRightUpArrow'">Left-Right-Up Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='bentArrow'">Bent Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartProcess'">Process</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartAlternateProcess'">Alternate Process</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartDecision'">Decision</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartProcess'">Process</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartPredefinedProcess'">Predefined Process</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartInputOutput'">Data</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartInternalStorage'">Internal Storage</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartDocument'">Document</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartMultidocument'">Multidocument</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartTerminator'">Terminator</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star4'">4-Point Star</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star5'">5-Point Star</xsl:when>
            <!--2011：6、7角形转换为8角形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star6'">8-Point Star</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star7'">8-Point Star</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star8'">8-Point Star</xsl:when>
            <!--2011-：10、12角形转换为16角形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star10'">16-Point Star</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star12'">16-Point Star</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star16'">16-Point Star</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star24'">24-Point Star</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='star32'">32-Point Star</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='ribbon'">Down Ribbon</xsl:when>
            
            <!--2013-12-04，wudi，云形标注，start-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='cloudCallout'">Cloud Callout</xsl:when>
            <!--end-->
            
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='line'">Line</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='straightConnector1'">Straight Connector</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='hexagon'">Hexagon</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='pentagon'">Regual Pentagon</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='can'">Can</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='cube'">Cube</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='bevel'">Bevel</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='foldedCorner'">Folded Corner</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='smileyFace'">Smiley Face</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='donut'">Donut</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='blockArc'">Block Arc</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='heart'">Heart</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='lightningBolt'">Lightning</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='sun'">Sun</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='moon'">Moon</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='arc'">Arc</xsl:when>
            <!--弧形-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='bracketPair'">Double Bracket</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='bracePair'">Double Brace</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='plaque'">Plaque</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftBracket'">Left Bracket</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='rightBracket'">Right Bracket</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftBrace'">Left Brace</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='rightBrace'">Right Brace</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='uturnArrow'">U-Turn Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftUpArrow'">Left-Up Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='bentUpArrow'">Bent-Up Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedRightArrow'">Curved Right Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedLeftArrow'">Curved Left Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedUpArrow'">Curved Up Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedDownArrow'">Curved Down Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='stripedRightArrow'">Striped Right Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='notchedRightArrow'">Notched Right Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='chevron'">Chevron Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='rightArrowCallout'">Right Arrow Callout</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftArrowCallout'">Left Arrow Callout</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='upArrowCallout'">Up Arrow Callout</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='downArrowCallout'">Down Arrow Callout</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftRightArrowCallout'">Left-Right Arrow Callout</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='upDownArrowCallout'">Up-Down Arrow Callout</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='circularArrow'">Circular Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartPreparation'">Preparation</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartManualInput'">Manual Input</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartManualOperation'">Manual Operation</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartConnector'">Connector</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartOffpageConnector'">Off-page Connector</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartPunchedCard'">Card</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartPunchedTape'">Punched Tape</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartSummingJunction'">Summing Junction</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartOr'">Or</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartCollate'">Collate</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartSort'">Sort</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartExtract'">Extract</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartMerge'">Merge</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartDelay'">Delay</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartMagneticDisk'">Magnetic Disk</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartDisplay'">Display</xsl:when>
            <!--zhaobj-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='ribbon2'">Up Ribbon</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='ellipseRibbon'">Curved Down Ribbon</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='ellipseRibbon2'">Curved Up Ribbon</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='verticalScroll'">Vertical Scroll</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='horizontalScroll'">Horizontal Scroll</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='wave'">Wave</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='doubleWave'">Double Wave</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='noSmoking'">No Symbol</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='plus'">Cross</xsl:when>
            <!--2011：加号、乘号、除号转换为十字-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='mathPlus'">Cross</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='mathMinus'">Rectangle</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='mathMultiply'">Cross</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='mathDivide'">Cross</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='irregularSeal1'">Explosion 1</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='irregularSeal2'">Explosion 2</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='wedgeRectCallout'">Rectangular Callout</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='wedgeRoundRectCallout'">Rounded Rectangular Callout</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='wedgeEllipseCallout'">Oval Callout</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='borderCallout1'">Line Callout2</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='borderCallout2'">Line Callout3</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='borderCallout3'">Line Callout4</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='accentCallout1'">Line Callout2(Accent Bar)</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='accentCallout2'">Line Callout3(Accent Bar)</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='accentCallout3'">Line Callout4(Accent Bar)</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='callout1'">Line Callout2(No Border)</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='callout2'">Line Callout3(No Border)</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='callout3'">Line Callout4(No Border)</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='accentBorderCallout1'">Line Callout2(Border and Accent Bar)</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='accentBorderCallout2'">Line Callout3(Border and Accent Bar)</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='accentBorderCallout3'">Line Callout4(Border and Accent Bar)</xsl:when>
            <!--xsl:when test="a:prstGeom/@prst='straightConnector1'">Arrow</xsl:when-->
            <!--xsl:when test="a:prstGeom/@prst='curvedConnector3'">Curve</xsl:when-->
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='bentConnector3'">Elbow Connector</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedConnector3'">Curved Connector</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedConnector4'">Curved Connector</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='homePlate'">Pentagon Arrow</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='quadArrowCallout'">Quad Arrow Callout</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartOnlineStorage'">Stored Data</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartMagneticTape'">Sequential Access Storage</xsl:when>
            <xsl:when test="./wp:spPr/a:prstGeom/@prst='flowChartMagneticDrum'">Direct Access Storage</xsl:when>
            <!---->
            <xsl:when test="./wp:spPr/a:custGeom/*">Curve</xsl:when>
            <xsl:otherwise>Rectangle</xsl:otherwise>
          </xsl:choose>
        </图:名称_801A>
        <图:生成软件_801B>UOFTranslator</图:生成软件_801B>
        <图:属性_801D>
          <xsl:choose>
            <xsl:when test="./wp:spPr/a:gradFill">
              <!--渐变填充-->
              <图:填充_804C>
                <xsl:apply-templates select="wp:spPr/a:gradFill"/>
              </图:填充_804C>
            </xsl:when>
            <xsl:when test="./wp:spPr/a:pattFill">
              <!--图案填充-->
              <图:填充_804C>
                <xsl:apply-templates select="wp:spPr/a:pattFill"/>
              </图:填充_804C>
            </xsl:when>
            <xsl:when test="./wp:spPr/a:solidFill">
              <!--纯色填充-->
              <图:填充_804C>
                <图:颜色_8004>
                  <xsl:apply-templates select="wp:spPr/a:solidFill"/>
                </图:颜色_8004>
              </图:填充_804C>
            </xsl:when>
            <xsl:when test="./wp:spPr/a:blipFill">
              <!--图片填充-->
              <图:填充_804C>
                <xsl:apply-templates select="wp:spPr/a:blipFill">
                  <xsl:with-param name="picturefrom" select="$picFrom"/>
                </xsl:apply-templates>
              </图:填充_804C>
            </xsl:when>
            <xsl:when test="./wp:spPr/a:grpFill">
              <!--组合图形填充-->
              <xsl:apply-templates select="./wp:spPr/a:grpFill"/>
            </xsl:when>
            <xsl:when test="./wp:spPr/a:noFill"/>

            <!--2014-04-29，wudi，修复无背景颜色图片转换后出现白色背景的BUG，start-->
            <xsl:when test="./pic:spPr/a:noFill"/>
            <!--end-->
            
            <!--有style的情况-->
            <!-- <xsl:when test="following-sibling::p:style/a:fillRef">-->
            <xsl:when test="./wp:style/a:fillRef[@idx!='0' and @idx!='1000']">
              <图:填充_804C>
                <图:颜色_8004>
                  <xsl:apply-templates select="wp:style/a:fillRef"/>
                </图:颜色_8004>
              </图:填充_804C>
            </xsl:when>
            <xsl:otherwise>
              <图:填充_804C>
                <图:颜色_8004>auto</图:颜色_8004>
              </图:填充_804C>
            </xsl:otherwise>
          </xsl:choose>
          <图:大小_8060>
            <xsl:attribute name="宽_C605">
              <xsl:if test="./wp:spPr/a:xfrm/a:ext/@cx">
                <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx div 12700"/>
              </xsl:if>
              <xsl:if test="./pic:spPr/a:xfrm/a:ext/@cx">
                <xsl:value-of select="./pic:spPr/a:xfrm/a:ext/@cx div 12700"/>
              </xsl:if>
            </xsl:attribute>
            <xsl:attribute name="长_C604">
              <xsl:if test="./wp:spPr/a:xfrm/a:ext/@cy">
                <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy div 12700"/>
              </xsl:if>
              <xsl:if test="./pic:spPr/a:xfrm/a:ext/@cy">
                <xsl:value-of select="./pic:spPr/a:xfrm/a:ext/@cy div 12700"/>
              </xsl:if>
            </xsl:attribute>
          </图:大小_8060>
          <!-- </xsl:if>-->
          <!--CXL线（包括线颜色、线粗线、线类型三个子元素）-->
          <图:线_8057>
            <xsl:choose>
              <xsl:when test="./wp:spPr/a:ln">
                <xsl:apply-templates select="wp:spPr/a:ln" mode="linecolor"/>
              </xsl:when>

              <!--2014-01-08，wudi，增加图片边框的转换，start-->
              <xsl:when test="./pic:spPr/a:ln">
                <xsl:apply-templates select="pic:spPr/a:ln" mode="linecolor"/>
              </xsl:when>
              <!--end-->

              <xsl:when test="./wp:style/a:lnRef/*">
                <xsl:for-each select="./wp:style/a:lnRef">
                  <图:线颜色_8058>
                    <xsl:call-template name="lnRefcolor"/>
                  </图:线颜色_8058>
                </xsl:for-each>
              </xsl:when>

              <!--2013-03-18，wudi，修复图片转换后出现黑色边框的BUG，start-->
              <!--2013-04-02，wudi，注释掉以下语句，不存在默认线条颜色-->
              <!--<xsl:when test ="not(./pic:spPr/a:ln) and not(./wp:spPr/a:ln) and not(./wp:style/a:lnRef/*)">
                <图:线颜色_8058>#000000</图:线颜色_8058>
              </xsl:when>-->
              <!--end-->

            </xsl:choose>
            <图:线类型_8059>
              <xsl:attribute name="线型_805A">
                <xsl:choose>
                  <xsl:when test="./wp:spPr/a:ln/@cmpd='sng'">single</xsl:when>
                  <xsl:when test="./wp:spPr/a:ln/@cmpd='dbl'">double</xsl:when>
                  <xsl:when test="./wp:spPr/a:ln/@cmpd='thickThin'">thin-thick</xsl:when>
                  <xsl:when test="./wp:spPr/a:ln/@cmpd='thinThick'">thick-thin</xsl:when>
                  <xsl:when test="./wp:spPr/a:ln/@cmpd='tri'">thick-between-thin</xsl:when>

                  <!--2014-01-08，wudi，增加图片边框的转换，start-->
                  <xsl:when test="./pic:spPr/a:ln/@cmpd='sng'">single</xsl:when>
                  <xsl:when test="./pic:spPr/a:ln/@cmpd='dbl'">double</xsl:when>
                  <xsl:when test="./pic:spPr/a:ln/@cmpd='thickThin'">thin-thick</xsl:when>
                  <xsl:when test="./pic:spPr/a:ln/@cmpd='thinThick'">thick-thin</xsl:when>
                  <xsl:when test="./pic:spPr/a:ln/@cmpd='tri'">thick-between-thin</xsl:when>
                  <!--end-->

                  <xsl:otherwise>single</xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
              <xsl:choose>
                <xsl:when test="./wp:spPr/a:ln/a:prstDash/@val">
                  <xsl:attribute name="虚实_805B">
                    <xsl:choose>
                      <xsl:when test="./wp:spPr/a:ln/a:prstDash/@val='solid'">
                        <xsl:value-of select="'solid'"/>
                      </xsl:when>
                      <!--圆点线-->
                      <xsl:when test="./wp:spPr/a:ln/a:prstDash/@val='sysDot'">
                        <xsl:value-of select="'round-dot'"/>
                      </xsl:when>
                      <!--方点线-->
                      <xsl:when test="./wp:spPr/a:ln/a:prstDash/@val='sysDash'">
                        <xsl:value-of select="'square-dot'"/>
                      </xsl:when>
                      <!--短划线->虚线-->
                      <xsl:when test="./wp:spPr/a:ln/a:prstDash/@val='dash'">
                        <xsl:value-of select="'dash'"/>
                      </xsl:when>
                      <!--划线-点->点虚线-->
                      <xsl:when test="./wp:spPr/a:ln/a:prstDash/@val='dashDot'">
                        <xsl:value-of select="'dash-dot'"/>
                      </xsl:when>
                      <!--长划线->长虚线-->
                      <xsl:when test="./wp:spPr/a:ln/a:prstDash/@val='lgDash'">
                        <xsl:value-of select="'long-dash'"/>
                      </xsl:when>
                      <!--长划线-点-->
                      <xsl:when test="./wp:spPr/a:ln/a:prstDash/@val='lgDashDot'">
                        <xsl:value-of select="'long-dash-dot'"/>
                      </xsl:when>
                      <!--长划线-点-点-->
                      <xsl:when test="./wp:spPr/a:ln/a:prstDash/@val='lgDashDotDot'">
                        <xsl:value-of select="'dash-dot-dot'"/>
                      </xsl:when>
                      <!--其它情况处理>
                      <xsl:when test="a:prstDash/@val='sysDashDot'">
                        <xsl:value-of select="'dash-dot-heavy'"/>
                      </xsl:when>
                      <xsl:when test="a:prstDash/@val='dot'">
                        <xsl:value-of select="'dotted'"/>
                      </xsl:when>
                      <xsl:when test="a:prstDash/@val='sysDashDotDot'">
                        <xsl:value-of select="'dash-dot-dot-heavy'"/>
                      </xsl:when-->
                    </xsl:choose>
                  </xsl:attribute>
                </xsl:when>

                <!--2014-01-08，wudi，增加图片边框的转换，start-->
                <xsl:when test="./pic:spPr/a:ln/a:prstDash/@val">
                  <xsl:attribute name="虚实_805B">
                    <xsl:choose>
                      <xsl:when test="./pic:spPr/a:ln/a:prstDash/@val='solid'">
                        <xsl:value-of select="'solid'"/>
                      </xsl:when>
                      <!--圆点线-->
                      <xsl:when test="./pic:spPr/a:ln/a:prstDash/@val='sysDot'">
                        <xsl:value-of select="'round-dot'"/>
                      </xsl:when>
                      <!--方点线-->
                      <xsl:when test="./pic:spPr/a:ln/a:prstDash/@val='sysDash'">
                        <xsl:value-of select="'square-dot'"/>
                      </xsl:when>
                      <!--短划线->虚线-->
                      <xsl:when test="./pic:spPr/a:ln/a:prstDash/@val='dash'">
                        <xsl:value-of select="'dash'"/>
                      </xsl:when>
                      <!--划线-点->点虚线-->
                      <xsl:when test="./pic:spPr/a:ln/a:prstDash/@val='dashDot'">
                        <xsl:value-of select="'dash-dot'"/>
                      </xsl:when>
                      <!--长划线->长虚线-->
                      <xsl:when test="./pic:spPr/a:ln/a:prstDash/@val='lgDash'">
                        <xsl:value-of select="'long-dash'"/>
                      </xsl:when>
                      <!--长划线-点-->
                      <xsl:when test="./pic:spPr/a:ln/a:prstDash/@val='lgDashDot'">
                        <xsl:value-of select="'long-dash-dot'"/>
                      </xsl:when>
                      <!--长划线-点-点-->
                      <xsl:when test="./pic:spPr/a:ln/a:prstDash/@val='lgDashDotDot'">
                        <xsl:value-of select="'dash-dot-dot'"/>
                      </xsl:when>
                    </xsl:choose>
                  </xsl:attribute>
                </xsl:when>
                <!--end-->
                
                <xsl:otherwise>
                  <xsl:attribute name="虚实_805B">solid</xsl:attribute>
                </xsl:otherwise>
              </xsl:choose>
            </图:线类型_8059>
            <xsl:if test="name(.)='wp:wsp'">
              <xsl:choose>
                <xsl:when test="./wp:spPr/a:ln/@w">
                  <图:线粗细_805C>
                    <xsl:value-of select="./wp:spPr/a:ln/@w div 12700"/>
                  </图:线粗细_805C>
                </xsl:when>
                <xsl:otherwise>

                  <!--2012-11-22，wudi，修正对预定义图形线粗细的不正确处理，start-->
                  <图:线粗细_805C>
                    <xsl:choose>
                      <xsl:when test="./wp:spPr/a:prstGeom/@prst='line'">0.75</xsl:when>
                      <!--直接连接符 2-->
                      <xsl:when test="./wp:spPr/a:prstGeom/@prst='straightConnector1'">0.75</xsl:when>
                      <!--直接箭头连接符 3，直接箭头连接符 4-->
                      <xsl:when test="./wp:spPr/a:prstGeom/@prst='arc'">0.75</xsl:when>
                      <!--弧形-->
                      <xsl:when test="./wp:spPr/a:prstGeom/@prst='bracketPair'">0.75</xsl:when>
                      <!--双括号 61-->
                      <xsl:when test="./wp:spPr/a:prstGeom/@prst='bracePair'">0.75</xsl:when>
                      <!--双大括号 62-->
                      <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftBracket'">0.75</xsl:when>
                      <!--左中括号 66-->
                      <xsl:when test="./wp:spPr/a:prstGeom/@prst='rightBracket'">0.75</xsl:when>
                      <!--右中括号 63-->
                      <xsl:when test="./wp:spPr/a:prstGeom/@prst='leftBrace'">0.75</xsl:when>
                      <!--左大括号 64-->
                      <xsl:when test="./wp:spPr/a:prstGeom/@prst='rightBrace'">0.75</xsl:when>
                      <!--右大括号 65-->
                      <xsl:when test="./wp:spPr/a:prstGeom/@prst='bentConnector3'">0.75</xsl:when>
                      <!--肘形连接符 5，肘形连接符6，肘形连接符 7-->
                      <xsl:when test="./wp:spPr/a:prstGeom/@prst='curvedConnector3'">0.75</xsl:when>
                      <!--曲线连接符8，曲线连接符9，曲线连接符11-->
                      <xsl:otherwise>2.0</xsl:otherwise>
                    </xsl:choose>
                  </图:线粗细_805C>
                  <!--end-->

                  <!--<图:线粗细_805C>2.0</图:线粗细_805C>-->
                  <!--原来是0.75，但发现WORD里预设为2磅-->
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
            <xsl:if test="name(.)='pic:pic'">
              <xsl:choose>
                <xsl:when test="./pic:spPr/a:ln/@w">
                  <!--cxl2011/12/9添加图片边框粗细-->
                  <图:线粗细_805C>
                    <xsl:value-of select="./pic:spPr/a:ln/@w div 12700"/>
                  </图:线粗细_805C>
                </xsl:when>
                <xsl:otherwise>
                  <图:线粗细_805C>0.0</图:线粗细_805C>
                  <!--原来是0.75，但发现WORD里预设为0.0磅-->
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
          </图:线_8057>
          <xsl:choose>
            <xsl:when test="./wp:spPr/a:xfrm/@rot">
              <图:旋转角度_804D>
                <!-- 旋转角度<xsl:value-of select="(21600000-a:xfrm/@rot) div 60000"/>-->
                <xsl:value-of select="./wp:spPr/a:xfrm/@rot div 60000"/>
              </图:旋转角度_804D>
            </xsl:when>
            <xsl:otherwise>
              <图:旋转角度_804D>0.0</图:旋转角度_804D>
            </xsl:otherwise>
          </xsl:choose>
          <!--zhaobj 箭头-->
          <xsl:if test="./wp:spPr/a:ln">
            <xsl:apply-templates select="wp:spPr/a:ln" mode="Arrow"/>
          </xsl:if>
          <!--<xsl:if test="wp:cNvGraphicFramePr/a:graphicFrameLocks/@noChangeAspect='1'">
            <图:锁定纵横比 uof:locID="g0028">0</图:锁定纵横比>
          </xsl:if>-->

          <图:缩放是否锁定纵横比_8055>
            <!--cxl,2012.3.6修改BUG-->
            <xsl:if test="$picFrom = 'grpsp'">
              <xsl:if test="./wp:cNvSpPr/a:spLocks/@noChangeAspect='1'">
                <xsl:value-of select="'true'"/>
              </xsl:if>
              <xsl:if test="./wp:cNvSpPr/a:spLocks/@noChangeAspect='0' or not(./wp:cNvSpPr/a:spLocks/@noChangeAspect)">
                <xsl:value-of select="'false'"/>
              </xsl:if>
            </xsl:if>
            <xsl:if test="$picFrom != 'grpsp'">
              <xsl:choose>
                <xsl:when test="./wp:cNvSpPr/a:spLocks/@noChangeAspect='1'">
                  <!--/../@locked-->
                  <xsl:value-of select="'true'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="'false'"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
          </图:缩放是否锁定纵横比_8055>
          <图:是否打印对象_804E>true</图:是否打印对象_804E>
          <!--<图:Web文字 uof:locID="g0033"/>-->
          <!--zhaobj 透明度a:solidFill 还有 a:gradFill a:pattFill a:blipFill的情况-->
          
          <!--2013-11-27，wudi，Strict标准下，val取值有变，修正，start-->
          <xsl:if test="./wp:spPr/a:ln/a:solidFill//a:alpha">
            <图:透明度_8050>
              <xsl:variable name="tmp">
                <xsl:value-of select="number(substring-before(./wp:spPr/a:ln/a:solidFill//a:alpha/@val,'%'))"/>
              </xsl:variable>
              <xsl:value-of select="round(100 - $tmp)"/>
            </图:透明度_8050>
          </xsl:if>
          <!--end-->

          <!--2012-01-25，wudi，OOX到UOF方向修复BUG：纯色填充的颜色透明度效果丢失，start-->
          
          <!--2013-11-27，wudi，Strict标准下，val取值有变，修正，start-->
          <xsl:if test="./wp:spPr/a:solidFill//a:alpha">
            <图:透明度_8050>
              <xsl:variable name="tmp">
                <xsl:value-of select="number(substring-before(./wp:spPr/a:solidFill//a:alpha/@val,'%'))"/>
              </xsl:variable>
              <xsl:value-of select="round(100 - $tmp)"/>
            </图:透明度_8050>
          </xsl:if>
          <!--end-->
          
          <!--end-->

          <xsl:if test="./wp:spPr/a:effectLst/a:outerShdw">
            <!--外部阴影-->
            <xsl:apply-templates select="wp:spPr/a:effectLst/a:outerShdw" mode="shadow"/>
          </xsl:if>
          <xsl:if test="./wp:spPr/a:effectLst/a:innerShdw">
            <!--内部阴影-->
            <xsl:apply-templates select="wp:spPr/a:effectLst/a:innerShdw" mode="shadow"/>
          </xsl:if>
          
          <!--2014-04-16，wudi，增加图片阴影效果的实现，start-->
          <xsl:if test="./pic:spPr/a:effectLst/a:outerShdw">
            <!--外部阴影-->
            <xsl:apply-templates select="pic:spPr/a:effectLst/a:outerShdw" mode="shadow"/>
          </xsl:if>
          <xsl:if test="./pic:spPr/a:effectLst/a:innerShdw">
            <!--内部阴影-->
            <xsl:apply-templates select="pic:spPr/a:effectLst/a:innerShdw" mode="shadow"/>
          </xsl:if>
          <!--end-->

          <!--cxl2011/11/20开始增加三维效果转化代码*****************************************************************************-->
          <xsl:if test="./wp:spPr/a:scene3d">
            <图:三维效果_8061>
              <xsl:apply-templates select="./wp:spPr/a:scene3d"/>
            </图:三维效果_8061>
          </xsl:if>
        </图:属性_801D>
      </图:预定义图形_8018>
      <!--cxl2011/11/19添加图片数据引用,其它数据引用按范围文档不转************************************************-->
      <xsl:if test="./pic:blipFill">
        <!--./wp:spPr/a:blipFill|-->
        <图:图片数据引用_8037>
          <xsl:variable name="picId">
            <xsl:if test="./pic:blipFill/a:blip/@r:embed">
              <xsl:value-of select="./pic:blipFill/a:blip/@r:embed"/>
            </xsl:if>
            <!--<xsl:if test="./wp:spPr/a:blipFill/a:blip/@r:embed">-->
            <!--待做案例验证-->
            <!--
              <xsl:value-of select="./wp:spPr/a:blipFill/a:blip/@r:embed"/>
            </xsl:if>-->
          </xsl:variable>
          <xsl:variable name="picName">
            <xsl:choose>
              <xsl:when test="$picFrom='document'">
                <xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="$picFrom='comments'">
                <xsl:value-of select="document('word/_rels/comments.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="$picFrom='endnotes'">
                <xsl:value-of select="document('word/_rels/endnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="$picFrom='footnotes'">
                <xsl:value-of select="document('word/_rels/footnotes.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="contains($picFrom,'header')">
                <xsl:variable name="hn" select="concat('word/_rels/',$picFrom,'.xml.rels')"/>
                <xsl:value-of select="document($hn)/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <xsl:when test="contains($picFrom,'footer')">
                <xsl:variable name="fn" select="concat('word/_rels/',$picFrom,'.xml.rels')"/>
                <xsl:value-of select="document($fn)/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
              <!--暂时先这样-->
              <xsl:when test="$picFrom='grpsp'">
                <xsl:value-of select="document('word/_rels/document.xml.rels')/rel:Relationships/rel:Relationship[@Id=$picId]/@Target"/>
              </xsl:when>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="pictype">
            <xsl:value-of select="substring-after($picName,'.')"/>
          </xsl:variable>
          <xsl:if test="$picFrom = 'grpsp'">
            <xsl:if test="wp:cNvPr/@id">
              <xsl:value-of select="concat('grpsppicObj',wp:cNvPr/@id)"/>
            </xsl:if>
            <xsl:if test="wpg:cNvPr/@id">
              <xsl:value-of select="concat('grpsppicObj',wpg:cNvPr/@id)"/>
            </xsl:if>
          </xsl:if>
          <xsl:if test="$picFrom != 'grpsp'">
            <xsl:value-of select="concat($picFrom,'Obj',$number * 2)"/>
            <!--substring-after($picName,'/')-->
          </xsl:if>
        </图:图片数据引用_8037>
      </xsl:if>

      <!--2013-12-03，wudi，预定义图形-控制点的转换，start-->
      <xsl:choose>
        <!--椭圆形标注，圆角矩形标注，矩形标注，云形标注-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='wedgeEllipseCallout' or ./wp:spPr/a:prstGeom/@prst='wedgeRoundRectCallout' or ./wp:spPr/a:prstGeom/@prst='wedgeRectCallout' or ./wp:spPr/a:prstGeom/@prst='cloudCallout') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x * 21600) div 100000 + 10800"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y * 21600) div 100000 + 10800"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--线形标注1，线形标注1（带强调线），线形标注1（无边框线），线形标注1（带边框和强调线）-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='borderCallout1' or ./wp:spPr/a:prstGeom/@prst='accentCallout1' or ./wp:spPr/a:prstGeom/@prst='callout1' or ./wp:spPr/a:prstGeom/@prst='accentBorderCallout1') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj4">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp4">
            <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp4)"/>
          </xsl:variable>
          <xsl:variable name="y2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x2 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y2 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x1 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y1 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--线形标注2，线形标注2（带强调线），线形标注2（无边框线），线形标注2（带边框和强调线）-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='borderCallout2' or ./wp:spPr/a:prstGeom/@prst='accentCallout2' or ./wp:spPr/a:prstGeom/@prst='callout2' or ./wp:spPr/a:prstGeom/@prst='accentBorderCallout2') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj4">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj5">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj5']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj6">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj6']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp4">
            <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp5">
            <xsl:value-of select="normalize-space(substring-after($adj5,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp6">
            <xsl:value-of select="normalize-space(substring-after($adj6,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp4)"/>
          </xsl:variable>
          <xsl:variable name="y2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="x3">
            <xsl:value-of select="number($tmp6)"/>
          </xsl:variable>
          <xsl:variable name="y3">
            <xsl:value-of select="number($tmp5)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x3 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y3 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x2 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y2 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x1 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y1 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--线形标注3，线形标注3（带强调线），线形标注3（无边框线），线形标注3（带边框和强调线）-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='borderCallout3' or ./wp:spPr/a:prstGeom/@prst='accentCallout3' or ./wp:spPr/a:prstGeom/@prst='callout3' or ./wp:spPr/a:prstGeom/@prst='accentBorderCallout3') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj4">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj5">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj5']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj6">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj6']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj7">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj7']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj8">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj8']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp4">
            <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp5">
            <xsl:value-of select="normalize-space(substring-after($adj5,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp6">
            <xsl:value-of select="normalize-space(substring-after($adj6,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp7">
            <xsl:value-of select="normalize-space(substring-after($adj7,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp8">
            <xsl:value-of select="normalize-space(substring-after($adj8,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp4)"/>
          </xsl:variable>
          <xsl:variable name="y2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="x3">
            <xsl:value-of select="number($tmp6)"/>
          </xsl:variable>
          <xsl:variable name="y3">
            <xsl:value-of select="number($tmp5)"/>
          </xsl:variable>
          <xsl:variable name="x4">
            <xsl:value-of select="number($tmp8)"/>
          </xsl:variable>
          <xsl:variable name="y4">
            <xsl:value-of select="number($tmp7)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x4 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y4 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x3 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y3 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x2 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y2 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x1 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y1 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--十字星，八角星，十六角星，二十四角星，三十二角星-->
        <!--五角星OOX有控制点，UOF无控制点-->
        <!--六角星，七角星转成八角星，十角星，十二角星转成十六角星-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='star4' or ./wp:spPr/a:prstGeom/@prst='star8' or ./wp:spPr/a:prstGeom/@prst='star16' or ./wp:spPr/a:prstGeom/@prst='star24' or ./wp:spPr/a:prstGeom/@prst='star32' or ./wp:spPr/a:prstGeom/@prst='star6' or ./wp:spPr/a:prstGeom/@prst='star7' or ./wp:spPr/a:prstGeom/@prst='star10' or ./wp:spPr/a:prstGeom/@prst='star12') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp">
            <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="10800 - ($x * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--上凸带形-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='ribbon2') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="10800 - ($x * 21600) div 200000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="21600 - ($y * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--前凸带形-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='ribbon') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="10800 - ($x * 21600) div 200000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--上凸弯带形-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='ellipseRibbon2') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:choose>
              <xsl:when test="number($tmp2) = 100000">
                <xsl:value-of select="number($tmp2) - 25000"/>
              </xsl:when>
              <xsl:when test="number($tmp2) = 25000">
                <xsl:value-of select="number($tmp2) + 4"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select ="number($tmp2)"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="10800 - ($x1 * 21600) div 200000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="21600 - ($y1 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x2 * 21600) div 100000 + 675"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--前凸弯带形-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='ellipseRibbon') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select ="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="10800 - ($x1 * 21600) div 200000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y1 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - ($x2 * 21600) div 100000 - 675"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--竖卷形，横卷形-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='verticalScroll' or ./wp:spPr/a:prstGeom/@prst='horizontalScroll') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp">
            <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--波形，双波形-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='wave' or ./wp:spPr/a:prstGeom/@prst='doubleWave') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 + ($y * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--右箭头，虚尾箭头，燕尾形箭头-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='rightArrow' or ./wp:spPr/a:prstGeom/@prst='stripedRightArrow' or ./wp:spPr/a:prstGeom/@prst='notchedRightArrow') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - ($x * 21600 * $height) div $width div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y * 10800) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--左箭头，左右箭头-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='leftArrow' or ./wp:spPr/a:prstGeom/@prst='leftRightArrow') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x * 21600 * $height) div $width div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y * 10800) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--上箭头-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='upArrow') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($y * 21600 * $width) div $height div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($x * 10800) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--上下箭头-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='upDownArrow') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="10800 - ($x * 10800) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="($y * 21600 * $width) div $height div 100000 "/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--下箭头-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='downArrow') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - ($y * 21600 * $width) div $height div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($x * 10800) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--十字箭头，丁字箭头-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='quadArrow' or ./wp:spPr/a:prstGeom/@prst='leftRightUpArrow') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="10800 - ($x1 * 10800) div 50000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y1 * 10800) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$x2 * 10800 div 50000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--圆角右箭头-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='bentArrow') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - ($x * $height * 21600) div $width div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="6079 - ($y * $height * 10800) div $width div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--直角双向箭头，直角上箭头-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='leftUpArrow' or ./wp:spPr/a:prstGeom/@prst='bentUpArrow') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - ($x1 * $height * 21600) div $width div 50000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="21600 - ($x1 * $height * 21600) div $width div 100000 + ($y1 * $height * 21600) div $width div 200000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$x2 * 10800 div 50000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--左弧形箭头，显示效果有差异-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='curvedRightArrow') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - $x1 * $width *  21600 div $height div 100000"/>-->
              <xsl:value-of select="21600 - $x1 * $width *  21600 div $height div 100000 + 250"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - ($x1 - $y1) * $width * 21600 div $height div 200000"/>-->
              <xsl:value-of select="21600 - ($x1 - $y1) * $width * 21600 div $height div 200000 + 250"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - $x2 * 21600 div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--右弧形箭头，显示效果有差异-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='curvedLeftArrow') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - $x1 * $width *  21600 div $height div 100000"/>-->
              <xsl:value-of select="21600 - $x1 * $width *  21600 div $height div 100000 + 250"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - ($x1 - $y1) * $width * 21600 div $height div 200000"/>-->
              <xsl:value-of select="21600 - ($x1 - $y1) * $width * 21600 div $height div 200000 + 250"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$x2 * 21600 div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--下弧形箭头，显示效果有差异-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='curvedUpArrow') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - $x1 * $height *  21600 div $width div 100000"/>-->
              <xsl:value-of select="21600 - $x1 * $height *  21600 div $width div 100000 + 10"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - ($x1 - $y1) * $height * 21600 div $width div 200000"/>-->
              <xsl:value-of select="21600 - ($x1 - $y1) * $height * 21600 div $width div 200000 + 10"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$x2 * 21600 div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--上弧形箭头，显示效果有差异-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='curvedDownArrow') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - $x1 * $height *  21600 div $width div 100000"/>-->
              <xsl:value-of select="21600 - $x1 * $height *  21600 div $width div 100000 + 10"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <!--<xsl:value-of select="21600 - $x2 * $x2 * 21600 div 100000 div 100000 div 2 - ($x1 - $y1) * $height * 21600 div $width div 200000"/>-->
              <xsl:value-of select="21600 - ($x1 - $y1) * $height * 21600 div $width div 200000 + 10"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - $x2 * 21600 div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--五边形，燕尾形-->
        <xsl:when test ="(./wp:spPr/a:prstGeom/@prst='homePlate' or ./wp:spPr/a:prstGeom/@prst='chevron') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp">
            <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - $x * 21600 * $height div $width div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--右箭头标注-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='rightArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj4">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp4">
            <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp4)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="y2">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x1 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y1 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - ($x2 * 21600 * $height) div $width div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y2 * 21600) div 200000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--左箭头标注-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='leftArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj4">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp4">
            <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp4)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="y2">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - ($x1 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y1 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x2 * 21600 * $height) div $width div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y2 * 21600) div 200000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--上箭头标注-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='upArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj4">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp4">
            <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp4)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="y2">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - ($x1 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y1 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x2 * 21600 * $width) div $height div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y2 * 21600) div 200000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--下箭头标注-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='downArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj4">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp4">
            <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp4)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="y2">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="($x1 * 21600) div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y1 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - ($x2 * 21600 * $width) div $height div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y2 * 21600) div 200000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--左右箭头标注-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='leftRightArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj4">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp4">
            <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp4)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="y2">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="10800 - ($x1 * 21600) div 200000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y1 * 21600) div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$x2 * 21600 * $height div $width div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y2 * 21600) div 200000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--十字箭头标注-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='quadArrowCallout') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj3">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj3']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj4">
            <xsl:value-of select ="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj4']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp3">
            <xsl:value-of select="normalize-space(substring-after($adj3,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp4">
            <xsl:value-of select="normalize-space(substring-after($adj4,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x1">
            <xsl:value-of select="number($tmp4)"/>
          </xsl:variable>
          <xsl:variable name="y1">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="x2">
            <xsl:value-of select="number($tmp3)"/>
          </xsl:variable>
          <xsl:variable name="y2">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="10800 - ($x1 * 21600) div 200000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y1 * 21600 * $height) div $width div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$x2 * 21600 div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="10800 - ($y2 * 21600 * $height) div $width div 200000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--等腰三角形，八边形，十字型，缺角矩形，立方体，棱台，太阳形，新月形，双括号，双大括号-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='triangle' or ./wp:spPr/a:prstGeom/@prst='octagon' or ./wp:spPr/a:prstGeom/@prst='plus' or ./wp:spPr/a:prstGeom/@prst='plaque' or ./wp:spPr/a:prstGeom/@prst='cube' or ./wp:spPr/a:prstGeom/@prst='bevel' or ./wp:spPr/a:prstGeom/@prst='sun' or ./wp:spPr/a:prstGeom/@prst='moon' or ./wp:spPr/a:prstGeom/@prst='bracketPair' or ./wp:spPr/a:prstGeom/@prst='bracePair') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp">
            <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$x * 21600 div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--平行四边形，梯形，六边形，同心圆，禁止符-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='parallelogram' or ./wp:spPr/a:prstGeom/@prst='trapezoid' or ./wp:spPr/a:prstGeom/@prst='hexagon' or ./wp:spPr/a:prstGeom/@prst='donut' or ./wp:spPr/a:prstGeom/@prst='noSmoking') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp">
            <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$x * 21600 * $height div $width div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--圆柱形-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='can') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp">
            <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$y * 21600 div 200000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--折角形-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='foldedCorner') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp">
            <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="21600 - $x * 21600 div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--笑脸-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='smileyFace') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp">
            <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="16200 + $y * 21600 div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--左中括号，右中括号-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='leftBracket' or ./wp:spPr/a:prstGeom/@prst='rightBracket') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp">
            <xsl:value-of select="normalize-space(substring-after($adj,'val'))"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$y * 21600 * $width div $height div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--左大括号，右大括号-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='leftBrace' or ./wp:spPr/a:prstGeom/@prst='rightBrace') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="adj2">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj2']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp1">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="tmp2">
            <xsl:value-of select="normalize-space(substring-after($adj2,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp1)"/>
          </xsl:variable>
          <xsl:variable name="y">
            <xsl:value-of select="number($tmp2)"/>
          </xsl:variable>
          <xsl:variable name="width">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cx"/>
          </xsl:variable>
          <xsl:variable name="height">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:ext/@cy"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$x * 21600 * $width div $height div 100000"/>
            </xsl:attribute>
            <xsl:attribute name="y_C607">
              <xsl:value-of select="$y * 21600 div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
        <!--肘形连接符，肘形箭头连接符，肘形双箭头连接符，曲线连接符，曲线箭头连接符，曲线双箭头连接符-->
        <xsl:when test="(./wp:spPr/a:prstGeom/@prst='bentConnector3' or ./wp:spPr/a:prstGeom/@prst='curvedConnector3') and ./*/a:prstGeom/a:avLst/a:gd">
          <xsl:variable name="adj1">
            <xsl:value-of select="./*/a:prstGeom/a:avLst/a:gd[@name = 'adj1']/@fmla"/>
          </xsl:variable>
          <xsl:variable name="tmp">
            <xsl:value-of select="normalize-space(substring-after($adj1,'val'))"/>
          </xsl:variable>
          <xsl:variable name="x">
            <xsl:value-of select="number($tmp)"/>
          </xsl:variable>
          <图:控制点_8039>
            <xsl:attribute name="x_C606">
              <xsl:value-of select="$x * 21600 div 100000"/>
            </xsl:attribute>
          </图:控制点_8039>
        </xsl:when>
      </xsl:choose>
      <!--end-->
      
      <!--翻转，cxl，2012.3.9互换x,y位置-->

      <!--2013-03-26，wudi，修复预定义图形方向不正确的BUG，此差异与永中软件有关，start-->
      <!--2013-03-28，wudi，修复功能测试OOX到UOF2方向ShapeStyle右边最下行的黄色矩形方向与原始不一致的BUG，增加限制条件Rectangular Callout-->
      <!--<xsl:choose>-->
        <!--2013-12-05，wudi，删除此处代码
        <xsl:when test ="./wp:spPr/a:prstGeom/@prst='borderCallout3' and not(./*/a:xfrm/@flipH)">
          <图:翻转_803A>
            <xsl:value-of select="'x'"/>
          </图:翻转_803A>
        </xsl:when>
        <xsl:when test ="./wp:spPr/a:prstGeom/@prst='accentCallout3' and not(./*/a:xfrm/@flipH)">
          <图:翻转_803A>
            <xsl:value-of select="'x'"/>
          </图:翻转_803A>
        </xsl:when>
        <xsl:when test ="./wp:spPr/a:prstGeom/@prst='callout3' and not(./*/a:xfrm/@flipH)">
          <图:翻转_803A>
            <xsl:value-of select="'x'"/>
          </图:翻转_803A>
        </xsl:when>
        <xsl:when test ="./wp:spPr/a:prstGeom/@prst='accentBorderCallout3' and not(./*/a:xfrm/@flipH)">
          <图:翻转_803A>
            <xsl:value-of select="'x'"/>
          </图:翻转_803A>
        </xsl:when>
        -->
        <!--2013-12-04，wudi，删除此处代码
        <xsl:when test ="./wp:spPr/a:prstGeom/@prst='wedgeRectCallout' and not(./*/a:xfrm/@flipH)">
          <图:翻转_803A>
            <xsl:value-of select="'x'"/>
          </图:翻转_803A>
        </xsl:when>
        -->
      <!--</xsl:choose>-->

      <xsl:if test="./*/a:xfrm/@flipH and not(./wp:spPr/a:prstGeom/@prst='borderCallout3') and not(./wp:spPr/a:prstGeom/@prst='accentCallout3') and not(./wp:spPr/a:prstGeom/@prst='callout3') and not(./wp:spPr/a:prstGeom/@prst='accentBorderCallout3')">
        <图:翻转_803A>
          <xsl:value-of select="'x'"/>
        </图:翻转_803A>
      </xsl:if>
      <xsl:if test="./*/a:xfrm/@flipV">
        <图:翻转_803A>
          <xsl:value-of select="'y'"/>
        </图:翻转_803A>
      </xsl:if>
      <!--end-->

      <xsl:for-each select="./wp:txbx">
        <xsl:apply-templates select="." mode="文本"/>
        <!--预定义图形中文本的转化-->
      </xsl:for-each>
      <xsl:if test="$picFrom ='grpsp'">
        <图:组合位置_803B>
          <xsl:attribute name="x_C606">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:off/@x div 12700"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="./wp:spPr/a:xfrm/a:off/@y div 12700"/>
          </xsl:attribute>
        </图:组合位置_803B>
      </xsl:if>
    </图:图形_8062>
  </xsl:template>
  <!--end-->
  
	<!--以下代码在paragraph.xsl中被调用，CXL 2011/12/5修改-->
	<xsl:template name="bodyAnchorYdy">
    <xsl:param name="filename"/>
		<xsl:variable name="number">
			<xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline | v:shape"/>
		</xsl:variable>
		<uof:锚点_C644>
      
      <!--2014-03-25，wudi，修改图形引用_C62E的取值方法，页眉页脚调整，start-->
      <xsl:attribute name="图形引用_C62E">
        <xsl:choose>
          <xsl:when test="$filename='document'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="$filename='comments'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="$filename='endnotes'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="$filename='footnotes'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          
          <!--2014-04-09，wudi，还原取值方法，新的取值方法会带来新的问题，影响页眉页脚转换，start-->
          <xsl:when test="contains($filename,'header')">
            <xsl:variable name="hn" select="substring-before($filename,'.xml')"/>
            <xsl:value-of select="concat($hn,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="contains($filename,'footer')">
            <xsl:variable name="fn" select="substring-before($filename,'.xml')"/>
            <xsl:value-of select="concat($fn,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <!--end-->
          
        </xsl:choose>
      </xsl:attribute>
      <!--end-->
      
			<!--<xsl:attribute name="字:类型">
				<xsl:if test="w10:wrap[type='none'] and w10:anchorlock">
					<xsl:value-of select="'inline'"/>
				</xsl:if>
				<xsl:if test="not(w10:wrap[type='none'] and w10:anchorlock)">
					<xsl:value-of select="'normal'"/>
				</xsl:if>
			</xsl:attribute>-->
      <xsl:variable name="style" select="@style"/>
      <uof:位置_C620>
				<xsl:variable name="H-temp">
					<xsl:if test="contains(@style,'mso-position-horizontal')">
						<xsl:variable name="temp" select="substring-after($style,'mso-position-horizontal:')"/>
						<xsl:value-of select="substring-before($temp,';')"/>
					</xsl:if>
					<xsl:if test="not (contains(@style,'mso-position-horizontal'))">
						<xsl:value-of select="'none'"/>
					</xsl:if>
				</xsl:variable>
				<xsl:variable name="HR-temp">
					<xsl:if test="contains(@style,'mso-position-horizontal-relative')">
						<xsl:variable name="temp" select="substring-after($style,'mso-position-horizontal-relative:')"/>
						<xsl:value-of select="substring-before($temp,';')"/>
					</xsl:if>
					<xsl:if test="not (contains(@style,'mso-position-horizontal-relative'))">
						<xsl:value-of select="'none'"/>
					</xsl:if>
				</xsl:variable>
				<xsl:variable name="V-temp">
					<xsl:if test="contains(@style,'mso-position-vertical')">
						<xsl:variable name="temp" select="substring-after($style,'mso-position-vertical:')"/>
						<xsl:value-of select="substring-before($temp,';')"/>
					</xsl:if>
					<xsl:if test="not (contains(@style,'mso-position-vertical'))">
						<xsl:value-of select="'none'"/>
					</xsl:if>
				</xsl:variable>
				<xsl:variable name="VR-temp">
					<xsl:if test="contains(@style,'mso-position-vertical-relative')">
						<xsl:variable name="temp" select="substring-after($style,'mso-position-vertical-relative:')"/>
            
            <!--2015-01-07，wudi，修复了位置-垂直转换BUG，start-->
            <xsl:if test="contains($temp,';')">
              <xsl:value-of select="substring-before($temp,';')"/>
            </xsl:if>
            <xsl:if test="not(contains($temp,';'))">
              <xsl:value-of select="$temp"/>
            </xsl:if>
            <!--end-->
            
					</xsl:if>
					<xsl:if test="not (contains(@style,'mso-position-vertical-relative'))">
						<xsl:value-of select="'none'"/>
					</xsl:if>
				</xsl:variable>
				<xsl:variable name="ML-temp">
					<xsl:variable name="temp" select="substring-after($style,'margin-left:')"/>
					<xsl:value-of select="substring-before($temp,';')"/>
				</xsl:variable>
				<xsl:variable name="MT-temp">
					<xsl:variable name="temp" select="substring-after($style,'margin-top:')"/>
					<xsl:value-of select="substring-before($temp,';')"/>
				</xsl:variable>
				<uof:水平_4106>
					<xsl:attribute name="相对于_410C">
						<xsl:choose>
							<xsl:when test="$HR-temp='margin'">
								<xsl:value-of select="'margin'"/>
							</xsl:when>
							<xsl:when test="$HR-temp='page'">
								<xsl:value-of select="'page'"/>
							</xsl:when>
							<xsl:when test="$HR-temp='char'">
								<xsl:value-of select="'char'"/>
							</xsl:when>
							<xsl:when test="$HR-temp='text'">
								<xsl:value-of select="'column'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'column'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
					<xsl:if test="not (contains(@style,'mso-position-horizontal'))">
						<uof:绝对_4107>
							<xsl:attribute name="值_4108">

                <!--2013-04-19，wudi，增加对对象集-图片，图表对象的处理，start-->
                
                <!--2014-03-26，wudi，增加w:pict的情况，start-->
                <xsl:if test ="parent::w:object or parent::w:pict">
                  <xsl:value-of select ="'0'"/>
                </xsl:if>
                <xsl:if test ="not(parent::w:object or parent::w:pict)">
                  <xsl:value-of select="$ML-temp"/>
                </xsl:if>
                <!--end-->
                
                <!--end-->

              </xsl:attribute>
						</uof:绝对_4107>
					</xsl:if>
					<xsl:if test="contains(@style,'mso-position-horizontal')">
            <uof:相对_4109>
							<xsl:attribute name="参考点_410A">
								<xsl:choose>
									<xsl:when test="$H-temp='center'">
										<xsl:value-of select="'center'"/>
									</xsl:when>
									<xsl:when test="$H-temp='left'">
										<xsl:value-of select="'left'"/>
									</xsl:when>
									<xsl:when test="$H-temp='right'">
										<xsl:value-of select="'right'"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="'left'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<!--<xsl:attribute name="字:值">
								<xsl:value-of select="$ML-temp"/>
							</xsl:attribute>-->
						</uof:相对_4109>
					</xsl:if>
				</uof:水平_4106>
        <uof:垂直_410D>
					<xsl:attribute name="相对于_C647">
						<xsl:choose>
							<xsl:when test="$VR-temp='margin'">
								<xsl:value-of select="'margin'"/>
							</xsl:when>
							<xsl:when test="$VR-temp='page'">
								<xsl:value-of select="'page'"/>
							</xsl:when>
							<xsl:when test="$VR-temp='line'">
								<xsl:value-of select="'line'"/>
							</xsl:when>
							<xsl:when test="$VR-temp='text'">
								<xsl:value-of select="'paragraph'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'paragraph'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
					<xsl:if test="not (contains(@style,'mso-position-vertical'))">
            <uof:绝对_4107>

              <!--2014-03-12，wudi，增加属性值标签，start-->
              <xsl:attribute name="值_4108">

                <!--2013-04-19，wudi，增加对对象集-图片，图表对象的处理，start-->
                
                <!--2014-03-26，wudi，增加w:pict的情况，start-->
                <xsl:if test ="parent::w:object or parent::w:pict">
                  <xsl:value-of select ="'0'"/>
                </xsl:if>
                <xsl:if test ="not(parent::w:object or parent::w:pict)">
                  <xsl:value-of select="$ML-temp"/>
                </xsl:if>
                <!--end-->
                
                <!--end-->

              </xsl:attribute>
              <!--end-->
              
						</uof:绝对_4107>
					</xsl:if>
					<xsl:if test="contains(@style,'mso-position-vertical')">
            <uof:相对_4109>
							<xsl:attribute name="参考点_410B">
								<xsl:choose>
									<xsl:when test="$V-temp='center'">
										<xsl:value-of select="'center'"/>
									</xsl:when>
									<xsl:when test="$V-temp='top'">
										<xsl:value-of select="'top'"/>
									</xsl:when>
									<xsl:when test="$V-temp='bottom'">
										<xsl:value-of select="'bottom'"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="'top'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<!--<xsl:attribute name="字:值">
								<xsl:value-of select="$MT-temp"/>
							</xsl:attribute>-->
						</uof:相对_4109>
					</xsl:if>
				</uof:垂直_410D>
			</uof:位置_C620>
      
      <uof:大小_C621>
        <!--2015-01-29，wudi，修改w-temp1和h-temp1的转换BUG，start-->
				<xsl:if test="@style">
					<xsl:if test="contains(@style,'width') or contains(@style,'WIDTH') or contains(@style,'height') or contains(@style,'HEIGHT')">
						<xsl:variable name="w-temp" select="substring-after($style,'width:')"/>
            
            <!--2014-03-24，wudi，考虑单位为in的情况，start-->
            <xsl:variable name="W-temp">
              <xsl:value-of select="substring($w-temp,1,15)"/>
            </xsl:variable>
            <xsl:variable name="w-temp1">
              <!--2015-03-19，wudi，将if选择改为choose，避免某些条件存在冲突，start-->
              <xsl:choose>
                <xsl:when test="contains($W-temp,'pt')">
                  <xsl:value-of select="substring-before($w-temp,'pt')"/>
                </xsl:when>

                <!--2014-03-26，wudi，添加缺失的倍数72，start-->
                <xsl:when test="contains($W-temp,'in')">
                  <xsl:value-of select="substring-before($w-temp,'in') * 72"/>
                </xsl:when>
                <!--end-->
              </xsl:choose>
              <!--end-->
            </xsl:variable>
            <!--<xsl:variable name="w-temp1" select="substring-before($w-temp,'pt;')"/>-->
            <!--end-->
            
						<xsl:variable name="h-temp" select="substring-after($style,'height:')"/>

            <!--2014-03-21，wudi，修改h-temp1的取值方式，考虑单位in和margin同时存在的情况，start-->
            <!--2013-04-19，wudi，针对对象集-图片，图表对象的转换，修改h-temp1的取值方式，start-->
            <xsl:variable name ="H-temp">
              <xsl:value-of select ="substring($h-temp,1,15)"/>
            </xsl:variable>
            <!--<xsl:variable name ="tmp">
              <xsl:value-of select ="string-length($H-temp)"/>
            </xsl:variable>
            <xsl:variable name ="tmp1">
              <xsl:value-of select ="substring($H-temp,$tmp - 1,2)"/>
            </xsl:variable>-->
            <xsl:variable name ="h-temp1">
              
              <!--2015-01-07，wudi，注释掉此部分代码，start-->
              <!--<xsl:if test ="$tmp1 ='pt'">
                <xsl:value-of select ="substring-before($h-temp,'pt')"/>
              </xsl:if>-->
              <!--end-->

              <!--2014-03-13，wudi，存在以in为单位的情况，1in=72pt，start-->

              <!--2015-01-07，wudi，注释掉此部分代码，start-->
              <!--<xsl:if test ="$tmp1 ='in'">
                <xsl:value-of select ="number(substring-before($h-temp,'in')) * 72"/>
              </xsl:if>-->
              <!--end-->

              <!--2015-03-19，wudi，将if选择改为choose，避免某些条件存在冲突，start-->
              <xsl:choose>
                <xsl:when test ="contains($H-temp,'pt')">
                  <xsl:value-of select ="substring-before($H-temp,'pt')"/>
                </xsl:when>

                <!--2014-03-26，wudi，添加缺失的倍数72，start-->
                <xsl:when test ="contains($H-temp,'in')">
                  <xsl:value-of select ="substring-before($H-temp,'in') * 72"/>
                </xsl:when>
                <!--end-->
              </xsl:choose>
              <!--end-->

              <!--end-->

            </xsl:variable>
            <!--end-->
            <!--end-->
            
            <!--2012-11-22，wudi，修正锚点大小属性‘宽_C605’和‘长_C604’取值弄反的错误，start-->
            <xsl:attribute name="宽_C605">
							<xsl:value-of select="$w-temp1"/>
            </xsl:attribute>
            <xsl:attribute name="长_C604">
							<xsl:value-of select="$h-temp1"/>
						</xsl:attribute>
            <!--end-->
            
					</xsl:if>
				</xsl:if>
        <!--end-->
      </uof:大小_C621>
      
			<uof:绕排_C622 环绕文字_C624="both">
				<xsl:variable name="wrapType">
					<xsl:if test="w10:wrap">
						<xsl:value-of select="w10:wrap/@type"/>
					</xsl:if>
					<xsl:if test="not(w10:wrap)">
						<xsl:value-of select="'none'"/>
					</xsl:if>
				</xsl:variable>
				<xsl:variable name="index">
					<xsl:if test="contains(@style,'z-index')">
						<xsl:variable name="temp" select="substring-after($style,'z-index:')"/>
						<xsl:value-of select="substring-before($temp,';')"/>
					</xsl:if>
					<xsl:if test="not (contains(@style,'z-index'))">
						<xsl:value-of select="'0'"/>
					</xsl:if>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="$wrapType='tight'">
						<xsl:attribute name="绕排方式_C623">
							<xsl:value-of select="'tight'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test="$wrapType='topAndBottom'">
						<xsl:attribute name="绕排方式_C623">
							<xsl:value-of select="'top-bottom'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test="$wrapType='through'">
						<xsl:attribute name="绕排方式_C623">
							<xsl:value-of select="'through'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test="$wrapType='square'">
						<xsl:attribute name="绕排方式_C623">
							<xsl:value-of select="'square'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test="$index &lt; '-250000000'">
						<xsl:attribute name="绕排方式_C623">
							<xsl:value-of select="'behind-text'"/>
						</xsl:attribute>
					</xsl:when>
					<xsl:when test="$index &gt; '250000000'">
						<xsl:attribute name="绕排方式_C623">
							<xsl:value-of select="'infront-of-text'"/>
						</xsl:attribute>
					</xsl:when>
				</xsl:choose>
			</uof:绕排_C622>
      <xsl:if test="@distT or @distL or @distR or @distB">
        <uof:边距_C628>
          <xsl:attribute name="上_C609">
            <xsl:value-of select="@distT div 12500"/>
          </xsl:attribute>
          <xsl:attribute name="左_C608">
            <xsl:value-of select="@distL div 12500"/>
          </xsl:attribute>
          <xsl:attribute name="右_C60A">
            <xsl:value-of select="@distR div 12500"/>
          </xsl:attribute>
          <xsl:attribute name="下_C60B">
            <xsl:value-of select="@distB div 12500"/>
          </xsl:attribute>
        </uof:边距_C628>
      </xsl:if>
		</uof:锚点_C644>
	</xsl:template>
	<xsl:template name="bodyAnchorPic"><!--图片对象锚点模板-->
		<xsl:param name="picType"/>
		<xsl:param name="filename"/>
		<xsl:variable name="originalId">
			<xsl:value-of select=".//a:blip/@r:embed"/>
		</xsl:variable>
		<xsl:variable name="number">
      
      <!--2014-03-18，wudi，增加文本框里插入图片的情形，start-->
      <xsl:if test="$filename='document' or $filename='txbody'">
        <xsl:if test="document('word/document.xml')/w:document//w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic
               or document('word/document.xml')/w:document//w:drawing/wp:anchor/a:graphic/a:graphicData/pic:pic">
          <xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline | v:shape"/>
        </xsl:if>
      </xsl:if>
      <!--end-->
      
			<!--<xsl:if test="not(document('word/document.xml')/w:document//w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic)">-->
      <xsl:if test="$filename!='document' and $filename!='txbody'">
				<xsl:call-template name="numberForEveryPart">
					<xsl:with-param name="numPart" select="$filename"/>
					<xsl:with-param name="ooxId" select="$originalId"/>
				</xsl:call-template>
			<!--</xsl:if>--></xsl:if>
		</xsl:variable>   
		<uof:锚点_C644>
      <xsl:attribute name="图形引用_C62E">
        <xsl:choose>
          <xsl:when test="$filename='document'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>

          <!--2014-03-18，wudi，增加文本框里插入图片的情形，start-->
          <xsl:when test="$filename='txbody'">
            <xsl:value-of select="concat('document','Obj',$number * 2 +1)"/>
          </xsl:when>
          <!--end-->
          
          <xsl:when test="$filename='comments'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1 )"/>
          </xsl:when>
          <xsl:when test="$filename='endnotes'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="$filename='footnotes'">
            <xsl:value-of select="concat($filename,'Obj',$number * 2 +1)"/>
          </xsl:when>
          <xsl:when test="contains($filename,'header')">
            <xsl:variable name="hn" select="substring-before($filename,'.xml')"/>
            <xsl:variable name="headerObjnumber"><!--CXL,2012.4.24，页眉页脚的图片引用要单独计数-->
              <xsl:value-of select="count(document(concat('word/',$filename))//v:rect) + count(document(concat('word/',$filename))//v:shape) + count(document(concat('word/',$filename))//wp:anchor) + count(document(concat('word/',$filename))//wp:inline)"/>
            </xsl:variable>
            <xsl:value-of select="concat($hn,'Obj',$headerObjnumber * 2 +1)"/>
          </xsl:when>
          <xsl:when test="contains($filename,'footer')">
            <xsl:variable name="fn" select="substring-before($filename,'.xml')"/>
            <xsl:variable name="footerObjnumber">
              <xsl:value-of select="count(document(concat('word/',$filename))//v:rect) + count(document(concat('word/',$filename))//v:shape) + count(document(concat('word/',$filename))//wp:anchor) + count(document(concat('word/',$filename))//wp:inline)"/>
            </xsl:variable>
            <xsl:value-of select="concat($fn,'Obj',$footerObjnumber * 2 +1)"/>
          </xsl:when>
        </xsl:choose>
      </xsl:attribute>
			<!--<xsl:attribute name="字:类型">
				<xsl:if test="$picType='inline'">
					<xsl:value-of select="'inline'"/>
				</xsl:if>
				<xsl:if test="$picType='anchor'">
					<xsl:value-of select="'normal'"/>
				</xsl:if>
			</xsl:attribute>-->
      
      <!--2014-03-28，wudi，uof:位置_C620的转换，修改路径查找方法，start-->
      <xsl:if test=".//wp:positionH or .//wp:positionV">
        <uof:位置_C620>
          <xsl:if test=".//wp:positionH">
            <uof:水平_4106>
              <xsl:attribute name="相对于_410C">
                <xsl:call-template name="positionHrelative">
                  <xsl:with-param name="val" select=".//wp:positionH/@relativeFrom"/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:if test="wp:positionH/wp:posOffset">
                <uof:绝对_4107>
                  <xsl:attribute name="值_4108">
                    <xsl:value-of select=".//wp:positionH/wp:posOffset div 12700"/>
                  </xsl:attribute>
                </uof:绝对_4107>
              </xsl:if>
              <xsl:if test="wp:positionH/wp:align">
                <uof:相对_4109>
                  <xsl:attribute name="参考点_410A">
                    <xsl:call-template name="positionHalign">
                      <xsl:with-param name="val" select=".//wp:positionH/wp:align"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </uof:相对_4109>
              </xsl:if>
            </uof:水平_4106>
          </xsl:if>
          <xsl:if test=".//wp:positionV">
            <uof:垂直_410D>
              <xsl:attribute name="相对于_C647">
                <xsl:call-template name="positionVrelative">
                  <xsl:with-param name="val" select=".//wp:positionV/@relativeFrom"/>
                </xsl:call-template>
              </xsl:attribute>
              <xsl:if test=".//wp:positionV/wp:posOffset">
                <uof:绝对_4107>
                  <xsl:attribute name="值_4108">
                    <xsl:value-of select=".//wp:positionV/wp:posOffset div 12700"/>
                  </xsl:attribute>
                </uof:绝对_4107>
              </xsl:if>
              <xsl:if test=".//wp:positionV/wp:align">
                <uof:相对_4109>
                  <xsl:attribute name="参考点_410B">
                    <xsl:call-template name="positionValign">
                      <xsl:with-param name="val" select=".//wp:positionV/wp:align"/>
                    </xsl:call-template>
                  </xsl:attribute>
                </uof:相对_4109>
              </xsl:if>
            </uof:垂直_410D>
          </xsl:if>
        </uof:位置_C620>
      </xsl:if>
      <!--end-->
      
      <uof:大小_C621>
        <xsl:attribute name="宽_C605">
          <xsl:value-of select="wp:extent/@cx div 12700"/>
        </xsl:attribute>
        <xsl:attribute name="长_C604">
          <xsl:value-of select="wp:extent/@cy div 12700"/>
        </xsl:attribute>
      </uof:大小_C621>		
     	
      <uof:绕排_C622>
				<xsl:if test="@behindDoc='1'">
					<xsl:attribute name="绕排方式_C623">
						<xsl:value-of select="'behind-text'"/>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test="@behindDoc='0'">
					<xsl:attribute name="绕排方式_C623">
						<xsl:value-of select="'infront-of-text'"/>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test="wp:wrapTight">
					<xsl:attribute name="绕排方式_C623">
						<xsl:value-of select="'tight'"/>
					</xsl:attribute>
					<xsl:if test="wp:wrapTight/@wrapText">
						<xsl:attribute name="环绕文字_C624">
							<xsl:call-template name="wraptext">
								<xsl:with-param name="val" select="wp:wrapTight/@wrapText"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
				</xsl:if>
				<xsl:if test="wp:wrapTopAndBottom">
					<xsl:attribute name="绕排方式_C623">
						<xsl:value-of select="'top-bottom'"/>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test="wp:wrapThrough">
					<xsl:attribute name="绕排方式_C623">
						<xsl:value-of select="'through'"/>
					</xsl:attribute>
					<xsl:if test="wp:wrapThrough/@wrapText">
						<xsl:attribute name="环绕文字_C624">
							<xsl:call-template name="wraptext">
								<xsl:with-param name="val" select="wp:wrapThrough/@wrapText"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
				</xsl:if>
				<xsl:if test="wp:wrapSquare">
					<xsl:attribute name="绕排方式_C623">
						<xsl:value-of select="'square'"/>
					</xsl:attribute>
					<xsl:if test="wp:wrapSquare/@wrapText">
						<xsl:attribute name="环绕文字_C624">
							<xsl:call-template name="wraptext">
								<xsl:with-param name="val" select="wp:wrapSquare/@wrapText"/>
							</xsl:call-template>
						</xsl:attribute>
					</xsl:if>
				</xsl:if>
			</uof:绕排_C622>
      
			<xsl:if test="@distT or @distL or @distR or @distB">
        <uof:边距_C628>
					<xsl:attribute name="上_C609">
						<xsl:value-of select="@distT div 12500"/>
					</xsl:attribute>
					<xsl:attribute name="左_C608">
						<xsl:value-of select="@distL div 12500"/>
					</xsl:attribute>
					<xsl:attribute name="右_C60A">
						<xsl:value-of select="@distR div 12500"/>
					</xsl:attribute>
					<xsl:attribute name="下_C60B">
						<xsl:value-of select="@distB div 12500"/>
					</xsl:attribute>
				</uof:边距_C628>
			</xsl:if>
      
			<xsl:if test="wp:cNvGraphicFramePr/a:graphicFrameLocks/@noResize='1'">
        <uof:是否锁定_C629 值="false"/>
			</xsl:if>
      <uof:保护_C62A 是否保护内容_C641="false" 是否保护位置_C642="false" 是否保护大小_C643="false"/>
			<xsl:if test="@allowOverlap='1'">
        <uof:是否允许重叠_C62B>true</uof:是否允许重叠_C62B>
			</xsl:if>
		</uof:锚点_C644>
	</xsl:template>
  
  
	<xsl:template name="numberForEveryPart">
		<xsl:param name="numPart"/>
		<xsl:param name="ooxId"/>
		<xsl:choose>
      
      <!--2014-03-18，注释掉此部分代码，多余，start-->
      <!--<xsl:when test="$numPart='document'">
        <xsl:variable name="nb">
          <xsl:apply-templates select="document('word/document.xml')/w:document//w:drawing[(wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (v:rect//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (v:shape//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId)]" mode="objNum"/>
        </xsl:variable>
        <xsl:value-of select="$nb"/>
      </xsl:when>-->
      <!--end-->
      
			<xsl:when test="$numPart='comments'">
				<xsl:variable name="number">
					<xsl:apply-templates select="document('word/comments.xml')/w:comments//w:drawing[(wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (v:rect//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (v:shape//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId)]" mode="objNum"/>
				</xsl:variable>
				<xsl:value-of select="$number"/>
			</xsl:when>
			<xsl:when test="$numPart='endnotes'">
				<xsl:variable name="number">
					<xsl:apply-templates select="document('word/endnotes.xml')/w:endnotes//w:drawing[(wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (v:rect//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (v:shape//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId)]" mode="objNum"/>
				</xsl:variable>
        <xsl:value-of select="$number"/>
			</xsl:when>
			<xsl:when test="$numPart='footnotes'">
				<xsl:variable name="number">
					<xsl:apply-templates select="document('word/footnotes.xml')/w:footnotes//w:drawing[(wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (v:rect//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (v:shape//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId)]" mode="objNum"/>
				</xsl:variable>
				<xsl:value-of select="$number"/>
			</xsl:when>
			<xsl:when test="contains($numPart,'header')">
				<xsl:variable name="hnn" select="substring-before($numPart,'.xml')"/>
				<xsl:variable name="hpath" select="concat('word/',$numPart)"/>
				<xsl:variable name="number">
					<xsl:apply-templates select="document($hpath)/w:hdr//w:drawing[(wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or(v:rect//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or(v:shape//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId)]" mode="objNum"/>
				</xsl:variable>
				<xsl:value-of select="$number"/>
			</xsl:when>
			<xsl:when test="contains($numPart,'footer')">
				<xsl:variable name="fnn" select="substring-before($numPart,'.xml')"/>
				<xsl:variable name="fpath" select="concat('word/',$numPart)"/>
				<xsl:variable name="number">
					<!--<xsl:apply-templates select="document($fpath)/w:ftr//[(wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or  (wp:anchor/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId) or (v:rect//a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip/@r:embed=$ooxId)]" mode="objNum"/>-->
				</xsl:variable>
				<xsl:value-of select="$number"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<!--
  <xsl:template name="numberForEveryPart">
    <xsl:param name ="numPart"/>
    <xsl:param name ="ooxId"/>
    <xsl:value-of select ="'100'"/>
  </xsl:template>
  -->
	<xsl:template match="w:drawing" mode="objNum">
		<xsl:for-each select="v:rect | wp:anchor | wp:inline | v:shape">
			<xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline | v:shape"/>
		</xsl:for-each>
	</xsl:template>
	
	<xsl:template name="wraptext">
		<xsl:param name="val"/>
		<xsl:choose>
			<xsl:when test="$val='bothSides'">
				<xsl:value-of select="'both'"/>
			</xsl:when>
			<xsl:when test="$val='left'">
				<xsl:value-of select="'left'"/>
			</xsl:when>
			<xsl:when test="$val='right'">
				<xsl:value-of select="'right'"/>
			</xsl:when>
			<xsl:when test="$val='largest'">
				<xsl:value-of select="'largest'"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="'both'"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="positionHrelative">
		<xsl:param name="val"/>
		<xsl:choose>
			<xsl:when test="$val='margin'">
				<xsl:value-of select="'margin'"/>
			</xsl:when>
			<xsl:when test="$val='page'">
				<xsl:value-of select="'page'"/>
			</xsl:when>
			<xsl:when test="$val='column'">
				<xsl:value-of select="'column'"/>
			</xsl:when>
			<xsl:when test="$val='character'">
				<xsl:value-of select="'char'"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="'margin'"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="positionHalign">
		<xsl:param name="val"/>
		<xsl:choose>
			<xsl:when test="$val='left'">
				<xsl:value-of select="'left'"/>
			</xsl:when>
			<xsl:when test="$val='right'">
				<xsl:value-of select="'right'"/>
			</xsl:when>
			<xsl:when test="$val='center'">
				<xsl:value-of select="'center'"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="'left'"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="positionVrelative">
		<xsl:param name="val"/>
		<xsl:choose>
			<xsl:when test="$val='margin'">
				<xsl:value-of select="'margin'"/>
			</xsl:when>
			<xsl:when test="$val='page'">
				<xsl:value-of select="'page'"/>
			</xsl:when>
			<xsl:when test="$val='paragraph'">
				<xsl:value-of select="'paragraph'"/>
			</xsl:when>
			<xsl:when test="$val='line'">
				<xsl:value-of select="'line'"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="'margin'"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="positionValign">
		<xsl:param name="val"/>
		<xsl:choose>
			<xsl:when test="$val='top'">
				<xsl:value-of select="'top'"/>
			</xsl:when>
			<xsl:when test="$val='bottom'">
				<xsl:value-of select="'bottom'"/>
			</xsl:when>
			<xsl:when test="$val='center'">
				<xsl:value-of select="'center'"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="'top'"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	<xsl:template name="lineType">
		<xsl:param name="linetype"/>
		<xsl:choose>
			<xsl:when test="$linetype='none'">
				<xsl:value-of select="'none'"/>
			</xsl:when>
			<xsl:when test="$linetype='single'">
				<xsl:value-of select="'single'"/>
			</xsl:when>
			<xsl:when test="$linetype='thick'">
				<xsl:value-of select="'thick'"/>
			</xsl:when>
			<xsl:when test="$linetype='double'">
				<xsl:value-of select="'double'"/>
			</xsl:when>
			<xsl:when test="$linetype='dotted'">
				<xsl:value-of select="'dotted'"/>
			</xsl:when>
			<xsl:when test="$linetype='dashed'">
				<xsl:value-of select="'dashed'"/>
			</xsl:when>
			<xsl:when test="$linetype='dotDash'">
				<xsl:value-of select="'dot-dash'"/>
			</xsl:when>
			<xsl:when test="$linetype='dotDotDash'">
				<xsl:value-of select="'dot-dot-dash'"/>
			</xsl:when>
			<xsl:when test="$linetype='wave'">
				<xsl:value-of select="'wave'"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="'single'"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--CXL2011/11/30修改zhaobj 箭头-->
  
  <!--2013-11-13，wudi，增加wp:spPr/a:ln，start-->
	<xsl:template match="wps:spPr/a:ln | wp:spPr/a:ln" mode="Arrow">
    <图:箭头_805D>
      <xsl:for-each select="a:headEnd | a:tailEnd">
        <xsl:if test="name(.)='a:headEnd'">
          <xsl:choose>
            <xsl:when test="@type='triangle'">
              <图:前端箭头_805E>
                <图:式样_8000>normal</图:式样_8000>
                <图:大小_8001>
                  <xsl:call-template name="Arrowsize"/>
                </图:大小_8001>
              </图:前端箭头_805E>
            </xsl:when>
            <xsl:when test="@type='arrow'">
              <图:前端箭头_805E>
                <图:式样_8000>open</图:式样_8000>
                <图:大小_8001>
                  <xsl:call-template name="Arrowsize"/>
                </图:大小_8001>
              </图:前端箭头_805E>
            </xsl:when>
            <xsl:when test="@type='stealth'">
              <图:前端箭头_805E>
                <图:式样_8000>stealth</图:式样_8000>
                <图:大小_8001>
                  <xsl:call-template name="Arrowsize"/>
                </图:大小_8001>
              </图:前端箭头_805E>
            </xsl:when>
            <xsl:when test="@type='diamond'">
              <图:前端箭头_805E>
                <图:式样_8000>diamond</图:式样_8000>
                <图:大小_8001>
                  <xsl:call-template name="Arrowsize"/>
                </图:大小_8001>
              </图:前端箭头_805E>
            </xsl:when>
            <xsl:when test="@type='oval'">
              <图:前端箭头_805E>
                <图:式样_8000>oval</图:式样_8000>
                <图:大小_8001>
                  <xsl:call-template name="Arrowsize"/>
                </图:大小_8001>
              </图:前端箭头_805E>
            </xsl:when>
            <xsl:otherwise> </xsl:otherwise>
            <!--<xsl:when test ="@type='none' or not(@type)"></xsl:when>-->
          </xsl:choose>
        </xsl:if>
        <xsl:if test="name(.)='a:tailEnd'">
          <xsl:choose>
            <xsl:when test="@type='triangle'">
              <图:后端箭头_805F>
                <图:式样_8000>normal</图:式样_8000>
                <图:大小_8001>
                  <xsl:call-template name="Arrowsize"/>
                </图:大小_8001>
              </图:后端箭头_805F>
            </xsl:when>
            <xsl:when test="@type='arrow'">
              <图:后端箭头_805F>
                <图:式样_8000>open</图:式样_8000>
                <图:大小_8001>
                  <xsl:call-template name="Arrowsize"/>
                </图:大小_8001>
              </图:后端箭头_805F>
            </xsl:when>
            <xsl:when test="@type='stealth'">
              <图:后端箭头_805F>
                <图:式样_8000>stealth</图:式样_8000>
                <图:大小_8001>
                  <xsl:call-template name="Arrowsize"/>
                </图:大小_8001>
              </图:后端箭头_805F>
            </xsl:when>
            <xsl:when test="@type='diamond'">
              <图:后端箭头_805F>
                <图:式样_8000>diamond</图:式样_8000>
                <图:大小_8001>
                  <xsl:call-template name="Arrowsize"/>
                </图:大小_8001>
              </图:后端箭头_805F>
            </xsl:when>
            <xsl:when test="@type='oval'">
              <图:后端箭头_805F>
                <图:式样_8000>oval</图:式样_8000>
                <图:大小_8001>
                  <xsl:call-template name="Arrowsize"/>
                </图:大小_8001>
              </图:后端箭头_805F>
            </xsl:when>
            <xsl:otherwise> </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:for-each>
    </图:箭头_805D>
	</xsl:template>
  <!--end-->
  
	<xsl:template name="Arrowsize">
		<xsl:choose>
			<xsl:when test="@w='sm' and @len='sm'">1</xsl:when>
			<xsl:when test="@w='sm' and @len='med'">2</xsl:when>
			<xsl:when test="@w='sm' and @len='lg'">3</xsl:when>
			<xsl:when test="@w='med' and @len='sm'">4</xsl:when>
			<xsl:when test="@w='med' and @len='med'">5</xsl:when>
			<xsl:when test="@w='med' and @len='lg'">6</xsl:when>
			<xsl:when test="@w='lg' and @len='sm'">7</xsl:when>
			<xsl:when test="@w='lg' and @len='med'">8</xsl:when>
			<xsl:when test="@w='lg' and @len='lg'">9</xsl:when>
			<xsl:otherwise>5</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--yuch 阴影 -->
	
	<xsl:template match="a:outerShdw" mode="shadow">
		<图:阴影_8051 是否显示阴影_C61C="true" 类型_C61D="single">
			
			<xsl:choose>
				<xsl:when test="a:srgbClr">
					<xsl:attribute name="颜色_C61E">
						<xsl:value-of select="concat('#',a:srgbClr/@val)"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="a:prstClr">
					<xsl:attribute name="颜色_C61E">
						<xsl:apply-templates select="a:prstClr"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="a:schemeClr">
					<xsl:attribute name="颜色_C61E">
						<xsl:apply-templates select="a:schemeClr"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
					<xsl:attribute name="颜色_C61E">
						<xsl:value-of select="'auto'"/>
					</xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>

			<xsl:choose>
        <!--2013-03-06，wudi，修改if判断条件，修复预定义图形阴影-透明度转换错误BUG，start-->
        
        <!--2013-11-22，wudi，Strcit标准下，val取值不同，修正，start-->
        
        <!--2014-01-08，wudi，透明度的取值有误，修正-->
				<xsl:when test="a:srgbClr/a:alpha">
          <xsl:variable name="tmp">
            <xsl:value-of select="number(substring-before(a:srgbClr/a:alpha/@val,'%'))"/>
          </xsl:variable>
					<xsl:if test="$tmp &lt; 20">
						<xsl:attribute name="透明度_C61F">0</xsl:attribute>
					</xsl:if>
					<xsl:if test="$tmp &gt; 19">
						<xsl:attribute name="透明度_C61F">49</xsl:attribute>
					</xsl:if>
				</xsl:when>
				<xsl:when test="a:prstClr/a:alpha">
          <xsl:variable name="tmp">
            <xsl:value-of select="number(substring-before(a:prstClr/a:alpha/@val,'%'))"/>
          </xsl:variable>
					<xsl:if test="$tmp &lt; 20">
						<xsl:attribute name="透明度_C61F">0</xsl:attribute>
					</xsl:if>
					<xsl:if test="$tmp &gt; 19">
						<xsl:attribute name="透明度_C61F">49</xsl:attribute>
					</xsl:if>
				</xsl:when>
				<xsl:when test="a:schemeClr/a:alpha">
          <xsl:variable name="tmp">
            <xsl:value-of select="number(substring-before(a:schemeClr/a:alpha/@val,'%'))"/>
          </xsl:variable>
					<xsl:if test="$tmp &lt; 20">
						<xsl:attribute name="透明度_C61F">0</xsl:attribute>
					</xsl:if>
					<xsl:if test="$tmp &gt; 19">
						<xsl:attribute name="透明度_C61F">49</xsl:attribute>
					</xsl:if>
				</xsl:when>
        <!--end-->
        
        <!--end-->
        
        <!--end-->
			</xsl:choose>
			
			
			<uof:偏移量_C61B>
				<xsl:if test="x">
					<xsl:attribute name="x_C606">
						<xsl:value-of select="round(x)"/>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test="y">
					<xsl:attribute name="y_C607">
						<xsl:value-of select="round(y)"/>
					</xsl:attribute>
				</xsl:if>
			</uof:偏移量_C61B>
		
		</图:阴影_8051>
	</xsl:template>
	<xsl:template match="a:innerShdw" mode="shadow">
		<图:阴影_8051 是否显示阴影_C61C="true" 类型_C61D="single">
			<xsl:choose>
				<xsl:when test="a:srgbClr">
					<xsl:attribute name="颜色_C61E">
						<xsl:value-of select="concat('#',a:srgbClr/@val)"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="a:prstClr">
					<xsl:attribute name="颜色_C61E">
						<xsl:apply-templates select="a:prstClr"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:when test="a:schemeClr">
					<xsl:attribute name="颜色_C61E">
						<xsl:apply-templates select="a:schemeClr"/>
					</xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
					<xsl:attribute name="颜色_C61E">
						<xsl:value-of select="'auto'"/>
					</xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>

			<xsl:choose>
        <!--2013-03-06，wudi，修改if判断条件，修复预定义图形阴影-透明度转换错误BUG，start-->
        
        <!--2013-11-22，wudi，Strcit标准下，val取值不同，修正，start-->

        <!--2014-01-08，wudi，透明度的取值有误，修正-->
        <xsl:when test="a:srgbClr/a:alpha">
          <xsl:variable name="tmp">
            <xsl:value-of select="number(substring-before(a:srgbClr/a:alpha/@val,'%'))"/>
          </xsl:variable>
          <xsl:if test="$tmp &lt; 20">
            <xsl:attribute name="透明度_C61F">0</xsl:attribute>
          </xsl:if>
          <xsl:if test="$tmp &gt; 19">
            <xsl:attribute name="透明度_C61F">49</xsl:attribute>
          </xsl:if>
        </xsl:when>
        <xsl:when test="a:prstClr/a:alpha">
          <xsl:variable name="tmp">
            <xsl:value-of select="number(substring-before(a:prstClr/a:alpha/@val,'%'))"/>
          </xsl:variable>
          <xsl:if test="$tmp &lt; 20">
            <xsl:attribute name="透明度_C61F">0</xsl:attribute>
          </xsl:if>
          <xsl:if test="$tmp &gt; 19">
            <xsl:attribute name="透明度_C61F">49</xsl:attribute>
          </xsl:if>
        </xsl:when>
        <xsl:when test="a:schemeClr/a:alpha">
          <xsl:variable name="tmp">
            <xsl:value-of select="number(substring-before(a:schemeClr/a:alpha/@val,'%'))"/>
          </xsl:variable>
          <xsl:if test="$tmp &lt; 20">
            <xsl:attribute name="透明度_C61F">0</xsl:attribute>
          </xsl:if>
          <xsl:if test="$tmp &gt; 19">
            <xsl:attribute name="透明度_C61F">49</xsl:attribute>
          </xsl:if>
        </xsl:when>
        <!--end-->
        
        <!--end-->

        <!--end-->
			</xsl:choose>

			<uof:偏移量_C61B>
        <xsl:if test="(@dir div 60000) = 0  or not(@dir)">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="6.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="0.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 45">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="6.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="6.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 90">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="0.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="6.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 135">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="-6.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="6.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 180">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="-6.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="0.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 225">
					<xsl:attribute name="x_C606">
						<xsl:value-of select="-6.0"/>
					</xsl:attribute>
					<xsl:attribute name="y_C607">
						<xsl:value-of select="-6.0"/>
					</xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 270">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="0.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="-6.0"/>
          </xsl:attribute>
        </xsl:if>
        <xsl:if test="(@dir div 60000) = 315">
          <xsl:attribute name="x_C606">
            <xsl:value-of select="6.0"/>
          </xsl:attribute>
          <xsl:attribute name="y_C607">
            <xsl:value-of select="-6.0"/>
          </xsl:attribute>
        </xsl:if>       
			</uof:偏移量_C61B>
			
		</图:阴影_8051>
	</xsl:template>
	<!--图填充cxl2011/12/8修改-->
	<xsl:template match="a:solidFill">
		<xsl:call-template name="colorChoice"/>
	</xsl:template>
	<xsl:template match="a:gradFill">
		<图:渐变_800D>
			<xsl:variable name="angle">
				<xsl:choose>
					<xsl:when test="a:lin">
						<xsl:value-of select="round(a:lin/@ang div 60000)"/><!--(360 - round(a:lin/@ang div 60000 div 45)*45+90) mod 360-->
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
			</xsl:variable>
			<xsl:attribute name="种子类型_8010">
				<xsl:choose>
					<xsl:when test="a:path">
						<xsl:choose>
							<xsl:when test="a:path/@path='rect'">square</xsl:when>
              
              <!--2012-01-25，wudi，修改oval为rectangle，start-->
							<xsl:when test="a:path/@path='circle'">rectangle</xsl:when>
              <!--end-->
              
							<xsl:when test="a:path/@path='shape'">square</xsl:when>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>linear</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
			<xsl:attribute name="起始浓度_8011">1.0</xsl:attribute>
			<xsl:attribute name="终止浓度_8012">1.0</xsl:attribute>
			<xsl:attribute name="渐变方向_8013">
        
        <!--2012-01-25，wudi，OOX到UOF方向修复BUG：渐变填充存在差异，start-->
        <xsl:if test ="$angle='0' or $angle='180'">
          <xsl:value-of select ="90"/>
        </xsl:if>
        <xsl:if test ="$angle='90' or $angle='270'">
          <xsl:value-of select ="0"/>
        </xsl:if>
        <xsl:if test ="$angle='135'">
          <xsl:value-of select ="315"/>
        </xsl:if>
        <xsl:if test ="$angle='225'">
          <xsl:value-of select ="45"/>
        </xsl:if>
        <xsl:if test ="not($angle='0') and not($angle='90') and not($angle='180') and not($angle='270') and not($angle='135') and not($angle='225')">
          <xsl:value-of select="$angle"/>
        </xsl:if>
        <!--end-->

      </xsl:attribute>
			<xsl:choose>
        
        <!--2012-01-25，wudi，修改test判定条件：去135，增315，start-->
				<xsl:when test="$angle='180' or $angle='225' or $angle='270' or $angle='315'">
          <!--end-->
          
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
        
        <!--2012-01-25，wudi，OOX到UOF方向修复BUG：渐变填充存在差异，start-->
        <xsl:when test ="a:path/@path='circle'">
          <xsl:attribute name ="种子X位置_8015">
            <xsl:if test ="a:path/a:fillToRect/@l">
              <xsl:value-of select ="a:path/a:fillToRect/@l div 1000"/>
            </xsl:if>
            <xsl:if test ="not(a:path/a:fillToRect/@l)">
              <xsl:value-of select ="0"/>
            </xsl:if>
          </xsl:attribute>
          <xsl:attribute name ="种子Y位置_8016">
            <xsl:if test ="a:path/a:fillToRect/@t">
              <xsl:value-of select ="a:path/a:fillToRect/@t div 1000"/>
            </xsl:if>
            <xsl:if test ="not(a:path/a:fillToRect/@t)">
              <xsl:value-of select ="0"/>
            </xsl:if>
          </xsl:attribute>
        </xsl:when>
        <!--end-->
        
			</xsl:choose>
		</图:渐变_800D>
	</xsl:template>
	<xsl:template match="a:fillRef">
		<xsl:call-template name="colorChoice"/>
	</xsl:template>
  <!--cxl2011/12/8修改-->
	<xsl:template match="a:pattFill">
		<图:图案_800A>
			<xsl:attribute name="类型_8008">
				<xsl:call-template name="patternFill">
					<xsl:with-param name="pa" select="1"/>
				</xsl:call-template>
			</xsl:attribute>
			<xsl:attribute name="前景色_800B">
				<xsl:call-template name="patternFill">
					<xsl:with-param name="pa" select="2"/>
				</xsl:call-template>
			</xsl:attribute>
			<xsl:attribute name="背景色_800C">
				<xsl:call-template name="patternFill">
					<xsl:with-param name="pa" select="3"/>
				</xsl:call-template>
			</xsl:attribute>
		</图:图案_800A>
	</xsl:template>
	<xsl:template name="patternFill">
		<xsl:param name="pa"/>
		<xsl:choose>
			<xsl:when test="$pa = '1'">
				<xsl:choose>
					<xsl:when test="@prst ='pct5'">ptn001</xsl:when>
					<xsl:when test="@prst ='pct10'">ptn002</xsl:when>
					<xsl:when test="@prst ='pct20'">ptn003</xsl:when>
					<xsl:when test="@prst ='pct25'">ptn004</xsl:when>
					<xsl:when test="@prst ='pct30'">ptn005</xsl:when>
					<xsl:when test="@prst ='pct40'">ptn006</xsl:when>
					<xsl:when test="@prst ='pct50'">ptn007</xsl:when>
					<xsl:when test="@prst ='pct60'">ptn008</xsl:when>
					<xsl:when test="@prst ='pct70'">ptn009</xsl:when>
					<xsl:when test="@prst ='pct75'">ptn010</xsl:when>
					<xsl:when test="@prst ='pct80'">ptn011</xsl:when>
					<xsl:when test="@prst ='pct90'">ptn012</xsl:when>
					<xsl:when test="@prst ='ltDnDiag'">ptn013</xsl:when>
					<xsl:when test="@prst ='ltUpDiag'">ptn014</xsl:when>
					<xsl:when test="@prst ='dkDnDiag'">ptn015</xsl:when>
					<xsl:when test="@prst ='dkUpDiag'">ptn016</xsl:when>
					<xsl:when test="@prst ='wdDnDiag'">ptn017</xsl:when>
					<xsl:when test="@prst ='wdUpDiag'">ptn018</xsl:when>
					<xsl:when test="@prst ='ltVert'">ptn019</xsl:when>
					<xsl:when test="@prst ='ltHorz'">ptn020</xsl:when>
					<xsl:when test="@prst ='narVert'">ptn021</xsl:when>
					<xsl:when test="@prst ='narHorz'">ptn022</xsl:when>
					<xsl:when test="@prst ='dkVert'">ptn023</xsl:when>
					<xsl:when test="@prst ='dkHorz'">ptn024</xsl:when>
					<xsl:when test="@prst ='dashDnDiag'">ptn025</xsl:when>
					<xsl:when test="@prst ='dashUpDiag'">ptn026</xsl:when>
					<xsl:when test="@prst ='dashHorz'">ptn027</xsl:when>
					<xsl:when test="@prst ='dashVert'">ptn028</xsl:when>
					<xsl:when test="@prst ='smConfetti'">ptn029</xsl:when>
					<xsl:when test="@prst ='lgConfetti'">ptn030</xsl:when>
					<xsl:when test="@prst ='zigZag'">ptn031</xsl:when>
					<xsl:when test="@prst ='wave'">ptn032</xsl:when>
					<xsl:when test="@prst ='diagBrick'">ptn033</xsl:when>
					<xsl:when test="@prst ='horzBrick'">ptn034</xsl:when>
					<xsl:when test="@prst ='weave'">ptn035</xsl:when>
					<xsl:when test="@prst ='plaid'">ptn036</xsl:when>
					<xsl:when test="@prst ='divot'">ptn037</xsl:when>
					<xsl:when test="@prst ='dotGrid'">ptn038</xsl:when>
					<xsl:when test="@prst ='dotDmnd'">ptn039</xsl:when>
					<xsl:when test="@prst ='shingle'">ptn040</xsl:when>
					<xsl:when test="@prst ='trellis'">ptn041</xsl:when>
					<xsl:when test="@prst ='sphere'">ptn042</xsl:when>
					<xsl:when test="@prst ='smGrid'">ptn043</xsl:when>
					<xsl:when test="@prst ='lgGrid'">ptn044</xsl:when>
					<xsl:when test="@prst ='smCheck'">ptn045</xsl:when>
					<xsl:when test="@prst ='lgCheck'">ptn046</xsl:when>
					<xsl:when test="@prst ='openDmnd'">ptn047</xsl:when>
					<xsl:when test="@prst ='solidDmnd'">ptn048</xsl:when>
				</xsl:choose>
			</xsl:when>
			<xsl:when test="$pa = '2'">
				<xsl:apply-templates select="a:fgClr"/>
			</xsl:when>
			<xsl:when test="$pa = '3'">
				<xsl:apply-templates select="a:bgClr"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<!--fgClr前景色-->
	<xsl:template match="a:fgClr">
		<xsl:call-template name="colorChoice"/>
	</xsl:template>
	<!--bgClr背景色-->
	<xsl:template match="a:bgClr">
		<xsl:call-template name="colorChoice"/>
	</xsl:template>
	<xsl:template match="a:blipFill">
		<xsl:param name="picturefrom"/>
    <xsl:param name ="Number"/>
    
		<xsl:variable name="number">
			<xsl:number format="1" level="any" count="v:rect | wp:anchor | wp:inline | v:shape"/>
		</xsl:variable>
		<图:图片_8005>
			<xsl:attribute name="位置_8006">
				<xsl:choose>
					<xsl:when test="a:stretch">
						<xsl:value-of select="'stretch'"/>
					</xsl:when>
					<xsl:when test="a:tile">
						<xsl:value-of select="'tile'"/>
					</xsl:when>
					<!--2010.04.12 myx add -->
					<xsl:when test="a:srcRect">tile</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="'center'"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
			<xsl:attribute name="图形引用_8007">
				<!-- <xsl:call-template name="findPicName"/>-->
        
        <!--2013-04-16，wudi，修复预定义图形转换BUG，五角星 17 和 矩形 10 带图片填充丢失，start-->
        
        <!--2014-03-06，wudi，增加心形 12和椭圆 9，start-->
				<xsl:if test="$picturefrom='grpsp' and not(ancestor::wp:wsp/wp:cNvPr/@name='五角星 17' or ancestor::wp:wsp/wp:cNvPr/@name='矩形 10' or ancestor::wp:wsp/wp:cNvPr/@name='心形 12' or ancestor::wp:wsp/wp:cNvPr/@name='椭圆 9')">
          <xsl:if test="../../wp:cNvPr/@id">
						<xsl:value-of select="concat('grpsppicObj',../../wp:cNvPr/@id)"/>
					</xsl:if>
				</xsl:if>
        <xsl:if test ="$picturefrom='grpsp' and (ancestor::wp:wsp/wp:cNvPr/@name='五角星 17' or ancestor::wp:wsp/wp:cNvPr/@name='矩形 10' or ancestor::wp:wsp/wp:cNvPr/@name='心形 12' or ancestor::wp:wsp/wp:cNvPr/@name='椭圆 9')">
          <xsl:value-of select="concat('documentObj',$number * 2)"/>
        </xsl:if>


        <!--2014-04-10，wudi，SmartArt的图片填充需特殊处理，start-->
        <xsl:if test ="$picturefrom='SmartArt'">
          <xsl:variable name="tmp">
            <xsl:value-of select="number(substring-after(./a:blip/@r:embed,'rId'))"/>
          </xsl:variable>
          <xsl:value-of select="concat($picturefrom,'Obj',($Number - 1) * 9 + $tmp)"/>
        </xsl:if>
        <!--end-->
        
				<xsl:if test="$picturefrom!='grpsp' and $picturefrom!='SmartArt'">
					<xsl:value-of select="concat($picturefrom,'Obj',$number * 2)"/>
				</xsl:if>
        <!--end-->
        
        <!--end-->
        
			</xsl:attribute>
		</图:图片_8005>
	</xsl:template>
	<xsl:template match="a:grpFill">
		<xsl:choose>
			<xsl:when test="../../../wpg:grpSpPr/a:solidFill">
				<xsl:apply-templates select="../../../wpg:grpSpPr/a:solidFill"/>
			</xsl:when>
			<xsl:when test="../../../wpg:grpSpPr/a:patternFill">
				<xsl:apply-templates select="../../../wpg:grpSpPr/a:patternFill"/>
			</xsl:when>
			<xsl:when test="../../../wpg:grpSpPr/a:pattFill">
				<xsl:apply-templates select="../../../wpg:grpSpPr/a:pattFill"/>
			</xsl:when>
			<xsl:when test="../../../wpg:grpSpPr/a:blipFill">
				<xsl:apply-templates select="../../../wpg:grpSpPr/a:blipFill">
					<xsl:with-param name="picturefrom" select="grpsp"/>
				</xsl:apply-templates>
			</xsl:when>
			<xsl:when test="../../../wpg:grpSpPr/a:grpFill">
				<xsl:apply-templates select="../../../wpg:grpSpPr/a:grpFill"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="findPicName">
		<xsl:variable name="findPictureName">
			<xsl:value-of select=".//a:blip/@r:embed"/>
		</xsl:variable>
		<xsl:call-template name="findRelatedRelationships">
			<xsl:with-param name="id">
				<xsl:value-of select="$findPictureName"/>
			</xsl:with-param>
		</xsl:call-template>
	</xsl:template>
	<xsl:template name="findRelatedRelationships">
		<xsl:param name="id"/>
		<!--<xsl:for-each select="ancestor::p:cSld | ancestor::p:txStyles">-->
		<xsl:apply-templates select="document('word/_rels/document.xml.rels')/rel:Relationships">
			<xsl:with-param name="targetID">
				<xsl:value-of select="$id"/>
			</xsl:with-param>
		</xsl:apply-templates>
		<!--  </xsl:for-each>-->
	</xsl:template>
	<xsl:template match="rel:Relationships">
		<xsl:param name="targetID"/>
		<xsl:for-each select="rel:Relationship">
			<xsl:if test="@Id = $targetID">
				<xsl:value-of select="substring-before(substring-after(@Target,'media/'),'.')"/>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>
	<!--zhaobj 线颜色-->
	<xsl:template match="a:ln" mode="linecolor">
		<xsl:choose>
			<xsl:when test="a:solidFill">
        <图:线颜色_8058>
					<xsl:choose>
						<xsl:when test="a:solidFill/a:srgbClr">
							<xsl:value-of select="concat('#',a:solidFill/a:srgbClr/@val)"/>
						</xsl:when>
						<xsl:when test="a:solidFill/a:schemeClr">
							<xsl:apply-templates select="a:solidFill" mode="lnschemeClr"/>
						</xsl:when>
						<xsl:otherwise>
							<!--转换成默认的accent1-->
							<xsl:value-of select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val"/>
						</xsl:otherwise>
					</xsl:choose>
				</图:线颜色_8058>
			</xsl:when>
			<!--zhaobj 2/24  a:gradFil 是渐进色，UOF中没有对应，转换为第一个颜色-->
			<xsl:when test="a:gradFill">
        <图:线颜色_8058>
					<xsl:choose>
						<xsl:when test="a:gradFill/a:gsLst">
							<xsl:apply-templates select="a:gradFill/a:gsLst" mode="firstcolor"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val)"/>
						</xsl:otherwise>
					</xsl:choose>
				</图:线颜色_8058>
			</xsl:when>
      
      <!--2014-04-15，wudi，修复文本框颜色转换不正确的BUG，start-->
			<xsl:when test="a:noFill">
        <!--<图:线颜色_8058>#FFFFFF</图:线颜色_8058>-->
			</xsl:when>
      <!--end-->
      
			<!--uof中没有图案填充的情况，直接转换为默认-->
			<xsl:when test="a:pattFill">
        <图:线颜色_8058>
					<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val)"/>
				</图:线颜色_8058>
			</xsl:when>
			<xsl:otherwise>
				<!--zhaobj 转换到lnRef中找，没有就到主题theme中  <图:线颜色 uof:locID="g0013">#000000</图:线颜色>-->
				<xsl:choose>
					<xsl:when test="parent::wps:spPr/parent::wps:wsp/wps:style/a:lnRef/*">
						<xsl:for-each select="parent::wps:spPr/parent::wps:wsp/wps:style/a:lnRef">
              <图:线颜色_8058>
								<xsl:call-template name="lnRefcolor"/>
							</图:线颜色_8058>
						</xsl:for-each>
					</xsl:when>
					<xsl:when test="parent::wps:spPr/parent::wps:wsp/wps:style/a:lnRef/@idx">
						<xsl:variable name="idx" select="parent::wps:spPr/parent::wps:wsp/wps:style/a:lnRef/@idx"/>
						<xsl:apply-templates select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position()=$idx]" mode="linethemecolor"/>
					</xsl:when>

          <!--2013-11-13，wudi，Strict标准下，命名空间前缀为wp，start-->
          <xsl:when test="parent::wp:spPr/parent::wp:wsp/wp:style/a:lnRef/*">
            <xsl:for-each select="parent::wp:spPr/parent::wp:wsp/wp:style/a:lnRef">
              <图:线颜色_8058>
                <xsl:call-template name="lnRefcolor"/>
              </图:线颜色_8058>
            </xsl:for-each>
          </xsl:when>
          <xsl:when test="parent::wp:spPr/parent::wp:wsp/wp:style/a:lnRef/@idx">
            <xsl:variable name="idx" select="parent::wp:spPr/parent::wp:wsp/wp:style/a:lnRef/@idx"/>
            <xsl:apply-templates select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position()=$idx]" mode="linethemecolor"/>
          </xsl:when>
          <!--end-->
          
					<xsl:otherwise>
            
            <!--2013-11-13，wudi，少了"图:线颜色"的标签，start-->
            <图:线颜色_8058>
              <xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val)"/>
            </图:线颜色_8058>
            <!--end-->
            
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="a:solidFill" mode="lnschemeClr">
		<xsl:call-template name="colorChoice"/>
	</xsl:template>
  
  <!--2013-11-27，wudi，Strict标准下，pos取值有变，修正，start-->
	<xsl:template match="a:gradFill/a:gsLst" mode="firstcolor">
		<xsl:for-each select="node()">
      <xsl:variable name="tmp">
        <xsl:value-of select="number(substring-before(@pos,'%'))"/>
      </xsl:variable>
			<xsl:if test="$tmp = '0'">
				<xsl:value-of select="concat('#',a:srgbClr/@val)"/>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>
  <!--end-->
  
	<!--lnref中的线颜色转换-->
	<xsl:template name="lnRefcolor">
		<xsl:choose>
			<xsl:when test="a:schemeClr">
				<xsl:call-template name="colorChoice"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:if test="parent::wps:spPr/parent::wps:wsp/wps:style/a:lnRef/@idx">
					<xsl:variable name="idx" select="parent::wps:spPr/parent::wps:wsp/wps:style/a:lnRef/@idx"/>
					<xsl:apply-templates select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position()=$idx]" mode="linethemecolor"/>
				</xsl:if>

        <!--2013-11-13，wudi，Strict标准下，命名空间前缀为wp，start-->
        <xsl:if test="parent::wp:spPr/parent::wp:wsp/wp:style/a:lnRef/@idx">
          <xsl:variable name="idx" select="parent::wp:spPr/parent::wp:wsp/wp:style/a:lnRef/@idx"/>
          <xsl:apply-templates select="document('word/theme/theme1.xml')/a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln[position()=$idx]" mode="linethemecolor"/>
        </xsl:if>
        <!--end-->
        
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="a:ln" mode="linethemecolor">
		<xsl:for-each select="a:solidFill | a:noFill | a:gradFill | a:blipFill | a:pattFill">
      <图:线颜色_8058>
				<xsl:call-template name="colorChoice"/>
			</图:线颜色_8058>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="colorChoice">
		<xsl:choose>
			<xsl:when test="a:srgbClr">
				<xsl:value-of select="concat('#',a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when test="a:hslClr">
				<xsl:apply-templates select="a:hslClr"/>
			</xsl:when>
			<xsl:when test="a:prstClr">
				<xsl:apply-templates select="a:prstClr"/>
			</xsl:when>
			<xsl:when test="a:schemeClr">
				<xsl:apply-templates select="a:schemeClr"/>
			</xsl:when>
			<xsl:when test="a:scrgbClr">
				<xsl:apply-templates select="a:scrgbClr"/>
			</xsl:when>
			<xsl:when test="a:sysClr">
				<xsl:apply-templates select="a:sysClr"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="a:prstClr">
		<xsl:choose>
			<xsl:when test="@val='aliceBlue'">#f0f8ff</xsl:when>
			<xsl:when test="@val='antiqueWhite'">#faebd7</xsl:when>
			<xsl:when test="@val='aqua'">#00ffff</xsl:when>
			<xsl:when test="@val='aquamarine'">#7fffd4</xsl:when>
			<xsl:when test="@val='azure'">#f0ffff</xsl:when>
			<xsl:when test="@val='beige'">#f5f5dc</xsl:when>
			<xsl:when test="@val='bisque'">#ffe4c4</xsl:when>
			<xsl:when test="@val='black'">#000000</xsl:when>
			<xsl:when test="@val='blanchedAlmond'">#ffebcd</xsl:when>
			<xsl:when test="@val='blue'">#0000ff</xsl:when>
			<xsl:when test="@val='blueViolet'">#8a2be2</xsl:when>
			<xsl:when test="@val='brown'">#a52a2a</xsl:when>
			<xsl:when test="@val='burlyWood'">#deb887</xsl:when>
			<xsl:when test="@val='cadetBlue'">#5f9ea0</xsl:when>
			<xsl:when test="@val='chartreuse'">#7fff00</xsl:when>
			<xsl:when test="@val='chocolate'">#d2691e</xsl:when>
			<xsl:when test="@val='coral'">#ff7f50</xsl:when>
			<xsl:when test="@val='cornflowerBlue'">#6495ed</xsl:when>
			<xsl:when test="@val='cornsilk'">#fff8dc</xsl:when>
			<xsl:when test="@val='crimson'">#dc143c</xsl:when>
			<xsl:when test="@val='cyan'">#00ffff</xsl:when>
			<xsl:when test="@val='deepPink'">#ff1493</xsl:when>
			<xsl:when test="@val='deepSkyBlue'">#00bfff</xsl:when>
			<xsl:when test="@val='dimGray'">#696969</xsl:when>
			<xsl:when test="@val='dkBlue'">#00008b</xsl:when>
			<xsl:when test="@val='dkCyan'">#008b8b</xsl:when>
			<xsl:when test="@val='dkGoldenrod'">#b8860b</xsl:when>
			<xsl:when test="@val='dkGray'">#a9a9a9</xsl:when>
			<xsl:when test="@val='dkGreen'">#006400</xsl:when>
			<xsl:when test="@val='dkKhaki'">#bdb767</xsl:when>
			<xsl:when test="@val='dkMagenta'">#8b008b</xsl:when>
			<xsl:when test="@val='dkOliveGreen'">#556b2f</xsl:when>
			<xsl:when test="@val='dkOrange'">#ff8c00</xsl:when>
			<xsl:when test="@val='dkOrchild'">#9932cc</xsl:when>
			<xsl:when test="@val='dkRed'">#8b0000</xsl:when>
			<xsl:when test="@val='dkSalmon'">#e9967a</xsl:when>
			<xsl:when test="@val='dkSeaGreen'">#8fbc8b</xsl:when>
			<xsl:when test="@val='dkSlateBlue'">#483d8b</xsl:when>
			<xsl:when test="@val='dkSlateGray'">#2f4f4f</xsl:when>
			<xsl:when test="@val='dkTurquoise'">#00cece</xsl:when>
			<xsl:when test="@val='dkViolet'">#9400d3</xsl:when>
			<xsl:when test="@val='dodgerBlue'">#1e90ff</xsl:when>
			<xsl:when test="@val='firebrick'">#b22222</xsl:when>
			<xsl:when test="@val='floralWhite'">#fffaf0</xsl:when>
			<xsl:when test="@val='forestGreen'">#228b22</xsl:when>
			<xsl:when test="@val='fuchsia'">#ff00ff</xsl:when>
			<xsl:when test="@val='gainsboro'">#dcdcdc</xsl:when>
			<xsl:when test="@val='ghostWhite'">#f8f8ff</xsl:when>
			<xsl:when test="@val='gold'">#ffd700</xsl:when>
			<xsl:when test="@val='goldenrod'">#daa520</xsl:when>
			<xsl:when test="@val='gray'">#808080</xsl:when>
			<xsl:when test="@val='green'">#008000</xsl:when>
			<xsl:when test="@val='greenYellow'">#adff2f</xsl:when>
			<xsl:when test="@val='honeydew'">#f0fff0</xsl:when>
			<xsl:when test="@val='hotPink'">#ff69b4</xsl:when>
			<xsl:when test="@val='indianRed'">#cd5c5c</xsl:when>
			<xsl:when test="@val='indigo'">#4b0082</xsl:when>
			<xsl:when test="@val='ivory'">#fffff0</xsl:when>
			<xsl:when test="@val='khaki'">#f0e68c</xsl:when>
			<xsl:when test="@val='lavender'">#e6e6fa</xsl:when>
			<xsl:when test="@val='lavenderBlush'">#fff0f5</xsl:when>
			<xsl:when test="@val='lawnGreen'">#7cfc00</xsl:when>
			<xsl:when test="@val='lemonChiffon'">#fffacd</xsl:when>
			<xsl:when test="@val='lime'">#00ff00</xsl:when>
			<xsl:when test="@val='limeGreen'">#32cd32</xsl:when>
			<xsl:when test="@val='linen'">#faf0e6</xsl:when>
			<xsl:when test="@val='ltBlue'">#add8e6</xsl:when>
			<xsl:when test="@val='ltCoral'">#f08080</xsl:when>
			<xsl:when test="@val='ltCyan'">#e0ffff</xsl:when>
			<xsl:when test="@val='ltGoldenrodYellow'">#fafa78</xsl:when>
			<xsl:when test="@val='ltGray'">#d3d3d3</xsl:when>
			<xsl:when test="@val='ltGreen'">#90ee90</xsl:when>
			<xsl:when test="@val='ltPink'">#ffb6c1</xsl:when>
			<xsl:when test="@val='ltSalmon'">#ffa07a</xsl:when>
			<xsl:when test="@val='ltSeaGreen'">#20b2aa</xsl:when>
			<xsl:when test="@val='ltSkyBlue'">#87cefa</xsl:when>
			<xsl:when test="@val='ltSlateGray'">#778899</xsl:when>
			<xsl:when test="@val='ltSteelBlue'">#b0c4de</xsl:when>
			<xsl:when test="@val='ltYellow'">#ffffe0</xsl:when>
			<xsl:when test="@val='magenta'">#ff00ff</xsl:when>
			<xsl:when test="@val='maroon'">#800000</xsl:when>
			<xsl:when test="@val='medAquamarine'">#66cdaa</xsl:when>
			<xsl:when test="@val='medBlue'">#0000cd</xsl:when>
			<xsl:when test="@val='medOrchid'">#ba55d3</xsl:when>
			<xsl:when test="@val='medPurple'">#9370db</xsl:when>
			<xsl:when test="@val='medSeaGreen'">#3cb371</xsl:when>
			<xsl:when test="@val='medSlateBlue'">#7b68ee</xsl:when>
			<xsl:when test="@val='medSpringGreen'">#00fa9a</xsl:when>
			<xsl:when test="@val='medTurquoise'">#48d1cc</xsl:when>
			<xsl:when test="@val='medVioletRed'">#c71585</xsl:when>
			<xsl:when test="@val='midnightBlue'">#191970</xsl:when>
			<xsl:when test="@val='mintCream'">#f5fffa</xsl:when>
			<xsl:when test="@val='mistyRose'">#ffe4e1</xsl:when>
			<xsl:when test="@val='moccasin'">#ffe4b5</xsl:when>
			<xsl:when test="@val='navajoWhite'">#ffdead</xsl:when>
			<xsl:when test="@val='navy'">#000080</xsl:when>
			<xsl:when test="@val='oldLace'">#fdf5e6</xsl:when>
			<xsl:when test="@val='olive'">#808000</xsl:when>
			<xsl:when test="@val='oliveDrab'">#6b8e23</xsl:when>
			<xsl:when test="@val='orange'">#ffa500</xsl:when>
			<xsl:when test="@val='orangeRed'">#ff4500</xsl:when>
			<xsl:when test="@val='orchid'">#da70d6</xsl:when>
			<xsl:when test="@val='paleGoldenrod'">#eee8aa</xsl:when>
			<xsl:when test="@val='paleGreen'">#98fb98</xsl:when>
			<xsl:when test="@val='paleTurquoise'">#afeeee</xsl:when>
			<xsl:when test="@val='paleVioletRed'">#db7093</xsl:when>
			<xsl:when test="@val='papayaWhip'">#ffefd5</xsl:when>
			<xsl:when test="@val='peachPuff'">#ffdab9</xsl:when>
			<xsl:when test="@val='peru'">#cd853f</xsl:when>
			<xsl:when test="@val='pink'">#ffc0cb</xsl:when>
			<xsl:when test="@val='plum'">#dda0dd</xsl:when>
			<xsl:when test="@val='powderBlue'">#b0e0e6</xsl:when>
			<xsl:when test="@val='purple'">#800080</xsl:when>
			<xsl:when test="@val='red'">#ff0000</xsl:when>
			<xsl:when test="@val='rosyBrown'">#bc8f8f</xsl:when>
			<xsl:when test="@val='royalBlue'">#4169e1</xsl:when>
			<xsl:when test="@val='saddleBrown'">#8b4513</xsl:when>
			<xsl:when test="@val='salmon'">#fa8072</xsl:when>
			<xsl:when test="@val='sandyBrown'">#f4a460</xsl:when>
			<xsl:when test="@val='seaGreen'">#2e8b57</xsl:when>
			<xsl:when test="@val='seaShell'">#fff5ee</xsl:when>
			<xsl:when test="@val='sienna'">#a0522d</xsl:when>
			<xsl:when test="@val='silver'">#c0c0c0</xsl:when>
			<xsl:when test="@val='skyBlue'">#87ceeb</xsl:when>
			<xsl:when test="@val='slateBlue'">#6a5acd</xsl:when>
			<xsl:when test="@val='slateGray'">#708090</xsl:when>
			<xsl:when test="@val='snow'">#fffafa</xsl:when>
			<xsl:when test="@val='springGreen'">#00ff7f</xsl:when>
			<xsl:when test="@val='steelBlue'">#4682b4</xsl:when>
			<xsl:when test="@val='tan'">#d2b48c</xsl:when>
			<xsl:when test="@val='teal'">#008080</xsl:when>
			<xsl:when test="@val='thistle'">#d8bfd8</xsl:when>
			<xsl:when test="@val='tomato'">#ff6347</xsl:when>
			<xsl:when test="@val='turquoise'">#40e0d0</xsl:when>
			<xsl:when test="@val='violet'">#ee82ee</xsl:when>
			<xsl:when test="@val='wheat'">#f5deb3</xsl:when>
			<xsl:when test="@val='white'">#ffffff</xsl:when>
			<xsl:when test="@val='whiteSmoke'">#f5f5f5</xsl:when>
			<xsl:when test="@val='yellow'">#ffff00</xsl:when>
			<xsl:when test="@val='yellowGreen'">#9acd32</xsl:when>
		</xsl:choose>
	</xsl:template>
	<!-- zhaobj主题色-->
	<xsl:template match="a:schemeClr">
		<xsl:variable name="color">
			<xsl:choose>
				<xsl:when test="@val = 'bg1'">
					<xsl:value-of select="document('word/settings.xml')/w:settings/w:clrSchemeMapping/@w:bg1"/>
				</xsl:when>
				<xsl:when test="@val = 'bg2'">
					<xsl:value-of select="document('word/settings.xml')/w:settings/w:clrSchemeMapping/@w:bg2"/>
				</xsl:when>
				<xsl:when test="@val = 'tx1'">
					<xsl:value-of select="document('word/settings.xml')/w:settings/w:clrSchemeMapping/@w:t1"/>
				</xsl:when>
				<xsl:when test="@val = 'tx2'">
					<xsl:value-of select="document('word/settings.xml')/w:settings/w:clrSchemeMapping/@w:t2"/>
				</xsl:when>
				<xsl:otherwise>
					<!--xsl:value-of select="'#4F81BD'"/-->
					<xsl:value-of select="@val"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:call-template name="Findcolor">
			<xsl:with-param name="colorvalue" select="$color"/>
		</xsl:call-template>
	</xsl:template>
	<xsl:template name="Findcolor">
		<xsl:param name="colorvalue"/>
		<xsl:choose>
			<xsl:when test="$colorvalue = 'dk1'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:dk1/a:sysClr/@lastClr)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'dark1'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:dk1/a:sysClr/@lastClr)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'lt1'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:lt1/a:sysClr/@lastClr)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'light1'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:lt1/a:sysClr/@lastClr)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'dk2'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:dk2/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'dark2'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:dk2/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'lt2'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:lt2/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'light2'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:lt2/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'accent1'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'accent2'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent2/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'accent3'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent3/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'accent4'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent4/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'accent5'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent5/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'accent6'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent6/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'hlink'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:hlink/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:when test="$colorvalue = 'folHlink'">
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:folHlink/a:srgbClr/@val)"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="concat('#',document('word/theme/theme1.xml')/a:theme/a:themeElements/a:clrScheme/a:accent1/a:srgbClr/@val)"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
  <!--预定义图形中文本的转换-->

  <!--2013-11-05，wudi，Strict标准下为wp前缀，加以区分，start-->
  <xsl:template match="wp:txbx" mode="文本">
    <图:文本_803C>
      <!--<xsl:attribute name ="是否为文本框_8046">
        <xsl:value-of select ="'false'"/>
      </xsl:attribute>-->
      <xsl:attribute name="是否自动换行_8047">
        <xsl:if test="following-sibling::wp:bodyPr/@wrap='none'">
          <xsl:value-of select="'false'"/>
        </xsl:if>
        <xsl:if test="following-sibling::wp:bodyPr/@wrap='square'">
          <xsl:value-of select="'true'"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name="是否大小适应文字_8048">
        <xsl:if test="following-sibling::wp:bodyPr/a:spAutoFit">
          <xsl:value-of select="'true'"/>
        </xsl:if>
        <xsl:if test="following-sibling::wp:bodyPr/a:noAutofit">
          <xsl:value-of select="'false'"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name="是否文字随对象旋转_8049">
        <xsl:choose>
          <xsl:when test="following-sibling::wp:bodyPr/@upright='1' or following-sibling::wp:bodyPr/@upright='true'">
            <xsl:value-of select="'false'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'true'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
      <图:边距_803D>
        
        <!--2014-06-04，wudi，修改转换比例，有变化，start-->
        <xsl:attribute name="左_C608">
          <!--cxl,2012.5.4修改转换比例-->
          <xsl:value-of select="following-sibling::wp:bodyPr/@lIns div 12700"/>
        </xsl:attribute>
        <xsl:attribute name="右_C60A">
          <xsl:value-of select="following-sibling::wp:bodyPr/@rIns div 12700"/>
        </xsl:attribute>
        <xsl:attribute name="上_C609">
          <xsl:value-of select="following-sibling::wp:bodyPr/@tIns div 12700"/>
        </xsl:attribute>
        <xsl:attribute name="下_C60B">
          <xsl:value-of select="following-sibling::wp:bodyPr/@bIns div 12700"/>
        </xsl:attribute>
        <!--end-->
        
      </图:边距_803D>
      <图:对齐_803E>
        <xsl:attribute name="水平对齐_421D">
          <xsl:if test="following-sibling::wp:bodyPr/@lIns">
            <xsl:value-of select="following-sibling::wp:bodyPr/@lIns div 9525"/>
          </xsl:if>
          <xsl:if test="not(following-sibling::wp:bodyPr/@lIns)">
            <xsl:value-of select="'left'"/>
          </xsl:if>
        </xsl:attribute>
        <xsl:attribute name="水平对齐_421D">
          <xsl:choose>
            <xsl:when test="following-sibling::wp:bodyPr/@vert='eaVert'">
              <xsl:value-of select="'left'"/>
            </xsl:when>
            <xsl:when test="following-sibling::wp:bodyPr/@vert='horn'">
              <xsl:value-of select="'right'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'left'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:attribute>
        <xsl:attribute name="文字对齐_421E">
          <!--原来叫垂直对齐-->
          <xsl:choose>
            <xsl:when test="following-sibling::wp:bodyPr/@anchor='t'">
              <xsl:value-of select="'top'"/>
            </xsl:when>
            <xsl:when test="following-sibling::wp:bodyPr/@anchor='ctr'">
              <xsl:value-of select="'center'"/>
            </xsl:when>
            <xsl:when test="following-sibling::wp:bodyPr/@anchor='b'">
              <xsl:value-of select="'bottom'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </图:对齐_803E>
      <图:文字排列方向_8042>
        <xsl:if test="following-sibling::wp:bodyPr/@vert='horz'">
          <xsl:choose>
            <xsl:when test="parent::wp:wsp/@normalEastAsianFlow='1'">
              <xsl:value-of select="'t2b-l2r-270e-0w'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'t2b-l2r-0e-0w'"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <xsl:if test="following-sibling::wp:bodyPr/@vert='eaVert'">
          <xsl:value-of select="'r2l-t2b-0e-90w'"/>
        </xsl:if>
        <xsl:if test="following-sibling::wp:bodyPr/@vert='vert'">
          <xsl:value-of select="'r2l-t2b-90e-90w'"/>
        </xsl:if>
        <xsl:if test="following-sibling::wp:bodyPr/@vert='vert270'">
          <xsl:value-of select="'l2r-b2t-270e-270w'"/>
        </xsl:if>
        <!--第五种方向-->
      </图:文字排列方向_8042>
      <图:内容_8043>
        
        <!--2014-04-08，wudi，之前没有考虑文本框里包含文字表的情况，添加之，start-->
        <xsl:for-each select=".//wne:txbxContent/w:tbl">
          <xsl:call-template name="table">
            <xsl:with-param name ="tblPartFrom" select ="'document'"/>
          </xsl:call-template>
        </xsl:for-each>
        <!--end-->
        
        <!--2015-04-26，wudi，修复文本内容转换BUG，针对sdt内容控件做特殊处理，start-->
        <xsl:for-each select=".//wne:txbxContent/w:p | .//wne:txbxContent/w:sdt/w:sdtContent/w:p">
          <xsl:call-template name="paragraph">
            <xsl:with-param name="pPartFrom" select="'txbody'"/>
          </xsl:call-template>
        </xsl:for-each>

        <!--2013-04-02，修复页眉页脚内容丢失的BUG，没有考虑子节点为w:sdt的情况，start-->
        <!--<xsl:for-each select =".//wne:txbxContent/w:sdt/w:sdtContent/w:p">
          <xsl:call-template name="paragraph">
            <xsl:with-param name="pPartFrom" select="'txbody'"/>
          </xsl:call-template>
        </xsl:for-each>-->
             
        <!--end-->
        <!--end-->

      </图:内容_8043>
    </图:文本_803C>
  </xsl:template>
  <!--end-->

  <!--2014-06-05，wudi，注释掉此段代码，此模板是用于Transitional版本的，这里不适用，start-->
	<!--<xsl:template match="wps:txbx" mode="文本">
    <图:文本_803C>
			--><!--<xsl:attribute name ="是否为文本框_8046">
        <xsl:value-of select ="'false'"/>
      </xsl:attribute>--><!--
      <xsl:attribute name="是否自动换行_8047">
        <xsl:if test="following-sibling::wps:bodyPr/@wrap='none'">
          <xsl:value-of select="'false'"/>
        </xsl:if>
        <xsl:if test="following-sibling::wps:bodyPr/@wrap='square'">
          <xsl:value-of select="'true'"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name="是否大小适应文字_8048">
        <xsl:if test="following-sibling::wps:bodyPr/a:spAutoFit">
          <xsl:value-of select="'true'"/>
        </xsl:if>
        <xsl:if test="following-sibling::wps:bodyPr/a:noAutofit">
          <xsl:value-of select="'false'"/>
        </xsl:if>
      </xsl:attribute>
      <xsl:attribute name="是否文字随对象旋转_8049">
        <xsl:choose>
          <xsl:when test="following-sibling::wps:bodyPr/@upright='1' or following-sibling::wps:bodyPr/@upright='true'">
            <xsl:value-of select="'false'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'true'"/>
          </xsl:otherwise>
        </xsl:choose>      
      </xsl:attribute>
      <图:边距_803D>

        --><!--2014-06-05，wudi，修改转换比例，有变化，start--><!--
			  <xsl:attribute name="左_C608">--><!--cxl,2012.5.4修改转换比例--><!--
				  <xsl:value-of select="following-sibling::wps:bodyPr/@lIns div 12700"/>
			  </xsl:attribute>
			  <xsl:attribute name="右_C60A">
				  <xsl:value-of select="following-sibling::wps:bodyPr/@rIns div 12700"/>
			  </xsl:attribute>
			  <xsl:attribute name="上_C609">
				  <xsl:value-of select="following-sibling::wps:bodyPr/@tIns div 12700"/>
			  </xsl:attribute>
			  <xsl:attribute name="下_C60B">
				  <xsl:value-of select="following-sibling::wps:bodyPr/@bIns div 12700"/>
			  </xsl:attribute>
        --><!--end--><!--
        
      </图:边距_803D>
      <图:对齐_803E>
			  <xsl:attribute name="水平对齐_421D">
          <xsl:if test="following-sibling::wps:bodyPr/@lIns">
            <xsl:value-of select="following-sibling::wps:bodyPr/@lIns div 9525"/>
          </xsl:if>
          <xsl:if test="not(following-sibling::wps:bodyPr/@lIns)">
            <xsl:value-of select="'left'"/>
          </xsl:if>
			  </xsl:attribute>
			  <xsl:attribute name="水平对齐_421D">
          <xsl:choose>
            <xsl:when test="following-sibling::wps:bodyPr/@vert='eaVert'">
              <xsl:value-of select="'left'"/>
            </xsl:when>
            <xsl:when test="following-sibling::wps:bodyPr/@vert='horn'">
              <xsl:value-of select="'right'"/>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'left'"/>
            </xsl:otherwise>
          </xsl:choose>				  
			  </xsl:attribute>
			  <xsl:attribute name="文字对齐_421E">--><!--原来叫垂直对齐--><!--
				  <xsl:choose>
					  <xsl:when test="following-sibling::wps:bodyPr/@anchor='t'">
						  <xsl:value-of select="'top'"/>
					  </xsl:when>
					  <xsl:when test="following-sibling::wps:bodyPr/@anchor='ctr'">
						  <xsl:value-of select="'center'"/>
					  </xsl:when>
					  <xsl:when test="following-sibling::wps:bodyPr/@anchor='b'">
						  <xsl:value-of select="'bottom'"/>
					  </xsl:when>
				  </xsl:choose>
			  </xsl:attribute>
      </图:对齐_803E>
      <图:文字排列方向_8042>
				<xsl:if test="following-sibling::wps:bodyPr/@vert='horz'">
					<xsl:choose>
						<xsl:when test="parent::wps:wsp/@normalEastAsianFlow='1'">
							<xsl:value-of select="'t2b-l2r-270e-0w'"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'t2b-l2r-0e-0w'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:if>
				<xsl:if test="following-sibling::wps:bodyPr/@vert='eaVert'">
					<xsl:value-of select="'r2l-t2b-0e-90w'"/>
				</xsl:if>
				<xsl:if test="following-sibling::wps:bodyPr/@vert='vert'">
					<xsl:value-of select="'r2l-t2b-90e-90w'"/>
				</xsl:if>
				<xsl:if test="following-sibling::wps:bodyPr/@vert='vert270'">
					<xsl:value-of select="'l2r-b2t-270e-270w'"/>
				</xsl:if>
				--><!--第五种方向--><!--
      </图:文字排列方向_8042>			
      <图:内容_8043>
			  <xsl:for-each select=".//w:txbxContent/w:p">
          <xsl:call-template name="paragraph">
            <xsl:with-param name="pPartFrom" select="'txbody'"/>
          </xsl:call-template>
			  </xsl:for-each>

        --><!--2013-04-02，修复页眉页脚内容丢失的BUG，没有考虑子节点为w:sdt的情况，start--><!--
        <xsl:for-each select =".//w:txbxContent/w:sdt/w:sdtContent/w:p">
          <xsl:call-template name="paragraph">
            <xsl:with-param name="pPartFrom" select="'txbody'"/>
          </xsl:call-template>
        </xsl:for-each>
        --><!--end--><!--
        
      </图:内容_8043>
		</图:文本_803C>
	</xsl:template>-->
  <!--end-->

  <!--三维效果转换模板 CXL2011年11月20日添加-->
  <xsl:template match="wps:spPr/a:scene3d | dsp:spPr/a:scene3d | wp:spPr/a:scene3d">
    <xsl:if test="not(../a:sp3d/@extrusionH)"><!--无深度-->
      <xsl:for-each select="./a:camera">
        <xsl:if test="./a:rot">
          <uof:角度_C635>
            <uof:x方向_C636>
              <xsl:value-of select="./a:rot/@lon div 60000"/>
            </uof:x方向_C636>
            <uof:y方向_C637>
              <xsl:value-of select="./a:rot/@lat div 60000"/>
            </uof:y方向_C637>
          </uof:角度_C635>
          <uof:方向_C63C>
            <uof:角度_C639>
              <xsl:value-of select="'none'"/>
            </uof:角度_C639>
            <xsl:choose>
              <xsl:when test="./@prst='orthographicFront'"><!--无旋转，只有在设置了三维效果时会出现-->
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='isometricLeftDown'"><!--等轴左下-->
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='isometricRightUp'"><!--等轴右上-->
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='isometricTopUp'"><!--等长顶部朝上-->
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='isometricBottomDown'"><!--等长底部朝下-->
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='isometricOffAxis1Left'"><!--离轴1左-->
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='isometricOffAxis1Right'"><!--离轴1右-->
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='isometricOffAxis1Top'"><!--离轴1上-->
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='isometricOffAxis2Left'"><!--离轴2左-->
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='isometricOffAxis2Right'"><!--离轴2右-->
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='isometricOffAxis2Top'"><!--离轴2上-->
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='perspectiveFront'"><!--前透视-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='perspectiveLeft'"><!--左透视-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='perspectiveRight'"><!--右透视-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='perspectiveBelow'"><!--下透视-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='perspectiveAbove'"><!--上透视-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='perspectiveRelaxedModerately'"><!--适度宽松透视-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='perspectiveRelaxed'"><!--宽松透视-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='perspectiveContrastingLeftFacing'"><!--左向对比透视-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='perspectiveContrastingRightFacing'"><!--右向对比透视-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='perspectiveHeroicExtremeLeftFacing'"><!--极左极大透视-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='perspectiveHeroicExtremeRightFacing'"><!--极右极大透视-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='obliqueTopLeft'"><!--倾斜左上-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='obliqueTopRight'"><!--倾斜右上-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='obliqueBottomLeft'"><!--倾斜左下-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:when test="./@prst='obliqueBottomRight'"><!--倾斜右下-->
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </xsl:when>
              <xsl:otherwise>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </xsl:otherwise>
            </xsl:choose>
          </uof:方向_C63C>
        </xsl:if>
        <xsl:if test="not(./a:rot)"><!--如果没有这个标签，则根据预设属性将值写死-->
          <xsl:choose>
            <xsl:when test="./@prst='orthographicFront'"><!--无旋转，只有在设置了三维效果时会出现-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricLeftDown'"><!--等轴左下-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>45.0</uof:x方向_C636>
                <uof:y方向_C637>35.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricRightUp'"><!--等轴右上-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>315.0</uof:x方向_C636>
                <uof:y方向_C637>35.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricTopUp'"><!--等长顶部朝上-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>314.7</uof:x方向_C636>
                <uof:y方向_C637>324.6</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricBottomDown'"><!--等长底部朝下-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>314.7</uof:x方向_C636>
                <uof:y方向_C637>35.4</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis1Left'"><!--离轴1左-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>64.0</uof:x方向_C636>
                <uof:y方向_C637>18.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis1Right'"><!--离轴1右-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>334.0</uof:x方向_C636>
                <uof:y方向_C637>18.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis1Top'"><!--离轴1上-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>306.5</uof:x方向_C636>
                <uof:y方向_C637>301.3</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis2Left'"><!--离轴2左-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>26.0</uof:x方向_C636>
                <uof:y方向_C637>18.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis2Right'"><!--离轴2右-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>296.0</uof:x方向_C636>
                <uof:y方向_C637>18.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis2Top'"><!--离轴2上-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>53.5</uof:x方向_C636>
                <uof:y方向_C637>301.3</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveFront'"><!--前透视-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveLeft'"><!--左透视-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>20.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveRight'"><!--右透视-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>340.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveBelow'"><!--下透视-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>20.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveAbove'"><!--上透视-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>340.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveRelaxedModerately'"><!--适度宽松透视-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>324.8</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveRelaxed'"><!--宽松透视-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>309.6</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveContrastingLeftFacing'"><!--左向对比透视-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>43.9</uof:x方向_C636>
                <uof:y方向_C637>10.4</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveContrastingRightFacing'"><!--右向对比透视-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>316.1</uof:x方向_C636>
                <uof:y方向_C637>10.4</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveHeroicExtremeLeftFacing'"><!--极左极大透视-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>34.5</uof:x方向_C636>
                <uof:y方向_C637>8.1</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveHeroicExtremeRightFacing'"><!--极右极大透视-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>325.5</uof:x方向_C636>
                <uof:y方向_C637>8.1</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='obliqueTopLeft'"><!--倾斜左上-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='obliqueTopRight'"><!--倾斜右上-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='obliqueBottomLeft'"><!--倾斜左下-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='obliqueBottomRight'"><!--倾斜右下-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:otherwise>
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
    <xsl:if test="../a:sp3d/@extrusionH"><!--深度-->
      <xsl:for-each select="./a:camera">
        <xsl:if test="./a:rot">
          <uof:角度_C635>
            <uof:x方向_C636>
              <xsl:value-of select="./a:rot/@lon div 60000"/>
            </uof:x方向_C636>
            <uof:y方向_C637>
              <xsl:value-of select="./a:rot/@lat div 60000"/>
            </uof:y方向_C637>
          </uof:角度_C635>         
          <xsl:choose>
            <xsl:when test="./@prst='orthographicFront'"><!--无旋转，只有在设置了三维效果时会出现-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'none'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='isometricLeftDown'"><!--等轴左下-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-right'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='isometricRightUp'"><!--等轴右上-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-left'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='isometricTopUp'"><!--等长顶部朝上-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-left'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='isometricBottomDown'"><!--等长底部朝下-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-right'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis1Left'"><!--离轴1左-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-right'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis1Right'"><!--离轴1右-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-left'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis1Top'"><!--离轴1上-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-left'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis2Left'"><!--离轴2左-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-right'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis2Right'"><!--离轴2右-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-left'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis2Top'"><!--离轴2上-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-right'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveFront'"><!--前透视-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'none'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveLeft'"><!--左透视-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-right'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveRight'"><!--右透视-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-left'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveBelow'"><!--下透视-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveAbove'"><!--上透视-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-bottom'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveRelaxedModerately'"><!--适度宽松透视-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-bottom'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveRelaxed'"><!--宽松透视-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-bottom'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveContrastingLeftFacing'"><!--左向对比透视-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-bottom-right'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveContrastingRightFacing'"><!--右向对比透视-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-bottom-left'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveHeroicExtremeLeftFacing'"><!--极左极大透视-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-bottom-right'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveHeroicExtremeRightFacing'"><!--极右极大透视-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-bottom-left'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='obliqueTopLeft'"><!--倾斜左上-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-left'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='obliqueTopRight'"><!--倾斜右上-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-right'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='obliqueBottomLeft'"><!--倾斜左下-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-bottom-left'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:when test="./@prst='obliqueBottomRight'"><!--倾斜右下-->
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-bottom-right'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:when>
            <xsl:otherwise>
              <uof:方向_C63C>
                <uof:角度_C639>
                  <xsl:value-of select="'to-top-right'"/>
                </uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
        <xsl:if test="not(./a:rot)">
          <!--如果没有这个标签，则根据预设属性将值写死-->
          <xsl:choose>
            <xsl:when test="./@prst='orthographicFront'"><!--无旋转，只有在设置了三维效果时会出现-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricLeftDown'"><!--等轴左下-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top-right</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>45.0</uof:x方向_C636>
                <uof:y方向_C637>35.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricRightUp'"><!--等轴右上-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top-left</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>315.0</uof:x方向_C636>
                <uof:y方向_C637>35.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricTopUp'"><!--等长顶部朝上-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top-left</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>314.7</uof:x方向_C636>
                <uof:y方向_C637>324.6</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricBottomDown'"><!--等长底部朝下-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top-right</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>314.7</uof:x方向_C636>
                <uof:y方向_C637>35.4</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis1Left'"><!--离轴1左-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top-right</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>64.0</uof:x方向_C636>
                <uof:y方向_C637>18.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis1Right'"><!--离轴1右-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top-left</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>334.0</uof:x方向_C636>
                <uof:y方向_C637>18.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis1Top'"><!--离轴1上-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top-left</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>306.5</uof:x方向_C636>
                <uof:y方向_C637>301.3</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis2Left'"><!--离轴2左-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top-right</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>26.0</uof:x方向_C636>
                <uof:y方向_C637>18.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis2Right'"><!--离轴2右-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top-left</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>296.0</uof:x方向_C636>
                <uof:y方向_C637>18.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='isometricOffAxis2Top'"><!--离轴2上-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top-right</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>53.5</uof:x方向_C636>
                <uof:y方向_C637>301.3</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveFront'"><!--前透视-->
              <uof:方向_C63C>
                <uof:角度_C639>none</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveLeft'"><!--左透视-->
              <uof:方向_C63C>
                <uof:角度_C639>to-right</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>20.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveRight'"><!--右透视-->
              <uof:方向_C63C>
                <uof:角度_C639>to-left</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>340.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveBelow'"><!--下透视-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>20.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveAbove'"><!--上透视-->
              <uof:方向_C63C>
                <uof:角度_C639>to-bottom</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>340.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveRelaxedModerately'">
              <!--适度宽松透视-->
              <uof:方向_C63C>
                <uof:角度_C639>to-bottom</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>324.8</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveRelaxed'"><!--宽松透视-->
              <uof:方向_C63C>
                <uof:角度_C639>to-bottom</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>309.6</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveContrastingLeftFacing'"><!--左向对比透视-->
              <uof:方向_C63C>
                <uof:角度_C639>to-bottom-right</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>43.9</uof:x方向_C636>
                <uof:y方向_C637>10.4</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveContrastingRightFacing'"><!--右向对比透视-->
              <uof:方向_C63C>
                <uof:角度_C639>to-bottom-left</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>316.1</uof:x方向_C636>
                <uof:y方向_C637>10.4</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveHeroicExtremeLeftFacing'"><!--极左极大透视-->
              <uof:方向_C63C>
                <uof:角度_C639>to-bottom-right</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>34.5</uof:x方向_C636>
                <uof:y方向_C637>8.1</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='perspectiveHeroicExtremeRightFacing'"><!--极右极大透视-->
              <uof:方向_C63C>
                <uof:角度_C639>to-bottom-left</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'perspective'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>325.5</uof:x方向_C636>
                <uof:y方向_C637>8.1</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='obliqueTopLeft'"><!--倾斜左上-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top-left</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='obliqueTopRight'"><!--倾斜右上-->
              <uof:方向_C63C>
                <uof:角度_C639>to-top-right</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='obliqueBottomLeft'"><!--倾斜左下-->
              <uof:方向_C63C>
                <uof:角度_C639>to-bottom-left</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:when test="./@prst='obliqueBottomRight'"><!--倾斜右下-->
              <uof:方向_C63C>
                <uof:角度_C639>to-bottom-right</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:when>
            <xsl:otherwise>
              <uof:方向_C63C>
                <uof:角度_C639>to-top-right</uof:角度_C639>
                <uof:方式_C63D>
                  <xsl:value-of select="'parallel'"/>
                </uof:方式_C63D>
              </uof:方向_C63C>
              <uof:角度_C635>
                <uof:x方向_C636>0.0</uof:x方向_C636>
                <uof:y方向_C637>0.0</uof:y方向_C637>
              </uof:角度_C635>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:for-each>
    </xsl:if>
    <xsl:if test="../a:sp3d">
      <xsl:if test="../a:sp3d/a:extrusionClr/a:srgbClr">
        <uof:颜色_C63F>
          <xsl:variable name ="color">
            <xsl:value-of select="../a:sp3d/a:extrusionClr/a:srgbClr/@val"/>
          </xsl:variable>
          <xsl:value-of select ="concat('#',$color)"/>
        </uof:颜色_C63F>
      </xsl:if>
      <xsl:if test="../a:sp3d/@extrusionH">
        <uof:深度_C63B>
          <xsl:value-of select="../a:sp3d/@extrusionH div 12700"/>
        </uof:深度_C63B>       
      </xsl:if>
      <xsl:for-each select="../a:sp3d">
        <uof:表面效果_C63E>
          <xsl:choose>
            <xsl:when test="../a:sp3d/@prstMaterial='legacyWireframe'">wire-frame</xsl:when><!--线框-->
            <!--以下效果均转化为UOF中亚光效果-->
            <xsl:when test="../a:sp3d/@prstMaterial='matte'">matte</xsl:when><!--亚光效果-->
            <xsl:when test="not(../a:sp3d/@prstMaterial)">matte</xsl:when><!--默认为暖色粗糙-->
            <xsl:when test="../a:sp3d/@prstMaterial='dkEdge'">matte</xsl:when><!--硬边缘-->
            <!--以下效果均转化为UOF中塑料效果-->
            <xsl:when test="../a:sp3d/@prstMaterial='plastic'">plastic</xsl:when><!--塑料效果-->
            <xsl:when test="../a:sp3d/@prstMaterial='flat'">plastic</xsl:when><!--平面-->
            <xsl:when test="../a:sp3d/@prstMaterial='powder'">plastic</xsl:when><!--粉-->
            <xsl:when test="../a:sp3d/@prstMaterial='translucentPowder'">plastic</xsl:when><!--半透明粉-->
            <!--以下效果均转化为UOF中金属效果-->
            <xsl:when test="../a:sp3d/@prstMaterial='metal'">metal</xsl:when><!--金属效果-->
            <xsl:when test="../a:sp3d/@prstMaterial='softEdge'">metal</xsl:when><!--柔边缘-->
            <xsl:when test="../a:sp3d/@prstMaterial='clear'">metal</xsl:when><!--最浅-->
            <xsl:otherwise>matte</xsl:otherwise>
          </xsl:choose>
        </uof:表面效果_C63E>
      </xsl:for-each>
    </xsl:if>
    <xsl:if test="not(../a:sp3d)">
      <uof:深度_C63B>0</uof:深度_C63B>
      <uof:表面效果_C63E>matte</uof:表面效果_C63E>
    </xsl:if>      
    <xsl:for-each select="a:lightRig">
      <uof:照明_C638>        
        <uof:角度_C639><!--照明角度-->
          <xsl:choose>
            <!--<xsl:when test="@dir='t'">0</xsl:when>--><!--dir属性只发现了取t值的情况-->
            <xsl:when test="(./a:rot/@rev div 60000)&gt;= 0 and (./a:rot/@rev div 60000)&lt;=30 ">0</xsl:when>
            <xsl:when test="(./a:rot/@rev div 60000)&gt;=31 and (./a:rot/@rev div 60000)&lt;=70">45</xsl:when>
            <xsl:when test="(./a:rot/@rev div 60000)&gt;=71 and (./a:rot/@rev div 60000)&lt;=120">90</xsl:when>
            <xsl:when test="(./a:rot/@rev div 60000)&gt;=121 and (./a:rot/@rev div 60000)&lt;=160">135</xsl:when>
            <xsl:when test="(./a:rot/@rev div 60000)&gt;=161 and (./a:rot/@rev div 60000)&lt;=200">180</xsl:when>
            <xsl:when test="(./a:rot/@rev div 60000)&gt;=201 and (./a:rot/@rev div 60000)&lt;=240">225</xsl:when>
            <xsl:when test="(./a:rot/@rev div 60000)&gt;=241 and (./a:rot/@rev div 60000)&lt;=300">270</xsl:when>
            <xsl:when test="(./a:rot/@rev div 60000)&gt;=301 and (./a:rot/@rev div 60000)&lt;=360">315</xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>          
        </uof:角度_C639>
        <uof:强度_C63A><!--照明效果-->
          <xsl:choose>
            <!--以下效果转化为UOF中的bright效果-->
            <xsl:when test="@rig='balanced'">bright</xsl:when><!--平衡-->
            <xsl:when test="@rig='soft'">bright</xsl:when><!--柔和-->
            <xsl:when test="@rig='contrasting'">bright</xsl:when><!--对比-->
            <xsl:when test="@rig='sunrise'">bright</xsl:when><!--日出-->
            <xsl:when test="@rig='flat'">bright</xsl:when><!--平面-->
            <xsl:when test="@rig='glow'">bright</xsl:when><!--发光-->
            <xsl:when test="@rig='brightRoom'">bright</xsl:when><!--明亮的房间-->
            <!--以下效果转化为UOF中的normal效果-->
            <xsl:when test="@rig='threePt'">normal</xsl:when><!--三点-->
            <xsl:when test="@rig='flood'">normal</xsl:when><!--强烈-->
            <xsl:when test="@rig='morning'">normal</xsl:when><!--早晨-->
            <xsl:when test="@rig='chilly'">normal</xsl:when><!--寒冷-->
            <xsl:when test="@rig='twoPt'">normal</xsl:when><!--两点-->
            <!--以下效果转化为UOF中的dim效果-->
            <xsl:when test="@rig='harsh'">dim</xsl:when><!--粗糙-->
            <xsl:when test="@rig='sunset'">dim</xsl:when><!--日落-->
            <xsl:when test="@rig='freezing'">dim</xsl:when><!--冰冻-->
            <xsl:otherwise>normal</xsl:otherwise>
          </xsl:choose>        
        </uof:强度_C63A>         
      </uof:照明_C638>
    </xsl:for-each>
    <uof:是否显示效果_C640>true</uof:是否显示效果_C640>    
  </xsl:template>
  
</xsl:stylesheet>
