<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
                xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                xmlns:dc="http://purl.org/dc/elements/1.1/"
                xmlns:dcterms="http://purl.org/dc/terms/"
                xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                xmlns:cus="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">

    <xsl:template name="sheetprop">
        <表:工作表属性_E80D>
            <!--页面设置-->
            <xsl:if test="//ws:tabColor[@rgb]">
                <xsl:apply-templates select="ws:worksheet/ws:sheetPr"/>
            </xsl:if>
            <xsl:choose>
                <xsl:when test="ws:worksheet">
                    <!--页面设置 修改添加，李杨2011-12-05-->
                    <表:页面设置_E7C1 名称_E7D4="页面设置">
                        <xsl:choose>
                            <xsl:when test ="ws:worksheet/ws:pageSetup">
                                <xsl:apply-templates select="ws:worksheet/ws:pageSetup"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <表:纸张_E7C2 长_C604="842.0" 宽_C605="595.3499755859375"/>
                                <表:纸张方向_E7C3>portrait</表:纸张方向_E7C3>
                                <表:缩放_E7C4>100</表:缩放_E7C4>
                                <表:页边距_E7C5 左_C608="89.8499984741211" 上_C609="72.0" 右_C60A="89.8499984741211"
                                    下_C60B="72.0"/>
                                <表:打印_E7CA 是否带网格线_E7CB="false" 是否带行号列标_E7CC="false" 是否按草稿方式_E7CD="false"
                                    是否先列后行_E7CE="true"/>
                                <表:批注打印方式_E7CF>none</表:批注打印方式_E7CF>
                                <表:错误单元格打印方式_E7D0>none</表:错误单元格打印方式_E7D0>
                                <表:垂直对齐方式_E701>top</表:垂直对齐方式_E701>
                                <表:水平对齐方式_E700>left</表:水平对齐方式_E700>
                            </xsl:otherwise>
                        </xsl:choose>
                    </表:页面设置_E7C1>
                    <xsl:apply-templates select="ws:worksheet/ws:sheetViews"/>
                </xsl:when>
                <xsl:when test="ws:chartsheet">
                    <xsl:apply-templates select="ws:chartsheet/ws:sheetPr"/>
                    <表:页面设置_E7C1 名称_E7D4="页面设置">
                        <xsl:apply-templates select="ws:chartsheet/ws:pageSetup"/>
                        <xsl:if test ="not(ws:chartsheet/ws:pageSetup) and (ws:chartsheet/ws:pageMargins)">
                            <xsl:apply-templates select ="ws:chartsheet/ws:pageMargins"/>
                        </xsl:if>
                        <!--<xsl:apply-templates select="ws:chartsheet/ws:pageMargins"/>
            <xsl:apply-templates select="ws:chartsheet/ws:printOptions"/>-->
                    </表:页面设置_E7C1>
                    <xsl:apply-templates select="ws:chartsheet/ws:sheetViews"/>
                </xsl:when>
            </xsl:choose>
            <xsl:if test="ws:worksheet/ws:picture">
                <表:背景填充_E830>
                    <图:图片_8005>
                        <xsl:variable name="rId">
                            <xsl:value-of select="./ws:worksheet/ws:picture/@r:id"/>
                        </xsl:variable>
                        <xsl:variable name="target">
                            <xsl:value-of select="./pr:Relationships/pr:Relationship[@Id = $rId]/@Target"/>
                        </xsl:variable>
                        <xsl:attribute name="图形引用_8007">
                            <xsl:value-of select="substring-before(substring-after($target,'../media/'),'.')"/>
                        </xsl:attribute>
                    </图:图片_8005>
                </表:背景填充_E830>
            </xsl:if>
            <!--页面设置中垂直方式和水平方式还没有写，跨文件，师姐的意思是不用转了，直接对应于UOF的锚点就行了-->
        </表:工作表属性_E80D>
    </xsl:template>

    <!--标签背景色-->
    <xsl:template match="ws:sheetPr">
        <xsl:if test="ws:tabColor[@rgb]">
            <表:标签背景色_E7C0>
                <xsl:variable name="color">
                    <xsl:value-of select="ws:tabColor/@rgb"/>
                </xsl:variable>
                <xsl:variable name="color2">
                    <xsl:value-of select="substring($color,3,6)"/>
                </xsl:variable>
                <xsl:variable name="color3">
                    <xsl:value-of select="concat('#',$color2)"/>
                </xsl:variable>
                <xsl:value-of select="$color3"/>
            </表:标签背景色_E7C0>
        </xsl:if>
        <!--添加标签背景色 主题色引用，李杨2012-3-9-->
        <xsl:if test ="ws:tabColor[@theme]">
            <表:标签背景色_E7C0>
                <xsl:variable name="the" select="ws:tabColor/@theme"/>
                <xsl:variable name="them" select="/ws:spreadsheets/a:theme/a:themeElements/a:clrScheme/*[position()=$the]/a:sysClr/@val"/>
                <xsl:variable name="them2" select="/ws:spreadsheets/a:theme/a:themeElements/a:clrScheme/*[position()=($the+1)]/a:srgbClr/@val"/>
                <xsl:choose>
                    <xsl:when test="$them='windowText'">
                        <xsl:value-of select="'#000000'"/>
                    </xsl:when>
                    <xsl:when test="$them='window'">
                        <xsl:value-of select="'#000000'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="concat('#',$them2)"/>
                    </xsl:otherwise>
                </xsl:choose>
            </表:标签背景色_E7C0>
        </xsl:if>
    </xsl:template>

    <xsl:template match="ws:pageSetup">
        <xsl:choose>

            <!-- update by 凌峰 BUG_3060：纸张大小宽度与高度弄反 20140308 start -->
            <xsl:when test="@paperSize">
                <xsl:if test="@orientation = 'portrait'">
                    <xsl:if test="@paperSize='1'">
                        <表:纸张_E7C2 宽_C605="612.0999755859375" 长_C604="792.0999755859375"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='3'">
                        <表:纸张_E7C2 宽_C605="792.0999755859375" 长_C604="1224.1500244140625"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='5'">
                        <表:纸张_E7C2 长_C604="612.0999755859375" 宽_C605="1008.1500244140625"/>
                    </xsl:if>
                    <!--添加executive纸张大小类型，李杨2012-3-9-->
                    <!--添加11x17、信封、双层日式明信片、16K、executive(JIS)、B5、8K的纸张大小类型，凌峰 2014-1-19-->
                    <xsl:if test ="@paperSize='7'">
                        <表:纸张_E7C2 长_C604="756.0999755859375" 宽_C605="521.9000244140625"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='8'">
                        <表:纸张_E7C2 宽_C605="842.0" 长_C604="1190.699951171875"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='9'">
                        <表:纸张_E7C2 宽_C605="595.3499755859375" 长_C604="842.0"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='11'">
                        <表:纸张_E7C2 长_C604="419.6000061035156" 宽_C605="595.3499755859375"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='12'">
                        <表:纸张_E7C2 宽_C605="728.5999755859375" 长_C604="1031.949951171875"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='13'">
                        <表:纸张_E7C2 宽_C605="515.9500122070312" 长_C604="728.5999755859375"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='17'">
                        <表:纸张_E7C2 宽_C605="792.0999755859375" 长_C604="1224.1500244140625"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='20'">
                        <表:纸张_E7C2 宽_C605="296.79998779296875" 长_C604="684.0999755859375"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='27'">
                        <表:纸张_E7C2 宽_C605="311.8500061035156" 长_C604="623.7000122070312"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='28'">
                        <表:纸张_E7C2 宽_C605="459.25" 长_C604="649.2000122070312"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='34'">
                        <表:纸张_E7C2 宽_C605="498.95001220703125" 长_C604="708.75"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='37'">
                        <表:纸张_E7C2 宽_C605="278.95001220703125" 长_C604="540.0499877929688"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='43'">
                        <表:纸张_E7C2 宽_C605="283.5" 长_C604="419.6000061035156"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='70'">
                        <表:纸张_E7C2 宽_C605="297.70001220703125" 长_C604="419.6000061035156"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='82'">
                        <表:纸张_E7C2 宽_C605="419.6000061035156" 长_C604="567.0"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='119'">
                        <表:纸张_E7C2 宽_C605="557.9500122070312" 长_C604="773.9500122070312"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='120'">
                        <表:纸张_E7C2 宽_C605="612.0999755859375" 长_C604="935.25"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='121'">
                        <表:纸张_E7C2 宽_C605="498.70001220703125" 长_C604="708.4500122070312"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='122'">
                        <表:纸张_E7C2 宽_C605="773.9500122070312" 长_C604="1116.1500244140625"/>
                    </xsl:if>
                </xsl:if>

                <xsl:if test="@orientation = 'landscape'">
                    <xsl:if test="@paperSize='1'">
                        <表:纸张_E7C2 宽_C605="792.09997558593755" 长_C604="612.099975585937"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='3'">
                        <表:纸张_E7C2 宽_C605="1224.1500244140625" 长_C604="792.0999755859375"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='5'">
                        <表:纸张_E7C2 长_C604="1008.1500244140625" 宽_C605="612.0999755859375"/>
                    </xsl:if>
                    <xsl:if test ="@paperSize='7'">
                        <表:纸张_E7C2 长_C604="521.9000244140625" 宽_C605="756.0999755859375"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='8'">
                        <表:纸张_E7C2 宽_C605="1190.699951171875" 长_C604="842.0"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='9'">
                        <表:纸张_E7C2 宽_C605="842.0" 长_C604="595.3499755859375"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='11'">
                        <表:纸张_E7C2 长_C604="595.3499755859375" 宽_C605="419.6000061035156"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='12'">
                        <表:纸张_E7C2 宽_C605="1031.949951171875" 长_C604="728.5999755859375"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='13'">
                        <表:纸张_E7C2 宽_C605="728.5999755859375" 长_C604="515.9500122070312"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='17'">
                        <表:纸张_E7C2 宽_C605="1224.1500244140625" 长_C604="792.0999755859375"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='20'">
                        <表:纸张_E7C2 宽_C605="684.0999755859375" 长_C604="296.79998779296875"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='27'">
                        <表:纸张_E7C2 宽_C605="623.7000122070312" 长_C604="311.8500061035156"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='28'">
                        <表:纸张_E7C2 宽_C605="649.2000122070312" 长_C604="459.25"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='34'">
                        <表:纸张_E7C2 宽_C605="708.75" 长_C604="498.95001220703125"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='37'">
                        <表:纸张_E7C2 宽_C605="540.0499877929688" 长_C604="278.95001220703125"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='43'">
                        <表:纸张_E7C2 宽_C605="419.6000061035156" 长_C604="283.5"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='70'">
                        <表:纸张_E7C2 宽_C605="419.6000061035156" 长_C604="297.70001220703125"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='82'">
                        <表:纸张_E7C2 宽_C605="567.0" 长_C604="419.6000061035156"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='119'">
                        <表:纸张_E7C2 宽_C605="773.9500122070312" 长_C604="557.9500122070312"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='120'">
                        <表:纸张_E7C2 宽_C605="935.25" 长_C604="612.0999755859375"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='121'">
                        <表:纸张_E7C2 宽_C605="708.4500122070312" 长_C604="498.70001220703125"/>
                    </xsl:if>
                    <xsl:if test="@paperSize='122'">
                        <表:纸张_E7C2 宽_C605="1116.1500244140625" 长_C604="773.9500122070312"/>
                    </xsl:if>
                </xsl:if>
                
            </xsl:when>
            <xsl:otherwise>
                <xsl:if test="@orientation = 'portrait'">
                    <表:纸张_E7C2 宽_C605="612.0999755859375" 长_C604="792.0999755859375"/>
                </xsl:if>
                <xsl:if test="@orientation = 'landscape'">
                    <表:纸张_E7C2 宽_C605="792.0999755859375" 长_C604="612.0999755859375"/>
                </xsl:if>
            </xsl:otherwise>
        </xsl:choose>
        <xsl:if test="@orientation">
            <表:纸张方向_E7C3>
                <xsl:value-of select="@orientation"/>
            </表:纸张方向_E7C3>
        </xsl:if>
        <xsl:if test="@scale">
            <表:缩放_E7C4>
                <xsl:value-of select="@scale"/>
            </表:缩放_E7C4>
        </xsl:if>
        <!--<xsl:if test="not(@scale)">
      <表:缩放_E7C4>
        <xsl:value-of select="'100'"/>
      </表:缩放_E7C4>
		</xsl:if>-->
        <!--页边距-->
        <xsl:apply-templates select="../ws:pageMargins"/>
        <!--<xsl:apply-templates select="ws:worksheet/ws:pageMargins"/>-->
        <!--页眉页脚-->
        <xsl:if test="../ws:headerFooter">
            <xsl:apply-templates select="../ws:headerFooter"/>
        </xsl:if>

        <表:打印_E7CA>
            <xsl:choose>
                <xsl:when test="parent::ws:worksheet/ws:printOptions/@gridLines='1'">
                    <xsl:attribute name="是否带网格线_E7CB">
                        <xsl:value-of select="'true'"/>
                    </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:attribute name="是否带网格线_E7CB">
                        <xsl:value-of select="'false'"/>
                    </xsl:attribute>
                </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
                <xsl:when test="parent::ws:worksheet/ws:printOptions/@headings='1'">
                    <xsl:attribute name="是否带行号列标_E7CC">
                        <xsl:value-of select="'true'"/>
                    </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:attribute name="是否带行号列标_E7CC">
                        <xsl:value-of select="'false'"/>
                    </xsl:attribute>
                </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
                <xsl:when test="@draft='1'">
                    <xsl:attribute name="是否按草稿方式_E7CD">
                        <xsl:value-of select="'true'"/>
                    </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:attribute name="是否按草稿方式_E7CD">
                        <xsl:value-of select="'false'"/>
                    </xsl:attribute>
                </xsl:otherwise>
            </xsl:choose>

            <xsl:choose>
                <xsl:when test="@pageOrder='overThenDown'">
                    <xsl:attribute name="是否先列后行_E7CE">
                        <xsl:value-of select="'false'"/>
                    </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:attribute name="是否先列后行_E7CE">
                        <xsl:value-of select="'true'"/>
                    </xsl:attribute>
                </xsl:otherwise>
            </xsl:choose>
        </表:打印_E7CA>

        <xsl:if test="@cellComments">
            <表:批注打印方式_E7CF>
                <xsl:if test="@cellComments='asDisplayed'">
                    <xsl:value-of select="'in-place'"/>
                </xsl:if>
                <xsl:if test="@cellComments='atEnd'">
                    <xsl:value-of select="'sheet-end'"/>
                </xsl:if>
                <!--<xsl:value-of select="'none'"/>-->
            </表:批注打印方式_E7CF>
        </xsl:if>
        <!--<xsl:if test="@cellComments">
      <表:批注打印方式_E7CF>true</表:批注打印方式_E7CF>
		</xsl:if>-->
        <xsl:if test="not(@cellComments)">
            <表:批注打印方式_E7CF>none</表:批注打印方式_E7CF>
        </xsl:if>
        <xsl:if test="not(@errors)">
            <表:错误单元格打印方式_E7D0>none</表:错误单元格打印方式_E7D0>
        </xsl:if>
        <xsl:if test="@errors">
            <表:错误单元格打印方式_E7D0>
                <xsl:if test="@errors='blank'">
                    <xsl:value-of select="'blank'"/>
                </xsl:if>
                <xsl:if test="@errors='dash'">
                    <xsl:value-of select="'dash'"/>
                </xsl:if>
                <xsl:if test="@errors='NA'">
                    <xsl:value-of select="'na'"/>
                </xsl:if>
                <xsl:if test="@errors='display'">
                    <xsl:value-of select="'none'"/>
                </xsl:if>
            </表:错误单元格打印方式_E7D0>
        </xsl:if>

        <!--垂直对齐方式,水平对齐方式  李杨2011-11-18-->
      <!--2014-4-28, update by Qihy, 页面设置-页边距-对齐方式不正确，start-->
      <xsl:if test="parent::ws:worksheet[ws:printOptions]">
        <xsl:apply-templates select="parent::ws:worksheet/ws:printOptions"/>
      </xsl:if>
      <!--2014-4-28 end-->

        <xsl:if test="@fitToWidth or @fitToHeight">
            <表:调整_E7D1>
                <xsl:if test="@fitToWidth">
                    <xsl:attribute name="页宽倍数_E7D3">
                        <xsl:value-of select="@fitToWidth"/>
                    </xsl:attribute>
                </xsl:if>
                <xsl:if test="@fitToHeight">
                    <xsl:attribute name="页高倍数_E7D2">
                        <xsl:value-of select="@fitToHeight"/>
                    </xsl:attribute>
                </xsl:if>
            </表:调整_E7D1>
        </xsl:if>
        <xsl:if test="not(@fitToWidth) and not(@fitToHeight)">
            <表:调整_E7D1 页高倍数_E7D2="1" 页宽倍数_E7D3="1"/>
        </xsl:if>
    </xsl:template>
    <xsl:template match="ws:pageMargins">
        <表:页边距_E7C5>
            <xsl:if test="@left">
                <xsl:attribute name="左_C608">
                    <xsl:value-of select="round(@left*72)"/>
                </xsl:attribute>
            </xsl:if>
            <xsl:if test="@right">
                <xsl:attribute name="右_C60A">
                    <xsl:value-of select="round(number(@right)*72)"/>
                </xsl:attribute>
            </xsl:if>
            <xsl:if test="@top">
                <xsl:attribute name="上_C609">
                    <xsl:value-of select="round(number(@top)*72)"/>
                </xsl:attribute>
            </xsl:if>
            <xsl:if test="@bottom">
                <xsl:attribute name="下_C60B">
                    <xsl:value-of select="round(number(@bottom)*72)"/>
                </xsl:attribute>
            </xsl:if>
        </表:页边距_E7C5>
    </xsl:template>
    <!--添加页眉页脚的水平对齐，以及内容text。 李杨2011-3-12-->
    <xsl:template name ="fh">
        <xsl:param name ="v2"/>
        <xsl:choose >
            <xsl:when test="contains($v2,'T')">
                <字:域开始_419E 类型_416E="time" 是否锁定_416F="false"/>
                <字:域代码_419F>
                    <字:段落_416B>
                        <字:句_419D>
                            <字:句属性_4158/>
                            <字:文本串_415B>TIME</字:文本串_415B>
                        </字:句_419D>
                        <字:句_419D>
                            <字:句属性_4158/>
                            <字:空格符_4161 个数_4162="1"/>
                            <字:文本串_415B>\@ "h:mm AM/PM"</字:文本串_415B>
                        </字:句_419D>
                    </字:段落_416B>
                </字:域代码_419F>
                <字:域结束_41A0/>
                <字:句_419D>
                    <字:句属性_4158/>
                    <字:文本串_415B>Sheet1</字:文本串_415B>
                </字:句_419D>
            </xsl:when>
            <xsl:when test="contains($v2,'F')">
                <字:域开始_419E 类型_416E="filename" 是否锁定_416F="false"/>
                <字:域代码_419F>
                    <字:段落_416B>
                        <字:句_419D>
                            <字:句属性_4158/>
                            <字:文本串_415B>FILENAME</字:文本串_415B>
                        </字:句_419D>
                    </字:段落_416B>
                </字:域代码_419F>
                <字:句_419D>
                    <字:句属性_4158/>
                    <字:文本串_415B>新建 Microsoft Office Excel 工作表.eio</字:文本串_415B>
                </字:句_419D>
                <字:域结束_41A0/>
                <字:句_419D>
                    <字:句属性_4158/>
                    <字:制表符_415E/>
                </字:句_419D>
                <字:句_419D>
                    <字:句属性_4158/>
                </字:句_419D>
            </xsl:when>
            <xsl:when test="contains($v2,'A')">
                <字:域结束_41A0/>
                <字:句_419D>
                    <字:句属性_4158/>
                    <字:文本串_415B>sheetName</字:文本串_415B>
                </字:句_419D>
            </xsl:when>
            <xsl:when test="contains($v2,'N')">
                <字:域开始_419E 类型_416E="numpages" 是否锁定_416F="false"/>
                <字:域代码_419F>
                    <字:段落_416B>
                        <字:句_419D>
                            <字:句属性_4158/>
                            <字:文本串_415B>NumPages</字:文本串_415B>
                        </字:句_419D>
                    </字:段落_416B>
                </字:域代码_419F>
                <字:句_419D>
                    <字:句属性_4158/>
                    <字:文本串_415B>2</字:文本串_415B>
                </字:句_419D>
            </xsl:when>
            <xsl:when test="contains($v2,'P')">
                <字:域开始_419E 类型_416E="page" 是否锁定_416F="false"/>
                <字:域代码_419F>
                    <字:段落_416B>
                        <字:句_419D>
                            <字:句属性_4158/>
                            <字:文本串_415B>Page</字:文本串_415B>
                        </字:句_419D>
                    </字:段落_416B>
                </字:域代码_419F>
                <字:句_419D>
                    <字:句属性_4158/>
                    <字:文本串_415B>1</字:文本串_415B>
                </字:句_419D>
                <字:域结束_41A0/>
            </xsl:when>
            <xsl:when test="contains($v2,'D')">
                <字:域开始_419E 类型_416E="date" 是否锁定_416F="false"/>
                <字:域代码_419F>
                    <字:段落_416B>
                        <字:句_419D>
                            <字:句属性_4158/>
                            <字:文本串_415B>DATE</字:文本串_415B>
                        </字:句_419D>
                        <字:句_419D>
                            <字:句属性_4158/>
                            <字:空格符_4161 个数_4162="1"/>
                            <字:文本串_415B>\@ "yyyy/MM/dd"</字:文本串_415B>
                        </字:句_419D>
                    </字:段落_416B>
                </字:域代码_419F>
                <字:句_419D>
                    <字:句属性_4158/>
                    <字:文本串_415B>2010年4月23日</字:文本串_415B>
                </字:句_419D>
                <字:域结束_41A0/>
            </xsl:when>
            <xsl:when test ="contains($v2,'Z')">
                <字:域开始_419E 是否锁定_416F="false" 类型_416E="Filename"/>
                <字:域代码_419F>
                    <字:段落_416B>
                        <字:句_419D>
                            <字:句属性_4158/>
                            <字:文本串_415B>FileName</字:文本串_415B>
                            <字:空格符_4161/>
                            <字:文本串_415B>\p</字:文本串_415B>
                        </字:句_419D>
                    </字:段落_416B>
                </字:域代码_419F>
                <字:句_419D>
                    <字:句属性_4158/>
                    <字:文本串_415B>新建 Microsoft Office Excel 工作表.uos</字:文本串_415B>
                </字:句_419D>
                <字:域结束_41A0/>
            </xsl:when>

        </xsl:choose>
    </xsl:template>
    <xsl:template match="ws:headerFooter">
        <!--页眉-->
        <xsl:if test ="ws:oddHeader">
            <xsl:variable name ="v">
                <xsl:value-of select ="translate(ws:oddHeader,'&amp;','')"/>
            </xsl:variable>
            <!--左-->
            <xsl:if test ="contains(ws:oddHeader,'&amp;L')">
                <表:页眉页脚_E7C6>
                    <xsl:attribute name ="位置_E7C9">header-left</xsl:attribute>
                    <xsl:variable name ="text2">
                        <xsl:value-of select ="substring-after(ws:oddHeader,'&amp;L')"/>
                    </xsl:variable>
                    <xsl:variable name ="text">
                        <xsl:choose>
                            <xsl:when test ="contains($text2,'&amp;C')">
                                <xsl:value-of select ="substring-before($text2,'&amp;C')"/>
                            </xsl:when>
                            <xsl:when test ="not(contains($text2,'&amp;C')) and contains($text2,'&amp;R')">
                                <xsl:value-of select ="substring-before($text2,'&amp;R')"/>
                            </xsl:when>
                            <xsl:otherwise >
                                <xsl:value-of select ="$text2"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:variable>
                    <字:段落_416B>
                        <字:段落属性_419B>
                            <字:对齐_417D 水平对齐_421D="center"/>
                            <字:边框_4133>
                                <uof:下_C616 线型_C60D="none"/>
                            </字:边框_4133>
                            <字:制表位设置_418F>
                                <字:制表位_4171 位置_4172="10.0" 类型_4173="left"/>
                                <字:制表位_4171 位置_4172="207.65" 类型_4173="center"/>
                            </字:制表位设置_418F>
                        </字:段落属性_419B>
                        <!--<字:句_419D>
              <字:制表符_415E/>
              <字:文本串_415B>
                -->
                        <!--<xsl:value-of select="translate(ws:oddHeader,'&amp;ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz','')"/>-->
                        <!--
                <xsl:value-of select ="substring-before($text,'A')"/>
              </字:文本串_415B>
            </字:句_419D>-->
                        <xsl:variable name ="v1">
                            <xsl:value-of select ="substring-after($v,'L')"/>
                        </xsl:variable>
                        <xsl:call-template name ="fh">
                            <xsl:with-param name ="v2" select ="substring($v1,1,1)"/>
                        </xsl:call-template>

                        <字:句_419D>
                            <字:文本串_415B>
                                <xsl:value-of select ="translate($text,'&amp;','')"/>
                            </字:文本串_415B>
                        </字:句_419D>
                    </字:段落_416B>
                </表:页眉页脚_E7C6>
            </xsl:if>
            <!--中-->
            <!-- 20130319 update by xuzhenwei BUG_2722:集成OO-UOF页面设置中页眉页脚的部分内容丢失 start -->
            <xsl:if test ="contains(ws:oddHeader,'&amp;C') or contains(ws:oddHeader,'&amp;P')">
                <表:页眉页脚_E7C6>
                    <xsl:attribute name ="位置_E7C9">header-center</xsl:attribute>
                    <xsl:variable name ="text2">
                        <xsl:choose>
                            <xsl:when test="contains(ws:oddHeader,'&amp;P')">
                                <xsl:value-of select ="substring-after(ws:oddHeader,'&amp;P')"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select ="substring-after(ws:oddHeader,'&amp;C')"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:variable>
                    <xsl:variable name ="text">
                        <xsl:choose>
                            <xsl:when test ="contains($text2,'&amp;R')">
                                <xsl:value-of select ="substring-before($text2,'&amp;R')"/>
                            </xsl:when>
                            <xsl:otherwise >
                                <xsl:value-of select ="$text2"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:variable>
                    <字:段落_416B>
                        <字:段落属性_419B>
                            <字:对齐_417D 水平对齐_421D="center"/>
                            <字:边框_4133>
                                <uof:下_C616 线型_C60D="none"/>
                            </字:边框_4133>
                            <字:制表位设置_418F>
                                <字:制表位_4171 位置_4172="10.0" 类型_4173="left"/>
                                <字:制表位_4171 位置_4172="207.65" 类型_4173="center"/>
                            </字:制表位设置_418F>
                        </字:段落属性_419B>
                        <!--<字:句_419D>
              <字:制表符_415E/>
              <字:文本串_415B>
                -->
                        <!--<xsl:value-of select="translate(ws:oddHeader,'&amp;ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz','')"/>-->
                        <!--
                <xsl:value-of select ="substring-before($text,'A')"/>
              </字:文本串_415B>
            </字:句_419D>-->
                        <xsl:variable name ="v1">
                            <xsl:choose>
                                <xsl:when test="contains(ws:oddHeader,'&amp;P')">
                                    <xsl:value-of select ="substring-after($v,'P')"/>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:value-of select ="substring-after($v,'C')"/>
                                </xsl:otherwise>
                            </xsl:choose>
                            <!-- end -->
                        </xsl:variable>
                        <xsl:call-template name ="fh">
                            <xsl:with-param name ="v2" select ="substring($v1,1,1)"/>
                        </xsl:call-template>
                        <字:句_419D>
                            <字:文本串_415B>
                                <xsl:value-of select ="translate($text,'&amp;','')"/>
                            </字:文本串_415B>
                        </字:句_419D>
                    </字:段落_416B>
                </表:页眉页脚_E7C6>
            </xsl:if>
            <!--右-->
            <xsl:if test ="contains(ws:oddHeader,'&amp;R')">
                <表:页眉页脚_E7C6>
                    <xsl:attribute name ="位置_E7C9">header-right</xsl:attribute>
                    <xsl:variable name ="text2">
                        <xsl:value-of select ="substring-after(ws:oddHeader,'&amp;R')"/>
                    </xsl:variable>
                    <!--<xsl:variable name ="text">
            <xsl:choose>
              <xsl:when test ="contains($text2,'&amp;C')">
                <xsl:value-of select ="substring-before($text2,'&amp;C')"/>
              </xsl:when>
              <xsl:when test ="not(contains($text2,'&amp;C')) and contains($text2,'&amp;R')">
                <xsl:value-of select ="substring-before($text2,'&amp;R')"/>
              </xsl:when>
              <xsl:otherwise >
                <xsl:value-of select ="$text2"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:variable>-->
                    <字:段落_416B>
                        <字:段落属性_419B>
                            <字:对齐_417D 水平对齐_421D="center"/>
                            <字:边框_4133>
                                <uof:下_C616 线型_C60D="none"/>
                            </字:边框_4133>
                            <字:制表位设置_418F>
                                <字:制表位_4171 位置_4172="10.0" 类型_4173="left"/>
                                <字:制表位_4171 位置_4172="207.65" 类型_4173="center"/>
                            </字:制表位设置_418F>
                        </字:段落属性_419B>
                        <!--<字:句_419D>
              <字:制表符_415E/>
              <字:文本串_415B>
                -->
                        <!--<xsl:value-of select="translate(ws:oddHeader,'&amp;ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz','')"/>-->
                        <!--
                <xsl:value-of select ="substring-before($text,'A')"/>
              </字:文本串_415B>
            </字:句_419D>-->
                        <xsl:variable name ="v1">
                            <xsl:value-of select ="substring-after($v,'R')"/>
                        </xsl:variable>
                        <xsl:call-template name ="fh">
                            <xsl:with-param name ="v2" select ="substring($v1,1,1)"/>
                        </xsl:call-template>
                        <字:句_419D>
                            <字:文本串_415B>
                                <xsl:value-of select ="translate($text2,'&amp;','')"/>
                            </字:文本串_415B>
                        </字:句_419D>
                    </字:段落_416B>
                </表:页眉页脚_E7C6>
            </xsl:if>
        </xsl:if>

        <!--页脚-->
        <xsl:if test="ws:oddFooter">
            <xsl:variable name ="v">
                <xsl:value-of select ="translate(ws:oddFooter,'&amp;','')"/>
            </xsl:variable>
            <!--左-->
            <xsl:if test ="contains(ws:oddFooter,'&amp;L')">
                <表:页眉页脚_E7C6>
                    <xsl:attribute name ="位置_E7C9">footer-left</xsl:attribute>
                    <xsl:variable name ="text2">
                        <xsl:value-of select ="substring-after(ws:oddFooter,'&amp;L')"/>
                    </xsl:variable>
                    <xsl:variable name ="text">
                        <xsl:choose>
                            <xsl:when test ="contains($text2,'&amp;C')">
                                <xsl:value-of select ="substring-before($text2,'&amp;C')"/>
                            </xsl:when>
                            <xsl:when test ="not(contains($text2,'&amp;C')) and contains($text2,'&amp;R')">
                                <xsl:value-of select ="substring-before($text2,'&amp;R')"/>
                            </xsl:when>
                            <xsl:otherwise >
                                <xsl:value-of select ="$text2"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:variable>
                    <字:段落_416B>
                        <字:段落属性_419B>
                            <字:对齐_417D 水平对齐_421D="center"/>
                            <字:边框_4133>
                                <uof:下_C616 线型_C60D="none"/>
                            </字:边框_4133>
                            <字:制表位设置_418F>
                                <字:制表位_4171 位置_4172="10.0" 类型_4173="left"/>
                                <字:制表位_4171 位置_4172="207.65" 类型_4173="center"/>
                            </字:制表位设置_418F>
                        </字:段落属性_419B>
                        <!--<字:句_419D>
              <字:制表符_415E/>
              <字:文本串_415B>
                -->
                        <!--<xsl:value-of select="translate(ws:oddHeader,'&amp;ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz','')"/>-->
                        <!--
                <xsl:value-of select ="substring-before($text,'A')"/>
              </字:文本串_415B>
            </字:句_419D>-->
                        <xsl:variable name ="v1">
                            <xsl:value-of select ="substring-after($v,'L')"/>
                        </xsl:variable>
                        <xsl:call-template name ="fh">
                            <xsl:with-param name ="v2" select ="substring($v1,1,1)"/>
                        </xsl:call-template>

                        <字:句_419D>
                            <字:文本串_415B>
                                <xsl:value-of select ="translate($text,'&amp;','')"/>
                                <!--<xsl:value-of select ="substring-after($text,'A')"/>-->
                            </字:文本串_415B>
                        </字:句_419D>
                    </字:段落_416B>
                </表:页眉页脚_E7C6>
            </xsl:if>
            <!--中-->
            <xsl:if test ="contains(ws:oddFooter,'&amp;C')">
                <表:页眉页脚_E7C6>
                    <xsl:attribute name ="位置_E7C9">footer-center</xsl:attribute>
                    <xsl:variable name ="text2">
                        <xsl:value-of select ="substring-after(ws:oddFooter,'&amp;C')"/>
                    </xsl:variable>
                    <xsl:variable name ="text">
                        <xsl:choose>
                            <xsl:when test ="contains($text2,'&amp;R')">
                                <xsl:value-of select ="substring-before($text2,'&amp;R')"/>
                            </xsl:when>
                            <xsl:otherwise >
                                <xsl:value-of select ="$text2"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:variable>
                    <字:段落_416B>
                        <字:段落属性_419B>
                            <字:对齐_417D 水平对齐_421D="center"/>
                            <字:边框_4133>
                                <uof:下_C616 线型_C60D="none"/>
                            </字:边框_4133>
                            <字:制表位设置_418F>
                                <字:制表位_4171 位置_4172="10.0" 类型_4173="left"/>
                                <字:制表位_4171 位置_4172="207.65" 类型_4173="center"/>
                            </字:制表位设置_418F>
                        </字:段落属性_419B>
                        <!--<字:句_419D>
              <字:制表符_415E/>
              <字:文本串_415B>
                -->
                        <!--<xsl:value-of select="translate(ws:oddHeader,'&amp;ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz','')"/>-->
                        <!--
                <xsl:value-of select ="substring-before($text,'A')"/>
              </字:文本串_415B>
            </字:句_419D>-->
                        <xsl:variable name ="v1">
                            <xsl:value-of select ="substring-after($v,'C')"/>
                        </xsl:variable>
                        <xsl:call-template name ="fh">
                            <xsl:with-param name ="v2" select ="substring($v1,1,1)"/>
                        </xsl:call-template>
                        <字:句_419D>

                            <字:文本串_415B>

                                <xsl:value-of select ="translate($text,'&amp;','')"/>

                                <!--<xsl:value-of select ="substring-after($text,'A')"/>-->
                            </字:文本串_415B>
                        </字:句_419D>
                    </字:段落_416B>
                </表:页眉页脚_E7C6>
            </xsl:if>
            <!--右-->
            <xsl:if test ="contains(ws:oddFooter,'&amp;R')">
                <表:页眉页脚_E7C6>
                    <xsl:attribute name ="位置_E7C9">footer-right</xsl:attribute>
                    <xsl:variable name ="text2">
                        <xsl:value-of select ="substring-after(ws:oddFooter,'&amp;R')"/>
                    </xsl:variable>
                    <xsl:variable name ="text">
                        <xsl:value-of select ="concat($text2,'A')"/>
                    </xsl:variable>
                    <字:段落_416B>
                        <字:段落属性_419B>
                            <字:对齐_417D 水平对齐_421D="center"/>
                            <字:边框_4133>
                                <uof:下_C616 线型_C60D="none"/>
                            </字:边框_4133>
                            <字:制表位设置_418F>
                                <字:制表位_4171 位置_4172="10.0" 类型_4173="left"/>
                                <字:制表位_4171 位置_4172="207.65" 类型_4173="center"/>
                            </字:制表位设置_418F>
                        </字:段落属性_419B>
                        <!--<字:句_419D>
              <字:制表符_415E/>
              <字:文本串_415B>
                <xsl:value-of select ="substring-before($text,'A')"/>
              </字:文本串_415B>
            </字:句_419D>-->
                        <xsl:variable name ="v1">
                            <xsl:value-of select ="substring-after($v,'R')"/>
                        </xsl:variable>
                        <xsl:call-template name ="fh">
                            <xsl:with-param name ="v2" select ="substring($v1,1,1)"/>
                        </xsl:call-template>
                        <字:句_419D>
                            <字:文本串_415B>
                                <xsl:value-of select ="translate($text2,'&amp;','')"/>
                            </字:文本串_415B>
                        </字:句_419D>
                    </字:段落_416B>
                </表:页眉页脚_E7C6>
            </xsl:if>
            <!--<表:页眉页脚_E7C6 位置_E7C9="footer-left">
        <xsl:variable name ="text">
          <xsl:choose >
            <xsl:when test ="contains(ws:oddFooter,'&amp;L') or contains(ws:oddFooter,'&amp;C') or contains(ws:oddFooter,'&amp;R')">
              <xsl:variable name ="text2">
                <xsl:value-of select ="translate(ws:oddFooter,'&amp;ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz','')"/>
              </xsl:variable>
              <xsl:value-of select ="concat($text2,'A')"/>
            </xsl:when>
            <xsl:otherwise >
              <xsl:value-of select ="translate(ws:oddFooter,'&amp;ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz','A')"/>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
				<字:段落_416B>
					<字:段落属性_419B>
            <字:对齐_417D>
              <xsl:attribute name ="水平对齐_421D">
                <xsl:choose >
                  <xsl:when test ="not(contains(ws:oddFooter,'&amp;L'))and not(contains(ws:oddFooter,'&amp;R'))">
                    <xsl:value-of select ="'center'"/>
                  </xsl:when>
                  <xsl:when test ="contains(ws:oddFooter,'&amp;L') and not(contains(ws:oddFooter,'&amp;R')) and not(contains(ws:oddFooter,'&amp;C'))">
                    <xsl:value-of select ="'left'"/>
                  </xsl:when>
                  <xsl:when test ="not(contains(ws:oddFooter,'&amp;L')) and contains(ws:oddFooter,'&amp;R') and not(contains(ws:oddFooter,'&amp;C'))">
                    <xsl:value-of select ="'right'"/>
                  </xsl:when>
                  <xsl:otherwise >
                    <xsl:value-of select ="'center'"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </字:对齐_417D>
						<字:边框_4133>
							<uof:下_C616 线型_C60D="none"/>
						</字:边框_4133>
					</字:段落属性_419B>
					<字:句_419D>
						<字:文本串_415B>
							-->
            <!--<xsl:value-of select="translate(ws:oddFooter,'&amp;ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz','')"/>-->
            <!--
              <xsl:value-of select="substring-before($text,'A')"/>
						</字:文本串_415B>
					</字:句_419D>
					<xsl:if test="contains(ws:oddFooter,'&amp;T')">
						<字:域开始_419E 类型_416E="time" 是否锁定_416F="false"/>
						<字:域代码_419F>
							<字:段落_416B>
								<字:句_419D>
									<字:句属性_4158/>
									<字:文本串_415B>TIME</字:文本串_415B>
								</字:句_419D>
								<字:句_419D>
									<字:句属性_4158/>
                  <字:空格符_4161 个数_4162="1"/>
									<字:文本串_415B>\@ "h:mm AM/PM"</字:文本串_415B>
								</字:句_419D>
							</字:段落_416B>
						</字:域代码_419F>
						<字:域结束_41A0/>
						<字:句_419D>
							<字:句属性_4158/>
							<字:文本串_415B>Sheet1</字:文本串_415B>
						</字:句_419D>
					</xsl:if>
					<xsl:if test="contains(ws:oddFooter,'&amp;F')">
						<字:域开始_419E 类型_416E="filename" 是否锁定_416F="false"/>
						<字:域代码_419F>
							<字:段落_416B>
								<字:句_419D>
									<字:句属性_4158/>
									<字:文本串_415B>FILENAME</字:文本串_415B>
								</字:句_419D>
							</字:段落_416B>
						</字:域代码_419F>
						<字:句_419D>
							<字:句属性_4158/>
							<字:文本串_415B>新建 Microsoft Office Excel 工作表.eio</字:文本串_415B>
						</字:句_419D>
						<字:域结束_41A0/>
						<字:句_419D>
							<字:句属性_4158/>
							<字:制表符_415E/>
						</字:句_419D>
						<字:句_419D>
							<字:句属性_4158/>
						</字:句_419D>
					</xsl:if>
					<xsl:if test="contains(ws:oddFooter,'&amp;A')">
						<字:域结束_41A0/>
						<字:句_419D>
							<字:句属性_4158/>
							<字:文本串_415B>Sheet1</字:文本串_415B>
						</字:句_419D>
					</xsl:if>
					<xsl:if test="contains(ws:oddFooter,'&amp;N')">
						<字:域开始_419E 类型_416E="numpages" 是否锁定_416F="false"/>
						<字:域代码_419F>
							<字:段落_416B>
								<字:句_419D>
									<字:句属性_4158/>
									<字:文本串_415B>NumPages</字:文本串_415B>
								</字:句_419D>
							</字:段落_416B>
						</字:域代码_419F>
						<字:句_419D>
							<字:句属性_4158/>
							<字:文本串_415B>2</字:文本串_415B>
						</字:句_419D>
					</xsl:if>
					<xsl:if test="contains(ws:oddFooter,'&amp;P')">
						<字:域开始_419E 类型_416E="page" 是否锁定_416F="false"/>
						<字:域代码_419F>
							<字:段落_416B>
								<字:句_419D>
									<字:句属性_4158/>
									<字:文本串_415B>Page</字:文本串_415B>
								</字:句_419D>
							</字:段落_416B>
						</字:域代码_419F>
						<字:句_419D>
							<字:句属性_4158/>
							<字:文本串_415B>1</字:文本串_415B>
						</字:句_419D>
						<字:域结束_41A0/>
					</xsl:if>
					<xsl:if test="contains(ws:oddFooter,'&amp;D')">
						<字:域开始_419E 类型_416E="date" 是否锁定_416F="false"/>
            <字:域代码_419F>
              <字:段落_416B>
                <字:句_419D>
                  <字:句属性_4158/>
                  <字:文本串_415B>DATE</字:文本串_415B>
                </字:句_419D>
                <字:句_419D>
                  <字:句属性_4158/>
                  <字:空格符_4161 个数_4162="1"/>
                  <字:文本串_415B>\@ "yyyy/MM/dd"</字:文本串_415B>
                </字:句_419D>
              </字:段落_416B>
            </字:域代码_419F>
            <字:句_419D>
              <字:句属性_4158/>
              <字:文本串_415B>2010年4月23日</字:文本串_415B>
            </字:句_419D>
						<字:域结束_41A0/>
					</xsl:if>
          <字:句_419D>
            <字:文本串_415B>
              <xsl:value-of select ="substring-after($text,'A')"/>
            </字:文本串_415B>
          </字:句_419D>
				</字:段落_416B>
			</表:页眉页脚_E7C6>-->
        </xsl:if>
    </xsl:template>
    <xsl:template match="ws:printOptions">
        <!--<表:打印_E7CA>
      <xsl:if test="@gridLines='1'">
        <xsl:attribute name="是否带网格线_E7CB">
          <xsl:value-of select="'true'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="not(@gridLines)">
        <xsl:attribute name="是否带网格线_E7CB">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="@headings">
        <xsl:attribute name="是否带行号列标_E7CC">
          <xsl:value-of select="'true'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="not(@headings)">
        <xsl:attribute name="是否带行号列标_E7CC">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="following-sibling::ws:pageSetup[@draft]">
        <xsl:attribute name="是否按草稿方式_E7CD">
          <xsl:value-of select="following-sibling::ws:pageSetup/@draft"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./parent::worksheet/ws:pageSetup[@pageOrder='overThenDown']">
        <xsl:attribute name="是否先列后行_E7CE">
          <xsl:value-of select="'false'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="./parent::worksheet/ws:pageSetup[@pageOrder='downThenOver']">
        <xsl:attribute name="是否先列后行_E7CE">
          <xsl:value-of select="'true'"/>
        </xsl:attribute>
      </xsl:if>
    </表:打印_E7CA>-->
        <xsl:if test="./@verticalCentered = '1'">
            <表:垂直对齐方式_E701>center</表:垂直对齐方式_E701>
        </xsl:if>
        <xsl:if test="./@horizontalCentered = '1'">
            <表:水平对齐方式_E700>center</表:水平对齐方式_E700>
        </xsl:if>
    </xsl:template>

    <xsl:template match="ws:sheetViews">
        <表:视图_E7D5>
            <xsl:if test="ws:sheetView/@workbookViewId">
                <xsl:attribute name="窗口标识符_E7E5">
                    <xsl:value-of select="ws:sheetView/@workbookViewId"/>
                </xsl:attribute>
            </xsl:if>
            <xsl:if test="ws:sheetView/ws:selection">
                <xsl:variable name="xuzq1" select="substring-after(@sqref,':')"/>
                <xsl:variable name="xuzq11" select="translate($xuzq1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
                <xsl:variable name="xuzq12" select="translate($xuzq1,'0123456789','')"/>
                <xsl:variable name="xuzqzz2">
                    <xsl:choose>
                        <xsl:when test="$xuzq12='A'">
                            <xsl:value-of select="'1'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='B'">
                            <xsl:value-of select="'2'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='C'">
                            <xsl:value-of select="'3'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='D'">
                            <xsl:value-of select="'4'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='E'">
                            <xsl:value-of select="'5'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='F'">
                            <xsl:value-of select="'6'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='G'">
                            <xsl:value-of select="'7'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='H'">
                            <xsl:value-of select="'8'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='I'">
                            <xsl:value-of select="'9'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='J'">
                            <xsl:value-of select="'10'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='K'">
                            <xsl:value-of select="'11'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='L'">
                            <xsl:value-of select="'12'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='M'">
                            <xsl:value-of select="'13'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='N'">
                            <xsl:value-of select="'14'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='O'">
                            <xsl:value-of select="'15'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='P'">
                            <xsl:value-of select="'16'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='Q'">
                            <xsl:value-of select="'17'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='R'">
                            <xsl:value-of select="'18'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='S'">
                            <xsl:value-of select="'19'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='T'">
                            <xsl:value-of select="'20'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='U'">
                            <xsl:value-of select="'21'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='V'">
                            <xsl:value-of select="'22'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='W'">
                            <xsl:value-of select="'23'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='X'">
                            <xsl:value-of select="'24'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='Y'">
                            <xsl:value-of select="'25'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq12='Z'">
                            <xsl:value-of select="'26'"/>
                        </xsl:when>
                    </xsl:choose>
                </xsl:variable>
                <xsl:variable name="xuzq2" select="substring-before(@sqref,':')"/>
                <xsl:variable name="xuzq21" select="translate($xuzq2,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
                <xsl:variable name="xuzq22" select="translate($xuzq2,'0123456789','')"/>
                <xsl:variable name="xuzqzz1">
                    <xsl:choose>
                        <xsl:when test="$xuzq22='A'">
                            <xsl:value-of select="'1'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='B'">
                            <xsl:value-of select="'2'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='C'">
                            <xsl:value-of select="'3'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='D'">
                            <xsl:value-of select="'4'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='E'">
                            <xsl:value-of select="'5'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='F'">
                            <xsl:value-of select="'6'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='G'">
                            <xsl:value-of select="'7'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='H'">
                            <xsl:value-of select="'8'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='I'">
                            <xsl:value-of select="'9'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='J'">
                            <xsl:value-of select="'10'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='K'">
                            <xsl:value-of select="'11'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='L'">
                            <xsl:value-of select="'12'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='M'">
                            <xsl:value-of select="'13'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='N'">
                            <xsl:value-of select="'14'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='O'">
                            <xsl:value-of select="'15'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='P'">
                            <xsl:value-of select="'16'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='Q'">
                            <xsl:value-of select="'17'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='R'">
                            <xsl:value-of select="'18'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='S'">
                            <xsl:value-of select="'19'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='T'">
                            <xsl:value-of select="'20'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='U'">
                            <xsl:value-of select="'21'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='V'">
                            <xsl:value-of select="'22'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='W'">
                            <xsl:value-of select="'23'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='X'">
                            <xsl:value-of select="'24'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='Y'">
                            <xsl:value-of select="'25'"/>
                        </xsl:when>
                        <xsl:when test="$xuzq22='Z'">
                            <xsl:value-of select="'26'"/>
                        </xsl:when>
                    </xsl:choose>
                </xsl:variable>

                <!--
				<表:选中 uof:locID="s0036" uof:attrList="值" 表:值="true">
					<xsl:value-of select="concat('R',$xuzqzz1,'C',$xuzq21,':','R',$xuzqzz2,'C',$xuzq11)"/>
				</表:选中>
        -->
            </xsl:if>
            <!--是否选中标签 李杨2011.11.4-->
            <xsl:if test ="ws:sheetView/@tabSelected='1'">
                <表:是否选中_E7D6>true</表:是否选中_E7D6>
            </xsl:if>
            <xsl:if test="ws:sheetView/ws:pane">
                <xsl:if test="ws:sheetView/ws:pane[not(@state) or (@state != 'frozenSplit' and @state !='frozen')]">
                    <!--拆分中增加属性:最上行,最左列,选中窗口  李杨2012.4.24-->
                    <表:拆分_E7D7>
                        <xsl:variable name="h" select="ws:sheetView/ws:pane/@topLeftCell"/>
                        <xsl:variable name="hx2" select="substring($h,1,1)"/>
                        <xsl:variable name="hx1" select="substring-after($h,$hx2)"/>
                        <xsl:variable name="hx3">
                            <xsl:choose>
                                <xsl:when test="$hx2='A'">
                                    <xsl:value-of select="'1'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='B'">
                                    <xsl:value-of select="'2'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='C'">
                                    <xsl:value-of select="'3'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='D'">
                                    <xsl:value-of select="'4'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='E'">
                                    <xsl:value-of select="'5'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='F'">
                                    <xsl:value-of select="'6'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='G'">
                                    <xsl:value-of select="'7'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='H'">
                                    <xsl:value-of select="'8'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='I'">
                                    <xsl:value-of select="'9'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='J'">
                                    <xsl:value-of select="'10'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='K'">
                                    <xsl:value-of select="'11'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='L'">
                                    <xsl:value-of select="'12'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='M'">
                                    <xsl:value-of select="'13'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='N'">
                                    <xsl:value-of select="'14'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='O'">
                                    <xsl:value-of select="'15'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='P'">
                                    <xsl:value-of select="'16'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='Q'">
                                    <xsl:value-of select="'17'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='R'">
                                    <xsl:value-of select="'18'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='S'">
                                    <xsl:value-of select="'19'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='T'">
                                    <xsl:value-of select="'20'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='U'">
                                    <xsl:value-of select="'21'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='V'">
                                    <xsl:value-of select="'22'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='W'">
                                    <xsl:value-of select="'23'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='X'">
                                    <xsl:value-of select="'24'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='Y'">
                                    <xsl:value-of select="'25'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='Z'">
                                    <xsl:value-of select="'26'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='AA'">
                                    <xsl:value-of select="'27'"/>
                                </xsl:when>
                                <xsl:when test="$hx2='AB'">
                                    <xsl:value-of select="'28'"/>
                                </xsl:when>
                            </xsl:choose>
                        </xsl:variable>
                        <xsl:attribute name="宽_C605">
                            <xsl:value-of select ="ws:sheetView/ws:pane/@xSplit div 1152.5 * 42"/>
                            <!--<xsl:value-of select="22 + ($hx3 - 1)*40"/>-->
                        </xsl:attribute>
                        <xsl:attribute name="长_C604">
                            <xsl:value-of select ="ws:sheetView/ws:pane/@ySplit div 300 * 12"/>
                            <!--<xsl:value-of select="10 + ($hx1 - 1)*11"/>-->
                        </xsl:attribute>
                        <!--<xsl:attribute name ="选中窗口_E833">
              <xsl:value-of select ="@activePane"/>
            </xsl:attribute>-->
                        <xsl:attribute name ="最上行_E831">
                            <xsl:value-of select ="$hx1"/>
                        </xsl:attribute>
                        <xsl:attribute name ="最左列_E832">
                            <xsl:value-of select ="translate($hx2,'ABCDEFGHI','123456789')"/>
                        </xsl:attribute>
                        <xsl:attribute name ="选中窗口_E833">
                            <xsl:if test ="ws:sheetView/ws:pane/@activePane='topRight'">
                                <xsl:value-of select ="'top-right'"/>
                            </xsl:if>
                            <xsl:if test ="ws:sheetView/ws:pane/@activePane='topLeft'">
                                <xsl:value-of select ="'top-left'"/>
                            </xsl:if>
                            <xsl:if test ="ws:sheetView/ws:pane/@activePane='bottomRight'">
                                <xsl:value-of select ="'bottom-right'"/>
                            </xsl:if>
                            <xsl:if test ="ws:sheetView/ws:pane/@activePane='bottomLeft'">
                                <xsl:value-of select ="'bottom-left'"/>
                            </xsl:if>
                        </xsl:attribute>

                    </表:拆分_E7D7>
                </xsl:if>
                <xsl:if test="ws:sheetView/ws:pane[(@state = 'frozenSplit' or @state='frozen') and not(@xSplit) and @ySplit]">
                    <!--冻结中需增加属性:最上行,最左列  李杨2011.11.4-->
                    <!-- xuzhenwei 20130107 修改永中2012冻结属性（行冻结和列冻结）失效，start -->
                    <表:冻结_E7D8>
                        <xsl:attribute name="行号_E7D9">
                          <xsl:choose>

                            <!-- 2014-3-26 update by Qihy, start-->
                            <!--<xsl:when test ="translate(ws:sheetView/ws:pane/@topLeftCell,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')!='1'">
                                  <xsl:value-of select ="translate(ws:sheetView/ws:pane/@topLeftCell,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','') - 1"/>
                                </xsl:when>-->
                            <!-- 2014-3-26 end-->

                            <xsl:when test ="ws:sheetView[@topLeftCell] and translate(ws:sheetView/@topLeftCell,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')!='1'">
                              <xsl:value-of select ="ws:sheetView/ws:pane/@ySplit + translate(ws:sheetView/@topLeftCell,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','') - 1"/>
                            </xsl:when>
                            <xsl:otherwise>
                              <xsl:value-of select="ws:sheetView/ws:pane/@ySplit"/>
                            </xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                      <!-- 2014-3-23 add by Qihy BUG_3149:oo-uof-oo 3149 (Transition)OOX->UOF->OOX 第二轮转换失败， start -->
                      <xsl:attribute name="最上行_E83D">
                        <xsl:value-of select ="translate(ws:sheetView/ws:pane/@topLeftCell,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
                      </xsl:attribute>
                      <xsl:attribute name="最左列_E83E">
                        <xsl:value-of select="1"/>
                      </xsl:attribute>
                      <!--2014-3-23 end -->
                      
                    </表:冻结_E7D8>
                </xsl:if>
                <xsl:if test="ws:sheetView/ws:pane[(@state = 'frozenSplit' or @state='frozen') and  @xSplit and not(@ySplit)]">
                    <表:冻结_E7D8>
                        <xsl:attribute name="列号_E7DA">
                          <xsl:choose>

                            <!-- 2014-3-26 update by Qihy, start-->
                            <!--<xsl:when test ="translate(ws:sheetView/ws:pane/@topLeftCell,'0123456789','')!='A'">
                                  <xsl:variable name="BCell" select ="translate(ws:sheetView/ws:pane/@topLeftCell,'0123456789','')"/>
                                  <xsl:choose>
                                    <xsl:when test="$BCell='B'">
                                      <xsl:value-of select="'1'"/>
                                    </xsl:when>
                                    <xsl:otherwise>
                                      <xsl:value-of select ="translate(ws:sheetView/ws:pane/@topLeftCell,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','') - 1"/>
                                    </xsl:otherwise>
                                  </xsl:choose>
                                </xsl:when>-->
                            <xsl:when test ="ws:sheetView[@topLeftCell] and translate(ws:sheetView/@topLeftCell,'0123456789','')!='A'">
                              <xsl:variable name ="colSeq">
                              <xsl:call-template name="Ts2Dec">
                                <xsl:with-param name="tsSource" select ="translate(ws:sheetView/@topLeftCell,'0123456789','')"/>
                              </xsl:call-template>
                              </xsl:variable>
                              <xsl:value-of select="ws:sheetView/ws:pane/@xSplit + $colSeq - 1"/>"
                            </xsl:when>
                            <!-- 2014-3-26 end-->
                            
                            <xsl:otherwise>
                              <xsl:value-of select="ws:sheetView/ws:pane/@xSplit"/>
                            </xsl:otherwise>
                          </xsl:choose>
                          
                        </xsl:attribute>

                     
                      <!-- 2014-3-23 add by Qihy BUG_3149:oo-uof-oo 3149 (Transition)OOX->UOF->OOX 第二轮转换失败， start -->
                      <xsl:attribute name="最上行_E83D">
                        <xsl:value-of select="1"/>
                      </xsl:attribute>
                      <xsl:attribute name="最左列_E83E">
                        <xsl:variable name="leftCol" select ="translate(ws:sheetView/ws:pane/@topLeftCell,'0123456789','')"/>
                        <xsl:call-template name="Ts2Dec">
                          <xsl:with-param name="tsSource" select="$leftCol"/>
                        </xsl:call-template>
                      </xsl:attribute>
                      <!--2014-3-23 end -->
                      
                    </表:冻结_E7D8>
                    <!-- xuzhenwei 20130107 修改永中2012冻结属性（行冻结和列冻结）失效，end -->
                </xsl:if>
                <xsl:if test="ws:sheetView/ws:pane[(@state = 'frozenSplit' or @state='frozen') and  @xSplit and @ySplit]">
                    <表:冻结_E7D8>
                        <xsl:attribute name="行号_E7D9">
                            <xsl:value-of select="ws:sheetView/ws:pane/@ySplit"/>
                        </xsl:attribute>
                        <xsl:attribute name="列号_E7DA">
                            <xsl:value-of select="ws:sheetView/ws:pane/@xSplit"/>
                        </xsl:attribute>

                      <!-- 2014-3-26 update by Qihy, start-->
                        <!-- 20130318 add by xuzhenwei BUG_2759:oo-uof 表格多行丢失 start -->
                        <!--<xsl:attribute name="最上行_E83D">
                            <xsl:value-of select="ws:sheetView/ws:pane/@ySplit + 1"/>
                        </xsl:attribute>
                        <xsl:attribute name="最左列_E83E">
                            <xsl:value-of select="ws:sheetView/ws:pane/@xSplit + 1"/>
                        </xsl:attribute>-->
                        <!-- end -->
                      <xsl:attribute name="最上行_E83D">
                        <xsl:value-of select ="translate(ws:sheetView/ws:pane/@topLeftCell,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
                      </xsl:attribute>
                      <xsl:attribute name="最左列_E83E">
                        <xsl:variable name="leftCol" select ="translate(ws:sheetView/ws:pane/@topLeftCell,'0123456789','')"/>
                        <xsl:call-template name="Ts2Dec">
                          <xsl:with-param name="tsSource" select="$leftCol"/>
                        </xsl:call-template>
                      </xsl:attribute>
                      <!-- 2014-3-26 end-->
                      
                    </表:冻结_E7D8>
                </xsl:if>
            </xsl:if>
          
          <!--2014-3-26 update by Qihy, 表:最上行_E7DB表:最左列_E7DC和获取值不正确， start-->
            <!--<xsl:if test ="ws:sheetView/ws:pane[@topLeftCell and (@state = 'frozen')]">
                <xsl:variable name="coll">
                    <xsl:value-of select="ws:sheetView/ws:pane/@topLeftCell"/>
                </xsl:variable>
                <xsl:variable name="col" select="translate($coll,'0123456789','')"/>
                <xsl:variable name="row" select="translate($coll,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
                --><!-- 20130318 add by xuzhenwei BUG_2759  BUG_2759:oo-uof 表格多行丢失 start --><!--
                <表:最上行_E7DB>
                    <xsl:value-of select="'1'"/>
                    --><!--最上行应该默认从1行开始 <xsl:value-of select ="$row - 1"/>--><!--
                </表:最上行_E7DB>
                <表:最左列_E7DC>
                    <xsl:value-of select ="'1'"/>
                    --><!--最左列应该默认从第1列开始 <xsl:value-of select ="$col"/>--><!--
                </表:最左列_E7DC>
                --><!-- end --><!--
            </xsl:if>
            <xsl:if test="ws:sheetView/ws:pane[@topLeftCell and not(@state = 'frozenSplit') and not(@state = 'frozen')]">
                --><!--<xsl:if test="ws:sheetView/ws:pane/@topLeftCell">--><!--
                <xsl:variable name="coll">
                    <xsl:value-of select="ws:sheetView/ws:pane/@topLeftCell"/>
                </xsl:variable>
                <xsl:variable name="col" select="translate($coll,'0123456789','')"/>
                <xsl:variable name="row" select="translate($coll,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
                <表:最上行_E7DB>
                    <xsl:value-of select="$row"/>
                </表:最上行_E7DB>
                <表:最左列_E7DC>
                    <xsl:call-template name="Ts2Dec">
                        <xsl:with-param name="tsSource" select="$col"/>
                    </xsl:call-template>
                </表:最左列_E7DC>
            </xsl:if>-->
          <xsl:if test ="ws:sheetView[@topLeftCell]">
            <表:最上行_E7DB>
              <xsl:value-of select ="translate(ws:sheetView/@topLeftCell,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
            </表:最上行_E7DB>
            <表:最左列_E7DC>
              <xsl:variable name ="colSeq" select="translate(ws:sheetView/@topLeftCell,'0123456789','')"/>
              <xsl:call-template name="Ts2Dec">
                <xsl:with-param name="tsSource" select ="$colSeq"/>
              </xsl:call-template>
            </表:最左列_E7DC>
          </xsl:if>
          <xsl:if test ="ws:sheetView[not(@topLeftCell)]">
            <表:最上行_E7DB>
              <xsl:value-of select ="1"/>
            </表:最上行_E7DB>
            <表:最左列_E7DC>
              <xsl:value-of select ="1"/>
            </表:最左列_E7DC>
          </xsl:if>
          <!--2014-3-26 end-->
          
            <!--yx,add,element seqence,2010.4.21-->
            <!--
      <表:最上行 uof:locID="s0039">1</表:最上行>
      <表:最左列 uof:locID="s0040">1</表:最左列>-->
            <xsl:if test="ws:sheetView[@view]">
                <表:当前视图类型_E7DD>
                    <xsl:variable name="sty">
                        <xsl:value-of select="ws:sheetView/@view"/>
                    </xsl:variable>
                    <xsl:if test="$sty = 'pageBreakPreview'">
                        <xsl:value-of select="'page'"/>
                    </xsl:if>
                    <xsl:if test="$sty != 'pageBreakPreview'">
                        <xsl:value-of select="'normal'"/>
                    </xsl:if>
                </表:当前视图类型_E7DD>
            </xsl:if>
            <xsl:if test="ws:sheetView[@showFormulas = '1']">
                <表:是否显示公式_E7DE>true</表:是否显示公式_E7DE>
            </xsl:if>
            <xsl:if test="ws:sheetView[not(@showFormulas) or @showFormulas = '0']">
                <表:是否显示公式_E7DE>false</表:是否显示公式_E7DE>
            </xsl:if>
            <!--<xsl:if test="ws:sheetView[not(@showGridLines) or @showGridLines = '1']">
        <表:是否显示网格_E7DF>true</表:是否显示网格_E7DF>
      </xsl:if>
      <xsl:if test="ws:sheetView[@showGridLines = '0']">
        <表:是否显示网格_E7DF>false</表:是否显示网格_E7DF>
      </xsl:if>-->
            <!--修改“表:是否显示网格_E7DF” 李杨2012-5-7-->
            <xsl:choose >
                <xsl:when test ="ws:sheetView[not(@showGridLines) or @showGridLines = '1']">
                    <表:是否显示网格_E7DF>true</表:是否显示网格_E7DF>
                </xsl:when>
                <xsl:otherwise >
                    <表:是否显示网格_E7DF>false</表:是否显示网格_E7DF>
                </xsl:otherwise>
            </xsl:choose>
            <xsl:if test ="ws:sheetView[@defaultGridColor='0']">
                <表:网格颜色_E7E0>
                    
                    <!-- update by 凌峰 BUG_3023:网格线颜色不是红色  20140306 start -->
                    <xsl:choose>
                        <xsl:when test="ws:sheetView[@colorId='0' or @colorId='8']">
                            <xsl:value-of select="'#000000'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='1' or @colorId='9']">
                            <xsl:value-of select="'#ffffff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='2' or @colorId='10']">
                            <xsl:value-of select="'#ff0000'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='3' or @colorId='11']">
                            <xsl:value-of select="'#00ff00'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='4' or @colorId='12']">
                            <xsl:value-of select="'#0000ff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='5' or @colorId='13']">
                            <xsl:value-of select="'#ffff00'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='6' or @colorId='14']">
                            <xsl:value-of select="'#ff00ff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='7' or @colorId='15']">
                            <xsl:value-of select="'#00ffff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='16']">
                            <xsl:value-of select="'#800000'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='17']">
                            <xsl:value-of select="'#008000'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='18']">
                            <xsl:value-of select="'#000080'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='19']">
                            <xsl:value-of select="'#808000'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='20']">
                            <xsl:value-of select="'#800080'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='21']">
                            <xsl:value-of select="'#008080'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='22']">
                            <xsl:value-of select="'#c0c0c0'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='23']">
                            <xsl:value-of select="'#00808080'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='24']">
                            <xsl:value-of select="'#999ff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='25']">
                            <xsl:value-of select="'#993366'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='26']">
                            <xsl:value-of select="'#ffffcc'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='27']">
                            <xsl:value-of select="'#ccffff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='28']">
                            <xsl:value-of select="'#660066'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='29']">
                            <xsl:value-of select="'#ff8080'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='30']">
                            <xsl:value-of select="'#0066cc'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='31']">
                            <xsl:value-of select="'#ccccffFF'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='32']">
                            <xsl:value-of select="'#000080'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='33']">
                            <xsl:value-of select="'#ff00ff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='34']">
                            <xsl:value-of select="'#ffff00'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='35']">
                            <xsl:value-of select="'#00ffff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='36']">
                            <xsl:value-of select="'#800080'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='37']">
                            <xsl:value-of select="'#800000'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='38']">
                            <xsl:value-of select="'#008080'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='39']">
                            <xsl:value-of select="'#0000ff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='40']">
                            <xsl:value-of select="'#00ccff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='41']">
                            <xsl:value-of select="'#ccffff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='42']">
                            <xsl:value-of select="'#ccffcc'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='43']">
                            <xsl:value-of select="'#ffff99'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='44']">
                            <xsl:value-of select="'#99ccff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='45']">
                            <xsl:value-of select="'#ff99cc'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='46']">
                            <xsl:value-of select="'#cc99ff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='47']">
                            <xsl:value-of select="'#ffcc99'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='48']">
                            <xsl:value-of select="'#3366ff'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='49']">
                            <xsl:value-of select="'#33cccc'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='50']">
                            <xsl:value-of select="'#99cc00'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='51']">
                            <xsl:value-of select="'#ffcc00'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='52']">
                            <xsl:value-of select="'#ff9900'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='53']">
                            <xsl:value-of select="'#ff6600'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='54']">
                            <xsl:value-of select="'#666699'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='55']">
                            <xsl:value-of select="'#969696'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='56']">
                            <xsl:value-of select="'#003366'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='57']">
                            <xsl:value-of select="'#339966'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='58']">
                            <xsl:value-of select="'#003300'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='59']">
                            <xsl:value-of select="'#333300'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='60']">
                            <xsl:value-of select="'#993300'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='61']">
                            <xsl:value-of select="'#993366'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='62']">
                            <xsl:value-of select="'#333399'"/>
                        </xsl:when>
                        <xsl:when test="ws:sheetView[@colorId='63']">
                            <xsl:value-of select="'#333333'"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:value-of select="auto"/>
                        </xsl:otherwise>
                    </xsl:choose>
                    <!--end-->
                    
                </表:网格颜色_E7E0>
            </xsl:if>
            <xsl:if test="ws:sheetView[@zoomScale]">
                <表:缩放_E7C4>
                    <xsl:variable name="scl">
                        <xsl:value-of select="ws:sheetView/@zoomScale"/>
                    </xsl:variable>
                    <xsl:value-of select="$scl"/>
                </表:缩放_E7C4>
            </xsl:if>
            <xsl:if test="ws:sheetView[not(@zoomScale)]">
                <表:缩放_E7C4>100</表:缩放_E7C4>
            </xsl:if>
            <xsl:if test="ws:sheetView[@zoomScaleSheetLayoutView]">
                <表:分页缩放_E7E1>
                    <xsl:variable name="ss" select="ws:sheetView/@zoomScaleSheetLayoutView"/>
                    <xsl:value-of select="$ss"/>
                </表:分页缩放_E7E1>
            </xsl:if>
            <xsl:if test="ws:sheetView[not(@zoomScaleSheetLayoutView)]">
                <表:分页缩放_E7E1>60</表:分页缩放_E7E1>
            </xsl:if>
            <xsl:if test="ws:sheetView/ws:selection[@sqref]">
                <表:选中区域_E7E2>
                    <xsl:variable name="cell" select="ws:sheetView/ws:selection/@sqref"/>
                    <xsl:choose>
                        <xsl:when test="contains($cell,':')">
                            <xsl:variable name="b" select="substring-before($cell,':')"/>
                            <xsl:variable name="a" select="substring-after($cell,':')"/>
                            <xsl:variable name="bc" select="translate($b,'0123456789','')"/>
                            <xsl:variable name="bn" select="translate($b,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
                            <xsl:variable name="bcc" select="concat('$',$bc)"/>
                            <xsl:variable name="bnc" select="concat('$',$bn)"/>
                            <xsl:variable name="ac" select="translate($a,'0123456789','')"/>
                            <xsl:variable name="an" select="translate($a,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','')"/>
                            <xsl:variable name="acc" select="concat('$',$ac)"/>
                            <xsl:variable name="anc" select="concat('$',$an)"/>
                            <xsl:variable name="bbb" select="concat($bcc,$bnc)"/>
                            <xsl:variable name="aaa" select="concat($acc,$anc)"/>
                            <xsl:if test="$bbb != $aaa">
                                <xsl:value-of select="concat($bbb,':',$aaa)"/>
                            </xsl:if>
                            <xsl:if test="$bbb = $aaa">
                                <xsl:value-of select="$bbb"/>
                            </xsl:if>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:variable name="f" select="substring($cell,1,1)"/>
                            <xsl:variable name="s" select="substring($cell,2)"/>
                            <xsl:variable name="ff" select="concat('$',$f)"/>
                            <xsl:variable name="ss" select="concat('$',$s)"/>
                            <xsl:value-of select="concat($ff,$ss)"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </表:选中区域_E7E2>
            </xsl:if>

            <xsl:if test="ws:sheetView/@showRowColHeaders='0'">
                <表:是否显示行号列标_E7E3>false</表:是否显示行号列标_E7E3>
            </xsl:if>

            <xsl:if test="ws:sheetView/@showZeros='0'">
                <表:是否显示零值_E7E4>false</表:是否显示零值_E7E4>
            </xsl:if>
        </表:视图_E7D5>
    </xsl:template>
    <!--列号的计算-->
    <xsl:template name="Ts2Dec">
        <xsl:param name="tsSource"/>
        <xsl:variable name="z" select="'000000'"/>
        <xsl:variable name="RowID" select="concat(substring($z,1,(string-length($z)-string-length($tsSource))),$tsSource)"/>
        <xsl:value-of select="$tsSource"/>
        <xsl:text>======</xsl:text>
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
</xsl:stylesheet>
