<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:ws="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                xmlns:式样="http://schemas.uof.org/cn/2009/styles"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:pc="http://schemas.openxmlformats.org/package/2006/content-types"
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:ori="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image">
    <xsl:output method ="xml" version ="1.1" encoding ="UTF-8" indent="yes"/>
    <xsl:template name="style">
        <式样:式样集_990B>
            <式样:字体集_990C>
                <!--缺省字体-->
                <!--Modified by LDM in 2011/04/21-->
                <!--字体声明中的字体族，UOF2.0中成为元素，并且增加了替换字体族元素  李杨2011-11-9-->
                <式样:字体声明_990D 标识符_9902="font_00000" 名称_9903="Times New Roman" 替换字体_9904="Times New Roman">
                    <式样:字体族_9900>auto</式样:字体族_9900>
                </式样:字体声明_990D>
                <式样:字体声明_990D 标识符_9902="font_00001" 名称_9903="宋体" 替换字体_9904="永中宋体">
                    <式样:字体族_9900>auto</式样:字体族_9900>
                </式样:字体声明_990D>
                <!--字体声明-->
                <!--Modified by LDM in 2011/04/21-->
                <!--此模板调用有问题  李杨2011-11-14-->
                <xsl:apply-templates select="/ws:spreadsheets/ws:styleSheet/ws:fonts"/>
              
              <!--2014-6-9, add by Qihy, sharedString.xml中定义的字体转换不正确， start-->
              <xsl:for-each select="ws:spreadsheets/ws:sst/ws:si">
                <xsl:call-template name="siRFont">
                  <xsl:with-param name="pos" select="position()"/>
                </xsl:call-template>
              </xsl:for-each>
              <!--2014-6-9 end--> 

                <!--20121211,gaoyuwei，解决OOXML到UOF文本框中艺术字大小不正确BUG（实际是字体丢失） start-->
                <xsl:if test="//xdr:wsDr/xdr:oneCellAnchor/xdr:sp/xdr:txBody">

                    <xsl:for-each select="//xdr:wsDr/xdr:oneCellAnchor">

                        <!--当不存在字体定义，OOXML 默认calibri-->
                        <xsl:if test="not(./xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:latin or a:ea  or a:cs)">
                            <式样:字体声明_990D 标识符_9902="font_00002" 名称_9903="Calibri" 替换字体_9904="Calibri">
                                <式样:字体族_9900>auto</式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>

                        <!--当txbox中存在字体定义，对应转换字体-->
                        <xsl:if test="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:latin or xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:ea  or xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:cs">
                            <xsl:variable name="sp_id">
                                <xsl:if test="ancestor::xdr:twoCellAnchor[xdr:sp]">
                                    <xsl:value-of select="ancestor::xdr:sp/xdr:nvSpPr/xdr:cNvPr/@id"/>
                                </xsl:if>
                            </xsl:variable>
                            <xsl:variable name="node_pos">
                                <xsl:value-of select="position()"/>
                            </xsl:variable>
                            <xsl:variable name="zitiyy">
                                <xsl:value-of select="concat($sp_id,$node_pos)"/>
                            </xsl:variable>
                            <!--中文字体存在两种情况a:ea和a:cs，此处为第一种-->
                            <xsl:if test="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:cs">
                                <式样:字体声明_990D>
                                    <xsl:attribute name="标识符_9902">
                                        <xsl:value-of select="concat($zitiyy,'ea')"/>
                                    </xsl:attribute>
                                    <xsl:choose>
                                      <xsl:when test="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:cs[@typeface!='' and @typeface!='+mn-lt' and @typeface!='+mj-lt' and @typeface!='+mj-ea' and @typeface!='+mn-ea' and @typeface!='+mj-cs' and @typeface!='+mn-cs']">
                                        <xsl:attribute name="名称_9903">
                                          <xsl:value-of select="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:ea/@typeface"/>
                                        </xsl:attribute>
                                        <xsl:attribute name="替换字体_9904">
                                          <xsl:value-of select="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:ea/@typeface"/>
                                        </xsl:attribute>
                                      </xsl:when>
                                      <xsl:otherwise>
                                        <xsl:attribute name="名称_9903">
                                          <xsl:value-of select="'宋体'"/>
                                        </xsl:attribute>
                                        <xsl:attribute name="替换字体_9904">
                                          <xsl:value-of select="'宋体'"/>
                                        </xsl:attribute>
                                      </xsl:otherwise>
                                    </xsl:choose>
                                    <式样:字体族_9900>
                                        <xsl:value-of select="'auto'"/>
                                    </式样:字体族_9900>
                                </式样:字体声明_990D>
                            </xsl:if>
                            <!--中文字体存在两种情况a:ea和a:cs,此处为第二种-->
                            <xsl:if test="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:ea">
                                <式样:字体声明_990D>
                                    <xsl:attribute name="标识符_9902">
                                        <xsl:value-of select="concat($zitiyy,'ea')"/>
                                    </xsl:attribute>
                                  
                                    <!--<xsl:attribute name="名称_9903">
                                        <xsl:value-of select="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:ea/@typeface"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="替换字体_9904">
                                        <xsl:value-of select="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:ea/@typeface"/>
                                    </xsl:attribute>-->
                                  <xsl:choose>
                                    <xsl:when test="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:ea[@typeface!='' and @typeface!='+mn-lt' and @typeface!='+mj-lt' and @typeface!='+mj-ea' and @typeface!='+mn-ea' and @typeface!='+mj-cs' and @typeface!='+mn-cs']">
                                      <xsl:attribute name="名称_9903">
                                        <xsl:value-of select="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:ea/@typeface"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="替换字体_9904">
                                        <xsl:value-of select="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:ea/@typeface"/>
                                      </xsl:attribute>
                                    </xsl:when>
                                    <xsl:otherwise>
                                      <xsl:attribute name="名称_9903">
                                        <xsl:value-of select="'宋体'"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="替换字体_9904">
                                        <xsl:value-of select="'宋体'"/>
                                      </xsl:attribute>
                                    </xsl:otherwise>
                                  </xsl:choose>
                                    <式样:字体族_9900>
                                        <xsl:value-of select="'auto'"/>
                                    </式样:字体族_9900>
                                </式样:字体声明_990D>
                            </xsl:if>
                            <xsl:if test="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:latin">
                                <式样:字体声明_990D>
                                    <xsl:attribute name="标识符_9902">
                                        <xsl:value-of select="concat($zitiyy,'latin')"/>
                                    </xsl:attribute>
                                  
                                  <!--2014-3-27, update by Qihy, BUG3159字体转换不正确， start-->
                                    <!--<xsl:attribute name="名称_9903">
                                        <xsl:value-of select="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:ea/@typeface"/>
                                    </xsl:attribute>
                                    <xsl:attribute name="替换字体_9904">
                                        <xsl:value-of select="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:ea/@typeface"/>
                                    </xsl:attribute>-->
                                  <xsl:choose>
                                    <xsl:when test="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:latin[@typeface!='' and @typeface!='+mn-lt' and @typeface!='+mj-lt' and @typeface!='+mj-ea' and @typeface!='+mn-ea' and @typeface!='+mj-cs' and @typeface!='+mn-cs']">
                                      <xsl:attribute name="名称_9903">
                                        <xsl:value-of select="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:latin/@typeface"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="替换字体_9904">
                                        <xsl:value-of select="xdr:sp/xdr:txBody/a:p/a:r/a:rPr/a:latin/@typeface"/>
                                      </xsl:attribute>
                                    </xsl:when>
                                    <xsl:otherwise>
                                      <xsl:attribute name="名称_9903">
                                        <xsl:value-of select="'Calibri'"/>
                                      </xsl:attribute>
                                      <xsl:attribute name="替换字体_9904">
                                        <xsl:value-of select="'Calibri'"/>
                                      </xsl:attribute>
                                    </xsl:otherwise>
                                  </xsl:choose>
                                  <!--2014-3-27 end-->
                                  
                                    <式样:字体族_9900>
                                        <xsl:value-of select="'auto'"/>
                                    </式样:字体族_9900>
                                </式样:字体声明_990D>
                            </xsl:if>
                        </xsl:if>
                    </xsl:for-each>

                </xsl:if>

                <!--end-->


                <!--图表的字体集-->
                <xsl:if test="/ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/pr:Relationships/pr:Relationship/@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'">
                    <xsl:for-each select="/ws:spreadsheets/ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace">
                        <xsl:variable name="fileName" select="translate(./@filename,'.xml','')"/>
                        <xsl:if test="./c:txPr/a:p/a:pPr/a:defRPr/a:ea">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('chartSpace_ea_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:txPr/a:p/a:pPrni /a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:txPr/a:p/a:pPr/a:defRPr/a:latin">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('chartSpace_la_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>

                        <!--分类轴字体-->
                        <xsl:if test="./c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('catAx_ea_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:ea">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('catAxTitle_ea_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:latin">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('catAxTitle_la_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('catAx_la_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>

                        <!--数值轴字体-->
                      
                      <!--2014-4-23, update by Qihy, 修改标识符标识符_9902取值，并增加数值轴标题字体的式样 start-->
                        <xsl:if test="./c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('valAx_ea_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:ea">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat($fileName,'.btj_szz_e_font')"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:latin">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat($fileName,'.btj_szz_l_font')"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                      
                      <xsl:if test="./c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin">
                        <式样:字体声明_990D>
                          <xsl:attribute name="标识符_9902">
                            <xsl:value-of select="concat($fileName,'.btj_szz_l_font')"/>
                          </xsl:attribute>
                          <xsl:attribute name="名称_9903">
                            <xsl:value-of select="./c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                          </xsl:attribute>
                          <xsl:attribute name="替换字体_9904">
                            <xsl:value-of select="./c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                          </xsl:attribute>
                          <式样:字体族_9900>
                            <xsl:value-of select="'auto'"/>
                          </式样:字体族_9900>
                        </式样:字体声明_990D>
                      </xsl:if>
                      <xsl:if test="./c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea">
                        <式样:字体声明_990D>
                          <xsl:attribute name="标识符_9902">
                            <xsl:value-of select="concat($fileName,'.btj_szz_l_font')"/>
                          </xsl:attribute>
                          <xsl:attribute name="名称_9903">
                            <xsl:value-of select="./c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                          </xsl:attribute>
                          <xsl:attribute name="替换字体_9904">
                            <xsl:value-of select="./c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                          </xsl:attribute>
                          <式样:字体族_9900>
                            <xsl:value-of select="'auto'"/>
                          </式样:字体族_9900>
                        </式样:字体声明_990D>
                      </xsl:if>
                      
                        <xsl:if test="./c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('valAx_la_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                      <!--2014-4-23 end-->
                      
                        <!--图例字体-->
                        <xsl:if test="./c:chart/c:legend/c:legendEntry/c:txPr/a:p/a:pPr/a:defRPr/a:ea">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('legendEntry_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:legend/c:legendEntry/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:legend/c:legendEntry/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:chart/c:legend/c:legendEntry/c:txPr/a:p/a:pPr/a:defRPr/a:latin">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('legendEntry_la_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:legend/c:legendEntry/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:legend/c:legendEntry/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:chart/c:legend/c:txPr/a:p/a:pPr/a:defRPr/a:ea">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('legend_ea_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:legend/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:legend/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:chart/c:legend/c:txPr/a:p/a:pPr/a:defRPr/a:latin">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('legend_la_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:legend//c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:legend//c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>

                        <!--数据表字体-->
                        <xsl:if test="./c:chart/c:plotArea/c:dTable/c:txPr/a:p/a:pPr/a:defRPr/a:ea">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('dTable_ea_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:dTable/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:dTable/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:chart/c:plotArea/c:dTable/c:txPr/a:p/a:pPr/a:defRPr/a:latin">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat('dTable_la_',$fileName)"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:dTable/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:plotArea/c:dTable/c:txPr/a:p/a:pPr/a:defRPr/a:la/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>

                        <!--图表标题字体-->
                      
                      <!--2014-4-23, update by Qihy, 式样中的标识符和图表中的式样引用标识符不对应，将titleDef_ea_chartX和titleDef_la_chartX改为chartX.btj_bt_l_font和chartX.btj_bt_e_font, start-->
                        <xsl:if test="./c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat($fileName,'.btj_bt_e_font')"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat($fileName,'.btj_bt_l_font')"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:la/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat($fileName,'.btj_bt_e_font')"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                        <xsl:if test="./c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin">
                            <式样:字体声明_990D>
                                <xsl:attribute name="标识符_9902">
                                    <xsl:value-of select="concat($fileName,'.btj_bt_l_font')"/>
                                </xsl:attribute>
                                <xsl:attribute name="名称_9903">
                                    <xsl:value-of select="./c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <xsl:attribute name="替换字体_9904">
                                    <xsl:value-of select="./c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin/@typeface"/>
                                </xsl:attribute>
                                <式样:字体族_9900>
                                    <xsl:value-of select="'auto'"/>
                                </式样:字体族_9900>
                            </式样:字体声明_990D>
                        </xsl:if>
                    </xsl:for-each>
                </xsl:if>
              <!--2014-4-23  end-->
              
            </式样:字体集_990C>
            <xsl:call-template name="Cellstyle"/>
            <!--<xsl:if test="/ws:spreadsheets/ws:styleSheet[ws:dxfs/ws:dxf]">
        <xsl:call-template name="dxfs"/>
      </xsl:if>-->
        </式样:式样集_990B>
    </xsl:template>

    <!--字体声明-->
    <xsl:template match="ws:fonts">
        <xsl:for-each select="ws:font/ws:name">
            <式样:字体声明_990D>
                <xsl:attribute name="标识符_9902">
                    <xsl:variable name="id" select="position()"/>
                    <xsl:value-of select="concat('font_',$id)"/>
                </xsl:attribute>
                <xsl:attribute name="名称_9903">
                    <xsl:value-of select="@val"/>
                </xsl:attribute>
                <!--<xsl:attribute name="替换字体_9904">
          <xsl:value-of select="'永中宋体'"/>
        </xsl:attribute>-->
                <式样:字体族_9900>
                    <xsl:value-of select="'auto'"/>
                </式样:字体族_9900>
            </式样:字体声明_990D>
        </xsl:for-each>
    </xsl:template>
  
  <!--2014-6-9, add by Qihy, 增加sharedString.xml中字体的转换， start-->
  <xsl:template name="siRFont">
    <xsl:param name="pos"/>
    <xsl:if test="ws:r/ws:rPr/ws:rFont">
      <xsl:for-each select="ws:r">
      <式样:字体声明_990D>
        <xsl:attribute name="标识符_9902">
          <xsl:value-of select="concat('font_', $pos, '_', position())"/>
        </xsl:attribute>
        <xsl:attribute name="名称_9903">
          <xsl:value-of select="./ws:rPr/ws:rFont/@val"/>
        </xsl:attribute>
        <式样:字体族_9900>
          <xsl:value-of select="'auto'"/>
        </式样:字体族_9900>
      </式样:字体声明_990D>
      </xsl:for-each>
    </xsl:if>
  </xsl:template>
  <xsl:template match="ws:r">
    <xsl:value-of select="position()"/>
  </xsl:template>
  <!--2014-6-9 end-->
  
    <!--单元格式样模板-->
    <!--Modified by LDM in 2011/04/21-->
    <xsl:template name="Cellstyle">
        <!--缺省单元格式样-->
        <!--Modified by LDM in 2011/04/21-->
        <式样:单元格式样集_9915>
            <式样:单元格式样_9916 标识符_E7AC="DEFSTYLE" 名称_E7AD="DEFSTYLE" 类型_E7AE="auto">
                <表:字体格式_E7A7>
                    <字:字体_4128 西文字体引用_4129="font_00000" 中文字体引用_412A="font_00001" 字号_412D="12.0"/>
                </表:字体格式_E7A7>
                <表:对齐格式_E7A8>
                    <表:水平对齐方式_E700>general</表:水平对齐方式_E700>
                    <表:垂直对齐方式_E701>center</表:垂直对齐方式_E701>
                    <表:文字排列方向_E703>t2b-l2r-0e-0w</表:文字排列方向_E703>
                </表:对齐格式_E7A8>
                <!--表:数字格式_E7A9 分类名称_E740="general" 格式码_E73F="general"/-->
            </式样:单元格式样_9916>

            <!--其它单元格式样-->
            <!--Modified by LDM in 2011/04/21-->
            <xsl:for-each select="./ws:spreadsheets/ws:styleSheet/ws:cellXfs/ws:xf">
                <xsl:variable name="cellStyleSeq">
                    <xsl:value-of select="position()"/>
                </xsl:variable>
                <式样:单元格式样_9916>
                    <xsl:attribute name="标识符_E7AC">
                        <xsl:value-of select="concat('CELLSTYLE_',$cellStyleSeq)"/>
                    </xsl:attribute>
                    <xsl:attribute name="名称_E7AD">
                        <xsl:value-of select="concat('CELLSTYLE_',$cellStyleSeq)"/>
                    </xsl:attribute>
                    <xsl:attribute name="类型_E7AE">
                        <xsl:value-of select="'auto'"/>
                    </xsl:attribute>

                    <!--字体格式-->
                    <!--Modified by LDM in 2011/04/21-->
                    <xsl:call-template name="FontStyle"/>

                    <!--对齐格式-->
                    <!--Modified by LDM in 2011/04/21-->
                    <xsl:choose>
                        <xsl:when test="./ws:alignment">
                            <xsl:call-template name="Alignment"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <表:对齐格式_E7A8>
                                <表:水平对齐方式_E700>general</表:水平对齐方式_E700>
                                <表:垂直对齐方式_E701>bottom</表:垂直对齐方式_E701>
                                <表:文字排列方向_E703>t2b-l2r-0e-0w</表:文字排列方向_E703>
                            </表:对齐格式_E7A8>
                        </xsl:otherwise>
                    </xsl:choose>

                    <!--数字格式 修改添加。李杨2012-3-19-->
                  <!--2014-3-20 update by Qihy, 对代码做了调整，该处涉及许多问题，start-->
                  <xsl:if test="./@numFmtId">
                    <xsl:variable name="fid">
                      <xsl:value-of select="./@numFmtId"/>
                    </xsl:variable>

                    <表:数字格式_E7A9>
                    <xsl:choose>
                      <xsl:when test ="$fid='11'">
                        <xsl:attribute name ="分类名称_E740">
                          <xsl:value-of select ="'scientific'"/>
                        </xsl:attribute>
                        <xsl:attribute name="格式码_E73F">
                          <xsl:value-of select ="'0.00E+00'"/>
                        </xsl:attribute>
                      </xsl:when>

                      <!--2014-5-6, add by Qihy, 修复bug3299数字单元格格式错误， start-->
                      <xsl:when test ="$fid='2'">
                        <xsl:attribute name ="分类名称_E740">
                          <xsl:value-of select ="'custom'"/>
                        </xsl:attribute>
                        <xsl:attribute name="格式码_E73F">
                          <xsl:value-of select ="'0.00'"/>
                        </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='3'">
                        <xsl:attribute name ="分类名称_E740">
                          <xsl:value-of select ="'custom'"/>
                        </xsl:attribute>
                        <xsl:attribute name="格式码_E73F">
                          <xsl:value-of select ="'#,##0'"/>
                        </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='13'">
                        <xsl:attribute name ="分类名称_E740">
                          <xsl:value-of select ="'custom'"/>
                        </xsl:attribute>
                        <xsl:attribute name="格式码_E73F">
                          <xsl:value-of select ="'# ??/??'"/>
                        </xsl:attribute>
                      </xsl:when>
                      <!--2014-5-6 end-->

                      <xsl:when test ="$fid='4'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'currency'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'#,##0.00'"/>
                          </xsl:attribute>
                      </xsl:when>
                        <!-- 20140415 update by lingfeng 增加数字格式 start -->
                        <xsl:when test ="$fid='5'">
                            <xsl:attribute name ="分类名称_E740">
                                <xsl:value-of select ="'custom'"/>
                            </xsl:attribute>
                            <xsl:attribute name="格式码_E73F">
                                <xsl:value-of select ="'$#,##0_);($#,##0)'"/>
                            </xsl:attribute>
                        </xsl:when>
                      
                      <!--2014-4-26, delete by Qihy, bug3283单元格格式不正确（176非固定值而是根据实例文档中的格式码获取）， start-->
                        <!--<xsl:when test ="$fid='176'">
                            <xsl:attribute name ="分类名称_E740">
                                <xsl:value-of select ="'custom'"/>
                            </xsl:attribute>
                            <xsl:attribute name="格式码_E73F">
                                <xsl:value-of select ="'$#,##0_);($#,##0)'"/>
                            </xsl:attribute>
                        </xsl:when>-->
                      <!--2014-4-26 end-->
                        <!--end-->
                      <xsl:when test ="$fid='10'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'percentage'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'0.00%'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='12'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'fraction'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'# ?/?'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='14'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'date'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'yyyy-m-d;@'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <!--NumberFormat.xlsx  F2列-->
                      <xsl:when test ="$fid='9'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'percentage'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'0%'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <!--G2-->
                      <xsl:when test ="$fid='43'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'accounting'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * &quot;-&quot;??_ ;_ @_ '"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='20'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'h:mm'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='19'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'h:mm:ss AM/PM'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='15'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'d-mmm-yy'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='16'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'d-mmm'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='17'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'mmm-yy'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='18'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'h:mm AM/PM'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='21'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'h:mm:ss'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='22'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'yyyy/m/d h:mm'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='27' or $fid='36' or $fid='50' or $fid='52' or $fid='57'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'yyyy&quot;年&quot;m&quot;月&quot;'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='28' or $fid='29' or $fid='51' or $fid='53' or $fid='54' or $fid='58'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'m&quot;月&quot;d&quot;日&quot;'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='30'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'m-d-yy'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='31'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='32'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'h&quot;时&quot;mm&quot;分&quot;'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='33'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'h&quot;时&quot;mm&quot;分&quot;ss&quot;秒&quot;'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='34' or $fid='55'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'上午/下午h&quot;时&quot;mm&quot;分&quot;'"/>
                          </xsl:attribute>
                      </xsl:when>
                      <xsl:when test ="$fid='35' or $fid='56'">
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'上午/下午h&quot;时&quot;mm&quot;分&quot;ss&quot;秒&quot;'"/>
                          </xsl:attribute>
                      </xsl:when>

                      <!--2014-3-23, add by Qihy, 增加格式码的转换， start-->
                      <xsl:when test="$fid ='45'">
                        <xsl:attribute name ="分类名称_E740">
                          <xsl:value-of select ="'custom'"/>
                        </xsl:attribute>
                        <xsl:attribute name="格式码_E73F">
                          <xsl:value-of select ="'mm:ss'"/>
                        </xsl:attribute>
                      </xsl:when>
                      <xsl:when test="$fid ='46'">
                        <xsl:attribute name ="分类名称_E740">
                          <xsl:value-of select ="'custom'"/>
                        </xsl:attribute>
                        <xsl:attribute name="格式码_E73F">
                          <xsl:value-of select ="'[h]:mm:ss'"/>
                        </xsl:attribute>
                      </xsl:when>
                      <xsl:when test="$fid ='47'">
                        <xsl:attribute name ="分类名称_E740">
                          <xsl:value-of select ="'custom'"/>
                        </xsl:attribute>
                        <xsl:attribute name="格式码_E73F">
                          <xsl:value-of select ="'mm:ss.0'"/>
                        </xsl:attribute>
                      </xsl:when>
                      <xsl:when test="$fid ='48'">
                        <xsl:attribute name ="分类名称_E740">
                          <xsl:value-of select ="'##0.0E+0'"/>
                        </xsl:attribute>
                        <xsl:attribute name="格式码_E73F">
                          <xsl:value-of select ="'mm:ss.0'"/>
                        </xsl:attribute>
                      </xsl:when>
                      <xsl:when test="$fid ='49'">
                        <xsl:attribute name ="分类名称_E740">
                          <xsl:value-of select ="'custom'"/>
                        </xsl:attribute>
                        <xsl:attribute name="格式码_E73F">
                          <xsl:value-of select ="'@'"/>
                        </xsl:attribute>
                      </xsl:when>
                      <!--2014-3-23 end-->
                      
                      <xsl:otherwise>
                        <xsl:for-each select="ancestor::ws:styleSheet/ws:numFmts/ws:numFmt">
                          <xsl:if test ="@numFmtId = $fid">
                            <xsl:attribute name ="分类名称_E740">
                              <xsl:value-of select ="'custom'"/>
                            </xsl:attribute>
                            <xsl:attribute name="格式码_E73F">
                              <xsl:choose>
                                <xsl:when test="contains(@formatCode,'h:mm\ AM/PM;@')">
                                  <xsl:value-of select="'h:mm AM/PM;@'"/>
                                </xsl:when>
                                <xsl:when test="contains(@formatCode,'h:mm:ss\ AM/PM;@') or contains(@formatCode,'h:mm:ss\ AM/PM')">
                                  <xsl:value-of select="'h:mm:ss AM/PM;@'"/>
                                </xsl:when>
                                <xsl:when test="contains(@formatCode,'上午/下午h&quot;时&quot;mm&quot;分&quot;;@')">
                                  <xsl:value-of select="'[DBNum1]上午/下午h&quot;时&quot;mm&quot;分&quot;;@'"/>
                                </xsl:when>
                                <xsl:when test="contains(@formatCode,'aaaa;@')">
                                  <xsl:value-of select="'aaaa;@'"/>
                                </xsl:when>
                                <xsl:when test="contains(@formatCode,'h&quot;时&quot;mm&quot;分&quot;;@')">
                                  <xsl:value-of select="'h&quot;时&quot;mm&quot;分&quot;;@'"/>
                                </xsl:when>
                                  <xsl:when test="contains(./@formatCode,'&quot;$&quot;#,##0_);\(&quot;$&quot;#,##0\)')">
                                      <xsl:value-of select="'$#,##0_);($#,##0)'"/>
                                  </xsl:when>
                                <xsl:otherwise>
                                  <xsl:value-of select="./@formatCode"/>
                                </xsl:otherwise>
                              </xsl:choose>
                            </xsl:attribute>
                          </xsl:if>
                        </xsl:for-each>
                      </xsl:otherwise>      
                    </xsl:choose>
                    </表:数字格式_E7A9>
                  </xsl:if>
                  
                    <!-- update by 凌峰 BUG_2977:符号丢失 20140316 start -->
                    <!--<xsl:if test="./@numFmtId">
                        <xsl:variable name="fid">
                            <xsl:value-of select="./@numFmtId"/>
                        </xsl:variable>

                        <xsl:for-each select="../../ws:numFmts/ws:numFmt[@numFmtId=$fid]">
                            <表:数字格式_E7A9>
                                <xsl:attribute name="分类名称_E740">
                                    <xsl:value-of select="./@formatType"/>
                                </xsl:attribute>
                                <xsl:variable name ="geshima" select ="@formatCode"/>
                                <xsl:attribute name="格式码_E73F">-->
                                  
                                  
                                    <!--<xsl:if test="contains($geshima,'&amp;')">
                    <xsl:variable name ="v" select ="translate($geshima,'&amp;','')"/>
                    <xsl:value-of select ="$v"/>
                  </xsl:if>
                  <xsl:if test ="not(contains($geshima,'&amp;'))">
                    <xsl:value-of select="@formatCode"/>
                  </xsl:if>
                  <xsl:value-of select="@formatCode"/>-->
                                  
                                  
                                    <!-- 20130226 update by xuzhenwei BUG_2669:单元格内容显示不正确 start -->
                                    <!--<xsl:choose>
                                        <xsl:when test="contains($geshima,'h:mm\ AM/PM;@')">
                                            <xsl:value-of select="'h:mm AM/PM;@'"/>
                                        </xsl:when>
                                        <xsl:when test="contains($geshima,'h:mm:ss\ AM/PM;@') or contains($geshima,'h:mm:ss\ AM/PM')">
                                            <xsl:value-of select="'h:mm:ss AM/PM;@'"/>
                                        </xsl:when>
                                        <xsl:when test="contains($geshima,'上午/下午h&quot;时&quot;mm&quot;分&quot;;@')">
                                            <xsl:value-of select="'[DBNum1]上午/下午h&quot;时&quot;mm&quot;分&quot;;@'"/>
                                        </xsl:when>
                                        <xsl:when test="contains($geshima,'aaaa;@')">
                                            <xsl:value-of select="'aaaa;@'"/>
                                        </xsl:when>
                                        <xsl:when test="contains($geshima,'h&quot;时&quot;mm&quot;分&quot;;@')">
                                            <xsl:value-of select="'h&quot;时&quot;mm&quot;分&quot;;@'"/>
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:value-of select="@formatCode"/>
                                        </xsl:otherwise>
                                    </xsl:choose>
                                    --><!-- end --><!--
                                </xsl:attribute>
                            </表:数字格式_E7A9>
                        </xsl:for-each>-->

                      <!--2014-3-20, update by Qihy, 注释掉，start-->
                        <!--<xsl:if test ="$fid='179'">
                            <表:数字格式_E7A9>
                                <xsl:attribute name ="分类名称_E740">
                                    <xsl:value-of select ="'currency'"/>
                                </xsl:attribute>
                                <xsl:attribute name="格式码_E73F">
                                    <xsl:value-of select ="'#,##0.000[$₮-450];\-#,##0.000[$₮-450]'"/>
                                </xsl:attribute>
                            </表:数字格式_E7A9>
                        </xsl:if>-->
                      <!--2014-3-20 end-->
                      
                        <!--end-->
                        
                        <!--<xsl:if test ="$fid='11'">
                            <表:数字格式_E7A9>
                                <xsl:attribute name ="分类名称_E740">
                                    <xsl:value-of select ="'scientific'"/>
                                </xsl:attribute>
                                <xsl:attribute name="格式码_E73F">
                                    <xsl:value-of select ="'0.00E+00'"/>
                                </xsl:attribute>
                            </表:数字格式_E7A9>
                        </xsl:if>
                        <xsl:if test ="$fid='4'">
                            <表:数字格式_E7A9>
                                <xsl:attribute name ="分类名称_E740">
                                    <xsl:value-of select ="'currency'"/>
                                </xsl:attribute>
                                <xsl:attribute name="格式码_E73F">
                                    <xsl:value-of select ="'#,##0.00'"/>
                                </xsl:attribute>
                            </表:数字格式_E7A9>
                        </xsl:if>
                        <xsl:if test ="$fid='10'">
                            <表:数字格式_E7A9>
                                <xsl:attribute name ="分类名称_E740">
                                    <xsl:value-of select ="'percentage'"/>
                                </xsl:attribute>
                                <xsl:attribute name="格式码_E73F">
                                    <xsl:value-of select ="'0.00%'"/>
                                </xsl:attribute>
                            </表:数字格式_E7A9>
                        </xsl:if>
                        <xsl:if test ="$fid='12'">
                            <表:数字格式_E7A9>
                                <xsl:attribute name ="分类名称_E740">
                                    <xsl:value-of select ="'fraction'"/>
                                </xsl:attribute>
                                <xsl:attribute name="格式码_E73F">
                                    <xsl:value-of select ="'# ?/?'"/>
                                </xsl:attribute>
                            </表:数字格式_E7A9>
                        </xsl:if>
                        <xsl:if test ="$fid='14'">
                            <表:数字格式_E7A9>
                                <xsl:attribute name ="分类名称_E740">
                                    <xsl:value-of select ="'date'"/>
                                </xsl:attribute>
                                <xsl:attribute name="格式码_E73F">
                                    <xsl:value-of select ="'yyyy-m-d;@'"/>
                                </xsl:attribute>
                            </表:数字格式_E7A9>
                        </xsl:if>-->
                        <!--NumberFormat.xlsx  F2列-->
                        <!--<xsl:if test ="$fid='9'">
                            <表:数字格式_E7A9>
                                <xsl:attribute name ="分类名称_E740">
                                    <xsl:value-of select ="'percentage'"/>
                                </xsl:attribute>
                                <xsl:attribute name="格式码_E73F">
                                    <xsl:value-of select ="'0%'"/>
                                </xsl:attribute>
                            </表:数字格式_E7A9>
                        </xsl:if>-->
                        <!--G2-->
                        <!--<xsl:if test ="$fid='43'">
                            <表:数字格式_E7A9>
                                <xsl:attribute name ="分类名称_E740">
                                    <xsl:value-of select ="'accounting'"/>
                                </xsl:attribute>
                                <xsl:attribute name="格式码_E73F">
                                    <xsl:value-of select ="'_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * &quot;-&quot;??_ ;_ @_ '"/>
                                </xsl:attribute>
                            </表:数字格式_E7A9>
                        </xsl:if>-->
                        <!--M2-->

                      <!-- 2014-3-11 add by Qihy,修复BUG2940-区域公式集多处错误中的金钱格式和字体颜色错误 start -->
                      <!--<xsl:if test ="$fid='176'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'￥#,##0;[Red]￥-#,##0'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='177'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'￥#,##0.00;[Red]￥-#,##0.00'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>-->
                      <!-- 2014-3-11 end -->
                      
                      <!-- 2014-3-10 update by Qihy, BUG_3059 ooxml-uof集成测试中(2010),创建日期显示有误,并增加其他日期时间格式码的转换 start -->
                        <!--<xsl:if test ="$fid='20'">
                            <表:数字格式_E7A9>
                                <xsl:attribute name ="分类名称_E740">
                                    <xsl:value-of select ="'custom'"/>
                                </xsl:attribute>
                                <xsl:attribute name="格式码_E73F">
                                    <xsl:value-of select ="'h:mm'"/>
                                </xsl:attribute>
                            </表:数字格式_E7A9>
                        </xsl:if>-->
                        <!-- 20130325 update by xuzhenwei BUG_2754:oo-uof 功能测试 部分单元格格式不正确 start 备注：自定义类型的日期格式写死-->
                        <!--<xsl:if test ="$fid='19'">
                            <表:数字格式_E7A9>
                                <xsl:attribute name ="分类名称_E740">
                                    <xsl:value-of select ="'custom'"/>
                                </xsl:attribute>
                                <xsl:attribute name="格式码_E73F">
                                    <xsl:value-of select ="'h:mm:ss AM/PM'"/>
                                </xsl:attribute>
                            </表:数字格式_E7A9>
                        </xsl:if>-->
                        <!-- end -->

                      <!-- 2014-3-3 update by Qihy, BUG_2940:oo-uof 功能测试 日期时间转换错误 start -->
                      <!--<xsl:if test ="$fid='15'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'d-mmm-yy'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='16'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'d-mmm'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='17'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'mmm-yy'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='18'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'h:mm AM/PM'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='21'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'h:mm:ss'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='22'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'yyyy/m/d h:mm'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      --><!-- 2014-3-3 end --><!--
                      <xsl:if test ="$fid='27' or $fid='36' or $fid='50' or $fid='52' or $fid='57'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'yyyy&quot;年&quot;m&quot;月&quot;'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='28' or $fid='29' or $fid='51' or $fid='53' or $fid='54' or $fid='58'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'m&quot;月&quot;d&quot;日&quot;'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='30'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'m-d-yy'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='31'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='32'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'h&quot;时&quot;mm&quot;分&quot;'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='33'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'h&quot;时&quot;mm&quot;分&quot;ss&quot;秒&quot;'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='34' or $fid='55'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'上午/下午h&quot;时&quot;mm&quot;分&quot;'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>
                      <xsl:if test ="$fid='35' or $fid='56'">
                        <表:数字格式_E7A9>
                          <xsl:attribute name ="分类名称_E740">
                            <xsl:value-of select ="'custom'"/>
                          </xsl:attribute>
                          <xsl:attribute name="格式码_E73F">
                            <xsl:value-of select ="'上午/下午h&quot;时&quot;mm&quot;分&quot;ss&quot;秒&quot;'"/>
                          </xsl:attribute>
                        </表:数字格式_E7A9>
                      </xsl:if>-->
                      <!-- 2014-3-10 end -->
          <!--<xsl:call-template name="NumFamat">
            <xsl:with-param name="fid" select="@numFmtId"/>
          </xsl:call-template>-->
                    <!--</xsl:if>-->
                  <!--2014-3-20 end-->
                  
                    <!--边框设置-->
                    <!--Modified by LDM in 2011/04/21-->
                    <xsl:if test="./@borderId">
                        <xsl:call-template name="CellBorder"/>
                    </xsl:if>

                    <!--填充设置 未完成-->
                  
                    <!--2014-3-23, update by Qihy, 修复BUG3113 互操作-ooxml-uof-oo（2010）内容差异， start-->
                    <!--<xsl:if test="(@applyFill and @applyFill=1) or (@fillId and @fillId !='0')">-->
                  
                  <!--2014-5-6, update by Qihy, 修复BUG3270和BUG3298 单元格背景色丢失问题， start-->
                  <!--<xsl:if test="(@applyFill and @applyFill=1) and (@fillId and @fillId !='0')">-->
                  <xsl:if test="((@applyFill and @applyFill='1') and (@fillId and @fillId !='0') and (@xfId and @xfId ='0')) or ((@applyFill and @applyFill='1') and (@fillId and @fillId !='0') and (@xfId and @xfId !='0'))or (not((@applyFill and @applyFill='1')) and (@fillId and @fillId !='0') and (@xfId and @xfId !='0'))">
                    <!--2014-5-6 end-->
                    
                    <!--2014-3-23 end-->
                    
                        <xsl:call-template name="tc"/>
                    </xsl:if>
                </式样:单元格式样_9916>
            </xsl:for-each>
            <xsl:if test="/ws:spreadsheets/ws:styleSheet[ws:dxfs/ws:dxf]">
                <xsl:call-template name="dxfs"/>
            </xsl:if>
        </式样:单元格式样集_9915>
    </xsl:template>

    <!--字体格式模板-->
    <!--Modified by LDM in 2011/04/21-->
    <xsl:template name="FontStyle">
        <表:字体格式_E7A7>
            <xsl:variable name="id">
                <xsl:value-of select="@fontId"/>
            </xsl:variable>
            <字:字体_4128>
                <xsl:if test="@fontId">
                    <xsl:attribute name="西文字体引用_4129">
                        <xsl:value-of select="concat('font_',number($id) + 1)"/>
                    </xsl:attribute>
                    <xsl:attribute name="中文字体引用_412A">
                        <xsl:value-of select="concat('font_',number($id) + 1)"/>
                    </xsl:attribute>
                    <xsl:for-each select="ancestor::ws:styleSheet/ws:fonts/ws:font[position()=$id + 1]">
                        <!--字号-->
                        <xsl:if test="./ws:sz/@val">
                            <xsl:attribute name="字号_412D">
                                <xsl:value-of select="./ws:sz/@val"/>
                            </xsl:attribute>
                        </xsl:if>
						
						
                        <!--zl 2015-4-15-->                        
                        <xsl:if test="./ws:color">
                            <xsl:attribute name="颜色_412F">
                                <xsl:variable name="COLOR" select="./ws:color"/>
                                <xsl:choose>
                                  <xsl:when test="$COLOR='#ffffff'">
                                    <xsl:value-of select="'#000000'"/>
                                  </xsl:when>
                                  <xsl:otherwise>
                                    <!--<xsl:apply-templates select="./ws:color"/>-->
                                    <xsl:value-of select="'#000000'"/>
                                  </xsl:otherwise>
                                </xsl:choose>
                            </xsl:attribute>
                        </xsl:if>
                      <!--zl 2015-4-15-->
						
						
                    </xsl:for-each>
                </xsl:if>
            </字:字体_4128>
            <xsl:for-each select="ancestor::ws:styleSheet/ws:fonts/ws:font[position()=$id + 1]">
                <!--粗体-->
                <xsl:if test="./ws:b[not(@val) or @val!=0]">
                    <字:是否粗体_4130>true</字:是否粗体_4130>
                </xsl:if>
                <!--斜体-->
                <xsl:if test="./ws:i[not(@val) or @val!=0]">
                    <字:是否斜体_4131>true</字:是否斜体_4131>
                </xsl:if>
                <!--删除线-->
                <xsl:if test="./ws:strike[not(@val) or @val!=0]">
                    <字:删除线_4135>single</字:删除线_4135>
                </xsl:if>
                <!--修改下划线 李杨2012-3-9-->
                <xsl:if test="./ws:u[not(@val) or @val!='none']">
                    <字:下划线_4136>
                        <xsl:attribute name="线型_4137">
                            <xsl:choose>
                                <xsl:when test="./ws:u[@val='double' or @val='doubleAccounting']">
                                    <xsl:value-of select="'double'"/>
                                </xsl:when>
                                <xsl:when test="./ws:u[@val='single' or @val='singleAccounting' ]">
                                    <xsl:value-of select="'single'"/>
                                </xsl:when>
                                <xsl:when test="not(./ws:u/@val)">
                                    <xsl:value-of select ="'single'"/>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:value-of select="'single'"/>
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:attribute>
                        <xsl:if test="./ws:color">
                            <xsl:attribute name="颜色_412F">
                                <xsl:apply-templates select="./ws:color"/>
                            </xsl:attribute>
                        </xsl:if>
                    </字:下划线_4136>
                </xsl:if>
                <!--空心字体-->
                <xsl:if test="./ws:outline[not(@val) or @val!=0]">
                    <字:是否空心_413E>true</字:是否空心_413E>
                </xsl:if>
                <!--字体阴影-->
                <xsl:if test="./ws:shadow[not(@val) or @val!=0]">
                    <字:是否阴影_4140>true</字:是否阴影_4140>
                </xsl:if>
                <!--上下标-->
                <xsl:if test="./ws:vertAlign[not(@val) or @val!=0]">
                    <字:上下标类型_4143>
                        <xsl:if test="./ws:vertAlign/@val='superscript'">
                            <xsl:value-of select="'sup'"/>
                        </xsl:if>
                        <xsl:if test="./ws:vertAlign/@val='subscript'">
                            <xsl:value-of select="'sub'"/>
                        </xsl:if>
                    </字:上下标类型_4143>
                </xsl:if>
            </xsl:for-each>
        </表:字体格式_E7A7>
    </xsl:template>

    <!--对齐格式模板-->
    <!--Modified by LDM in 2011/04/21-->
  <xsl:template name="Alignment">
    <表:对齐格式_E7A8>
      <xsl:choose>
        <xsl:when test="not(./ws:alignment/@horizontal) and not(./ws:alignment/@vertical) and ./ws:alignment/@textRotation=255">
          <表:水平对齐方式_E700>general</表:水平对齐方式_E700>
          <表:垂直对齐方式_E701>bottom</表:垂直对齐方式_E701>
          <表:文字排列方向_E703>r2l-t2b-0e-90w</表:文字排列方向_E703>
          <表:文字旋转角度_E704>0</表:文字旋转角度_E704>
        </xsl:when>
        <xsl:otherwise>
          <!--水平对齐方式-->
          <表:水平对齐方式_E700>
            <xsl:choose>
              <xsl:when test="./ws:alignment[@horizontal]">
                <xsl:variable name="horiAlign" select="./ws:alignment/@horizontal"/>
                <xsl:choose>
                  <xsl:when test="$horiAlign='centerContinuous'">
                    <xsl:value-of select="'center-across-selection'"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="$horiAlign"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:when>
              <xsl:otherwise>
                <!-- update by xuzhenwei BUG_2668、BUG_2903:内容格式显示不正确 2013-01-20 start
            <xsl:value-of select="'left'"/> -->
                <!-- end -->
              </xsl:otherwise>
            </xsl:choose>
          </表:水平对齐方式_E700>
          <!--垂直对齐方式-->
          <表:垂直对齐方式_E701>
            <xsl:choose>
              <xsl:when test="./ws:alignment[@vertical]">
                <xsl:value-of select="ws:alignment/@vertical"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="'bottom'"/>
              </xsl:otherwise>
            </xsl:choose>
          </表:垂直对齐方式_E701>

          <!--缩进-->
          <xsl:if test="ws:alignment[@indent]">
            <表:缩进_E702>
              <xsl:value-of select="ws:alignment/@indent"/>
            </表:缩进_E702>
          </xsl:if>
          <!--文字旋转角度-->
          <xsl:if test="ws:alignment[@textRotation]">
            <表:文字旋转角度_E704>
              <xsl:variable name="dre" select="ws:alignment/@textRotation"/>
              <xsl:choose>
                <!--修改 李杨2012-3-9-->
                <xsl:when test="$dre&gt;90 and $dre&lt;180">
                  <xsl:value-of select="90 - $dre"/>
                </xsl:when>
                <xsl:when test="$dre=180">
                  <xsl:value-of select="-90"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="$dre"/>
                </xsl:otherwise>
              </xsl:choose>
            </表:文字旋转角度_E704>
          </xsl:if>
          <!--自动换行-->
          <xsl:if test="ws:alignment[@wrapText]">
            <表:是否自动换行_E705>true</表:是否自动换行_E705>
          </xsl:if>
          <!--缩小字体填充-->
          <xsl:if test="ws:alignment[@shrinkToFit]">
            <表:是否缩小字体填充_E706>true</表:是否缩小字体填充_E706>
          </xsl:if>
        </xsl:otherwise>
      </xsl:choose>
    </表:对齐格式_E7A8>
  </xsl:template>

    <!--数字格式模板 未完成-->

    <xsl:template name="NumFamat">
        <xsl:param name="fid"/>
        <表:数字格式_E7A9>
            <xsl:variable name="vvxa">
                <xsl:value-of select="/ws:spreadsheets/ws:styleSheet/ws:numFmts/ws:numFmt[@numFmtId=$fid]"/>
            </xsl:variable>

            <xsl:choose>
                <xsl:when test="contains(./@formatCode,'&quot;$&quot;#,##0_);\(&quot;$&quot;#,##0\)')">
                    <xsl:attribute name="分类名称_E740">
                        <xsl:value-of select="custom"/>
                    </xsl:attribute>
                    <xsl:attribute name="格式码_E73F">
                        <xsl:value-of select="'$#,##0_);($#,##0)'"/>
                    </xsl:attribute>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:attribute name="分类名称_E740">
                        <xsl:value-of select="./@formatType"/>
                    </xsl:attribute>
                    <xsl:attribute name="格式码_E73F">
                        <xsl:value-of select="./@formatCode"/>
                    </xsl:attribute>
                </xsl:otherwise>
            </xsl:choose>

        </表:数字格式_E7A9>
    </xsl:template>
    <!--
      <xsl:choose>
        <xsl:when test="contains($vvxa,'&quot;￥')">
          <xsl:attribute name="表:分类名称">
            <xsl:value-of select="'currency'"/>
          </xsl:attribute>
          <xsl:attribute name="表:格式码">
            <xsl:value-of select="'[$￥-804]#,##0.00;[$￥-804]-#,##0.00 '"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains($vvxa,'&quot;US')">
          <xsl:attribute name="表:分类名称">
            <xsl:value-of select="'accounting'"/>
          </xsl:attribute>
          <xsl:attribute name="表:格式码">
            <xsl:value-of select="'_-US$* #,##0.00_ ;_-US$* -#,##0.00 ;_-US$* &quot;-&quot;??_ ;_-@_  '"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains($vvxa,'&quot;年')">
          <xsl:attribute name="表:分类名称">
            <xsl:value-of select="'date'"/>
          </xsl:attribute>
          <xsl:attribute name="表:格式码">
            <xsl:value-of select="'[DBNum1]yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;;@'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains($vvxa,'mm')">
          <xsl:attribute name="表:分类名称">
            <xsl:value-of select="'time'"/> 
          </xsl:attribute>
          <xsl:attribute name="表:格式码">
            <xsl:value-of select="'h:mm AM/PM;@'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains($vvxa,'0000') or contains($vvxa,'DBNum')">
          <xsl:attribute name="表:分类名称">
            <xsl:value-of select="'specialization'"/>
          </xsl:attribute>
          <xsl:attribute name="表:格式码">
            <xsl:value-of select="'[DBNum1]General'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains($vvxa,'E+')">
          <xsl:attribute name="表:分类名称">
            <xsl:value-of select="'scientific'"/> 
          </xsl:attribute>
          <xsl:attribute name="表:格式码">
            <xsl:value-of select="'0.00E+00'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains($vvxa,'?/')">
          <xsl:attribute name="表:分类名称">
            <xsl:value-of select="'fraction'"/>
          </xsl:attribute>
          <xsl:attribute name="表:格式码">
            <xsl:value-of select="'# ?/?'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains($vvxa,'0.0000%')">
          <xsl:attribute name="表:分类名称">
            <xsl:value-of select="'fraction'"/>
          </xsl:attribute>
          <xsl:attribute name="表:格式码">
            <xsl:value-of select="'# ?/?'"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:when test="contains($vvxa,'0_ ')">
          <xsl:attribute name="表:分类名称">
            <xsl:value-of select="'number'"/>
          </xsl:attribute>
          <xsl:attribute name="表:格式码">
            <xsl:value-of select="'0.00_ '"/>
          </xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:attribute name="表:分类名称"> 
            <xsl:value-of select="'general'"/>
          </xsl:attribute>
          <xsl:attribute name="表:格式码">
            <xsl:value-of select="'general'"/>
          </xsl:attribute>
        </xsl:otherwise>
      </xsl:choose>
      -->
    <!--单元格边框模板-->
    <!--Modified by LDM in 2011/04/21-->
    <xsl:template name="CellBorder">
        <!--边框中增加了属性：阴影类型  李杨2011-11-9-->
        <表:边框_4133>
            <xsl:variable name="borderId" select="@borderId"/>
            <xsl:for-each select="ancestor::ws:styleSheet/ws:borders/ws:border[position()=$borderId+1]">
              
                <!-- update by 凌峰 BUG_2976：单元格式样边框部分效果丢失 20140126 start -->
                <xsl:if test="./ws:left/@style and not(./ws:left/@style='none')">
                    <uof:左_C613>
                        <xsl:variable name="style" select="./ws:left/@style"/>
                        <xsl:call-template name="BorderStyle">
                            <xsl:with-param name="style" select="$style"/>
                        </xsl:call-template>
                        <!--边框颜色-->
                        <xsl:if test="./ws:left/ws:color">
                            <xsl:attribute name="颜色_C611">
                                <xsl:apply-templates select="./ws:left/ws:color"/>
                            </xsl:attribute>
                        </xsl:if>
                        <!--边框宽度，默认转换，OOXML中不能设置-->
                        <xsl:attribute name="宽度_C60F">
                            <xsl:value-of select="'1.0'"/>
                        </xsl:attribute>
                    </uof:左_C613>
                </xsl:if>
                <xsl:if test="./ws:top/@style and not(./ws:top/@style='none')">
                    <uof:上_C614>
                        <xsl:variable name="style" select="./ws:top/@style"/>
                        <xsl:call-template name="BorderStyle">
                            <xsl:with-param name="style" select="$style"/>
                        </xsl:call-template>
                        <!--边框颜色-->
                        <xsl:if test="./ws:top/ws:color">
                            <xsl:attribute name="颜色_C611">
                                <xsl:apply-templates select="./ws:top/ws:color"/>
                            </xsl:attribute>
                        </xsl:if>
                        <!--边框宽度，默认转换，OOXML中不能设置-->
                        <xsl:attribute name="宽度_C60F">
                            <xsl:value-of select="'1.0'"/>
                        </xsl:attribute>
                    </uof:上_C614>
                </xsl:if>
                <xsl:if test="./ws:right/@style and not(./ws:right/@style='none')">
                    <uof:右_C615>
                        <xsl:variable name="style" select="./ws:right/@style"/>
                        <xsl:call-template name="BorderStyle">
                            <xsl:with-param name="style" select="$style"/>
                        </xsl:call-template>
                        <!--边框颜色-->
                        <xsl:if test="./ws:right/ws:color">
                            <xsl:attribute name="颜色_C611">
                                <xsl:apply-templates select="./ws:right/ws:color"/>
                            </xsl:attribute>
                        </xsl:if>
                        <!--边框宽度，默认转换，OOXML中不能设置-->
                        <xsl:attribute name="宽度_C60F">
                            <xsl:value-of select="'1.0'"/>
                        </xsl:attribute>
                    </uof:右_C615>
                </xsl:if>
                <xsl:if test="./ws:bottom/@style and not(./ws:bottom/@style='none')">
                    <uof:下_C616>
                        <xsl:variable name="style" select="./ws:bottom/@style"/>
                        <xsl:call-template name="BorderStyle">
                            <xsl:with-param name="style" select="$style"/>
                        </xsl:call-template>
                        <!--边框颜色-->
                        <xsl:if test="./ws:bottom/ws:color">
                            <xsl:attribute name="颜色_C611">
                                <xsl:apply-templates select="./ws:bottom/ws:color"/>
                            </xsl:attribute>
                        </xsl:if>
                        <!--边框宽度，默认转换，OOXML中不能设置-->
                        <xsl:attribute name="宽度_C60F">
                            <xsl:value-of select="'1.0'"/>
                        </xsl:attribute>
                    </uof:下_C616>
                </xsl:if>
                <xsl:if test="./@diagonalDown='1'">
                    <uof:对角线1_C617>
                        <xsl:variable name="style" select="./ws:diagonal/@style"/>
                        <xsl:call-template name="BorderStyle">
                            <xsl:with-param name="style" select="$style"/>
                        </xsl:call-template>
                        <!--边框颜色-->
                        <xsl:if test="./ws:diagonal/ws:color">
                            <xsl:attribute name="颜色_C611">
                                <xsl:apply-templates select="./ws:diagonal/ws:color"/>
                            </xsl:attribute>
                        </xsl:if>
                        <!--边框宽度，默认转换，OOXML中不能设置-->
                        <xsl:attribute name="宽度_C60F">
                            <xsl:value-of select="'1.0'"/>
                        </xsl:attribute>
                    </uof:对角线1_C617>
                </xsl:if>
                <xsl:if test="./@diagonalUp='1'">
                    <uof:对角线2_C618>
                        <xsl:variable name="style" select="./ws:diagonal/@style"/>
                        <xsl:call-template name="BorderStyle">
                            <xsl:with-param name="style" select="$style"/>
                        </xsl:call-template>
                        <!--边框颜色-->
                        <xsl:if test="./ws:diagonal/ws:color">
                            <xsl:attribute name="颜色_C611">
                                <xsl:apply-templates select="./ws:diagonal/ws:color"/>
                            </xsl:attribute>
                        </xsl:if>
                        <!--边框宽度，默认转换，OOXML中不能设置-->
                        <xsl:attribute name="宽度_C60F">
                            <xsl:value-of select="'1.0'"/>
                        </xsl:attribute>
                    </uof:对角线2_C618>
                    <!--BUG_2976 end-->
                </xsl:if>
            </xsl:for-each>
        </表:边框_4133>
    </xsl:template>

    <!--边框类型转换 OOXML TO UOF-->
    <!--Modified by LDM in 2011/04/21-->
    <xsl:template name="BorderStyle">
        <xsl:param name="style"/>
        <xsl:choose>
            <xsl:when test="$style='double'">
                <xsl:attribute name="线型_C60D">
                    <xsl:value-of select="'double'"/>
                </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
                <xsl:attribute name="线型_C60D">
                    <xsl:value-of select="'single'"/>
                </xsl:attribute>
                <xsl:attribute name="虚实_C60E">
                    <xsl:choose>
                        <xsl:when test="$style='dashDot' or $style='slantDashDot' or $style='mediumDashDot'">
                            <xsl:value-of select="'dash-dot'"/>
                        </xsl:when>
                        <xsl:when test="$style='dashDotDot'">
                            <xsl:value-of select="'dash-dot-dot'"/>
                        </xsl:when>
                        <xsl:when test="$style='dashed'">
                            <xsl:value-of select="'dash'"/>
                        </xsl:when>
                        <xsl:when test="$style='dotted' or $style='hair'">
                            <xsl:value-of select="'square-dot'"/>
                        </xsl:when>
                        <xsl:when test="$style='mediumDashDotDot'">
                            <xsl:value-of select="'dash-dot-dot'"/>
                        </xsl:when>
                        <xsl:when test="$style='mediumDashed'">
                            <xsl:value-of select="'square-dot'"/>
                        </xsl:when>
                    </xsl:choose>
                </xsl:attribute>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template name="tc">
        <表:填充_E7A3 是否填充随图形旋转_8067="false">
            <xsl:variable name="fillid" select="@fillId"/>
            <!--图案-->
            <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill">
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill[@patternType='solid']">
                    <图:颜色_8004>
                        <!-- 20130515 update by xuzhenwei 单元格背景颜色填充丢失 start -->
                        <xsl:choose>
                            <xsl:when test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor[@theme]">
                                <xsl:variable name="themeId" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor/@theme"/>
                              <xsl:variable name="themeTint" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor/@tint"/>
                                <xsl:choose>
                                    <xsl:when test="$themeId=0">
                                        <xsl:value-of select="'#ffffff'"/>
                                    </xsl:when>
                                    <xsl:when test="$themeId=1">
                                        <xsl:value-of select="'#000000'"/>
                                    </xsl:when>
                                    <!--李杨添加 2012-3-12-->
                                    <xsl:when test ="$themeId=2">
                                        <xsl:value-of select ="'#f5f5f8'"/>
                                    </xsl:when>
                                    <xsl:when test ="$themeId=3">
                                      
                                      <!--2014-3-24, update by Qihy, 取值错误，start-->
                                        <!--<xsl:value-of select ="'#082e4'"/>-->
                                      <xsl:value-of select ="'#082e54'"/>
                                      <!--2014-3-24 end-->
                                      
                                    </xsl:when>
                                  <xsl:when test ="$themeId=5 and $themeTint=0.59999389629810485">
                                    <xsl:value-of select ="'#e9afa3'"/>
                                  </xsl:when>
                                  <!--2014-5-1, update by 凌峰, 修复BUG3181 oox-uof-oox单元格背景颜色丢失， start-->
                                  <xsl:when test ="$themeId=7 and $themeTint=0.79998168889431442">
                                    
                                    <!--2014-5-29, update by Qihy, 颜色取值错误， start-->
                                    <!--<xsl:value-of select ="'#DCDCDC'"/>-->
                                    <xsl:value-of select ="'#e8e0ff'"/>
                                    <!--2014-5-29 end-->
                                    
                                  </xsl:when>
                                  <!--2014-5-1 end-->
                                  
                                  <!--2014-5-31, update by Qihy, 颜色取值错误， start-->
                                  <xsl:when test ="$themeId=9 and $themeTint=0.79998168889431442">
                                    <xsl:value-of select ="'#F6E7DC'"/>
                                  </xsl:when>
                                  <!--2014-5-31 end-->

                                  <xsl:otherwise>
                                        <xsl:variable name="color" select="/ws:spreadsheets//a:theme/a:themeElements/a:clrScheme/*[position()=$themeId + 1]/a:srgbClr/@val"/>
                                        
                                        <!-- update by 凌峰 BUG_2978：单元格背景色丢失  20140308 start-->
                                        <xsl:value-of select="concat('#',translate($color,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
                                    </xsl:otherwise>
                                </xsl:choose>
                            </xsl:when>
                            <xsl:when test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor[@rgb]">
                                <xsl:variable name="col" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor/@rgb"/>
                                <xsl:variable name="coll" select="substring($col,3,8)"/>
                                <xsl:value-of select="concat('#',translate($coll,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
                                <!--end-->
                                
                            </xsl:when>
                            <xsl:when test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor[@indexed]">
                                <xsl:variable name="indexedId" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor/@indexed"/>
                              
                              <!--2014-3-22 update by Qihy, indexedId取值对应的颜色值不正确， start-->
                                <!--<xsl:choose>
                                    <xsl:when test="$indexedId=0">
                                        <xsl:value-of select="'#ffffff'"/>
                                    </xsl:when>
                                    <xsl:when test="$indexedId=1">
                                        <xsl:value-of select="'#000000'"/>
                                    </xsl:when>
                                    <xsl:when test ="$indexedId=2">
                                        <xsl:value-of select ="'#f5f5f8'"/>
                                    </xsl:when>
                                    <xsl:when test ="$indexedId=3">
                                        <xsl:value-of select ="'#082e54'"/>
                                    </xsl:when>
                                    <xsl:when test ="$indexedId=41">
                                        <xsl:value-of select ="'#c1f8ff'"/>
                                    </xsl:when>
                                    <xsl:otherwise>
                                        <xsl:value-of select="'#000000'"/>
                                    </xsl:otherwise>
                                </xsl:choose>-->
                              <xsl:call-template name ="indexColor">
                                <xsl:with-param name="indexed" select="$indexedId"/>
                              </xsl:call-template>
                              <!--2014-3-22 end-->
                              
                            </xsl:when>
                        </xsl:choose>
                    </图:颜色_8004>
                </xsl:if>
                <xsl:if test ="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill[@patternType='none']">
                    <图:颜色_8004>
                        <xsl:value-of select ="'#ffffff'"/>
                    </图:颜色_8004>
                </xsl:if>
                <!--5yue5ri xiayanxiachecked-->
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill[@patternType!='solid' and  not(@patternType='none')]">
                    <xsl:variable name="type" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill[@patternType!='solid']/@patternType"/>
                    <图:图案_800A>
                        <xsl:attribute name="类型_8008">
                            <!-- update by xuzhenwei 2013-01-09 sourceforge bug_2667:图案显示不正确 start-->
                            <xsl:call-template name="tatype">
                                <xsl:with-param name="tttype" select="$type"/>
                            </xsl:call-template>
                            <!-- end -->
                        </xsl:attribute>
                        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill[@patternType!='solid']/ws:fgColor">
                            <xsl:variable name="aa">
                                <xsl:call-template name="fgColor">
                                    <xsl:with-param name="fillid2" select="$fillid"/>
                                </xsl:call-template>
                            </xsl:variable>
                            <xsl:attribute name="前景色_800B">
                              <xsl:if test="$aa = 'eeece1'">
                                <xsl:value-of select="000000"/>
                              </xsl:if>
                              <xsl:if test="not($aa = 'eeece1')">
                                <xsl:value-of select="$aa"/>
                              </xsl:if>
                            </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill[@patternType!='solid']/ws:bgColor">
                            <xsl:variable name="bb">
                                <xsl:call-template name="bgColor">
                                    <xsl:with-param name="fillid2" select="$fillid"/>
                                </xsl:call-template>
                            </xsl:variable>
                          
                          <!--2014-3-24, update by Qihy, 属性名称错误， start-->
                          <!--<xsl:attribute name="图:背景色">-->
                          <xsl:attribute name="背景色_800C">
                          <!--2014-3-24 end-->
                            
                                <xsl:value-of select="$bb"/>
                            </xsl:attribute>
                        </xsl:if>
                    </图:图案_800A>
                </xsl:if>
            </xsl:if>
            <!--渐变-->
            <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill">
                <!--2yue 2 夏艳霞 gaide-->
                <!--yx,add path,2010.5.6-->
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[@type='path']">
                    <图:渐变_800D>
                        <xsl:call-template name="becolor"/>
                        <xsl:attribute name="种子类型_8010">
                            <xsl:value-of select="'rectangle'"/>
                        </xsl:attribute>
                        <xsl:attribute name="起始浓度_8011">1.0</xsl:attribute>
                        <xsl:attribute name="终止浓度_8012">1.0</xsl:attribute>
                        <xsl:attribute name="渐变方向_8013">0</xsl:attribute>
                        <xsl:attribute name="边界_8014">50</xsl:attribute>

                        <xsl:choose>
                            <xsl:when test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[not(@top) and not(@bottom) and not(@left) and not(@right)]">
                                <xsl:attribute name="种子X位置_8015">30</xsl:attribute>
                                <xsl:attribute name="种子Y位置_8016">30</xsl:attribute>
                            </xsl:when>
                            <xsl:when test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[not(@top) and not(@bottom)]">
                                <xsl:attribute name="种子X位置_8015">60</xsl:attribute>
                                <xsl:attribute name="种子Y位置_8016">30</xsl:attribute>
                            </xsl:when>
                            <xsl:when test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[not(@left) and not(@right)]">
                                <xsl:attribute name="种子X位置_8015">30</xsl:attribute>
                                <xsl:attribute name="种子Y位置_8016">60</xsl:attribute>
                            </xsl:when>

                            <xsl:otherwise>
                                <xsl:choose>
                                    <xsl:when test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[@left!='1' and @right!='1' and @top!='1' and @bottom!='1']">
                                        <xsl:attribute name="种子X位置_8015">50</xsl:attribute>
                                        <xsl:attribute name="种子Y位置_8016">50</xsl:attribute>
                                    </xsl:when>
                                    <xsl:otherwise>
                                        <xsl:attribute name="种子X位置_8015">60</xsl:attribute>
                                        <xsl:attribute name="种子Y位置_8016">30</xsl:attribute>
                                    </xsl:otherwise>
                                </xsl:choose>
                            </xsl:otherwise>

                        </xsl:choose>
                    </图:渐变_800D>
                </xsl:if>
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[@type!='path' and @left]">
                    <图:渐变_800D 种子类型_8010="square" 起始浓度_8011="1.0" 终止浓度_8012="1.0" 渐变方向_8013="0" 边界_8014="5" 种子X位置_8015="100" 种子Y位置_8016="100">
                        <xsl:call-template name="becolor"/>
                    </图:渐变_800D>
                </xsl:if>
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[@type!='path' and not(@left)]">
                    <图:渐变_800D 种子类型_8010="square" 起始浓度_8011="1.0" 终止浓度_8012="1.0" 渐变方向_8013="0" 边界_8014="5" 种子X位置_8015="30" 种子Y位置_8016="30">
                        <xsl:call-template name="becolor"/>
                    </图:渐变_800D>
                </xsl:if>
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[not(@degree) and not(@type)]">
                    <图:渐变_800D 种子类型_8010="axial" 起始浓度_8011="1.0" 终止浓度_8012="1.0" 渐变方向_8013="90" 边界_8014="50" 种子X位置_8015="100" 种子Y位置_8016="100">
                        <xsl:call-template name="becolor"/>
                    </图:渐变_800D>
                </xsl:if>
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[@degree and @degree='45']">
                    <图:渐变_800D 种子类型_8010="axial" 起始浓度_8011="100.0" 终止浓度_8012="100.0" 渐变方向_8013="45" 边界_8014="80" 种子X位置_8015="17" 种子Y位置_8016="100">
                        <xsl:call-template name="becolor"/>
                    </图:渐变_800D>
                </xsl:if>
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[@degree and @degree='90']">
                    <图:渐变_800D 种子类型_8010="axial" 起始浓度_8011="100.0" 终止浓度_8012="100.0" 渐变方向_8013="0" 边界_8014="100" 种子X位置_8015="100" 种子Y位置_8016="100">
                        <xsl:call-template name="becolor"/>
                    </图:渐变_800D>
                </xsl:if>
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[@degree and @degree='135']">
                    <图:渐变_800D 种子类型_8010="axial" 起始浓度_8011="100.0" 终止浓度_8012="100.0" 渐变方向_8013="315" 边界_8014="100" 种子X位置_8015="100" 种子Y位置_8016="100">
                        <xsl:call-template name="becolor"/>
                    </图:渐变_800D>
                </xsl:if>
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[@degree and @degree='180']">
                    <图:渐变_800D 种子类型_8010="linear" 起始浓度_8011="1.0" 终止浓度_8012="1.0" 渐变方向_8013="270" 边界_8014="50" 种子X位置_8015="100" 种子Y位置_8016="100">
                        <xsl:call-template name="becolor2"/>
                    </图:渐变_800D>
                </xsl:if>
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[@degree and @degree='225']">
                    <图:渐变_800D 种子类型_8010="linear" 起始浓度_8011="1.0" 终止浓度_8012="1.0" 渐变方向_8013="225" 边界_8014="50" 种子X位置_8015="100" 种子Y位置_8016="100">
                        <xsl:call-template name="becolor2"/>
                    </图:渐变_800D>
                </xsl:if>
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[@degree and @degree='270']">
                    <图:渐变_800D 种子类型_8010="linear" 起始浓度_8011="1.0" 终止浓度_8012="1.0" 渐变方向_8013="360" 边界_8014="80" 种子X位置_8015="100" 种子Y位置_8016="100">
                        <xsl:call-template name="becolor2"/>
                    </图:渐变_800D>
                </xsl:if>
                <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[@degree and @degree='315']">
                    <图:渐变_800D 种子类型_8010="linear" 起始浓度_8011="1.0" 终止浓度_8012="1.0" 渐变方向_8013="315" 边界_8014="80" 种子X位置_8015="100" 种子Y位置_8016="100">
                        <xsl:call-template name="becolor2"/>
                    </图:渐变_800D>
                </xsl:if>
                <!--xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill[@degree and @degree='360']">
					<图:渐变 uof:locID="g0037" uof:attrList="起始色 终止色 种子类型 起始浓度 终止浓度 渐变方向 边界 种子X位置 种子Y位置 类型" 图:种子类型="linear" 图:起始浓度="1.0" 图:终止浓度="1.0" 图:渐变方向="360" 图:边界="80" 图:种子X位置="100" 图:种子Y位置="100">
						<xsl:call-template name="becolor"/>
					</图:渐变>
				</xsl:if-->
            </xsl:if>
        </表:填充_E7A3>
    </xsl:template>
    <xsl:template name="fgColor">
        <xsl:param name="fillid2"/>
        <xsl:variable name="fillid" select="$fillid2"/>
        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()= $fillid + 1]/ws:patternFill/ws:fgColor[@theme]">
            <xsl:variable name="themeId" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor/@theme"/>
          <xsl:variable name="tint" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor/@tint"/>
          <xsl:choose>
                <!--李杨修改 2012-3-12-->
                <xsl:when test="$themeId=0">
                    <xsl:value-of select="'#ffffff'"/>
                </xsl:when>
                <xsl:when test="$themeId=1">
                    <xsl:value-of select="'#000000'"/>
                </xsl:when>
            <xsl:when test="$themeId=3 and $tint=-0.499984740745262">
              <xsl:value-of select="'#00203f'"/>
            </xsl:when>
            <xsl:when test="$themeId=5 and $tint=-0.499984740745262">
              <xsl:value-of select="'#800000'"/>
            </xsl:when>
            <xsl:when test="$themeId=8 and $tint=-0.499984740745262">
              <xsl:value-of select="'#006662'"/>
            </xsl:when>
                <xsl:otherwise>
                    <xsl:variable name="color" select="/ws:spreadsheets//a:theme/a:themeElements/a:clrScheme/*[position()=$themeId+1]/a:srgbClr/@val"/>
                    <xsl:value-of select="concat('#',translate($color,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:if>
        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor[@rgb]">
            <xsl:variable name="col" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor/@rgb"/>
            <xsl:variable name="coll" select="substring($col,3,8)"/>
            <xsl:value-of select="concat('#',translate($coll,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
        </xsl:if>
      
      <!--2014-3-24, add by Qihy, 增加当fgColor属性为indexed时，颜色的转换， start-->
      <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor[@indexed]">
        <xsl:call-template name ="indexColor">
          <xsl:with-param name="indexed" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:fgColor/@indexed"/>
        </xsl:call-template>
      </xsl:if>
      <!--2014-3-24 end-->
      
    </xsl:template>
    <xsl:template name="bgColor">
        <xsl:param name="fillid2"/>
        <xsl:variable name="fillid" select="$fillid2"/>
        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:bgColor[@theme]">
            <xsl:variable name="themeId" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:bgColor/@theme"/>
          <xsl:choose>
                <xsl:when test="$themeId=0">
                    <xsl:value-of select="'#000000'"/>
                </xsl:when>
                <xsl:when test="$themeId=1">
                    <xsl:value-of select="'#ffffff'"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:variable name="color" select="/ws:spreadsheets//a:theme/a:themeElements/a:clrScheme/*[position()=$themeId]/a:srgbClr/@val"/>
                    <xsl:value-of select="concat('#',translate($color,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:if>
        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:bgColor[@rgb]">
            <xsl:variable name="col" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:bgColor/@rgb"/>
            <xsl:variable name="coll" select="substring($col,3,8)"/>
            <xsl:value-of select="concat('#',translate($coll,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
        </xsl:if>
      
      <!--2014-3-24, add by Qihy, 增加当bgColor属性为indexed时，颜色的转换， start-->
      <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:bgColor[@indexed]">
        <xsl:call-template name ="indexColor">
          <xsl:with-param name="indexed" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:patternFill/ws:bgColor/@indexed"/>
        </xsl:call-template>
      </xsl:if>
      <!--2014-3-24 end-->
      
    </xsl:template>
    <!-- update by xuzhenwei 2013-01-09 sourceforge bug_2667:图案显示不正确 start bug_2724再修正20130326 -->
    <xsl:template name="tatype">
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
                <xsl:value-of select="'ptn004'"/>
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
                <xsl:value-of select="'ptn041'"/>
            </xsl:when>
            <xsl:when test="$tttype='darkUp'">
                <xsl:value-of select="'ptn016'"/>
            </xsl:when>
            <xsl:when test="$tttype='shingle'">
                <xsl:value-of select="'ptn031'"/>
            </xsl:when>
            <xsl:when test="$tttype='darkGrid'">
                <xsl:value-of select="'ptn045'"/>
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
            <xsl:otherwise>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    <!-- end -->
    <xsl:template name="ttype2">
        <xsl:param name="tttype"/>
        <xsl:choose>
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

    <xsl:template name="becolor">
        <xsl:variable name="fillid3" select="@fillId"/>
        <xsl:attribute name="起始色_800E">
            <xsl:call-template name="qss1">
                <xsl:with-param name="fillid" select="$fillid3"/>
            </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="终止色_800F">
            <xsl:call-template name="zzs1">
                <xsl:with-param name="fillid" select="$fillid3"/>
            </xsl:call-template>
        </xsl:attribute>
    </xsl:template>
    <xsl:template name="qss1">
        <xsl:param name="fillid"/>
        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[position()=1]/ws:color[@theme]">
            <xsl:call-template name="themeColor">
                <xsl:with-param name="themeId" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[position()=1]/ws:color/@theme"/>
            </xsl:call-template>
        </xsl:if>
        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[position()=1]/ws:color[@rgb]">
            <xsl:variable name="rgbcolor" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[position()=1]/ws:color/@rgb"/>
            <xsl:variable name="endcolor">
                <xsl:value-of select="concat('#',translate($rgbcolor,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
            </xsl:variable>
            <xsl:value-of select="concat('#',translate(substring($endcolor,4,6),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
        </xsl:if>
    </xsl:template>
    <xsl:template name="zzs1">
        <xsl:param name="fillid"/>
        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[2]/ws:color[@theme]">
            <xsl:call-template name="themeColor">
                <xsl:with-param name="themeId" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[2]/ws:color/@theme"/>
            </xsl:call-template>
        </xsl:if>
        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[2]/ws:color[@rgb]">
            <xsl:variable name="rgbcolor" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[2]/ws:color/@rgb"/>
            <xsl:variable name="endcolor">
                <xsl:value-of select="concat('#',translate($rgbcolor,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
            </xsl:variable>
            <xsl:value-of select="concat('#',translate(substring($endcolor,4,6),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
        </xsl:if>
    </xsl:template>
    <xsl:template name="becolor2">
        <xsl:variable name="fillid3" select="@fillId"/>
        <xsl:attribute name="起始色_800E">
            <xsl:call-template name="qss2">
                <xsl:with-param name="fillid" select="$fillid3"/>
            </xsl:call-template>
        </xsl:attribute>
        <xsl:attribute name="终止色_800F">
            <xsl:call-template name="zzs2">
                <xsl:with-param name="fillid" select="$fillid3"/>
            </xsl:call-template>
        </xsl:attribute>
    </xsl:template>
    <xsl:template name="qss2">
        <xsl:param name="fillid"/>
        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[2]/ws:color[@theme]">
            <xsl:call-template name="themeColor">
                <xsl:with-param name="themeId" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[2]/ws:color/@theme"/>
            </xsl:call-template>
        </xsl:if>
        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[2]/ws:color[@rgb]">
            <xsl:variable name="rgbcolor" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[2]/ws:color/@rgb"/>
            <xsl:variable name="endcolor">
                <xsl:value-of select="concat('#',translate($rgbcolor,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
            </xsl:variable>
            <xsl:value-of select="concat('#',translate(substring($endcolor,4,6),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
        </xsl:if>
    </xsl:template>
    <xsl:template name="zzs2">
        <xsl:param name="fillid"/>
        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[position()=1]/ws:color[@theme]">
            <xsl:call-template name="themeColor">
                <xsl:with-param name="themeId" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[position()=1]/ws:color/@theme"/>
            </xsl:call-template>
        </xsl:if>
        <xsl:if test="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[position()=1]/ws:color[@rgb]">
            <xsl:variable name="rgbcolor" select="ancestor::ws:styleSheet/ws:fills/ws:fill[position()=$fillid+1]/ws:gradientFill/ws:stop[position()=1]/ws:color/@rgb"/>
            <xsl:variable name="endcolor">
                <xsl:value-of select="concat('#',translate($rgbcolor,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
            </xsl:variable>
            <xsl:value-of select="concat('#',translate(substring($endcolor,4,6),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
        </xsl:if>
    </xsl:template>
    <xsl:template name="themeColor">
        <xsl:param name="themeId"/>
        <xsl:choose>
            <xsl:when test="$themeId=0">
                <xsl:value-of select="'#ffffff'"/>
            </xsl:when>
            <xsl:when test="$themeId=1">
                <xsl:value-of select="'#000000'"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:variable name="color" select="/ws:spreadsheets//a:theme/a:themeElements/a:clrScheme/*[position()=$themeId+1]/a:srgbClr/@val"/>
                <xsl:value-of select="concat('#',translate($color,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template match="a:defRPr" mode="czitiji">
        <xsl:choose>
            <xsl:when test="a:latin">
                <式样:字体声明_990D 标识符_9902="tbq_l_font">
                    <xsl:attribute name="名称_9903">
                        <xsl:value-of select="@typeface"/>
                    </xsl:attribute>
                    <式样:字体族_9900>
                        <xsl:value-of select="@typeface"/>
                    </式样:字体族_9900>
                </式样:字体声明_990D>
            </xsl:when>
            <xsl:when test="a:ea">
                <式样:字体声明_990D 标识符_9902="tbq_e_font">
                    <xsl:attribute name="名称_9903">
                        <xsl:value-of select="@typeface"/>
                    </xsl:attribute>
                    <式样:字体族_9900>
                        <xsl:value-of select="@typeface"/>
                    </式样:字体族_9900>
                </式样:字体声明_990D>
            </xsl:when>
        </xsl:choose>
    </xsl:template>
    <xsl:template name="dxfs">
        <xsl:for-each select="/ws:spreadsheets/ws:styleSheet/ws:dxfs/ws:dxf">
            <!--<式样:单元格式样集_9915>-->
            <式样:单元格式样_9916>
                <xsl:attribute name="标识符_E7AC">
                    <xsl:value-of select="concat('tjgsh_',position())"/>
                </xsl:attribute>
                <xsl:attribute name="名称_E7AD">
                    <xsl:value-of select="concat('tjgsh_',position())"/>
                </xsl:attribute>
                <xsl:attribute name="类型_E7AE">
                    <xsl:value-of select="'auto'"/>
                </xsl:attribute>
                <xsl:if test="ws:font">
                    <表:字体格式_E7A7>
                        <字:字体_4128>
                            <!--<xsl:attribute name ="西文字体引用_4129">font_00000</xsl:attribute>
                <xsl:attribute name ="中文字体引用_412A">font_00001</xsl:attribute>
                <xsl:attribute name ="字号_412D">12.0</xsl:attribute>-->
                            <xsl:if test="ws:font/ws:color">
                                <xsl:attribute name="颜色_412F">
                                    <xsl:apply-templates select="ws:font/ws:color"/>
                                </xsl:attribute>
                            </xsl:if>
                        </字:字体_4128>
                        <xsl:if test="ws:font/ws:b[not(@val) or @val!=0]">
                            <字:是否粗体_4130>true</字:是否粗体_4130>
                        </xsl:if>
                        <xsl:if test="ws:font/ws:i[not(@val) or @val!=0]">
                            <字:是否斜体_4131>true</字:是否斜体_4131>
                        </xsl:if>
                        <xsl:if test="ws:font/ws:strike[not(@val) or @val!=0]">
                            <字:删除线_4135>single</字:删除线_4135>
                        </xsl:if>
                        <xsl:if test="ws:font/ws:u[not(@val='none')]">
                            <字:下划线_4136>
                                <xsl:attribute name="线型_4137">
                                    <xsl:choose>
                                        <xsl:when test="ws:font/ws:u[@val='double' or @val='doubleAccounting' ]">
                                            <xsl:value-of select="'double'"/>
                                        </xsl:when>
                                        <xsl:when test="ws:font/ws:u[@val='single' or @val='singleAccounting' ]">
                                            <xsl:value-of select="'single'"/>
                                        </xsl:when>
                                        <xsl:otherwise>
                                            <xsl:value-of select="'single'"/>
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:attribute>
                            </字:下划线_4136>
                        </xsl:if>
                    </表:字体格式_E7A7>
                </xsl:if>
                <!--<表:对齐格式_E7A8>
            <表:水平对齐方式_E700>general</表:水平对齐方式_E700>
            <表:垂直对齐方式_E701>center</表:垂直对齐方式_E701>
            <表:文字排列方向_E703>t2b-l2r-0e-0w</表:文字排列方向_E703>
          </表:对齐格式_E7A8>
          <表:数字格式_E7A9 分类名称_E740="general" 格式码_E73F="general"/>-->
              <xsl:if test="ws:fill">
                    <表:填充_E7A3>
                        <xsl:attribute name ="是否填充随图形旋转_8067">false</xsl:attribute>

                      <!--2014-3-26, add by Qihy, 增加图案填充，背景色和前景色填充情况不全， start-->
                      <xsl:if test="ws:fill/ws:patternFill[@patternType] and ws:fill/ws:patternFill[@patternType!='solid' and  not(@patternType='none')]">
                        <xsl:variable name="type" select="ws:fill/ws:patternFill/@patternType"/>
                        <图:图案_800A>
                          <xsl:attribute name="类型_8008">
                            <xsl:call-template name="ttype2">
                              <xsl:with-param name="tttype" select="$type"/>
                            </xsl:call-template>
                          </xsl:attribute>
                          <xsl:attribute name="前景色_800B">
                            <xsl:choose>
                              <xsl:when test ="ws:fill/ws:patternFill/ws:fgColor[@rgb]">
                                <xsl:value-of select="concat('#',translate(substring(ws:fill/ws:patternFill/ws:fgColor/@rgb,3,8),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>"
                              </xsl:when>
                              <xsl:when test ="ws:fill/ws:patternFill/ws:fgColor[@indexed]">
                                <xsl:call-template name="indexColor">
                                  <xsl:with-param name="indexed" select ="ws:fill/ws:patternFill/ws:fgColor/@indexed"/>
                                </xsl:call-template>
                              </xsl:when>
                              <xsl:when test="not(ws:fill/ws:patternFill/ws:fgColor)">
                                <xsl:value-of select="'#000000'"/>
                              </xsl:when>
                            </xsl:choose>
                          </xsl:attribute>
                          <xsl:attribute name="背景色_800C">
                            <xsl:choose>
                              <xsl:when test ="ws:fill/ws:patternFill/ws:bgColor[@rgb]">
                                <xsl:value-of select="concat('#',translate(substring(ws:fill/ws:patternFill/ws:bgColor/@rgb,3,8),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>"
                              </xsl:when>
                              <xsl:when test ="ws:fill/ws:patternFill/ws:bgColor[@indexed]">
                                <xsl:call-template name="indexColor">
                                  <xsl:with-param name="indexed" select ="ws:fill/ws:patternFill/ws:bgColor/@indexed"/>
                                </xsl:call-template>
                              </xsl:when>
                              <xsl:when test="not(ws:fill/ws:patternFill/ws:bgColor)">
                                <xsl:value-of select="'#ffffff'"/>
                              </xsl:when>
                            </xsl:choose>
                          </xsl:attribute>
                        </图:图案_800A>
                      </xsl:if>
                      <xsl:if test="ws:fill/ws:patternFill[not(@patternType)] or ws:fill/ws:patternFill[@patternType = 'none' or @patternType = 'solid']">
                      <图:颜色_8004>
                            <xsl:if test="ws:fill/ws:patternFill/ws:bgColor[@theme]">
                                <xsl:variable name="themeId" select="ws:fill/ws:patternFill/ws:bgColor/@theme"/>
                                <xsl:choose>
                                    <xsl:when test="$themeId=0">
                                        <xsl:value-of select="'#000000'"/>
                                    </xsl:when>
                                    <xsl:when test="$themeId=1">
                                        <xsl:value-of select="'#ffffff'"/>
                                    </xsl:when>
                                    <xsl:otherwise>
                                        <xsl:variable name="color" select="/ws:spreadsheets//a:theme/a:themeElements/a:clrScheme/*[position()=$themeId+1]/a:srgbClr/@val"/>

                                        <!-- update by 凌峰 BUG_2978：单元格背景色丢失  20140308 start-->
                                        <xsl:value-of select="concat('#',translate($color,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
                                    </xsl:otherwise>
                                </xsl:choose>
                            </xsl:if>
                            <xsl:if test="ws:fill/ws:patternFill/ws:bgColor[@rgb]">
                                <xsl:variable name="col" select="ws:fill/ws:patternFill/ws:bgColor/@rgb"/>
                                <xsl:variable name="coll" select="substring($col,3,8)"/>
                                <xsl:value-of select="concat('#',translate($coll,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
                                <!--end-->
                                
                            </xsl:if>

                          <!--2014-3-24, update by Qihy, 颜色填充取值为indexed时的转换， start-->
                          <xsl:if test="ws:fill/ws:patternFill/ws:bgColor[@indexed]">
                            <xsl:call-template name="indexColor">
                              <xsl:with-param name="indexed" select ="@indexed"/>
                            </xsl:call-template>
                          </xsl:if>
                          <!--2014-3-24 end-->
                        
                    </图:颜色_8004>
                      </xsl:if>
                      <!--2014-3-26 end-->
                      
                  </表:填充_E7A3>
              </xsl:if>
                <xsl:if test="ws:border">
                    <表:边框_4133>
                      
                      <!--2014-4-27, update by Qihy, bug3264, 图案填充时，边框线丢失， start-->
                      <xsl:if test="ws:border/ws:left/@style and not(./ws:left/@style='none')">
                        <uof:左_C613>
                          <xsl:variable name="style" select="ws:border/ws:left/@style"/>
                          <xsl:call-template name="BorderStyle">
                            <xsl:with-param name="style" select="$style"/>
                          </xsl:call-template>
                          <!--边框颜色-->
                          <xsl:if test="ws:border/ws:left/ws:color">
                            <xsl:attribute name="颜色_C611">
                              <xsl:apply-templates select="ws:border/ws:left/ws:color"/>
                            </xsl:attribute>
                          </xsl:if>
                          <!--边框宽度，默认转换，OOXML中不能设置-->
                          <xsl:attribute name="宽度_C60F">
                            <xsl:value-of select="'1.0'"/>
                          </xsl:attribute>
                        </uof:左_C613>
                      </xsl:if>
                      <xsl:if test="ws:border/ws:top/@style and not(ws:border/ws:top/@style='none')">
                        <uof:上_C614>
                          <xsl:variable name="style" select="ws:border/ws:top/@style"/>
                          <xsl:call-template name="BorderStyle">
                            <xsl:with-param name="style" select="$style"/>
                          </xsl:call-template>
                          <!--边框颜色-->
                          <xsl:if test="ws:border/ws:top/ws:color">
                            <xsl:attribute name="颜色_C611">
                              <xsl:apply-templates select="ws:border/ws:top/ws:color"/>
                            </xsl:attribute>
                          </xsl:if>
                          <!--边框宽度，默认转换，OOXML中不能设置-->
                          <xsl:attribute name="宽度_C60F">
                            <xsl:value-of select="'1.0'"/>
                          </xsl:attribute>
                        </uof:上_C614>
                      </xsl:if>
                      <xsl:if test="ws:border/ws:right/@style and not(ws:border/ws:right/@style='none')">
                        <uof:右_C615>
                          <xsl:variable name="style" select="ws:border/ws:right/@style"/>
                          <xsl:call-template name="BorderStyle">
                            <xsl:with-param name="style" select="$style"/>
                          </xsl:call-template>
                          <!--边框颜色-->
                          <xsl:if test="ws:border/ws:right/ws:color">
                            <xsl:attribute name="颜色_C611">
                              <xsl:apply-templates select="ws:border/ws:right/ws:color"/>
                            </xsl:attribute>
                          </xsl:if>
                          <!--边框宽度，默认转换，OOXML中不能设置-->
                          <xsl:attribute name="宽度_C60F">
                            <xsl:value-of select="'1.0'"/>
                          </xsl:attribute>
                        </uof:右_C615>
                      </xsl:if>
                      <xsl:if test="ws:border/ws:bottom/@style and not(ws:border/ws:bottom/@style='none')">
                        <uof:下_C616>
                          <xsl:variable name="style" select="ws:border/ws:bottom/@style"/>
                          <xsl:call-template name="BorderStyle">
                            <xsl:with-param name="style" select="$style"/>
                          </xsl:call-template>
                          <!--边框颜色-->
                          <xsl:if test="ws:border/ws:bottom/ws:color">
                            <xsl:attribute name="颜色_C611">
                              <xsl:apply-templates select="ws:border/ws:bottom/ws:color"/>
                            </xsl:attribute>
                          </xsl:if>
                          <!--边框宽度，默认转换，OOXML中不能设置-->
                          <xsl:attribute name="宽度_C60F">
                            <xsl:value-of select="'1.0'"/>
                          </xsl:attribute>
                        </uof:下_C616>
                      </xsl:if>
                      <xsl:if test="ws:border/@diagonalDown='1'">
                        <uof:对角线1_C617>
                          <xsl:variable name="style" select="ws:border/ws:diagonal/@style"/>
                          <xsl:call-template name="BorderStyle">
                            <xsl:with-param name="style" select="$style"/>
                          </xsl:call-template>
                          <!--边框颜色-->
                          <xsl:if test="ws:border/ws:diagonal/ws:color">
                            <xsl:attribute name="颜色_C611">
                              <xsl:apply-templates select="ws:border/ws:diagonal/ws:color"/>
                            </xsl:attribute>
                          </xsl:if>
                          <!--边框宽度，默认转换，OOXML中不能设置-->
                          <xsl:attribute name="宽度_C60F">
                            <xsl:value-of select="'1.0'"/>
                          </xsl:attribute>
                        </uof:对角线1_C617>
                      </xsl:if>
                      <xsl:if test="ws:border/@diagonalUp='1'">
                        <uof:对角线2_C618>
                          <xsl:variable name="style" select="ws:border/ws:diagonal/@style"/>
                          <xsl:call-template name="BorderStyle">
                            <xsl:with-param name="style" select="$style"/>
                          </xsl:call-template>
                          <!--边框颜色-->
                          <xsl:if test="ws:border/ws:diagonal/ws:color">
                            <xsl:attribute name="颜色_C611">
                              <xsl:apply-templates select="ws:border/ws:diagonal/ws:color"/>
                            </xsl:attribute>
                          </xsl:if>
                          <!--边框宽度，默认转换，OOXML中不能设置-->
                          <xsl:attribute name="宽度_C60F">
                            <xsl:value-of select="'1.0'"/>
                          </xsl:attribute>
                        </uof:对角线2_C618>
                      </xsl:if>
                      <!--2014-4-27 end-->
                      
                    </表:边框_4133>

                  <!--2014-5-29, add by Qihy, 单元格格式化不正确(此处可能考虑不是很全面，自定义格式码尚未考虑)， start-->
                  <xsl:if test="ws:numFmt">
                    <表:数字格式_E7A9>
                      <xsl:attribute name="分类名称_E740">
                        <xsl:value-of select="'custom'"/>
                      </xsl:attribute>
                      <xsl:attribute name="格式码_E73F">
                        <xsl:value-of select="ws:numFmt/@formatCode"/>
                      </xsl:attribute>
                    </表:数字格式_E7A9>
                  </xsl:if>
                  <!--2014-5-29 end-->
                </xsl:if>
            </式样:单元格式样_9916>
            <!--</式样:单元格式样集_9915>-->
        </xsl:for-each>
    </xsl:template>

    <xsl:template name="ssty"></xsl:template>

    <!--颜色设置模板-->
    <!--Modified by LDM in 2011/04/21-->
  <xsl:template match="ws:color" name ="Color">
        <xsl:if test="@theme">
            <xsl:variable name="theme" select="@theme"/>
            <xsl:call-template name="ThemeColor">
                <xsl:with-param name="col" select="$theme"/>
            </xsl:call-template>
        </xsl:if>
        <xsl:if test="@rgb">
            <xsl:variable name="color">
                <xsl:value-of select="@rgb"/>
            </xsl:variable>
            <xsl:value-of select="concat('#',translate(substring-after($color,'FF'),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))"/>
        </xsl:if>
        <xsl:if test ="@auto">
            <xsl:value-of select ="'auto'"/>
        </xsl:if>
      
      <!--2014-3-22 update by Qihy, start-->
        <!--李杨添加，不知道indexed代表什么意思。2012-5-17-->
        <!--<xsl:choose>
            <xsl:when test ="@indexed='64'">
                <xsl:value-of select ="'auto'"/>
            </xsl:when>
            <xsl:when test ="@indexed='62'">
                <xsl:value-of select ="'#2f4f2f'"/>
            </xsl:when>
            --><!--<xsl:otherwise >
        <xsl:value-of select ="'auto'"/>
      </xsl:otherwise>--><!--
        </xsl:choose>-->

      <xsl:if test ="@indexed">
        <xsl:call-template name ="indexColor">
          <xsl:with-param name="indexed" select="@indexed"/>
        </xsl:call-template>
      </xsl:if>
  </xsl:template>
  <!--2014-3-22 end-->
  
    <!--Theme Color值转换为RGB Color值-->
    <!--Modified by LDM in 2011/04/21-->
    <xsl:template name="ThemeColor">
        <xsl:param name="col"/>
        <xsl:choose>
            <xsl:when test="$col='0'">
                <xsl:value-of select="'#ffffff'"/>
            </xsl:when>
            <xsl:when test="$col='1'">
                <xsl:value-of select="'#000000'"/>
            </xsl:when>
            <xsl:when test="$col='2'">
                <xsl:value-of select="'#1f497d'"/>
            </xsl:when>
            <xsl:when test="$col='3'">
                <xsl:value-of select="'#1f497d'"/>
                <!--<xsl:value-of select="'#EEECE1'"/>-->
            </xsl:when>
            <xsl:when test="$col='4'">
                <xsl:value-of select="'#4f81bd'"/>
            </xsl:when>
            <xsl:when test="$col='5'">
                <xsl:value-of select="'#c0504c'"/>
            </xsl:when>
            <xsl:when test="$col='6'">
                <xsl:value-of select="'#9bbb59'"/>
            </xsl:when>
            <xsl:when test="$col='7'">
                <xsl:value-of select="'#8064a2'"/>
            </xsl:when>
            <xsl:when test="$col='8'">
                <xsl:value-of select="'#4bacc6'"/>
            </xsl:when>
            <xsl:when test="$col='9'">
                <xsl:value-of select="'#f79646'"/>
            </xsl:when>
            <xsl:when test="$col='10'">
              
              <!--2014-3-29, update by Qihy, bug3159, 超链接颜色不正确， start-->
                <!--<xsl:value-of select="'#ff0000'"/>-->
              <xsl:value-of select="'#0000ff'"/>
              <!--2014-3-29 end-->
              
            </xsl:when>
            <xsl:when test="$col='11'">
                <xsl:value-of select="'#800080'"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="'#ffffff'"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    <xsl:template name="default_runstyle_hyperlink">
        <式样:句式样集_990F>
            <式样:句式样_9910 标识符_4100="CELLSTYLE_bh" 名称_4101="" 类型_4102="default">
                <字:字体_4128 西文字体引用_4129="font_0" 中文字体引用_412A="font_1" 字号_412D="12.0" 字:相对字号="12.0" 颜色_412F="#0000ff"/>
                <字:下划线_4136 线型_4137="single" 颜色_412F="#0000ff"/>
            </式样:句式样_9910>
            <式样:句式样_9910 标识符_4100="CELLSTYLE_ah" 名称_4101="" 类型_4102="default">
                <字:字体_4128 西文字体引用_4129="font_0" 中文字体引用_412A="font_1" 字号_412D="12.0" 字:相对字号="12.0" 颜色_412F="#800080"/>
                <字:下划线_4136 线型_4137="single" 颜色_412F="#800080"/>
            </式样:句式样_9910>
        </式样:句式样集_990F>
        <!---->
    </xsl:template>

  <!--2014-3-22 add by Qihy, 增加indexed取值对应的颜色值， start-->
  <xsl:template name ="indexColor">
    <xsl:param name ="indexed"/>
    <xsl:choose>
      <xsl:when test="ancestor::ws:spreadsheets/ws:styleSheet/ws:colors/ws:indexedColors/ws:rgbColor[position() = $indexed+1]/@rgb">
        <xsl:value-of select ="concat('#', substring(ancestor::ws:spreadsheets/ws:styleSheet/ws:colors/ws:indexedColors/ws:rgbColor[position() = $indexed+1]/@rgb, 3, 8))"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$indexed='0' or $indexed='8'">
            <xsl:value-of select="'#000000'"/>
          </xsl:when>
          <xsl:when test="$indexed='1' or $indexed='9'">
            <xsl:value-of select="'#ffffff'"/>
          </xsl:when>
          <xsl:when test="$indexed='2' or $indexed='10'">
            <xsl:value-of select="'#ff0000'"/>
          </xsl:when>
          <xsl:when test="$indexed='3' or $indexed='11'">
            <xsl:value-of select="'#00ff00'"/>
          </xsl:when>
          <xsl:when test="$indexed='4' or $indexed='12'">
            <xsl:value-of select="'#0000ff'"/>
          </xsl:when>
          <xsl:when test="$indexed='5' or $indexed='13'">
            <xsl:value-of select="'#ffff00'"/>
          </xsl:when>
          <xsl:when test="$indexed='6' or $indexed='14'">
            <xsl:value-of select="'#ff00ff'"/>
          </xsl:when>
          <xsl:when test="$indexed='7' or $indexed='15'">
            <xsl:value-of select="'#00ffff'"/>
          </xsl:when>
          <xsl:when test="$indexed='16'">
            <xsl:value-of select="'#800000'"/>
          </xsl:when>
          <xsl:when test="$indexed='17'">
            <xsl:value-of select="'#008000'"/>
          </xsl:when>
          <xsl:when test="$indexed='18'">
            <xsl:value-of select="'#000080'"/>
          </xsl:when>
          <xsl:when test="$indexed='19'">
            <xsl:value-of select="'#808000'"/>
          </xsl:when>
          <xsl:when test="$indexed='20'">
            <xsl:value-of select="'#800080'"/>
          </xsl:when>
          <xsl:when test="$indexed='21'">
            <xsl:value-of select="'#008080'"/>
          </xsl:when>
          <xsl:when test="$indexed='22'">
            <xsl:value-of select="'#c0c0c0'"/>
          </xsl:when>
          <xsl:when test="$indexed='23'">
            <xsl:value-of select="'#808080'"/>
          </xsl:when>
          <xsl:when test="$indexed='24'">
            <xsl:value-of select="'#999ff'"/>
          </xsl:when>
          <xsl:when test="$indexed='25'">
            <xsl:value-of select="'#993366'"/>
          </xsl:when>
          <xsl:when test="$indexed='26'">
            <xsl:value-of select="'#ffffcc'"/>
          </xsl:when>
          <xsl:when test="$indexed='27'">
            <xsl:value-of select="'#ccffff'"/>
          </xsl:when>
          <xsl:when test="$indexed='28'">
            <xsl:value-of select="'#660066'"/>
          </xsl:when>
          <xsl:when test="$indexed='29'">
            <xsl:value-of select="'#ff8080'"/>
          </xsl:when>
          <xsl:when test="$indexed='30'">
            <xsl:value-of select="'#0066cc'"/>
          </xsl:when>
          <xsl:when test="$indexed='31'">
            <xsl:value-of select="'#ccccff'"/>
          </xsl:when>
          <xsl:when test="$indexed='32'">
            <xsl:value-of select="'#000080'"/>
          </xsl:when>
          <xsl:when test="$indexed='33'">
            <xsl:value-of select="'#ff00ff'"/>
          </xsl:when>
          <xsl:when test="$indexed='34'">
            <xsl:value-of select="'#ffff00'"/>
          </xsl:when>
          <xsl:when test="$indexed='35'">
            <xsl:value-of select="'#00ffff'"/>
          </xsl:when>
          <xsl:when test="$indexed='36'">
            <xsl:value-of select="'#800080'"/>
          </xsl:when>
          <xsl:when test="$indexed='37'">
            <xsl:value-of select="'#800000'"/>
          </xsl:when>
          <xsl:when test="$indexed='38'">
            <xsl:value-of select="'#008080'"/>
          </xsl:when>
          <xsl:when test="$indexed='39'">
            <xsl:value-of select="'#0000ff'"/>
          </xsl:when>
          <xsl:when test="$indexed='40'">
            <xsl:value-of select="'#00ccff'"/>
          </xsl:when>
          <xsl:when test="$indexed='41'">
            <xsl:value-of select="'#ccffff'"/>
          </xsl:when>
          <xsl:when test="$indexed='42'">
            <xsl:value-of select="'#ccffcc'"/>
          </xsl:when>
          <xsl:when test="$indexed='43'">
            <xsl:value-of select="'#ffff99'"/>
          </xsl:when>
          <xsl:when test="$indexed='44'">
            <xsl:value-of select="'#99ccff'"/>
          </xsl:when>
          <xsl:when test="$indexed='45'">
            <xsl:value-of select="'#ff99cc'"/>
          </xsl:when>
          <xsl:when test="$indexed='46'">
            <xsl:value-of select="'#cc99ff'"/>
          </xsl:when>
          <xsl:when test="$indexed='47'">
            <xsl:value-of select="'#ffcc99'"/>
          </xsl:when>
          <xsl:when test="$indexed='48'">
            <xsl:value-of select="'#3366ff'"/>
          </xsl:when>
          <xsl:when test="$indexed='49'">
            <xsl:value-of select="'#33cccc'"/>
          </xsl:when>
          <xsl:when test="$indexed='50'">
            <xsl:value-of select="'#99cc00'"/>
          </xsl:when>
          <xsl:when test="$indexed='51'">
            <xsl:value-of select="'#ffcc00'"/>
          </xsl:when>
          <xsl:when test="$indexed='52'">
            <xsl:value-of select="'#ff9900'"/>
          </xsl:when>
          <xsl:when test="$indexed='53'">
            <xsl:value-of select="'#ff6600'"/>
          </xsl:when>
          <xsl:when test="$indexed='54'">
            <xsl:value-of select="'#666699'"/>
          </xsl:when>
          <xsl:when test="$indexed='55'">
            <xsl:value-of select="'#969696'"/>
          </xsl:when>
          <xsl:when test="$indexed='56'">
            <xsl:value-of select="'#003366'"/>
          </xsl:when>
          <xsl:when test="$indexed='57'">
            <xsl:value-of select="'#339966'"/>
          </xsl:when>
          <xsl:when test="$indexed='58'">
            <xsl:value-of select="'#003300'"/>
          </xsl:when>
          <xsl:when test="$indexed='59'">
            <xsl:value-of select="'#333300'"/>
          </xsl:when>
          <xsl:when test="$indexed='60'">
            <xsl:value-of select="'#993300'"/>
          </xsl:when>
          <xsl:when test="$indexed='61'">
            <xsl:value-of select="'#993366'"/>
          </xsl:when>
          <xsl:when test="$indexed='62'">
            <xsl:value-of select="'#333399'"/>
          </xsl:when>
          <xsl:when test="$indexed='63'">
            <xsl:value-of select="'#333333'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'auto'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--2014-3-22 end-->
  
    <!--
       <xsl:template match="ws:si" mode="zitiji2">
    <xsl:variable name="pos" select="position()"/>
    <xsl:for-each select="ws:r">
      <xsl:variable name="pp">
        <xsl:value-of select="position()"/>
      </xsl:variable>
      <xsl:if test="ws:rPr">
        <xsl:if test="ws:rPr/ws:rFont">
          <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
            <xsl:attribute name="uof:标识符">
              <xsl:variable name="id" select="$pos"/>
              <xsl:value-of select="concat('font:',$id,$pp)"/>
            </xsl:attribute>
            <xsl:attribute name="uof:名称">
              <xsl:value-of select="ws:rPr/ws:rFont/@val"/>
            </xsl:attribute>
            <xsl:attribute name="uof:字体族">
              <xsl:value-of select="ws:rPr/ws:rFont/@val"/>
            </xsl:attribute>
          </uof:字体声明>
        </xsl:if>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
      <xsl:if test="/ws:spreadsheets/workbook.xml.rels/rel:Relationships/rel:Relationship[@Target='sharedStrings.xml']">
        <xsl:if test="/ws:spreadsheets/ws:sst/ws:si[ws:r]">
          <xsl:apply-templates select="/ws:spreadsheets/ws:sst/ws:si" mode="zitiji2"/>
        </xsl:if>
      </xsl:if>
  -->
</xsl:stylesheet>

<!--标题分类轴字体（默认字体引用的情况）2010.3.4-->
<!--
          <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea]">
            <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
              <xsl:attribute name="uof:标识符">
                <xsl:value-of select="concat($sstarget,'.btj_flz_e_font')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:名称">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="uof:字体族">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
              </xsl:attribute>
            </uof:字体声明>
          </xsl:if>
          <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin]">
            <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
              <xsl:attribute name="uof:标识符">
                <xsl:value-of select="concat($sstarget,'.btj_flz_l_font')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:名称">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="uof:字体族">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </uof:字体声明>
          </xsl:if>

          <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea]">
            <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
              <xsl:attribute name="uof:标识符">
                <xsl:value-of select="concat($sstarget,'.btj_flz_e_font')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:名称">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="uof:字体族">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea/@typeface"/>
              </xsl:attribute>
            </uof:字体声明>
          </xsl:if>
          <xsl:if test="not(ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea]) and (ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea])">
            <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
              <xsl:attribute name="uof:标识符">
                <xsl:value-of select="concat($sstarget,'.btj_flz_e_font')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:名称">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="uof:字体族">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea/@typeface"/>
              </xsl:attribute>
            </uof:字体声明>
          </xsl:if>
          <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin]">
            <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
              <xsl:attribute name="uof:标识符">
                <xsl:value-of select="concat($sstarget,'.btj_flz_l_font')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:名称">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="uof:字体族">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin/@typeface"/>
              </xsl:attribute>
            </uof:字体声明>
          </xsl:if>
          <xsl:if test="not(ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin]) and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin]">
            <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
              <xsl:attribute name="uof:标识符">
                <xsl:value-of select="concat($sstarget,'.btj_flz_l_font')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:名称">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="uof:字体族">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </uof:字体声明>
          </xsl:if>
-->
<!--标题数值轴字体（默认字体引用的情况）2010.3.4-->

          <!--<xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea]">
            <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
              <xsl:attribute name="uof:标识符">
                <xsl:value-of select="concat($sstarget,'.btj_szz_e_font')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:名称">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="uof:字体族">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
              </xsl:attribute>
            </uof:字体声明>
          </xsl:if>
          <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin]">
            <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
              <xsl:attribute name="uof:标识符">
                <xsl:value-of select="concat($sstarget,'.btj_szz_l_font')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:名称">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="uof:字体族">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </uof:字体声明>
          </xsl:if>

          <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea]">
            <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
              <xsl:attribute name="uof:标识符">
                <xsl:value-of select="concat($sstarget,'.btj_szz_e_font')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:名称">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="uof:字体族">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea/@typeface"/>
              </xsl:attribute>
            </uof:字体声明>
          </xsl:if>
          <xsl:if test="not(ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea]) and (ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea])">
            <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
              <xsl:attribute name="uof:标识符">
                <xsl:value-of select="concat($sstarget,'.btj_szz_e_font')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:名称">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="uof:字体族">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea/@typeface"/>
              </xsl:attribute>
            </uof:字体声明>
          </xsl:if>
          <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin]">
            <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
              <xsl:attribute name="uof:标识符">
                <xsl:value-of select="concat($sstarget,'.btj_szz_l_font')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:名称">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="uof:字体族">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin/@typeface"/>
              </xsl:attribute>
            </uof:字体声明>
          </xsl:if>
          <xsl:if test="not(ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin]) and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin]">
            <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
              <xsl:attribute name="uof:标识符">
                <xsl:value-of select="concat($sstarget,'.btj_szz_l_font')"/>
              </xsl:attribute>
              <xsl:attribute name="uof:名称">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
              <xsl:attribute name="uof:字体族">
                <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin/@typeface"/>
              </xsl:attribute>
            </uof:字体声明>
          </xsl:if>-->
      

<!--
      <xsl:for-each select="/ws:spreadsheets/ws:spreadsheet">
        <xsl:for-each select="ws:Drawings/xdr:wsDr/pr:Relationships/pr:Relationship">
          <xsl:variable name="target" select="@Target"/>
<xsl:if test="contains($target,'chart')">
  <xsl:variable name="targetq">
    <xsl:value-of select="substring-after($target,'..')"/>
  </xsl:variable>
  <xsl:variable name="starget">
    <xsl:value-of select="substring-after($target,'charts/')"/>
  </xsl:variable>
  <xsl:variable name="sstarget">
    <xsl:value-of select="substring-before($starget,'.')"/>
  </xsl:variable>
  <xsl:variable name="targetnew">
    <xsl:value-of select="concat('xlsx/xl',$targetq)"/>
  </xsl:variable>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:txPr/a:p/a:pPr/a:defRPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.tbq_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:txPr/a:p/a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.tbq_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.flz_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.flz_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.flz_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.flz_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.szz_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.szz_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.szz_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.szz_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>

  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:legend/c:legendEntry/c:txPr/a:p/a:pPr/a:defRPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.tuli_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:legendEntry/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:legendEntry/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:legend/c:legendEntry/c:txPr/a:p/a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.tuli_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:legendEntry/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:legendEntry/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:legend/c:txPr/a:p/a:pPr/a:defRPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.tuli_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:legend/c:txPr/a:p/a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.tuli_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend//c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:legend/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>

  <xsl:if test="./c:chart/c:plotArea/c:dTable/c:txPr/a:p/a:pPr/a:defRPr/a:ea">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.sjb_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:dTable/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:dTable/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:dTable/c:txPr/a:p/a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.sjb_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:dTable/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:dTable/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_bt_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_bt_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>

  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_bt_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="not(ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea]) and (ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea])">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_bt_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_bt_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="not(ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin]) and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_bt_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>

  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_flz_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_flz_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>

  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_flz_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="not(ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea]) and (ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea])">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_flz_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_flz_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="not(ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin]) and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_flz_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>

  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_szz_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_szz_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:txPr/a:p/a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>

  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_szz_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="not(ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:ea]) and (ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea])">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_szz_e_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:ea/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_szz_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
  <xsl:if test="not(ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/a:rPr/a:latin]) and ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')][c:chart/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin]">
    <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族">
      <xsl:attribute name="uof:标识符">
        <xsl:value-of select="concat($sstarget,'.btj_szz_l_font')"/>
      </xsl:attribute>
      <xsl:attribute name="uof:名称">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
      <xsl:attribute name="uof:字体族">
        <xsl:value-of select="ancestor-or-self::ws:spreadsheet/ws:Drawings/xdr:wsDr/c:chartSpace[@filename=substring-after($target,'charts/')]/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:r/preceding-sibling::a:pPr/a:defRPr/a:latin/@typeface"/>
      </xsl:attribute>
    </uof:字体声明>
  </xsl:if>
</xsl:if>
</xsl:for-each>
</xsl:for-each>


<xsl:if test="/ws:spreadsheets/pc:Types/*[contains(@PartName,'comments')]">
  <xsl:for-each select="/ws:spreadsheets/pc:Types/*[contains(@PartName,'comments')]">
    <xsl:variable name="pname" select="@PartName"/>
    <xsl:variable name="pname2" select="substring-after($pname,'/xl/')"/>
    <xsl:variable name="pname4" select="substring-before($pname2,'.')"/>
    <xsl:variable name="pname3" select="concat('xlsx',$pname)"/>
    <xsl:for-each select="./ancestor-or-self::ws:spreadsheet/ws:comments/ws:commentList/ws:comment/ws:comments/ws:commentList/ws:comment">
      <xsl:variable name="cpos">
        <xsl:number count="ws:comment" level="single"/>
      </xsl:variable>
      <xsl:for-each select="ws:text">
        <xsl:variable name="tpos">
          <xsl:number count="ws:text" level="single"/>
        </xsl:variable>
        <xsl:for-each select="ws:r">
          <xsl:variable name="rpos">
            <xsl:number count="ws:r" level="single"/>
          </xsl:variable>
          <uof:字体声明 uof:locID="u0041" uof:attrList="标识符 名称 字体族 替换字体">
            <xsl:attribute name="uof:标识符">
              <xsl:value-of select="concat($pname4,$cpos,$tpos,$rpos)"/>
            </xsl:attribute>
            <xsl:attribute name="uof:名称">
              <xsl:value-of select="ws:rPr/ws:rFont/@val"/>
            </xsl:attribute>
            <xsl:attribute name="uof:字体族">
              <xsl:value-of select="ws:rPr/ws:rFont/@val"/>
            </xsl:attribute>
          </uof:字体声明>
        </xsl:for-each>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:for-each>
</xsl:if>
-->