<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:fo="http://www.w3.org/1999/XSL/Format"
                xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
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
                xmlns:pzip="urn:u2o:xmlns:post-processings:special"
                xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">


  <xsl:template match="uof:单元格内容集合">
    <xsl:if test="*">
      <sst xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">
        <xsl:variable name="cnt" select="count(uof:单元格内容)"/>
        <xsl:attribute name="count">
          <xsl:value-of select="$cnt"/>
        </xsl:attribute>
        <xsl:attribute name="uniqueCount">
          <xsl:value-of select="$cnt"/>
        </xsl:attribute>
        <xsl:for-each select="./uof:单元格内容">
          
          <!--2014-6-3, update by Qihy, 一个单元格内字体式样不同的情况下，转换不正确，start-->
          <xsl:variable name="rowSeq" select="substring-before(@uof:cellID, '_')"/>
          <xsl:variable name="colSeq" select="substring-after(@uof:cellID, '_')"/>
          <xsl:variable name="sheetId" select="@uof:sheetId"/>
          <si>
            <xsl:for-each select="字:文本串_415B">
            <xsl:variable name="pos" select="position()"/>
            <xsl:if test="ancestor::uof:UOF/表:电子表格文档_E826/表:工作表集/表:单工作表/表:工作表_E825[@标识符_E7AC=$sheetId]/表:工作表内容_E80E/表:行_E7F1[@行号_E7F3=$rowSeq]/表:单元格_E7F2[@列号_E7BC=$colSeq]/表:数据_E7B3//字:句_419D/字:句属性_4158">
              <r>
              <xsl:for-each select="ancestor::uof:UOF/表:电子表格文档_E826/表:工作表集/表:单工作表/表:工作表_E825[@标识符_E7AC=$sheetId]/表:工作表内容_E80E/表:行_E7F1[@行号_E7F3=$rowSeq]/表:单元格_E7F2[@列号_E7BC=$colSeq]/表:数据_E7B3/字:句_419D[position()=$pos]">
                <xsl:if test="字:句属性_4158">
                  <rPr>
                    <xsl:apply-templates select="字:句属性_4158"/>
                  </rPr>
                </xsl:if>
                
                <!--2014-6-9, update by Qihy, PMG_EventBudget.xlsx数据丢失， start-->
                <xsl:if test="string-length(normalize-space(字:文本串_415B)) &lt;= 0">
                  <t xml:space="preserve"> </t>
                </xsl:if>
                <xsl:if test="string-length(normalize-space(字:文本串_415B)) &gt; 0">
                <t>
                  <xsl:value-of select="字:文本串_415B"/>
                </t>
                </xsl:if>
                <!--2014-6-9 end-->

              </xsl:for-each>
              </r>
            </xsl:if>
            <xsl:if test="not(ancestor::uof:UOF/表:电子表格文档_E826/表:工作表集/表:单工作表/表:工作表_E825[@标识符_E7AC=$sheetId]/表:工作表内容_E80E/表:行_E7F1[@行号_E7F3=$rowSeq]/表:单元格_E7F2[@列号_E7BC=$colSeq]/表:数据_E7B3//字:句_419D/字:句属性_4158)">
              <t>
                <xsl:value-of select="."/>
              </t>
            </xsl:if>
            </xsl:for-each>
          </si>
        </xsl:for-each>
      </sst>
    </xsl:if>
  </xsl:template>
  <xsl:template name="GetText">
    <xsl:param name="count"/>
    <xsl:param name="value"/>
    <xsl:param name="seq"/>
    <xsl:if test="number($seq) = $count">
      <xsl:value-of select="$value"/>
    </xsl:if>
    <xsl:if test="number($seq) &lt; number($count)">
      <xsl:call-template name="GetText">
        <xsl:with-param name="count" select="$count"/>
        <xsl:with-param name="seq" select="$seq + 1"/>
        <xsl:with-param name="value" select="concat($value,.//字:文本串_415B[number($seq) + 1])"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>
  <xsl:template match="字:句属性_4158">
    <xsl:if test="字:是否粗体_4130 and 字:是否粗体_4130/text()='true'">
      <b/>
    </xsl:if>
    <xsl:if test="字:字体_4128">
      <xsl:variable name="fnt" select="字:字体_4128/@中文字体引用_412A"/>
      <xsl:variable name="font" select="ancestor::uof:UOF/式样:式样集_990B/式样:字体集_990C/式样:字体声明_990D[@标识符_9902 = $fnt]/@名称_9903"/>
      <xsl:if test="字:字体_4128[@字号_412D]">
        <sz>
         <xsl:attribute name="val">
          <xsl:value-of select="字:字体_4128/@字号_412D"/>
         </xsl:attribute>
        </sz>
      </xsl:if>
      <xsl:if test="not(字:字体_4128[@字号_412D])">
        <sz val="11"/>
      </xsl:if>
      <rFont>
        <xsl:attribute name="val">
          <xsl:value-of select="$font"/>
        </xsl:attribute>
      </rFont>
      <family val="2"/>
      <!--<charset val="134"/>-->
    </xsl:if>
    <xsl:if test="字:是否斜体_4131 and 字:是否斜体_4131/text()='true'">
        <i/>
    </xsl:if>
    <xsl:if test="字:删除线_4135 and 字:删除线_4135/text() != 'none'">
      <strike/>
    </xsl:if>
    <xsl:if test="字:空心">
      <xsl:if test="字:空心[@值='true']">
        <outline/>
      </xsl:if>
    </xsl:if>
    <xsl:if test="字:阴影">
      <xsl:if test="字:阴影[@值='true']">
        <shadow/>
      </xsl:if>
    </xsl:if>
    <xsl:if test="字:字体_4128">
      <xsl:if test="字:字体_4128[@颜色_412F] and 字:字体_4128[@颜色_412F!='auto']">
        <xsl:variable name="color" select="字:字体_4128/@颜色_412F"/>
        <xsl:variable name="color2" select="substring-after($color,'#')"/>
        <xsl:variable name="color3" select="translate($color2,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLNOPQRSTUVWXYZ')"/>
        <xsl:variable name="color4" select="concat('FF',$color3)"/>
        <color>
          <xsl:attribute name="rgb">
            <xsl:if test="$color2!=''">
              <xsl:value-of select="$color4"/>
            </xsl:if>
          </xsl:attribute>
        </color>
      </xsl:if>
      <xsl:if test="not(字:字体_4128[@颜色_412F]) or substring-after(字:字体_4128/@颜色_412F,'#')=''">
        <color theme="1"/>
      </xsl:if>
    </xsl:if>
    <xsl:if test="字:下划线_4136 and 字:下划线_4136[@线型_4137= 'single']">
        <u/>
    </xsl:if>
    <xsl:if test="字:下划线_4136 and 字:下划线_4136[@线型_4137= 'double']">
      <u>
        <xsl:attribute name="val">
          <xsl:value-of select="'double'"/>
        </xsl:attribute>
      </u>
    </xsl:if>
    <xsl:if test="字:上下标类型_4143/text()='sub'">
      <vertAlign val="subscript"/>
    </xsl:if>
    <xsl:if test="字:上下标类型_4143/text()='sup'">
      <vertAlign val="superscript"/>
    </xsl:if>
    <!--<scheme val="minor"/>-->
  </xsl:template>
  <!--2014-6-3 end-->
  
</xsl:stylesheet>

<!--Need Consideration Codes-->
<!--Marked by LDM in 2011/01/21-->
<!--
          <phoneticPr fontId="1" type="noConversion"/>
<xsl:template match="uof:所有的句ssi">
  <sst xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">
    <xsl:variable name="cnt" select="count(uof:si的编号)"/>
    <xsl:attribute name="count">
      <xsl:value-of select="$cnt"/>
    </xsl:attribute>
    <xsl:attribute name="uniqueCount">
      <xsl:value-of select="$cnt"/>
    </xsl:attribute>
    <xsl:for-each select="uof:si的编号">
      <si>
        <xsl:if test="uof:单元格里的句si[字:句]">
          <xsl:for-each select="uof:单元格里的句si/字:句">
            <xsl:if test="not(字:句属性) and 字:文本串">
              <t>
                <xsl:value-of select="字:文本串"/>
              </t>
            </xsl:if>
            <xsl:if test="字:句属性 and 字:文本串">
              <r>
                <rPr>
                  <xsl:apply-templates select="字:句属性"/>
                </rPr>
                <t>
                  <xsl:value-of select="字:文本串"/>
                </t>
              </r>
            </xsl:if>
          </xsl:for-each>
        </xsl:if>
        <phoneticPr fontId="1" type="noConversion"/>
      </si>
    </xsl:for-each>
  </sst>
</xsl:template>
-->

<!--Need Consideration-->
<!--Not Finished-->
<!--Marked by LDM in 2010/12/19-->
<!--
          <xsl:if test="uof:单元格里的句si[字:句]">
            <xsl:for-each select="uof:单元格里的句si/字:句">
              <xsl:if test="not(字:句属性) and 字:文本串">
                <t>
                  <xsl:value-of select="字:文本串"/>
                </t>
              </xsl:if>
              <xsl:if test="字:句属性 and 字:文本串">
                <r>
                  <rPr>
                    <xsl:apply-templates select="字:句属性"/>
                  </rPr>
                  <t>
                    <xsl:value-of select="字:文本串"/>
                  </t>
                </r>
              </xsl:if>
            </xsl:for-each>
          </xsl:if>
          -->