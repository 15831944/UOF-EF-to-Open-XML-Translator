<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:字="http://schemas.uof.org/cn/2003/uof-wordproc" xmlns:uof="http://schemas.uof.org/cn/2003/uof" xmlns:图="http://schemas.uof.org/cn/2003/graph" xmlns:表="http://schemas.uof.org/cn/2003/uof-spreadsheet" xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <xsl:template name="contentTypes">
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>

      <!--picture-->
      <xsl:if test="//uof:其他对象[@uof:公共类型='jpg'] or //uof:其他对象[@uof:私有类型]">
        <Default Extension="jpeg" ContentType="image/jpeg"/>
      </xsl:if>
      <xsl:if test="//uof:其他对象[@uof:公共类型='jpeg']">
        <Default Extension="jpeg" ContentType="image/jpeg"/>
      </xsl:if>
      <xsl:if test="//uof:其他对象[@uof:公共类型='png']">
        <Default Extension="jpeg" ContentType="image/jpeg"/>
      </xsl:if>
      <xsl:if test="//uof:其他对象[@uof:公共类型='bmp']">
        <Default Extension="png" ContentType="image/png"/>
      </xsl:if>
      <xsl:if test="//uof:其他对象[@uof:公共类型='gif']">
        <Default Extension="gif" ContentType="image/gif"/>
      </xsl:if>
      <xsl:if test="//uof:其他对象[@uof:公共类型='wmf']">
        <Default Extension="wmf" ContentType="image/x-wmf"/>
      </xsl:if>

      <!--styles-->
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      <!--theme-->
      <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
      <!--matadata-->
      <xsl:if test="uof:元数据/uof:编辑时间|uof:元数据/uof:创建应用程序|uof:元数据/uof:文档模板|uof:元数据/uof:公司名称|uof:元数据/uof:经理名称|uof:元数据/uof:页数|uof:元数据/uof:字数|uof:元数据/uof:行数|uof:元数据/uof:段落数">
        <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
      </xsl:if>
      <xsl:if test="uof:元数据/uof:标题|uof:元数据/uof:主题|uof:元数据/uof:创建者|uof:元数据/uof:最后作者|uof:元数据/uof:摘要|uof:元数据/uof:创建日期|uof:元数据/uof:编辑次数|uof:元数据/uof:分类|uof:元数据/uof:关键字集">
        <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
      </xsl:if>
      <xsl:if test="uof:元数据/uof:用户自定义元数据集">
        <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
      </xsl:if>
      <!--workbook-->
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <xsl:for-each select="uof:电子表格/表:工作表集/表:单工作表">
        <Override>
          <xsl:attribute name="PartName">
            <xsl:value-of select="concat('/xl/worksheets/sheet',@uof:sheetNo,'.xml')"/>
          </xsl:attribute>
          <xsl:attribute name="ContentType">
            <xsl:value-of select="'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'"/>
          </xsl:attribute>
        </Override>
      </xsl:for-each>
    </Types>
  </xsl:template>
</xsl:stylesheet>
