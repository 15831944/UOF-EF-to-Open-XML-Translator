<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet
  version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:uof="http://schemas.uof.org/cn/2009/uof"
  xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
  xmlns:演="http://schemas.uof.org/cn/2009/presentation"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:图="http://schemas.uof.org/cn/2009/graph"
  xmlns:超链="http://schemas.uof.org/cn/2009/hyperlinks"
  
  xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
  xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
  xmlns:ws="http://purl.oclc.org/ooxml/spreadsheetml/main"
  
  xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">

  <!--Hyperlinks Template-->
  <xsl:template name="Hyperlinks">
    <xsl:if test="./ws:spreadsheets/ws:spreadsheet/ws:worksheet/ws:hyperlinks/ws:hyperlink">
      <超链:链接集_AA0B>
        <!--修改 李杨2012-3-16-->
        <xsl:for-each select="ws:spreadsheets/ws:spreadsheet">
          <!-- update by xuzhenwei BUG_2479:超级链接式样设置失效 start 暂时没起作用，需要补充完善 -->
          <xsl:variable name="rList">
            <xsl:for-each select="./ws:worksheet/ws:sheetData/ws:row">
              <xsl:for-each select="//c">
                <xsl:value-of select="concat(@r,'xxx',@s,' ')"/>
              </xsl:for-each>
            </xsl:for-each>
          </xsl:variable>
          <xsl:for-each select ="ws:worksheet/ws:hyperlinks/ws:hyperlink">
            <xsl:variable name ="sheetName">
              <xsl:value-of select ="../../../@sheetName"/>
            </xsl:variable>
            <xsl:variable name="seq" select="position()"/>
            <xsl:if test="@r:id">
              <xsl:variable name="refId" select="@r:id"/>
              <超链:超级链接_AA0C>
                <xsl:attribute name="标识符_AA0A">
                  <xsl:value-of select="concat($sheetName,'_Hyperlink_',$seq)"/>
                </xsl:attribute>
                <超链:链源_AA00>
                  <xsl:value-of select="concat('hlkref_',generate-id(.))"/>
                </超链:链源_AA00>
                <xsl:if test="ancestor::ws:worksheet/following-sibling::pr:Relationships/pr:Relationship/@Id=$refId">
                  <超链:目标_AA01>
                    <xsl:value-of select="ancestor::ws:worksheet/following-sibling::pr:Relationships/pr:Relationship[@Id=$refId]/@Target"/>
                  </超链:目标_AA01>
                </xsl:if>
                <xsl:if test="contains($rList,@ref)">
                  <超链:式样_AA02>
                    <xsl:variable name="bf" select="substring-after($rList,@ref+'xxx')"/>
                    <xsl:variable name="styleposition" select="substring-before($bf,' ')"/>
                    <xsl:attribute name="未访问式样引用_AA03">
                      <xsl:value-of select="concat('CELLSTYLE_',$styleposition)"/>
                    </xsl:attribute>
                    <xsl:attribute name="已访问式样引用_AA04">
                      <xsl:value-of select="concat('CELLSTYLE_',$styleposition)"/>
                    </xsl:attribute>
                  </超链:式样_AA02>
                </xsl:if>
                <xsl:if test="@tooltip">
                  <超链:提示_AA05>
                    <xsl:value-of select="@tooltip"/>
                  </超链:提示_AA05>
                </xsl:if>
              </超链:超级链接_AA0C>
            </xsl:if>
            <xsl:if test="@location">
              <xsl:choose>
                <xsl:when test="contains(@location,'书签')">
                  <超链:超级链接_AA0C>
                    <xsl:variable name="bookmark">
                      <xsl:value-of select="@location"/>
                    </xsl:variable>
                    <xsl:attribute name="标识符_AA0A">
                      <xsl:value-of select="concat('ID_Hyperlink_',$seq)"/>
                    </xsl:attribute>
                    <超链:链源_AA00>
                      <xsl:value-of select="concat('hlkref_',generate-id(.))"/>
                    </超链:链源_AA00>
                    <超链:目标_AA01>
                      <xsl:variable name="targetTemp">
                        <xsl:if test="//ws:workbook/ws:definedNames/ws:definedName[@name=$bookmark]">
                          <xsl:value-of select="translate(//ws:workbook/ws:definedNames/ws:definedName[@name=$bookmark],'$','')"/>
                        </xsl:if>
                        <xsl:if test="not(//ws:workbook/ws:definedNames/ws:definedName[@name=$bookmark])">
                          <xsl:value-of select="'Sheet2!D28'"/>
                        </xsl:if>
                      </xsl:variable>
                      <xsl:value-of  select="$targetTemp"/>
                    </超链:目标_AA01>
                    <xsl:if test="contains($rList,@ref)">
                      <超链:式样_AA02>
                        <xsl:variable name="bf" select="substring-after($rList,@ref+'xxx')"/>
                        <xsl:variable name="styleposition" select="substring-before($bf,' ')"/>
                        <xsl:attribute name="未访问式样引用_AA03">
                          <xsl:value-of select="concat('CELLSTYLE_',$styleposition)"/>
                        </xsl:attribute>
                        <xsl:attribute name="已访问式样引用_AA04">
                          <xsl:value-of select="concat('CELLSTYLE_',$styleposition)"/>
                        </xsl:attribute>
                      </超链:式样_AA02>
                    </xsl:if>
                    <xsl:if  test="@tooltip">
                      <超链:提示_AA05>
                        <xsl:value-of  select="@tooltip"/>
                      </超链:提示_AA05>
                    </xsl:if>
                  </超链:超级链接_AA0C>
                </xsl:when>
                <xsl:otherwise>
                  <超链:超级链接_AA0C>
                    <xsl:attribute name="标识符_AA0A">
                      <xsl:value-of select="concat($sheetName,'_Hyperlink_',$seq)"/>
                    </xsl:attribute>
                    <超链:链源_AA00>
                      <xsl:value-of select="concat('hlkref_',generate-id(.))"/>
                    </超链:链源_AA00>
                    <超链:目标_AA01>
                      <xsl:variable name="targetTemp">
                        <xsl:value-of select="ancestor::ws:worksheet/following-sibling::pr:Relationships/pr:Relationship[@Id=$seq]/@Target"/>
                      </xsl:variable>
                      
                      <!--2013-11-12 Qihy 超链接-链接到本地文档已定义名称时丢失-->
                      <xsl:variable name ="targetTemp1" select ="@location"/>
                      <xsl:variable name="locationTemp">
                         <xsl:choose>
                          <xsl:when test ="contains($targetTemp1, '!')">
                            <xsl:value-of select="$targetTemp1"/>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:value-of select ="//ws:workbook/ws:definedNames/ws:definedName[@name = $targetTemp1]"/>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:variable>
                      <!--2013-11-12 Qihy 超链接-链接到本地文档已定义名称时丢失-->
                      
                      <xsl:if test="contains($targetTemp,'file:///')">
                        <xsl:value-of select="concat(translate($targetTemp,'file:///',''),'/',$locationTemp)"/>
                      </xsl:if>
                      <xsl:if test="not(contains($targetTemp,'file:///'))">
                        <xsl:value-of select="concat('',$locationTemp)"/>
                      </xsl:if>
                    </超链:目标_AA01>
                    <超链:式样_AA02>
                    <xsl:if test="contains($rList,@ref)">
                        <xsl:variable name="bf" select="substring-after($rList,@ref+'xxx')"/>
                        <xsl:variable name="styleposition" select="substring-before($bf,' ')"/>
                        <xsl:attribute name="未访问式样引用_AA03">
                          <xsl:value-of select="concat('CELLSTYLE_',$styleposition)"/>
                        </xsl:attribute>
                        <xsl:attribute name="已访问式样引用_AA04">
                          <xsl:value-of select="concat('CELLSTYLE_',$styleposition)"/>
                        </xsl:attribute>
                    </xsl:if>
                    </超链:式样_AA02>
                    <!-- end -->
                    <xsl:if  test="@tooltip">
                      <超链:提示_AA05>
                        <xsl:value-of  select="@tooltip"/>
                      </超链:提示_AA05>
                    </xsl:if>
                  </超链:超级链接_AA0C>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:if>
          </xsl:for-each>
        </xsl:for-each>
      </超链:链接集_AA0B>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>
