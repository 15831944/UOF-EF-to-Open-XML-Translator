<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:uof="http://schemas.uof.org/cn/2009/uof"
                xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
                xmlns:演="http://schemas.uof.org/cn/2009/presentation"
                xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
                xmlns:图="http://schemas.uof.org/cn/2009/graph"
                
                xmlns:ws="http://purl.oclc.org/ooxml/spreadsheetml/main"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
	<xsl:template name="sheetfilter">
		<xsl:if test="ws:worksheet/ws:autoFilter">
      <表:筛选集_E83A>
        <表:筛选_E80F>
          <xsl:choose>
            <xsl:when test="ws:worksheet/ws:autoFilter/ws:filterColumn/ws:customFilters">
              <xsl:attribute name="类型_E83B">auto</xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="类型_E83B">auto</xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:apply-templates select="ws:worksheet/ws:autoFilter"/>
        </表:筛选_E80F>
      </表:筛选集_E83A>
		</xsl:if>
	</xsl:template>
  <xsl:template name ="ColIndex_C2N">
    <xsl:param name="colChar"/>
    <xsl:variable name="colNum">
      <xsl:value-of select="translate($colChar,'ABCDEFGHI','123456789')"/>
    </xsl:variable>
    <xsl:value-of select="$colNum"/>
  </xsl:template>
	<xsl:template match="ws:autoFilter">
    <xsl:variable name="ref">
      <xsl:value-of select="./@ref"/>
    </xsl:variable>
    <表:范围_E810>
      <xsl:value-of select="$ref"/>
    </表:范围_E810>
		<xsl:if test="ws:filterColumn or ws:sortState">
			<xsl:for-each select="ws:filterColumn">

        <表:条件_E811>
          <xsl:attribute name="列号_E819">
            <!--<xsl:variable name="colNum">
              <xsl:call-template name="ColIndex_C2N">
                <xsl:with-param name="colChar" select="substring($ref,1,1)"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:value-of select="$colNum"/>-->
            <xsl:value-of select ="@colId +1"/>
          </xsl:attribute>

          <xsl:choose>
            <!--<xsl:when test ="../sortState/sortCondition/@sorBy = 'cellColor'">-->
              <xsl:when test="ws:colorFilter">
              <表:颜色_E818>
                <xsl:variable name="colorId">
                  <xsl:value-of select="./ws:colorFilter/@dxfId"/>
                  <!--<xsl:value-of select ="../sortState/sortCondition/@dxfId"/>-->
                </xsl:variable>
                <xsl:variable name="colorValue">
                  <xsl:value-of select="./ancestor::ws:spreadsheets/ws:styleSheet/ws:dxfs/ws:dxf[position() = number($colorId) + 1]/ws:fill/ws:patternFill/ws:fgColor/@rgb"/>
                </xsl:variable>
                <xsl:value-of select="concat('#',substring-after($colorValue,'FF'))"/>
              </表:颜色_E818>
            </xsl:when>
            <!-- 20130518 add by xuzhenwei 修改数据过滤，start -->
            <!-- 按val值进行过滤 -->
            <xsl:when test="ws:filters/ws:filter">
               <xsl:for-each select="ws:filters/ws:filter">
                 <表:普通_E812>
                   <xsl:attribute name="类型_E7B6">
                     <xsl:value-of select="'value'"/>
                   </xsl:attribute>
                   <xsl:attribute name="值_E813">
                     <xsl:value-of select="@val"/>
                   </xsl:attribute>
                 </表:普通_E812>
               </xsl:for-each>
            </xsl:when>
            <!-- 数据过滤，top排序 -->
            <xsl:when test="ws:top10">
              <表:普通_E812>
                <xsl:attribute name="类型_E7B6">
                  <xsl:value-of select="'top-item'"/>
                </xsl:attribute>
                <xsl:attribute name="值_E813">
                  <xsl:value-of select="ws:top10/@val"/>
                </xsl:attribute>
              </表:普通_E812>
            </xsl:when>
            <!-- end -->
            <xsl:otherwise>
              <表:自定义_E814>
                <xsl:attribute name="类型_E83C">
                  <xsl:choose>
                    <xsl:when test ="ws:customFilters/@and and ws:customFilters/@and='1'">
                      <xsl:value-of select ="'and'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="'or'"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
                <xsl:for-each select="ws:filters/ws:filter">
                  <xsl:if test ="position()=1 or position()=2">
                  <表:操作条件_E815>
                    <表:操作码_E816>equal-to</表:操作码_E816>
                    <表:值_E817>
                      <xsl:value-of select ="@val"/>
                    </表:值_E817>
                  </表:操作条件_E815>
                  </xsl:if>
                </xsl:for-each>
                <xsl:for-each select="ws:customFilters/ws:customFilter">
                  <表:操作条件_E815>
                    <表:操作码_E816>
                      <xsl:variable name="op">
                        <xsl:if test="@operator = 'equal'">
                          <xsl:value-of select="'equal-to'"/>
                        </xsl:if>
                        <!--<xsl:if test="not(@operator)">
                          <xsl:value-of select="'not-equal-to'"/>
                        </xsl:if>-->
                        <xsl:if test="@operator = 'greaterThan'">
                          <xsl:value-of select="'greater-than'"/>
                        </xsl:if>
                        <xsl:if test="@operator = 'greaterThanOrEqual'">
                          <xsl:value-of select="'greater-than-or-equal-to'"/>
                        </xsl:if>
                        <xsl:if test="@operator = 'lessThan'">
                          <xsl:value-of select="'less-than'"/>
                        </xsl:if>
                        <xsl:if test="@operator = 'lessThanOrEqual'">
                          <xsl:value-of select="'less-than-or-equal-to'"/>
                        </xsl:if>
                        <xsl:if test="@operator = 'notEqual'">
                          <xsl:value-of select="'not-equal-to'"/>
                        </xsl:if>
                        <xsl:choose>
                          <xsl:when test="starts-with(@val,'*')">
                            <xsl:variable name="cut">
                              <xsl:value-of select="substring-after(@val,'*')"/>
                            </xsl:variable>
                            <xsl:if test="contains($cut,'*')">
                              <xsl:if test="not(@operator)">
                                <xsl:value-of select="'contain'"/>
                              </xsl:if>
                              <xsl:if test="@operator = 'notEqual'">
                                <xsl:value-of select="'not-contain'"/>
                              </xsl:if>
                            </xsl:if>
                            <xsl:if test="not(contains($cut,'*'))">
                              <xsl:if test="not(@operator)">
                                <xsl:value-of select="'end-with'"/>
                              </xsl:if>
                              <xsl:if test="@operator = 'notEqual'">
                                <xsl:value-of select="'not-end-with'"/>
                              </xsl:if>
                            </xsl:if>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:if test="@operator and contains(@val,'*')">
                              <xsl:value-of select="'not-start-with'"/>
                            </xsl:if>
                            <xsl:if test="not(@operator) and contains(@val,'*')">
                              <xsl:value-of select="'start-with'"/>
                            </xsl:if>
                          </xsl:otherwise>
                        </xsl:choose>
                      </xsl:variable>
                      <xsl:value-of select="$op"/>
                    </表:操作码_E816>
                    <表:值_E817>
                      <xsl:value-of select="translate(@val,'*','')"/>
                    </表:值_E817>
                  </表:操作条件_E815>
                </xsl:for-each>
              </表:自定义_E814>
            </xsl:otherwise>
          </xsl:choose>

        </表:条件_E811>
			</xsl:for-each>

      <!--<xsl:for-each select ="ws:sortState">
        <表:条件_E811>
          <xsl:attribute name="列号_E819">
            <xsl:value-of select ="@colId +1"/>
          </xsl:attribute>
          <xsl:if test ="../sortState/sortCondition/@sorBy = 'cellColor'">
            <表:颜色_E818>
              <xsl:variable name="colorId">
                --><!--<xsl:value-of select="./ws:colorFilter/@dxfId"/>--><!--
                <xsl:value-of select ="../sortState/sortCondition/@dxfId"/>
              </xsl:variable>
              <xsl:variable name="colorValue">
                <xsl:value-of select="./ancestor::ws:spreadsheets/ws:styleSheet/ws:dxfs/ws:dxf[position() = number($colorId) + 1]/ws:fill/ws:patternFill/ws:fgColor/@rgb"/>
              </xsl:variable>
              <xsl:value-of select="substring-after($colorValue,'FF')"/>
            </表:颜色_E818>
          </xsl:if>
        </表:条件_E811>
      </xsl:for-each>-->
		</xsl:if>
	</xsl:template>
</xsl:stylesheet>
