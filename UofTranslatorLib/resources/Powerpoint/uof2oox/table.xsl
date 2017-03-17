<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:pzip="urn:u2o:xmlns:post-processings:special" 
				xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
				xmlns:fo="http://www.w3.org/1999/XSL/Format"
				xmlns:app="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" 
				xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
				xmlns:dc="http://purl.org/dc/elements/1.1/" 
				xmlns:dcterms="http://purl.org/dc/terms/" 
				xmlns:dcmitype="http://purl.org/dc/dcmitype/"
				xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
				xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
				xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" 
				xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
 xmlns:uof="http://schemas.uof.org/cn/2009/uof"
xmlns:表="http://schemas.uof.org/cn/2009/spreadsheet"
xmlns:演="http://schemas.uof.org/cn/2009/presentation"
xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
xmlns:图="http://schemas.uof.org/cn/2009/graph"
xmlns:规则="http://schemas.uof.org/cn/2009/rules">
  <!--start liuyin 20130110 修改幻灯片配色方案效果丢失 添加命名空间   xmlns:规则="http://schemas.uof.org/cn/2009/rules"-->
  
  
  <xsl:import href="txBody.xsl"/>
  <xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
  <xsl:template name="table">
    <xsl:param name="tblobj"/>
    <p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr>
          <xsl:attribute name="id">
            <xsl:value-of select="substring($tblobj,4,5)"/>
          </xsl:attribute>
          <xsl:attribute name="name">表格</xsl:attribute>
            <!--<xsl:value-of select="表格"/>-->
          
        </p:cNvPr>
        <p:cNvGraphicFramePr>
          <a:graphicFrameLocks noGrp="1"/>
        </p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <xsl:if test="@图形引用_C62E=$tblobj">
        <xsl:call-template name="xfrm"/>
      </xsl:if>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <xsl:for-each select=".//图:内容_8043/字:文字表_416C">
            <a:tbl>
             <xsl:apply-templates select="字:文字表属性_41CC"/>              
              <xsl:call-template name="tblGrid"/>
              <xsl:call-template name="tblRow1">
                <xsl:with-param name="totalColCount">
                  <xsl:value-of select="count(./字:文字表属性_41CC/字:列宽集_41C1/字:列宽_41C2)"/>
                </xsl:with-param>
                <xsl:with-param name="totalRowCount" select="count(./字:行_41CD)"/>
                <xsl:with-param name="Rows" select="./child::字:行_41CD"/>
                <xsl:with-param name="curRow" select="1"/>
                <xsl:with-param name ="vmergestr" select="''"/>
              </xsl:call-template>
            </a:tbl>
          </xsl:for-each>
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>
  </xsl:template>
  <xsl:template match="字:文字表属性_41CC">
    <a:tblPr firstRow="1" bandRow="1">      
      <xsl:if test="字:填充_4134">        
         <!--<xsl:call-template name="fill"/>-->
       
      </xsl:if>
         <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
    </a:tblPr>
  </xsl:template>
  <xsl:template name="tblGrid">
    <a:tblGrid>
      <xsl:for-each select="字:文字表属性_41CC/字:列宽集_41C1/字:列宽_41C2">
        <a:gridCol>
          <xsl:attribute name="w">
            
            <!--start liuyin 20130112 修改文字表高度和列宽不正确-->
            <!--<xsl:value-of select="round(. * 24)"/>-->
            <xsl:value-of select="round(. * 12698)"/>
            <!--end liuyin 20130112 修改文字表高度和列宽不正确-->
            
          </xsl:attribute>
        </a:gridCol>
      </xsl:for-each>
    </a:tblGrid>
  </xsl:template>
  <!--liuyangyang 2015-03-24 修复表格跨行转换问题 start-->
  <xsl:template name="getPreVms">
    <xsl:param name="colCount"/>
    <xsl:param name="rowCell"/>
    <xsl:param name="vmergestr1"/>
    <xsl:param name="vmergestr2"/>
    <xsl:param name="curCol"/>
    <xsl:param name="curCell"/>
    <xsl:param name="hmergeCount"/>
    <xsl:choose>
      <xsl:when test="$curCol &lt;=$colCount">
        <xsl:variable name="nextVMerge" select="number(substring-before($vmergestr1,'*'))"/>
        <xsl:variable name="vmergeCount" select="number(substring-before(substring-after($vmergestr1,'*'),'*'))"/>
        <xsl:choose>
          <xsl:when test="$curCol=$nextVMerge">
            
            
            <xsl:call-template name="getPreVms">
              <xsl:with-param name="colCount" select="$colCount"/>
              <xsl:with-param name="rowCell" select="$rowCell"/>
              <xsl:with-param name="vmergestr1" select="substring-after(substring-after($vmergestr1,'*'),'*')"/>
              <xsl:with-param name="vmergestr2">
                <xsl:if test="$vmergeCount=1">
                  <xsl:value-of select="$vmergestr2"/>
                </xsl:if>
                <xsl:if test="$vmergeCount &gt;1">
                  <xsl:value-of select="concat($vmergestr2,$nextVMerge,'*',$vmergeCount - 1,'*')"/>
                </xsl:if>
              </xsl:with-param>
              <xsl:with-param name="curCol" select="$curCol + 1"/>
              <xsl:with-param name="curCell" select="$curCell"/>
              <xsl:with-param name="hmergeCount" select="$hmergeCount - 1"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$hmergeCount &gt; 1 and $curCol != $nextVMerge">
            <xsl:call-template name="getPreVms">
              <xsl:with-param name="colCount" select="$colCount"/>
              <xsl:with-param name="rowCell" select="$rowCell"/>
              <xsl:with-param name="vmergestr1" select="$vmergestr1"/>
              <xsl:with-param name="vmergestr2">
                <xsl:value-of select="$vmergestr2"/>
              </xsl:with-param>
              <xsl:with-param name="curCol" select="$curCol + 1"/>
              <xsl:with-param name="curCell" select="$curCell"/>
              <xsl:with-param name="hmergeCount" select="$hmergeCount - 1"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:for-each select="$rowCell[$curCell]">
              <xsl:call-template name="getPreVms">
                <xsl:with-param name="colCount" select="$colCount"/>
                <xsl:with-param name="rowCell" select="$rowCell"/>
                <xsl:with-param name="vmergestr1" select="$vmergestr1"/>
                <xsl:with-param name="vmergestr2">
                  <xsl:choose>
                    <xsl:when test="字:单元格属性_41B7/字:跨行_41A6 and not(字:单元格属性_41B7/字:跨列_41A7)">
                      <xsl:value-of select="concat($vmergestr2,$curCol,'*',字:单元格属性_41B7/字:跨行_41A6,'*')"/>
                    </xsl:when>
                    <xsl:when test="字:单元格属性_41B7/字:跨行_41A6 and 字:单元格属性_41B7/字:跨列_41A7">
                      <xsl:call-template name="getHVM">
                        <xsl:with-param name="curCol" select="$curCol"/>
                        <xsl:with-param name="vmergeCount" select="number(字:单元格属性_41B7/字:跨行_41A6)"/>
                        <xsl:with-param name="hmergeCount" select="number(字:单元格属性_41B7/字:跨列_41A7)"/>
                        <xsl:with-param name="vmergeStr" select="$vmergestr2"/>
                      </xsl:call-template>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$vmergestr2"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
                <xsl:with-param name="curCol" select="$curCol + 1"/>
                <xsl:with-param name="curCell" select="$curCell+ 1"/>
                <xsl:with-param name="hmergeCount">
                  <xsl:if test="字:单元格属性_41B7/字:跨列_41A7">
                    <xsl:value-of select="number(字:单元格属性_41B7/字:跨列_41A7)"/>
                  </xsl:if>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$vmergestr2"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="transtblCell">
    <xsl:param name="colCount"/>
    <xsl:param name="rowCell"/>
    <xsl:param name="vmergestr1"/>
    <xsl:param name="curCol"/>
    <xsl:param name="curCell"/>
    <xsl:param name="hmergeCount"/>
    <xsl:param name="preRowSpan"/>
    <xsl:choose>
      <xsl:when test="$curCol &lt;=$colCount">
        <xsl:variable name="nextVMerge" select="number(substring-before($vmergestr1,'*'))"/>
        <xsl:variable name="vmergeCount" select="number(substring-before(substring-after($vmergestr1,'*'),'*'))"/>
        <xsl:choose>
          <xsl:when test="$curCol=$nextVMerge">
            <xsl:choose>
              <xsl:when test="$hmergeCount and $hmergeCount &gt; 1">
                <xsl:call-template name="transtblMergeCellContent">
                  <xsl:with-param name="vmerge" select="'yes'"/>
                  <xsl:with-param name="hmerge" select="'yes'"/>
                  <xsl:with-param name="rowSpan" select="0"/>
                </xsl:call-template>
              </xsl:when>
              <xsl:otherwise>
                <xsl:call-template name="transtblMergeCellContent">
                  <xsl:with-param name="vmerge" select="'yes'"/>
                  <xsl:with-param name="hmerge" select="'no'"/>
                  <xsl:with-param name="rowSpan" select="0"/>
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
            <xsl:call-template name="transtblCell">
              <xsl:with-param name="colCount" select="$colCount"/>
              <xsl:with-param name="rowCell" select="$rowCell"/>
              <xsl:with-param name="vmergestr1" select="substring-after(substring-after($vmergestr1,'*'),'*')"/>
              <xsl:with-param name="curCol" select="$curCol + 1"/>
              <xsl:with-param name="curCell" select="$curCell"/>
              <xsl:with-param name="hmergeCount" select="$hmergeCount - 1"/>
              <xsl:with-param name="preRowSpan" select ="0"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$hmergeCount &gt; 1 and $curCol != $nextVMerge">
            <xsl:call-template name="transtblMergeCellContent">
              <xsl:with-param name="vmerge" select="'no'"/>
              <xsl:with-param name="hmerge" select="'yes'"/>
              <xsl:with-param name="rowSpan">
                <xsl:if test="not($preRowSpan) or $preRowSpan&lt; 1 ">
                  <xsl:value-of select="0"/>
                </xsl:if>
                <xsl:if test="$preRowSpan &gt; 1">
                  <xsl:value-of select="$preRowSpan"/>
                </xsl:if>
              </xsl:with-param>
            </xsl:call-template>
            <xsl:call-template name="transtblCell">
              <xsl:with-param name="colCount" select="$colCount"/>
              <xsl:with-param name="rowCell" select="$rowCell"/>
              <xsl:with-param name="vmergestr1" select="$vmergestr1"/>
              <xsl:with-param name="curCol" select="$curCol + 1"/>
              <xsl:with-param name="curCell" select="$curCell"/>
              <xsl:with-param name="hmergeCount" select="$hmergeCount - 1"/>
              <xsl:with-param name="preRowSpan" select ="$preRowSpan"/>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:for-each select="$rowCell[$curCell]">
              <xsl:call-template name="cell"/>
              <xsl:call-template name="transtblCell">
                <xsl:with-param name="colCount" select="$colCount"/>
                <xsl:with-param name="rowCell" select="$rowCell"/>
                <xsl:with-param name="vmergestr1" select="$vmergestr1"/>
                <xsl:with-param name="curCol" select="$curCol + 1"/>
                <xsl:with-param name="curCell" select="$curCell+ 1"/>
                <xsl:with-param name="hmergeCount">
                  <xsl:choose>
                    <xsl:when test="字:单元格属性_41B7/字:跨列_41A7">
                      <xsl:value-of select="number(字:单元格属性_41B7/字:跨列_41A7)"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="0"/>
                    </xsl:otherwise>
                  </xsl:choose>
                  
                </xsl:with-param>
                <xsl:with-param name="preRowSpan">
                  <xsl:choose>
                    <xsl:when test="字:单元格属性_41B7/字:跨行_41A6">
                      <xsl:value-of select="number(字:单元格属性_41B7/字:跨行_41A6)"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="0"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:for-each>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  
  <xsl:template name="getHVM">
    <xsl:param name="curCol"/>
    <xsl:param name="vmergeCount"/>
    <xsl:param name="hmergeCount"/>
    <xsl:param name="vmergeStr"/>
    <xsl:choose>
      <xsl:when test="$hmergeCount &gt; 0">
        <xsl:call-template name="getHVM">
          <xsl:with-param name="curCol" select="$curCol +1"/>
          <xsl:with-param name="vmergeCount" select="$vmergeCount"/>
          <xsl:with-param name="hmergeCount" select="$hmergeCount - 1"/>
          <xsl:with-param name="vmergeStr" select="concat($vmergeStr,$curCol,'*',$vmergeCount,'*')"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select ="$vmergeStr"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  
  
  <xsl:template name="transtblMergeCellContent">
    <xsl:param name="vmerge"/>
    <xsl:param name="hmerge"/>
    <xsl:param name="rowSpan"/>
    <a:tc>
      <xsl:if test="$rowSpan &gt;1">
        <xsl:attribute name="rowSpan">
          <xsl:value-of select="$rowSpan"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="$vmerge = 'yes'">
        <xsl:attribute name="vMerge">
          <xsl:value-of select="'1'"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="$hmerge = 'yes'">
        <xsl:attribute name="hMerge">
          <xsl:value-of select="'1'"/>
        </xsl:attribute>
      </xsl:if>
      <a:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:endParaRPr lang="zh-CN" altLang="en-US" dirty="0"/>
        </a:p>
      </a:txBody>
      <a:tcPr/>
    </a:tc>
  </xsl:template>
  <xsl:template name="tblRow1">
    <xsl:param name="totalColCount"/>
    <xsl:param name="totalRowCount"/>
    <xsl:param name="Rows"/>
    <xsl:param name="curRow"/>
    <xsl:param name="vmergestr"/>
    <a:tr>
      <xsl:attribute name="h">
        <xsl:choose>

          <!--start liuyin 20130112 修改文字表高度和列宽不正确-->
          <!--<xsl:when test="字:表行属性_41BD/字:高度_41B8/@固定值_41B9 * 20">
					 <xsl:value-of select="round(字:表行属性_41BD/字:高度_41B8/@固定值_41B9 * 20)"/>-->
          <xsl:when test="$Rows[$curRow]/字:表行属性_41BD/字:高度_41B8/@固定值_41B9">
            <xsl:value-of select="round($Rows[$curRow]/字:表行属性_41BD/字:高度_41B8/@固定值_41B9 * 12698)"/>
            <!--end liuyin 20130112 修改文字表高度和列宽不正确-->

          </xsl:when>
          <xsl:otherwise>

            <!--start liuyin 20130112 修改文字表高度和列宽不正确-->
            <!--<xsl:value-of select="round(字:表行属性_41BD/字:高度_41B8/@最小值_41BA * 20)"/>-->
            <xsl:value-of select="round($Rows[$curRow]/字:表行属性_41BD/字:高度_41B8/@最小值_41BA * 12698)"/>
            <!--end liuyin 20130112 修改文字表高度和列宽不正确-->

          </xsl:otherwise>
        </xsl:choose>

      </xsl:attribute>
        <xsl:call-template name="transtblCell">
          <xsl:with-param name="colCount" select="$totalColCount"/>
          <xsl:with-param name="rowCell" select="$Rows[$curRow]/child::字:单元格_41BE"/>
          <xsl:with-param name="vmergestr1" select="$vmergestr"/>
          <xsl:with-param name="curCol" select="1"/>
          <xsl:with-param name="curCell" select="1"/>
          <xsl:with-param name="hmergeCount" select="0"/>
          <xsl:with-param name="preRowSpan" select="0"/>
        </xsl:call-template>
      
      
    </a:tr>
    <xsl:variable name="tmpvms">
      <xsl:call-template name="getPreVms">
        <xsl:with-param name="colCount" select="$totalColCount"/>
        <xsl:with-param name="rowCell" select="$Rows[$curRow]/child::字:单元格_41BE"/>
        <xsl:with-param name="vmergestr1" select="$vmergestr"/>
        <xsl:with-param name="vmergestr2" select="''"/>
        <xsl:with-param name="curCol" select="1"/>
        <xsl:with-param name="curCell" select="1"/>
        <xsl:with-param name="hmergeCount" select="0"/>
      </xsl:call-template>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="$curRow &lt; $totalRowCount">
        <xsl:call-template name="tblRow1">
          <xsl:with-param name="totalColCount" select="$totalColCount"/>
          <xsl:with-param name="totalRowCount" select="$totalRowCount"/>
          <xsl:with-param name="Rows" select="$Rows"/>
          <xsl:with-param name="curRow" select="$curRow + 1"/>
          <xsl:with-param name="vmergestr" select="$tmpvms"/>
           
        </xsl:call-template>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  <!--end liuyangyang 2015-03-24 修复表格跨行转换问题 -->
  
  <xsl:template name="tblRow">
    <xsl:for-each select="字:行_41CD">
      <a:tr>
		 <xsl:attribute name="h">
			 <xsl:choose>
         
         <!--start liuyin 20130112 修改文字表高度和列宽不正确-->
				 <!--<xsl:when test="字:表行属性_41BD/字:高度_41B8/@固定值_41B9 * 20">
					 <xsl:value-of select="round(字:表行属性_41BD/字:高度_41B8/@固定值_41B9 * 20)"/>-->
         <xsl:when test="字:表行属性_41BD/字:高度_41B8/@固定值_41B9 * 1270">
           <xsl:value-of select="round(字:表行属性_41BD/字:高度_41B8/@固定值_41B9 * 12698)"/>
           <!--end liuyin 20130112 修改文字表高度和列宽不正确-->
         
         </xsl:when>
				 <xsl:otherwise>
           
           <!--start liuyin 20130112 修改文字表高度和列宽不正确-->
					 <!--<xsl:value-of select="round(字:表行属性_41BD/字:高度_41B8/@最小值_41BA * 20)"/>-->
           <xsl:value-of select="round(字:表行属性_41BD/字:高度_41B8/@最小值_41BA * 12698)"/>
           <!--end liuyin 20130112 修改文字表高度和列宽不正确-->
         
         </xsl:otherwise>
			 </xsl:choose>
          
        </xsl:attribute>
        <!--     若是有跨列，调用递归处理单元格 -->
        <xsl:for-each select="字:单元格_41BE">
          <xsl:call-template name="cell"/>
          <xsl:if test="字:单元格属性_41B7/字:跨列_41A7">
            <xsl:call-template name="Mergecell">
              <xsl:with-param name="Mergenum" select="字:单元格属性_41B7/字:跨列_41A7"/>
            </xsl:call-template>
          </xsl:if>
       </xsl:for-each>
      </a:tr>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="cell">
    <a:tc>
      <xsl:if test="字:单元格属性_41B7/字:跨列_41A7">
        <xsl:attribute name="gridSpan">
          <xsl:value-of select="字:单元格属性_41B7/字:跨列_41A7"/>
        </xsl:attribute>
      </xsl:if>
      <xsl:if test="字:单元格属性_41B7/字:跨行_41A6">
        <xsl:attribute name="rowSpan">
          <xsl:value-of select="字:单元格属性_41B7/字:跨行_41A6"/>
        </xsl:attribute>
      </xsl:if>
      
      <xsl:call-template name="txBody"/>
      <xsl:apply-templates select="字:单元格属性_41B7"/>
    </a:tc>
  </xsl:template>
  <xsl:template name="Mergecell">
    <xsl:param name="Mergenum"/>
    <xsl:if test="$Mergenum != 1 ">
      <a:tc hMerge="1">
        <a:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:endParaRPr lang="zh-CN" altLang="en-US" dirty="0"/>
          </a:p>
        </a:txBody>
        <a:tcPr/>
      </a:tc>
      <xsl:call-template name="Mergecell">
        <xsl:with-param name="Mergenum" select="$Mergenum - 1"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="字:单元格属性_41B7">
    <a:tcPr>
      <xsl:attribute name="marL">
        <xsl:value-of select="round(字:单元格边距_41A4/@左_C608 * 12698)"/>
      </xsl:attribute>
      <xsl:attribute name="marR">
        <xsl:value-of select="round(字:单元格边距_41A4/@右_C60A * 12698)"/>
      </xsl:attribute>
      <xsl:attribute name="marT">
        <xsl:value-of select="round(字:单元格边距_41A4/@上_C609 * 12698)"/>
      </xsl:attribute>
      <xsl:attribute name="marB">
        <xsl:value-of select="round(字:单元格边距_41A4/@下_C60B * 12698)"/>
      </xsl:attribute>

      <!--
      
      单元格对齐方式；左对齐，右对齐，居中在段落属性中
      -->
      <xsl:if test="not(字:垂直对齐方式_41A5='top')">
        <xsl:attribute name="anchor">
          <xsl:choose>
            <xsl:when test="字:垂直对齐方式_41A5='center'">
              <xsl:value-of select="'ctr'"/>
            </xsl:when>
            <xsl:when test="字:垂直对齐方式_41A5='bottom'">
              <xsl:value-of select="'b'"/>
            </xsl:when>
          </xsl:choose>
        </xsl:attribute>
      </xsl:if>
      
      <xsl:if test ="字:边框_4133">
        <xsl:apply-templates select ="字:边框_4133"/>
        
      </xsl:if>

		<xsl:choose>
			<xsl:when test="字:填充_4134/图:颜色_8004='auto'">
				<a:noFill/>
			</xsl:when>
      
      <!--start liuyin 20130112 修改当文字表未设置填充时，转换后出现填充颜色-->
      <!--修复文字表渐变填充丢失 liqiuling 2013-2-25  start-->
      <xsl:when test="not(字:填充_4134/图:颜色_8004)and not(字:填充_4134/图:图案_800A)and not(字:填充_4134/图:渐变_800D)">
        <xsl:choose>
          <xsl:when test="not(../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A)">
            <a:noFill/>
          </xsl:when>
          <!--start liuyin 20130112 修改图案填充出错-->
          <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008">
            <a:pattFill>
              <xsl:attribute name="prst">
                <xsl:call-template name="tablefillName"/>
              </xsl:attribute>
              <a:fgClr>
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:choose>
                      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@前景色_800B='auto'">
                        <xsl:value-of select="'000000'"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="substring(../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@前景色_800B,2,6)"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </xsl:attribute>
                </a:srgbClr>
              </a:fgClr>
              <a:bgClr>
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:choose>
                      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@背景色_800C='auto'">
                        <xsl:value-of select="'ffffff'"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="substring(../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@背景色_800C,2,6)"/>
                      </xsl:otherwise>
                    </xsl:choose>

                  </xsl:attribute>
                </a:srgbClr>
              </a:bgClr>
            </a:pattFill>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="fill"/>
          </xsl:otherwise>
          <!--end liuyin 20130112 修改图案填充出错-->
        </xsl:choose>
      </xsl:when>
      <!--end liuyin 20130112 修改当文字表未设置填充时，转换后出现填充颜色-->
      <!--修复文字表渐变填充丢失 liqiuling 2013-2-25  start-->
			<xsl:otherwise>
				<xsl:call-template name="fill"/>
			</xsl:otherwise>
		</xsl:choose>
      <!--start liuyin 20130112 修改图案填充出错-->
        <!--<xsl:if test="字:填充_4134">
        <xsl:call-template name="fill"/>
      </xsl:if>-->
      <!--end liuyin 20130112 修改图案填充出错-->
    </a:tcPr>
  </xsl:template>
  <xsl:template match ="字:边框_4133">
    <xsl:if test ="uof:左_C613">

      <xsl:apply-templates select ="uof:左_C613"/>
    </xsl:if>
    <xsl:if test ="uof:右_C615">
      <xsl:apply-templates select ="uof:右_C615"/>

    </xsl:if>
    <xsl:if test ="uof:上_C614">
      <xsl:apply-templates select ="uof:上_C614"/>

    </xsl:if>
    <xsl:if test ="uof:下_C616">

      <xsl:apply-templates select ="uof:下_C616"/>
    </xsl:if>
    
    <!--start liuyin 20130112 修改单元格内部出现斜上斜下框线-->
    <xsl:if test ="uof:对角线1_C617 and uof:对角线1_C617/@线型_C60D !='none'">
      <xsl:apply-templates select ="uof:对角线1_C617"/>
    </xsl:if>
    <xsl:if test ="uof:对角线2_C618 and uof:对角线2_C618/@线型_C60D !='none'">
      <xsl:apply-templates select ="uof:对角线2_C618"/>
    </xsl:if>
    <!--end liuyin 20130112 修改单元格内部出现斜上斜下框线-->
  </xsl:template>
  <xsl:template match ="uof:左_C613">
    <a:lnL>
      <xsl:call-template name="border"/>
    </a:lnL>
  </xsl:template>
  <xsl:template match ="uof:右_C615">
    <a:lnR>
      <xsl:call-template name="border"/>
    </a:lnR>
  </xsl:template>
  <xsl:template match ="uof:上_C614">
    <a:lnT>
      <xsl:call-template name="border"/>
    </a:lnT>
  </xsl:template>
  <xsl:template match ="uof:下_C616">
    <a:lnB>
      <xsl:call-template name="border"/>
    </a:lnB>
  </xsl:template>
  <xsl:template match ="uof:对角线1_C617">
    <a:lnTlToBr>
      <xsl:call-template name="border"/>
    </a:lnTlToBr>
  </xsl:template>
  <xsl:template match ="uof:对角线2_C618">
    <a:lnBlToTr>
      <xsl:call-template name="border"/>
    </a:lnBlToTr>
  </xsl:template>
  
  
  <xsl:template name="border">
    <xsl:attribute name="w">
      <xsl:value-of select="@宽度_C60F * 12700"/>
    </xsl:attribute>
    <xsl:attribute name="cap">
      <!--
      暂时不转 UOF不支持？
      需要对应一下描述      
      -->
      <xsl:value-of select="'flat'"/>
    </xsl:attribute>
    <xsl:attribute name="cmpd">
      <!--
      需要对应一下描述 根据 @uof:类型 判断 @cmpd和a:prstDash/@val
      的值
      -->
      <xsl:value-of select="'sng'"/>
    </xsl:attribute>
    <xsl:attribute name="algn">
       <xsl:value-of select="'ctr'"/>
    </xsl:attribute>
    <xsl:if test="@颜色_C611">
      <a:solidFill>
        <a:srgbClr>
          <xsl:attribute name="val">
            <xsl:choose>
              <xsl:when test="@颜色_C611!='auto'">
                <xsl:variable name="RGB">
                  <xsl:value-of select="@颜色_C611"/>
                </xsl:variable>
                <xsl:value-of select="substring-after($RGB,'#')"/>
              </xsl:when>
              <xsl:when test="@颜色_C611='auto'">000000</xsl:when>
            </xsl:choose>
          </xsl:attribute>
        </a:srgbClr>
      </a:solidFill>
    </xsl:if>
    <xsl:if test ="not(@颜色_C611)">
      <a:noFill/>
    </xsl:if>
    <a:prstDash>
      <xsl:attribute name="val">
        <xsl:choose>
          <!--start liuyin 20130112 修改文字表边框线经转换虚线效果丢失-->
          <xsl:when test="@线型_C60D='single' and not(@虚实_C60E)">
            <xsl:value-of select="'solid'"/>
          </xsl:when>
          <!--
          <xsl:when test="@线型_C60D='dotted'">
            <xsl:value-of select="'dot'"/>
          </xsl:when>
          <xsl:when test="@线型_C60D='dotted-heavy'">
            <xsl:value-of select="'sysDot'"/>
          </xsl:when>
          <xsl:when test="@线型_C60D='dot-dash'">
            <xsl:value-of select="'dashDot'"/>
          </xsl:when>
          <xsl:when test="@线型_C60D='dot-dot-dash'">
            <xsl:value-of select="'sysDashDotDot'"/>
          </xsl:when>
          <!- -
          加的
          - ->
          <xsl:when test="@线型_C60D='dash'">
            <xsl:value-of select="'dash'"/>
          </xsl:when>
          <xsl:when test="@线型_C60D='dash-dot-heavy'">
            <xsl:value-of select="'dashDot'"/>
          </xsl:when>
          <xsl:when test="@线型_C60D='dash-long'">
            <xsl:value-of select="'lgDash'"/>
          </xsl:when>
          <xsl:when test="@线型_C60D='dash-long-heavy'">
            <xsl:value-of select="'lgDashDot'"/>
          </xsl:when>
          <xsl:when test="@线型_C60D='dash-dot-dot-heavy'">
            <xsl:value-of select="'lgDashDotDot'"/>
          </xsl:when>
          <xsl:when test="@线型_C60D='dash-dot-heavy'">
            <xsl:value-of select="'sysDashDot'"/>
          </xsl:when>
          <xsl:when test="@线型_C60D='dash-dot-dot'">
            <xsl:value-of select="'sysDashDotDot'"/>
          </xsl:when>-->
          
          <xsl:when test="@虚实_C60E='dotted'">
          <xsl:value-of select="'dot'"/>
        </xsl:when>
        <xsl:when test="@虚实_C60E='dotted-heavy'">
          <xsl:value-of select="'sysDot'"/>
        </xsl:when>
        <xsl:when test="@虚实_C60E='dot-dash'">
          <xsl:value-of select="'dashDot'"/>
        </xsl:when>
        <xsl:when test="@虚实_C60E='dot-dot-dash'">
          <xsl:value-of select="'sysDashDotDot'"/>
        </xsl:when>
        <xsl:when test="@虚实_C60E='dash'">
          <xsl:value-of select="'dash'"/>
        </xsl:when>
        <xsl:when test="@虚实_C60E='dash-dot-heavy'">
          <xsl:value-of select="'dashDot'"/>
        </xsl:when>
        <xsl:when test="@虚实_C60E='dash-long'">
          <xsl:value-of select="'lgDash'"/>
        </xsl:when>
        <xsl:when test="@虚实_C60E='dash-long-heavy'">
          <xsl:value-of select="'lgDashDot'"/>
        </xsl:when>
        <xsl:when test="@虚实_C60E='dash-dot-dot-heavy'">
          <xsl:value-of select="'lgDashDotDot'"/>
        </xsl:when>
        <xsl:when test="@虚实_C60E='dash-dot-heavy'">
          <xsl:value-of select="'sysDashDot'"/>
        </xsl:when>
        <xsl:when test="@虚实_C60E='dash-dot-dot'">
          <xsl:value-of select="'sysDashDotDot'"/>
        </xsl:when>
          <xsl:when test="@虚实_C60E='square-dot'">
            <xsl:value-of select="'sysDot'"/>
          </xsl:when>
          <!--end liuyin 20130112 修改文字表边框线经转换虚线效果丢失-->
          
          <xsl:otherwise>
            <xsl:value-of select="'solid'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:attribute>
    </a:prstDash>
    <xsl:if test ="name(.)!='uof:对角线1_C617' and name(.)!='uof:对角线2_C618' ">

      <a:round/>
      <a:headEnd type="none" w="med" len="med"/>
      <a:tailEnd type="none" w="med" len="med"/>

    </xsl:if>
  </xsl:template>
  <xsl:template name="xfrm">
	  
	  <xsl:variable name="obj">
		  <xsl:value-of select="@图形引用_C62E"/>
   </xsl:variable>
    <xsl:if test="uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108 or uof:位置_C620/uof:垂直_410/uof:绝对_4107/@值_4108 or uof:大小_C621/@长_C604 or uof:大小_C621/@宽_C605">
		
		<xsl:choose>
        <xsl:when test="图:图形_8062/图:文本_803C/图:内容_8043/字:文字表_416C">
          <p:xfrm>
            <!-- 09.10.30 added by myx -->
            <xsl:if test=".//图:旋转角度_804D!='0.0'">
              <xsl:attribute name="rot">  
				
                <xsl:choose>
                  <xsl:when test="图:图形/图:翻转/@图:方向='y' or 图:图形/图:翻转/@图:方向='x'">
					
                    <xsl:value-of select="round((360-.//图:属性/图:旋转角度)*60000)"/>
                  </xsl:when>
                  <xsl:otherwise>
					 
                    <xsl:value-of select="round(.//图:属性/图:旋转角度*60000)"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <!--4月15日蒋俊彦修改-->
            <!--<xsl:if test="图:图形/图:翻转">-->
			  <xsl:if test="//图:图形_8062/图:翻转_803A">
              <xsl:call-template name="flip"/>
            </xsl:if>
            <!--原来是uof：X坐标和Y坐标 改为 锚点下 uof：位置-->
            <xsl:if test="uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108 or uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108">
              <a:off>
                <xsl:attribute name="x">
                  <xsl:choose>
                    <xsl:when test="contains(uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108,'-')">
                      <xsl:value-of select="'0'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="round(uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108*12700)"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
                <xsl:attribute name="y">
                  <xsl:choose>
                    <xsl:when test="contains(uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108,'-')">
                      <xsl:value-of select="'0'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="round(uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108*12700)"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </a:off>
            </xsl:if>
            <xsl:if test="uof:大小_C621/@宽_C605 or uof:大小_C621/@长_C604">
              <a:ext>
                <xsl:attribute name="cx">
                  <xsl:value-of select="round(uof:大小_C621/@宽_C605*12700)"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="round(uof:大小_C621/@长_C604*12700)"/>
                </xsl:attribute>
              </a:ext>
            </xsl:if>
          </p:xfrm>
        </xsl:when>
        <xsl:otherwise>
          <a:xfrm>
            <!-- 09.10.30 added by myx -->
            <xsl:if test=".//图:旋转角度_804D!='0.0'">
              <xsl:attribute name="rot">
                <!--2011-5-26 luowentian-->
                <xsl:choose>
                  <xsl:when test=".//图:翻转_803A='y' or .//图:翻转_803A='x'">
                    <xsl:value-of select="round((360-.//图:属性_801D/图:旋转角度_804D)*60000)"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="round(.//图:属性_801D/图:旋转角度_804D*60000)"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:if>
            <!--4月15日蒋俊彦修改-->
            <!--<xsl:if test="图:图形/图:翻转">-->
			  <!--<2012.02.05 lijuan>-->
			  <xsl:if test=".//图:翻转_803A">
              <xsl:call-template name="flip"/>
            </xsl:if>
            <!---->
			  <!-- 由于没有把图形 和母版的页脚 时间 锚点引入到版式下 所以出现问题 李娟-->
			  <!--<xsl:for-each select="//uof:锚点_C644">@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@-->
            <xsl:if test="./uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108 or uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108">
              <a:off>
                <xsl:attribute name="x">
                  <xsl:choose>
                    <xsl:when test="contains(uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108,'-')">
                      <xsl:value-of select="'0'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="round(uof:位置_C620/uof:水平_4106/uof:绝对_4107/@值_4108*12700)"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
                <xsl:attribute name="y">
                  <xsl:choose>
                    <xsl:when test="contains(uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108,'-')">
                      <xsl:value-of select="'0'"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="round(uof:位置_C620/uof:垂直_410D/uof:绝对_4107/@值_4108*12700)"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:attribute>
              </a:off>
            </xsl:if>
            <xsl:if test="./uof:大小_C621/@长_C604 or ./uof:大小_C621/@宽_C605">
              <a:ext>
                <xsl:attribute name="cx">
                  <xsl:value-of select="round(uof:大小_C621/@宽_C605*12700)"/>
                </xsl:attribute>
                <xsl:attribute name="cy">
                  <xsl:value-of select="round(uof:大小_C621/@长_C604*12700)"/>
                </xsl:attribute>
              </a:ext>
            </xsl:if>
			  <!--</xsl:for-each>-->
          </a:xfrm>
        </xsl:otherwise>
      </xsl:choose>
		
    </xsl:if>
  </xsl:template>
  <xsl:template name="fill">
	  
    <xsl:if test=".//图:颜色_8004 and substring(.//图:颜色_8004,2,6)!='ffffff'">
      
      <!--start liuyin 20130110 修改幻灯片配色方案效果丢失-->
      <xsl:variable name="colorSchemename">
        <xsl:value-of select="../@配色方案引用_6C12"/>
      </xsl:variable>
      <!--end liuyin 20130110 修改幻灯片配色方案效果丢失-->
      
      <!--规则:配色方案集_6C11/规则:配色方案_6BE4[@标识符_6B0A=../@配色方案引用_6C12]-->
      <a:solidFill>
        <a:srgbClr>
          <xsl:choose>
            
            <!--liuyangyang 2015-04-01 修改背景填充色丢失的bug start-->
            <xsl:when test="./图:颜色_8004 !=''">
              <xsl:attribute name="val">
                <xsl:choose>
                  <xsl:when test="./图:颜色_8004='auto'">
                    <xsl:value-of select="'000000'"/>
                  </xsl:when>
                  <xsl:otherwise>
                    <xsl:value-of select="substring(./图:颜色_8004,2,6)"/>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:attribute>
            </xsl:when>
            <!--end   liuyangyang 2015-04-01 修改背景填充色丢失的bug-->
            
            <!--start liuyin 20130110 修改幻灯片配色方案效果丢失-->
            <xsl:when test="/uof:UOF/uof:演示文稿/演:公用处理规则/规则:配色方案集_6C11/规则:配色方案_6BE4[@标识符_6B0A=$colorSchemename]/规则:填充_6B06">
              <xsl:attribute name="val">
                <xsl:value-of select="substring(/uof:UOF/uof:演示文稿/演:公用处理规则/规则:配色方案集_6C11/规则:配色方案_6BE4[@标识符_6B0A=$colorSchemename]/规则:填充_6B06,2,6)"/>
              </xsl:attribute>
            </xsl:when>
            <!--end liuyin 20130110 修改幻灯片配色方案效果丢失-->
            <!--修改幻灯片背景不正确  liqiuling 2013-03-29 start-->
			      <xsl:when test=".//图:颜色_8004='auto'">
              
              <!--start liuyin 20130331 修改bug2781,side4背景填充不正确-->
              <!--<xsl:attribute name="val">ffffff</xsl:attribute>-->
              <xsl:attribute name="val"><xsl:value-of  select="substring(//规则:配色方案集_6C11/规则:配色方案_6BE4/规则:填充_6B06,2,6)"/></xsl:attribute>
              <!--end liuyin 20130331 修改bug2781,side4背景填充不正确-->
              
            </xsl:when>
            <!--修改幻灯片背景不正确  liqiuling 2013-03-29 start-->
            <xsl:otherwise>
              <xsl:attribute name="val">
                <xsl:value-of select="substring(.//图:颜色_8004,2,6)"/>
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
          <xsl:if test=".//图:属性_801D/图:透明度_8050">
            <a:alpha>
              <xsl:attribute name="val">
                <xsl:value-of select="round((100-.//图:属性_801D/图:透明度_8050)*1000)"/>
              </xsl:attribute>
            </a:alpha>
          </xsl:if>
        </a:srgbClr>
      </a:solidFill>
    </xsl:if>
    <!--处理除白色填充-->
    <xsl:if test="substring(.//图:颜色_8004,2,6)='ffffff'">
      <a:solidFill>
        <a:srgbClr val="ffffff"/>
      </a:solidFill>
    </xsl:if>
    <!--处理渐变填充-->
    <xsl:if test=".//图:渐变_800D">
      <xsl:variable name="angle" select=".//图:渐变_800D/@渐变方向_8013"/>
      <a:gradFill flip="none" rotWithShape="0">

        <!--start liuyin 20130110 修改预定义图形渐变填充转换错误-->
          <xsl:choose>
            <xsl:when test=".//图:渐变_800D/@种子类型_8010='radar'">
              <a:gsLst>
              <a:gs pos="0">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(.//图:渐变_800D/@起始色_800E,2,6)"/>
                  </xsl:attribute>
                  <!--2011-4-4罗文甜-->
                  <xsl:if test=".//图:属性_801D/图:透明度_8050">
                    <a:alpha>
                      <xsl:attribute name="val">
                        <xsl:value-of select="(100-.//图:属性_801D/图:透明度_8050)*1000"/>
                      </xsl:attribute>
                    </a:alpha>
                  </xsl:if>
                </a:srgbClr>
              </a:gs>
              <a:gs pos="50000">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(.//图:渐变_800D/@终止色_800F,2,6)"/>
                  </xsl:attribute>
                </a:srgbClr>
              </a:gs>
              <a:gs pos="100000">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(.//图:渐变_800D/@起始色_800E,2,6)"/>
                  </xsl:attribute>
                </a:srgbClr>
              </a:gs>
            </a:gsLst>
            </xsl:when>
            <xsl:when test=".//图:渐变_800D/@种子类型_8010='axial'">
             <a:gsLst>
              <a:gs pos="0">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(.//图:渐变_800D/@起始色_800E,2,6)"/>
                  </xsl:attribute>
                  <!--2011-4-4罗文甜-->
                  <xsl:if test=".//图:属性_801D/图:透明度_8050">
                    <a:alpha>
                      <xsl:attribute name="val">
                        <xsl:value-of select="(100-.//图:属性_801D/图:透明度_8050)*1000"/>
                      </xsl:attribute>
                    </a:alpha>
                  </xsl:if>
                </a:srgbClr>
              </a:gs>
              <a:gs pos="50000">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(.//图:渐变_800D/@终止色_800F,2,6)"/>
                  </xsl:attribute>
                </a:srgbClr>
              </a:gs>
              <a:gs pos="100000">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(.//图:渐变_800D/@起始色_800E,2,6)"/>
                  </xsl:attribute>
                </a:srgbClr>
              </a:gs>
            </a:gsLst>
            </xsl:when>
            <xsl:when test=".//图:渐变_800D/@种子类型_8010='rectangle'">
             <a:gsLst> 
              <a:gs pos="0">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(.//图:渐变_800D/@起始色_800E,2,6)"/>
                  </xsl:attribute>
                  <!--2011-4-4罗文甜-->
                  <xsl:if test=".//图:属性_801D/图:透明度_8050">
                    <a:alpha>
                      <xsl:attribute name="val">
                        <xsl:value-of select="(100-.//图:属性_801D/图:透明度_8050)*1000"/>
                      </xsl:attribute>
                    </a:alpha>
                  </xsl:if>
                </a:srgbClr>
              </a:gs>
              <a:gs pos="100000">
                <a:srgbClr>
                  <xsl:attribute name="val">
                    <xsl:value-of select="substring(.//图:渐变_800D/@终止色_800F,2,6)"/>
                  </xsl:attribute>
                </a:srgbClr>
              </a:gs>
             </a:gsLst>
              <a:path path="rect">
                <a:fillToRect>
                  <xsl:attribute name="l">
                    <xsl:value-of select=".//图:渐变_800D/@种子X位置_8015*1000"/>
                  </xsl:attribute>
                  <xsl:attribute name="t">
                    <xsl:value-of select=".//图:渐变_800D/@种子Y位置_8016*1000"/>
                  </xsl:attribute>
                  <xsl:attribute name="r">
                    <xsl:value-of select="100000-.//图:渐变_800D/@种子X位置_8015*1000"/>
                  </xsl:attribute>
                  <xsl:attribute name="b">
                    <xsl:value-of select="100000-.//图:渐变_800D/@种子Y位置_8016*1000"/>
                  </xsl:attribute>
                </a:fillToRect>
              </a:path>
            </xsl:when>
            
            <!--start liuyin 20130110 修改母版背景渐变-预设效果错误错误-->
            <xsl:when test="../../../演:母版集_6C0C/演:母版_6C0D/演:背景_6B2C/图:渐变_800D/@类型_8008='rainbow1'">
              <a:gsLst>
                <a:gs pos="0">
                  <a:srgbClr val="DB214D"/>
                </a:gs>
                <a:gs pos="14000">
                  <a:srgbClr val="DD7F2C"/>
                </a:gs>
                <a:gs pos="32000">
                  <a:srgbClr val="E1F500"/>
                </a:gs>
                <a:gs pos="50000">
                  <a:srgbClr val="00A13B"/>
                </a:gs>
                <a:gs pos="67000">
                  <a:srgbClr val="1AAC8C"/>
                </a:gs>
                <a:gs pos="83000">
                  <a:srgbClr val="36B2C6"/>
                </a:gs>
                <a:gs pos="100000">
                  <a:srgbClr val="756BA4"/>
                </a:gs>
              </a:gsLst>
              <a:lin ang="5400000" scaled="1"/>
            </xsl:when>

            <xsl:when test=".//图:渐变_800D/@种子类型_8010='linear'">
              <a:gsLst>
                <xsl:if test=".//图:渐变_800D/@起始色_800E">
                <a:gs pos="0">
                  <a:srgbClr>
                    <xsl:attribute name="val">
                      <xsl:value-of select="substring(.//图:渐变_800D/@起始色_800E,2,6)"/>
                    </xsl:attribute>
                    <!--2011-4-4罗文甜-->
                    <xsl:if test=".//图:属性_801D/图:透明度_8050">
                      <a:alpha>
                        <xsl:attribute name="val">
                          <xsl:value-of select="(100-.//图:属性_801D/图:透明度_8050)*1000"/>
                        </xsl:attribute>
                      </a:alpha>
                    </xsl:if>
                  </a:srgbClr>
                </a:gs>
                <a:gs pos="100000">
                  <a:srgbClr>
                    <xsl:attribute name="val">
                      <xsl:value-of select="substring(.//图:渐变_800D/@终止色_800F,2,6)"/>
                    </xsl:attribute>
                  </a:srgbClr>
                </a:gs>
                </xsl:if>
                
                <!--start liuyin 20130331 修改bug2781,side2，填充不正确-->
                <xsl:if test="not(.//图:渐变_800D/@起始色_800E)">
                  <a:gs pos="0">
                    <a:srgbClr val="000000"/>
                  </a:gs>
                  <a:gs pos="100000">
                    <a:srgbClr val="FFFFFF"/>
                  </a:gs>
                  <!--<a:gs pos="0">
                    <a:srgbClr val="DB214D"/>
                  </a:gs>
                  <a:gs pos="14000">
                    <a:srgbClr val="DD7F2C"/>
                  </a:gs>
                  <a:gs pos="32000">
                    <a:srgbClr val="E1F500"/>
                  </a:gs>
                  <a:gs pos="50000">
                    <a:srgbClr val="00A13B"/>
                  </a:gs>
                  <a:gs pos="67000">
                    <a:srgbClr val="1AAC8C"/>
                  </a:gs>
                  <a:gs pos="83000">
                    <a:srgbClr val="36B2C6"/>
                  </a:gs>
                  <a:gs pos="100000">
                    <a:srgbClr val="756BA4"/>
                  </a:gs>-->
                  <!--end liuyin 20130331 修改bug2781,side2，填充不正确-->
                </xsl:if>
              </a:gsLst>
              <a:lin ang="5400000" scaled="1"/>
            </xsl:when>
            <!--end liuyin 20130110 修改母版背景渐变-预设效果错误错误-->
            
            <xsl:otherwise>
              <xsl:choose>
                <xsl:when test="$angle='135' or $angle='180' or $angle='225' or $angle='270'">
                 <a:gsLst>
                  <a:gs pos="100000">
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:value-of select="substring(.//图:渐变_800D/@起始色_800E,2,6)"/>
                      </xsl:attribute>
                      <!--2011-4-4罗文甜-->
                      <xsl:if test=".//图:属性_801D/图:透明度_8050">
                        <a:alpha>
                          <xsl:attribute name="val">
                            <xsl:value-of select="(100-.//图:属性_801D/图:透明度_8050)*1000"/>
                          </xsl:attribute>
                        </a:alpha>
                      </xsl:if>
                    </a:srgbClr>
                  </a:gs>
                  <a:gs pos="0">
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:value-of select="substring(.//图:渐变_800D/@终止色_800F,2,6)"/>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:gs>
                 </a:gsLst>
                </xsl:when>
                <xsl:otherwise>
                 <a:gsLst>
                  <a:gs pos="0">
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:value-of select="substring(.//图:渐变_800D/@起始色_800E,2,6)"/>
                      </xsl:attribute>
                      <!--2011-4-4罗文甜-->
                      <xsl:if test=".//图:属性_801D/图:透明度_8050">
                        <a:alpha>
                          <xsl:attribute name="val">
                            <xsl:value-of select="(100-.//图:属性_801D/图:透明度_8050)*1000"/>
                          </xsl:attribute>
                        </a:alpha>
                      </xsl:if>
                    </a:srgbClr>
                  </a:gs>
                  <a:gs pos="100000">
                    <a:srgbClr>
                      <xsl:attribute name="val">
                        <xsl:value-of select="substring(.//图:渐变_800D/@终止色_800F,2,6)"/>
                      </xsl:attribute>
                    </a:srgbClr>
                  </a:gs>
                 </a:gsLst>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:otherwise>
            
          </xsl:choose>
        <xsl:choose>
          <xsl:when test=".//图:渐变_800D/@种子类型_8010='square'">
            <a:path>
              <!--<xsl:attribute name="path">shape</xsl:attribute>-->
              <xsl:attribute name="path">rect</xsl:attribute>
              <xsl:variable name="x" select=".//图:渐变_800D/@种子X位置_8015"/>
              <xsl:variable name="y" select=".//图:渐变_800D/@种子Y位置_8016"/>
              <!--<xsl:choose>
                <xsl:when test =".//图:渐变/@图:种子X位置='100'and .//图:渐变/@图:种子Y位置='100'">
                  <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
                </xsl:when>
                <xsl:when test =".//图:渐变/@图:种子X位置='30'and .//图:渐变/@图:种子Y位置='30'">
                  <a:fillToRect l="0" t="0" r="85000" b="85000"/>
                </xsl:when>
                <xsl:otherwise>
                  <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
                </xsl:otherwise>
              </xsl:choose>-->
              <xsl:choose>
                <xsl:when test="$x='30' and $y='30'">
                  <a:fillToRect r="100000" b="100000"/>
                </xsl:when>
                <xsl:when test="$x='30' and $y='60'">
                  <a:fillToRect t="100000" r="100000"/>
                </xsl:when>
                <xsl:when test="$x='60' and $y='30'">
                  <a:fillToRect l="100000" b="100000"/>
                </xsl:when>
                <xsl:when test="$x='60' and $y='60'">
                  <a:fillToRect l="100000" t="100000"/>
                </xsl:when>
              </xsl:choose>
            </a:path>
          </xsl:when>
          <xsl:when test=".//图:渐变_800D/@种子类型_8010='axial'">
            <a:lin>
              <xsl:attribute name="ang">
                <!--4月16日jjy改渐变颜色-->
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
          </xsl:when>
           <xsl:otherwise>
          </xsl:otherwise>
        <!--end liuyin 20130110 修改预定义图形渐变填充转换错误-->
        </xsl:choose>
        <a:tileRect/>
      </a:gradFill>
    </xsl:if>
    <!--2月20日改:处理纹理和图片填充-->
    <xsl:if test=".//图:图片_8005">
      <a:blipFill rotWithShape="0">
        <a:blip>
          <!--10.23 黎美秀修改
            <xsl:attribute name="r:embed">
            <xsl:value-of select="concat('rId',substring-after(.//图:图片/@图:图形引用,'OBJ'))"/>
          </xsl:attribute>
          
          -->
          <xsl:attribute name="r:embed">
            <xsl:value-of select="concat('rId',.//图:图片_8005/@图形引用_8007)"/>
          </xsl:attribute>
          <!--2011-4-4罗文甜-->
          <xsl:if test=".//图:属性_801D/图:透明度_8050">
            <a:alphaModFix>
              <xsl:attribute name="amt">
                <xsl:value-of select="(100-.//图:属性_801D/图:透明度_8050)*1000"/>
              </xsl:attribute>
            </a:alphaModFix>
          </xsl:if>
        </a:blip>
        <xsl:if test=".//图:图片_8005/@位置_8006='title'">
          <a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>
        </xsl:if>
        <a:srcRect/>
        <xsl:if test=".//图:图片_8005/@位置_8006='stretch'">
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </xsl:if>
      </a:blipFill>
    </xsl:if>
    <!--处理图案填充-->
    <xsl:if test=".//图:图案_800A">
      <a:pattFill>
        <xsl:attribute name="prst">
          <xsl:call-template name="fillName"/>
        </xsl:attribute>
        <a:fgClr>
          <a:srgbClr>
            <!--start liuyin 20130112 修改图案填充前景色为auto时出错-->
            <xsl:attribute name="val">
              <xsl:choose>
                <xsl:when test=".//图:图案_800A/@前景色_800B='auto'">
                  <xsl:value-of select="'000000'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="substring(.//图:图案_800A/@前景色_800B,2,6)"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:attribute>
            <!--end liuyin 20130112 修改图案填充前景色为auto时出错-->
          </a:srgbClr>
        </a:fgClr>
        <a:bgClr>
          <a:srgbClr>
            <xsl:attribute name="val">
              <!--start liuyin 20130112 修改图案填充前景色为auto时出错-->
              <xsl:choose>
                <xsl:when test=".//图:图案_800A/@背景色_800C='auto'">
                  <xsl:value-of select="'ffffff'"/>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:value-of select="substring(.//图:图案_800A/@背景色_800C,2,6)"/>
                </xsl:otherwise>
              </xsl:choose>
              <!--end liuyin 20130112 修改图案填充前景色为auto时出错-->
            </xsl:attribute>
          </a:srgbClr>
        </a:bgClr>
      </a:pattFill>
    </xsl:if>

  </xsl:template>
  <!--start liuyin 20130112 修改图案填充出错-->
  <xsl:template name="tablefillName">
    <xsl:choose>
      <!--2011-3-20罗文甜，增加图案类型的判断-->
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn001'">pct5</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn002'">pct10</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn003'">pct20</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn004'">pct25</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn005'">pct30</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn006'">pct40</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn007'">pct50</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn008'">pct60</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn009'">pct70</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn010'">pct75</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn011'">pct80</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn012'">pct90</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn013'">ltDnDiag</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn014'">ltUpDiag</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn015'">dkDnDiag</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn016'">dkUpDiag</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn017'">wdDnDiag</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn018'">wdUpDiag</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn019'">ltVert</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn020'">ltHorz</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn021'">narVert</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn022'">narHorz</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn023'">dkVert</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn024'">dkHorz</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn025'">dashDnDiag</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn026'">dashUpDiag</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn027'">dashVert</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn028'">dashHorz</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn029'">smConfetti</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn030'">lgConfetti</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn031'">zigZag</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn032'">wave</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn033'">diagBrick</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn034'">horzBrick</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn035'">weave</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn036'">plaid</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn037'">divot</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn038'">dotGrid</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn039'">dotDmnd</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn040'">shingle</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn041'">trellis</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn042'">sphere</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn043'">smGrid</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn044'">lgGrid</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn045'">smCheck</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn046'">lgCheck</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn047'">openDmnd</xsl:when>
      <xsl:when test="../../../字:文字表属性_41CC/字:填充_4134/图:图案_800A/@类型_8008='ptn048'">solidDmnd</xsl:when>
      <xsl:otherwise>pct5</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--start liuyin 20130112 修改图案填充出错-->
  
  <xsl:template name="fillName">
    <xsl:choose>
      <!--2011-3-20罗文甜，增加图案类型的判断-->
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
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn027'">dashVert</xsl:when>
      <xsl:when test=".//图:图案_800A/@类型_8008='ptn028'">dashHorz</xsl:when>
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
      <xsl:otherwise>pct5</xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <xsl:template name="flip">
    <!--2011-3-28罗文甜：修改Bug-->
    <xsl:if test=".//图:翻转_803A='x'">
      <xsl:attribute name="flipH">1</xsl:attribute>
    </xsl:if>
    <xsl:if test=".//图:翻转_803A='y'">
      <xsl:attribute name="flipV">1</xsl:attribute>
    </xsl:if>
    <xsl:if test=".//图:翻转_803A='xy'">
      <xsl:attribute name="flipV">1</xsl:attribute>
      <xsl:attribute name="flipH">1</xsl:attribute>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>
