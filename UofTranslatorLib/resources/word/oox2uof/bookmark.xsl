<?xml version="1.0" encoding="utf-8"?>
<!--
* Copyright (c) 2006, Beihang University, China
* All rights reserved.
*
* Redistribution and use in source and binary forms, with or without
* modification, are permitted provided that the following conditions are met:
*
*     * Redistributions of source code must retain the above copyright
*       notice, this list of conditions and the following disclaimer.
*     * Redistributions in binary form must reproduce the above copyright
*       notice, this list of conditions and the following disclaimer in the
*       documentation and/or other materials provided with the distribution.
*     * Neither the name of Clever Age, nor the names of its contributors may
*       be used to endorse or promote products derived from this software
*       without specific prior written permission.
*
* THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
* EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
* WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
* DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
* DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
* (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
* LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
* ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
* (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
* SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
-->
<!--
<Author>WuDi</Author>
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:字="http://schemas.uof.org/cn/2009/wordproc"
  xmlns:书签="http://schemas.uof.org/cn/2009/bookmarks">
  <xsl:output encoding="UTF-8" indent="yes" method="xml" version="1.0"/>

  <!--书签模板-->
  <xsl:template name="bookmarks">
    <xsl:for-each select="w:bookmarkStart">
      <!--<xsl:choose>
        <xsl:when test="name(.)='w:bookmarkStart'">-->        
          <书签:书签_9105>
            <xsl:attribute name="名称_9103">
              <xsl:value-of select="@w:name"/>
            </xsl:attribute>
            <书签:区域_9100>
              <xsl:attribute name="区域引用_41CE">
                <xsl:value-of select="concat('bk_',@w:id)"/>
              </xsl:attribute>
            </书签:区域_9100>
          </书签:书签_9105>
        <!--</xsl:when>
      </xsl:choose>-->
    </xsl:for-each>
  </xsl:template>

  <!--2012-11-23，wudi，OOX到UOF方向的书签实现，start-->
  <xsl:template name="bookmarkStart">
    <字:句_419D>
      <字:区域开始_4165>
        <xsl:attribute name="标识符_4100">
          <xsl:value-of select= "concat('bk_',@w:id)"/>
        </xsl:attribute>
        <xsl:attribute name="名称_4166">
          <xsl:value-of select="@w:name"/>
        </xsl:attribute>
        <xsl:attribute name="类型_413B">
          <xsl:value-of select="'bookmark'"/>
        </xsl:attribute>
      </字:区域开始_4165>
    </字:句_419D>
  </xsl:template>

  <xsl:template name="bookmarkEnd">
    <字:句_419D>
      <字:区域结束_4167>
        <xsl:attribute name="标识符引用_4168">
          <xsl:value-of select= "concat('bk_',@w:id)"/>
        </xsl:attribute>
      </字:区域结束_4167>
    </字:句_419D>
  </xsl:template>
  <!--end-->
  
</xsl:stylesheet>
