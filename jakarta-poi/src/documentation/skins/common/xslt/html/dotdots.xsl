<?xml version="1.0" encoding="utf-8"?>
<!-- Copyright (C) 2004 The Apache Software Foundation. All rights reserved. -->

<!--
Contains the 'dotdots' template, which, given a path, will output a set of
directory traversals to get back to the source directory. Handles both '/' and
'\' directory separators.

Examples:
  Input                           Output 
    index.html                    ""
    dir/index.html                "../"
    dir/subdir/index.html         "../../"
    dir//index.html              "../"
    dir/                          "../"
    dir//                         "../"
    \some\windows\path            "../../"
    \some\windows\path\           "../../../"
    \Program Files\mydir          "../"

Cannot handle ..'s in the path, so don't expect 'dir/subdir/../index.html' to
work.

jefft@apache.org
-->

<xsl:stylesheet
  version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <xsl:template name="dotdots">
    <xsl:param name="path"/>
    <xsl:variable name="dirs" select="normalize-space(translate(concat($path, 'x'), ' /\', '_  '))"/>
    <!-- The above does the following:
       o Adds a trailing character to the path. This prevents us having to deal
         with the special case of ending with '/'
       o Translates all directory separators to ' ', and normalize spaces,
		 cunningly eliminating duplicate '//'s. We also translate any real
		 spaces into _ to preserve them.
    -->
    <xsl:variable name="remainder" select="substring-after($dirs, ' ')"/>
    <xsl:if test="$remainder">
      <xsl:text>../</xsl:text>
      <xsl:call-template name="dotdots">
        <xsl:with-param name="path" select="translate($remainder, ' ', '/')"/>
		<!-- Translate back to /'s because that's what the template expects. -->
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

<!--
  Uncomment to test.
  Usage: saxon dotdots.xsl dotdots.xsl path='/my/test/path'

  <xsl:param name="path"/>
  <xsl:template match="/">
    <xsl:message>Path: <xsl:value-of select="$path"/></xsl:message>
    <xsl:call-template name="dotdots">
      <xsl:with-param name="path" select="$path"/>
    </xsl:call-template>
  </xsl:template>
 -->

</xsl:stylesheet>
