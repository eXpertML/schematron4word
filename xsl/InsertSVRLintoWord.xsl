<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:math="http://www.w3.org/2005/xpath-functions/math"
  xmlns:map="http://www.w3.org/2005/xpath-functions/map"
  xmlns:xd="http://www.oxygenxml.com/ns/doc/xsl"
  xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage" 
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"  
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" 
  xmlns:o="urn:schemas-microsoft-com:office:office" 
  xmlns:v="urn:schemas-microsoft-com:vml" 
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" 
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
  xmlns:svrl="http://purl.oclc.org/dsdl/svrl"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:s4w="https://github.com/eXpertML/schematron4word"
  exclude-result-prefixes="#all"
  version="3.0">
  <!-- xsl xs math map xd mc w14 r m o v w15 a svrl -->
  
  <xd:doc scope="stylesheet">
    <xd:desc>
      <xd:p><xd:b>Created on:</xd:b> Sep 1, 2025</xd:p>
      <xd:p><xd:b>Author:</xd:b> TFJH, OUP</xd:p>
      <xd:p>This stylesheet takes a word document as its primary input and an SVRL as its secondary input.  It produces an annotated copy of the word document where SVRL error messages have been inserted as comment annotations.</xd:p>
    </xd:desc>
    <xd:param name="SVRL">XML document node containing an SVRL report for some schematron operation on the input document.  For convenience, three other parameters are made available to set this by providing items other than a document node:</xd:param>
    <xd:param name="SVRLurl">String containing a URL to an SVRL XML document.  This static parameter is used as an alternative way of populating $SVRL.</xd:param>
    <xd:param name="SVRLtext">String value of an unparsed XML document.  This static parameter is used as an alternative way of populating $SVRL, and is particularly needed due to restrictions in the Microsoft Add-In environment when transforming with SaxonJs.</xd:param>
    <xd:param name="SVRLnode">Element value of some SVRL document/fragment.  This is particularly useful for testing with XSpec parameters.  Since XSpec scripts need to be able to set this parameter dynamically, it cannot be declared as static, and must therefore be the last alternative for population of $SVRL.</xd:param>
  </xd:doc>
  
  <xsl:mode on-no-match="shallow-copy"/>
  
  <xsl:param name="SVRLurl" as="xs:string?" static="yes"/>
  
  <xsl:param name="SVRLtext" as="xs:string?" static="yes"/>
  
  <xsl:param name="SVRLnode" as="element()?"/>
  
  <xsl:param name="SVRL" as="document-node()">
    <xsl:choose>
      <xsl:when test="exists($SVRLurl) and doc-available($SVRLurl)">
        <xsl:sequence select="doc($SVRLurl)"/>
      </xsl:when>
      <xsl:when test="exists($SVRLtext)">
        <xsl:sequence select="parse-xml($SVRLtext)"/>
      </xsl:when> 
      <xsl:otherwise>
        <xsl:document>
          <xsl:sequence select="$SVRLnode"/>
        </xsl:document>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:param>
  
  <xsl:variable name="errs" as="element()*" select="$SVRL//svrl:successful-report|$SVRL//svrl:failed-assert"/>
  
  <xsl:variable name="keepComments" select="/pkg:package/pkg:part/pkg:xmlData/w:comments/w:comment[not(@w:initials='S4W')]"/>
  
  <xsl:key name="commentByID" match="w:comment" use="@w:id"/>
  
  <xsl:variable name="docRoot" select="/"/>
  
  <xsl:key name="errorByElement" match="svrl:failed-assert | svrl:successful-report">
    <xsl:variable name="El" as="element()*">
      <xsl:evaluate xpath="'$docRoot' || @location">
        <xsl:with-param name="docRoot" select="$docRoot"/>
      </xsl:evaluate>
    </xsl:variable>
    <xsl:sequence select="outermost($El) ! generate-id(.)"/>
  </xsl:key>
  
  <xd:doc>
    <xd:desc>Adds error-comments in around text runs</xd:desc>
  </xd:doc>
  <xsl:template match="w:r[key('errorByElement', generate-id(.), $SVRL)]">
    <xsl:call-template name="s4w:addCommentRef">
      <xsl:with-param name="contents" tunnel="yes">
        <xsl:next-match/>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>Adds error-comments within paragraphs</xd:desc>
  </xd:doc>
  <xsl:template match="w:p[key('errorByElement', generate-id(.), $SVRL)]">
    <xsl:variable name="preliminaries" select="w:pPr"/>
    <xsl:copy>
      <xsl:apply-templates select="@*"/>
      <xsl:apply-templates select="$preliminaries"/>
      <xsl:call-template name="s4w:addCommentRef">
        <xsl:with-param name="contents" tunnel="yes">
          <xsl:apply-templates select="* except $preliminaries"/>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:copy>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>Named template to recursively add comment references</xd:desc>
    <xd:param name="contents">The content of the text that is inside the comment range(s)</xd:param>
    <xd:param name="commentIDs">A list of SVRL error comment IDs that need their references adding.</xd:param>
  </xd:doc>
  <xsl:template name="s4w:addCommentRef">
    <xsl:param name="contents" as="node()*" tunnel="yes"/>
    <xsl:param name="commentIDs" as="xs:string*" select="key('errorByElement', generate-id(.), $SVRL) ! s4w:SVRL-id(.)"/>
    <xsl:choose>
      <xsl:when test="empty($commentIDs)">
        <xsl:sequence select="$contents"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="SVRL-id" select="head($commentIDs)"/>
        <w:commentRangeStart w:id="{$SVRL-id}"/>
        <xsl:call-template name="s4w:addCommentRef">
          <xsl:with-param name="commentIDs" select="tail($commentIDs)"/>
        </xsl:call-template>
        <w:commentRangeEnd w:id="{$SVRL-id}"/>
        <w:r>
          <w:commentReference w:id="{$SVRL-id}"/>
        </w:r>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>Creates comment part in word package (as required)</xd:desc>
  </xd:doc>
  <xsl:template match="pkg:package">
    <xsl:copy>
      <xsl:apply-templates select="@* | node()"/>
      <xsl:where-populated>
        <pkg:part pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml">
          <xsl:where-populated>
            <pkg:xmlData>
              <xsl:where-populated>
                <w:comments>
                  <xsl:if test="empty(/pkg:package/pkg:part/pkg:xmlData/w:comments)">
                    <xsl:namespace name="w" select="'http://schemas.openxmlformats.org/wordprocessingml/2006/main'"/>
                    <xsl:apply-templates select="$SVRL" mode="comments"/>
                  </xsl:if>
                </w:comments>
              </xsl:where-populated>
            </pkg:xmlData>
          </xsl:where-populated>
        </pkg:part>
      </xsl:where-populated>
    </xsl:copy>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>Adds comments to existing comment parts</xd:desc>
  </xd:doc>
  <xsl:template match="w:comments">
    <xsl:copy>
      <xsl:apply-templates select="@*|node()"/>
      <xsl:apply-templates select="$SVRL" mode="comments"/>
    </xsl:copy>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>Remove old S4W comments</xd:desc>
  </xd:doc>
  <xsl:template match="w:comment[@w:initials='S4W']"/>
  
  <xd:doc>
    <xd:desc>Renumber comments we want to keep</xd:desc>
  </xd:doc>
  <xsl:template match="w:comment[not(@w:initials='S4W')]">
    <xsl:copy>
      <xsl:apply-templates select="@* except @w:id"/>
      <xsl:attribute name="w:id" select="index-of($keepComments, .) - 1"/>
      <xsl:apply-templates select="node()"/>
    </xsl:copy>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>Remove old S4W comment anchors</xd:desc>
  </xd:doc>
  <xsl:template 
    match="w:commentRangeStart[key('commentByID', @w:id)/@w:initials='S4W']
          |w:commentRangeEnd[key('commentByID', @w:id)/@w:initials='S4W']
          |w:r/w:commentReference[key('commentByID', @w:id)/@w:initials='S4W']"/>
  
  <xd:doc>
    <xd:desc>Re-number comments we're keeping</xd:desc>
  </xd:doc>
  <xsl:template match="@w:id[
      not(key('commentByID', .)/@w:initials = 'S4W')
      and (
        parent::w:commentRangeStart or
        parent::w:commentRangeEnd or
        parent::w:commentReference
      )
    ]">
    <xsl:attribute name="w:id" select="index-of($keepComments, key('commentByID', .)) - 1"/>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>Adds a relationship to the comments document, if none exists</xd:desc>
  </xd:doc>
  <xsl:template match="rel:Relationships[empty(rel:Relationship[@Target='comments.xml'])][exists(($SVRL//svrl:successful-report|$SVRL//svrl:failed-assert))]">
    <xsl:copy>
      <xsl:apply-templates select="@*|node()"/>
      <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships" Id="{generate-id(.)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:mode name="comments" on-no-match="shallow-skip"/>
  
  <xd:doc>
    <xd:desc>Generates comment content from SVRL errors</xd:desc>
  </xd:doc>
  <xsl:template match="svrl:successful-report|svrl:failed-assert" mode="comments">
    <w:comment w:id="{s4w:SVRL-id(.)}" w:author="QA" w:initials="S4W">
      <w:p>
        <w:r>
          <w:annotationRef/>
        </w:r>
        <xsl:apply-templates mode="#current"/>
      </w:p>
    </w:comment>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>Adds SVRL error text to comments</xd:desc>
  </xd:doc>
  <xsl:template match="svrl:text" mode="comments">
    <w:r>
      <w:t><xsl:value-of select="."/></w:t>
    </w:r>
  </xsl:template>
  
  <xd:doc>
    <xd:desc>Function to create comment ID from a given schematron error</xd:desc>
    <xd:param name="err">The SVRL error elment</xd:param>
    <xd:return>Returns an ID for the error/comment as a string.</xd:return>
  </xd:doc>
  <xsl:function name="s4w:SVRL-id" as="xs:string">
    <xsl:param name="err" as="element(*)"/>
    <xsl:if test="$err[not(self::svrl:successful-report or self::svrl:failed-assert)]">
      <xsl:sequence select="error(xs:QName('s4w:ERR001'), 'Function expects an SVRL error'), $err"/>
    </xsl:if>
    <xsl:sequence select="xs:string(index-of($errs ! generate-id(.), generate-id($err)) - 1 + count($keepComments))"/>
  </xsl:function>
  
</xsl:stylesheet>