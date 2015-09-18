<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Const dbdns = "./"
Const Htmldns = "./"
Const Base_Target = ""
%>
<!--#include file="AppCode/Conn.asp"-->
<!--#include file="AppCode/fun/function.asp"-->
<!--#include file="AppCode/Class/Ok3w_Article.asp"-->
<!--#include file="vbs.asp"-->
<%
ID=myCdbl(Request.QueryString("id"))
Set Article = New Ok3w_Article
Call Article.HitsAdd(ID)
Call Article.Load(ID)
ClassID = ""
%>
<style type="text/css">
<!--
.wen1 {
	font-family: "Times New Roman", Times, serif;
	font-size: 14px;
	line-height: 28px;
	font-weight: bold;
}
-->
</style>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0" class="wen1">
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="86%"><%Call OutThisPageContent(Article.ID,Article.Content,"asp")%></td>
  </tr>
</table>


