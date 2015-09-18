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
	font-size: 16px;
	line-height: 22px;
	font-weight: normal;
}
-->
</style>
<table width="661"  border="0" cellpadding="28" cellspacing="1" bordercolor="#FEFEFE" bgcolor="#8F8F8F">
  <tr>
    <td height="184" bgcolor="#FFFFFF" class="wen1">
      <%Call OutThisPageContent(Article.ID,Article.Content,"asp")%>
    </td>
  </tr>
</table>
