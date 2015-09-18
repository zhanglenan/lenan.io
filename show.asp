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
id=myCdbl(Request.QueryString("id"))
Set Article = New Ok3w_Article
Call Article.HitsAdd(Id)
Call Article.Load(Id)
If Article.IsPass=0 Then Call Page_Err("文章已经关闭")
If Article.IsDelete=1 Then Call Page_Err("文章已经删除")
ClassID = Article.ClassID
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=Article.Title%></title>
<meta name="keywords" content="<%=Replace(Article.Keywords,"|",",")%>" />
<meta name="description" content="<%=Article.Description%>" />
<script language="JavaScript" src="./js/js.js" type="text/javascript"></script>
<link href="./css/css2.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	background-image: url(images/di.gif);
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>

<body>
<!--#include file="head.htm"-->
<table width="1024" border="0" align="center" cellpadding="0" cellspacing="0" id="__01">
  <tr>
    <td width="17" valign="top" background="images/index_15.jpg"> <img src="images/index_15.jpg" width="17" height="660" alt="" /></td>
    <td width="1110" valign="top" bgcolor="#FFFFFF"><table width="80%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="97%"  border="0" align="center" cellpadding="10" cellspacing="0" bordercolor="#FEFEFE" bgcolor="#8F8F8F">
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" style="table-layout: fixed; word-wrap:break-word;">
              <tr>
                <td><div style="text-align:center;">
                    <h1><%=Article.Title%></h1>
                  </div>
                    <div class="xlist">
                      <hr />
                      <div align="center"><%=Format_Time(Article.AddTime,1)%> &nbsp;&nbsp;by：<%=Article.ComeFrom%> </div>
                    </div>
                    <div class="zoom">
                      <%If Article.Description="" Then%>
                      <div class="content_line">
                        <!--分隔线-->
                      </div>
                      <%Else%>
                      <%End If%>
                      <%If Article.vUserGroupID=0 And Article.vUserJifen=0 Then%>
                      <%Call OutThisPageContent(Article.ID,Article.Content,"html")%>
                      <%Else%>
                      <iframe name="p_view" id="p_view" scrolling="No" src="./c/p_view.asp?id=<%=Article.ID%>" height="200" onload="SetCwinHeight(this)" width="100%" frameborder="0"></iframe>
                      <%End If%>
                      <%If Article.IsUserAdd=1 Then%>
                      <div style="color:#0000FF;">感谢<strong><%=Article.Inputer%></strong>投稿</div>
                      <%End If%>
                  </div></td>
              </tr>
            </table>
       
          </td>
        </tr>
      </table>
        <p class="zoom">&nbsp;</p></td>
  </tr>
</table>    </td>
    <td width="18" valign="top" background="images/index_17.jpg"> <img src="images/index_17.jpg" width="18" height="660" alt="" /></td>
  </tr>
  <tr>
    <td colspan="3"> <img src="images/index_18.jpg" width="1024" height="5" alt="" /></td>
  </tr>
</table><!--#include file="foot.asp"-->
</body>
</html>
