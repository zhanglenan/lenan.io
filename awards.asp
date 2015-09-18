<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Const dbdns = "./"
Const Htmldns = "./"
Const Base_Target = "target=""_blank"""
%>
<!--#include file="AppCode/Conn.asp"-->
<!--#include file="AppCode/fun/function.asp"-->
<!--#include file="vbs.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>awards</title>
<meta name="keywords" content="<%=Application("Ok3w_SiteKeyWords")%>">
<meta name="description" content="<%=Application("Ok3w_SiteDescription")%>">
<script language="javascript" src="js/js.js"></script>
<link href="css/css2.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	background-image: url(images/di.GIF);
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</head>

<body><!--#include file="head.htm"-->
<table width="1024" border="0" align="center" cellpadding="0" cellspacing="0" id="__01">
  <tr>
    <td> <img src="images/index_15.jpg" width="17" height="660" alt="" /></td>
    <td valign="top" bgcolor="#FFFFFF">
      <table width="100%" height="60"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="62%">&nbsp;</td>
          <td width="38%"><div align="right"><img src="images/tit3.gif" width="195" height="34" /></div></td>
        </tr>
        <tr>
          <td colspan="2"><table width="948"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="2" bgcolor="#8F8F8F"></td>
            </tr>
          </table></td>
        </tr>
      </table>
      <table width="97%" height="422"  border="0" align="center" cellpadding="10" cellspacing="0" bordercolor="#FEFEFE" bgcolor="#8F8F8F">
        <tr>
          <td height="184" valign="top" bgcolor="#FFFFFF"><table width="948"  border="0" align="center" cellpadding="10" cellspacing="1" bordercolor="#FEFEFE" bgcolor="#8F8F8F">
            <tr>
              <td bgcolor="#FFFFFF"><p class="zoom">
                <span class="box_list">
                <%Call Ok3w_Article_List(5,5,1,100,False,False,False,"",False,"")%>
                </span>              </p>                </td>
            </tr>
          </table>
          </td>
        </tr>
    </table></td>
    <td> <img src="images/index_17.jpg" width="18" height="660" alt="" /></td>
  </tr>
  <tr>
    <td colspan="3"> <img src="images/index_18.jpg" width="1024" height="5" alt="" /></td>
  </tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
