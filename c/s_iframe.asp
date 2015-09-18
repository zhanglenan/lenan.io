<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Const dbdns = "../"
Const Htmldns = "../"
%>
<!--#include file="../AppCode/Conn.asp"-->
<!--#include file="../AppCode/fun/function.asp"-->
<!--#include file="../vbs.asp"-->
<%
cid = Trim(Request.QueryString("cid"))
If cid<>"" Then cid = Cint(cid)
rx = Cint(Request.QueryString("rx"))
cx = Cint(Request.QueryString("cx"))
lx = Cint(Request.QueryString("lx"))
cmd = CBool(Request.QueryString("cmd"))
dt = CBool(Request.QueryString("dt"))
ft = Cint(Request.QueryString("ft"))
dh = CBool(Request.QueryString("dh"))
op = Request.QueryString("op")
size = Cint(Request.QueryString("size"))
color = Request.QueryString("color")
mcolor = Request.QueryString("mcolor")
uline = CBool(Request.QueryString("uline"))
hline = Cint(Request.QueryString("hline"))
bcolor = Request.QueryString("bcolor")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Ok3w外部调用</title>
<base target="_blank" />
<style type="text/css">
body{margin:0px; padding:0px; background-color:#<%=bcolor%>;}
body,td{font-size:<%=size%>px; line-height:<%=hline%>px;}
a{color:#<%=color%>; <%If uline Then%>text-decoration:underline;<%Else%>text-decoration:none;<%End If%>}
a:hover { color: #<%=mcolor%>; <%If uline Then%>text-decoration:underline;<%Else%>text-decoration:none;<%End If%>}
</style>
</head>

<body>
<%Call Ok3w_Soft_List(cid,rx,cx,lx,cmd,False,dt,ft,dh,op)%>
</body>
</html>
<%
Set Rs = Nothing
Call CloseConn()
%>