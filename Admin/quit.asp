<%dbdns="../"%>
<!--#include file="chk.asp"-->
<%
Call SaveAdminLog("°²È«ÍË³ö")
Call CloseConn()
Set Admin = Nothing

Session("Ok3w_eWebEditor")=""
Session("Ok3w_Caiji") = ""

Response.Cookies("Ok3w")("AdminId") = ""
Response.Cookies("Ok3w")("AdminName") = ""
Response.Cookies("Ok3w")("AdminPwd") = ""
Response.Cookies("Ok3w")("GroupId") = ""
						
Response.Write("<script language=""javascript"">")
Response.Write("parent.location.href=""ad_login.asp"";")
Response.Write("</script>")
Response.End()
%>
