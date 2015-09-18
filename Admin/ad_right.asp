<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%dbdns="../"%>
<!--#include file="chk.asp"-->
<%
Function IsObjInstalled(ObjName)
	On Error ReSume Next
	Set Obj = Server.CreateObject(ObjName)
	If Err.Number<>0 Then
		Err.Clear
		IsObjInstalled = False
		Else
			IsObjInstalled = True
	End If
	On Error Goto 0
End Function

Function DbSize()
	On Error ReSume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFile(Server.MapPath(dbdns & SysSiteDbPath))
	DbSize = f.Size
	Set f = Nothing
	Set fso = Nothing
	On Error Goto 0
End Function
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理系统</title>
<link rel="stylesheet" type="text/css" href="images/Style.css">
<style type="text/css">
<!--
.STYLE1 {
	color: #0000FF;
	font-weight: bold;
	font-size: 14px;
}
.STYLE2 {
	color: #FF0000;
	font-size: 14px;
}
-->
</style>
</head>

<body>
<!--#include file="top.asp"-->
<br />
<table cellspacing="0" cellpadding="0" width="98%" align="center" border="0">
  <tbody>
    <tr>
      <td style="PADDING-LEFT: 2px; HEIGHT: 22px" 
    background="images/tab_top_bg.gif"><table cellspacing="0" cellpadding="0" width="477" border="0">
        <tbody>
          <tr>
            <td><table height="22" cellspacing="0" cellpadding="0" border="0">
              <tbody>
                <tr>
                  <td width="3"><img id="tabImgLeft__0" height="22" 
                  src="images/tab_active_left.gif" width="3" /></td>
                  <td 
                background="images/tab_active_bg.gif" class="tab"><strong class="mtitle">网站管理后台 &gt;&gt; 管理首页 </strong></td>
                  <td width="3"><img id="tabImgRight__0" height="22" 
                  src="images/tab_active_right.gif" 
            width="3" /></td>
                </tr>
              </tbody>
            </table></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr>
      <td bgcolor="#ffffff"><table cellspacing="0" cellpadding="0" width="100%" border="0">
        <tbody>
          <tr>
            <td width="1" background="images/tab_bg.gif"><img height="1" 
            src="images/tab_bg.gif" width="1" /></td>
            <td width="100%" 
          valign="top" 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><div id="tabContent__0" style="DISPLAY: block; VISIBILITY: visible">
              <table cellspacing="1" cellpadding="1" width="100%" align="center" 
            bgcolor="#8ccebd" border="0">
                <tbody>
                  <tr>
                    <td 
                style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
                valign="top" bgcolor="#fffcf7">
<fieldset style="padding:10px;">				
				<table width="100%" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="100"><span class="STYLE1">&gt;&gt;欢迎您使用 <%=SysSiteName%> 网站后台管理系统&lt;&lt;</span><br />
                            <br /><br />
                            <a href="http://ok3w.glzy8.com/book.asp" target="_blank" style="color:#006600; font-size:14px; text-decoration:underline;">如果你有什么技术问题，请点这里查看使用说明，或是在线咨询</a></td>
                      </tr>
                      <tr>
                        <td><table border="0" cellpadding="5" cellspacing="0">
                            <tbody>
                              <tr>
                                <td align="right"><strong>软件版本：</strong></td>
                                <td>Ok3w V4.8 （ASP+<%=Db_Type%> 免费版 <a href="http://ok3w.glzy8.com/a/26.html" target="_blank">点这里升级成为正式版？</a>）</td>
                              </tr>
                              <tr>
                                <td align="right"><strong>版权所有：</strong></td>
                                <td>ok3w.glzy8.com</td>
                              </tr>
							  <tr>
                                <td align="right"><strong>演示地址：</strong></td>
                                <td><a href="http://ok3w.glzy8.com/" target="_blank">http://ok3w.glzy8.com/</a></td>
                              </tr>
                              <tr>
                                <td align="right"><strong>设计开发：</strong></td>
                                <td>zhengbi</td>
                              </tr>
                              <tr>
                                <td align="right"><strong>联系QQ：</strong></td>
                                <td>124895502</td>
                              </tr>
                              <tr>
                                <td align="right"><strong>联系邮箱：</strong></td>
                                <td>zhengbi888@yahoo.com.cn</td>
                              </tr>
                            </tbody>
                        </table></td>
                      </tr>
                    </table>
</fieldset>					
                      <br />			
<fieldset style="padding:10px;">
<legend>服务器信息</legend>
<br />
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td height="25" bgcolor="#FFFFFF">&nbsp;服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
    <td height="25" bgcolor="#FFFFFF">&nbsp;脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td height="25" bgcolor="#FFFFFF">&nbsp;站点物理路径：<%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
    <td height="25" bgcolor="#FFFFFF">&nbsp;FSO文本读写：<%If Not IsObjInstalled("Scripting.FileSystemObject") Then%><font color="#FF0066"><b>×</b></font><%else%><b>√</b><%end if%></td>
  </tr>
  <tr>
    <td height="25" bgcolor="#FFFFFF">&nbsp;Jmail组件支持：<%If Not IsObjInstalled("JMail.SMTPMail") Then%><font color="#FF0066"><b>×</b></font><%else%><b>√</b><%end if%></td>
    <td height="25" bgcolor="#FFFFFF">&nbsp;CDONTS组件支持：<%If Not IsObjInstalled("CDONTS.NewMail") Then%><font color="#FF0066"><b>×</b></font><%else%><b>√</b><%end if%></td>
  </tr>
  <tr>
    <td height="25" bgcolor="#FFFFFF">&nbsp;数据库使用：<b style="color:blue"><%If Db_Type="ACC" Then%>ACCESS(<font color="red"><%=Round(DbSize()/1024^2,2)%>M</font>)<%Else%>Sql Server 2000<%End If%></b></td>
    <td height="25" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
</table>
</fieldset>
<br />
<fieldset style="padding:10px;">
<legend>升级公告</legend>
<br />
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><iframe scrolling="no" height="20" width="100%" frameborder="0" src="http://ok3w.glzy8.com/gg.asp?v=4.7"></iframe></td>
  </tr>
</table>
</fieldset>	
				
				</td>
                  </tr>
                </tbody>
              </table>
            </div></td>
            <td width="1" background="images/tab_bg.gif"><img height="1" 
            src="images/tab_bg.gif" width="1" /></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr>
      <td background="images/tab_bg.gif" bgcolor="#ffffff"><img height="1" 
      src="images/tab_bg.gif" width="1" /></td>
    </tr>
  </tbody>
</table>
<br>
</body>
</html>
<%
Call CloseConn()
Set Rs = Nothing
Set Admin = Nothing
%>

