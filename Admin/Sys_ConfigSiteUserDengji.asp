<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%dbdns="../"%>
<!--#include file="chk.asp"-->
<!--#include file="../AppCode/Class/Ok3w_SiteConfig.asp"-->
<!--#include file="../AppCode/fun/function.asp"-->
<%
Call CheckAdminFlag(5)

Select Case Trim(Request.Form("action"))
	Case "SiteUserDengji"
		Set SiteConfig = New Ok3w_SiteConfig
		Call SiteConfig.SiteUserDengji()
		Call SaveAdminLog("�༭ϵͳ����")
		Call CloseConn()
		Call ActionOk("Sys_ConfigSiteUserDengji.asp")
End Select
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨����ϵͳ</title>
<link rel="stylesheet" type="text/css" href="images/Style.css"></head>

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
            <td width="147"><table height="22" cellspacing="0" cellpadding="0" border="0">
              <tbody>
                <tr>
                  <td width="3"><img id="tabImgLeft__0" height="22" 
                  src="images/tab_active_left.gif" width="3" /></td>
                  <td 
                background="images/tab_active_bg.gif" class="tab"><strong class="mtitle">��Ա�ȼ�����ֹ���</strong></td>
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
            <td 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          valign="top"><div id="tabContent__0" style="DISPLAY: block; VISIBILITY: visible">
              <table cellspacing="1" cellpadding="1" width="100%" align="center" 
            bgcolor="#8ccebd" border="0">
                <tbody>
                  <tr>
                    <td 
                style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
                valign="top" bgcolor="#fffcf7"><table border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                      <form id="form1" name="form1" method="post" action="">
                        
                        
                        <tr>
                          <td align="right" bgcolor="#EBEBEB">�ȼ����</td>
                          <td bgcolor="#EBEBEB">�Զ�������</td>
                          <td bgcolor="#EBEBEB">�ﵽ�ü���������ͻ���</td>
                          <td bgcolor="#EBEBEB">�ȼ�ͼƬ</td>
                          </tr>
<%For i=1 To 12%>						  
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��<strong><%=i%></strong>��</td>
                          <td bgcolor="#FFFFFF"><input name="name" type="text" id="name" value="<%=GetUserDengjiValue(0,i)%>"></td>
                          <td bgcolor="#FFFFFF"><input name="jifen" type="text" id="jifen" size="8" value="<%=GetUserDengjiValue(1,i)%>"></td>
                          <td bgcolor="#FFFFFF"><input name="pic" type="text" id="pic" value="<%=GetUserDengjiValue(2,i)%>"></td>
                          </tr>
<%Next%>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��ע</td>
                          <td colspan="3" bgcolor="#FFFFFF">Ĭ��ʮ�������𣬲����õģ����ƴ����ռ��ɣ�����Ҫ����С��������롣</td>
                          </tr>
                        <tr>
                          <td colspan="4" align="center" bgcolor="#FFFFFF"><input name="action" type="hidden" id="action" value="SiteUserDengji">
                            <input type="submit" name="Submit" value="����"></td>
                          </tr>
                      </form>
                    </table></td>
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
</body>
</html>
<%
Call CloseConn()
Set Admin = Nothing
Set SiteConfig = Nothing
%>
