<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%dbdns="../"%>
<!--#include file="chk.asp"-->
<!--#include file="../AppCode/Class/Ok3w_SiteConfig.asp"-->
<!--#include file="../AppCode/fun/function.asp"-->
<%
Call CheckAdminFlag(1)

Select Case Trim(Request.Form("action"))
	Case "edit"
		Set SiteConfig = New Ok3w_SiteConfig
		Call SiteConfig.Edit()
		Call SaveAdminLog("�༭ϵͳ����")
		Call CloseConn()
		Call ActionOk("Sys_Config.asp")
End Select

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
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨����ϵͳ</title>
<link rel="stylesheet" type="text/css" href="images/Style.css">
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
            <td width="147"><table height="22" cellspacing="0" cellpadding="0" border="0">
              <tbody>
                <tr>
                  <td width="3"><img id="tabImgLeft__0" height="22" 
                  src="images/tab_active_left.gif" width="3" /></td>
                  <td 
                background="images/tab_active_bg.gif" class="tab"><strong class="mtitle">��̨����ϵͳ</strong></td>
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
                valign="top" bgcolor="#fffcf7"><table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                      <form id="form1" name="form1" method="post" action="">
                        <tr>
                          <td colspan="2" bgcolor="#EBEBEB"><strong>��վ��������</strong></td>
                        </tr>
                        <tr>
                          <td width="18%" align="right" bgcolor="#FFFFFF">��վ�ر�</td>
                          <td width="82%" bgcolor="#FFFFFF"><input name="SiteIsClose" type="checkbox" id="SiteIsClose" value="1" <%If Application("Ok3w_SiteIsClose")="1" Then%>checked="checked"<%End If%> /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��վ�ر�ԭ��</td>
                          <td bgcolor="#FFFFFF"><input name="SiteCloseNote" type="text" value="<%=Application("Ok3w_SiteCloseNote")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td colspan="2" align="left" bgcolor="#EBEBEB"><strong>��վ��������</strong></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��վ����</td>
                          <td bgcolor="#FFFFFF"><input name="SiteName" type="text" value="<%=Application("Ok3w_SiteName")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��վ����</td>
                          <td bgcolor="#FFFFFF"><input name="SiteTitle" type="text" value="<%=Application("Ok3w_SiteTitle")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��վ�ؼ���</td>
                          <td bgcolor="#FFFFFF"><input name="SiteKeyWords" type="text" value="<%=Application("Ok3w_SiteKeyWords")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��վ����</td>
                          <td bgcolor="#FFFFFF"><input name="SiteDescription" type="text" value="<%=Application("Ok3w_SiteDescription")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��վ��ַ</td>
                          <td bgcolor="#FFFFFF"><input name="SiteUrl" type="text" value="<%=Application("Ok3w_SiteUrl")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">&nbsp;TCP/IP������</td>
                          <td bgcolor="#FFFFFF"><input name="SiteTCPIP" type="text" id="SiteTCPIP" value="<%=Application("Ok3w_SiteTCPIP")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">�����������</td>
                          <td bgcolor="#FFFFFF"><input name="SiteIsGuest" type="checkbox" id="SiteIsGuest" value="1" <%if Application("Ok3w_SiteIsGuest")="1" then%>checked="checked"<%End If%> />
                            ��Ҫ����˺���ʾ</td>
                        </tr>
                        
                        
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">ȫ�ֹؼ���</td>
                          <td bgcolor="#FFFFFF"><textarea name="SitePublicKeyWords" cols="80" rows="5" id="SitePublicKeyWords"><%=Application("Ok3w_SitePublicKeyWords")%></textarea>
                            <br>
                            һ��һ�����ؼ��ʺ������á�|���ָ�</td>
                        </tr>
                        <tr style="display:none">
                          <td align="right" bgcolor="#FFFFFF">��վ����Ŀ¼</td>
                          <td bgcolor="#FFFFFF"><input name="SitePath" type="text" id="SitePath" value="<%=Application("Ok3w_SitePath")%>" size="20" /></td>
                        </tr>
                        <tr style="display:none">
                          <td align="right" bgcolor="#FFFFFF">���ݿ�����Ŀ¼</td>
                          <td bgcolor="#FFFFFF"><input name="SiteDbPath" type="text" id="SiteDbPath" value="<%=Application("Ok3w_SiteDbPath")%>" size="20" /></td>
                        </tr>
                        <tr style=" display:none">
                          <td align="right" bgcolor="#FFFFFF">��̨��ַ</td>
                          <td bgcolor="#FFFFFF"><input name="SiteAdminPath" type="text" value="<%=Application("Ok3w_SiteAdminPath")%>" size="20" /></td>
                        </tr>
                        <tr>
                          <td colspan="2" bgcolor="#EBEBEB"><strong>��ϵ��ʽ����</strong></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��˾����</td>
                          <td bgcolor="#FFFFFF"><input name="SiteCoName" type="text" id="SiteCoName" value="<%=Application("Ok3w_SiteCoName")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��ϵ��</td>
                          <td bgcolor="#FFFFFF"><input name="SiteLinkMan" type="text" id="SiteLinkMan" value="<%=Application("Ok3w_SiteLinkMan")%>" size="20" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">�绰</td>
                          <td bgcolor="#FFFFFF"><input name="SiteTel" type="text" value="<%=Application("Ok3w_SiteTel")%>" size="20" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">����</td>
                          <td bgcolor="#FFFFFF"><input name="SiteFax" type="text" value="<%=Application("Ok3w_SiteFax")%>" size="20" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��ַ</td>
                          <td bgcolor="#FFFFFF"><input name="SiteAddress" type="text" value="<%=Application("Ok3w_SiteAddress")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">�ʱ�</td>
                          <td bgcolor="#FFFFFF"><input name="SiteZip" type="text" value="<%=Application("Ok3w_SiteZip")%>" size="10" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">����</td>
                          <td bgcolor="#FFFFFF"><input name="SiteEmail" type="text" value="<%=Application("Ok3w_SiteEmail")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">QQ</td>
                          <td bgcolor="#FFFFFF"><input name="SiteQQ" type="text" value="<%=Application("Ok3w_SiteQQ")%>" size="20" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
                          <td bgcolor="#FFFFFF"><input name="SiteUserDengji" type="hidden" id="SiteUserDengji" value="<%=Application("Ok3w_SiteUserDengji")%>">
                          <input name="action" type="hidden" id="action" value="edit" />
                              <input type="submit" name="Submit" value="�� ��" style="font-size:14px;" /></td>
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

