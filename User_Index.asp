<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Const dbdns = "./"
Const Htmldns = "./"
Const Base_Target = ""
%>
<!--#include file="AppCode/Conn.asp"-->
<!--#include file="AppCode/fun/function.asp"-->
<!--#include file="AppCode/Pager.asp"-->
<!--#include file="AppCode/fun/md5.asp"-->
<!--#include file="AppCode/Class/Ok3w_User.asp"-->
<!--#include file="AppCode/Class/Ok3w_Article.asp"-->
<!--#include file="vbs.asp"-->
<%
ClassID = ""

Set User = New Ok3w_User
Set Article = New Ok3w_Article

User_Name = Replace(User.IsLogin(),"'","")
If User_Name = "" Then
	Call MessageBox("�㻹û�е�½���ǵ�½�Ѿ���ʱ��","./")
End If

a = Trim(Request.QueryString("a"))
b = Trim(Request.QueryString("b"))

If a="user_edit" And b="save" Then
	Call  User.UserEdit()
	Call MessageBox("�����ɹ�","?a=user_eidt&rnd=" & Now())
End If

If a="a_edit" And b="save" Then
	a_id = Trim(Request.QueryString("a_id"))
	Call Article.User_Article_Save(a_id)
	Call MessageBox("�����ɹ�","?a=a_list&rnd=" & Now())
End If

If a="a_del" Then
	a_id = Cdbl(Request.QueryString("a_id"))
	Call Article.User_Article_Del(a_id,User_Name)
	Call MessageBox("�����ɹ�","?a=a_list&rnd=" & Now())
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��������_<%=Application("Ok3w_SiteName")%></title>
<script language="javascript" src="js/js.js"></script>
<link href="css/css2.css" rel="stylesheet" type="text/css" />
</head>

<body>
<!--#include file="head.asp"-->
<table width="970" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#FFFFFF"><table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2" align="left" valign="top"><div class="pathnav">����ǰλ�ã�<a href="./">��վ��ҳ</a> &gt;&gt; �������� </div></td>
  </tr>
  <tr>
    <td width="200" align="left" valign="top"><table width="190" border="0" cellspacing="0" cellpadding="10">
      <tr>
        <td align="center" class="adminleft" style="font-weight: bold">����˵�</td>
      </tr>
      <tr>
        <td align="center" height="15"></td>
      </tr>
      <tr>
        <td align="center" class="adminleft"><a href="User_Index.asp">��������</a></td>
      </tr>
      <tr>
        <td align="center" height="15"></td>
      </tr>
      <tr>
        <td align="center" class="adminleft"><a href="?a=a_edit">Ͷ��/�Ƽ�����</a></td>
      </tr>
      <tr>
        <td align="center" height="15"></td>
      </tr>
      <tr>
        <td align="center" class="adminleft"><a href="?a=a_list">��������</a></td>
      </tr>
      
      <tr>
        <td align="center" height="15"></td>
      </tr>
      <tr>
        <td align="center" class="adminleft"><a href="save.asp?action=LoginOut">��ȫ�˳�</a></td>
      </tr>
    </table></td>
    <td align="left" valign="top">
	
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="cbox">
        
        <tr>
          <td valign="top"  style="padding:5px;">
<%
Select Case a
	Case "a_edit"
		a_id = Trim(Request.QueryString("a_id"))
		Call Article_Edit(a_id)
	Case "a_list"
		Call Article_List()
	Case Else
		Call user_edit()
End Select
%>
<%Private Sub user_edit()%>	  
		  <div style="padding:8px;border:1px solid #EBEBEB; background-color:#f2f6fb; font-size:14px; font-weight:bold;">��������</div>
		  <div style="border:1px solid #EBEBEB; margin-top:4px;">
		    
		      <table border="0" cellspacing="0" cellpadding="5"><form id="form1" name="form1" method="post" action="?a=user_edit&b=save">
<%
Sql = "select * from Ok3w_User where User_Name='" & User_Name & "'"
Rs.Open Sql,Conn,1,1
%>	  
<script language="javascript" src="images/DatePicker/WdatePicker.js"></script>			  
<script language="javascript">
function ChkReg(frm)
{
	if(frm.User_Password.value.trim()=="")
	{
		alert("ԭ���벻��Ϊ�գ�������");
		frm.User_Password.focus();
		return false;
	}
	
	if(frm.User_Password21.value.trim()!="")
	{
		if(frm.User_Password21.value.trim().length<6 || frm.User_Password21.value.trim().length>20)
		{
			alert("������ֻ����6-20λ�ַ�");
			frm.User_Password21.focus();
			return false;
		}
		
		if(frm.User_Password21.value != frm.User_Password22.value)
		{
			alert("�������������벻һ�£�������");
			frm.User_Password22.focus();
			return false;
		}
	}
	
	var exp2=/^[a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+\.[a-zA-Z0-9._-]*$/
	if(frm.Mail.value.match(exp2)==null)
	{
		alert("�����ʼ�����ȷ��������");
		frm.Mail.focus();
		return false;
	}
	frm.bntSubmit.disabled=true;
	
	frm.submit();
}
</script>			  
                <tr>
                  <td colspan="2" class="red12">��½��Ϣ�����</td>
                </tr>
                <tr>
                  <td align="right">�û�����</td>
                  <td><img src="<%=GetUserDengjiValue(2,GetUserDengjiID(Rs("Jifen")))%>" /> <strong><%=Rs("User_Name")%></strong> ע��ʱ�䣺<strong><%=Rs("RegTime")%></strong> ����½��<strong><%= Rs("LastLoginTime")%></strong></td>
                </tr>
                <tr>
                  <td align="right">���֣�</td>
                  <td><h1><%=Rs("Jifen")%></h1></td>
                </tr>
                <tr>
                  <td align="right">ԭ���룺</td>
                  <td><input name="User_Password" type="password" class="ginput" id="User_Password" size="20" maxlength="20" />
                    <span class="red12real">*</span>6-20λ�ַ�</td>
                </tr>
                <tr>
                  <td align="right">�����룺</td>
                  <td><input name="User_Password21" type="password" class="ginput" id="User_Password21" size="20" maxlength="20" />
                    <span class="red12real">*</span>6-20λ�ַ�������������</td>
                </tr>
                <tr>
                  <td align="right">���������룺</td>
                  <td><input name="User_Password22" type="password" class="ginput" id="User_Password22" size="20" maxlength="20" />
                    <span class="red12real">*</span></td>
                </tr>
                <tr>
                  <td align="right">���䣺</td>
                  <td><input name="Mail" type="text" class="ginput" id="Mail" value="<%=Rs("Mail")%>" size="35" maxlength="50" />
                    <span class="red12real">*</span>����������䣬������ϵ</td>
                </tr>
                <tr>
                  <td colspan="2" class="red12">������Ϣ��ѡ�</td>
                </tr>
                <tr>
                  <td align="right">������</td>
                  <td><input name="Name" type="text" class="ginput" id="Name" value="<%=Rs("Name")%>" size="10" maxlength="8" /></td>
                </tr>
                <tr>
                  <td align="right">�Ա�</td>
                  <td><input type="radio" name="Sex" value="��" <%If Rs("Sex")="��" Then%>checked="checked"<%End If%>/>
                    ��
                      <input type="radio" name="Sex" value="Ů" <%If Rs("Sex")="Ů" Then%>checked="checked"<%End If%>/>
                      Ů 
                      <input name="Sex" type="radio" value="����" <%If Rs("Sex")="����" Then%>checked="checked"<%End If%> />
                      ����</td>
                </tr>
                <tr>
                  <td align="right">���������գ�</td>
                  <td><input name="Birthday" type="text" class="ginput" id="Birthday" onClick="WdatePicker()" value="<%= Rs("Birthday")%>" size="10" maxlength="8" /></td>
                </tr>
                <tr>
                  <td align="right">�绰��</td>
                  <td><input name="Tel" type="text" class="ginput" id="Tel" value="<%=Rs("Tel")%>" size="15" maxlength="15" /></td>
                </tr>
                <tr>
                  <td align="right">QQ��</td>
                  <td><input name="QQ" type="text" class="ginput" id="QQ" value="<%=Rs("QQ")%>" size="15" maxlength="15" /></td>
                </tr>
                <tr>
                  <td align="right">��ϵ��ַ��</td>
                  <td><input name="Address" type="text" class="ginput" id="Address" value="<%=Rs("Address")%>" size="35" maxlength="50" /></td>
                </tr>
                <tr>
                  <td align="right">�ʱࣺ</td>
                  <td><input name="Zip" type="text" class="ginput" id="Zip" value="<%=Rs("Zip")%>" size="10" maxlength="6" /></td>
                </tr>
                <tr>
                  <td align="right">���ҽ��ܣ�</td>
                  <td><textarea name="Content" cols="60" rows="8" class="ginput" id="Content"><%=Rs("Content")%></textarea></td>
                </tr>
                
                <tr>
                  <td>&nbsp;</td>
                  <td><input name="bntSubmit" type="button" class="gbnt" id="bntSubmit" onClick="ChkReg(this.form);" value="Ok�������޸�" />
                    <input name="action" type="hidden" id="action" value="Reg" /></td>
                </tr>
     <%Rs.Close%>        
</form>
              </table>
	      </div>
<%End Sub%>	
<!--/////////////////////////////////////////////////////////////////////////////////////////////////////////////-->
<%Private Sub Article_List()%>
<div style="padding:8px;border:1px solid #EBEBEB; background-color:#f2f6fb; font-size:14px; font-weight:bold; margin-bottom:4px;">��������</div>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">				  
                        <tr style="border-bottom:1px solid #AACCEE; background-color:#EFEFEF; line-height:170%; text-align:center;">
                          <td style="border-bottom:1px solid #AEE1DC; background-color:#EFEFEF; line-height:170%; text-align:center;">����</td>
                          <td style="border-bottom:1px solid #AEE1DC; background-color:#EFEFEF; line-height:170%; text-align:center;">ʱ��</td>
                          <td style="border-bottom:1px solid #AEE1DC; background-color:#EFEFEF; line-height:170%; text-align:center;">״̬</td>
                          <td style="border-bottom:1px solid #AEE1DC; background-color:#EFEFEF; line-height:170%; text-align:center;">����</td>
                        </tr>
<%
Sql = "Select * from Ok3w_Article where Inputer='" & Replace(User_Name,"'","") & "' order by ID desc"
Rs.Open Sql,Conn,1,1
If Not(Rs.Eof And Rs.Bof) Then
Do While Not Rs.Eof
%>						
                        <tr bgcolor="#FFFFFF">
                          <td style="line-height:170%; font-size:14px;"><%=Rs("Title")%></td>
                          <td><%=Rs("AddTime")%></td>
                          <td><%If Rs("IsPass")=1 Then%>��ͨ<%Else%><font color="#FF0000">����</font><%End If%></td>
                          <td><%If Rs("IsPass")=0 Then%><a href="?a=a_edit&a_id=<%=Rs("ID")%>" style="text-decoration:underline;">�޸�/�鿴</a> <a href="?a=a_del&a_id=<%=Rs("ID")%>" style="text-decoration:underline;" onClick="if(!confirm('���Ҫɾ����')) return false;">ɾ��</a>
                          <%Else%>&nbsp;<%End If%></td>
                        </tr>
<%
	Rs.MoveNext
Loop
Else
%>						
                        <tr bgcolor="#FFFFFF">
                          <td colspan="4" align="center">��������</td>
                        </tr>
<%End If
Rs.Close
%>
                    </table>
<%End Sub%>
<!--///////////////////////////////////////////////////////////////////////////////////////////-->	
<%Private Sub Article_Edit(ID)
If ID<>"" Then
	Call Article.Load(Cdbl(ID))
	If Article.Inputer<>User_Name Then
		Call MessageBox("�Բ������޷��޸ı��˵�����","")
	End If
	If Article.IsPass=1 Then
		Call MessageBox("�Բ����Ѿ���ͨ�����£��������޸�","")
	End If
End If
%>
<div style="padding:8px;border:1px solid #EBEBEB; background-color:#f2f6fb; font-size:14px; font-weight:bold; margin-bottom:4px;">Ͷ��/�Ƽ�����</div>
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0">	
<form id="form1" name="form1" method="post" action="?a=a_edit&b=save&a_id=<%=ID%>">
<script language="javascript">
String.prototype.trim = function(){ return this.replace(/(^\s*)|(\s*$)/g, "");}
function ChkForm(frm)
{
	if(frm.Title.value.trim()=="")
	{
		alert("���ⲻ��Ϊ�գ�������");
		frm.Title.focus();
		return false;
	}
	if(eWebEditor1.eWebEditor.document.body.innerHTML.trim()=="")
	{
		alert("���ݲ���Ϊ�գ�������");
		eWebEditor1.eWebEditor.focus();
		return false;
	}
	
	frm.bntSubmit.disabled = true;	
	frm.submit();
}
</script>			  
                        <tr bgcolor="#FFFFFF">
                          <td align="right">&nbsp;</td>
                          <td colspan="3" style="line-height:170%;">һ������Ա��˺󣬽��ṫ����ʾ������<br />
                            ����
                            ���ͨ�����㽫�޷��ٶ����½��б༭��<br />
                          ���������ת�أ�������ע����Դ�����ߡ�</td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="right">¼����</td>
                          <td colspan="3"><input name="textfield2" type="text" disabled="disabled" class="ginput" value="<%=User_Name%>" size="25" maxlength="20" /></td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td width="70" align="right">����</td>
                          <td colspan="3"><select name="ClassID">
                            <%Call InitClassSelectOption(1,0,Article.ClassID)%>
                          </select></td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="right">����<span class="red12">*</span></td>
                          <td colspan="3"><input name="Title" type="text" class="ginput" id="Title" value="<%=Article.Title%>" size="50" maxlength="50" /></td>
              </tr>
                        

                        <tr bgcolor="#FFFFFF">
                          <td align="right">����<span class="red12">*</span></td>
                          <td colspan="3">
						  <input name="Content" type="hidden" id="Content" value="<%=Server.HTMLEncode(Replace(Article.Content,"upfiles/","../upfiles/"))%>" />
						  <input name="UpFiles" type="hidden" id="UpFiles">
						  <IFRAME ID="eWebEditor1" SRC="editor/ewebeditor.htm?id=Content&style=user&savepathfilename=UpFiles" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="300" style="border:1px solid #CCCCCC;"></IFRAME>							</td>
              </tr>
                        
                        <tr bgcolor="#FFFFFF">
                          <td align="right">��Դ</td>
                          <td><input name="ComeFrom" type="text" class="ginput" id="ComeFrom" value="<%=Article.ComeFrom%>" size="25" maxlength="50" /></td>
                          <td align="right">����</td>
                          <td><input name="Author" type="text" class="ginput" id="Author" value="<%=Article.Author%>" size="25" maxlength="50" /></td>
                        </tr>
                        
                        
                        <tr bgcolor="#FFFFFF">
                          <td align="right">&nbsp;</td>
                          <td colspan="3"><input name="bntSubmit" type="button" class="gbnt" id="bntSubmit" onClick="ChkForm(this.form);" value="Ok�������ύ"/></td>
              </tr>
			  </form>
            </table>
<%End Sub%>
<!--///////////////////////////////////////////////////////////////////////////////////////////-->						
  		  </td>
        </tr>
      </table>    </td>
  </tr>
</table></td>
  </tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>