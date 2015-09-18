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
<title><%=Application("Ok3w_SiteTitle")%></title>
<meta name="keywords" content="<%=Application("Ok3w_SiteKeyWords")%>">
<meta name="description" content="<%=Application("Ok3w_SiteDescription")%>">
<script language="javascript" src="js/js.js"></script>
<link href="css/css2.css" rel="stylesheet" type="text/css" />
</head>

<body>
<!--#include file="head.asp"-->
<table width="970" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td bgcolor="#FFFFFF">
    <table width="950" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-bottom:8px;">
      <tr>
        <td width="7%"><strong>站内公告：</strong></td>
        <td width="93%"><%Call Ok3w_Article_Gundong("",10,880,12,120)%></td>
      </tr>  </table></td>
    </tr>
  </table>
<table width="970" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td bgcolor="#FFFFFF">
    <table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td align="left" valign="top"><table width="305" border="0" cellpadding="0" cellspacing="0" class="box">
          <tr>
            <td class="box_title"><h3><%=GetClassName(1)%></h3>
                <span><a href="list.asp?id=1">更多&gt;&gt;</a></span></td>
          </tr>
          <tr>
            <td><%Call Ok3w_Article_IndexClassImg(1,72,72,1,12)%>
			<div class="box_list"><%Call Ok3w_Article_List(1,5,1,17,False,False,False,"",False,"")%></div></td>
          </tr>
        </table>
          <table width="305" border="0" cellpadding="0" cellspacing="0" class="box">
            <tr>
              <td class="box_title"><h3><%=GetClassName(2)%></h3>
                  <span><a href="list.asp?id=2">更多&gt;&gt;</a></span></td>
            </tr>
            <tr>
              <td><%Call Ok3w_Article_IndexClassImg(2,72,72,1,12)%>
			  <div class="box_list"><%Call Ok3w_Article_List(2,5,1,17,False,False,False,"",False,"")%></div></td>
            </tr>
          </table></td>
        <td align="center" valign="top"><div style="width:320px; height:230px; background-color:#f5fafe;" class="box">
          <%Call Ok3w_Article_ImgFlash("",320,243)%>
        </div>
          <table width="322" border="0" cellpadding="0" cellspacing="0" class="box">
            <tr>
              <td class="box_title"><h3><%=GetClassName(3)%></h3>
                  <span><a href="list.asp?id=3">更多&gt;&gt;</a></span></td>
            </tr>
            <tr>
              <td><%Call Ok3w_Article_IndexClassImg(3,72,72,1,12)%>
			  <div class="box_list">
                <%Call Ok3w_Article_List(3,5,1,17,False,False,False,"",False,"")%>
              </div></td>
            </tr>
          </table></td>
        <td align="right" valign="top">
		<table width="305" border="0" align="right" cellpadding="0" cellspacing="0" class="box">
          <tr>
            <td width="80" class="box_title"><h3><div class="hdover" id="tab1" onmouseover="settab('tab',1,2)">最近更新</div></h3></td>
            <td width="80" class="box_title"><h3><div class="hdout" id="tab2" onmouseover="settab('tab',2,2)">点击排行</div></h3></td>
            <td width="125" class="box_title">&nbsp;</td>
          </tr>
          <tr>
            <td colspan="3"><div class="box_list" style="padding-bottom:7px;">
			<div id="tab_con1" style="display:block;"><%Call Ok3w_Article_List("",20,1,16,False,False,True,0,False,"")%></div>
			<div id="tab_con2" style="display:none;"><%Call Ok3w_Article_List("",20,1,14,False,False,False,"",True,"hot")%></div>
			</div></td>
            </tr>
        </table>
		</td>
      </tr>  </table></td>
    </tr>
  </table>
<table width="970" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td bgcolor="#FFFFFF">
    <table width="950" border="0" align="center" cellpadding="0" cellspacing="0" class="box">
      <tr>
        <td align="center"><%Call Ok3w_Article_ImgGD("",1,12,935,135,142,100,False,"new",120)%></td>
      </tr>  </table></td>
    </tr>
  </table>
<table width="970" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td bgcolor="#FFFFFF">
    <table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><table width="305" border="0" cellpadding="0" cellspacing="0" class="box">
          <tr>
            <td class="box_title"><h3><%=GetClassName(4)%></h3><span><a href="list.asp?id=4">更多&gt;&gt;</a></span></td>
          </tr>
          <tr>
            <td><%Call Ok3w_Article_IndexClassImg(4,72,72,1,12)%>
			<div class="box_list"><%Call Ok3w_Article_List(4,5,1,17,False,False,False,"",False,"")%></div></td>
          </tr>
        </table></td>
        <td><table width="322" border="0" align="center" cellpadding="0" cellspacing="0" class="box">
          <tr>
            <td class="box_title"><h3><%=GetClassName(5)%></h3><span><a href="list.asp?id=5">更多&gt;&gt;</a></span></td>
          </tr>
          <tr>
            <td><%Call Ok3w_Article_IndexClassImg(5,72,72,1,12)%>
			<div class="box_list"><%Call Ok3w_Article_List(5,5,1,17,False,False,False,"",False,"")%></div></td>
          </tr>
        </table></td>
        <td><table width="305" border="0" align="right" cellpadding="0" cellspacing="0" class="box">
          <tr>
            <td class="box_title"><h3><%=GetClassName(6)%></h3><span><a href="list.asp?id=6">更多&gt;&gt;</a></span></td>
          </tr>
          <tr>
            <td><%Call Ok3w_Article_IndexClassImg(6,72,72,1,12)%>
			<div class="box_list"><%Call Ok3w_Article_List(6,5,1,17,False,False,False,"",False,"")%></div></td>
          </tr>
        </table></td>
      </tr>  </table></td>
    </tr>
  </table>
<table width="970" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td bgcolor="#FFFFFF">
    <table width="950" border="0" align="center" cellpadding="0" cellspacing="0" class="box">
      <tr>
        <td class="box_title"><h3><a href="http://www.glzy8.com/" target="_blank">友情连接</a></h3></td>
      </tr>
      <tr>
        <td><%Call  Ok3w_Site_Link(27,9,1,1)%><%Call  Ok3w_Site_Link(27,8,0,1)%>
		</td>
      </tr>  </table></td>
    </tr>
  </table>
<!--#include file="foot11.asp"-->
</body>
</html>
