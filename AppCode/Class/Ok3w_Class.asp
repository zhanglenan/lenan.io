<%
Class Ok3w_Class
    Public ID
    Public ChannelID
    Public SortName
    Public ParentID
    Public SortPath
    Public OrderID
	Public IsPic
	Public PageSize
	Public IsNav
	Public gotoURL
	
	Public Function IsHaveNextClass(ClassId)
		Sql = "select count(ID) from Ok3w_Class where ParentID=" & ClassId
		If Conn.Execute(Sql)(0)=0 Then 
			IsHaveNextClass = False
			Else
				IsHaveNextClass = True
		End If
	End Function
	
    '��Ӽ�¼
    Public Sub Add()
        Call GetFormData()
'        Sql = "select count(*) from Ok3w_Class where SortName='" & SortName & "' and ParentID=" & ParentID
'        If Conn.Execute(Sql)(0)<>0 Then
'			Session("ErrMsg") = "��ͬ�ķ����Ѿ����ڣ��벻Ҫ�ظ���ӡ�"
'			Call ActionOk("Class_Manage.asp")
'        End If
        Set oRs = Server.CreateObject("Adodb.RecordSet")
        Sql = "select * from Ok3w_Class where 1=2"
        oRs.Open Sql,Conn,1,3
        oRs.AddNew()
        Call UpdateRs(oRs)
        oRs.Update()
        oRs.Close()
        Set oRs = Nothing
		
		ID = Conn.Execute("select max(ID) from Ok3w_Class")(0)
		Sql = "update Ok3w_Class set SortPath=SortPath+'" & ID & ",' where ID=" & ID
		Conn.Execute Sql
    End Sub
	
    '�޸ļ�¼
    Public Sub Edit()
        Call GetFormData()
'        Sql = "select count(*) from Ok3w_Class where SortName='" & SortName & "'  and ParentID=" & ParentID & " and ID<>" & ID
'        If Conn.Execute(Sql)(0)<>0 Then
'           	Session("ErrMsg") = "��ͬ�ķ����Ѿ����ڣ��벻Ҫ�ظ���ӡ�"
'			Call ActionOk("Class_Manage.asp")
'        End If
        Set oRs = Server.CreateObject("Adodb.RecordSet")
        Sql = "select * from Ok3w_Class where ID= " & ID
        oRs.Open Sql,Conn,1,3
        Call UpdateRs(oRs)
        oRs.Update
        oRs.Close
        Set oRs = Nothing
    End Sub
	
    'ɾ����¼
    Public Sub Del()
		Call GetFormData()
		Sql = "delete from Ok3w_Class where  ID=" & ID
		Conn.Execute Sql
		
		Sql = "delete from Ok3w_Class where  SortPath like '%," & ID & ",%'"
		Conn.Execute Sql
		
		Sql = "delete from Ok3w_Article where  SortPath like '%," & ID & ",%'"
		Conn.Execute Sql
    End Sub
	
    '���ձ�
    Private Sub GetFormData()
		ID = Request.Form("ID")
		If ID="" Then ID = GetMaxClassID() + 1
        ChannelID = Request.Form("ChannelID")
		OrderID = Request.Form("OrderID")
		ParentID = Request.Form("ParentID")
        SortName = Request.Form("SortName")
        SortPath = Request.Form("SortPath")
		IsPic = Request.Form("IsPic")
		If IsPic = "" Then IsPic = 0
		PageSize = Request.Form("PageSize")
		If PageSize = "" Then PageSize = 20
		IsNav = Request.Form("IsNav")
		If IsNav = "" Then IsNav = 0
		gotoURL = Request.Form("gotoURL")
    End Sub
	
    '���¼�¼��
    Private Sub UpdateRs(ByRef Rs)
		Rs("ID") = ID
        Rs("ChannelID") = ChannelID
        Rs("SortName") = SortName
        Rs("ParentID") = ParentID
        Rs("SortPath") = SortPath
        Rs("OrderID") = OrderID
		Rs("IsPic") = IsPic
		Rs("PageSize") = PageSize
		Rs("IsNav") = IsNav
		Rs("gotoURL") = gotoURL
    End Sub
	
    '�Ӽ�¼���ж�����
    Private Sub GetRs(ByRef Rs)
        ID = Rs("ID")
        ChannelID = Rs("ChannelID")
        SortName = Rs("SortName")
        ParentID = Rs("ParentID")
        SortPath = Rs("SortPath")
        OrderID = Rs("OrderID")
		IsPic = Rs("IsPic")
		PageSize = Rs("PageSize")
		IsNav = Rs("IsNav")
		gotoURL = Rs("gotoURL")
    End Sub
	
	Public Function GetMaxClassID()
        Set clsRs = Server.CreateObject("Adodb.RecordSet")
        Sql = "select max(ID) from Ok3w_Class"
        clsRs.Open Sql,Conn,0,1
        If IsNull(clsRs(0)) Then
        GetMaxClassID = 0
            Else
                GetMaxClassID = clsRs(0)
        End If
        clsRs.Close
        Set clsRs = Nothing
    End Function
	
	Public Function GetMaxClassOrder(ParentID)
        Set clsRs = Server.CreateObject("Adodb.RecordSet")
        Sql = "select max(OrderID) from Ok3w_Class where ParentID=" & ParentID
        clsRs.Open Sql,Conn,0,1
        If IsNull(clsRs(0)) Then
        GetMaxClassOrder = 0
            Else
                GetMaxClassOrder = clsRs(0)
        End If
        clsRs.Close
        Set clsRs = Nothing
    End Function
End Class
%>	