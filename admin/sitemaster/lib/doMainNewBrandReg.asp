<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###############################################
' PageName : main_manager.asp
' Discription : ����Ʈ ���� ����
' History : 2008.04.11 ������ : �Ǽ������� ����
'			2009.04.19 �ѿ�� 2009�� �°� ����
'           2009.12.21 ������ : ���ں� �÷��� ���� ��� �߰�
'			2012.02.08 ������ : �̴ϴ޷� ��ü
'           2013.09.28 ������ : 2013������ - �߰����� �ʵ� �߰�
'           2015.04.07 ������ : 2015������ - �߰����� �ʵ� �߰�
'           2018-01-15 ����ȭ : ���� PC��� ���� �߰�
'			2019.08.23 �ѿ�� ����(���� �귣�尡 ����ϵǰ�, �귣��� ���۽�Ʈ���� ���� ������. üũ/���/���� ��������)
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim idx, MainCopy, BrandID, Main_Image, BrandName
Dim StartDate, EndDate, DispOrder, Isusing, mode
	
	
	idx = requestCheckVar(request("idx"),10)
	MainCopy = requestCheckVar(request("MainCopy"),128)
	BrandID = requestCheckVar(request("BrandID"),32)
	Main_Image = requestCheckVar(request("Main_Image"),128)
	StartDate = requestCheckVar(request("StartDate"),10)
	EndDate = requestCheckVar(request("EndDate"),10)
	DispOrder = requestCheckVar(request("DispOrder"),3)
	Isusing = requestCheckVar(request("Isusing"),1)

	if idx="" then idx=0
	If idx=0 Then
	mode = "add"
	Else
	mode = "edit"
	End If

dim sqlStr, Evt_Img1, Evt_Img2, Evt_Img3
dim referer
	referer = request.ServerVariables("HTTP_REFERER")

If BrandID <> "" Then
	BrandID = trim(BrandID)
	sqlStr = "select top 1 socname from [db_user].[dbo].[tbl_user_c]"
	sqlStr = sqlStr + " where userid='" + CStr(BrandID) + "'"

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF Or rsget.BOF) Then
		BrandName =  rsget("socname")
	else
		response.write "<script>alert('�귣��ID�� �߸��Ǿ����ϴ�.\n�ٽ� Ȯ���� �ּ���.');</script>"
		response.write "<script>self.location.replace('"& referer &"')</script>"
		rsget.close() : dbget.close() : response.end
	End If
	rsget.Close
else
	response.write "<script>alert('�귣��ID�� �Է����ּ���.');</script>"
	response.write "<script>self.location.replace('"& referer &"')</script>"
	dbget.close() : response.end
End If

if (mode = "add") then
    sqlStr = " insert into [db_sitemaster].[dbo].[tbl_main_new_brand]" + VbCrlf
    sqlStr = sqlStr & " (BrandID, BrandName, MainCopy, Main_Image, StartDate, EndDate, DispOrder, Isusing, RegUser)" + VbCrlf
    sqlStr = sqlStr & "		select" & vbcrlf
	sqlStr = sqlStr & "		userid,socname,'" + MainCopy + "','" + Main_Image + "','" + StartDate + "'" + VbCrlf
	sqlStr = sqlStr & "		,'" + EndDate + "','" + DispOrder + "','" + Isusing + "','" +  session("ssBctCname") + "'" + VbCrlf
	sqlStr = sqlStr & "		from [db_user].[dbo].[tbl_user_c]" & vbcrlf
	sqlStr = sqlStr & "		where userid = '"& BrandID &"'" & vbcrlf

	'response.write sqlStr & "<br>"
	'response.end
    dbget.Execute sqlStr

	sqlStr = "select IDENT_CURRENT('[db_sitemaster].[dbo].[tbl_main_new_brand]') as idx"

	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not Rsget.Eof then
		idx = rsget("idx")
	end if
	rsget.close

elseif mode = "edit" then
	sqlStr = "update b" & vbcrlf
	sqlStr = sqlStr & " set b.BrandID=c.userid" + VbCrlf
	sqlStr = sqlStr & " ,b.BrandName=c.socname" + VbCrlf
	sqlStr = sqlStr & " ,b.MainCopy='" + MainCopy + "'" + VbCrlf
	sqlStr = sqlStr & " ,b.Main_Image='" + Main_Image + "'" + VbCrlf
	sqlStr = sqlStr & " ,b.StartDate='" + StartDate + "'" + VbCrlf
	sqlStr = sqlStr & " ,b.EndDate='" + EndDate + "'" + VbCrlf
	sqlStr = sqlStr & " ,b.DispOrder='" + DispOrder + "'" + VbCrlf
	sqlStr = sqlStr & " ,b.Isusing='" + Isusing + "'" + VbCrlf
	sqlStr = sqlStr & " ,b.LastUser='" + session("ssBctCname") + "'" + VbCrlf
	sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_main_new_brand] b" & VbCrlf
	sqlStr = sqlStr & " join [db_user].[dbo].[tbl_user_c] c" & VbCrlf
	sqlStr = sqlStr & " 	on b.BrandID = c.userid" & VbCrlf
	sqlStr = sqlStr & " where idx="& idx &"" & VbCrlf

	'response.write sqlStr & "<br>"
	'response.end
	dbget.Execute sqlStr
end if

	response.write "<script>alert('����Ǿ����ϴ�.');</script>"
	response.write "<script>opener.location.reload();self.close();</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->