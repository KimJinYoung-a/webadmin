<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/TenQuizCls.asp" -->
<%
dim mode, sqlStr

dim id
dim jukyo
dim jukyocd
dim startdate
dim enddate
dim chkDays
dim useyn
dim taskStatus
dim regdate
dim lastupdate
dim regUser
dim updateUser

dim adminName

mode = request("mode")
id	= request("id")
jukyo	= request("jukyo")
jukyocd	= request("jukyocd")
startdate	= request("startdate")
enddate	= request("enddate")
chkDays	= request("chkDays")
useyn	= request("useyn")
taskStatus	= request("taskStatus")
regdate	= request("regdate")
lastupdate	= request("lastupdate")
regUser	= request("regUser")
updateUser	= request("updateUser")

public Function GetAdminName(adminid)	
	If IsNull(adminid) Or adminid="" Then Exit Function
	On Error Resume Next
	dim SqlStr

	sqlStr = " Select top 1 username "
	sqlStr = sqlStr & " From db_partner.dbo.tbl_user_tenbyten "
	sqlStr = sqlStr & " where userid = '"& adminid &"'"
	rsget.CursorLocation = adUseClient
	rsget.CursorType=adOpenStatic
	rsget.Locktype=adLockReadOnly
	rsget.Open sqlStr, dbget

	If Not(rsget.bof Or rsget.eof) Then
		GetAdminName = rsget("username")
	End If
	rsget.close
	On Error goto 0
End Function

adminName = GetAdminName(session("ssBctId"))		

'// ��忡 ���� �б�
Select Case mode
	Case "mod"
		sqlStr = "Update DB_USER.DBO.tbl_mileage_auto_extinction_master " &_
				" 	Set jukyo ='" & jukyo & "'" &_
				" 	,jukyocd ='" & jukyocd & "'" &_
				" 	,startdate ='" & startdate & "'" &_
				" 	,enddate ='" & enddate & "'" &_
				" 	,chk_days ='" & chkDays & "'" &_
				" 	,useyn =" & useyn &_					
				" 	,task_status ='" & taskStatus & "'" &_
				" 	,lastupdate = getdate() "&_
				" 	,update_user ='" & adminName & "'" &_
				" Where id =" & id

		dbget.Execute(sqlStr)
	Case "add"
		'�ű� ���
		sqlStr = "SELECT count(*) as cnt from DB_USER.DBO.tbl_mileage_auto_extinction_master where jukyo = '" & jukyo & "' " 

		rsget.Open sqlStr, dbget, 1
		If rsget("cnt") >= 1 Then
		%>
			<script>
			<!--
				// ������� ����
				alert("������ �Ҹ� ���䰡 �����մϴ�. �ٸ� ������ �־��ּ���.");
				history.back()
			//-->
			</script>
		<%
			response.end
			rsget.Close
		End If
		
		sqlStr = "Insert Into DB_USER.DBO.tbl_mileage_auto_extinction_master " &_
					" ( jukyo, jukyocd, startdate, enddate, chk_days, useyn, reg_user, regdate" &_
					" ) values "&_
					" ('" & jukyo &"'" &_
					" ,'" & jukyocd &"'" &_
					" ,'" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,'" & chkDays & "'" &_
					" ,'" & useyn & "'" &_
					" ,'" & adminName & "'" &_										
					" , getdate()" &_										
					")"		
		dbget.Execute(sqlStr)
End Select
%>
<% If mode = "add"  Or mode = "mod" then%>
<script>
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	window.opener.document.location.href = window.opener.document.URL;    // �θ�â ���ΰ�ħ
	 self.close();        // �˾�â �ݱ�
//-->
</script>
<% Else %>
<script language="javascript">
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	self.location = "index.asp";
//-->
</script>
<% End If %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
