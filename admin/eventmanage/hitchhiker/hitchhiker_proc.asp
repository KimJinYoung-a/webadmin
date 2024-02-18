<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	Dim idx, evt_code, mevt_code, Hvol, startdate, enddate, isusing, query, notidx, delidate
	idx 		= Request("idx")
	evt_code	= Request("evt_code")
	mevt_code	= Request("m_evt_code")
	Hvol		= Request("Hvol")
	startdate	= requestCheckvar(Request("startdate"),10)
	enddate		= requestCheckvar(Request("enddate"),10)
	isusing		= requestCheckvar(Request("isusing"),1)
	delidate		= Request("delidate")
	enddate = enddate & " 23:59:59"

	If idx <> "" Then
		query = "UPDATE db_event.dbo.tbl_vip_hitchhiker " & _
				 "		SET " & _
				 "			Hvol = '" & Hvol & "', " & _
				 "			evt_code = '" & evt_code & "', " & _
				 "			mevt_code = '" & mevt_code & "', " & _
				 "			startdate = '" & startdate & "', " & _
				 "			enddate = '" & enddate & "', " & _
				 "			delidate = '" & delidate & "', " & _
				 "			isusing = '" & isusing & "' " & _
				 "	WHERE idx = '" & idx & "' "
		dbget.execute query

		'나머지는 사용안함 처리
		If isusing = "Y" Then
			query = "UPDATE db_event.dbo.tbl_vip_hitchhiker " & _
					 "		SET " & _
					 "			isusing = 'N' " & _
					 "	WHERE idx <> '" & idx & "' "
			dbget.execute query
		End If
	Else
		query = "INSERT INTO db_event.dbo.tbl_vip_hitchhiker(Hvol, evt_code, mevt_code, startdate, enddate, delidate, regdate) " & _
				 "	VALUES('" & Hvol & "','" & evt_code & "','" & mevt_code & "','" & startdate & "','" & enddate & "','" & delidate & "', '" & date() & "')"
		dbget.execute query
	End If

	Response.Write "<script>alert('저장되었습니다.');opener.location.reload();window.close();</script>"
	dbget.close()
	Response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->