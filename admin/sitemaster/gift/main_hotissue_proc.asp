<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : GIFT 메인 HOT ISSUE 관리
' Hieditor : 서동석 생성
'			 2022.07.08 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim vQuery, vIdx, vThemeIdx, vSubject, vSDate, vEDate, vIsUsing, vSortNo
	vIdx = requestCheckVar(getNumeric(request("idx")),10)
	vThemeIdx = Request("themeidx")
	vSubject = Request("subject")
	vSDate = Request("sdate")
	vEDate = Request("edate")
	vIsUsing = Request("isusing")
	vSortNo = Request("sortno")
	
	
	If vIdx = "" Then
		if vSubject <> "" and not(isnull(vSubject)) then
			vSubject = ReplaceBracket(vSubject)
		end If

		vQuery = "INSERT INTO [db_board].dbo.tbl_giftmain_hotissue(themeIdx,subject,startdate,enddate,isUsing,sortNo,reguserid) " & _
				 "VALUES('" & vThemeIdx & "','" & html2db(vSubject) & "','" & vSDate & "','" & vEDate & "','" & vIsUsing & "','" & vSortNo & "','" & session("ssBctId") & "')"
		dbget.execute vQuery
	Else
		if vSubject <> "" and not(isnull(vSubject)) then
			vSubject = ReplaceBracket(vSubject)
		end If

		vQuery = "UPDATE [db_board].dbo.tbl_giftmain_hotissue SET " & _
				 " themeIdx = '" & vThemeIdx & "', " & _
				 " subject = '" & html2db(vSubject) & "', " & _
				 " startdate = '" & vSDate & "', " & _
				 " enddate = '" & vEDate & " 23:59:59', " & _
				 " isUsing = '" & vIsUsing & "', " & _
				 " sortNo = '" & vSortNo & "' " & _
				 " WHERE idx = '" & vIdx & "'"
		dbget.execute vQuery
	End If
%>
<script type='text/javascript'>opener.location.reload();window.close()</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->