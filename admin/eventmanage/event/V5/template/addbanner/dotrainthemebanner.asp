<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###############################################
' PageName : dotrainthemebanner.asp
' Discription : I형 기차 템플릿 이미지 등록
' History : 2019.02.12 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	Dim eventid , mode , device , idx, saveafter
	Dim slideimg , sorting , isusing '슬라이드 이미지
	Dim menuidx
	Dim sqlStr, GroupItemCheck

	Dim sIdx, sSortNo, sIsUsing, i '//슬라이드

	idx = requestCheckVar(Request.form("idx"),10)
	eventid = requestCheckVar(Request.form("eventid"),10)
	mode = requestCheckVar(Request.form("mode"),6)
	device = requestCheckVar(Request.form("device"),1)
	slideimg = requestCheckVar(Request.form("slideimg"),200)
	GroupItemCheck = requestCheckVar(Request.form("GroupItemCheck"),1)
	menuidx = requestCheckvar(request("menuidx"),10)
'Response.write mode & "<br>"
Select Case mode
	 Case "TI"
		'slide이미지 신규 등록
		sqlStr = "Insert Into [db_event].[dbo].[tbl_event_multi_contents] " &_
					" (menuidx, imgurl,grouptype) values " &_
					" ('" & menuidx  & "'" &_
					" ,'" & slideimg &"','B')"
		dbget.Execute(sqlStr)
		saveafter="TI"
	Case "TU"
		'//리스트에서수정
		for i=1 to request.form("chkIdx").count
			sIdx = request.form("chkIdx")(i)
			sSortNo = request.form("sort"&sIdx)
			sIsUsing = request.form("use"&sIdx)

			if sSortNo="" then sSortNo="0"
			if sIsUsing="" then sIsUsing="N"

			sqlStr = sqlStr & " Update [db_event].[dbo].[tbl_event_multi_contents] Set "
			sqlStr = sqlStr & " viewidx=" & sSortNo & ""
			sqlStr = sqlStr & " ,isusing='" & sIsUsing & "'"
			sqlStr = sqlStr & " Where idx='" & sIdx & "';" & vbCrLf
		Next

		If sqlStr <> "" then
			dbget.Execute sqlStr
		Else
			Call Alert_return("저장할 내용이 없습니다.")
			dbget.Close: Response.End
		End If 
		saveafter="TU"
	Case "TD" '삭제
		sIdx = request.form("chkIdx")

		sqlStr = "delete from db_event.dbo.tbl_event_multi_contents Where idx='"& sIdx &"'"
		dbget.Execute sqlStr
End Select

	sqlStr = " Update [db_event].[dbo].[tbl_event_multi_contents_master]"
	sqlStr = sqlStr & " Set GroupItemType='B'"
	sqlStr = sqlStr & " ,GroupItemCheck='" + Cstr(GroupItemCheck) + "'"
	sqlStr = sqlStr & " Where idx=" & menuidx
	dbget.Execute sqlStr
%>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("<%=chkiif(mode="TD","삭제 완료.","수정/저장 완료.")%>");
	self.location = "pop_train_theme_addbanner.asp?eC=<%=eventid%>&smode=<%=mode%>&saveafter=<%=saveafter%>&menuidx=<%=menuidx%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
