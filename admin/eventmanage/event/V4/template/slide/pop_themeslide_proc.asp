<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###############################################
' PageName : pop_slide_proc.asp
' Discription : 모바일 slide process
' History : 2016-02-16 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	Dim eventid , mode , device , idx, saveafter
	Dim slideimg , linkurl , sorting , isusing , sbgimg '슬라이드 이미지
	Dim topimg , btmimg , topaddimg '템플릿 이미지들
	Dim btmYN , btmcode , btmaddimg , pcadd1 , moadd1 , gubun
	Dim sqlStr

	Dim sIdx, sSortNo, sIsUsing, i , slinkurl '//슬라이드

	idx 		= requestCheckVar(Request.form("idx"),10)
	eventid 	= requestCheckVar(Request.form("eventid"),10)
	mode 		= requestCheckVar(Request.form("mode"),6)
	device 		= requestCheckVar(Request.form("device"),1)
	slideimg 	= requestCheckVar(Request.form("slideimg"),200)
	sbgimg	 	= requestCheckVar(Request.form("bgslideimg"),200)
	linkurl 	= requestCheckVar(Request.form("linkurl"),200)

	topimg 		= requestCheckVar(Request.form("topimg"),200)
	btmimg 		= requestCheckVar(Request.form("btmimg"),200)
	topaddimg 	= requestCheckVar(Request.form("topaddimg"),200)

	btmYN		= requestCheckVar(Request.form("btmYN"),1)
	btmcode		= html2db(Request.form("btmcode"))
	btmaddimg	= requestCheckVar(Request.form("btmaddimg"),200)
	pcadd1		= requestCheckVar(Request.form("pcadd1"),200)
	moadd1		= requestCheckVar(Request.form("moadd1"),200)
	gubun		= requestCheckVar(Request.form("gubun"),200)
'Response.write mode & "<br>"
Select Case mode
	 Case "SI"
		'slide이미지 신규 등록
		sqlStr = "Insert Into db_event.dbo.tbl_event_slide_addimage " &_
					" (evt_code, device , slideimg , bgimg , linkurl ) values " &_
					" ('" & eventid  & "'" &_
					" ,'" & device &"'" &_
					" ,'" & slideimg &"'" &_
					" ,'" & sbgimg &"'" &_
					" ,'" & linkurl &"')"
		dbget.Execute(sqlStr)

		sqlStr = "if not Exists(select idx from db_event.dbo.tbl_event_slide_template where evt_code="&eventid&" and device='"&device&"')" + vbcrlf
		sqlStr = sqlStr + " begin" + vbcrlf
		sqlStr = sqlStr + "Insert Into db_event.dbo.tbl_event_slide_template " + vbcrlf
		sqlStr = sqlStr + " (evt_code, device , topimg , btmimg , topaddimg , btmYN , btmcode , btmaddimg , pcadd1 , gubun ) values " + vbcrlf
		sqlStr = sqlStr + " ('" & eventid &"'" + vbcrlf
		sqlStr = sqlStr + " ,'" & device &"'" + vbcrlf
		sqlStr = sqlStr + " ,'" & topimg &"'" + vbcrlf
		sqlStr = sqlStr + " ,'" & btmimg &"'" + vbcrlf
		sqlStr = sqlStr + " ,'" & topaddimg &"'" + vbcrlf
		sqlStr = sqlStr + " ,'" & btmYN &"'" + vbcrlf
		sqlStr = sqlStr + " ,'" & btmcode &"'" + vbcrlf
		sqlStr = sqlStr + " ,'" & btmaddimg &"'" + vbcrlf
		sqlStr = sqlStr + " ,'" & pcadd1 &"'" + vbcrlf
		sqlStr = sqlStr + " ,'" & gubun &"'" + vbcrlf
		sqlStr = sqlStr + ")" + vbcrlf
		sqlStr = sqlStr + " end" + vbcrlf
		'Response.write sqlStr
		'Response.end
		dbget.Execute(sqlStr)
		saveafter="SI"
	Case "SU"
		'//리스트에서수정
		for i=1 to request.form("chkIdx").count
			sIdx = request.form("chkIdx")(i)
			sSortNo = request.form("sort"&sIdx)
			sIsUsing = request.form("use"&sIdx)
			slinkurl = request.form("linkurl"&sIdx)
			if sSortNo="" then sSortNo="0"
			if sIsUsing="" then sIsUsing="N"

			sqlStr = sqlStr & " Update db_event.dbo.tbl_event_slide_addimage Set "
			sqlStr = sqlStr & " sorting=" & sSortNo & ""
			sqlStr = sqlStr & " ,isusing='" & sIsUsing & "'"
			sqlStr = sqlStr & " ,linkurl='" & slinkurl & "'"
			sqlStr = sqlStr & " Where idx='" & sIdx & "';" & vbCrLf
		Next

		If sqlStr <> "" then
			dbget.Execute sqlStr
		Else
			Call Alert_return("저장할 내용이 없습니다.")
			dbget.Close: Response.End
		End If 
		saveafter="SU"
	Case "SD" '삭제
		sIdx = request.form("chkIdx")

		sqlStr = "delete from db_event.dbo.tbl_event_slide_addimage Where idx='"& sIdx &"' and device = '"& device &"'"
		dbget.Execute sqlStr
End Select
%>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("<%=chkiif(mode="SD","삭제 완료.","수정/저장 완료.")%>");
	self.location = "pop_<%=chkiif(device="W","pcweb","mobile")%>_themeslide.asp?eC=<%=eventid%>&smode=<%=mode%>&saveafter=<%=saveafter%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
