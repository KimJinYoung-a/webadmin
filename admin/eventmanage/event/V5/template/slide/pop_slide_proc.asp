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
	Dim eventid , mode , device , idx
	Dim slideimg , linkurl , sorting , isusing , sbgimg '슬라이드 이미지
	Dim topimg , btmimg , topaddimg '템플릿 이미지들
	Dim btmYN , btmcode , btmaddimg , pcadd1 , moadd1 , gubun , isarrow
	Dim videoSize, videoLink '동영상 관련
	Dim sqlStr, menuidx, contentsAlign

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

	videoSize	= requestCheckVar(Request.form("videosize"),1)
	videoLink	= requestCheckVar(Request.form("videolink"),250)

	isarrow		= requestCheckVar(Request.form("isarrow"),1)
	menuidx = requestCheckvar(request.form("menuidx"),16)
	if menuidx="" or isnull(menuidx) then menuidx=0

Select Case mode
	 Case "SI"
		'slide이미지 신규 등록
		sqlStr = "Insert Into db_event.dbo.tbl_event_slide_addimage " &_
					" (evt_code, device, slideimg, bgimg, linkurl, menuidx) values " &_
					" ('" & eventid  & "'" &_
					" ,'" & device &"'" &_
					" ,'" & slideimg &"'" &_
					" ,'" & sbgimg &"'" &_
					" ,'" & linkurl &"'" &_
					" ,'" & menuidx &"')"
		dbget.Execute(sqlStr)
		sqlStr = "IF NOT EXISTS(SELECT idx FROM db_event.dbo.tbl_event_multi_contents WHERE menuidx=" & menuidx  & " and device='"& device &"')" & vbCrLf
		sqlStr = sqlStr & "	BEGIN" & vbCrLf
		sqlStr = sqlStr & "		Insert Into db_event.dbo.tbl_event_multi_contents(menuidx, device , imgurl)" & vbCrLf
		sqlStr = sqlStr & "  	values('" & menuidx  & "','" & device &"','" & slideimg &"')" & vbCrLf
		sqlStr = sqlStr & "	END"
		dbget.Execute(sqlStr)
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
	
	Case "SD" '삭제
		sIdx = request.form("chkIdx")

		sqlStr = "delete from db_event.dbo.tbl_event_slide_addimage Where idx='"& sIdx &"' and device = '"& device &"'"
		dbget.Execute sqlStr
		sqlStr = "IF NOT EXISTS(SELECT top 1 idx FROM db_event.dbo.tbl_event_slide_addimage WHERE menuidx=" & menuidx  & " AND device='"& device &"')" & vbCrLf
		sqlStr = sqlStr & "	BEGIN" & vbCrLf
		sqlStr = sqlStr & "		DELETE FROM db_event.dbo.tbl_event_multi_contents WHERE menuidx=" & menuidx  & " AND device='"& device &"'" & vbCrLf
		sqlStr = sqlStr & "	END"
		dbget.Execute(sqlStr)
	Case "I"
		'template 신규 등록
		sqlStr = "Insert Into db_event.dbo.tbl_event_slide_template " &_
					" (evt_code, device, topimg, btmimg, topaddimg, btmYN, btmcode, btmaddimg, pcadd1, gubun, videosize, videolink, isarrow, menuidx) values " &_
					" ('" & eventid &"'" &_
					" ,'" & device &"'" &_
					" ,'" & topimg &"'" &_
					" ,'" & btmimg &"'" &_
					" ,'" & topaddimg &"'" &_
					" ,'" & btmYN &"'" &_
					" ,'" & btmcode &"'" &_
					" ,'" & btmaddimg &"'" &_
					" ,'" & pcadd1 &"'" &_
					" ,'" & gubun &"'" &_
					" ,'" & videoSize &"'" &_
					" ,'" & videoLink &"'" &_								
					" ,'" & isarrow &"'" &_
					" ," & menuidx &"" &_
					")"
		dbget.Execute(sqlStr)

	Case "U"
		'내용 수정
		sqlStr = "Update db_event.dbo.tbl_event_slide_template " &_
				" Set topimg='" & topimg & "'" &_
				" 	,btmimg='" & btmimg & "'" &_
				" 	,topaddimg='" & topaddimg & "'" &_
				" 	,btmYN='" & btmYN & "'" &_
				" 	,btmcode='" & btmcode & "'" &_
				" 	,btmaddimg='" & btmaddimg & "'" &_
				" 	,pcadd1='" & pcadd1 & "'" &_
				" 	,moadd1='" & moadd1 & "'" &_
				" 	,gubun='" & gubun & "'" &_
				" 	,videosize='" & videoSize & "'" &_
				" 	,videolink='" & videoLink & "'" &_						
				" 	,isarrow='" & isarrow & "'" &_
				" Where idx='" & idx &"' and evt_code = '"& eventid &"' and device = '"& device &"'"
		dbget.Execute(sqlStr)
	
End Select

if gubun="1" or gubun="2" then
	contentsAlign = "2"
else
	contentsAlign = "1"
end if

'--3.theme 수정
sqlStr = "UPDATE [db_event].[dbo].[tbl_event_md_theme]" & vbCrlf
sqlStr = sqlStr + " SET contentsAlign='" & contentsAlign & "'" & vbCrlf
sqlStr = sqlStr + " where evt_code=" & eventid
dbget.execute sqlStr

%>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("<%=chkiif(mode="SD","삭제 완료.","수정/저장 완료.")%>");
	self.location = "pop_<%=chkiif(device="W","pcweb","mobile")%>_slide.asp?eC=<%=eventid%>&menuidx=<%=menuidx%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->