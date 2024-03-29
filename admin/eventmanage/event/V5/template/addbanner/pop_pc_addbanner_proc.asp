<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###############################################
' PageName : pop_pc_addbanner_proc.asp
' Discription : PC slide process
' History : 2017-12-14 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	Dim eventid , mode , idx , gubun
	Dim bimg , btitle , blink , bdate_flag , bst_date , bed_date , isusing '슬라이드 이미지
	Dim sqlStr, menuidx
	Dim sIdx, sSortNo, sIsUsing, i , sBlink , sGubun , sBtitle , sbst_date , sbed_date , sbimg , sbdate_flag '//슬라이드
	Dim sDt , eDt, viewidx

	mode 		= requestCheckVar(Request.form("mode"),6)

	idx 		= requestCheckVar(Request.form("idx"),10)
	eventid 	= requestCheckVar(Request.form("eventid"),10)
	gubun		= requestCheckVar(Request.form("gubun"),1)

	bimg 		= requestCheckVar(Request.form("bimg"),200)
	btitle 		= trim(requestCheckVar(Request.form("btitle"),200))
	blink		= Trim(requestCheckVar(Request.form("blink"),200))

	bdate_flag	= requestCheckVar(Request.form("bdate_flag"),1)
	bst_date	= requestCheckVar(Request.form("bst_date"),10)
	bed_date	= requestCheckVar(Request.form("bed_date"),10)

	isusing		= requestCheckVar(Request.form("isusing"),1)
	menuidx = requestCheckvar(request("menuidx"),16)
	if menuidx="" or isnull(menuidx) then menuidx=0

'	sDt			= requestCheckvar(request("sDt"),10) '//이벤트 시작일
'	eDt			= requestCheckvar(request("eDt"),10) '//이벤트 종료일

'//// 사용중인 이미지 갯수만 저장
Sub fnevtaddimgcnt()
	Dim imgcnt : imgcnt = 0
	sqlStr = "SELECT count(*) FROM db_event.dbo.tbl_event_pc_addbanner where evt_code = '"&eventid&"' and isusing = 'Y' and menuidx=" & menuidx
	rsget.Open sqlStr,dbget,1
	IF Not rsget.Eof Then
		imgcnt = rsget(0)
	End If
	rsget.close()

	sqlStr = "update db_event.dbo.tbl_event_display set evt_pc_addimg_cnt = "& imgcnt &" where evt_code = '"& eventid &"'" 
	dbget.Execute(sqlStr)
End sub

Select Case mode
	 Case "SI"
		'slide이미지 신규 등록
		sqlStr = "Insert Into db_event.dbo.tbl_event_pc_addbanner " &_
					" (evt_code, gubun, bimg, btitle, blink, bdate_flag, bst_date, bed_date, isusing, menuidx) values " &_
					" ('" & eventid  & "'" &_
					" ,'" & gubun &"'" &_
					" ,'" & bimg &"'" &_
					" ,'" & btitle &"'" &_
					" ,'" & blink &"'" &_
					" ,'" & bdate_flag &"'" &_
					" ,'" & bst_date &"'" &_
					" ,'" & bed_date &"'" &_
					" ,'Y'" &_
					" ,'" & menuidx &"'" &_
					")"
		dbget.Execute(sqlStr)

	    Call fnevtaddimgcnt()

		sqlStr = "IF NOT EXISTS(SELECT idx FROM db_event.dbo.tbl_event_multi_contents WHERE menuidx=" & menuidx  & " and device='W')" & vbCrLf
		sqlStr = sqlStr & "	BEGIN" & vbCrLf
		sqlStr = sqlStr & "		Insert Into db_event.dbo.tbl_event_multi_contents(menuidx, device , imgurl)" & vbCrLf
		sqlStr = sqlStr & "  	values('" & menuidx  & "','W','" & bimg &"')" & vbCrLf
		sqlStr = sqlStr & "	END"
		dbget.Execute(sqlStr)

	Case "SU"
		'//리스트에서수정
		for i=1 to request.form("chkIdx").count
			sIdx = request.form("chkIdx")(i)
			sGubun = request.form("gubun"&sIdx)
			sbimg = request.form("bimg"&sIdx)
			sBtitle = request.form("btitle"&sIdx)
			sIsUsing = request.form("isusing"&sIdx)
			sBlink = request.form("blink"&sIdx)
			sbdate_flag = request.form("bdate_flag"&sIdx)
			sbst_date = request.form("bst_date"&sIdx)
			sbed_date = request.form("bed_date"&sIdx)
			if sIsUsing="" then sIsUsing="N"
			viewidx = Request.form("viewidx")(i)

			sqlStr = sqlStr & " Update db_event.dbo.tbl_event_pc_addbanner Set "
			sqlStr = sqlStr & " gubun='" & sGubun & "'"
			sqlStr = sqlStr & " ,bimg='" & sbimg & "'"
			sqlStr = sqlStr & " ,Btitle='" & sBtitle & "'"
			sqlStr = sqlStr & " ,isusing='" & sIsUsing & "'"
			sqlStr = sqlStr & " ,blink='" & sBlink & "'"
			sqlStr = sqlStr & " ,bst_date='" & sbst_date & "'"
			sqlStr = sqlStr & " ,bed_date='" & sbed_date & "'"
			sqlStr = sqlStr & " ,bdate_flag='" & sbdate_flag & "'"
			sqlStr = sqlStr & " ,viewidx='" & viewidx & "'"
			sqlStr = sqlStr & " Where idx='" & sIdx & "';" & vbCrLf
		Next

		If sqlStr <> "" then
			dbget.Execute sqlStr

		    Call fnevtaddimgcnt()
		Else
			Call Alert_return("저장할 내용이 없습니다.")
			dbget.Close: Response.End
		End If 
	
	Case "SD" '삭제
		sIdx = request.form("chkIdx")

		sqlStr = "delete from db_event.dbo.tbl_event_pc_addbanner Where idx='"& sIdx &"'"
		dbget.Execute sqlStr

	    Call fnevtaddimgcnt()

		sqlStr = "IF NOT EXISTS(SELECT top 1 idx FROM db_event.dbo.tbl_event_pc_addbanner WHERE menuidx=" & menuidx  & ")" & vbCrLf
		sqlStr = sqlStr & "	BEGIN" & vbCrLf
		sqlStr = sqlStr & "		DELETE FROM db_event.dbo.tbl_event_multi_contents WHERE menuidx=" & menuidx  & " AND device='W'" & vbCrLf
		sqlStr = sqlStr & "	END"
		dbget.Execute(sqlStr)

End Select
%>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("<%=chkiif(mode="SD","삭제 완료.","수정/저장 완료.")%>");
	//self.location = "pop_pc_addbanner.asp?eC=<%=eventid%>&sDt=<%=sDt%>&eDt=<%=eDt%>";
	self.location = "pop_pc_addbanner.asp?eC=<%=eventid%>&menuidx=<%=menuidx%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->