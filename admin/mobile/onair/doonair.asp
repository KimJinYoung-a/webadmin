<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doMainTopKeyword.asp
' Discription : 메인 탑키워드 처리 페이지
' History : 2013.12.16 이종화 생성
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, onairtitle, linkinfo
Dim qdate
Dim startdate , enddate , gubun , isusing

Dim listidx , itemid , sortnum
Dim subidx

Dim ctitle , cper , cnum , cgubun

menupos		= Request("menupos")
isusing		= Request("isusing")
mode		= Request("mode")
gubun		= Request("gubun")
allusing	= Request("allusing")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
idx			= Request("idx")
onairtitle		= html2db(Request("onairtitle"))
linkinfo	= html2db(Request("linkinfo"))

qdate	=	Request("prevDate")

startdate			= Request("StartDate")& " " &Request("sTm")
enddate			= Request("EndDate")& " " &Request("eTm")

listidx		= Request("listidx")
subidx		= Request("subidx")
itemid		= Request("subItemid")
sortnum	= Request("sortnum")

ctitle			= Request("ctitle")
cper			= Request("cper")
cnum		= Request("cnum")
cgubun		= Request("cgubun")

if SortNo="" then	SortNo = html2db(Request("SortNo"))

''response.write gubun &"<br>"
'response.write mode &"<br>"
'response.write idx &"<br>"
'response.write itemid &"<br>"
'response.write subidx &"<br>"
'response.write sortnum &"<br>"
''response.write startdate &"<br>"
''response.write enddate &"<br>"
'response.End

'// 모드에 따른 분기
Select Case mode
	 Case "subadd"
		'subitem 신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_onair_item " &_
					" (listidx, itemid , sortnum) values " &_
					" ('" & listidx  & "'" &_
					" ,'" & itemid &"'" &_
					" ,'" & sortnum &"')"
		dbget.Execute(sqlStr)

	 Case "submodify"
		'subitem 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_onair_item " &_
				" Set itemid='" & itemid & "'" &_
				" 	,sortnum='" & sortnum & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" Where subidx=" & subidx
		dbget.Execute(sqlStr)

		'// 페이지정보 최종 수정자 업데이트
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_onair_list " + VbCrlf
		sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
		sqlStr = sqlStr + " where idx=" + cstr(listidx)

'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

	Case "add"
		'신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_onair_list " &_
					" (gubun, onairtitle, startdate , enddate , adminid , ctitle , cper , cgubun , cnum) values " &_
					" ('" & gubun  & "'" &_
					" ,'" & onairtitle &"'" &_
					" ,'" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & ctitle &"'" &_
					" ,'" & cper &"'" &_
					" ,'" & cgubun &"'" &_
					" ,'" & cnum &"')"
		dbget.Execute(sqlStr)

	Case "modify"
		'내용 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_onair_list " &_
				" Set gubun='" & gubun & "'" &_
				" 	,onairtitle='" & onairtitle & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,ctitle='" & ctitle & "'" &_
				" 	,cper='" & cper & "'" &_
				" 	,cnum='" & cnum & "'" &_
				" 	,cgubun='" & cgubun & "'" &_
				" Where idx=" & idx
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "quickadd"
		'신규 등록
		Dim qi , qd
		Dim mdate : mdate = Date
'		Dim tempdate
		Dim extratime1 , extratime2
		Dim existscatecodecnt

'		For qd = 0 To qdate-1 '// qdate는 1일부터
'
'		tempdate = dateadd("d",qd,mdate)
'
'		sqlStr = "SELECT count(*) as cnt"
'		sqlStr = sqlStr & " from db_sitemaster.dbo.tbl_mobile_main_onair_list"
'		sqlStr = sqlStr & " WHERE isusing='Y' and convert(varchar(10),startdate,120) between '"& tempdate &"' and '"& tempdate &"'"
'		
'		response.write sqlStr & "<BR>"
'		rsget.Open sqlStr, dbget, 1
'		If Not rsget.Eof then
'			existscatecodecnt = rsget("cnt")
'		End If
'		rsget.Close
'
'		If existscatecodecnt > 0 Then
'			Response.Write  "<script>"
'			Response.Write  "	alert('이미 등록되어 있는 날짜가 있습니다. 이중 등록은 불가능 합니다.');"
'			Response.Write  "	self.location = 'index.asp?menupos="&menupos&"';"
'			Response.Write  "</script>"
'			dbget.close()	:	response.End
'			Exit For '// out
'		End If
'
'		Next 

		'response.end

		'//qdate (일별수가 늘어 난다면 일수 만큼 생성)
		'For  qd = 0 To qdate-1 '// qdate는 1일부터
			For qi = 0 To 3

				If qi = 0 Then 
						extratime1 = dateadd("d",0,qdate) &" 08:00:00"
						extratime2 = dateadd("d",0,qdate) &" 11:59:59"
				ElseIf qi = 1 Then
						extratime1 = dateadd("d",0,qdate) &" 12:00:00"
						extratime2 = dateadd("d",0,qdate) &" 17:59:59"
				ElseIf qi = 2 Then
						extratime1 = dateadd("d",0,qdate) &" 18:00:00"
						extratime2 = dateadd("d",0,qdate) &" 22:59:59"
				ElseIf qi = 3 Then
						extratime1 = dateadd("d",0,qdate) &" 23:00:00"		
						extratime2 = dateadd("d",1,qdate) &" 07:59:59" '익일 7시까지
				End If 

			sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_onair_list " &_
					" (gubun, onairtitle, startdate , enddate , adminid) values " &_
					" ('" & qi+1 & "'" &_
					" ,'" & extratime1 &" 시작 onair 입니다'" &_
					" ,'" & extratime1 &"'" &_
					" ,'" & extratime2 &"'" &_
					" ,'" & session("ssBctId") &"')"
			dbget.Execute(sqlStr)
			Next 
		'Next 

End Select

%>
<% If mode = "subadd"  Or mode = "submodify" then%>
<script>
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	window.opener.document.location.href = window.opener.document.URL;    // 부모창 새로고침
	 self.close();        // 팝업창 닫기
//-->
</script>
<% Else %>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "index.asp?menupos=<%=menupos%>&prevDate=<%=qdate%>";
//-->
</script>
<% End If %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
