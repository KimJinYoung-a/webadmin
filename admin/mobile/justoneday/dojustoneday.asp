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

Dim listidx , itemid , sortnum , saleper , saleprice
Dim subidx

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

saleper	   = Request("saleper")
saleprice	= Request("saleprice")


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
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_justoneday_item " &_
					" (listidx, itemid , sortnum) values " &_
					" ('" & listidx  & "'" &_
					" ,'" & itemid &"'" &_
					" ,'" & sortnum &"')"
		dbget.Execute(sqlStr)

	 Case "submodify"
		'subitem 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_justoneday_item " &_
				" Set itemid='" & itemid & "'" &_
				" 	,sortnum='" & sortnum & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" Where subidx=" & subidx
		dbget.Execute(sqlStr)

		'// 페이지정보 최종 수정자 업데이트
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_justoneday_list " + VbCrlf
		sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
		sqlStr = sqlStr + " where idx=" + cstr(listidx)

'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

	Case "add"
		'신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_justoneday_list " &_
					" (onairtitle, startdate , enddate , adminid , saleper , saleprice) values " &_
					" ('" & onairtitle &"'" &_
					" ,'" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & saleper &"'" &_
					" ,'" & saleprice &"')"
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "modify"
		'내용 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_justoneday_list " &_
				" Set onairtitle='" & onairtitle & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,saleper='" & saleper & "'" &_
				" 	,saleprice='" & saleprice & "'" &_
				" Where idx=" & idx
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "quickadd"
		'신규 등록
		Dim extratime1 , extratime2 , titledate
			
		titledate	  = dateadd("d",0,qdate)
		extratime1 = dateadd("d",0,qdate) &" 00:00:00"
		extratime2 = dateadd("d",0,qdate) &" 23:59:59"


		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_justoneday_list " &_
				" (onairtitle, startdate , enddate , adminid) values " &_
				" ('" & titledate &" 시작 justoneday 입니다'" &_
				" ,'" & extratime1 &"'" &_
				" ,'" & extratime2 &"'" &_
				" ,'" & session("ssBctId") &"')"
		dbget.Execute(sqlStr)

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
