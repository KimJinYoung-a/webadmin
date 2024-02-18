<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doexhibition.asp
' Discription : exhibition 처리 페이지
' History : 2016.04.07 이종화 생성
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, exhibitiontitle, linkinfo
Dim qdate
Dim startdate , enddate , isusing

Dim listidx , itemid , sortnum , topview , linkurl
Dim subidx

menupos			= Request("menupos")
isusing			= Request("isusing")
mode			= Request("mode")
allusing		= Request("allusing")
selIdx			= Replace(Request("selIdx")," ","")
SortNo			= Replace(Request("arrSort")," ","")
idx				= getNumeric(Request("idx"))
exhibitiontitle	= html2db2017(Request("exhibitiontitle"))
linkinfo		= html2db2017(Request("linkinfo"))

qdate			= Request("prevDate")

startdate		= Request("StartDate")& " " &Request("sTm")&":00:00"
enddate			= Request("EndDate")& " " &Request("eTm")&":59:59"

listidx			= getNumeric(Request("listidx"))
subidx			= getNumeric(Request("subidx"))
itemid			= getNumeric(Request("subItemid"))
sortnum			= getNumeric(Request("sortnum"))

topview			= Request("topview") '//바로 노출여부 Y N
linkurl			= Trim(Request("linkurl")) '//링크

if SortNo="" Then SortNo = html2db2017(Request("SortNo"))

'// 모드에 따른 분기
Select Case mode
	 Case "subadd"
		'subitem 신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_exhibition_item " &_
					" (listidx, itemid , sortnum) values " &_
					" ('" & listidx  & "'" &_
					" ,'" & itemid &"'" &_
					" ,'" & sortnum &"')"
		dbget.Execute(sqlStr)

	 Case "submodify"
		'subitem 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_exhibition_item " &_
				" Set itemid='" & itemid & "'" &_
				" 	,sortnum='" & sortnum & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" Where subidx=" & subidx
		dbget.Execute(sqlStr)

		'// 페이지정보 최종 수정자 업데이트
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_exhibition_list " + VbCrlf
		sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
		sqlStr = sqlStr + " where idx=" + cstr(listidx)

'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

	Case "add"
		'신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_exhibition_list " &_
					" (exhibitiontitle, startdate , enddate , linkurl , adminid ) values " &_
					" (N'" & exhibitiontitle &"'" &_
					" ,'" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,N'" & linkurl &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					")"
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "modify"
		'내용 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_exhibition_list " &_
				" Set exhibitiontitle=N'" & exhibitiontitle & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,topview='" & topview & "'" &_
				" 	,linkurl=N'" & linkurl & "'" &_
				" Where idx=" & idx
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "quickadd"
		'신규 등록
		Dim extratime1 , extratime2 , titledate

		titledate	  = dateadd("d",0,qdate)
		extratime1 = dateadd("d",0,qdate) &" 00:00:00"
		extratime2 = dateadd("d",0,qdate) &" 23:59:59"


		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_exhibition_list " &_
				" (exhibitiontitle, startdate , enddate , adminid) values " &_
				" ('수정요망 --&gt;" & titledate &" 시작 메인 기획전 입니다'" &_
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
