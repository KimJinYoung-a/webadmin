<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : domdpick.asp
' Discription : mdpick 처리 페이지
' History : 2013.12.16 이종화 생성
' 		  : 2018.09.06 최종원 frontimage 삽입 코드 추가
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, mdpicktitle, linkinfo
Dim qdate
Dim startdate , enddate , gubun , isusing

Dim listidx , itemid , sortnum , topview
Dim subidx
dim userlevelgubun
dim frontimage, isLowestPrice

menupos	= Request("menupos")
isusing		= Request("isusing")
mode		= Request("mode")
gubun		= Request("gubun")
allusing	= Request("allusing")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
idx			= getNumeric(Request("idx"))
mdpicktitle	= html2db2017(Request("mdpicktitle"))
linkinfo	= html2db2017(Request("linkinfo"))

qdate		=	Request("prevDate")

startdate	= Request("StartDate")& " " &Request("sTm")
enddate		= Request("EndDate")& " " &Request("eTm")
frontimage		= Request("frontimage")
isLowestPrice	= Request("islowestprice")

listidx		= getNumeric(Request("listidx"))
subidx		= getNumeric(Request("subidx"))
itemid		= getNumeric(Request("subItemid"))
sortnum		= getNumeric(Request("sortnum"))

topview 	= getNumeric(Request("topview"))
userlevelgubun = getNumeric(Request("userlevelgubun"))

if SortNo="" then	SortNo = html2db2017(Request("SortNo"))

'// 모드에 따른 분기
Select Case mode
	 Case "subadd"
		'subitem 신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_mdpick_item " &_
					" (listidx, itemid , sortnum , topview , gubun, frontimg, islowestprice) values " &_
					" ('" & listidx  & "'" &_
					" ,'" & itemid &"'" &_
					" ,'" & sortnum &"'" &_
					" ,'" & topview & "'" &_
					" ,'" & gubun & "'" &_
					" ,'" & frontimage &"'" &_
					" ,'" & isLowestPrice &"')"					
		dbget.Execute(sqlStr)

	 Case "submodify"
		'subitem 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_mdpick_item " &_
				" Set itemid='" & itemid & "'" &_
				" 	,sortnum='" & sortnum & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,gubun='" & gubun & "'" &_
				" 	,topview='" & topview & "'" &_
				" 	,frontimg='" & frontimage & "'" &_
				" 	,isLowestPrice='" & isLowestPrice & "'" &_				
				" Where subidx=" & subidx
		dbget.Execute(sqlStr)

		'// 페이지정보 최종 수정자 업데이트
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_mdpick_list " + VbCrlf
		sqlStr = sqlStr + " Set lastadminid='" & session("ssBctId") & "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate() " + VbCrlf
		sqlStr = sqlStr + " where idx=" + cstr(listidx)

'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

	Case "add"
		'신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_mdpick_list " &_
					" (mdpicktitle, startdate , enddate , adminid , userlevelgubun) values " &_
					" (N'" & mdpicktitle &"'" &_
					" ,'" & startdate &"'" &_
					" ,'" & enddate &"'" &_
					" ,'" & session("ssBctId") &"'" &_
					" ,'" & userlevelgubun &"'" &_
					")"
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "modify"
		'내용 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_main_mdpick_list " &_
				" Set mdpicktitle=N'" & mdpicktitle & "'" &_
				" 	,startdate='" & startdate & "'" &_
				" 	,enddate='" & enddate & "'" &_
				" 	,lastadminid='" & session("ssBctId") & "'" &_
				" 	,lastupdate=getdate()" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,topview='" & topview & "'" &_
				" 	,userlevelgubun='" & userlevelgubun & "'" &_				
				" Where idx=" & idx
		'response.write sqlStr
		dbget.Execute(sqlStr)

	Case "quickadd"
		'신규 등록
		Dim extratime1 , extratime2 , titledate

		titledate	  = dateadd("d",0,qdate)
		extratime1 = dateadd("d",0,qdate) &" 00:00:00"
		extratime2 = dateadd("d",0,qdate) &" 23:59:59"


		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_main_mdpick_list " &_
				" (mdpicktitle, startdate , enddate , adminid) values " &_
				" ('수정요망 --&gt;" & titledate &" 시작 mdpick 입니다'" &_
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
