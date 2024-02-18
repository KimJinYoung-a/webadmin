<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : doMainTopKeyword.asp
' Discription : 메인 탑키워드 처리 페이지
' History : 2008.04.18 허진원 생성
'           2022.07.01 한용민 수정(isms취약점수정, 소스표준화)
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'// 변수 선언 및 파라메터 접수
dim menupos, mode, selIdx, SortNo, allusing, sqlStr, siteDiv
dim arrIdx, arrSort, lp
dim idx, keyword, linkinfo

menupos		= Request("menupos")
mode		= Request("mode")
allusing	= Request("allusing")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
idx			= Request("idx")
keyword		= html2db(Request("keyword"))
linkinfo	= html2db(Request("linkinfo"))
siteDiv		= Request("siteDiv")

if SortNo="" then	SortNo = html2db(Request("SortNo"))

'// 모드에 따른 분기
Select Case mode
	Case "changeUsing"
		'사용여부 일괄 변경
		if selIdx<>"" then
			sqlStr = "Update [db_sitemaster].[dbo].tbl_maintopKeyword " &_
					" Set isusing='" & allusing & "'" &_
					" Where idx in (" & selIdx & ")"
			dbget.Execute(sqlStr)
		end if

	Case "changeSort"
		'표시순서 일괄 변경
		if selIdx<>"" then
			arrIdx = split(selIdx,",")
			arrSort = split(SortNo,",")

			for lp=0 to ubound(arrIdx)
				sqlStr = sqlStr & "Update [db_sitemaster].[dbo].tbl_maintopKeyword " &_
						" Set sortNo=" & arrSort(lp) &_
						" Where idx=" & arrIdx(lp) & ";" & vbCrLf
			next
			dbget.Execute(sqlStr)
		end if

	Case "add"
		if keyword <> "" and not(isnull(keyword)) then
			keyword = ReplaceBracket(keyword)
		end If
		if linkinfo <> "" and not(isnull(linkinfo)) then
			linkinfo = ReplaceBracket(linkinfo)
		end If

		'신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_maintopKeyword " &_
				" (siteDiv, keyword, linkinfo, SortNo) values " &_
				" ('" & siteDiv & "'" &_
				" ,'" & keyword & "'" &_
				" ,'" & linkinfo & "'" &_
				" ," & SortNo & ")"
		dbget.Execute(sqlStr)

	Case "modify"
		if keyword <> "" and not(isnull(keyword)) then
			keyword = ReplaceBracket(keyword)
		end If
		if linkinfo <> "" and not(isnull(linkinfo)) then
			linkinfo = ReplaceBracket(linkinfo)
		end If

		'내용 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_maintopKeyword " &_
				" Set siteDiv='" & siteDiv & "'" &_
				"	,keyword='" & keyword & "'" &_
				" 	,linkinfo='" & linkinfo & "'" &_
				" 	,SortNo=" & SortNo &_
				" Where idx=" & idx
		dbget.Execute(sqlStr)

End Select

%>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "main_TopKeyword.asp?menupos=<%=menupos%>&siteDiv=<%=siteDiv%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
