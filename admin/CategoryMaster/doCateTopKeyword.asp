<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doCateTopKeyword.asp
' Discription : 카테고리 탑키워드 처리 페이지
' History : 2008.03.31 허진원 생성
'         : 2008.10.27 중카테고리 처리 추가(허진원)
'         : 2009.04.16 관련상품 추가(허진원)
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, cdl, cdm, keyword, linkinfo, itemid

menupos		= Request("menupos")
mode		= Request("mode")
allusing	= Request("allusing")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
itemid		= Request("itemid")
idx			= Request("idx")
cdl			= Request("cdl")
if cdl="110" then
	cdm			= Request("cdm")
end if
keyword		= html2db(Request("keyword"))
linkinfo	= html2db(Request("linkinfo"))

if SortNo="" then	SortNo = html2db(Request("SortNo"))

'// 모드에 따른 분기
Select Case mode
	Case "changeUsing"
		'사용여부 일괄 변경
		if selIdx<>"" then
			sqlStr = "Update [db_sitemaster].[dbo].tbl_category_keyword " &_
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
				sqlStr = sqlStr & "Update [db_sitemaster].[dbo].tbl_category_keyword " &_
						" Set sortNo=" & arrSort(lp) &_
						" Where idx=" & arrIdx(lp) & ";" & vbCrLf
			next
			dbget.Execute(sqlStr)
		end if

	Case "add"
		'신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_category_keyword " &_
				" (cdl, cdm, keyword, linkinfo, itemid, SortNo) values " &_
				" ('" & cdl & "'" &_
				" ,'" & cdm & "'" &_
				" ,'" & keyword & "'" &_
				" ,'" & linkinfo & "'" &_
				" ,'" & itemid & "'" &_
				" ," & SortNo & ")"
		dbget.Execute(sqlStr)

	Case "modify"
		'내용 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_category_keyword " &_
				" Set cdl='" & cdl & "'" &_
				" 	,cdm='" & cdm & "'" &_
				" 	,keyword='" & keyword & "'" &_
				" 	,linkinfo='" & linkinfo & "'" &_
				" 	,itemid='" & itemid & "'" &_
				" 	,SortNo=" & SortNo &_
				" Where idx=" & idx
		dbget.Execute(sqlStr)

End Select

%>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "category_left_topKeyword.asp?menupos=<%=menupos%>&cdl=<%=cdl%>&cdm=<%=cdm%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
