<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doMainTopKeyword.asp
' Discription : 메인 탑키워드 처리 페이지
' History : 2008.04.18 허진원 생성
'           2012.01.09 허진원 : 사이트구분 추가
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode, selIdx, SortNo, allusing, sqlStr
dim arrIdx, arrSort, lp
dim idx, itemid, comment, userid, cate_large, cate_mid

menupos		= Request("menupos")
mode		= Request("mode")
allusing	= Request("allusing")
selIdx		= Replace(Request("selIdx")," ","")
SortNo		= Replace(Request("arrSort")," ","")
idx			= Request("idx")
itemid		= html2db(Request("itemid"))
comment	= html2db(Request("comment"))
userid	= html2db(Request("userid"))
cate_large	= Format00(3,request("cate_large"))
cate_mid	= Format00(3,Request("cate_mid"))


if SortNo="" then	SortNo = html2db(Request("SortNo"))

'// 모드에 따른 분기
Select Case mode
	Case "changeUsing"
		'사용여부 일괄 변경
		if selIdx<>"" then
			sqlStr = "Update [db_sitemaster].[dbo].tbl_main_review " &_
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
				sqlStr = sqlStr & "Update [db_sitemaster].[dbo].tbl_main_review" &_
						" Set sortNo=" & arrSort(lp) &_
						" Where idx=" & arrIdx(lp) & ";" & vbCrLf
			next
			dbget.Execute(sqlStr)
		end if

	Case "add"
		'신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_main_review " &_
				" (itemid, comment, SortNo, userid, cate_large, cate_mid) values " &_
				" ('" & itemid & "'" &_
				" ,'" & comment & "'" &_
				" ,'" & SortNo & "'" &_
				" ,'" & userid & "'" &_
				" ,'" & cate_large & "'" &_
				" ,'" & cate_mid & "')"
		dbget.Execute(sqlStr)

	Case "modify"
		'내용 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_main_review " &_
				" Set itemid='" & itemid & "'" &_
				"	,comment='" & comment & "'" &_
				" 	,SortNo=" & SortNo &_
				" 	,userid='" & userid & "'" &_
				" Where idx=" & idx
		dbget.Execute(sqlStr)

End Select

%>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "main_review.asp?menupos=<%=menupos%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
