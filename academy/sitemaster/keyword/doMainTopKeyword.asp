<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<%
'###############################################
' PageName : doMainTopKeyword.asp
' Discription : 메인 탑키워드 처리 페이지
' History : 2009.09.16 한용민 10x10어드민 이전후 변경
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode, selIdx, SortNo, allusing, sqlStr ,arrIdx, arrSort, lp , keyword_gubun
dim idx, keyword, linkinfo
	menupos		= RequestCheckvar(Request("menupos"),10)
	mode		= RequestCheckvar(Request("mode"),16)
	allusing	= RequestCheckvar(Request("allusing"),1)
	keyword_gubun	= RequestCheckvar(Request("keyword_gubun"),10)
	selIdx		= Replace(Request("selIdx")," ","")
	SortNo		= Replace(Request("arrSort")," ","")
	idx			= RequestCheckvar(Request("idx"),10)
	keyword		= html2db(RequestCheckvar(Request("keyword"),32))
	linkinfo	= html2db(Request("linkinfo"))
  	if selIdx <> "" then
		if checkNotValidHTML(selIdx) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if SortNo <> "" then
		if checkNotValidHTML(SortNo) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	if SortNo="" then	SortNo = html2db(Request("SortNo"))

'// 모드에 따른 분기
Select Case mode
	Case "changeUsing"
		'사용여부 일괄 변경
		if selIdx<>"" then
			sqlStr = "Update [db_academy].[dbo].tbl_maintopKeyword " &_
					" Set isusing='" & allusing & "'" &_
					" Where idx in (" & selIdx & ")"
			dbacademyget.Execute(sqlStr)
		end if

	Case "changeSort"
		'표시순서 일괄 변경
		if selIdx<>"" then
			arrIdx = split(selIdx,",")
			arrSort = split(SortNo,",")

			for lp=0 to ubound(arrIdx)
				sqlStr = sqlStr & "Update [db_academy].[dbo].tbl_maintopKeyword " &_
						" Set sortNo=" & arrSort(lp) &_
						" Where idx=" & arrIdx(lp) & ";" & vbCrLf
			next
			dbacademyget.Execute(sqlStr)
		end if

	Case "add"
		'신규 등록
		sqlStr = "Insert Into [db_academy].[dbo].tbl_maintopKeyword " &_
				" (keyword, keyword_gubun,linkinfo, SortNo) values " &_
				" ('" & keyword & "'" &_
				" ," & keyword_gubun & "" &_
				" ,'" & linkinfo & "'" &_
				" ," & SortNo & ")"
		
		'response.write sqlStr &"<Br>"
		dbacademyget.Execute(sqlStr)

	Case "modify"
		'내용 수정
		sqlStr = "Update [db_academy].[dbo].tbl_maintopKeyword " &_
				" Set keyword='" & keyword & "'" &_
				" 	,keyword_gubun='" & keyword_gubun & "'" &_
				" 	,linkinfo='" & linkinfo & "'" &_				
				" 	,SortNo=" & SortNo &_
				" Where idx=" & idx

		'response.write sqlStr &"<Br>"				
		dbacademyget.Execute(sqlStr)

End Select

%>

<script language="javascript">

	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "main_TopKeyword.asp?menupos=<%=menupos%>";

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->