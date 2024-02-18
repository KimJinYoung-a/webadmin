<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : Category_left_bestBrand.asp
' Discription : 카테고리 좌측 베스트 브랜드
' History : 2008.04.02 한용민 텐바이텐어드민 이전/수정
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<%
dim mode,cdl,itemid, SortNo, selIdx, arrIdx, arrSort ,refer, menupos, lp , sqlStr
	menupos = request("menupos")
	mode = request("mode")
	cdl = request("cdl")
	itemid = trim(request("itemid"))
	selIdx = Request("selIdx")
	SortNo = Request("SortNo")

if right(itemid,1)="," then
	itemid = left(itemid,len(itemid)-1)
end if

'// 모드별 분기 //
Select Case mode
	Case "del"
		'삭제처리
		sqlStr = "Update [db_academy].dbo.tbl_category_left_bestbrand"
		sqlStr = sqlStr + " Set isusing='N' "
		sqlStr = sqlStr + " where idx in (" + selIdx + ")"
	
		dbacademyget.Execute(sqlStr)

	Case "changeSort"
		'표시순서 일괄 변경
		if selIdx<>"" then
			arrIdx = split(selIdx,",")
			arrSort = split(SortNo,",")

			for lp=0 to ubound(arrIdx)
				sqlStr = sqlStr & "Update [db_academy].dbo.tbl_category_left_bestbrand " &_
						" Set sortNo=" & arrSort(lp) &_
						" Where idx=" & arrIdx(lp) & ";" & vbCrLf
			next
			dbacademyget.Execute(sqlStr)
		end if
end Select

refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">

	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "<%=refer%>";

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->