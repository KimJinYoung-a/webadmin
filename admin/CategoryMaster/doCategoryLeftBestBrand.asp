<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doCategoryLeftBestBrand.asp
' Discription : 카테고리 베스트 브랜드 처리 페이지
' History : 2008.04.02 허진원 : 생성
'###############################################

dim mode,cdl,itemid, SortNo, selIdx, arrIdx, arrSort
dim refer, menupos, lp

menupos = request("menupos")
mode = request("mode")
cdl = request("cdl")
itemid = trim(request("itemid"))
selIdx = Request("selIdx")
SortNo = Request("SortNo")

if right(itemid,1)="," then
	itemid = left(itemid,len(itemid)-1)
end if

dim sqlStr

'// 모드별 분기 //
Select Case mode
	Case "del"
		'삭제처리
		sqlStr = "Update [db_sitemaster].[dbo].tbl_category_left_bestbrand"
		sqlStr = sqlStr + " Set isusing='N' "
		sqlStr = sqlStr + " where idx in (" + selIdx + ")"
	
		dbget.Execute(sqlStr)

	Case "changeSort"
		'표시순서 일괄 변경
		if selIdx<>"" then
			arrIdx = split(selIdx,",")
			arrSort = split(SortNo,",")

			for lp=0 to ubound(arrIdx)
				sqlStr = sqlStr & "Update [db_sitemaster].[dbo].tbl_category_left_bestbrand " &_
						" Set sortNo=" & arrSort(lp) &_
						" Where idx=" & arrIdx(lp) & ";" & vbCrLf
			next
			dbget.Execute(sqlStr)
		end if

end Select

refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "<%=refer%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->