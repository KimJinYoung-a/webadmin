<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doCategoryLeftBigchance.asp
' Discription : 카테고리 빅찬스 처리 페이지
' History : 2008.03.31 허진원 : 생성
'           2008.07.25 허진원 수정 : 상품 정렬순서 추가
'###############################################

dim mode,cdl,cdm, itemid, selIdx, sortNo
dim i, refer, menupos
dim arrSortNo, arrIdx

menupos = request("menupos")
mode = request("mode")
cdl = request("cdl")
cdm = request("cdm")
itemid = trim(request("itemid"))
selIdx = Request("selIdx")
sortNo = Request("arrSort")

if right(itemid,1)="," then
	itemid = left(itemid,len(itemid)-1)
end if

dim sqlStr

'// 모드별 분기 //
Select Case mode
	Case "del"
		'선택상품 삭제
		sqlStr = "delete from [db_sitemaster].[dbo].tbl_category_left_bigchance"
		sqlStr = sqlStr + " where idx in (" + selIdx + ")"
	
		rsget.Open sqlStr,dbget,1

		refer = request.ServerVariables("HTTP_REFERER")

	Case "sort"
		'선택상품 정렬번호 적용
		arrIdx = split(selIdx,",")
		arrSortNo = split(sortNo,",")

		sqlStr = ""
		for i=0 to ubound(arrIdx)
			sqlStr = sqlStr & "Update [db_sitemaster].[dbo].tbl_category_left_bigchance"
			sqlStr = sqlStr & " Set sortNo=" & arrSortNo(i)
			sqlStr = sqlStr & " where idx=" & arrIdx(i) & "; " & vbCrLf
		next

		'response.Write sqlStr
		rsget.Open sqlStr,dbget,1

		refer = request.ServerVariables("HTTP_REFERER")

	Case "add"
		'신규 상품 추가
		if cdl<>"110" then
			sqlStr = "insert into [db_sitemaster].[dbo].tbl_category_left_bigchance"
			sqlStr = sqlStr + " (cdl, itemid)"
			sqlStr = sqlStr + " select  '" + cdl + "', itemid"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item"
			sqlStr = sqlStr + " where itemid in (" + itemid + ")"
			sqlStr = sqlStr + " and itemid not in ("
			sqlStr = sqlStr + " select itemid from [db_sitemaster].[dbo].tbl_category_left_bigchance"
			sqlStr = sqlStr + " where cdl='" + cdl + "'"
			sqlStr = sqlStr + " and itemid in (" + itemid + ")"
			sqlStr = sqlStr + ")"
		else
			'감성채널일 경우 중분류 추가
			sqlStr = "insert into [db_sitemaster].[dbo].tbl_category_left_bigchance"
			sqlStr = sqlStr + " (cdl, cdm, itemid)"
			sqlStr = sqlStr + " select  '" + cdl + "', '" + cdm + "', itemid"
			sqlStr = sqlStr + " from [db_item].[dbo].tbl_item"
			sqlStr = sqlStr + " where itemid in (" + itemid + ")"
			sqlStr = sqlStr + " and itemid not in ("
			sqlStr = sqlStr + " select itemid from [db_sitemaster].[dbo].tbl_category_left_bigchance"
			sqlStr = sqlStr + " where cdl='" + cdl + "'"
			sqlStr = sqlStr + " and cdm='" + cdm + "'"
			sqlStr = sqlStr + " and itemid in (" + itemid + ")"
			sqlStr = sqlStr + ")"
		end if

		rsget.Open sqlStr,dbget,1

		refer = "category_left_Bigchance.asp?menupos=" & menupos
end Select

%>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "<%=refer%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
