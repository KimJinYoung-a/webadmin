<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim sqlStr
dim idx, CreateIDX
idx	= Request("idx")

    'Just1Day 복사
    sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_just1day2018_list" & VbCrlf
    sqlStr = sqlStr & " (title, startdate, enddate, adminid, isusing, maxsaleper" & VbCrlf
    sqlStr = sqlStr & ", type, bannerimage, linkurl, workertext, platform)" & VbCrlf
    sqlStr = sqlStr & " select top 1 title, startdate, enddate, adminid, isusing, maxsaleper, type" & VbCrlf
    sqlStr = sqlStr & " , bannerimage, linkurl, workertext, 'mobile'" & VbCrlf
    sqlStr = sqlStr & " from [db_sitemaster].[dbo].tbl_just1day2018_list where idx=" & idx & VbCrlf
    dbget.Execute(sqlStr)

    sqlStr = "select SCOPE_IDENTITY()"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    IF not rsget.EOF THEN
        CreateIDX = rsget(0)
    End IF
    rsget.Close

    '상품 복사
    sqlStr = "Insert Into [db_sitemaster].[dbo].[tbl_just1day2018_item]" & VbCrlf
    sqlStr = sqlStr & " (listidx, title, itemid, frontimage, price, saleper, adminid, isusing, sortnum)" & VbCrlf
    sqlStr = sqlStr & " select "& CreateIDX &", title, itemid, frontimage, price, saleper, adminid, isusing, sortnum" & VbCrlf
    sqlStr = sqlStr & " from [db_sitemaster].[dbo].[tbl_just1day2018_item] where listidx=" & idx & VbCrlf
    sqlStr = sqlStr & " order by subidx asc"& VbCrlf
    dbget.Execute(sqlStr)

%>
<script>
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	window.opener.document.location.href = window.opener.document.URL;    // 부모창 새로고침
	 self.close();        // 팝업창 닫기
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->