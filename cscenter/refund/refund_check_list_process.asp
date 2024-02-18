<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim mode
dim asidList
dim AssignedRow


mode		= trim(request("mode"))
asidList	= trim(request("asidList"))


dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr


Select Case mode
	Case "ipjumRefund"
		'' 제휴몰 구매확정 후 환불
		sqlStr = " update [db_cs].[dbo].[tbl_new_as_list] "
		sqlStr = sqlStr + " set title = '제휴몰 구매확정 후 환불' "
		sqlStr = sqlStr + "where id in (" & asidList & ") "
		''response.write sqlStr
		dbget.Execute sqlStr, AssignedRow
	Case "ipjumDiffRefund"
		'' 고객입금 차액환불
		sqlStr = " update [db_cs].[dbo].[tbl_new_as_list] "
		sqlStr = sqlStr + " set title = '고객입금 차액환불' "
		sqlStr = sqlStr + "where id in (" & asidList & ") "
		''response.write sqlStr
		dbget.Execute sqlStr, AssignedRow
	Case "prdDiffRefund"
		'' 상품대금 차액환불
		sqlStr = " update [db_cs].[dbo].[tbl_new_as_list] "
		sqlStr = sqlStr + " set title = '상품대금 차액환불' "
		sqlStr = sqlStr + "where id in (" & asidList & ") "
		''response.write sqlStr
		dbget.Execute sqlStr, AssignedRow
	Case "csDelivRefund"
		'' CS서비스 - 배송비환불
		sqlStr = " update [db_cs].[dbo].[tbl_new_as_list] "
		sqlStr = sqlStr + " set title = 'CS서비스 - 무통장 환불(배송비)' "
		sqlStr = sqlStr + "where id in (" & asidList & ") "
		''response.write sqlStr
		dbget.Execute sqlStr, AssignedRow
	Case "upcheJungsanRefund"
		'' 업체정산 및 고객환불
		sqlStr = " update [db_cs].[dbo].[tbl_new_as_list] "
		sqlStr = sqlStr + " set title = '업체정산 및 고객환불' "
		sqlStr = sqlStr + "where id in (" & asidList & ") "
		''response.write sqlStr
		dbget.Execute sqlStr, AssignedRow
	Case Else
		''
End Select

%>
<script language='javascript'>
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
