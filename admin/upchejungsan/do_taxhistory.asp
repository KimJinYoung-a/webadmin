<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idx, mode
dim sqlstr, resultrow

idx = request("idx")
mode = request("mode")


if mode="delhistory" then
	sqlstr = "update  [db_jungsan].[dbo].tbl_tax_history_master" + VbCrlf
	sqlstr = sqlstr + " set deleteyn='Y'" + VbCrlf
	sqlstr = sqlstr + " where idx=" + CStr(idx) + " " + VbCrlf

	dbget.execute sqlstr, resultrow
	
	''정산 테이블 계산서INDEX - NULL
	
	sqlstr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master" + VbCrlf
    sqlstr = sqlstr + " set taxlinkidx=NULL" + VbCrlf
    sqlstr = sqlstr + " ,neotaxno=NULL" + VbCrlf
    sqlstr = sqlstr + " where taxlinkidx=" + CStr(idx) + " " + VbCrlf
    sqlstr = sqlstr + " and neotaxno in (select tax_no from [db_jungsan].[dbo].tbl_tax_history_master where idx=" + CStr(idx) + " )"

    dbget.execute sqlstr
end if
%>

<% if mode="delhistory" then %>
<script language='javascript'>
alert('발행 로그가 취소되었습니다(<%= resultrow %>건). \n발행된 세금계산서를 승인 취소하시기 바랍니다.');
window.close();
</script>
<% end if %>

<!-- #include virtual="/lib/db/dbclose.asp" -->