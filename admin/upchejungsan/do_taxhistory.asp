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
	
	''���� ���̺� ��꼭INDEX - NULL
	
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
alert('���� �αװ� ��ҵǾ����ϴ�(<%= resultrow %>��). \n����� ���ݰ�꼭�� ���� ����Ͻñ� �ٶ��ϴ�.');
window.close();
</script>
<% end if %>

<!-- #include virtual="/lib/db/dbclose.asp" -->