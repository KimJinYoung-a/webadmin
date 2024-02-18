<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
dim itemid : itemid = requestCheckvar(request("itemid"),10)
dim actType: actType = requestCheckvar(request("actType"),10)

dim sqlStr, AssignedRow
sqlStr = "exec [db_item].[dbo].[usp_Ten_QuickDlvITem_Set] "&itemid&",'"&actType&"','"&session("ssBctId")&"'"
dbget.Execute sqlStr

%>

<script>
    alert('수정되었습니다.')
    opener.location.reload();
    window.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->