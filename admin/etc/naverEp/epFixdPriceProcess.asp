<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim sqlStr
Dim mode, itemid, fixedcash, useyn
mode	    = requestCheckVar(request("mode"), 32)
itemid	    = requestCheckVar(request("itemid"), 10)
fixedcash   = requestCheckVar(request("fixedcash"), 10)
useyn       = requestCheckVar(request("useyn"), 10)
'rw mode
'rw itemid
'rw fixedcash

if (mode="add") then
    sqlStr = "exec db_temp.[dbo].[usp_TEN_NV_FixedPriceAdd] "&itemid&","&fixedcash&",'"&session("ssBctId")&"'"
    dbget.Execute sqlStr
elseif (mode="edit") then
    sqlStr = "exec db_temp.[dbo].[usp_TEN_NV_FixedPriceEdit] "&itemid&","&fixedcash&",'"&session("ssBctId")&"'"
    dbget.Execute sqlStr
elseif (mode="useyn") then
    sqlStr = "exec db_temp.[dbo].[usp_TEN_NV_FixedPriceUseEdit] "&itemid&",'"&useyn&"','"&session("ssBctId")&"'"
    dbget.Execute sqlStr
else
    rw "ERROR:"&mode
    dbget.Close()
    response.end
end if
%>
<script>
    alert("수정되었습니다.");
    opener.location.reload();
    window.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->