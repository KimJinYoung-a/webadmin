<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim itemid, mode
dim sellyn, dispyn, packyn, usingyn

itemid = request.Form("itemid")
mode   = request.Form("mode")
sellyn = request.Form("sellyn")
dispyn = request.Form("dispyn")

dim obuyprice,oneitem

dim i
dim sqlStr
if mode="arr" then
	itemid = split(itemid,"|")
	sellyn = split(sellyn,"|")
	dispyn = split(dispyn,"|")

	for i=lBound(itemid)+1 to Ubound(itemid)
	    
	    if (itemid(i)<>"") and (sellyn(i)<>"") then
    		sqlStr = "update [db_item].[dbo].tbl_item" + VbCrlf
    		sqlStr = sqlStr + " set sellyn='" + CStr(sellyn(i)) + "'" + VbCrlf
    		sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
    		sqlStr = sqlStr + " where itemid=" + CStr(itemid(i))
            
            dbget.Execute sqlStr
        end if
	next
end if


dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
alert('수정되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->