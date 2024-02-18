<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
Dim sqlStr, taxdate, iMsg, menupos, idx
menupos		= Request("menupos")
taxdate	= requestCheckvar(Request("taxdate"),10)
idx 	= requestCheckvar(Request("idx"),10)

If iMsg = "" Then
    sqlStr = ""
    sqlStr = sqlStr & " UPDATE [db_sitemaster].[dbo].[tbl_taxdate_manage] SET "
    sqlStr = sqlStr & " taxdate = '"&taxdate&"' "
    sqlStr = sqlStr & " ,lastUpdate = getdate() "
    sqlStr = sqlStr & " ,regUserid = '"&session("ssBctID")&"' "
    sqlStr = sqlStr & " WHERE idx = '"&idx&"'"
	rsget.Open sqlStr,dbget,1
	iMsg = "저장하였습니다."
End If 
%>
<script language="javascript">
<% If (iMsg <> "") Then %>
alert("<%=iMsg %>");
document.parent.reload();
<% End If %>

</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->