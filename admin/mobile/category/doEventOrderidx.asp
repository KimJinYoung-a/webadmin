<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###########################################################
' Description :  ����� ī�װ� ���� �̺�Ʈ ���� ����
' History : 2020.12.02 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'// ���� ���� �� ���ް� ����
Dim idxarr, orderidxarr, cnt, sqlStr, ix

	idxarr = request.form("idxarr")
	orderidxarr	= request.form("orderidxarr")
    if idxarr <> "" then
        idxarr = split(idxarr,",")
        cnt = ubound(idxarr)
        orderidxarr = split(orderidxarr,",")
        for ix=0 to cnt	
            sqlStr = "UPDATE [db_sitemaster].[dbo].tbl_display_catemain_ex" & VbCrlf
            sqlStr = sqlStr & " SET view_order = " & Cstr(orderidxarr(ix)) & VbCrlf
            sqlStr = sqlStr &	"	WHERE idx=" & Cstr(idxarr(ix))
            dbget.execute sqlStr
        next
    end if
response.write "<script>parent.location.reload();</script>"
response.write "<script>alert('����Ǿ����ϴ�.');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->