<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�� ���� �̹��� ���� ���� ó��
' History : 2019.11.07 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
dim idx, lp, sqlStr, refer
    refer = request.ServerVariables("HTTP_REFERER")

	Set idx 	= request.form("idx")

    for lp=1 to idx.count
        sqlStr = "UPDATE db_sitemaster.dbo.tbl_brand_image Set "
        sqlStr = sqlStr & " isusing=" & split(idx(lp),"/")(1) & ", "
        sqlStr = sqlStr & " lastupdate=getdate(), "
        sqlStr = sqlStr & " lastadminid='" & session("ssBctId") & "' "
        sqlStr = sqlStr & " Where idx=" & split(idx(lp),"/")(0)
        dbget.execute sqlStr
    next

    Call Alert_Move("����Ǿ����ϴ�.",refer)
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->