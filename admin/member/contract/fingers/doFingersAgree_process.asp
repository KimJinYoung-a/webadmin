<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ΰŽ� ��� ����
' Hieditor : 2016.08.10 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/fingersUpcheAgreeCls.asp"-->
<%
dim sqlStr , AssignedRow
dim mode : mode = request("mode")
dim agreeIdx : agreeIdx  = requestCheckvar(request("agreeIdx"),10) 

if (mode="del") then
    sqlStr = "update db_partner.dbo.tbl_partner_fingers_agreeHist"&vbCRLF
    sqlStr = sqlStr&" set deldate=isNULL(deldate,getdate())"&vbCRLF
    sqlStr = sqlStr&" where agreeIdx="&agreeIdx&vbCRLF
    sqlStr = sqlStr&" and agreedate is NULL"
    
    dbget.Execute sqlStr,AssignedRow
    
    if (AssignedRow>0) then
        response.write "<script>alert('���� ó�� �Ǿ����ϴ�.');opener.location.reload();window.close()</script>"
    else
        response.write "<script>alert('�۾��� ������ �߻��Ͽ����ϴ�.');history.back()</script>"
    end if
elseif (mode="") then
    
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->