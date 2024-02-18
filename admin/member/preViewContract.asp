<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->

<%
dim sqlStr
dim ContractID
ContractID     = request("ContractID")


dim onecontract
set onecontract = new CPartnerContract
onecontract.FRectContractID = ContractID

if ContractID<>"" then
    onecontract.getOneContract
end if

if onecontract.FResultCount<1 then
    response.write "<script>alert('권한이 없거나 유효한 계약번호가 아닙니다.');</script>"
    dbget.close()	:	response.End    
end if

'Response.Buffer=true
'Response.Expires=0
'Response.ContentType = "application/msword" 
'Response.AddHeader "Content-Disposition", "attachment;filename=" & onecontract.FOneItem.FcontractName & "(" & onecontract.FOneItem.FContractNo & ")" & ".doc"

%>

<%= onecontract.FOneItem.FContractContents %>

<%
set onecontract = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
