<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->

<%

dim ContractType
ContractType     = request("ContractType")


dim onecontractProtoType
set onecontractProtoType = new CPartnerContract
onecontractProtoType.FRectContractType = ContractType
onecontractProtoType.getOneContractProtoType


'Response.Buffer=true
Response.Expires=0
Response.ContentType = "application/msword"
Response.AddHeader "Content-Disposition", "attachment;filename=" & onecontractProtoType.FOneItem.FcontractName & "_¿øº»" & ".doc"

%>

<%= onecontractProtoType.FOneItem.FContractContents %>

<%
set onecontractProtoType = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->