<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->

<%
dim ContractType
ContractType     = request("ContractType")


dim onecontractProtoType 
set onecontractProtoType = new CPartnerContract
onecontractProtoType.FRectContractType = ContractType
onecontractProtoType.getOneContractProtoType

%>

<%= onecontractProtoType.FOneItem.FContractContents %>

<%
set onecontractProtoType = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->