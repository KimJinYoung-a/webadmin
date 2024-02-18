<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="EUC-KR" %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->

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
Session.CodePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->