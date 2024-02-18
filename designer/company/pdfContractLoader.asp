<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->

<%
dim sqlStr
dim ContractID, key
ContractID     = requestCheckVar(request("ContractID"),100)
key = requestCheckVar(request("key"),100)

if (key<>"123123") then response.end

dim onecontract
set onecontract = new CPartnerContract
onecontract.FRectContractID = ContractID
onecontract.FRectMakerid = session("ssBctID")

if ContractID<>"" then
    onecontract.getOneContract
end if

if onecontract.FResultCount<1 then
    response.write "<script>alert('권한이 없거나 유효한 계약번호가 아닙니다.');</script>"
    dbget.close()	:	response.End
else
    ''다운로드시 상태변경
    sqlStr = "update db_partner.dbo.tbl_partner_contract"
    sqlStr = sqlStr & " set contractState=3"
    sqlStr = sqlStr & " ,confirmDate=IsNULL(confirmDate,getdate())"
    sqlStr = sqlStr & " where ContractID=" & ContractID
    sqlStr = sqlStr & " and contractState=1"

    ''dbget.Execute sqlStr
end if


%>

<%= onecontract.FOneItem.FContractContents %>

<%
set onecontract = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->