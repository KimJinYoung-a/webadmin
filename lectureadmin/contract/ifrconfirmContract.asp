<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/fingersUpcheAgreeCls.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim agreeIdx : agreeIdx=requestCheckvar(request("agreeIdx"),10)
dim gkey : gkey=request("gkey")
dim ekey : ekey=request("ekey")
dim chkcf : chkcf=requestCheckvar(request("chkcf"),10)

if (ekey="") then
    response.write "암호화 키가 올바르지 않습니다."
    dbget.close()	:	response.end
end if

if (UCASE(ekey)<>UCASE(MD5("TBTCTR"&agreeIdx&gkey))) then
    response.write "암호화 키가 올바르지 않습니다."
    dbget.close()	:	response.end
end if


dim makerid : makerid = session("ssBctID")
dim groupid : groupid = getPartnerId2GroupID(makerid)


dim onecontract
set onecontract = new CFingersUpcheAgree
onecontract.FRectagreeIdx = agreeIdx

if agreeIdx<>"" then
    onecontract.getOneFingersUpcheAgree
end if

if onecontract.FResultCount<1 then
    response.write "권한이 없거나, 유효한 계약번호가 아닙니다."
    dbget.close()	:	response.End
end if

%>

<%= onecontract.FOneItem.getContractContents %>

<% if (chkcf="1")and(onecontract.FOneItem.IsPrivContractAddItem) then %>
<div style='page-break-before:always'></div>
<%= getPriContractContents(onecontract.FOneItem.Fcompanyname) %>
<% end if %>

<%
set onecontract = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
