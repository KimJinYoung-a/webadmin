<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 핑거스 계약 관리
' Hieditor : 2016.08.10 서동석 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/classes/partners/fingersUpcheAgreeCls.asp"-->
<%
dim sqlStr
dim agreeIdx

agreeIdx  = requestCheckvar(request("agreeIdx"),10)

dim onecontract
set onecontract = new CFingersUpcheAgree
onecontract.FRectDelInclude = "on"
onecontract.FRectagreeIdx = agreeIdx

if agreeIdx<>"" then
    onecontract.getOneFingersUpcheAgree
end if

if onecontract.FResultCount<1 then
    response.write "권한이 없거나, 유효한 계약번호가 아닙니다."
    dbget.close()	:	response.End
end if


dim itypeName : itypeName = onecontract.FoneItem.getContractTypeAgreeName
%>
<%= onecontract.FOneItem.getContractContents %>

<% if (onecontract.FOneItem.IsPrivContractAddItem) then %>
<br><br>
<div style='page-break-before:always'></div>
<%= getPriContractContents(onecontract.FOneItem.Fcompanyname) %>
<% end if %>
<%
set onecontract = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
