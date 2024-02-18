<%@  codepage="949" language="VBScript" %>
<% option explicit %>
<%Response.Addheader "P3P","policyref='/w3c/p3p.xml', CP='NOI DSP LAW NID PSA ADM OUR IND NAV COM'"%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.CharSet="EUC-KR"
Response.ContentType="text/html;charset=euc-kr"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<%
dim sqlStr
dim ctrKey
ctrKey     = requestCheckvar(request("ctrKey"),10)


dim onecontract
set onecontract = new CPartnerContract
onecontract.FRectCtrKey = ctrKey

if ctrKey<>"" then
    onecontract.GetOneContractMaster
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

<%= replace(onecontract.FOneItem.FContractContents,"$$IMAGE1$$", DocuSignStampBase64) %>

<% if (CPrvContract) and (onecontract.FOneItem.IsDefaultContract) then %>
    <% if (onecontract.FOneItem.FsignType<>"D") Then %>
        <div style='page-break-before:always'></div>
        <%= getPriContractContents(onecontract.FOneItem.FB_UPCHENAME) %>
    <% end if %>
<% end if %>    
<%
set onecontract = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
    Session.CodePage = 949
%>