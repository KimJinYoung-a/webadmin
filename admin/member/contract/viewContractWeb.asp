<%@ language=vbscript %>
<% option explicit %>
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
    response.write "<script>alert('������ ���ų� ��ȿ�� ����ȣ�� �ƴմϴ�.');</script>"
    dbget.close()	:	response.End
end if

'Response.Buffer=true
'Response.Expires=0
'Response.ContentType = "application/msword"
'Response.AddHeader "Content-Disposition", "attachment;filename=" & onecontract.FOneItem.FcontractName & "(" & onecontract.FOneItem.FContractNo & ")" & ".doc"

%>

<%= onecontract.FOneItem.FContractContents %>

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
