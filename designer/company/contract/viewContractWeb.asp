<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<%
dim sqlStr
dim ctrKey,gkey,ekey,chkcf
ctrKey  = requestCheckvar(request("ctrKey"),10)
gkey    = request("gkey")
ekey    = request("ekey")
chkcf   = request("chkcf")

if (ekey="") then
    response.write "��ȣȭ Ű�� �ùٸ��� �ʽ��ϴ�."
    response.end
end if

if (UCASE(ekey)<>UCASE(MD5("TBTCTR"&ctrKey&gkey))) then
    response.write "��ȣȭ Ű�� �ùٸ��� �ʽ��ϴ�."
    response.end
end if


dim onecontract
set onecontract = new CPartnerContract
onecontract.FRectCtrKey = ctrKey

if ctrKey<>"" then
    onecontract.GetOneContractMaster
end if

if onecontract.FResultCount<1 then
    response.write "������ ���ų�, ��ȿ�� ����ȣ�� �ƴմϴ�."
    dbget.close()	:	response.End
end if

'Response.Buffer=true
'Response.Expires=0
'Response.ContentType = "application/msword"
'Response.AddHeader "Content-Disposition", "attachment;filename=" & onecontract.FOneItem.FcontractName & "(" & onecontract.FOneItem.FContractNo & ")" & ".doc"

%>

<%= onecontract.FOneItem.FContractContents %>

<% if (CPrvContract) and (onecontract.FOneItem.IsDefaultContract) then %>
    <% if (TRUE) then %><div style='page-break-before:always'>&nbsp;</div><% end if %>
<%= getPriContractContents(onecontract.FOneItem.FB_UPCHENAME) %>
<% end if %>
<%
set onecontract = Nothing

''��ü�� �ٿ�ε� �� ��� Ȯ���� �÷��� ������Ʈ
if (chkcf="1") then
    sqlStr=" update db_partner.dbo.tbl_partner_ctr_master"
    sqlStr=sqlStr&" set confirmdate=IsNULL(confirmdate,getdate())"
    sqlStr=sqlStr&" ,ctrState=(CASE WHEN ctrState in (1,2) then 3 else ctrState end )"
    sqlStr=sqlStr&" where ctrKey="&ctrKey

    dbget.Execute sqlStr
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
