<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesignerNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/checkPartnerLog.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_baljucls.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->

<%
''���� .2008-05-20
If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('������ ����Ǿ����ϴ�.');</script>"
    dbget.close()	:	response.End
end if

dim iMakerid : iMakerid = session("ssBctId")
dim iCustSheetType

IF (iMakerid="victoria001") or (iMakerid="toms001") THEN
    iCustSheetType = 1                      '''���� v0�� ��ǰ��+�ɼǸ� ����
ELSEIF (iMakerid="thegirin") THEN
    iCustSheetType = 2                      '''������� ���� �ڷ�?
ELSE
    iCustSheetType = 0
END IF

dim ojumun
dim ix,sql
Dim listitemlist,listitem,listitemcount
dim iSall, SheetType

listitem  =  Replace(request("orderserial"), " ", "")
iSall     =  requestCheckVar(request("isall"), 32)
SheetType =  request("SheetType"), 32)

set ojumun = new CJumunMaster

ojumun.FRectOrderSerial = listitem
ojumun.FRectIsAll       = iSall
ojumun.FRectDesignerID = session("ssBctID")
ojumun.ReDesignerSelectBaljuList

dim oGift, i, j
set oGift = new COrderGift

''Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"

function replaceXlText(org)
    dim reText
    reText = replace(org,"<","&lt;")
    replaceXlText = replace(reText,">","&gt;")
end function

function replaceVictoria001Type(imanageCode,ioptionName)
    imanageCode = NULL2Blank(imanageCode)
    ioptionName = NULL2Blank(ioptionName)

    IF (InStr(ioptionName,"[W]")>1) or (InStr(ioptionName,"[K]")>1) or (InStr(ioptionName,"[M]")>1) then
        ioptionName = Left(ioptionName,2)
    ENd IF

    IF (ioptionName<>"") then
        replaceVictoria001Type = imanageCode & " [" & ioptionName &"]"
    ELSE
        replaceVictoria001Type = imanageCode
    END IF


end function

Dim bufStr
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
	  xmlns:x="urn:schemas-microsoft-com:office:excel"
	  xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<title>��������</title>
<style>
    <!--
	br
	    {mso-data-placement:same-cell;}
	tr
	    {mso-height-source:auto;
	    mso-ruby-visibility:none;}
	td
	    {white-space:normal;}
	-->
</style>
</head>

<body leftmargin="10">
<% IF (iCustSheetType=1) then %>
<table width=1200 cellspacing=0 cellpadding=1 border=0>
<tr>
	<td align="center" x:str >�ֹ���ȣ</td>
	<td align="center" x:str >�ֹ���</td>
	<td align="center" x:str >�����ڸ�</td>
	<td align="center" x:str >��������ȭ</td>
	<td align="center" x:str >�������ڵ���</td>
	<td align="center" x:str >�������̸���</td>
	<td align="center" x:str >������</td>
	<td align="center" x:str >��������ȭ</td>
	<td align="center" x:str >�������ڵ���</td>
	<td align="center" x:str >�����ȣ</td>
	<td align="center" x:str >������ּ�</td>
	<td align="center" x:str >������ǻ���</td>
	<td align="center" x:str >�ù��ȣ</td>
	<td align="center" x:str >��ǰ���̵�</td>
	<td align="center" x:str >��ǰ�� �ɼǸ�</td>
	<td align="center" x:str >�ǸŰ�</td>
	<td align="center" x:str >����</td>
	<td align="center" x:str >�ֹ����۸޼���</td>
	<td align="center" x:str >����ǰ</td>
	<td align="center" x:str >��������(�ö��)</td>
	<td align="center" x:str >ī�帮��(�ö��)</td>
	<td align="center" x:str >�޼���(�ö��)</td>
	<td align="center" x:str >�����»��(�ö��)</td>
	<td align="center" x:str >��ǰ��</td>
	<td align="center" x:str >�ɼǸ�</td>
</tr>
<% for ix=0 to ojumun.FResultCount - 1 %>
<tr>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FOrderSerial %></td>
	<td align="center" x:str><%= Left(CStr(ojumun.FMasterItemList(ix).FRegDate),10) %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FBuyName %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FBuyPhone %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FBuyHp %></td>
	<td align="center" x:str></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqName %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqPhone %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqHp %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqZipCode %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqZipAddr %><%= " " %><%= ojumun.FMasterItemList(ix).FReqAddress %></td>
	<td align="center" x:str><%= db2html(ojumun.FMasterItemList(ix).FComment) %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).Fsongjangno %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).Fitemid %></td>
	<td align="center" x:str><%= replaceVictoria001Type(ojumun.FMasterItemList(ix).FupcheManageCode,ojumun.FMasterItemList(ix).FItemoptionName) %></td>
	<td align="center" x:num="<%= ojumun.FMasterItemList(ix).FItemCost %>" ><%= ojumun.FMasterItemList(ix).FItemCost %></td>
	<td align="center" x:num="<%= ojumun.FMasterItemList(ix).FItemNo %>" ><%= ojumun.FMasterItemList(ix).FItemNo %></td>
	<% if (ojumun.FMasterItemList(ix).FItemNo>1) then %>
	<td align="center" x:str >
	    <% if (Not IsNULL(ojumun.FMasterItemList(ix).Frequiredetail)) then %>
	    <% if ojumun.FMasterItemList(ix).Frequiredetail<>"" then %>
	    <% for i=0 to ojumun.FMasterItemList(ix).FItemNo-1 %>
	    [<%= i+ 1 %>�� ��ǰ ����]<br>
	    <%= nl2br(replaceXlText(splitValue(ojumun.FMasterItemList(ix).Frequiredetail,CAddDetailSpliter,i))) %>
	    <br>
	    <% next %>
	    <% end if %>
	    <% end if %>
	</td>
	<% else %>
	<td align="center" x:str ><%= nl2br(replaceXlText(Replace(ojumun.FMasterItemList(ix).Frequiredetail, CAddDetailSpliter, ""))) %></td>
	<% end if %>
	<%
	oGift.FRectOrderSerial = ojumun.FMasterItemList(ix).FOrderSerial
    oGift.FRectMakerid = session("ssBctId")
    oGift.FRectGiftDelivery = "Y"
    oGift.GetOneOrderGiftlist
	%>
	<% if (oGift.FResultCount>0) then %>
	<td align="center" x:str>
    <% for j=0 to oGift.FResultCount -1 %>
        <%= oGift.FItemList(j).GetEventConditionStr %><% if j>oGift.FResultCount -1 then %><br><% end if %>
    <% next %>
	</td>
	<% ELSE %>
	<td align="center" x:str></td>
	<% end if %>
<% if Not IsNULL(ojumun.FMasterItemList(ix).Freqdate) then %>
	<td align="center" x:str><%= Left(CStr(ojumun.FMasterItemList(ix).Freqdate),10) %>�� <%= (ojumun.FMasterItemList(ix).GetReqTimeText) %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).getCardribbonName %></td>
	<td align="center" x:str><%= replaceXlText(db2html(ojumun.FMasterItemList(ix).Fmessage)) %></td>
	<td align="center" x:str><%= db2html(ojumun.FMasterItemList(ix).Ffromname) %></td>
<% else %>
    <td align="center" x:str></td>
    <td align="center" x:str></td>
    <td align="center" x:str></td>
    <td align="center" x:str></td>
<% end if %>
    <td align="center" x:str><%= ojumun.FMasterItemList(ix).FItemName %></td>
    <td align="center" x:str><%= (ojumun.FMasterItemList(ix).FItemoptionName) %></td>
</tr>
<% next %>
</table>
<% ELSE %>

<table width=1200 cellspacing=0 cellpadding=1 border=0>
<tr>
	<td align="center" x:str >��ȭ��</td>
	<td align="center" x:str >�����ȣ</td>
	<td align="center" x:str >�ּ�</td>
	<td align="center" x:str >��ȭ��ȣ</td>
	<td align="center" x:str >�ڵ���</td>
	<td align="center" x:str >��ǰ��</td>
	<td align="center" x:str >���</td>
	<td align="center" x:str >����</td>
	<td align="center" x:str >�ݾ�</td>
</tr>
<% for ix=0 to ojumun.FResultCount - 1 %>
<tr>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqName %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqZipCode %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqZipAddr %> <%= ojumun.FMasterItemList(ix).FReqAddress %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqPhone %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqHp %></td>
    <% IF (NULL2blank(ojumun.FMasterItemList(ix).FItemoptionName)<>"") THEN %>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FItemName %> <%= (ojumun.FMasterItemList(ix).FItemoptionName) %> (<%= ojumun.FMasterItemList(ix).FItemNo %> ��)</td>
	<% ELSE %>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FItemName %> (<%= ojumun.FMasterItemList(ix).FItemNo %> ��)</td>
	<% end if %>

	<%
	bufstr = ""
	oGift.FRectOrderSerial = ojumun.FMasterItemList(ix).FOrderSerial
    oGift.FRectMakerid = session("ssBctId")
    oGift.FRectGiftDelivery = "Y"
    oGift.GetOneOrderGiftlist

	bufstr = bufstr & db2html(ojumun.FMasterItemList(ix).FComment)

	'' Gift
	if (oGift.FResultCount>0) then
         for j=0 to oGift.FResultCount -1
            bufstr = bufstr & oGift.FItemList(j).GetEventConditionStr
            if j>oGift.FResultCount -1 then bufstr = bufstr &"<br>" end if
         next
	end if

	if (ojumun.FMasterItemList(ix).FItemNo>1) then
	    if (Not IsNULL(ojumun.FMasterItemList(ix).Frequiredetail)) then
	    if ojumun.FMasterItemList(ix).Frequiredetail<>"" then
	    for i=0 to ojumun.FMasterItemList(ix).FItemNo-1
	        bufstr = bufstr & "["& i+ 1 &"�� ��ǰ ����]<br>"
	        bufstr = bufstr & nl2br(replaceXlText(splitValue(ojumun.FMasterItemList(ix).Frequiredetail,CAddDetailSpliter,i)))
	        bufstr = bufstr & "<br>"
	    next
	    end if
	    end if
	else
	    bufstr = bufstr & db2html(ojumun.FMasterItemList(ix).FComment)
	    bufstr = bufstr & nl2br(replaceXlText(Replace(ojumun.FMasterItemList(ix).Frequiredetail, CAddDetailSpliter, "")))
	end if
	%>
	<td align="center" x:str ><%= bufstr %></td>
    <td align="center" x:str></td>
    <td align="center" x:str></td>
</tr>
<% next %>
</table>
<% END IF %>
</body>
</html>
<%
set ojumun = Nothing
set oGift = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
