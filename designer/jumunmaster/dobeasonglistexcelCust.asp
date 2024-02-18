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
''통합 .2008-05-20
If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('세션이 종료되었습니다.');</script>"
    dbget.close()	:	response.End
end if

dim iMakerid : iMakerid = session("ssBctId")
dim iCustSheetType

IF (iMakerid="victoria001") or (iMakerid="toms001") THEN
    iCustSheetType = 1                      '''기존 v0에 상품명+옵션명 통합
ELSEIF (iMakerid="thegirin") THEN
    iCustSheetType = 2                      '''대한통운 송장 자료?
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
<title>송장파일</title>
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
	<td align="center" x:str >주문번호</td>
	<td align="center" x:str >주문일</td>
	<td align="center" x:str >구매자명</td>
	<td align="center" x:str >구매자전화</td>
	<td align="center" x:str >구매자핸드폰</td>
	<td align="center" x:str >구매자이메일</td>
	<td align="center" x:str >수령인</td>
	<td align="center" x:str >수령인전화</td>
	<td align="center" x:str >수령인핸드폰</td>
	<td align="center" x:str >우편번호</td>
	<td align="center" x:str >배송지주소</td>
	<td align="center" x:str >배송유의사항</td>
	<td align="center" x:str >택배번호</td>
	<td align="center" x:str >상품아이디</td>
	<td align="center" x:str >상품명 옵션명</td>
	<td align="center" x:str >판매가</td>
	<td align="center" x:str >수량</td>
	<td align="center" x:str >주문제작메세지</td>
	<td align="center" x:str >사은품</td>
	<td align="center" x:str >배송희망일(플라워)</td>
	<td align="center" x:str >카드리본(플라워)</td>
	<td align="center" x:str >메세지(플라워)</td>
	<td align="center" x:str >보내는사람(플라워)</td>
	<td align="center" x:str >상품명</td>
	<td align="center" x:str >옵션명</td>
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
	    [<%= i+ 1 %>번 상품 문구]<br>
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
	<td align="center" x:str><%= Left(CStr(ojumun.FMasterItemList(ix).Freqdate),10) %>일 <%= (ojumun.FMasterItemList(ix).GetReqTimeText) %></td>
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
	<td align="center" x:str >수화인</td>
	<td align="center" x:str >우편번호</td>
	<td align="center" x:str >주소</td>
	<td align="center" x:str >전화번호</td>
	<td align="center" x:str >핸드폰</td>
	<td align="center" x:str >상품명</td>
	<td align="center" x:str >비고</td>
	<td align="center" x:str >착불</td>
	<td align="center" x:str >금액</td>
</tr>
<% for ix=0 to ojumun.FResultCount - 1 %>
<tr>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqName %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqZipCode %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqZipAddr %> <%= ojumun.FMasterItemList(ix).FReqAddress %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqPhone %></td>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FReqHp %></td>
    <% IF (NULL2blank(ojumun.FMasterItemList(ix).FItemoptionName)<>"") THEN %>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FItemName %> <%= (ojumun.FMasterItemList(ix).FItemoptionName) %> (<%= ojumun.FMasterItemList(ix).FItemNo %> 개)</td>
	<% ELSE %>
	<td align="center" x:str><%= ojumun.FMasterItemList(ix).FItemName %> (<%= ojumun.FMasterItemList(ix).FItemNo %> 개)</td>
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
	        bufstr = bufstr & "["& i+ 1 &"번 상품 문구]<br>"
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
