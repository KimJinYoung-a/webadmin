<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->

<%
dim id
id = requestCheckVar(request("id"),10)

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster

	if (ocsaslist.FOneItem.Fdeleteyn = "Y") then
	    response.write "<script>alert(" + Chr(34) + "이미 삭제된 내역입니다." + Chr(34) + ")</script>"
	    response.write "이미 삭제된 내역입니다."
	    dbget.close()	:	response.End
	elseif (ocsaslist.FOneItem.Fcurrstate = "B007") then
		response.write "<script>alert(" + Chr(34) + "이미 완료된 내역입니다." + Chr(34) + ")</script>"
		response.write "이미 완료된 내역입니다."
		dbget.close()	:	response.End
	end if
end if


dim orefund
set orefund = New CCSASList
orefund.FRectCsAsID = id

if (id<>"") then
    orefund.GetOneRefundInfo
end if

''주문 마스타
dim oordermaster
set oordermaster = new COrderMaster

if (ocsaslist.FResultCount>0) then
    IF (ocsaslist.FOneItem.Frefminusorderserial<>"") then
        oordermaster.FRectOrderSerial = ocsaslist.FOneItem.Frefminusorderserial
    ELSE
        oordermaster.FRectOrderSerial = ocsaslist.FOneItem.FOrderSerial
    ENd IF

    oordermaster.QuickSearchOrderMaster
end if

''주문 디테일
dim oorderdetail
set oorderdetail = new COrderMaster

if (oordermaster.FResultCount>0) then
    oorderdetail.FRectOrderSerial = ocsaslist.FOneItem.FOrderSerial
    oorderdetail.QuickSearchOrderDetail
end if

if (ocsaslist.FResultCount<1) or (orefund.FResultCount<1) then
    response.write "<script>alert('환불내역이 없거나 유효하지 않은 내역입니다.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

if (ocsaslist.FOneItem.FCurrstate<>"B001") then
    response.write "<script>alert('접수 상태가 아닙니다.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

'' 신용카드 취소만 가능
'if (orefund.FOneItem.Freturnmethod<>"R100") then
'    response.write "<script>alert('현재 신용카드 거래만 취소 가능합니다.');</script>"
'    response.write "<script>window.close();</script>"
'    dbget.close()	:	response.End
'end if

''기프팅은 수기취소 '' 제휴마케팅통해 수기취소후 완료 //2013/01/14
if (oordermaster.FOneItem.Faccountdiv="550") then
    response.write "<script>alert('기프팅 취소는 수기로만 가능합니다.\n\n제휴마케팅 담당자 통해 기프팅에 취소요청후 수기완료 하시기 바랍니다.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

'' IniPay 만 취소만 가능 INIMX_CARD INIMX_ISP_
dim IsInicisTID : IsInicisTID = False
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="IniTechPG_")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_CARD")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_ISP_")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIswtCARD")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIswtISP_")
IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,6)="Stdpay")
''IsInicisTID = IsInicisTID or (Left(orefund.FOneItem.FpaygateTid,10)="INIMX_AUTH")

if (oordermaster.FOneItem.Fpggubun <> "KA") and (oordermaster.FOneItem.Fpggubun <> "KK") and (oordermaster.FOneItem.Fpggubun <> "CH") and (oordermaster.FOneItem.Fpggubun <> "NP") and (oordermaster.FOneItem.Fpggubun <> "PY") and (oordermaster.FOneItem.Fpggubun <> "TS") and Not IsInicisTID AND orefund.FOneItem.Freturnmethod<>"R400" AND orefund.FOneItem.Freturnmethod<>"R420" then
    response.write "<script>alert('이니시스,카카오PAY, 네이버페이, 페이코, 토스 거래만 취소 가능합니다.["&oordermaster.FOneItem.Fpggubun&"]aaaa');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if
''rw oordermaster.FOneItem.FOrderSERIAL

if (oordermaster.FResultCount>0) then
    if (oordermaster.FOneItem.FCancelYn="N") and (oordermaster.FOneItem.Fjumundiv<>"9") and (orefund.FOneItem.Freturnmethod<>"R120") and (orefund.FOneItem.Freturnmethod<>"R022") and (orefund.FOneItem.Freturnmethod<>"R420")  then
        response.write "<script>alert('반품주문 또는 주문이 취소된 경우\n\n신용카드일부취소 또는 핸드폰부분취소만 취소 가능합니다.');</script>"
        response.write "<script>window.close();</script>"
        dbget.close()	:	response.End
    end if
end if

dim i
dim IsDirectCancelAvail
IsDirectCancelAvail = True

for i=0 to oorderdetail.FResultCount - 1
    if (oorderdetail.FItemList(i).FItemId<>0) then
        if (Not (IsNULL(oorderdetail.FItemList(i).Fcurrstate) or (oorderdetail.FItemList(i).Fcurrstate<3))) then
            IsDirectCancelAvail = False
        end if
    end if
next

dim CancelCase , etcCancelCase

if (Left(ocsaslist.FOneItem.FOrderSerial,1)="A") or (Left(ocsaslist.FOneItem.FOrderSerial,1)="B") then
    CancelCase="강좌최소"
else
    CancelCase="배송전취소"

    if (oordermaster.FOneItem.Fjumundiv="9") then
        etcCancelCase = "반품"
    elseif (orefund.FOneItem.Freturnmethod="R120") or (orefund.FOneItem.Freturnmethod="R022") or (orefund.FOneItem.Freturnmethod="R420") then
        etcCancelCase = "배송전부분취소"
    end if
end if
%>
<script language='javascript'>
function ActCancel(frm){
    if (frm.msg.value.length<1){
        alert('취소사유를 입력해 주세요.');
        frm.msg.focus();
        return;
    }

    if ((frm.returnmethod.value=="R120") || (frm.returnmethod.value=="R022") || (frm.returnmethod.value=="R420")) {
        //부분취소(신용카드, 핸드폰)
        if (confirm('부분 취소 진행 하시겠습니까?')){
            frm.action="pop_PartialCardCancel_process_test.asp";
            frm.submit();
        }
    }else{
        if (confirm('승인 취소 하시겠습니까?')){
            frm.action="pop_CardCancel_process_20150812.asp";
            frm.submit();
        }
    }
}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmCanncel" method="post" action="pop_CardCancel_process_20150812.asp">
<input type="hidden" name="id" value="<%= id %>">
<input type="hidden" name="returnmethod" value="<%= orefund.FOneItem.Freturnmethod %>">
<% if (oordermaster.FResultCount>0) then %>
<input type="hidden" name="rdsite" value="<%= oordermaster.FOneItem.Frdsite%>">
<input type="hidden" name="buyemail" value="<%= oordermaster.FOneItem.Fbuyemail%>">
<% end if %>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">취소자</td>
    <td bgcolor="#FFFFFF">
        <%= session("ssBctID") %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">주문번호</td>
    <td bgcolor="#FFFFFF">
        <%= ocsaslist.FOneItem.FOrderSerial %>

        <% if (ocsaslist.FOneItem.Frefminusorderserial<>"") then %>
        (마이너스 주문번호 : <%= ocsaslist.FOneItem.Frefminusorderserial %>)
        <% end if %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">상태</td>
    <td bgcolor="#FFFFFF">
    <% if (oordermaster.FResultCount>0) then %>
        <font color="<%= oordermaster.FOneItem.CancelYnColor %>"><%= oordermaster.FOneItem.CancelYnName %></font> <font color="<%= oordermaster.FOneItem.IpkumDivColor %>"><%= oordermaster.FOneItem.IpkumDivName %>

        <% if (oordermaster.FOneItem.Fjumundiv="9") then %>
        <font color=red><strong>[마이너스 주문]</strong></font>
        <% end if %>
    <% end if %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">구매자ID</td>
    <td bgcolor="#FFFFFF">
        <%= ocsaslist.FOneItem.FUserID %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">취소방식</td>
    <td bgcolor="#FFFFFF">
        <%= orefund.FOneItem.FreturnmethodName %>
        <% if (orefund.FOneItem.Freturnmethod="R120") or (orefund.FOneItem.Freturnmethod="R022") then %>
        (<strong><%= orefund.FOneItem.Freturnmethod %></strong>)
        <% else %>
		(<%= orefund.FOneItem.Freturnmethod %>)
		<% end if %>
    </td>
</tr>
<tr height="25">
    <td width="90" bgcolor="<%= adminColor("topbar") %>">취소금액</td>
    <td bgcolor="#FFFFFF">
        <%= FormatNumber(orefund.FOneItem.Frefundrequire,0) %>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">PG사 ID</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text_ro" name="tid" value="<%= orefund.FOneItem.FpaygateTid %>" size="60" readonly>
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">취소사유</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text" name="msg" value="<%= ChkIIF(IsDirectCancelAvail,CancelCase, etcCancelCase) %>" size="50" maxlength="60" >
    	<% if (oordermaster.FResultCount>0) then %>
    	<% if ((C_ADMIN_AUTH) and (oordermaster.FOneItem.Fjumundiv="9")) or (session("ssBctID")="icommang") or (session("ssBctID")="iroo4")  then %>
    	<input type="checkbox" name="force" >금액검토안함
    	<% end if %>
    	<% end if %>
    </td>
</tr>

<tr height="25">
    <td colspan="2" align="center" bgcolor="#FFFFFF">
    <% if (orefund.FOneItem.Freturnmethod="R120") or (orefund.FOneItem.Freturnmethod="R022") then %>
    <input type="button" class="button" value=" 승인 부분 취소 " onClick="ActCancel(frmCanncel)">
    <% else %>
    <input type="button" class="button" value=" 승인 취소 " onClick="ActCancel(frmCanncel)">
    <% end if %>
    </td>
</tr>
</form>
</table>
<%
set ocsaslist = Nothing
set orefund = Nothing
set oordermaster = Nothing
set oorderdetail = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
