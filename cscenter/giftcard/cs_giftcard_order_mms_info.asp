<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_giftcard_ordercls.asp" -->
<%

dim giftorderserial
giftorderserial = RequestCheckVar(request("giftorderserial"),11)

'==============================================================================
dim oGiftOrder

set oGiftOrder = new cGiftCardOrder

if (giftorderserial <> "") then
	oGiftOrder.FRectGiftOrderSerial = giftorderserial

	oGiftOrder.getCSGiftcardOrderDetail
end if

dim ix, i
dim tmpvalue, tmpselected

%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script>
function SubmitForm() {
	if (validate(frm)==false) {
		return ;
	}

    if (confirm("저장하시겠습니까?") == true) {
        frm.submit();
    }
}

document.title = "MMS 정보";
</script>


<!-- 구매자정보 -->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" onsubmit="return false;" action="cs_giftcard_order_info_edit_process.asp">
    <input type="hidden" name="mode" value="modifymmsinfo">
    <input type="hidden" name="giftorderserial" value="<%= oGiftOrder.FOneItem.FgiftOrderSerial %>">
    <input type="hidden" name="bookingYn" value="<%= oGiftOrder.FOneItem.FbookingYn %>">
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
	    <td colspan="2">
	        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	    		<tr>
	    			<td width="100">
	    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>MMS 정보</b>
				    </td>
				    <td align="right">
				        <input type="button" value="저장하기" class="csbutton" onClick="SubmitForm();">
				    </td>
				</tr>
			</table>
	    </td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">예약전송</td>
	    <td bgcolor="#FFFFFF">
	    	<% if (oGiftOrder.FOneItem.FbookingYn = "Y") then %>
	    		예약전송
	    	<% else %>
	    		즉시전송
	    	<% end if %>
	    </td>
	</tr>
	<% if (oGiftOrder.FOneItem.FbookingYn = "Y") then %>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">예약일시</td>
	    <td bgcolor="#FFFFFF">
	    	<input type="text" class="text" name="bookingDate" value="<%= Left(oGiftOrder.FOneItem.FbookingDate,10) %>" id="[off,off,off,off][예약일시]" size=10 readonly ><a href="javascript:calendarOpen(frm.bookingDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
	    	<select class=select name=bookingDateHH>
	    	<% for i = 8 to 20 %>
	    		<%
	    		tmpvalue = Right(("0" & i), 2)
	    		tmpselected = ""
	    		if (oGiftOrder.FOneItem.FbookingDate <> "") then
	    			if (Hour(oGiftOrder.FOneItem.FbookingDate) = i) then
	    				tmpselected = "selected"
	    			end if
	    			'
	    		end if
	    		%>
	    		<option value="<%= tmpvalue %>" <%= tmpselected %>><%= tmpvalue %></option>
	    	<% next %>
	    	</select>
	    	시
	    </td>
	</tr>
	<% end if %>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">보내는분HP</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text" name="sendhp" id="[on,off,1,16][보내는분HP]" value="<%= oGiftOrder.FOneItem.Fsendhp %>"></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">받는분HP</td>
	    <td bgcolor="#FFFFFF"><input type="text" class="text" name="reqhp" id="[on,off,1,16][받는분HP]" value="<%= oGiftOrder.FOneItem.Freqhp %>"></td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">MMS 제목</td>
	    <td bgcolor="#FFFFFF">
	    	<input type="text" class="text" name="MMSTitle" id="[on,off,1,32][MMS제목]" size="50" value="<%= oGiftOrder.FOneItem.FMMSTitle %>">
	    </td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">MMS 내용</td>
	    <td bgcolor="#FFFFFF">
	    	<textarea name="MMSContent" cols=45 rows=8 id="[on,off,1,32][MMS내용]"><%= oGiftOrder.FOneItem.FMMSContent %></textarea>
	    </td>
	</tr>
	</form>
</table>
<!-- 구매자정보 -->



<%
set oGiftOrder = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->