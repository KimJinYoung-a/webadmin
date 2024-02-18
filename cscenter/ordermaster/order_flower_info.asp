<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->

<%

'// 문자열을 잘라 원하는 위치의 값을 반환 //
function SplitValue(orgStr,delim,pos)
    dim buf
    SplitValue = ""
    if IsNULL(orgStr) then Exit function
    if (Len(delim)<1) then Exit function
    buf = split(orgStr,delim)
    
    if UBound(buf)<pos then Exit function
    
    SplitValue = buf(pos)
end function

Sub DrawFlowerOneDateBox(byval yyyy,mm,dd,tt)
	dim buf,i

	buf = "<select name='yyyy'>"
    for i=2007 to cint(year(dateadd("yyyy",1,now)))
		if (CStr(i)=CStr(yyyy)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + ">" + CStr(i) + "</option>"
		end if
	next
    buf = buf + "</select>년 "

    buf = buf + "<select name='mm' >"
    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"'>" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>월 "

    buf = buf + "<select name='dd' >"
    for i=1 to 31
		if (Format00(2,i)=Format00(2,dd)) then
	    buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
        buf = buf + "<option value='" + Format00(2,i) + "'>" + Format00(2,i) + "</option>"
		end if
    next
    buf = buf + "</select>일 "


    buf = buf & "<select name='tt' >"
    for i=9 to 18
		if (Format00(2,i)=Format00(2,tt)) then
        buf = buf & "<option value='" & CStr(i) & "' selected>" & CStr(i) & "~" & CStr(i + 2) & "</option>"
		else
        buf = buf & "<option value='" & CStr(i) & "'>" & CStr(i) & "~" & CStr(i + 2) & "</option>"
		end if
    next
    buf = buf & "</select>시 "

    response.write buf
end Sub

dim ojumun, orderserial, AlertMsg, IsOldOrder
orderserial = requestCheckVar(request("orderserial"),11)

set ojumun = new COrderMaster

ojumun.FRectOrderSerial = orderserial
ojumun.QuickSearchOrderMaster

if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
    
    if (ojumun.FResultCount>0) then
        IsOldOrder = true
        AlertMsg = "6개월 이전 주문입니다."
    end if
    
end if

dim ix

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


document.title = "플라워 배송 정보";
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" onsubmit="return false;" action="order_info_edit_process.asp">
    <input type="hidden" name="mode" value="modifyflowerinfo">
    <input type="hidden" name="orderserial" value="<%= ojumun.FOneItem.FOrderSerial %>">
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
	    <td colspan="2">
	        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	    		<tr>
	    			<td width="160">
	    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>플라워 배송 정보</b>
				    </td>    				    
				    <td align="right">
				    	<input type="button" value="저장하기" class="button" onclick="javascript:SubmitForm();" <%= chkIIF(IsOldOrder,"disabled","") %>>
				    </td>
				</tr>
			</table>
	    </td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">선택</td>
	    <td bgcolor="#FFFFFF">
	        <input type="radio" name="cardribbon" value="1" <% if ojumun.FOneItem.Fcardribbon="1" then response.write "checked" %> >카드
	        &nbsp;
	        <input type="radio" name="cardribbon" value="2" <% if ojumun.FOneItem.Fcardribbon="2" then response.write "checked" %> >리본
	        &nbsp;
	        <input type="radio" name="cardribbon" value="3" <% if ojumun.FOneItem.Fcardribbon="3" then response.write "checked" %> >없음
	    </td>
	</tr>
	<tr height="25">
	    <td colspan="2" bgcolor="#FFFFFF">
	        <textarea id="[off,off,off,off][메세지]" class="textarea" rows="5" cols="45" name="message"><%= ojumun.FOneItem.Fmessage %></textarea>
	        <br>
	        
	    </td>
	</tr>
	<tr height="25">
	    <td bgcolor="<%= adminColor("topbar") %>">FROM</td>
	    <td bgcolor="#FFFFFF">
	        <input id="[on,off,1,16][FromName]" type="text" class="text" name="fromname" value="<%= ojumun.FOneItem.Ffromname %>" size="20" maxlength="20">
	    </td>
	</tr>
	<tr height="85" bgcolor="#FFFFFF">
	    <td bgcolor="<%= adminColor("topbar") %>">희망일</td>
	    <td bgcolor="#FFFFFF">
	        <% DrawFlowerOneDateBox SplitValue(ojumun.FOneItem.Freqdate,"-",0),SplitValue(ojumun.FOneItem.Freqdate,"-",1),SplitValue(ojumun.FOneItem.Freqdate,"-",2), ojumun.FOneItem.Freqtime %>
	        <br>
	        * 플라워 휴일 배송은 특별공지가<br>
	          없는 날은 불가능합니다.
        </td>
	</tr>
	</form>
</table>
<%
set ojumun = Nothing
%>

<script language='javascript'>
    <% if (AlertMsg<>"") then %>
        alert('<%= AlertMsg %>');
    <% end if %>
</script>    
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->