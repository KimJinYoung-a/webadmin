<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/baljucls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<%
dim orderserial,deliveryno,obalju
dim didx, mode
didx = request("didx")
mode = request("mode")

deliveryno = request("deliveryno")
orderserial = request("orderserial")

if (orderserial = "") then
        orderserial = "-"
end if

dim omasterwithcs
set omasterwithcs = new COldMiSend
omasterwithcs.FRectOrderSerial = orderserial
omasterwithcs.FRectDeliveryNo = deliveryno
omasterwithcs.GetOneOrderMasterWithCS

orderserial = omasterwithcs.FOneItem.FOrderSerial

set obalju = New CBalju
obalju.FRectOrderSerial = orderserial
obalju.GetMiSendOrderDetail

dim i
%>
<script language='javascript'>
function DelMiSend(frm){
	var ret = confirm('�����Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="del";
		frm.submit();
	}
}

function SaveMiSend(frm){
	if (frm.ipgodate.value.length>0){
		if (frm.ipgodate.value.length!=10){
			alert('�԰������� ��Ȯ�� �Է��ϼ���.');
			frm.ipgodate.focus();
			return;
		}
	}

	var ret = confirm('�����Ͻðڽ��ϱ�?');

	if (ret){
		frm.submit();
	}
}
function calender_open() {
       document.all.cal.style.display="";
}

function SearchThis(){
	location.href="/admin/ordermaster/misendmaster_main.asp?orderserial=" + frmsearch.orderserial.value;
}

</script>
<style type="text/css">
<!--
td { font-size:9pt; font-family:Verdana;}

.button {
	font-family: "����", "����";
	font-size: 10px;
	background-color: #E4E4E4;
	border: 1px solid #000000;
	color: #000000;
	height: 20px;
}
//-->
</style>



<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<form name=frmsearch>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�ֹ���ȣ : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size=13 >
	        	<input type="button" class="button_s" value="�˻�" onClick="SearchThis()">
	        	&nbsp;&nbsp;
	        	<% if omasterwithcs.FOneItem.FCancelyn="Y" then %>
				<b><font color="#CC3333">��� �ֹ����Դϴ�.</font></b>
				<script language='javascript'>alert('��ҵ� �ŷ� �Դϴ�.');</script>
				<% else %>
				���� �ֹ����Դϴ�.
				<% end if %>
		</td>
	</tr>
	</form>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">��ǰ�ڵ�</td>
		<td width="50">�̹���</td>
		<td>��ǰ��<font color="blue">[�ɼ�]</font></td>
		<td width="40">����</td>
		<td width="30">���<br>����</td>
		<td width="80">��������</td>
		<td width="30">D+</td>
		<td width="80">��������</td>
		<td width="80">�������</td>
		<td width="80">��û����</td>
		<td width="80">ó�����</td>
		<td width="80">ó������</td>
	</tr>
	<form name="frmmisend" method="post" action="domisendmaster_main.asp">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<input type="hidden" name="mode" value="">
	<% for i=0 to Ubound(obalju.FBaljuDetailList) -1 %>
	<tr align="center" bgcolor="FFFFFF">
		<% if obalju.FBaljuDetailList(i).IsUpcheBeasong then %>
		<td ><font color="red"><%= obalju.FBaljuDetailList(i).FItemID %></font></td>
		<% else %>
		<td ><%= obalju.FBaljuDetailList(i).FItemID %></td>
		<% end if %>
		<td><img src="<%= obalju.FBaljuDetailList(i).FImageSmall %>" width="50" height="50"></td>
		<td align="left">
			<%= obalju.FBaljuDetailList(i).FItemName %>
			<% if obalju.FBaljuDetailList(i).FItemOptionName<>"" then %>
			<br>
			<font color="blue">[<%= obalju.FBaljuDetailList(i).FItemOptionName %>]</font>
			<% end if %>
		</td>
		<td><%= obalju.FBaljuDetailList(i).FItemNo %></td>
		<td><font color="<%= obalju.FBaljuDetailList(i).CancelYnColor %>"><%= obalju.FBaljuDetailList(i).CancelYnName %></font></td>
		<td>2009-04-05</td>
		<td>D+2</td> <!-- D+2 �̻��ϰ��, ���������� ǥ�� -->
		<% if Not IsNull(obalju.FBaljuDetailList(i).FmiSendCode) and (CStr(obalju.FBaljuDetailList(i).FDetailIDx)=Cstr(didx)) then %>
		<td><font color="red">�Է���</font></td>
		<% else %>
		<td><font color="<%= obalju.FBaljuDetailList(i).getMiSendCodeColor %>"><%= obalju.FBaljuDetailList(i).getMiSendCodeName %></font></td>
		<% end if %>
		<td><%= obalju.FBaljuDetailList(i).FmiSendIpgodate %></td>
		<td><%= obalju.FBaljuDetailList(i).FrequestString %></td>
		<% if (obalju.FBaljuDetailList(i).FmiSendCode <> "") then %>
		<input type="hidden" name="didx" value="<%= obalju.FBaljuDetailList(i).FDetailIDx %>">
		<% end if %>
		<td>
		  <% if (obalju.FBaljuDetailList(i).FmiSendCode <> "") then %>
		  <input type="text" name="finishstr" value="<%= obalju.FBaljuDetailList(i).FfinishString %>" size="10">
		  <% end if %>
		</td>
		<td>
		  <% if (obalju.FBaljuDetailList(i).FmiSendCode <> "") then %>
		      <% if (obalju.FBaljuDetailList(i).FmiSendState = "7") then %>
		      �Ϸ�
		      <input type=hidden name=state value="7">
		      <% else %>
		  <select name="state">
		    <option value="0" <% if (obalju.FBaljuDetailList(i).FmiSendState = "0") then response.write "selected" end if %>>��ó��</option>
		    <option value="1" <% if (obalju.FBaljuDetailList(i).FmiSendState = "1") then response.write "selected" end if %>>SMS�Ϸ�</option>
		    <option value="2" <% if (obalju.FBaljuDetailList(i).FmiSendState = "2") then response.write "selected" end if %>>�ȳ�Mail�Ϸ�</option>
		    <option value="3" <% if (obalju.FBaljuDetailList(i).FmiSendState = "3") then response.write "selected" end if %>>��ȭ�Ϸ�</option>
		   <!-- <option value="3" <% if (obalju.FBaljuDetailList(i).FmiSendState = "3") then response.write "selected" end if %>>��۽�ó��</option> -->
		    <option value="6" <% if (obalju.FBaljuDetailList(i).FmiSendState = "6") then response.write "selected" end if %>>CSó���Ϸ�</option>
		  </select>
		      <% end if %>
		  <% end if %>
		</td>
	</tr>
	<% next %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center"><input type="button" value=" ó���Է� " onclick="document.frmmisend.submit();"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</form>
</table>
<!-- ǥ �ϴܹ� ��-->


<%
set omasterwithcs = Nothing
set obalju = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->