<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����� ��ǰ����Ʈ
' Hieditor : �̻� ����
'			 2019.01.16 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<%
dim orderserial,deliveryno, detailcancelyn, didx, mode, reload, i
dim sellsite
	didx = requestCheckvar(request("didx"),20)
	mode = request("mode")
	detailcancelyn = requestCheckvar(request("detailcancelyn"),2)
	reload = requestCheckvar(request("reload"),2)
	deliveryno = requestCheckvar(request("deliveryno"),30)
	orderserial = requestCheckvar(request("orderserial"),20)

if (orderserial = "") then
    orderserial = "-"
end if
if reload="" and detailcancelyn="" then detailcancelyn="Y"

dim omasterwithcs
set omasterwithcs = new COldMiSend
	omasterwithcs.FRectOrderSerial = orderserial
	omasterwithcs.FRectDeliveryNo = deliveryno
	omasterwithcs.GetOneOrderMasterWithCS

orderserial = omasterwithcs.FOneItem.FOrderSerial
sellsite = omasterwithcs.FOneItem.Fsitename

dim omisendList
set omisendList = new COldMiSend
	omisendList.FRectOrderSerial = orderserial
	omisendList.frectdetailcancelyn = detailcancelyn
	omisendList.GetMiSendOrderDetailList


'// ���밡�� API
dim availApiCS : availApiCS = "stockout"
if (sellsite = "coupang") or (sellsite = "interpark") then
	availApiCS = "cancel"
end if

%>
<script type="text/javascript">

function confirmSubmit(){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        document.frmmisend.submit();
    }
}

function popMisendInput(iidx){
    var popwin = window.open('/partner/jumunmaster/popMisendInput.asp?idx=' + iidx,'popMisendInput','width=1200,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popSendCallChange(iidx){
    if (confirm('���Բ� �ȳ���ȭ�� ��Ƚ��ϱ�?')){
        frmmisendOne.detailIDx.value=iidx;
        frmmisendOne.submit();
    }
}

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
	frmsearch.submit();
}

function jsSendStockOut(detailidx) {
	var sellsite = '<%= sellsite %>';
    var orderserial = '<%= orderserial %>';
	var url;

	switch (sellsite) {
		case 'ssg':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutOne&detailidx=" + detailidx + '&orderserial=' + orderserial;
			break;
		case '11st1010':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutOne&detailidx=" + detailidx + '&orderserial=' + orderserial;
			break;
		case 'nvstorefarm':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutOne&detailidx=" + detailidx + '&orderserial=' + orderserial;
			break;
        case 'gmarket1010':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=cancelAll&detailidx=" + detailidx + '&orderserial=' + orderserial;
			break;
		case 'WMP':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=cancelAll&detailidx=" + detailidx + '&orderserial=' + orderserial;
			break;
		case 'wmpfashion':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=cancelAll&detailidx=" + detailidx + '&orderserial=' + orderserial;
			break;
		default:
			alert('�������� �ʴ� ���޸��Դϴ�.[' + sellsite + ']');
			return;
	}

	if (confirm('��ǰ��� : �����Ͻðڽ��ϱ�?')) {
		var popwin = window.open(url,'_blank','width=500,height=300');
		popwin.focus();
	}
}

function jsSendStockCnclOut(detailidx) {
	var sellsite = '<%= sellsite %>';
	var url;

	switch (sellsite) {
		case 'ssg':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutCnclOne&detailidx=" + detailidx;
			break;
		//case 'lotteCom':
		//	url = "http://wapi.10x10.co.kr/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutCnclOne&detailidx=" + detailidx;
		//	break;
		default:
			alert('�������� �ʴ� ���޸��Դϴ�.[' + sellsite + ']');
			return;
	}

	if (confirm('��ǰ������ : �����Ͻðڽ��ϱ�?')) {
		var popwin = window.open(url,'_blank','width=500,height=300');
		popwin.focus();
	}
}

function jsSendStockOutAll() {
	var sellsite = '<%= sellsite %>';
	var orderserial = '<%= orderserial %>';
	var url;

	switch (sellsite) {
		case 'ssg':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutAll&orderserial=" + orderserial;
			break;
		//case 'lotteCom':
		//	url = "http://wapi.10x10.co.kr/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=stockoutAll&orderserial=" + orderserial;
		//	break;
		default:
			alert('�������� �ʴ� ���޸��Դϴ�.[' + sellsite + ']');
			return;
	}

	if (confirm('��ǰ��� : �����Ͻðڽ��ϱ�?')) {
		var popwin = window.open(url,'_blank','width=500,height=300');
		popwin.focus();
	}
}

function jsSendCancelAll() {
	var sellsite = '<%= sellsite %>';
	var orderserial = '<%= orderserial %>';
	var url;
	var arrdetailidx = "";
	var arritemno = "";
	var chk, orderitemno, itemno, i, j, k;

	for (i = 0; ; i++) {
		chk = document.getElementById('chk_' + i);
		orderitemno = document.getElementById('orderitemno_' + i);
		itemno = document.getElementById('itemno_' + i);

		if (chk == undefined) { break; }
		if (chk.disabled == true) { continue; }
		if (chk.checked != true) { continue; }

		if ((itemno.value == "") || (itemno.value*0 != 0)) {
			alert('��Ҽ����� ���ڸ� �����մϴ�.');
			itemno.focus();
			return;
		}
		if (itemno.value*1 <= 0) {
			alert('��Ҽ����� 0���� Ŀ���մϴ�.');
			itemno.focus();
			return;
		}
		if (itemno.value*1 > orderitemno.value*1) {
			alert('��Ҽ����� �ֹ��������� Ŭ �� �����ϴ�.');
			itemno.focus();
			return;
		}

		arrdetailidx = arrdetailidx + ',' + chk.value;
		arritemno = arritemno + ',' + itemno.value;
	}

	if (arrdetailidx == "") {
		alert('���õ� ��ǰ�� �����ϴ�.');
		return;
	}

	switch (sellsite) {
		case 'coupang':
			url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=cancelAll&orderserial=" + orderserial + "&detailidx=" + arrdetailidx + "&itemno=" + arritemno;
			break;
		case 'interpark':
			if (confirm('������ũ�� ���ü����� �����ϰ� ���û�ǰ ���ΰ� ��ҵ˴ϴ�.\n\n�����Ͻðڽ��ϱ�?')) {
				url = "<%=apiURL%>/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=cancelAll&orderserial=" + orderserial + "&detailidx=" + arrdetailidx + "&itemno=" + arritemno;
			} else {
				return;
			}

			break;
		//case 'cjmall':
		//	url = "http://wapi.10x10.co.kr/outmall/order/xSite_CS_Stock_Snd_Process.asp?sellsite=" + sellsite + "&mode=cancelAll&orderserial=" + orderserial + "&detailidx=" + arrdetailidx + "&itemno=" + arritemno;
		//	break;
		default:
			alert('�������� �ʴ� ���޸��Դϴ�.[' + sellsite + ']');
			return;
	}

	if (confirm('�ֹ���� : �����Ͻðڽ��ϱ�?')) {
		var popwin = window.open(url,'_blank','width=500,height=300');
		popwin.focus();
	}
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

<!-- �˻� ���� -->
<form name="frmsearch" style="margin:0px;" method="get">
<input type="hidden" name="reload" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>" height="25">�˻�<br>����</td>
		<td align="left">
			�ֹ���ȣ : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size=13 >
        	<% if omasterwithcs.FOneItem.FCancelyn<>"N" then %>
				<b><font color="#CC3333">[����ֹ�]</font></b>
				<script language='javascript'>alert('��ҵ� �ŷ� �Դϴ�.');</script>
			<% else %>
				[�����ֹ�]
			<% end if %>

			&nbsp;
			&nbsp;
			���� : <%= omasterwithcs.FOneItem.FBuyName %>

			<% if C_CriticInfoUserLV1 then %>
				&nbsp;
				�ڵ�����ȣ : <%= omasterwithcs.FOneItem.FBuyHp %>
				&nbsp;
				�̸��� : <%= omasterwithcs.FOneItem.FBuyEmail %>
		    <% else %>
				&nbsp;
				�ڵ�����ȣ : XXX-XXX-XXXX
				&nbsp;
				�̸��� :
			<% end if %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="SearchThis();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left" height="25">
			<input type="checkbox" value="Y" name="detailcancelyn" <% if detailcancelyn="Y" then response.write " checked" %> > ����ֹ�����
		</td>
	</tr>
</table>
</form>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10px; padding-bottom:10px;">
	<tr>
		<td align="left">
			<input type="button" class="csbutton" value="ó����������" onclick="confirmSubmit();">
			<% if (sellsite<>"10x10") then %>
			&nbsp;
			���޸� : <%= sellsite %>
            <%
            select case sellsite
                case "ssg"
                    response.write "(��ǰ�� or �ϰ�ǰ�� ���۰���, ��ǰ�� ǰ�������� ���۰���)"
                case "11st1010"
                    response.write "(��ǰ�� ǰ�� ���۰���)"
                case "nvstorefarm"
                    response.write "(��ǰ�� ��ҽ�û����/��ǰ�ֹ���� ���۰���)"
                case "coupang"
                    response.write "(���û�ǰ ���޸� ������� ����, ���� �Ϻ���� ����)"
                case "interpark"
                    response.write "(���û�ǰ ���޸� ������� ����, ���� �Ϻ���� ����)"
                case "gmarket1010"
                    response.write "(���û�ǰ ���޸� ������� ����, �����������)"
                case "WMP", "wmpfashion"
                    response.write "(���û�ǰ ���޸� ������� ����, �����������)"
                case else
                    response.write "(API �۾�����)"
            end select
            %>
			<% end if %>
		</td>
		<td align="right">
			<% if (sellsite<>"10x10") then %>
			<% if (availApiCS = "stockout") then %>
			<input type="button" class="csbutton" value="ǰ������ ���޸� �ϰ�����" onClick="jsSendStockOutAll()" <%= CHKIIF(C_ADMIN_AUTH, "", "disabled") %>>
			<% end if %>
			<% if (availApiCS = "cancel") then %>
			<input type="button" class="csbutton" value="���û�ǰ ���޸� �������" onClick="jsSendCancelAll()">
			<% end if %>
			<% end if %>
		</td>
	</tr>
</table>
<!-- �׼� �� -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"></td>
		<td>�귣��</td>
		<td width="50">��ǰ�ڵ�</td>
		<td width="50">�̹���</td>
		<td>��ǰ��<font color="blue">[�ɼ�]</font></td>
		<td width="30">�ֹ�<br>����</td>
		<td width="30">����<br>����</td>
		<td width="30">���<br>����</td>
		<td width="80">��������</td>
		<td width="30">�ҿ�<br>�ϼ�</td>
		<td width="60">�������</td>
		<td width="100">��������</td>
		<td width="80">�������</td>
		<td width="120">����/��ü<br>�ۼ��޸�</td>
		<td width="35">SMS</td>
		<td width="35">MAIL</td>
		<td width="35">CALL</td>
		<td width="35">����<br />API</td>
		<td width="100">CSó������</td>
		<td width="85">CSó���޸�</td>
	</tr>
	<form name="frmmisend" method="post" action="misendmaster_main_process.asp">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<input type="hidden" name="mode" value="">
	<% for i=0 to omisendList.FResultCount -1 %>

	<% if omisendList.FItemList(i).FItemLackNo<>0 then %>
	<tr align="center" bgcolor="<%= adminColor("pink") %>">
	<% else %>
	<tr align="center" bgcolor="FFFFFF">
	<% end if %>
		<td>
			<% if (sellsite<>"10x10") and (availApiCS = "cancel") then %>
			<input type="checkbox" name="chk" id="chk_<%= i %>" value="<%= omisendList.FItemList(i).Fidx %>" <%= CHKIIF(omisendList.FItemList(i).FMisendReason="05", "", "disabled")%>>
			<% end if %>
		</td>
		<td><%= omisendList.FItemList(i).FMakerid %></td>
		<td>
			<% if omisendList.FItemList(i).FIsUpcheBeasong="Y" then %>
			<font color="red"><%= omisendList.FItemList(i).FItemID %></font>
			<% else %>
			<%= omisendList.FItemList(i).FItemID %>
			<% end if %>
		</td>
		<td><img src="<%= omisendList.FItemList(i).FSmallImage %>" width="50" height="50"></td>
		<td align="left">
			<%= omisendList.FItemList(i).FItemName %>
			<% if omisendList.FItemList(i).FItemOptionName<>"" then %>
			<br>
			<font color="blue">[<%= omisendList.FItemList(i).FItemOptionName %>]</font>
			<% end if %>
		</td>
		<td><%= omisendList.FItemList(i).FItemNo %></td>
		<td>
			<% if omisendList.FItemList(i).FItemLackNo=0 then %>
			-
			<% else %>
			<input type="hidden" name="orderitemno" id="orderitemno_<%= i %>" value="<%= omisendList.FItemList(i).FItemNo %>">
			<input type="text" class="text" name="itemno" id="itemno_<%= i %>" value="<%= omisendList.FItemList(i).FItemLackNo %>" size="1">
			<% end if %>
		</td>
		<td>
		    <%= fnColor(omisendList.FItemList(i).FDetailCancelYn,"cancelyn") %>
		</td>
		<td>
		    <% if IsNULL(omisendList.FItemList(i).FbaljuDate) then %>

		    <% else %>
		    <%= Left(omisendList.FItemList(i).FbaljuDate,10) %>
		    <% end if %>
		</td>
		<td>
		    <!-- D+2 �̻��ϰ��, ���������� ǥ�� -->
		    <%
'				If (Not IsNULL(omisendList.FItemList(i).getBeasongDPlusDate)) and (omisendList.FItemList(i).getBeasongDPlusDate<>"")  then
'					if (omisendList.FItemList(i).getBeasongDPlusDate>=2) then
'						response.write "<strong><font color='Red'>"& omisendList.FItemList(i).getBeasongDPlusDateStr &"</font></strong>"
'					else
'					response.write omisendList.FItemList(i).getBeasongDPlusDateStr
'			   		end if
'				else
'					response.write omisendList.FItemList(i).getBeasongDPlusDateStr
'				end if

				If (Not IsNULL(omisendList.FItemList(i).FDday)) and (omisendList.FItemList(i).FDday<>"")  then
					if (omisendList.FItemList(i).FDday>=2) then
						response.write "<strong><font color='Red'>"& omisendList.FItemList(i).getNewBeasongDPlusDateStr &"</font></strong>"
					else
    		    		response.write omisendList.FItemList(i).getNewBeasongDPlusDateStr
			   		end if
				else
					response.write omisendList.FItemList(i).getNewBeasongDPlusDateStr
				end if
			%>
		</td>
		<td>
		    <font color="<%= omisendList.FItemList(i).getUpCheDeliverStateColor %>"><%= omisendList.FItemList(i).getUpCheDeliverStateName %></font>
		</td>
		<td>
			<% if (Trim(omisendList.FItemList(i).FPrevMisendReason) <> "") then %>
				<%= MiSendCodeToName(omisendList.FItemList(i).FPrevMisendReason) %><br>
				-&gt;
			<% end if %>
			<% if Not IsNull(omisendList.FItemList(i).FMisendReason) and (CStr(omisendList.FItemList(i).FIDx)=Cstr(didx)) then %>
				<font color="red">�Է���</font>
			<% else %>
				<font color="<%= omisendList.FItemList(i).getMiSendCodeColor %>"><%= omisendList.FItemList(i).getMiSendCodeName %></font>
				<% if True or (omisendList.FItemList(i).FMisendReason = "05") then %>
				<br><acronym title="<%= omisendList.FItemList(i).FMiRegDate %>"><%= omisendList.FItemList(i).FMiRegUserid %></acrpnym>
				<% end if %>
			<% end if %>
			<% if Not IsNull(omisendList.FItemList(i).Freqreguserid) then %>
				<br /><%= omisendList.FItemList(i).Freqreguserid %>
			<% end if %>
		</td>
		<td>
			<% if (omisendList.FItemList(i).FMisendReason<>"") and (omisendList.FItemList(i).FMisendReason<>"00") and (omisendList.FItemList(i).FMisendReason<>"05") then %>
				<%= omisendList.FItemList(i).FmiSendIpgodate %>
			<% end if %>
		</td>
		<td>
			<%= omisendList.FItemList(i).FrequestString %>
			<%= nl2br(omisendList.FItemList(i).FupcheRequestString) %>
		</td>
		<td>
		    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')"><%= omisendList.FItemList(i).FisSendSMS %></a>
		    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="N") and (omisendList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')">N</a>
		    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="Y") and (omisendList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')">N</a>
    	    <% end if %>
		</td>
		<td>
		    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')"><%= omisendList.FItemList(i).FisSendEmail %></a>
		    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="N") and (omisendList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')">N</a>
		    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="Y") and (omisendList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')">N</a>
    	    <% end if %>
		</td>
		<td>
		    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
		        <% if (omisendList.FItemList(i).FisSendCall<>"Y") then %>
		            <a href="javascript:popSendCallChange('<%= omisendList.FItemList(i).FIDx %>')"><%= omisendList.FItemList(i).FisSendCall %></a>
		        <% else %>
		            <%= omisendList.FItemList(i).FisSendCall %>
		        <% end if %>
    	    <% end if %>
		</td>
		<td>
			<% if (sellsite<>"10x10") then %>
				<%= omisendList.FItemList(i).FisSendAPI %><br />
				<% if (omisendList.FItemList(i).FMisendReason="05") and (omisendList.FItemList(i).FisSendAPI = "N") then %>
				<input type="button" class="csbutton" value="����" onClick="jsSendStockOut(<%= omisendList.FItemList(i).Fidx %>)">
				<% elseif (omisendList.FItemList(i).FisSendAPI = "Y") then %>
				<input type="button" class="csbutton" value="���" onClick="jsSendStockCnclOut(<%= omisendList.FItemList(i).Fidx %>)">
                <% elseif (orderserial = "20030XXX942396") then %>
                <input type="button" class="csbutton" value="����" onClick="jsSendStockOut(<%= omisendList.FItemList(i).Fidx %>)">
				<% end if %>
			<% end if %>
		</td>

		<% if (omisendList.FItemList(i).FMisendReason <> "") then %>
		<input type="hidden" name="didx" value="<%= omisendList.FItemList(i).FIDx %>">
		<% end if %>

		<td>
		<% if (omisendList.FItemList(i).FMisendReason <> "") then %>

			<input type=hidden name=prevstate value="<%= omisendList.FItemList(i).FmiSendState %>">

		      <% if (omisendList.FItemList(i).FmiSendState = "7") then %>
		      �Ϸ�
		      <input type=hidden name=state value="7">
		      <% else %>
		  	<select class="select" name="state">
				<option value="0" <% if (omisendList.FItemList(i).FmiSendState = "0") then response.write "selected" end if %>>��ó��</option>
				<!-- <option value="1" <% if (omisendList.FItemList(i).FmiSendState = "1") then response.write "selected" end if %>>SMS�Ϸ�</option> -->
				<!-- <option value="2" <% if (omisendList.FItemList(i).FmiSendState = "2") then response.write "selected" end if %>>�ȳ�Mail�Ϸ�</option> -->
				<!-- <option value="3" <% if (omisendList.FItemList(i).FmiSendState = "3") then response.write "selected" end if %>>��ȭ�Ϸ�</option> -->
				<!-- <option value="3" <% if (omisendList.FItemList(i).FmiSendState = "3") then response.write "selected" end if %>>��۽�ó��</option> -->
				<option value="4" <% if (omisendList.FItemList(i).FmiSendState = "4") then response.write "selected" end if %>>���ȳ�</option><!-- �ű�(SMS/mail/��ȭ��) -->
				<option value="6" <% if (omisendList.FItemList(i).FmiSendState = "6") then response.write "selected" end if %>>CSó���Ϸ�</option>
		  	</select>
		      <% end if %>
		  <% end if %>
		</td>
		<td>
		  <% if (omisendList.FItemList(i).FMisendReason <> "") then %>
		  <input type="text" class="text" name="finishstr" value="<%= omisendList.FItemList(i).FfinishString %>" size="10">
		  <% end if %>
		</td>
	</tr>
	<% next %>
	</form>
</table>

<form name="frmmisendOne" method="post" action="misendmaster_main_process.asp">
<input type="hidden" name="mode" value="SendCallChange">
<input type="hidden" name="orderserial" value="<%= orderserial %>">
<input type="hidden" name="detailIDx" value="">
</form>
<!-- ǥ �ϴܹ� ��-->

<%
set omasterwithcs = Nothing
set omisendList = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
