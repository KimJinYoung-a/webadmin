<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/oldmisendcls.asp"-->
<%
dim orderserial,deliveryno
dim didx, mode
didx = requestCheckVar(request("didx"),10)
mode = requestCheckVar(request("mode"),16)

deliveryno = requestCheckVar(request("deliveryno"),16)
orderserial = requestCheckVar(request("orderserial"),16)

if (orderserial = "") then
    orderserial = "-"
end if

dim omasterwithcs
set omasterwithcs = new COldMiSend
omasterwithcs.FRectOrderSerial = orderserial
omasterwithcs.FRectDeliveryNo = deliveryno
omasterwithcs.GetOneOrderMasterWithCS

orderserial = omasterwithcs.FOneItem.FOrderSerial

dim omisendList
set omisendList = new COldMiSend
omisendList.FRectOrderSerial = orderserial
omisendList.GetMiSendOrderDetailList

dim i
%>
<script language='javascript'>
function confirmSubmit(){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        document.frmmisend.submit();
    }
}

function popMisendInput(iidx){
    var popwin = window.open('popMisendInput.asp?idx=' + iidx,'popMisendInput','width=440,height=300,scrollbars=yes,resizable=yes');
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

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name=frmsearch>
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
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
			&nbsp;
			�ڵ�����ȣ : <%= omasterwithcs.FOneItem.FBuyHp %>
			&nbsp;
			�̸��� : <%= omasterwithcs.FOneItem.FBuyEmail %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="SearchThis();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="csbutton" value="ó����������" onclick="confirmSubmit();" disabled>
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
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
		<td><font color="red"><b><% if omisendList.FItemList(i).FItemLackNo=0 then response.write "-" else  response.write  omisendList.FItemList(i).FItemLackNo end if%></b></font></td>
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
		    <% if (Not IsNULL(omisendList.FItemList(i).getBeasongDPlusDate)) and (omisendList.FItemList(i).getBeasongDPlusDate<>"")  then %>
    		    <% if (omisendList.FItemList(i).getBeasongDPlusDate>=2) then %>
    		    <strong><font color="Red"><%= omisendList.FItemList(i).getBeasongDPlusDateStr %></font></strong>
    		    <% else %>
    		    <%= omisendList.FItemList(i).getBeasongDPlusDateStr %>
    		    <% end if %>
		    <% else %>
		    <%= omisendList.FItemList(i).getBeasongDPlusDateStr %>
		    <% end if %>
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
				<% if (omisendList.FItemList(i).FMisendReason = "05") then %>
				<br><acronym title="<%= omisendList.FItemList(i).FMiRegDate %>"><%= omisendList.FItemList(i).FMiRegUserid %></acrpnym>
				<% end if %>
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
		    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="Y") and (omisendList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')">N</a>
    	    <% end if %>
<!--
    	    <% if omisendList.FItemList(i).FMisendReason = "05" then %>
	    		<a href="javascript:PopCSSMSSend('<%= omisendList.FItemList(i).FBuyHp %>','<%= omisendList.FItemList(i).FOrderSerial %>','<%= omisendList.FItemList(i).FUserID %>','[�ٹ����� ǰ���ȳ�]�ֹ��Ͻ� ��ǰ�� <%= omisendList.FItemList(i).FItemName %>(<%= omisendList.FItemList(i).FItemID %>)��ǰ�� ǰ���Ǿ� �߼��� �Ұ��մϴ�.���ο� ������ ��� �˼��մϴ�');">N</a>
	    	<% elseif omisendList.FItemList(i).FMisendReason = "03" then %>
	    		<a href="javascript:PopCSSMSSend('<%= omisendList.FItemList(i).FBuyHp %>','<%= omisendList.FItemList(i).FOrderSerial %>','<%= omisendList.FItemList(i).FUserID %>','[�ٹ����� ��������ȳ�]�ֹ��Ͻ� ��ǰ�� <%= omisendList.FItemList(i).FItemName %>(<%= omisendList.FItemList(i).FItemID %>)��ǰ�� <%= omisendList.FItemList(i).FmiSendIpgodate %>�� �߼۵� �����Դϴ�');">N</a>
	    	<% elseif omisendList.FItemList(i).FMisendReason = "01" then %>
	    		<a href="javascript:PopCSSMSSend('<%= omisendList.FItemList(i).FBuyHp %>','<%= omisendList.FItemList(i).FOrderSerial %>','<%= omisendList.FItemList(i).FUserID %>','[�ٹ����� ��������ȳ�]�ֹ��Ͻ� ��ǰ�� <%= omisendList.FItemList(i).FItemName %>(<%= omisendList.FItemList(i).FItemID %>)��ǰ�� <%= omisendList.FItemList(i).FmiSendIpgodate %>�� �߼۵� �����Դϴ�');"> N </a>
	    	<% else %>
	    	    N
	    	<% end if %>
 -->
		</td>
		<td>
		    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')"><%= omisendList.FItemList(i).FisSendEmail %></a>
		    <% elseif (omisendList.FItemList(i).FIsUpcheBeasong="Y") and (omisendList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMisendInput('<%= omisendList.FItemList(i).FIDx %>')">N</a>
    	    <% end if %>

<!--
    			<% if omisendList.FItemList(i).FMisendReason = "05" then %>
    	    		<a href="javascript:PopCSMailSend('<%= omisendList.FItemList(i).FBuyEmail %>','<%= omisendList.FItemList(i).FOrderSerial %>','<%= omisendList.FItemList(i).FUserID %>');">N</a>
    	    	<% elseif omisendList.FItemList(i).FMisendReason = "03" then %>
    	    		<a href="javascript:PopCSMailSend('<%= omisendList.FItemList(i).FBuyEmail %>','<%= omisendList.FItemList(i).FOrderSerial %>','<%= omisendList.FItemList(i).FUserID %>');">N</a>
    	    	<% elseif omisendList.FItemList(i).FMisendReason = "01" then %>
    	    		<a href="javascript:PopCSMailSend('<%= omisendList.FItemList(i).FBuyEmail %>','<%= omisendList.FItemList(i).FOrderSerial %>','<%= omisendList.FItemList(i).FUserID %>');"> N </a>
    	    	<% else %>
    	    	    N
    	    	<% end if %>
-->
		</td>
		<td>
		    <% if omisendList.FItemList(i).FMisendReason<>"" then %>
		        <% if (omisendList.FItemList(i).FisSendCall<>"Y") then %>
		            <a href="javascript:popSendCallChange('<%= omisendList.FItemList(i).FIDx %>')"><%= omisendList.FItemList(i).FisSendCall %></a>
		        <% else %>
		            <%= omisendList.FItemList(i).FisSendCall %>
		        <% end if %>

    	    <% end if %>
<!--
		    <% if omisendList.FItemList(i).FisSendCall="Y" then %>
		        <%= omisendList.FItemList(i).FisSendCall %>
		    <% else %>
    			<% if (omisendList.FItemList(i).FMisendReason<>"") and (omisendList.FItemList(i).FMisendReason<>"00") then %>
    			N
    			<% end if %>
    		<% end if %>
  -->
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

<!-- #include virtual="/cscenterv2/lib/poptail.asp"-->
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
