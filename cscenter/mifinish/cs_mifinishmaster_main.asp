<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : cs����
' History	:  2007.06.01 �̻� ����
'              2023.11.15 �ѿ�� ����(6�������� �����͵� ó�������ϰ� ���� ����)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_mifinishcls.asp"-->
<%
dim asid, orderserial, csdetailidx, ocsmifinishmaster, ocsmifinishDetailList, isChulgoState, i
	asid = requestCheckVar(request("asid"),10)
	csdetailidx = request("csdetailidx")

set ocsmifinishmaster = new CCSMifinishMaster
	ocsmifinishmaster.FRectAsid = asid
	ocsmifinishmaster.GetOneCSMaster

	if ocsmifinishmaster.FtotalCount < 1 then
		ocsmifinishmaster.FRectAsid = asid
		ocsmifinishmaster.FRectorder6MonthBefore = "Y"
		ocsmifinishmaster.GetOneCSMaster
	end if

orderserial = ocsmifinishmaster.FOneItem.FOrderSerial

set ocsmifinishDetailList = new CCSMifinishMaster
	ocsmifinishDetailList.FRectAsid = asid
	ocsmifinishDetailList.getMiFinishCSDetailList

	if ocsmifinishDetailList.FTotalCount < 1 then
		ocsmifinishDetailList.FRectAsid = asid
		ocsmifinishDetailList.FRectorder6MonthBefore = "Y"
		ocsmifinishDetailList.getMiFinishCSDetailList
	end if

isChulgoState = (ocsmifinishmaster.FOneItem.Fdivcd = "A000") or (ocsmifinishmaster.FOneItem.Fdivcd = "A100")

%>
<script type="text/javascript">

function confirmSubmit(){
    if (confirm('���� �Ͻðڽ��ϱ�?')) {
    	var arrfinishstr = document.getElementsByName("finishstr");

    	for (var i = 0; i < arrfinishstr.length; i++) {
    		// ��ǥ �ٲٱ�
    		arrfinishstr[i].value = arrfinishstr[i].value.replace(/,/g, "_XX_");
    	}

        document.frmmisend.submit();
    }
}

function popMifinishInput(csdetailidx) {
    var popwin = window.open('/cscenter/mifinish/popMifinishInput.asp?csdetailidx=' + csdetailidx,'popMifinishInput','width=1400,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popSendCallChange(iidx){
    if (confirm('���Բ� �ȳ���ȭ�� ��Ƚ��ϱ�?')){
        frmmisendOne.csdetailidx.value=iidx;
        frmmisendOne.submit();
    }
}

function SearchThis(){
	location.href="/cscenter/mifinish/cs_mifinishmaster_main.asp?asid=" + frmsearch.asid.value;
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
<form name="frmsearch" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			ASID : <input type="text" class="text" name="asid" value="<%= asid %>" size=13 >
        	<% if ocsmifinishmaster.FOneItem.Fdeleteyn<>"N" then %>
			<b><font color="#CC3333">[���CS]</font></b>
			<script language='javascript'>alert('��ҵ� CS �Դϴ�.');</script>
			<% else %>
			[����CS]
			<% end if %>
			&nbsp;
			&nbsp;
			���� : <font color="<%= ocsmifinishmaster.FOneItem.getDivcdColor %>"><%= ocsmifinishmaster.FOneItem.getDivcdStr %></font>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="SearchThis();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" height="30">
		<td align="left">
			�ֹ���ȣ : <%= orderserial %>
			&nbsp;
			���� : <%= ocsmifinishmaster.FOneItem.FBuyName %>
			&nbsp;
			�ڵ�����ȣ : <%= ocsmifinishmaster.FOneItem.FBuyHp %>
			&nbsp;
			�̸��� : <%= ocsmifinishmaster.FOneItem.FBuyEmail %>
		</td>
	</tr>
</table>
</form>
<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="csbutton" value="ó����������" onclick="confirmSubmit();">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- �׼� �� -->

<form name="frmmisend" method="post" action="cs_mifinishmaster_main_process.asp" style="margin:0px;">
<input type="hidden" name="asid" value="<%= asid %>">
<input type="hidden" name="mode" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">��ǰ�ڵ�</td>
		<td width="50">�̹���</td>
		<td>��ǰ��<br><font color="blue">[�ɼ�]</font></td>
		<td width="40">����<br>����</td>
		<td width="40">����<br>����</td>
		<td width="80">ó��������</td>
		<td width="60">�ҿ�<br>�ϼ�</td>
		<td width="80">��ó������</td>
		<td width="80">ó��������</td>
		<td width="80">����/��ü<br>�ۼ��޸�</td>
		<td width="35">SMS</td>
		<td width="35">MAIL</td>
		<td width="35">CALL</td>
		<td width="120">CSó������</td>
		<td width="100">CSó���޸�</td>
	</tr>
	<% for i=0 to ocsmifinishDetailList.FResultCount -1 %>

	<% if ocsmifinishDetailList.FItemList(i).FItemLackNo<>0 then %>
	<tr align="center" bgcolor="<%= adminColor("pink") %>">
	<% else %>
	<tr align="center" bgcolor="FFFFFF">
	<% end if %>
		<td>
			<% if ocsmifinishDetailList.FItemList(i).FIsUpcheBeasong="Y" then %>
			<font color="red"><%= ocsmifinishDetailList.FItemList(i).FItemID %></font>
			<% else %>
			<%= ocsmifinishDetailList.FItemList(i).FItemID %>
			<% end if %>
		</td>
		<td><img src="<%= ocsmifinishDetailList.FItemList(i).FSmallImage %>" width="50" height="50"></td>
		<td align="left">
			<%= ocsmifinishDetailList.FItemList(i).FItemName %>
			<% if ocsmifinishDetailList.FItemList(i).FItemOptionName<>"" then %>
			<br>
			<font color="blue">[<%= ocsmifinishDetailList.FItemList(i).FItemOptionName %>]</font>
			<% end if %>
		</td>
		<td><%= ocsmifinishDetailList.FItemList(i).FRegItemNo %></td>
		<td><font color="red"><b><% if ocsmifinishDetailList.FItemList(i).FItemLackNo=0 then response.write "-" else  response.write  ocsmifinishDetailList.FItemList(i).FItemLackNo end if%></b></font></td>
		<td>
		    <% if IsNULL(ocsmifinishDetailList.FItemList(i).FRegdate) then %>

		    <% else %>
		    <%= Left(ocsmifinishDetailList.FItemList(i).FRegdate,10) %>
		    <% end if %>
		</td>
		<td>
		    <!-- D+2 �̻��ϰ��, ���������� ǥ�� -->
		    <% if (Not IsNULL(ocsmifinishDetailList.FItemList(i).getDPlusDate)) and (ocsmifinishDetailList.FItemList(i).getDPlusDate<>"")  then %>
    		    <% if (ocsmifinishDetailList.FItemList(i).getDPlusDate>=2) then %>
    		    <strong><font color="Red"><%= ocsmifinishDetailList.FItemList(i).getDPlusDateStr %></font></strong>
    		    <% else %>
    		    <%= ocsmifinishDetailList.FItemList(i).getDPlusDateStr %>
    		    <% end if %>
		    <% else %>
		    	<%= ocsmifinishDetailList.FItemList(i).getDPlusDateStr %>
		    <% end if %>
		</td>
		<td>
			<% if Not IsNull(ocsmifinishDetailList.FItemList(i).FMifinishReason) and (CStr(ocsmifinishDetailList.FItemList(i).Fcsdetailidx)=Cstr(csdetailidx)) then %>
				<font color="red">�Է���</font>
			<% else %>
				<font color="<%= ocsmifinishDetailList.FItemList(i).getMiFinishCodeColor %>"><%= ocsmifinishDetailList.FItemList(i).getMiFinishCodeName %></font>
			<% end if %>
		</td>
		<td>
			<% if (ocsmifinishDetailList.FItemList(i).FMifinishReason<>"") and (ocsmifinishDetailList.FItemList(i).FMifinishReason<>"00") and (ocsmifinishDetailList.FItemList(i).FMifinishReason<>"05") then %>
				<%= ocsmifinishDetailList.FItemList(i).FMifinishipgodate %>
			<% end if %>
		</td>
		<td><%= ocsmifinishDetailList.FItemList(i).FrequestString %></td>
		<td>
		    <% if ocsmifinishDetailList.FItemList(i).FMifinishReason<>"" then %>
		        <a href="javascript:popMifinishInput('<%= ocsmifinishDetailList.FItemList(i).Fcsdetailidx %>')"><%= ocsmifinishDetailList.FItemList(i).FisSendSMS %></a>
		    <% elseif (ocsmifinishDetailList.FItemList(i).FIsUpcheBeasong="Y") and (ocsmifinishDetailList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMifinishInput('<%= ocsmifinishDetailList.FItemList(i).Fcsdetailidx %>')">N</a>
    	    <% end if %>
		</td>
		<td>
		    <% if ocsmifinishDetailList.FItemList(i).FMifinishReason<>"" then %>
		        <a href="javascript:popMifinishInput('<%= ocsmifinishDetailList.FItemList(i).Fcsdetailidx %>')"><%= ocsmifinishDetailList.FItemList(i).FisSendEmail %></a>
		    <% elseif (ocsmifinishDetailList.FItemList(i).FIsUpcheBeasong="Y") and (ocsmifinishDetailList.FItemList(i).FCurrstate<7) then %>
		        <a href="javascript:popMifinishInput('<%= ocsmifinishDetailList.FItemList(i).Fcsdetailidx %>')">N</a>
    	    <% end if %>
		</td>
		<td>
		    <% if ocsmifinishDetailList.FItemList(i).FMifinishReason<>"" and isChulgoState then %>
		        <% if (ocsmifinishDetailList.FItemList(i).FisSendCall<>"Y") then %>
		            <a href="javascript:popSendCallChange('<%= ocsmifinishDetailList.FItemList(i).Fcsdetailidx %>')"><%= ocsmifinishDetailList.FItemList(i).FisSendCall %></a>
		        <% else %>
		            <%= ocsmifinishDetailList.FItemList(i).FisSendCall %>
		        <% end if %>

    	    <% end if %>
		</td>

		<% if (ocsmifinishDetailList.FItemList(i).FMifinishReason <> "") then %>
		<input type="hidden" name="arrcsdetailidx" value="<%= ocsmifinishDetailList.FItemList(i).Fcsdetailidx %>">
		<% end if %>

		<td>
		  <% if (ocsmifinishDetailList.FItemList(i).FMifinishReason <> "") then %>
		      <% if (ocsmifinishDetailList.FItemList(i).FMiFinishState = "7") then %>
		      �Ϸ�
		      <input type=hidden name=state value="7">
		      <% else %>
		  	<select class="select" name="state">
				<option value="0" <% if (ocsmifinishDetailList.FItemList(i).FMiFinishState = "0") then response.write "selected" end if %>>��ó��</option>
				<option value="4" <% if (ocsmifinishDetailList.FItemList(i).FMiFinishState = "4") then response.write "selected" end if %>>���ȳ�</option><!-- �ű�(SMS/mail/��ȭ��) -->
				<option value="6" <% if (ocsmifinishDetailList.FItemList(i).FMiFinishState = "6") then response.write "selected" end if %>>CSó���Ϸ�</option>
		  	</select>
		      <% end if %>
		  <% end if %>
		</td>
		<td>
		  <% if (ocsmifinishDetailList.FItemList(i).FMifinishReason <> "") then %>
		  <input type="text" class="text" name="finishstr" value="<%= ocsmifinishDetailList.FItemList(i).FfinishString %>" size="10">
		  <% end if %>
		</td>
	</tr>
	<% next %>
</table>
</form>

<form name="frmmisendOne" method="post" action="cs_mifinishmaster_main_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="SendCallChange">
<input type="hidden" name="asid" value="<%= asid %>">
<input type="hidden" name="csdetailidx" value="">
</form>

<%
set ocsmifinishmaster = Nothing
set ocsmifinishDetailList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->