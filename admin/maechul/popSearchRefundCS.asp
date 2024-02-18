<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<%

dim targetGbn, orderserial, suborderserial, orgorderserial, chgorderserial, reqDate, reqPrice
dim page

targetGbn 		= requestCheckvar(request("targetGbn"),32)
orderserial 	= requestCheckvar(request("orderserial"),32)
suborderserial 	= requestCheckvar(request("suborderserial"),32)
orgorderserial 	= requestCheckvar(request("orgorderserial"),32)
chgorderserial 	= requestCheckvar(request("chgorderserial"),32)
reqDate 		= requestCheckvar(request("reqDate"),32)
reqPrice 		= requestCheckvar(request("reqPrice"),32)

if page="" then page=1


'// ============================================================================
dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FPageSize = 20
ocsaslist.FCurrPage = page

if (orgorderserial <> "") then
	ocsaslist.FRectOrderSerial = orgorderserial
else
	ocsaslist.FRectOrderSerial = orderserial
end if

ocsaslist.GetCSASMasterListByProcedure


'// ============================================================================
'// ���γ���
Dim oCPGData
set oCPGData = new CPGData
	oCPGData.FPageSize = 50
	oCPGData.FCurrPage = 1
	''��Ī�� �Ŀ��� 77�� ������ ��Ī���� �Ѿ���� �ð��� �־ ���� ��Ī�� ������ ǥ�þʵȴ�.
	oCPGData.FRectShowJumunLog = "Y"

	oCPGData.FRectSearchField = "orderserial"
	if (orgorderserial <> "") then
		oCPGData.FRectSearchText = orgorderserial
	else
		oCPGData.FRectSearchText = orderserial
	end if

    oCPGData.getPGDataList_ON


'// ============================================================================
'// ���γ���(�ΰŽ� ��ǰ)
Dim oCPGDataACA
set oCPGDataACA = new CPGData
	oCPGDataACA.FPageSize = 50
	oCPGDataACA.FCurrPage = 1
	''��Ī�� �Ŀ��� 77�� ������ ��Ī���� �Ѿ���� �ð��� �־ ���� ��Ī�� ������ ǥ�þʵȴ�.
	oCPGDataACA.FRectShowJumunLog = "Y"

	oCPGDataACA.FRectSearchField = "orderserial"
	if (orgorderserial <> "") then
		'// �ΰŽ��� ��ǰ�� ��� ���̳ʽ� �ֹ���ȣ�� ��Ī�Ǿ� �ִ�.
		oCPGDataACA.FRectSearchText = orderserial
		oCPGDataACA.getPGDataList_ON
	end if


'// ============================================================================
'// OFF���γ���(�ΰŽ� ���� �������)
Dim oCPGDataACAbyHand
set oCPGDataACAbyHand = new CPGData
	oCPGDataACAbyHand.FPageSize = 50
	oCPGDataACAbyHand.FCurrPage = 1
	''��Ī�� �Ŀ��� 77�� ������ ��Ī���� �Ѿ���� �ð��� �־ ���� ��Ī�� ������ ǥ�þʵȴ�.
	oCPGDataACAbyHand.FRectExcMatchFinish = "Y"
	oCPGDataACAbyHand.FRectDateType = "A"
	oCPGDataACAbyHand.FRectStartdate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	oCPGDataACAbyHand.FRectEndDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())) + 1, 1)
	oCPGDataACAbyHand.FRectshopid = "cafe003"
	''oCPGDataACAbyHand.FRectSearchField = "cardPrice"
	''oCPGDataACAbyHand.FRectSearchText = reqPrice

	if (targetGbn = "AC") then
		'// �ΰŽ� ���� ī��� ������
		oCPGDataACAbyHand.getPGDataList_OFF
	end if


'// ============================================================================
'// ��ȯ�ֹ�
Dim oCChgPGData
set oCChgPGData = new CPGData
	oCChgPGData.FPageSize = 50
	oCChgPGData.FCurrPage = 1
	''��Ī�� �Ŀ��� 77�� ������ ��Ī���� �Ѿ���� �ð��� �־ ���� ��Ī�� ������ ǥ�þʵȴ�.
	''oCPGData.FRectIncJumunLog = "Y"

	oCChgPGData.FRectSearchField = "orderserial"
	if (chgorderserial <> "") then
		oCChgPGData.FRectSearchText = chgorderserial
		oCChgPGData.getPGDataList_ON
	end if


'// ============================================================================
dim i

%>

<script language='javascript'>

function SubmitSearch(frm) {

	/*
	if (frm.serchjeokyoyn.checked == true) {
		if (frm.jeokyo.value == "") {
			alert("���並 �Է��ϼ���");
			frm.jeokyo.focus();
			return;
		}
	}

	if (frm.serchtxammountyn.checked == true) {
		if (frm.txammount.value == "") {
			alert("�Աݾ��� �Է��ϼ���");
			frm.txammount.focus();
			return;
		}

		if (frm.txammount.value*0 != 0) {
			alert("�ݾ��� ���ڸ� �����մϴ�.");
			frm.txammount.focus();
			return;
		}
	}
	*/

	document.frm.submit();
}

function SubmitMatch(divcd, asid, refundPrice, finishDate) {
	var orderserial, orgorderserial, reqDate, reqPrice;

	orderserial = "<%= orderserial %>";
	orgorderserial = "<%= orgorderserial %>";
	reqDate = "<%= Left(reqDate, 10) %>";
	reqPrice = "<%= reqPrice %>";

	if ((divcd != "A003") && (divcd != "A007") && (divcd != "R000")) {
		alert("�߸��� �����Դϴ�.");
		return;
	}

	if (reqDate != finishDate) {
		if (confirm("��û���ڿ� ó�����ڰ� �ٸ��ϴ�.\n\n�����Ͻðڽ��ϱ�?") != true) {
			return;
		}
	}

	if (refundPrice*1 != reqPrice*-1) {
		if (confirm("�ݾ� ����ġ [��û : " + reqPrice*-1 + ", ȯ�� : " + refundPrice*1 + ", ���� " + (reqPrice*-1 - refundPrice*1)+ "]\n\n�����Ͻðڽ��ϱ�?") != true) {
			return;
		}
	}

	if (confirm("��Ī�Ͻðڽ��ϱ�?") == true) {
		var frm = document.frmAct;
		if (divcd == "R000") {
			frm.mode.value = "matchNoRefund";
		} else if (divcd == "A007") {
			frm.mode.value = "matchRefundByPGdataOn";
		} else if (divcd == "A003") {
			frm.mode.value = "matchRefundByBankOn";
		} else {
			alert("�߸��� �����Դϴ�.");
			return;
		}

		frm.asid.value = asid;

		frm.submit();
	}
}

function SubmitMatchDeposit(asid, refundPrice, finishDate) {
	var frm = document.frmAct;

	if (confirm("��Ī�Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "matchRefundByDepositOn";
		frm.asid.value = asid;
		frm.appprice.value = refundPrice;
		frm.appDate.value = finishDate;

		frm.submit();
	}
}

function SubmitMatchReBank(asid, refundPrice, finishDate) {
	var frm = document.frmAct;

	if (confirm("��Ī�Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "matchRefundByReBankOn";
		frm.asid.value = asid;
		frm.appprice.value = refundPrice;
		frm.appDate.value = finishDate;

		frm.submit();
	}
}

function SubmitMatchPGOff(pggubun, pgkey) {
	var frm = document.frmAct;

	if (confirm("��Ī�Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "matchByPGdataOff";
		frm.pggubun.value = pggubun;
		frm.pgkey.value = pgkey;

		frm.submit();
	}
}

function SubmitMatchPG(refundPrice, pgkey, pgcskey) {
	var orderserial, orgorderserial, reqDate, reqPrice;

	orderserial = "<%= orderserial %>";
	orgorderserial = "<%= orgorderserial %>";
	reqPrice = "<%= reqPrice %>";

	if (refundPrice != reqPrice) {
		if (confirm("�ݾ� ����ġ [��û : " + reqPrice + ", ���� : " + refundPrice + ", ���� " + (reqPrice*1 - refundPrice*1)+ "]\n\n�����Ͻðڽ��ϱ�?") != true) {
			return;
		}
	}

	if (confirm("��Ī�Ͻðڽ��ϱ�?.") == true) {
		var frm = document.frmAct;

		frm.mode.value = "matchByPGdataOn";

		frm.pgkey.value = pgkey;
		frm.pgcskey.value = pgcskey;
		frm.appprice.value = refundPrice;

		frm.submit();
	}
}



/*




function SubmitDisMatch(frm) {
	if (confirm("��Ī���� �����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "dismatch";
		frm.submit();
	}
}
*/

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="1" width="50" height="60" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<% if (orgorderserial <> "") then %>
			<b>���ֹ���ȣ</b> : <%= orgorderserial %>
			<% else %>
			<b>�ֹ���ȣ</b> : <%= orderserial %>
			<% end if %>

			&nbsp;
			<b>��û��</b> : <acronym title="<%= reqDate %>"><%= Left(reqDate,10) %></acronym>
			&nbsp;
			<b>��û�ݾ�</b> : <%= FormatNumber(reqPrice, 0) %> ��
		</td>
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

[�¶��� CS����]<br />
* <font color="red">PG �� ���γ����� ���� ����</font>�� ��Ī�����մϴ�.<br />
(PG�� ���γ����� �ִ� ���, PG�� ���γ����� ��Ī�ؾ� �մϴ�.)
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= ocsaslist.FTotalCount %></b>
		</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="50" align="center">Idx</td>
		<td width="100" align="center">����</td>
		<td width="90" align="center">���ֹ���ȣ</td>
		<td width="90" align="center">Site</td>
		<td align="center">����</td>
		<td width="75" align="center">����</td>
		<td width="70" align="center">ȯ�ұݾ�</td>
		<td width="80" align="center">�����</td>
		<td width="80" align="center">ó����</td>
		<td width="30" align="center">����</td>
		<td width="80" align="center">���</td>
	</tr>
<% for i = 0 to (ocsaslist.FResultCount - 1) %>
    <% if (ocsaslist.FItemList(i).Fdeleteyn <> "N") then %>
    <tr bgcolor="#EEEEEE" style="color:gray" align="center" height="25">
    <% else %>
    <tr bgcolor="#FFFFFF" align="center" height="25">
    <% end if %>
        <td height="20" nowrap><%= ocsaslist.FItemList(i).Fid %></td>
        <td nowrap align="left"><acronym title="<%= ocsaslist.FItemList(i).GetAsDivCDName %>"><font color="<%= ocsaslist.FItemList(i).GetAsDivCDColor %>"><%= ocsaslist.FItemList(i).GetAsDivCDName %></font></acronym></td>
        <td nowrap>
			<%= ocsaslist.FItemList(i).Forgorderserial %>
			<% if (ocsaslist.FItemList(i).Forderserial <> ocsaslist.FItemList(i).Forgorderserial) then %>
				+
			<% end if %>
        </td>
        <td nowrap><%= ocsaslist.FItemList(i).FExtsitename %></td>
        <td nowrap align="left"><acronym title="<%= ocsaslist.FItemList(i).Ftitle %>"><%= ocsaslist.FItemList(i).Ftitle %></acronym></td>
        <td nowrap><font color="<%= ocsaslist.FItemList(i).GetCurrstateColor %>"><%= ocsaslist.FItemList(i).GetCurrstateName %></font></td>
        <td nowrap align="right"><%= FormatNumber(ocsaslist.FItemList(i).Frefundrequire,0) %></td>
        <td nowrap><acronym title="<%= ocsaslist.FItemList(i).Fregdate %>"><%= Left(ocsaslist.FItemList(i).Fregdate,10) %></acronym></td>
        <td nowrap><acronym title="<%= ocsaslist.FItemList(i).Ffinishdate %>"><%= Left(ocsaslist.FItemList(i).Ffinishdate,10) %></acronym></td>
        <td nowrap>
        <% if ocsaslist.FItemList(i).Fdeleteyn="Y" then %>
        <font color="red">����</font>
        <% elseif ocsaslist.FItemList(i).Fdeleteyn="C" then %>
        <font color="red"><strong>���</strong></font>
        <% end if %>
        </td>
		<td>
			<% if (ocsaslist.FItemList(i).Fdivcd = "A003") and (ocsaslist.FItemList(i).Freturnmethod = "R910") and (ocsaslist.FItemList(i).Fdeleteyn = "N") and (ocsaslist.FItemList(i).Fcurrstate = "B007") then %>
				<input type="button" class="button" value="��Ī" class="csbutton" style="width:30px;" onclick="SubmitMatchDeposit(<%= ocsaslist.FItemList(i).Fid %>, '<%= ocsaslist.FItemList(i).Frefundrequire %>', '<%= Left(ocsaslist.FItemList(i).Ffinishdate,10) %>');">
			<% end if %>
			<!--
			<% if (ocsaslist.FItemList(i).Fdivcd = "A003") or (ocsaslist.FItemList(i).Fdivcd = "A007") then %>
				<% if (ocsaslist.FItemList(i).Fdeleteyn = "N") and (ocsaslist.FItemList(i).Fcurrstate = "B007") then %>
					<input type="button" class="button" value="��Ī" class="csbutton" style="width:30px;" onclick="SubmitMatch('<%= ocsaslist.FItemList(i).Fdivcd %>', <%= ocsaslist.FItemList(i).Fid %>, '<%= ocsaslist.FItemList(i).Frefundrequire %>', '<%= Left(ocsaslist.FItemList(i).Ffinishdate,10) %>');">
				<% end if %>
				<% if (ocsaslist.FItemList(i).Fdeleteyn <> "N") and (ocsaslist.FItemList(i).Fcurrstate <> "B007") then %>
					<input type="button" class="button" value="ȯ�Ҿ���" class="csbutton" style="width:60px;" onclick="SubmitMatch('R000', <%= ocsaslist.FItemList(i).Fid %>, '<%= ocsaslist.FItemList(i).Frefundrequire %>', '<%= Left(ocsaslist.FItemList(i).Ffinishdate,10) %>');">
				<% end if %>
			<% end if %>
			-->
		</td>
    </tr>
<% next %>
</table>

<p>

[���γ��� : �¶���+�ΰŽ�]
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= oCPGData.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCPGData.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">PG��</td>
	<td width="80">PG��id</td>
	<td width="80">�������</td>
	<!--
	<td width="270">PG��KEY</td>
	<td width="270">PG��CSKEY</td>
	-->
	<td width="60">����</td>
	<td width="150">����(���)����</td>
	<td width="60">�ŷ���</td>
	<td width="80">����Ʈ</td>
	<td width="100">���ֹ���ȣ</td>
	<td width="60">CSIDX</td>

	<td width="150">�ֹ��α�</td>
	<!--
	<td width="80">�����</td>
	-->
	<td>���</td>
</tr>

<% for i=0 to oCPGData.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCPGData.FItemList(i).FPGgubun %></td>
	<td><%= oCPGData.FItemList(i).FPGuserid %></td>
	<td><%= oCPGData.FItemList(i).GetAppMethodName %></td>
	<!--
	<td><%= oCPGData.FItemList(i).FPGkey %></td>
	<td><%= oCPGData.FItemList(i).FPGCSkey %></td>
	-->
	<td>
		<font color="<%= oCPGData.FItemList(i).GetAppDivCodeColor %>"><%= oCPGData.FItemList(i).GetAppDivCodeName %></font>
	</td>
	<td>
		<% if Not IsNull(oCPGData.FItemList(i).FcancelDate) then %>
			<%= oCPGData.FItemList(i).FcancelDate %>
		<% else %>
			<%= oCPGData.FItemList(i).FappDate %>
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(oCPGData.FItemList(i).FappPrice, 0) %>&nbsp;</td>
	<td>
		<%= oCPGData.FItemList(i).Fsitename %>
	</td>
	<td><%= oCPGData.FItemList(i).Forderserial %></td>
	<td><%= oCPGData.FItemList(i).Fcsasid %></td>

	<td><%= oCPGData.FItemList(i).GetFullLogOrderSerial %></td>

	<!--
	<td><%= Left(oCPGData.FItemList(i).Fregdate, 10) %></td>
	-->
	<td>
		<% if IsNull(oCPGData.FItemList(i).Flogorderserial) then %>
			<input type="button" class="button" value="��Ī" class="csbutton" style="width:30px;" onclick="SubmitMatchPG(<%= oCPGData.FItemList(i).FappPrice %>, '<%= oCPGData.FItemList(i).FPGkey %>', '<%= oCPGData.FItemList(i).FPGCSkey %>');">
		<% end if %>
	</td>
</tr>
<% next %>
</table>

<% if (oCChgPGData.FresultCount > 0) then %>
[��ȯ�ֹ�]
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= oCChgPGData.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCChgPGData.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">PG��</td>
	<td width="80">PG��id</td>
	<td width="55">�������</td>
	<!--
	<td width="270">PG��KEY</td>
	<td width="270">PG��CSKEY</td>
	-->
	<td width="60">����</td>
	<td width="150">����(���)����</td>
	<td width="60">�ŷ���</td>
	<td width="80">����Ʈ</td>
	<td width="100">���ֹ���ȣ</td>
	<td width="60">CSIDX</td>
	<!--
	<td width="80">�����</td>
	-->
	<td>���</td>
</tr>

<% for i=0 to oCChgPGData.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCChgPGData.FItemList(i).FPGgubun %></td>
	<td><%= oCChgPGData.FItemList(i).FPGuserid %></td>
	<td><%= oCChgPGData.FItemList(i).GetAppMethodName %></td>
	<!--
	<td><%= oCChgPGData.FItemList(i).FPGkey %></td>
	<td><%= oCChgPGData.FItemList(i).FPGCSkey %></td>
	-->
	<td>
		<font color="<%= oCChgPGData.FItemList(i).GetAppDivCodeColor %>"><%= oCChgPGData.FItemList(i).GetAppDivCodeName %></font>
	</td>
	<td>
		<% if Not IsNull(oCChgPGData.FItemList(i).FcancelDate) then %>
			<%= oCChgPGData.FItemList(i).FcancelDate %>
		<% else %>
			<%= oCChgPGData.FItemList(i).FappDate %>
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(oCChgPGData.FItemList(i).FappPrice, 0) %>&nbsp;</td>
	<td>
		<%= oCChgPGData.FItemList(i).Fsitename %>
	</td>
	<td><%= oCChgPGData.FItemList(i).Forderserial %></td>
	<td><%= oCChgPGData.FItemList(i).Fcsasid %></td>
	<!--
	<td><%= Left(oCChgPGData.FItemList(i).Fregdate, 10) %></td>
	-->
	<td>
		<input type="button" class="button" value="��Ī" class="csbutton" style="width:30px;" onclick="SubmitMatchPG(<%= oCChgPGData.FItemList(i).FappPrice %>, '<%= oCChgPGData.FItemList(i).FPGkey %>', '<%= oCChgPGData.FItemList(i).FPGCSkey %>');">
	</td>
</tr>
<% next %>
</table>
<% end if %>

<% if (oCPGDataACA.FresultCount > 0) then %>
&nbsp;<br>
[�ΰŽ���ǰ]
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= oCPGDataACA.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCPGDataACA.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">PG��</td>
	<td width="80">PG��id</td>
	<td width="55">�������</td>
	<!--
	<td width="270">PG��KEY</td>
	<td width="270">PG��CSKEY</td>
	-->
	<td width="60">����</td>
	<td width="150">����(���)����</td>
	<td width="60">�ŷ���</td>
	<td width="80">����Ʈ</td>
	<td width="100">���ֹ���ȣ</td>
	<td width="60">CSIDX</td>
	<!--
	<td width="80">�����</td>
	-->
	<td>���</td>
</tr>

<% for i=0 to oCPGDataACA.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCPGDataACA.FItemList(i).FPGgubun %></td>
	<td><%= oCPGDataACA.FItemList(i).FPGuserid %></td>
	<td><%= oCPGDataACA.FItemList(i).GetAppMethodName %></td>
	<!--
	<td><%= oCPGDataACA.FItemList(i).FPGkey %></td>
	<td><%= oCPGDataACA.FItemList(i).FPGCSkey %></td>
	-->
	<td>
		<font color="<%= oCPGDataACA.FItemList(i).GetAppDivCodeColor %>"><%= oCPGDataACA.FItemList(i).GetAppDivCodeName %></font>
	</td>
	<td>
		<% if Not IsNull(oCPGDataACA.FItemList(i).FcancelDate) then %>
			<%= oCPGDataACA.FItemList(i).FcancelDate %>
		<% else %>
			<%= oCPGDataACA.FItemList(i).FappDate %>
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(oCPGDataACA.FItemList(i).FappPrice, 0) %>&nbsp;</td>
	<td>
		<%= oCPGDataACA.FItemList(i).Fsitename %>
	</td>
	<td><%= oCPGDataACA.FItemList(i).Forderserial %></td>
	<td><%= oCPGDataACA.FItemList(i).Fcsasid %></td>
	<!--
	<td><%= Left(oCPGDataACA.FItemList(i).Fregdate, 10) %></td>
	-->
	<td>
		<input type="button" class="button" value="��Ī" class="csbutton" style="width:30px;" onclick="SubmitMatchPG(<%= oCPGDataACA.FItemList(i).FappPrice %>, '<%= oCPGDataACA.FItemList(i).FPGkey %>', '<%= oCPGDataACA.FItemList(i).FPGCSkey %>');">
	</td>
</tr>
<% next %>
</table>
<% end if %>

<% if (oCPGDataACAbyHand.FresultCount > 0) then %>
&nbsp;<br>
[�ΰŽ� ���� ī������]
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= oCPGDataACAbyHand.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCPGDataACAbyHand.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">PG��</td>
	<td width="80">PG��Key</td>
	<td width="55">ī�屸��</td>
	<!--
	<td width="270">PG��KEY</td>
	<td width="270">PG��CSKEY</td>
	-->
	<td width="60">����</td>
	<td width="150">����(���)����</td>
	<td width="60">�ŷ���</td>
	<td width="100">�ֹ���ȣ</td>
	<td width="80">�Աݿ�����</td>
	<!--
	<td width="80">�����</td>
	-->
	<td>���</td>
</tr>

<% for i=0 to oCPGDataACAbyHand.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCPGDataACAbyHand.FItemList(i).FPGgubun %></td>
	<td><%= oCPGDataACAbyHand.FItemList(i).FPGkey %></td>
	<td><%= oCPGDataACAbyHand.FItemList(i).FcardGubun %></td>
	<!--
	<td><%= oCPGDataACAbyHand.FItemList(i).FPGkey %></td>
	<td><%= oCPGDataACAbyHand.FItemList(i).FPGCSkey %></td>
	-->
	<td>
		<font color="<%= oCPGDataACAbyHand.FItemList(i).GetAppDivCodeColor %>"><%= oCPGDataACAbyHand.FItemList(i).GetAppDivCodeName %></font>
	</td>
	<td>
		<%= oCPGDataACAbyHand.FItemList(i).FappDate %>
	</td>
	<td align="right"><%= FormatNumber(oCPGDataACAbyHand.FItemList(i).FcardPrice, 0) %>&nbsp;</td>
	<td><%= oCPGDataACAbyHand.FItemList(i).Forderserial %></td>
	<td><%= oCPGDataACAbyHand.FItemList(i).Fipkumdate %></td>
	<!--
	<td><%= Left(oCPGDataACAbyHand.FItemList(i).Fregdate, 10) %></td>
	-->
	<td>
		<input type="button" class="button" value="��Ī" class="csbutton" style="width:30px;" onclick="SubmitMatchPGOff('<%= oCPGDataACAbyHand.FItemList(i).FPGGubun %>', '<%= oCPGDataACAbyHand.FItemList(i).FPGkey %>');">
	</td>
</tr>
<% next %>
</table>
<% end if %>

<form name="frmAct" method="post" action="refundMatchRefund_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<input type="hidden" name="suborderserial" value="<%= suborderserial %>">
	<input type="hidden" name="orgorderserial" value="<%= orgorderserial %>">
	<input type="hidden" name="asid" value="">
	<input type="hidden" name="pggubun" value="">
	<input type="hidden" name="pgkey" value="">
	<input type="hidden" name="pgcskey" value="">
	<input type="hidden" name="appprice" value="">
	<input type="hidden" name="appDate" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
