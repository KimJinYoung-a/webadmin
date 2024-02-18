<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%
dim orderserial
dim mode
Dim sqlStr, rowCount

orderserial = request("orderserial")

if (orderserial = "") then
    orderserial = "-"
end if


'// ============================================================================
'// CS ����
'// ============================================================================
dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FPageSize = 50
ocsaslist.FCurrPage = 1

ocsaslist.FRectOrderSerial = orderserial

ocsaslist.GetCSASMasterListByProcedure


'// ============================================================================
'// ���޻� �ֹ�����
'// ============================================================================

Dim ArrExtSiteOrderDetailList

sqlStr = " select "
sqlStr = sqlStr & " 	T.sellSite "
sqlStr = sqlStr & " 	, T.orderserial "
sqlStr = sqlStr & " 	, T.outmallorderserial "
sqlStr = sqlStr & " 	, T.OrgDetailKey  "
sqlStr = sqlStr & " 	, T.matchitemid "
sqlStr = sqlStr & " 	, T.matchitemoption "
sqlStr = sqlStr & " 	, T.orderitemname "
sqlStr = sqlStr & " 	, T.orderitemoptionname "
sqlStr = sqlStr & " 	, T.itemordercount "
sqlStr = sqlStr & " 	, d.itemno "
sqlStr = sqlStr & " 	, m.cancelyn "
sqlStr = sqlStr & " 	, d.cancelyn "
sqlStr = sqlStr & " 	, D.beasongdate "
sqlStr = sqlStr & " 	, IsNull(D.currstate, 0) as currstate "
sqlStr = sqlStr & " 	, IsNULL(T.sendState,0) as sendState "
sqlStr = sqlStr & " 	, T.matchState as matchState "				'// 15
sqlStr = sqlStr & " 	, T.outmallorderseq "
sqlStr = sqlStr & " from "
sqlStr = sqlStr & " 	db_temp.dbo.tbl_xSite_TMPOrder T "
sqlStr = sqlStr & "  	left Join db_order.dbo.tbl_order_detail D "
sqlStr = sqlStr & "  	on "
sqlStr = sqlStr & " 		T.orderserial=D.orderserial "
sqlStr = sqlStr & "  		and T.matchitemid=D.itemid "
sqlStr = sqlStr & "  		and T.matchitemoption=D.itemoption "
sqlStr = sqlStr & "  	left Join db_order.dbo.tbl_order_master M "
sqlStr = sqlStr & "  	on "
sqlStr = sqlStr & " 		D.orderserial=M.orderserial "
sqlStr = sqlStr & " where "
sqlStr = sqlStr & " 	1 = 1 "
sqlStr = sqlStr & " 	and T.orderserial = '" + CStr(orderserial) + "' "
sqlStr = sqlStr & " order by "
sqlStr = sqlStr & " 	T.orderserial, T.matchitemid, T.matchitemoption  "
''response.write sqlStr

rsget.Open sqlStr,dbget,1
if Not rsget.Eof then
	ArrExtSiteOrderDetailList = rsget.getRows
end if
rsget.Close


'// ============================================================================
dim omasterwithcs
set omasterwithcs = new COldMiSend
omasterwithcs.FRectOrderSerial = orderserial
omasterwithcs.GetOneOrderMasterWithCS

orderserial = omasterwithcs.FOneItem.FOrderSerial

dim omisendList
set omisendList = new COldMiSend
omisendList.FRectOrderSerial = orderserial
omisendList.GetMiSendOrderDetailList


dim validtenitemno
validtenitemno = 0

for i = 0 to omisendList.FResultCount - 1
	if (omisendList.FItemList(i).FDetailCancelYn <> "Y") then
		validtenitemno = validtenitemno + omisendList.FItemList(i).FItemNo
	end if
next

dim i
dim prevmatchitemid, prevmatchitemoption, IsCSDetail
%>
<script language='javascript'>
function SetCancelAllOrder() {
	var frm = document.frmact;
    if (confirm("����ó�� �Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "cancelall";
        frm.submit();
    }
}

function SetCancelSelectedOrder(isforce) {
	var arrchk = "";
	var validextitemno = 0;
	var validtenitemno = <%= validtenitemno %>;
	for (var i = 0; ; i++) {
		var chk = document.getElementById("chk_" + i);
		var extitemno = document.getElementById("extitemno_" + i);
		if (chk == undefined) {
			break;
		}

		if (chk.checked == true) {
			arrchk = arrchk + "," + chk.value;
		} else {
			validextitemno = validextitemno + extitemno.value*1;
		}
	}

	if (arrchk == "") {
		alert("���õ� �ֹ��� �����ϴ�.");
		return;
	}

	if (isforce != true) {
		if (validextitemno != validtenitemno) {
			if (validextitemno > validtenitemno) {
				alert("�����ֹ�����(" + validextitemno + ")�� �ٹ����� �ֹ�����(" + validtenitemno + ")�� ��ġ���� �ʽ��ϴ�.\n\n��ǰ�� ��ҵ� ��� ���޸� �ֹ������� �����ϼ���");
			} else {
				alert("�����ֹ�����(" + validextitemno + ")�� �ٹ����� �ֹ�����(" + validtenitemno + ")�� ��ġ���� �ʽ��ϴ�.\n\n��ǰ�� ������ ��� �����ǰ�� �����ϼ���.");
			}

			return;
		}
	}

	var frm = document.frmact;
	if (confirm("���� �����ֹ��� �����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "cancelselected";
		frm.arrchk.value = arrchk;
		frm.submit();
	}
}

function ModifyMatchItem(outmallorderseq, extitemid, extitemoption, extitemno) {
	// =========================================================================
	// ���� ����
	// =========================================================================
	var validextitemno = extitemno;
	var validtenitemno = 0;
	for (var i = 0; ; i++) {
		var chk = document.getElementById("chkten_" + i);
		var itemno = document.getElementById("tenitemno_" + i);
		if (chk == undefined) {
			break;
		}

		if (chk.checked) {
			validtenitemno = validtenitemno + itemno.value*1;
		}
	}

	if (validextitemno != validtenitemno) {
		if (validextitemno > validtenitemno) {
			alert("�����ֹ�����(" + validextitemno + ")�� �ٹ����� �ֹ�����(" + validtenitemno + ")�� ��ġ���� �ʽ��ϴ�.\n\n��ǰ�� ��ҵ� ��� ���޸� �ֹ������� �����ϼ���");
		} else {
			alert("�����ֹ�����(" + validextitemno + ")�� �ٹ����� �ֹ�����(" + validtenitemno + ")�� ��ġ���� �ʽ��ϴ�.\n\n��ǰ�� ������ ��� �����ǰ�� �����ϼ���.");
		}

		return;
	}

	// =========================================================================
	// ���û�ǰ Ȯ��
	// =========================================================================
	var selecteditemid = 0;
	var selecteditemoption = "";
	var selecteditemno = 0;
	var selecteditemcount = 0;

	for (var i = 0; ; i++) {
		var chkten = document.getElementById("chkten_" + i);
		var tenitemid = document.getElementById("tenitemid_" + i);
		var tenitemoption = document.getElementById("tenitemoption_" + i);
		var tenitemno = document.getElementById("tenitemno_" + i);

		if (chkten == undefined) {
			break;
		}

		if (chkten.checked == true) {
			if (selecteditemcount < 1) {
				selecteditemid = tenitemid.value;
				selecteditemoption = tenitemoption.value;
			}
			selecteditemno = selecteditemno + tenitemno.value*1;
			selecteditemcount = selecteditemcount + 1;
		}
	}

	if (selecteditemcount < 1) {
		alert("�ٹ����� �ֹ��������� ����� ��ǰ�� �����ϼ���.");
		return;
	}

	/*
	// �ΰ��� �̻��� ��ǰ���� ������ �ֹ������� �����Ѱ��
	// ����� ��ǰ ��θ� üũ�ؼ� ������ ���ϰ�, ������ ������ ù��° ��ǰ�� �����ȣ ����
	if (selecteditemcount > 1) {
		alert("������ ��ǰ�� �ϳ��� ���� �����մϴ�.");
		return;
	}
	*/

	if (selecteditemno*1 != extitemno*1) {
		if (confirm("����!!\n\n������ �ٸ��ϴ�.\n������ �����Է� �Ͻðڽ��ϱ�?") != true) {
			return;
		}
		// alert("����Ұ� - ������ �ٸ��ϴ�." + selecteditemno);
		// return;
	}

	// =========================================================================
	// ���� ��ǰ�� �̹� ���޸� �ֹ� ��ǰ�� �ִ��� Ȯ��
	// =========================================================================
	for (var i = 0; ; i++) {
		var chk = document.getElementById("chk_" + i);
		var itemid = document.getElementById("extitemid_" + i);
		var itemoption = document.getElementById("extitemoption_" + i);
		if (chk == undefined) {
			break;
		}

		if ((itemid.value*1 == selecteditemid*1) && (itemoption.value*1 == selecteditemoption*1)) {
			alert("�ߺ����� : ������ ��ǰ�� �̹� ���޸� �ֹ������� �ֽ��ϴ�.");
			return;
		}
	}

	var frm = document.frmact;
	if (confirm("���� ��ǰ���� �����ǰ�� �����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "modifymatchitem";
		frm.chk.value = outmallorderseq;
		frm.itemid.value = selecteditemid;
		frm.itemoption.value = selecteditemoption;
		frm.submit();
	}
}

function ModifyMatchItemNo(outmallorderseq, extitemid, extitemoption, extitemno) {
	// =========================================================================
	var validextitemno = 0;
	var validtenitemno = <%= validtenitemno %>;
	for (var i = 0; ; i++) {
		var chk = document.getElementById("chk_" + i);
		var itemno = document.getElementById("extitemno_" + i);
		if (chk == undefined) {
			break;
		}

		validextitemno = validextitemno + itemno.value*1;
	}

	if (validextitemno != validtenitemno) {
		if (validextitemno > validtenitemno) {
			alert("�����ֹ�����(" + validextitemno + ")�� �ٹ����� �ֹ�����(" + validtenitemno + ")�� ��ġ���� �ʽ��ϴ�.\n\n��ǰ�� ��ҵ� ��� ���޸� �ֹ������� �����ϼ���");
		} else {
			alert("�����ֹ�����(" + validextitemno + ")�� �ٹ����� �ֹ�����(" + validtenitemno + ")�� ��ġ���� �ʽ��ϴ�.\n\n��ǰ�� ������ ��� �����ǰ�� �����ϼ���.");
		}

		return;
	}

	// =========================================================================
	var changeditemid = 0;
	var changeditemoption = "";
	for (var i = 0; ; i++) {
		var chk = document.getElementById("chk_" + i);
		var extchangeitemid = document.getElementById("extitemid_" + i);
		var extchangeitemoption = document.getElementById("extitemoption_" + i);
		var extchangeitemno = document.getElementById("extitemno_" + i);
		var extchangetenitemno = document.getElementById("exttenitemno_" + i);

		if (chk == undefined) {
			break;
		}

		if ((extchangeitemno.value*1 + extitemno*1) == extchangetenitemno.value*1) {
			changeditemid = extchangeitemid.value;
			changeditemoption = extchangeitemoption.value;
			break;
		}
	}

	if (changeditemid == 0) {
		alert("����Ұ� : �ý����� ����");
		return;
	}

	var frm = document.frmact;
	if (confirm("������ �þ ��ǰ���� �����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "modifymatchitemno";
		frm.chk.value = outmallorderseq;
		frm.itemid.value = changeditemid;
		frm.itemoption.value = changeditemoption;
		frm.submit();
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
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="SearchThis();">
		</td>
	</tr>
	</form>
</table>

<p>

<br><b>[CSó������]</b>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height="25" width="80">idx</td>
		<td width="120">����</td>
		<td width="100">���ֹ���ȣ</td>
		<td width="100">Site</td>
		<td>����</td>
		<td width="35">����</td>
		<td width="80">�����</td>
		<td width="80">ó����</td>
		<td width="220">���ü���</td>
		<td width="30">����</td>
	</tr>
<% for i = 0 to (ocsaslist.FResultCount - 1) %>
    <% if (ocsaslist.FItemList(i).Fdeleteyn <> "N") then %>
    <tr bgcolor="#EEEEEE" style="color:gray" align="center">
    <% else %>
    <tr bgcolor="#FFFFFF" align="center">
    <% end if %>
        <td height="25" nowrap><%= ocsaslist.FItemList(i).Fid %></td>
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
        <td nowrap><acronym title="<%= ocsaslist.FItemList(i).Fregdate %>"><%= Left(ocsaslist.FItemList(i).Fregdate,10) %></acronym></td>
        <td nowrap><acronym title="<%= ocsaslist.FItemList(i).Ffinishdate %>"><%= Left(ocsaslist.FItemList(i).Ffinishdate,10) %></acronym></td>
		<td nowrap>
			<% Call drawSelectBoxDeliverCompany ("songjangdiv", ocsaslist.FItemList(i).Fsongjangdiv) %>
			<%= ocsaslist.FItemList(i).Fsongjangno %>
		</td>
        <td nowrap>
			<% if ocsaslist.FItemList(i).Fdeleteyn="Y" then %>
			<font color="red">����</font>
			<% elseif ocsaslist.FItemList(i).Fdeleteyn="C" then %>
			<font color="red"><strong>���</strong></font>
			<% end if %>
        </td>
    </tr>
<% next %>

</table>

<p>

<br><b>[�ٹ������ֹ�����]</b>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<form name="frmtensite">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"></td>
		<td width="50">��ǰ�ڵ�</td>
		<td width="50">�̹���</td>
		<td>��ǰ��<font color="blue">[�ɼ�]</font></td>
		<td width="40">�ֹ�<br>����</td>
		<td width="35">���<br>����</td>
		<td width="60">�������</td>
		<td width="80">��������</td>
		<td width="120">�����ȣ</td>
	</tr>
	<% for i=0 to omisendList.FResultCount -1 %>
	<% if omisendList.FItemList(i).FItemLackNo<>0 then %>
	<tr align="center" bgcolor="<%= adminColor("pink") %>">
	<% else %>
	<tr align="center" bgcolor="FFFFFF">
	<% end if %>
		<td><input type="checkbox" name="chkten" id="chkten_<%= i %>" value="" <% if omisendList.FItemList(i).FDetailCancelYn = "Y" then %>disabled<% end if %> ></td>
		<input type="hidden" name="tenitemid" id="tenitemid_<%= i %>" value="<%= omisendList.FItemList(i).FItemID %>">
		<input type="hidden" name="tenitemoption" id="tenitemoption_<%= i %>" value="<%= omisendList.FItemList(i).FItemOption %>">
		<input type="hidden" name="tenitemno" id="tenitemno_<%= i %>" value="<%= omisendList.FItemList(i).FItemNo %>">
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
		    <%= fnColor(omisendList.FItemList(i).FDetailCancelYn,"cancelyn") %>
		</td>
		<td>
		    <font color="<%= omisendList.FItemList(i).getUpCheDeliverStateColor %>"><%= omisendList.FItemList(i).getUpCheDeliverStateName %></font>
		</td>
		<td>
				<font color="<%= omisendList.FItemList(i).getMiSendCodeColor %>"><%= omisendList.FItemList(i).getMiSendCodeName %></font>
		</td>
		<td>
			<% if (omisendList.FItemList(i).FSongjangno <> "") then %>
				<%= omisendList.FItemList(i).FSongjangdiv %>
				<%= omisendList.FItemList(i).FSongjangno %>
			<% end if %>
		</td>
	</tr>
	<% next %>
	</form>
</table>

<p>

<br><b>[���޸��ֹ�����]</b>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<form name="frmextsite">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="20"></td>
		<td width="30">�ֹ�<br>����</td>
		<td width="70">���޻�</td>
    	<td>�����ֹ���ȣ</td>
      	<td>����<br>��ǰ�ڵ�</td>
      	<td width="50">��ǰ�ڵ�</td>
      	<td width="50">�ɼ��ڵ�</td>
      	<td align="left">��ǰ��<br><font color="blue">[�ɼǸ�]</font></td>
        <td width="30">����<br>�ֹ�<br>����</td>
		<td width="30">����<br>�ֹ�<br>����</td>
		<td width="30">����<br>���<br>����</td>
		<td width="30">���<br>����</td>
		<td width="30">����<br>����</td>
      	<td>���</td>
    </tr>
<% if (IsArray(ArrExtSiteOrderDetailList)) THEN %>
<%
rowCount = UBound(ArrExtSiteOrderDetailList,2)

for i=0 to rowCount
	'// ��ǰ�ڵ�, �ɼ��ڵ� ��� �����ϸ� ���޸� CS���̴�.
	IsCSDetail = (prevmatchitemid = ArrExtSiteOrderDetailList(4,i)) and (prevmatchitemoption = ArrExtSiteOrderDetailList(5,i))
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><input type="checkbox" name="chk" id="chk_<%= i %>" value="<%= ArrExtSiteOrderDetailList(16,i) %>" <% if ArrExtSiteOrderDetailList(14,i) = "1" or ArrExtSiteOrderDetailList(13,i) = "7" and Not IsCSDetail then %>disabled<% end if %> ></td>
	<input type="hidden" name="extitemid" id="extitemid_<%= i %>" value="<%= ArrExtSiteOrderDetailList(4,i) %>">
	<input type="hidden" name="extitemoption" id="extitemoption_<%= i %>" value="<%= ArrExtSiteOrderDetailList(5,i) %>">
	<input type="hidden" name="extitemno" id="extitemno_<%= i %>" value="<%= ArrExtSiteOrderDetailList(8,i) %>">
	<input type="hidden" name="exttenitemno" id="exttenitemno_<%= i %>" value="<%= ArrExtSiteOrderDetailList(9,i) %>">
    <td height="45">
		<% if IsCSDetail then %>
			<font color=red>CS</font>
		<% else %>
			����
		<% end if %>
	</td>
	<td><%= ArrExtSiteOrderDetailList(0,i) %></td>
    <td><%= ArrExtSiteOrderDetailList(2,i) %></td>
    <td><%= ArrExtSiteOrderDetailList(3,i) %></td>
    <td><%= ArrExtSiteOrderDetailList(4,i) %></td>
	<td><%= ArrExtSiteOrderDetailList(5,i) %></td>
    <td align="left">
		<%= ArrExtSiteOrderDetailList(6,i) %>
		<% if (ArrExtSiteOrderDetailList(5,i) <> "0000") then %>
			<br><font color="blue">[<%= ArrExtSiteOrderDetailList(7,i) %>]</font>
		<% end if %>
	</td>
    <td align="center"><%= ArrExtSiteOrderDetailList(8,i) %></td>
    <td align="center">
		<% if (ArrExtSiteOrderDetailList(8,i) <> ArrExtSiteOrderDetailList(9,i)) then %><font color="red"><% end if %>
		<% if (ArrExtSiteOrderDetailList(11,i) <> "Y") then %>
			<%= ArrExtSiteOrderDetailList(9,i) %>
		<% end if %>
	</td>
	<td>
		<% if ArrExtSiteOrderDetailList(15,i) = "D" then %>
		���
		<% end if %>
	</td>
	<td align="center">
		<% if ArrExtSiteOrderDetailList(13,i) = "7" then %>
			���<br>�Ϸ�
		<% end if %>
	</td>
	<td align="center">
		<% if ArrExtSiteOrderDetailList(14,i) = "1" then %>
			Y
		<% end if %>
	</td>
    <td align="center">
		<input type="button" class="csbutton" value="��ǰ����" onClick="ModifyMatchItem(<%= ArrExtSiteOrderDetailList(16,i) %>, <%= ArrExtSiteOrderDetailList(4,i) %>, '<%= ArrExtSiteOrderDetailList(5,i) %>', <%= ArrExtSiteOrderDetailList(8,i) %>)" <% if ArrExtSiteOrderDetailList(14,i) = "1" then %>disabled<% end if %> >
		<input type="button" class="csbutton" value="�����߰�" onClick="ModifyMatchItemNo(<%= ArrExtSiteOrderDetailList(16,i) %>, <%= ArrExtSiteOrderDetailList(4,i) %>, '<%= ArrExtSiteOrderDetailList(5,i) %>', <%= ArrExtSiteOrderDetailList(8,i) %>)" <% if ArrExtSiteOrderDetailList(14,i) = "1" then %>disabled<% end if %> >
	</td>
</tr>
<%
prevmatchitemid = ArrExtSiteOrderDetailList(4,i)
prevmatchitemoption = ArrExtSiteOrderDetailList(5,i)
%>
<% next %>
<% ELSE %>
<tr>
    <td colspan="11" align="center">[�˻� ����� �����ϴ�.]</td>
</tr>
<% end if %>
</form>
</table>

<p>

<input type="button" class="csbutton" value="���� ���޸� �ֹ�����" onClick="SetCancelSelectedOrder(false)">
&nbsp;
<input type="button" class="csbutton" value="[��������] ���� ���޸� �ֹ�����" onClick="SetCancelSelectedOrder(true)">
&nbsp;
<input type="button" class="csbutton" value="���޸��ֹ� ��ü����" onClick="SetCancelAllOrder()" <%if (omasterwithcs.FOneItem.FCancelyn <> "Y") then %>disabled<% end if %> >

<p>

<form name="frmact" method="post" action="etcSiteOrderProc.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="orderserial" value="<%= orderserial %>">
<input type="hidden" name="arrchk" value="">
<input type="hidden" name="chk" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
</form>
<!-- ǥ �ϴܹ� ��-->


<%
set omasterwithcs = Nothing
set omisendList = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
