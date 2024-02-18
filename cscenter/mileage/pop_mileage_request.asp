<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs���� ���ϸ��� ����
' Hieditor : �̻� ����
'			 2023.09.05 �ѿ�� ����(�ҽ� ǥ���ڵ����� ����. �������� Ŭ�� ���� �߰�.)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<%
dim userid, orderserial, mileage, jukyo, i, buf, gubun01, gubun02, gubun01name, gubun02name, page, ojumun, defaultCSRefundLimit
dim ix,iy, omakeridList
	userid = requestCheckvar(request("userid"),32)
	orderserial = requestCheckvar(request("orderserial"),32)
	mileage = requestCheckvar(request("mileage"),32)
	jukyo = requestCheckvar(request("jukyo"),32)
	gubun01 = requestCheckvar(request("gubun01"),32)
	gubun02 = requestCheckvar(request("gubun02"),32)
	gubun01name = requestCheckvar(request("gubun01name"),32)
	gubun02name = requestCheckvar(request("gubun02name"),32)

if (userid = "") then
    response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
    dbget.close()	:	response.End
end if

page = 1

set ojumun = new COrderMaster
ojumun.FPageSize = 5
ojumun.FCurrPage = page
ojumun.FRectUserID = userid
ojumun.FRectOrderSerial = orderserial
ojumun.QuickSearchOrderList

'' ���� 6���� ���� ���� �˻�
if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderList

    if (ojumun.FResultCount>0) then
        response.write "<script>alert('6���� ���� �ֹ��Դϴ�.');</script>"
    end if
end if

if (orderserial = "") and (ojumun.FResultCount=1) then
	orderserial = ojumun.FItemList(0).FOrderSerial
end if

set omakeridList = new COrderMaster
if (orderserial <> "") then
	omakeridList.FRectOrderSerial = orderserial
	omakeridList.getUpcheBeasongMakerList
end if

defaultCSRefundLimit = GetUserRefundAuthLimit(session("ssBctId"))

%>
<script type="text/javascript">

function SubmitForm(){
	var jukyo;

	if (document.frm.orderserial.value == "") {
		alert("���� �ֹ���ȣ�� �Է��ϼ���.");
		return;
	}

	if (document.frm.gubun01.value == "") {
		alert("���������� �����ϼ���.");
		return;
	}

	if (document.frm.mileage.value == "") {
		alert("�������� ��Ȯ�� �Է��ϼ���.");
		return;
	}

	if (document.frm.mileage.value*0 != 0) {
		alert("�������� ��Ȯ�� �Է��ϼ���.");
		return;
	}

	if (document.frm.mileage.value == 0) {
		alert("�������� 0 �� �� �� �����ϴ�.");
		return;
	}

	if (document.frm.mileage.value*1 > <%= defaultCSRefundLimit %>) {
		alert("<%= FormatNumber(defaultCSRefundLimit,0) %> ���ϸ����� �ʰ��Ͽ� ������ �� �����ϴ�.\n�ε��� �� ���� �������� �ο��ؾ� �� ��� ��Ʈ�忡�� �������ּ���.");
		return;
	}

	<% if omakeridList.FresultCount > 0 then %>
	if (((document.frm.gubun01.value != "C004") || (document.frm.gubun02.value != "CD13")) && (document.frm.requiremakerid.value == "")) {
		alert("���� �귣�带 �����ϼ���.");
		return;
	}
	<% end if %>

	if (document.frm.contents_jupsu.value == "") {
		alert("���������� �����ϴ�.");
		return;
	}

	if (confirm("������û �Ͻðڽ��ϱ�?") == true) {
		document.frm.submit();
	}
}

function SubmitFormForce() {
	if (document.frm.gubun01.value == "") {
		alert("���������� �����ϼ���.");
		return;
	}

	if (document.frm.mileage.value == "") {
		alert("�������� ��Ȯ�� �Է��ϼ���.");
		return;
	}

	if (document.frm.mileage.value*0 != 0) {
		alert("�������� ��Ȯ�� �Է��ϼ���.");
		return;
	}

	if (document.frm.mileage.value == 0) {
		alert("�������� 0 �� �� �� �����ϴ�.");
		return;
	}

	if (confirm("���� �Ͻðڽ��ϱ�?") == true) {
		document.frm.mode.value = "requestForce";
		document.frm.submit();
	}
}

function SetOrderSerial(orderserial){
	var frm = document.frm;
	var DisplayMakerID = false;
	frm.orderserial.value = orderserial;

	if (frm.contents_jupsu.value != "") {
		if (confirm("���ú귣�带 ǥ���Ͻðڽ��ϱ�?\n(�Էµ� ���������� ������ϴ�.)")) {
			DisplayMakerID = true;
		}
	} else {
		DisplayMakerID = true;
	}

	if (DisplayMakerID == true) {
		frm.method = "get";
		frm.action = "";
		frm.submit();
	}
}

function selectGubun(value_gubun01, value_gubun02, value_gubun01name, value_gubun02name, name_gubun01, name_gubun02, name_gubun01name, name_gubun02name ,name_frm, targetDiv) {
    var frm = eval(name_frm);

    eval("document." + name_frm + "." + name_gubun01).value = value_gubun01;
    eval("document." + name_frm + "." + name_gubun02).value = value_gubun02;
    eval("document." + name_frm + "." + name_gubun01name).value = value_gubun01name;
    eval("document." + name_frm + "." + name_gubun02name).value = value_gubun02name;
}

function regcontents_jupsu(contents_jupsu) {
	frm.contents_jupsu.value=contents_jupsu;
}

</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<b>���ϸ��� ������û</b> ������û�� �Ͻø�, CSó������Ʈ�� ��ϵǸ�, ������ ���ΰ� �Բ� �����˴ϴ�.
		<br>* <font color=red><%= FormatNumber(defaultCSRefundLimit,0) %> ���ϸ���</font> �ʰ� �����Ұ�
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- �׼� �� -->

<form name="frm" method="post" action="/cscenter/mileage/domodifymileage.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="request">
<input type="hidden" name="userid" value="<%= userid %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="left">
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">���̵� :</td>
  		<td bgcolor="#FFFFFF" width="40%" >
  			<b><%= userid %></b>
  		</td>
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">�����ֹ���ȣ :</td>
  		<td bgcolor="#FFFFFF" width="40%"  >
  			<input type=text name=orderserial value="<%= orderserial %>">
  		</td>
	</tr>
	<tr align="left">
		<td height="30" bgcolor="<%= adminColor("tabletop") %>">�������� :</td>
  		<td bgcolor="#FFFFFF" colspan="3">
			<input type="hidden" name="gubun01" value="<%= gubun01 %>">
			<input type="hidden" name="gubun02" value="<%= gubun02 %>">
			<input class="text_ro" type="text" name="gubun01name" value="<%= gubun01name %>" size="16" Readonly >
			&gt;
			<input class="text_ro" type="text" name="gubun02name" value="<%= gubun02name %>" size="16" Readonly >
			&nbsp;
			[<a href="javascript:selectGubun('C006','CF06','��������','�������','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">�������</a>]
			[<a href="javascript:selectGubun('C004','CD05','����','ǰ��','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">ǰ��</a>]
			[<a href="javascript:selectGubun('C006','CF01','��������','���߼�','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">�����</a>]
			[<a href="javascript:selectGubun('C005','CE03','��ǰ����','��ǰ��Ͽ���','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">��ǰ��Ͽ���</a>]
			[<a href="javascript:selectGubun('C004','CD12','����','��ü����ҷ�','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">��ü����ҷ�</a>]
			[<a href="javascript:selectGubun('C004','CD14','����','��Ÿ��ü����','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">��Ÿ��ü����</a>]
			&nbsp;
			[<a href="javascript:selectGubun('C004','CD13','����','CS����','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">CS����</a>]
			<!--
			[<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">��Ÿ</a>]
			-->
  		</td>
	</tr>
	<tr align="left">
  		<td height="30" bgcolor="<%= adminColor("tabletop") %>">������ :</td>
  		<td bgcolor="#FFFFFF" >
			<input type=text name=mileage value="<%= mileage %>">
  		</td>
  		<td height="30" bgcolor="<%= adminColor("tabletop") %>">���ú귣�� :</td>
  		<td bgcolor="#FFFFFF" >
			<select class="select" name="requiremakerid">
				<option></option>
				<% if orderserial <> "" and omakeridList.FResultCount > 0 then %>
				<% for i=0 to omakeridList.FResultCount - 1 %>
				<option value="<%= omakeridList.FItemList(i).Fmakerid %>"><%= CHKIIF(omakeridList.FItemList(i).Fmakerid="10x10logistics", "�ٹ����ٹ��", omakeridList.FItemList(i).Fmakerid) %></option>
				<% next %>
				<% end if %>
			</select>
		</td>
	</tr>
	<tr align="left">
  		<td height="30" bgcolor="<%= adminColor("tabletop") %>">�������� :</td>
  		<td bgcolor="#FFFFFF" colspan="3">
			<textarea class='textarea' id="contents_jupsu" name="contents_jupsu" cols="80" rows="6"></textarea>
			<% if C_ADMIN_AUTH then %>
				<br>
				[<a href="#" onclick="regcontents_jupsu('�̺�Ʈ ���ϸ��� ����'); return;">�̺�Ʈ���ϸ�������</a>]
				[<a href="#" onclick="regcontents_jupsu('�ֹ� ���ϸ��� ����'); return;">�ֹ����ϸ�������</a>]
			<% end if %>
  		</td>
	</tr>
</table>
</form>

<% if orderserial = "" then %>
	<br />
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td height=25>����</td>
		<td>�ֹ���ȣ</td>
		<td>������</td>
		<td>�����Ѿ�</td>
		<td>�������</td>
		<td>�ŷ�����</td>
		<td>�ֹ���</td>
		<td>���</td>
	</tr>

	<% if (ojumun.FresultCount > 0) then %>
		<% for i=0 to ojumun.FResultCount - 1 %>
		<tr align="center" bgcolor="#FFFFFF">
			<td height=25><%= ojumun.FItemList(ix).CancelYnName %></td>
			<td><%= ojumun.FItemList(i).FOrderSerial %></td>
			<td><%= ojumun.FItemList(i).FBuyName %></td>
			<td><%= FormatNumber(ojumun.FItemList(i).FTotalSum,0) %></td>
			<td><%= ojumun.FItemList(i).JumunMethodName %></td>
			<td><%= ojumun.FItemList(i).IpkumDivName %></td>
			<td><acronym title="<%= ojumun.FItemList(i).FRegDate %>"><%= Left(ojumun.FItemList(i).FRegDate,10) %></acronym></td>
			<td><input type=button class="button" value="����" onclick="SetOrderSerial('<%= ojumun.FItemList(i).FOrderSerial %>')"></td>
		</tr>
		<% next %>
	<% else %>
		<tr align="center" bgcolor="#FFFFFF">
			<td colspan="8"> �˻��� ����� �����ϴ�.</td>
		</tr>
	<% end if %>
	</table>
<% end if %>

<!-- �׼� ���� -->
<br />
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="center">
		<input type="button" class="button" value="������û" onClick="SubmitForm();">

		<% if C_CSPowerUser or C_ADMIN_AUTH then %>
			&nbsp;
			<input type="button" class="button" value="����(������)" onClick="SubmitFormForce();"> ��������Ʈ���̻�, �����ڱ���
		<% end if %>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
