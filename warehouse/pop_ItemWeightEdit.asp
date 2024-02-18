<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �ؿܹ�� ��ǰ ����
' History : 2008.03.26 ������ ����
'			2016.05.27 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/items/itemVolumncls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim itembarcode, prcAfter, sqlStr, mode, deliverOverseas, menupos, itemgubun, itemid, itemoption, i, regtype
dim optionsetYN

	itembarcode = requestCheckVar(request("itembarcode"),20)
	prcAfter = requestCheckVar(request("prcAfter"),32)
	mode = requestCheckVar(request("mode"),32)
	menupos		= requestCheckVar(getNumeric(request("menupos")),10)
	regtype = requestCheckVar(request("regtype"),32)
	optionsetYN=False
'' ByWeightProc/BySizeProc
if (mode = "") then
	mode = "ByWeightProc"
end if

if regtype="" then regtype="I"

'������ڵ� �˻�
if Len(itembarcode)>=12 then
	sqlStr = "select top 1 b.* " + VbCrlf
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_stock b " + VbCrlf
	sqlStr = sqlStr + " where b.barcode='" + CStr(itembarcode) + "' " + VbCrlf

	'response.write sqlStr & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		itemgubun = rsget("itemgubun")
		itemid = rsget("itemid")
		itemoption = rsget("itemoption")
	else
		itemgubun 	= BF_GetItemGubun(itembarcode)
		itemid 		= BF_GetItemId(itembarcode)
		itemoption 	= BF_GetItemOption(itembarcode)
	end if
	rsget.Close
else
	itemgubun="10"
	itemid = itembarcode
	itemoption="0000"
end if

if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim oitembar
set oitembar = new CItemInfo
	oitembar.FRectItemID = itemid
	if itemid<>"" then
		oitembar.GetOneItemInfo
	    if itemgubun = "10" and oitembar.FResultCount>0 then
		    regtype = oitembar.FOneItem.FitemManageType
	    end if
	end if

if itemgubun = "" then
	Response.Write "�����Է� �Ұ� : ��ǰ������ ����"
	Response.end
end if

if itemid<>"" and itemgubun <> "10" then
	oitembar.FRectItemGubun = itemgubun
	oitembar.FRectItemID =  itemid
	oitembar.FRectItemOption =  itemoption
	oitembar.GetOneItemInfoOffline
	regtype = "O"
	''Response.end
end if

dim k, oitemoption, oOptionMultipleType, oOptionMultiple

set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" and itemgubun = "10" then
	oitemoption.GetItemOptionInfo
end If

set oOptionMultipleType = new CItemOptionMultiple
oOptionMultipleType.FRectItemID = itemid
if itemid<>"" and itemgubun = "10" then
    oOptionMultipleType.GetOptionTypeInfo
end if

set oOptionMultiple = new CitemOptionMultiple
oOptionMultiple.FRectItemID = itemid
if itemid<>"" and itemgubun = "10" then
    oOptionMultiple.GetOptionMultipleInfo
end if
%>
<script language="javascript" type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function InputWeightInfo(frm){
//	//���赵 ��� ����	'/2016.05.27 �ѿ�� �߰�
//	if(frm.isUcDeli.value=='True') {
//		alert('�ٹ����� ��� ��ǰ�� ���Ը� �Է��� �� �ֽ��ϴ�.\n\n�ٸ� ��ǰ�� �������ּ���.');
//		return;
//	}

	if (!frm.overSeaYn.value){
		alert('�ؿܹ�� ���θ� �������ּ���.');
		frm.overSeaYn.focus();
		return;
	}

	if (confirm('������ �����Ͻðڽ��ϱ�?')){
		frm.mode.value = "ByWeightProc";
		frm.submit();
	}
}

function InputSizeInfo(frm){
//	//���赵 ��� ����	'/2016.05.27 �ѿ�� �߰�
//	if(frm.isUcDeli.value=='True') {
//		alert('�ٹ����� ��� ��ǰ�� ����� �Է��� �� �ֽ��ϴ�.\n\n�ٸ� ��ǰ�� �������ּ���.');
//		return;
//	}

	if (!frm.itemWeight.value.length){
		alert('��ǰ ���Ը� ��Ȯ�� �Է��ϼ���.');
		frm.itemWeight.focus();
		return;
	}
	if (frm.itemWeight.value*0 != 0) {
		alert('��ǰ ���Դ� ���ڸ� �����մϴ�.');
		frm.itemWeight.value="";
		frm.itemWeight.focus();
		return;
	}

	if (!frm.volX.value.length){
		alert('��ǰ ����� ��Ȯ�� �Է��ϼ���.');
		frm.volX.focus();
		return;
	}

	if (!frm.volY.value.length){
		alert('��ǰ ����� ��Ȯ�� �Է��ϼ���.');
		frm.volY.focus();
		return;
	}

	if (!frm.volZ.value.length){
		alert('��ǰ ����� ��Ȯ�� �Է��ϼ���.');
		frm.volZ.focus();
		return;
	}

	if (frm.volX.value*0 != 0) {
		alert('��ǰ ������� ���ڸ� �����մϴ�.');
		frm.volX.value="";
		frm.volX.focus();
		return;
	}

	if (frm.volY.value*0 != 0) {
		alert('��ǰ ������� ���ڸ� �����մϴ�.');
		frm.volY.value="";
		frm.volY.focus();
		return;
	}
	if (frm.volZ.value*0 != 0) {
		alert('��ǰ ������� ���ڸ� �����մϴ�.');
		frm.volZ.value="";
		frm.volZ.focus();
		return;
	}

	if (confirm('��ǰ ����� �����Ͻðڽ��ϱ�?')){
	    frm.mode.value = "BySizeProc";
		frm.submit();
	}
}

function Research(frm){
	frm.submit();
}

function getOnliad() {
<%
	if (oitembar.FResultCount>0) then
		if Not(oitembar.FOneItem.IsUpcheBeasong) and (prcAfter="") then
			if (mode = "ByWeightProc") then
%>
    document.frmitemWeight.itemWeight.select();
    document.frmitemWeight.itemWeight.focus();
<%
			else
%>
    document.frmitemWeight.volX.select();
    document.frmitemWeight.volX.focus();
<%
			end if
		else
%>
    document.frmbar.itembarcode.select();
    document.frmbar.itembarcode.focus();
<%
		end if
	else
%>
    document.frmbar.itembarcode.select();
    document.frmbar.itembarcode.focus();
<% end if %>
}

function fnCheckRegType(regtype){
	if(regtype=="I"){
		$("#itemW").show();
		$("#itemS").show();
		$("#itemOPT").hide();
	}
	else{
		$("#itemW").hide();
		$("#itemS").hide();
		$("#itemOPT").show();
	}
}

function InputOptionSizeInfo(frm){
	if(frm.regtype[0].checked){
		if (!frm.itemWeight.value.length){
			alert('��ǰ ���Ը� ��Ȯ�� �Է��ϼ���.');
			frm.itemWeight.focus();
			return;
		}
		if (frm.itemWeight.value*0 != 0) {
			alert('��ǰ ���Դ� ���ڸ� �����մϴ�.');
			frm.itemWeight.value="";
			frm.itemWeight.focus();
			return;
		}
		if (!frm.volX.value.length){
			alert('��ǰ ����� ��Ȯ�� �Է��ϼ���.');
			frm.volX.focus();
			return;
		}
		if (!frm.volY.value.length){
			alert('��ǰ ����� ��Ȯ�� �Է��ϼ���.');
			frm.volY.focus();
			return;
		}
		if (!frm.volZ.value.length){
			alert('��ǰ ����� ��Ȯ�� �Է��ϼ���.');
			frm.volZ.focus();
			return;
		}
		if (frm.volX.value*0 != 0) {
			alert('��ǰ ������� ���ڸ� �����մϴ�.');
			frm.volX.value="";
			frm.volX.focus();
			return;
		}
		if (frm.volY.value*0 != 0) {
			alert('��ǰ ������� ���ڸ� �����մϴ�.');
			frm.volY.value="";
			frm.volY.focus();
			return;
		}
		if (frm.volZ.value*0 != 0) {
			alert('��ǰ ������� ���ڸ� �����մϴ�.');
			frm.volZ.value="";
			frm.volZ.focus();
			return;
		}
		if (confirm('��ǰ ����� �����Ͻðڽ��ϱ�?')){
			frm.mode.value = "ByOptionSameSizeProc";
			frm.submit();
		}
	}
	else{
		if (confirm('��ǰ ����� �����Ͻðڽ��ϱ�?')){
			frm.mode.value = "ByOptionSizeProc";
			frm.submit();
		}
	}
}

window.onload=getOnliad;
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<!-- ��ܹ� ���� -->
<form name="frmbar" method=get>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
					<font color="red">&nbsp;<strong>��ǰ�����Է�</strong></font>
				</td>
				<td align="right">
					<input type="text" class="text"  name="itembarcode" value="<%= itembarcode %>" size=14 maxlength=14 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ Research(frmbar); return false;}">
					<input type="button" class="button" value="�˻�" onclick="Research(frmbar)" >
				</td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->
</form>
<% if oitembar.FResultCount>0 then %>
<tr bgcolor="#FFFFFF">
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
	<td colspan="2"><%= oitembar.FOneItem.Fmakername %>(<%= oitembar.FOneItem.Fmakerid %>)</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
	<td colspan="2"><%= oitembar.FOneItem.FItemName %></td>
</tr>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�̹���</td>
	<td colspan="2"><img src="<%= oitembar.FOneItem.Flistimage %>" width="100" height="100" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
</tr>

<form name="frmitemWeight" method=post  action="/warehouse/itemWeight_process.asp" style="margin:0px;">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="mode" value="">
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�ǸŰ�</td>
	<td colspan="2"><%= FormatNumber(oitembar.FOneItem.FSellcash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">��۱���</td>
	<td colspan="2">
		<%=oitembar.FOneItem.GetDeliveryName%>
		<input type="hidden" name="isUcDeli" value="<%=oitembar.FOneItem.IsUpcheBeasong %>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">�ؿܹ�� ����</td>
	<td colspan="2">
		<%
		' �ؿ����� �ϰ��
		if oitembar.FOneItem.Fdeliverfixday = "G" then
			deliverOverseas = "N"
		else
			deliverOverseas = oitembar.FOneItem.FdeliverOverseas
		end if

		if Not(oitembar.FOneItem.IsUpcheBeasong) and (oitembar.FOneItem.FitemWeight<=0) then
			drawSelectBoxUsingYN "overSeaYn", "Y"
			Response.Write "[���� ����: <font color=darkred><b>"
			Response.Write oitembar.FOneItem.FdeliverOverseas
			Response.Write "</b></font>]"
		else
			drawSelectBoxUsingYN "overSeaYn", deliverOverseas
			Response.Write "[���� ����: <font color=darkred><b>"
			Response.Write oitembar.FOneItem.FdeliverOverseas
			Response.Write "</b></font>]"
		end if
		%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">���尡�� ����</td>
	<td colspan="2">
	<%
		if Not(oitembar.FOneItem.IsUpcheBeasong) and (oitembar.FOneItem.FvolX<=0) then
			drawSelectBoxUsingYN "pojangok", "N"
			Response.Write " [���� ����: <font color=darkred><b>"
			Response.Write oitembar.FOneItem.Fpojangok
			Response.Write "</b></font>]"
		else
			drawSelectBoxUsingYN "pojangok", oitembar.FOneItem.Fpojangok
		end if
	%>
	<input type="button" class="button" value="����" onclick="InputWeightInfo(frmitemWeight);">
	</td>
</tr>
<% If itemoption="0000" And oitemoption.FResultCount<1 Then %>
<input type="hidden" name="itemoption" value="<%= itemoption %>">
<% Else %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#E6E6E6">��ǰ ������<br>�������</td>
	<td colspan="2">
		<input type="radio" name="regtype" value="I"<% If regtype="I" Then Response.write " checked" %> onClick="fnCheckRegType('I');">�ϰ���� <input type="radio" name="regtype" value="O" <% If regtype="O" Then Response.write " checked" %> onClick="fnCheckRegType('O');">�ɼǰ������ <input type="button" class="button" value="����" onclick="InputOptionSizeInfo(frmitemWeight);">
	</td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF" id="itemW">
	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ����</td>
	<td colspan="2">
		<%
		'/���赵 ��� ����	'/2016.05.27 �ѿ�� �߰�
		'if Not(oitembar.FOneItem.IsUpcheBeasong) then
		%>
			<input type="text" class="text" name="itemWeight" id="itemWeight" value="<%= oitembar.FOneItem.FitemWeight %>" size="6" AUTOCOMPLETE="off" style="text-align:right;">g
			&nbsp;�� ���Դ� �׷�(g)���� �Է� (��:1.5Kg��1500g)
		<% 'else %>
			<!--<input type="text" class="text" name="itemWeight" value="<%'= oitembar.FOneItem.FitemWeight %>" size="6" readonly onKeyPress="if (event.keyCode == 13){ InputWeightInfo(frmitemWeight); return false;}" style="text-align:right;">g
			&nbsp;�� �ٹ����ٹ�� ��ǰ�� ���Ը� �Է��� �� �ֽ��ϴ�.-->
		<% 'end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="itemS">
	<td bgcolor="#E6E6E6">��ǰ������</td>
	<td colspan="2">
		<input type="text" class="text" name="volX" id="volX" value="<%= oitembar.FOneItem.FvolX %>" size="2" AUTOCOMPLETE="off" style="text-align:right;">
		*
		<input type="text" class="text" name="volY" id="volY" value="<%= oitembar.FOneItem.FvolY %>" size="2" AUTOCOMPLETE="off" style="text-align:right;">
		*
		<input type="text" class="text" name="volZ" id="volZ" value="<%= oitembar.FOneItem.FvolZ %>" size="2" AUTOCOMPLETE="off" style="text-align:right;">
		cm
		&nbsp;�� ��Ƽ����(cm)�� �Է� <% If itemoption="0000" And oitemoption.FResultCount<1 Then %><input type="button" class="button" value="����" onclick="InputSizeInfo(frmitemWeight);"><% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" id="itemOPT" style="display:none">
	<td bgcolor="#E6E6E6">�ɼǺ�</td>
	<td colspan="2">
		<% if oitemoption.FResultCount<1 then %>
		<% else %>
			<table width="100%" align="center" cellpadding="3" cellspacing="1" border="0" class="a" bgcolor="#999999">
				<tr align="center" bgcolor="#E6E6E6">
					<td width="60">�ɼ��ڵ�</td>
					<td>�ɼǻ󼼸�</td>
					<td width="80">����</td>
					<td width="150">��ǰ������</td>
				</tr>
				<% for k=0 to oitemoption.FResultCount -1 %>
				<tr align="center" bgcolor="#FFFFFF">
					<td><%= oitemoption.FItemList(k).Fitemoption %><input type="hidden" name="itemoption" value="<%= oitemoption.FItemList(k).Fitemoption %>"></td>
					<td><%= oitemoption.FItemList(k).Foptionname %></td>
					<td width="80"><input type="text" class="text" name="oitemWeight" id="MitemWeight" value="<%= oitemoption.FItemList(k).FitemWeight %>" size="6" style="text-align:right;">g</td>
					<td width="100"><input type="text" class="text" name="ovolX" id="MvolX" value="<%= oitemoption.FItemList(k).FvolX %>" size="2" style="text-align:right;">*<input type="text" class="text" name="ovolY" id="MvolY" value="<%= oitemoption.FItemList(k).FvolY %>" size="2" style="text-align:right;">*<input type="text" class="text" name="ovolZ" id="MvolZ" value="<%= oitemoption.FItemList(k).FvolZ %>" size="2" style="text-align:right;">cm</td>
				</tr>
				<% Next %>
				<% optionsetYN = True %>
			<table>
		<% end if %>
	</td>
</tr>
</form>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="3" align="center">
		�˻������ �����ϴ�

		<!-- <br>
		���� 10�ڵ�(�¶��ε�ϻ�ǰ)�� ����� �����մϴ�.
		<br>90�ڵ��� ��� ������ǰ������ �̿��ϼ���. -->
	</td>
</tr>
<% end if %>
</table>
<% if optionsetYN Then %>
<script>
$(function(){
	fnCheckRegType('<%= regtype %>');
	$("input[name='regtype']:radio[value='<%= regtype %>']").prop('checked',true);
});
</script>
<% End If %>

<form name="frmsavebar" method=post action="barcode_input_process.asp" style="margin:0px;">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="publicbarcode" value="">
</form>

<%
set oitembar = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
