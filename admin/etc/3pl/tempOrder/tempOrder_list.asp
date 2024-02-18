<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs �޸�
' History : 2007.01.01 �̻� ����
'           2016.12.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/3pl/tempOrderCls.asp" -->
<!-- #include virtual="/lib/classes/3pl/common.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim i, useyn, companyid, matchState
Dim page, research
	companyid	= requestCheckVar(request("companyid"),32)
	page     	= requestCheckVar(request("page"),10)
	research    = requestCheckVar(request("research"),10)
	matchState  = requestCheckVar(request("matchState"),10)

If page = "" Then page = 1

if (research = "") then
	matchState = "I"
end if


dim oCTPLTempOrder
set oCTPLTempOrder = New CTPLTempOrder
	oCTPLTempOrder.FCurrPage				= page
	oCTPLTempOrder.FRectCompanyID			= companyid
	oCTPLTempOrder.FPageSize				= 20
	oCTPLTempOrder.FRectMatchState			= matchState

oCTPLTempOrder.GetTPLTempOrderList()


dim isCheckBoxDisable, pOrderSerial

%>
<script type="text/javascript">
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function jsPopModi(companyid, prdcode) {
	var popwin = window.open("pop_product_modify.asp?companyid=" + companyid + "&prdcode=" + prdcode,"jsPopModi","width=600 height=400 scrollbars=auto resizable=yes");
	popwin.focus();
}

function jsSubmit(frm) {
	frm.submit();
}

function apiOrderProcess(){
	var frm = document.frm;
	var companyid, partnercompanyid;

	var obj = frm.partnercompany;

	if (obj.value === '') {
		alert('���޻縦 �����ϼ���.');
		return;
	}

	var v = obj.options[obj.selectedIndex].text;

    if (confirm(""+v+"���� �ֹ� ���� ��� �Ͻðڽ��ϱ�?")) {
		v = obj.value.split(',');
		companyid = v[0];
		partnercompanyid = v[1];

		frm = document.frmXSiteOrder;
		frm.mode.value = "getxsiteorderlist";
		frm.companyid.value = companyid;
		frm.partnercompanyid.value = partnercompanyid;
		frm.action = "3PLSiteOrder_Ins_Process.asp"
		frm.submit();
    }
}

function CheckProduct(o) {
	var frm;
	if (o.checked) {
		hL(o);
	} else {
		dL(o);
	}
}

function fnCheckValidAll(bool, comp) {
	var obj;
	for (var i = 0; ; i++) {
		obj = document.getElementById("cksel_" + i);
		if (obj == undefined) { break; }
		if (obj.disabled == true) { continue; }
		obj.checked = bool;
		CheckProduct(obj);
	}
}

function SubmitInputOrder() {
    var checkedOrderSerial = "";
	var obj;
	var frm = document.frmXSiteOrder;

	for (var i = 0; ; i++) {
		obj = document.getElementById("cksel_" + i);

		if (obj == undefined) { break; }
		if (obj.disabled == true) { continue; }
		if (obj.checked == true) {
			checkedOrderSerial = checkedOrderSerial + "," + obj.value;
		}
	}

    if (checkedOrderSerial == "") {
        alert('���� �ֹ��� �����ϴ�.');
        return;
    }

    if (confirm('�ֹ��� �Է� �Ͻðڽ��ϱ�?')) {
        frm.mode.value = "add";
		frm.action = "orderInput_Process.asp";

		frm.arrOutMallOrderSerial.value = checkedOrderSerial;
        frm.submit();
    }
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" height="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ���� : <% Call SelectBoxCompanyID("companyid", companyid, CHKIIF(useyn="Y", "Y", "")) %>
	    &nbsp;&nbsp;
	    * ó������ :
		<select class="select" name="matchState">
			<option value='' <%= chkIIF(matchState="","selected","") %> >��ü</option>
	     	<option value='I' <%= chkIIF(matchState="I","selected","") %> >�������</option>
	     	<option value='O' <%= chkIIF(matchState="O","selected","") %> >�ֹ��Է¿Ϸ�</option>
     	</select>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSubmit(frm);">
	</td>
</tr>
</table>

<div style="float: left; padding:5px;">
	<input type="button" class="button" value="1. ���� ���" onClick="jsPopModi('', '')" disabled>
	<% Call SelectBoxApiInput(companyid, "partnercompany", "", "Y") %>
	<input type="button" class="button" value="1. API���� ���" onClick="apiOrderProcess()">
</div>
<div style="float: right; padding:5px;">
	<input type="button" class="button" value="2. ���ó��� �ֹ��Է�" onClick="SubmitInputOrder()">
</div>

</form>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="21">
		�˻���� : <b><%= FormatNumber(oCTPLTempOrder.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCTPLTempOrder.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="20"><input type="checkbox" name="chkAll" onclick="fnCheckValidAll(this.checked);"></td>
	<td>����</td>
    <td>���޸�</td>
    <td>����<br />�ֹ���ȣ</td>
    <td>����<br />�ֹ���</td>
    <td>�ֹ���<br />������</td>
    <td>�����ڵ�</td>
    <td>����<br />��ǰ�ڵ�</td>
    <td>����<br />�ɼ�</td>
    <td>���޻�ǰ��<br />��ǰ��</td>
    <td>���޿ɼǸ�<br />�ɼǸ�</td>
    <td>����</td>
	<td>����<br />�ֹ���ȣ</td>
    <td>��Ī����</td>
    <td>���</td>
</tr>
<% if (oCTPLTempOrder.FResultCount > 0) then %>
	<% for i = 0 to (oCTPLTempOrder.FResultCount - 1) %>
<%
isCheckBoxDisable = False
if (oCTPLTempOrder.FItemList(i).Fprdcode = "") or IsNull(oCTPLTempOrder.FItemList(i).Fprdcode) then
	'// �����ڵ� ��Ī����
	isCheckBoxDisable = True
elseif (oCTPLTempOrder.FItemList(i).FOrderSerial <> "") then
	'// �ֹ��Է¿Ϸ�
	isCheckBoxDisable = True
end if
%>
	<tr align="center" bgcolor="#FFFFFF" height="25">
		<td>
			<input type="checkbox" id="cksel_<%= i %>" name="cksel" value="<%= oCTPLTempOrder.FItemList(i).FOutMallOrderSerial %>" onclick="CheckProduct(this);" <%= CHKIIF(isCheckBoxDisable = "Y", "disabled" ,"")%> >
		</td>
		<td><%= oCTPLTempOrder.FItemList(i).Fcompanyid %></td>
		<td><%= oCTPLTempOrder.FItemList(i).FSellSiteName %></td>
		<td><%= oCTPLTempOrder.FItemList(i).FOutMallOrderSerial %></td>
		<td><%= oCTPLTempOrder.FItemList(i).FOrgDetailKey %></td>
		<td>
			<%= oCTPLTempOrder.FItemList(i).FOrderName %><br />
			<%= oCTPLTempOrder.FItemList(i).FReceiveName %>
		</td>
		<td><%= oCTPLTempOrder.FItemList(i).Fprdcode %></td>
		<td><%= oCTPLTempOrder.FItemList(i).ForderItemID %></td>
		<td><%= oCTPLTempOrder.FItemList(i).ForderItemOption %></td>
		<td>
			<%= oCTPLTempOrder.FItemList(i).ForderItemName %>
			<% if (oCTPLTempOrder.FItemList(i).ForderItemName<>oCTPLTempOrder.FItemList(i).Fprdname) then %>
			<br /><font color="#FF0000"><%= oCTPLTempOrder.FItemList(i).Fprdname %></font>
			<% end if %>
		</td>
		<td>
			<%= oCTPLTempOrder.FItemList(i).ForderItemOptionName %>
			<% if (oCTPLTempOrder.FItemList(i).ForderItemOptionName<>oCTPLTempOrder.FItemList(i).Fprdoptionname) then %>
			<br /><font color="#FF0000"><%= oCTPLTempOrder.FItemList(i).Fprdoptionname %></font>
			<% end if %>
		</td>
		<td><%= oCTPLTempOrder.FItemList(i).FItemOrderCount %></td>
		<td><%= oCTPLTempOrder.FItemList(i).FOrderSerial %></td>
		<td><%= oCTPLTempOrder.FItemList(i).getmatchStateString() %></td>
		<td></td>
    </tr>
	<% next %>
	<tr height="20">
	    <td colspan="21" align="center" bgcolor="#FFFFFF">
	        <% if oCTPLTempOrder.HasPreScroll then %>
			<a href="javascript:goPage('<%= oCTPLTempOrder.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oCTPLTempOrder.StartScrollPage to oCTPLTempOrder.FScrollCount + oCTPLTempOrder.StartScrollPage - 1 %>
	    		<% if i>oCTPLTempOrder.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oCTPLTempOrder.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="21">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>

</table>

<form name="frmXSiteOrder" method="post" action="">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="companyid" value="">
	<input type="hidden" name="partnercompanyid" value="">
	<input type="hidden" name="arrOutMallOrderSerial" value="">
</form>

<%
set oCTPLTempOrder = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
