<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/shinhan/lib/just1DayCls.asp"-->
<%
'###############################################
' PageName : Just1Day_write.asp
' Discription : ��������� ����Ʈ ������ ���/����
' History : 2009.10.27 ������ ����
'###############################################

dim justDate,mode,i
mode=request("mode")
justDate=request("justDate")

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--

document.domain = "10x10.co.kr";

function editcont(){
    //���µ��� ���� ������ ���;;
    var frm=document.inputfrm;
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
    
}

function subcheck(){
	var frm=document.inputfrm;

	if(!frm.justDate.value) {
		alert("������ ��¥�� �������ּ���!");
		return;
	} else {
		<% If session("ssAdminPsn") <> "7" Then %>
		if(frm.justDate.value<='<%=date%>') {
			alert("��ǰ�� ����/����� ���� ������ ��¥�� �����մϴ�.");
			return;
		}
		<% End If %>
	}

	if(!frm.itemid.value) {
		alert("����� ��ǰ�� �������ּ���!");
		return;
	}

	if(!frm.salePrice.value) {
		alert("��ǰ�� ���αݾ��� �Է����ּ���!");
		frm.salePrice.focus();
		return;
	} else {
		if(parseInt(frm.salePrice.value)>=parseInt(frm.orgPrice.value)) {
			alert("�ǸŰ����� ���ξ��� ũ�ų� ���� ���� �����ϴ�.\n���ξ��� Ȯ�����ּ���.");
			return;
		}
	}

	if (!frm.saleSuplyCash.value) {
		alert("��ǰ�� ���Աݾ��� �Է����ּ���!");
		frm.saleSuplyCash.focus();
		return;
	}
    
    // ���԰��� �����ǸŰ� ���� Ŭ �� ����
    if (frm.saleSuplyCash.value*1>frm.salePrice.value*1) {
		alert("��ǰ�� ���Աݾ��� �Է����ּ���!\n�ظ��Ա޾��� �Ǹ� �ݾ� ���� Ŭ �� �����ϴ�.");
		frm.saleSuplyCash.focus();
		return;
	}
	
	if(!frm.limitNo.value) {
		alert("�������� �Ǹ��� ������ �Է����ּ���.\n\n�� �����ǸŰ� �ƴ϶�� 0�� �Է����ּ���.");
		frm.limitNo.focus();
		return;
	}
    
    //eastone �߰� �ǸŰ�0,���԰�0 ���ε�� ����.
    if ((frm.salePrice.value=="0")&&(frm.saleSuplyCash.value=="0")){
        if (!confirm('�����ǸŰ� 0, ���θ��԰� 0���� ��Ͻ� ���� ���� �ʽ��ϴ�. ����Ͻðڽ��ϱ�?')){
            return;
        }
    }
    
	frm.submit();
}

function popItemWindow(tgf){
	var popup_item = window.open("/common/pop_singleItemSelect.asp?target=" + tgf + "&ptype=just1day", "popup_item", "width=800,height=500,scrollbars=yes,status=no");
	popup_item.focus();
}

function putPercent(){
	var pct, frm = document.inputfrm;
	if(frm.orgPrice.value==0||frm.salePrice.value==0) {
		pct = "0%";
	}
	else {
		pct = 1 - (frm.salePrice.value / frm.orgPrice.value);
		pct = pct * 100;
		pct = Math.round(pct*10) / 10 
		pct = pct + "%";
	}
	frm.saleRate.value= pct;
}

function delitems(){
	var frm = document.inputfrm;
	if (confirm('�� �������� �����Ͻðڽ��ϱ�?')) {
		frm.mode.value="delete";
		frm.submit();
	}
}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="doJust1Day_Process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<% =mode %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<font color="red"><b>����Ʈ ������ ���/����</b></font>
	</td>
</tr>
<% if mode="add" then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��¥</td>
	<td bgcolor="#FFFFFF">
        <input id="justDate" name="justDate" value="<%=justDate%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="justDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "justDate", trigger    : "justDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ǰ</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="itemid" value="" size="10" readonly>
		<input type="button" class="button" value="ã��" onClick="popItemWindow('inputfrm')">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		���αݾ� <input type="text" class="text" name="salePrice" value="" size="10" style="text-align:right" onkeyup="putPercent()">��
		/ �ǸŰ� <input type="text" class="text_ro" name="orgPrice" value="0" size="8" readonly style="text-align:right">��,
		������ <input type="text" class="text_ro" name="saleRate" value="0%" size="4" readonly style="text-align:center">
		<br>���Աݾ� <input type="text" class="text" name="saleSuplyCash" value="" size="8" style="text-align:right">��
		<br>(���Աݾ� 0�̸� ���� ��ǰ ���԰� ���)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="limitNo" value="100" size="4" style="text-align:right">
		(�������� 0���� ������ ������ ���� �Ǹŵ˴ϴ�.)
		<input type="hidden" name="limitYn" value="">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		<textarea name="justDesc" class="textarea" cols="80" rows="3"></textarea>
	</td>
</tr>
<% elseif mode="edit" then %>
<%
	dim fmainitem
	set fmainitem = New Cjust1Day
	fmainitem.FCurrPage = 1
	fmainitem.FPageSize=1
	fmainitem.FRectDate=justDate
	fmainitem.Getjust1DayList
%>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��¥</td>
	<td bgcolor="#FFFFFF">
		<b><%=fmainitem.FItemList(0).FjustDate%></b>
		<input type="hidden" name="justDate" value="<%=fmainitem.FItemList(0).FjustDate%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ǰ</td>
	<td bgcolor="#FFFFFF">
		<%= "[" & fmainitem.FItemList(0).Fitemid & "] " & fmainitem.FItemList(0).Fitemname %>
		<input type="hidden" name="itemid" value="<%=fmainitem.FItemList(0).Fitemid%>">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		���αݾ� <input type="text" class="text" name="salePrice" value="<%= fmainitem.FItemList(0).FjustSalePrice %>" size="10" style="text-align:right" onkeyup="putPercent()">��
		/ �ǸŰ� <input type="text" class="text_ro" name="orgPrice" value="<%= fmainitem.FItemList(0).ForgPrice %>" size="8" readonly style="text-align:right">��,
		������ <input type="text" class="text_ro" name="saleRate" value="<%= FormatPercent(1-(fmainitem.FItemList(0).FjustSalePrice/fmainitem.FItemList(0).ForgPrice),1) %>" size="4" readonly style="text-align:center">
		<br>���Ա޾� <input type="text" class="text" name="saleSuplyCash" value="<%= fmainitem.FItemList(0).FsaleSuplyCash %>" size="8" style="text-align:right">��
		<br>(���Աݾ� 0�̸� ���� ��ǰ ���԰� ���)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		�������� <input type="text" class="text" name="limitNo" value="<%= fmainitem.FItemList(0).FlimitNo %>" size="4" style="text-align:right">
		- �Ǹż� <input type="text" class="text_ro" name="limitSold" value="<%= fmainitem.FItemList(0).FlimitSold %>" size="3" readonly style="text-align:right">
		(�������� 0���� ������ ������ ���� �Ǹŵ˴ϴ�.)
		<input type="hidden" name="limitYn" value="">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF">
		<textarea name="justDesc" class="textarea" cols="80" rows="3"><%= fmainitem.FItemList(0).FjustDesc %></textarea>
		<input type="button" value=" ���� ���� " class="button" onclick="editcont();">
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"> &nbsp;&nbsp;
		<% if mode="edit" then %><input type="button" value=" ���� " class="button" onclick="delitems();"> &nbsp;&nbsp;<% end if %>
		<input type="button" value=" ��� " class="button" onclick="history.back();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
