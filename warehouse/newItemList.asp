<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �Ż�ǰ����Ʈ
' History : ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newItemCls.asp"-->
<%

dim i, j
dim purchasetype, startDT, endDT, mwdiv
dim yyyy1, yyyy2, mm1, mm2, dd1, dd2, fromDate, toDate
dim designer

yyyy1   = RequestCheckVar(request("yyyy1"),32)
mm1     = RequestCheckVar(request("mm1"),32)
dd1     = RequestCheckVar(request("dd1"),32)
yyyy2   = RequestCheckVar(request("yyyy2"),32)
mm2     = RequestCheckVar(request("mm2"),32)
dd2     = RequestCheckVar(request("dd2"),32)
designer     = RequestCheckVar(request("designer"),32)

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), CStr(Day(Now()) - 14))
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), CStr(Day(Now()) - 7))

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	yyyy2 = Cstr(Year(toDate))
	mm2 = Cstr(Month(toDate))
	dd2 = Cstr(day(toDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2)
end if

purchasetype = RequestCheckVar(request("purchasetype"), 32)
mwdiv = RequestCheckVar(request("mwdiv"), 32)

if (purchasetype = "") then
	purchasetype = "1"
	mwdiv = "M"
end if

startDT = Left(fromDate,10)
endDT = Left(toDate,10)

dim oCNewItem
set oCNewItem = new CNewItem

oCNewItem.FRectPurchaseType = purchasetype
oCNewItem.FRectStartDT = startDT
oCNewItem.FRectEndDT = endDT
oCNewItem.FRectMWDiv = mwdiv
oCNewItem.FRectMakerID = designer

oCNewItem.GetNewItemList()

dim totipgocnt, totonsellcnt, totoffsellcnt

%>
<script language='javascript'>
function popViewCurrentStock(itemgubun, itemid, itemoption) {
	var popwin;
	popwin = window.open('/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popViewCurrentStock','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			* ��������:
			<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",purchasetype,"" %>
			<select class="select" name="mwdiv">
				<option value="">����+��Ź</option>
				<option value="M" <%= CHKIIF(mwdiv="M", "selected", "") %>>����</option>
				<option value="W" <%= CHKIIF(mwdiv="W", "selected", "") %>>��Ź</option>
			</select>
			* �귣�� : <% drawSelectBoxDesignerwithName "designer", designer %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			* �Ⱓ(������) :
			<% DrawDateBoxdynamic yyyy1, "yyyy1", yyyy2, "yyyy2", mm1, "mm1", mm2, "mm2", dd1, "dd1", dd2, "dd2" %>
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p />

* �ִ� 1000�������� ǥ�õ˴ϴ�.

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="100">��������</td>
		<td width="150">�귣��ID</td>
    	<td width="80">��ǰID</td>
    	<td>��ǰ��</td>
		<td>�ɼǸ�</td>
    	<td width="80">�԰���</td>
    	<td width="80">������</td>
		<td width="50">�԰��</td>
		<td width="50">ON�Ǹ�</td>
		<td width="50">OFF�Ǹ�</td>
		<td width="50">���Ǹŷ�</td>
		<td>���</td>
    </tr>
<% for i=0 to oCNewItem.FResultCount - 1 %>
	<%
	totipgocnt = totipgocnt + oCNewItem.FItemList(i).Fipgocnt
	totonsellcnt = totonsellcnt + oCNewItem.FItemList(i).Fonsellcnt
	totoffsellcnt = totoffsellcnt + oCNewItem.FItemList(i).Foffsellcnt
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= getBrandPurchaseType(oCNewItem.FItemList(i).Fpurchasetype) %></td>
		<td><%= oCNewItem.FItemList(i).Fmakerid %></td>
		<td><%= oCNewItem.FItemList(i).Fitemid %></td>
		<td>
			<a href="javascript:popViewCurrentStock('<%= oCNewItem.FItemList(i).Fitemgubun %>', '<%= oCNewItem.FItemList(i).Fitemid %>', '<%= oCNewItem.FItemList(i).Fitemoption %>');">
				<%= oCNewItem.FItemList(i).Fitemname %>
			</a>
		</td>
		<td><%= oCNewItem.FItemList(i).Fitemoptionname %></td>
		<td><%= oCNewItem.FItemList(i).Fipgodate %></td>
		<td><%= oCNewItem.FItemList(i).FsellSTDate %></td>
		<td><%= oCNewItem.FItemList(i).Fipgocnt %></td>
		<td><%= oCNewItem.FItemList(i).Fonsellcnt %></td>
		<td><%= oCNewItem.FItemList(i).Foffsellcnt %></td>
		<td><%= (oCNewItem.FItemList(i).Fonsellcnt + oCNewItem.FItemList(i).Foffsellcnt) %></td>
		<td></td>
	</tr>
<% next %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="4">�հ�</td>
		<td><%= oCNewItem.FResultCount %></td>
		<td></td>
		<td></td>
		<td><%= totipgocnt %></td>
		<td><%= totonsellcnt %></td>
		<td><%= totoffsellcnt %></td>
		<td><%= (totonsellcnt + totoffsellcnt) %></td>
		<td></td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
