<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mdMenu/itemcheckCls.asp" -->
<%
Dim oCheck, research, i, page, dispCate, itemid, makerid, maxDepth
dispCate	= requestCheckvar(request("disp"),16)
research	= requestCheckvar(request("research"),2)
itemid  	= request("itemid")
makerid		= requestCheckvar(request("makerid"),32)
maxDepth	= 1

Dim nowsDate
Dim iSD : iSD	= requestCheckvar(request("iSD"),10)
'�ٹ����� ��ǰ�ڵ� ����Ű�� �˻��ǰ�
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If

If iSD = "" Then
	nowsDate = Left(dateadd("d",-4,Now()), 7) & "-01"
Else
	nowsDate = Left(iSD, 7) & "-01"
End If

SET oCheck = new cCheck
	oCheck.FRectCateCode	= dispCate
	oCheck.FRectItemid		= itemid
	oCheck.FRectMakerid		= makerid
	oCheck.FRectNowsDate	= nowsDate
	oCheck.getBuycashPrimeList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function pop_optionEdit(v){
	var pwin = window.open('/admin/itemmaster/itemlist.asp?itemid='+v,'popOutMallEtcLink','width=1200,height=700,scrollbars=yes,resizable=yes');
	pwin.focus();
}
function pop_extsitejungsancheck(iitemid,iitemoption,iitemcost,ibuycash){
	var pwin2 = window.open('','pop_extsitejungsancheck');
	pwin2.location.href="/admin/etc/extsitejungsan_check.asp?itemid="+iitemid+"&ordsch=on&itemcost="+iitemcost+"&buycash="+ibuycash+"&itemoption="+iitemoption+"&mallsellcash="+iitemcost;

	
	pwin2.focus();
}

</script>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<!-- #include virtual="/admin/etc/difforder/gubunTab.asp"-->
<input type="hidden" name="vTab" value="<%= vTab %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2"><h2>�ػ�ǰ �� �ɼ��߰��ݾ��� ���ް��� �Ҽ������� ��Ÿ�� �ִ� ��</h2></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		���� ���ؿ� : 
		<input id="iSD" name="iSD" value="<%=nowsDate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absbottom" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "iSD", trigger    : "iSD_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		&nbsp;&nbsp;
		��ǰ�ڵ� : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;&nbsp;
		����ī�װ� : <!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->
		&nbsp;&nbsp;
		�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
	</td>
	<td align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
</form>

<br>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		�˻���� : <b><%= FormatNumber(oCheck.FResultCount,0) %></b>
	</td>

</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ǰ�ڵ�</td>
	<td>�귣��ID</td>
	<td>�ǸŰ�</td>
	<td>�������ܰ�</td>
	<td>���ް�</td>
	<td>��ǰ����������</td>
	<td>1Depth����ī�װ���</td>
	<td>����</td>
</tr>
<% If oCheck.FResultCount > 0 Then %>
<% For i=0 to oCheck.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><a href="<%=wwwURL%>/<%= oCheck.FItemList(i).FItemID %>" target="_blank"><%= oCheck.FItemList(i).FItemID %></a><br /><%= oCheck.FItemList(i).FItemoption %></td>
	<td><%= oCheck.FItemList(i).FMakerid %></td>
	<td><%= FormatNumber(oCheck.FItemList(i).FSellCash, 0) %></td>
	<td><%= oCheck.FItemList(i).FitemcostCouponNotApplied %></td>
	<td><%= oCheck.FItemList(i).FBuycash %></td>
	<td><%= oCheck.FItemList(i).FLastupdate %></td>
	<td><%= oCheck.FItemList(i).FCatename %></td>
	<td>
		<input type="button" class="button" value="Check" onclick="pop_optionEdit('<%= oCheck.FItemList(i).FItemID %>');">
		<input type="button" class="button" value="Check2" onclick="pop_extsitejungsancheck('<%= oCheck.FItemList(i).FItemID %>','<%= oCheck.FItemList(i).FItemoption %>','<%= oCheck.FItemList(i).FSellCash %>','<%= replace(Formatnumber(oCheck.FItemList(i).FBuycash,0),",","") %>');">
	</td>

</tr>
<% Next %>
<% Else %>
<tr height="50">
    <td colspan="8" align="center" bgcolor="#FFFFFF">
		�����Ͱ� �����ϴ�
    </td>
</tr>
<% End If %>
</table>
<% SET oCheck = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->