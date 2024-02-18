<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/db/dbTPLHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/tplTempOrderCls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim tplcompanyid
Dim orderserial, outmallorderserial
Dim sellsite : sellsite = requestCheckvar(request("sellsite"),10)
Dim matchState : matchState = requestCheckvar(request("matchState"),10)
Dim research : research = requestCheckvar(request("research"),10)
Dim page : page = requestCheckvar(request("page"),10)
Dim regyyyymmdd : regyyyymmdd = requestCheckvar(request("regyyyymmdd"),10)
if (research="") then matchState="I"
if (page="") then page=1

tplcompanyid		= requestCheckvar(request("tplcompanyid"),32)
orderserial			= requestCheckvar(request("orderserial"),32)
outmallorderserial	= requestCheckvar(request("outmallorderserial"),32)

Dim otmpOrder
set otmpOrder = new CTplTempOrder
otmpOrder.FPageSize = 200					'�迭�Է��� ������ ������ ������ ���� ����(CallDBSendRequestModifyOnlineSellAfterMulti ����)
otmpOrder.FCurrPage = page
otmpOrder.FRectTPLCompanyID  = tplcompanyid
otmpOrder.FRectSellSite   = sellsite
otmpOrder.FRectMatchState = matchState
otmpOrder.FRectorderserial			= orderserial
otmpOrder.FRectoutmallorderserial	= outmallorderserial
otmpOrder.FRectregYYYYMMDD 			= regyyyymmdd
otmpOrder.getOnlineTmpOrderList


Dim i, pOrderSerial, isNewOrderLine
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function xlOnlineOrderUpload(){
    var winFile = window.open("popRegFile.asp","popFile","width=400, height=300 ,scrollbars=yes,resizable=yes");
	winFile.focus();
}

function xlOnlineSiteReg(){
    var winFile = window.open("/order/orderInput/popRegSite.asp","xlOnlineSiteReg","width=600, height=400 ,scrollbars=yes,resizable=yes");
	winFile.focus();
}

function popMatchItem(){
    var params = "";
    var popWin = window.open("/company/partnercompany/partneritemlink_modify.asp" + params,"popitemLink","width=800, height=600 ,scrollbars=yes,resizable=yes");
    popWin.focus();
}

function chkThis(comp){
    AnCheckClick(comp);
}

function chkValidAll(){
    var frm = document.frmArr;

}


// ============================================================================
function CheckProduct(o) {
	var frm;

	if (o.checked) {
		hL(o);
	 } else {
		dL(o);
		document.frmBuyTop.chk.checked = false;
	}
}

function CheckTop(o) {
	var frm;

	if (o.checked) {
		SelectAll();
	 } else {
		DeselectAll();
	}
}

function DeselectAll() {
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (frm.chk.disabled == false) {
				frm.chk.checked = false;
				CheckProduct(frm.chk);
			}
		}
	}
}

function SelectAll() {
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (frm.chk.disabled == false) {
				frm.chk.checked = true;
				CheckProduct(frm.chk);
			}
		}
	}
}
// ============================================================================



function SubmitInputOrder(){
	var frm;
	var frmarr = document.frmArrupdate;

	frmarr.seqarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (frm.chk.checked == true) {
				frmarr.seqarr.value = frmarr.seqarr.value + "|" + frm.OutMallOrderSeq.value
			}
		}
	}

	if (frmarr.seqarr.value == "") {
		alert("���õ� �ֹ��� �����ϴ�.");
		return;
	}

	var ret = confirm('���� �ֹ��Է��Ͻðڽ��ϱ�?');

	if (ret != true){
		return;
	}

	frmarr.submit();
	// alert(frmarr.seqarr.value);
}

function SubmitDeleteOrderOne(OutMallOrderSeq) {

	var ret = confirm('�����Ͻðڽ��ϱ�?');

	if (ret != true){
		return;
	}

	var frm;
	var frmarr = document.frmArrupdate;

	frmarr.seqarr.value = "";

	frmarr.seqarr.value = "|" + OutMallOrderSeq;

	frmarr.mode.value = "deleteoneorder";
	frmarr.submit();
	// alert(frmarr.seqarr.value);
}

function AddNewPartnerItemLinkWithOrder(SellSite, orderItemID, orderItemName, orderItemOption, orderItemOptionName) {
	var popwin = window.open("/company/partnercompany/partneritemlink_modify_frame.asp?SellSite=" + SellSite + "&orderItemID=" + orderItemID + "&orderItemName=" + orderItemName + "&orderItemOption=" + orderItemOption + "&orderItemOptionName=" + orderItemOptionName,"AddNewPartnerItemLinkWithOrder","width=900 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

$(function() {
	var CAL_Start = new Calendar({
		inputField : "regyyyymmdd", trigger    : "regyyyymmdd_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
});
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">

	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>" class="td_br">�˻�<br>����</td>
		<td align="left">
		    ó������ :
			<select class="select" name="matchState">
			<option value='' <%= chkIIF(matchState="","selected","") %> >��ü</option>
	     	<option value='I' <%= chkIIF(matchState="I","selected","") %> >�������</option>
	     	<option value='P' <%= chkIIF(matchState="P","selected","") %> >��ǰ��Ī�Ϸ�</option>
	     	<option value='O' <%= chkIIF(matchState="O","selected","") %> >�ֹ��Է¿Ϸ�</option>
	     	</select>
			* �ֹ���ȣ:<input type="text" name="orderserial" value="<%=orderserial%>" size="20" maxlength="22"  >
			&nbsp;&nbsp;
			* �����ֹ���ȣ:<input type="text" name="outmallorderserial" value="<%= outmallorderserial %>" size="20" maxlength="22" >
			&nbsp;&nbsp;
			* �ֹ��Է��� :
			<input id="regyyyymmdd" name="regyyyymmdd" value="<%=regyyyymmdd%>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="regyyyymmdd_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>" class="td_br">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" style="padding-top:10;">
	<tr height="25">
		<td align="left">
			<input type="button" class="button" value="1. ���� ���" onClick="xlOnlineOrderUpload();" <%= CHKIIF(C_ADMIN_AUTH, "", "disabled") %>>
		</td>
		<td align="right">
			<input type="button" class="button" value="2. ���ó����ֹ��Է�" onClick="SubmitInputOrder()">
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="left" bgcolor="#F4F4F4" height="25">
		<td colspan="15" class="td_br">
			�˻���� : <b><%= otmpOrder.FTotalcount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= otmpOrder.FTotalPage %></b>
		</td>
	</tr>
	<form name="frmBuyTop" method="post" action="return false">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="20"><input type="checkbox" name="chk" onclick="CheckTop(this);"></td>
	    <td width="150">�Ǹż��θ�</td>
    	<td width="100">�����ֹ���ȣ</td>
		<td width="100">�����ֹ���</td>

      	<td>���޻�ǰ��<br />��ǰ��</td>
      	<td>���޿ɼǸ�<br>�ɼǸ�</td>

      	<td width="100">��ǰ�ڵ�</td>
      	<td width="100">�����ڵ�</td>
		<td width="120">3PL<br />�ֹ���ȣ</td>

      	<td>��ǰ��Ī</td>
    </tr>
    </form>
    <% for i=0 to otmpOrder.FresultCount -1 %>
    <form name="frmBuyPrc_<%= i %>" method="post" action="return false">
    <input type="hidden" name="OutMallOrderSeq" value="<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>">
	<tr align="center">
    	<td class="td_br"><input type="checkbox" name="chk" onclick="CheckProduct(this);" <% if (otmpOrder.FItemList(i).IsItemMatched <> true) or (otmpOrder.FItemList(i).FmatchState = "O") then %>disabled<% end if %>></td>
    	<td class="td_br"><%= otmpOrder.FItemList(i).FSellSiteName %></td>
      	<td class="td_br"><%= otmpOrder.FItemList(i).FOutMallOrderSerial %></td>
		<td class="td_br"><%= otmpOrder.FItemList(i).FOrgDetailKey %></td>

      	<td class="td_br"><%= otmpOrder.FItemList(i).ForderItemName %><br><%= otmpOrder.FItemList(i).FItemName %></td>
		<td class="td_br"><%= otmpOrder.FItemList(i).ForderItemOptionName %><br><%= otmpOrder.FItemList(i).FItemOptionName %></td>

		<td class="td_br">
			<%
			if (otmpOrder.FItemList(i).Fitemid <> "") then
				response.write BF_MakeTenBarcode(otmpOrder.FItemList(i).Fitemgubun, otmpOrder.FItemList(i).Fitemid, otmpOrder.FItemList(i).Fitemoption)
			end if
			%>
		</td>
		<td class="td_br"><%= otmpOrder.FItemList(i).Fbarcode %></td>
		<td class="td_br">
			<%= otmpOrder.FItemList(i).Forderserial %>
			<%
			if IsNull(otmpOrder.FItemList(i).Forderserial) then
				if IsNull(otmpOrder.FItemList(i).ForderItemName) then
					response.write "<font color='red'>��ǰ�� ����!!</font><br />"
				end if
				if IsNull(otmpOrder.FItemList(i).FItemOrderCount) then
					response.write "<font color='red'>��ǰ���� ����!!</font><br />"
				end if
				if IsNull(otmpOrder.FItemList(i).Fitemgubun) then
					response.write "<font color='red'>��ǰ�ڵ��Ī ����!!</font><br />"
				end if
			end if
			%>
		</td>

      	<td class="td_br">
      		<%= otmpOrder.FItemList(i).getmatchStateString %>
			<% if (otmpOrder.FItemList(i).getmatchStateString = "�����Է�") then %>
				<br><input type="button" class="button" value="����" onclick="SubmitDeleteOrderOne('<%= otmpOrder.FItemList(i).FOutMallOrderSeq %>');">
			<% end if %>
      	</td>
    </tr>
    </form>
    <%
    pOrderSerial = otmpOrder.FItemList(i).FOutMallOrderSerial
    %>
    <% next %>
	<tr align="center" bgcolor="#EEEEEE" height="25">
		<td class="td_br" colspan="15">
			<% if otmpOrder.HasPreScroll then %>
			<a href="javascript:NextPage('<%= otmpOrder.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + otmpOrder.StartScrollPage to otmpOrder.FScrollCount + otmpOrder.StartScrollPage - 1 %>
    			<% if i>otmpOrder.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if otmpOrder.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>

<form name="frmArrupdate" method="post" action="orderInput_process.asp">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="seqarr" value="">
</form>

<p>
<%
set otmpOrder = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
