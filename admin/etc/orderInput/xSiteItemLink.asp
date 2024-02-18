<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
Dim sellsite : sellsite = requestCheckvar(request("sellsite"),32)
Dim chkDiff : chkDiff = requestCheckvar(request("chkDiff"),10)
Dim chkDiffPrc : chkDiffPrc = requestCheckvar(request("chkDiffPrc"),10)
Dim research : research = requestCheckvar(request("research"),10)
Dim page : page = requestCheckvar(request("page"),10)
Dim itemidarr : itemidarr = requestCheckvar(request("itemidarr"),300)
Dim outmallitemidarr : outmallitemidarr = requestCheckvar(request("outmallitemidarr"),300)

if (page="") then page=1

Dim otmpItem
set otmpItem = new CxSiteTempLinkItem
otmpItem.FPageSize = 20
otmpItem.FCurrPage = page
otmpItem.FRectSellSite   = sellsite
otmpItem.FRectitemidarr  = itemidarr
otmpItem.FRectoutmallitemidarr = outmallitemidarr
otmpItem.FRectStateDiff = chkDiff
otmpItem.FRectPriceDiff = chkDiffPrc

otmpItem.xSiteTempLinkItemList

Dim i, pOrderSerial, isNewOrderLine
%>
<script language='javascript'>
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function xlOnlineOrderUpload(){
    var winFile = window.open("/admin/etc/orderInput/popRegFile.asp","popFile","width=400, height=300 ,scrollbars=yes,resizable=yes");
	winFile.focus();
}

function popMatchItem(mallid,itemid,itemoption){
    if (mallid.length<1){
        alert('���θ� ���� �˻� �� ��� �� �� �ֽ��ϴ�.')
        return;
    }

    var params = "?mallid="+mallid+"&itemid="+itemid+"&itemoption="+itemoption;
    var popWin = window.open("/admin/etc/orderInput/partneritemlink_modify.asp" + params,"popitemLink","width=800, height=600 ,scrollbars=yes,resizable=yes");
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

function fnCheckValidAll(bool, comp){
    var frm = comp.form;

    if (!comp.length){
        if (comp.disabled==false){
            comp.checked = bool;
            AnCheckClick(comp);
        }
    }else{
        for (var i=0;i<comp.length;i++){
            if (comp[i].disabled==false){
                comp[i].checked = bool;
                AnCheckClick(comp[i]);
            }
        }
    }
}

function xlOnlineOrderLotteiMall(){
    var frm = document.frmSvArr;
    frm.mode.value="ltimallreg";
    frm.submit();
}

function SubmitInputOrder(frm){
    var checkedExists = false;
    if (!frm.cksel.length){
        if (frm.cksel.checked){
            checkedExists = true;
        }
    }else{

        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                checkedExists = true;
                break;
            }
        }
    }

    if (!checkedExists){
        alert('���� �ֹ��� �����ϴ�.');
        return;
    }

    if (confirm('�ֹ��� �Է� �Ͻðڽ��ϱ�?')){
        frm.mode.value="add";
        frm.submit();
    }
}

function popUpload(mallid){
	if (mallid.length<1){
        alert('���θ� ���� �˻� �� ��� �� �� �ֽ��ϴ�.')
        return;
    }
	var params = "?mallid="+mallid;
	var popWin2 = window.open("/admin/etc/orderInput/popEtcExcelRegFile.asp" + params,"popUpload","width=800, height=600 ,scrollbars=yes,resizable=yes");
	popWin2.focus();
}

function AddNewPartnerItemLinkWithOrder(SellSite, orderItemID, orderItemName, orderItemOption, orderItemOptionName) {
	var popwin = window.open("/company/partnercompany/partneritemlink_modify_frame.asp?SellSite=" + SellSite + "&orderItemID=" + orderItemID + "&orderItemName=" + orderItemName + "&orderItemOption=" + orderItemOption + "&orderItemOptionName=" + orderItemOptionName,"AddNewPartnerItemLinkWithOrder","width=900 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popMatchItemOptionEdit(outMallorderSeq,Matchitemid){
    var popwin = window.open("popMatchItemOptionEdit.asp?outMallorderSeq="+outMallorderSeq+"&Matchitemid="+Matchitemid,"popMatchItemOptionEdit","width=900 height=580 scrollbars=yes resizable=yes");
    popwin.focus();
}

function delInputOrder(outMallorderSeq,OutMallOrderSerial,orderItemID,orderItemOption){
    if (!confirm('���� �Ͻðڽ��ϱ�?')){
        return;
    }
    var popwin = window.open("OrderInput_Process.asp?mode=delpInputOrder&outMallorderSeq="+outMallorderSeq+"&OutMallOrderSerial="+OutMallOrderSerial+"&orderItemID="+orderItemID+"&orderItemOption="+orderItemOption,"OrderInput_Process","width=100 height=100 scrollbars=yes resizable=yes");
    popwin.focus();
}

function chgComp(comp){
    var frm = comp.form;

    //frm.sellsite.disabled = (comp.checked);
    //frm.matchState.disabled = (comp.checked);
    //frm.orderserial.disabled = (comp.checked);
    //frm.outmallorderserial.disabled = (comp.checked);
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">

	<tr align="center">
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>" class="td_br">�˻�<br>����</td>
		<td align="left" class="td_br">
		    ���θ� ���� :

		    <% call drawSelectBoxXSiteHandItemPartner("sellsite", sellsite) %>

	     	<input type="checkbox" name="chkDiff" <%= CHKIIF(chkDiff="on","checked","") %> > ��ǰ�ǸŻ��� �ٸ�����
            <input type="checkbox" name="chkDiffPrc" <%= CHKIIF(chkDiffPrc="on","checked","") %> > ��ǰ���� �ٸ�����
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>" class="td_br">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr>
	    <td class="td_br">
	        TEN��ǰ��ȣ:<input type="text" name="itemidarr" value="<%=itemidarr%>" size="30" maxlength="100"  >
	     	&nbsp;
	     	���޻�ǰ��ȣ:<input type="text" name="outmallitemidarr" value="<%= outmallitemidarr %>" size="30" maxlength="100" >
	    </td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->


<!--
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr height="25">
		<td align="left">
			<input type="button" class="button" value="1. ���� ���" onClick="xlOnlineOrderUpload();">
			&nbsp;
            <input type="button" class="button" value="�Ե�iMall�ֹ� �ӽõ��" onClick="xlOnlineOrderLotteiMall();">
		</td>
		<td align="right">
			<input type="button" class="button" value="2. ���ó����ֹ��Է�" onClick="SubmitInputOrder(frmSvArr)">
		</td>
	</tr>
</table>
-->
<p>


<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl" >
	<tr height="25">
		<td colspan="11" class="td_br">
			�˻���� : <b><%= otmpItem.FTotalcount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= otmpItem.FTotalPage %></b>
		</td>
		<td align="right" class="td_br">
			<input type="button" class="button" value="�ϰ����" onclick="popUpload(document.frm.sellsite.value);">
		    <input type="button" class="button" value="��ǰ���� �űԵ��" onclick="popMatchItem(document.frm.sellsite.value, '');">
		</td>

	</tr>
	<form name="frmSvArr" method="post" action="OrderInput_Process.asp">
	<input type="hidden" name="mode" value="add">
	<tr align="center" class="tr_tablebar">
	    <!--
	    <td width="20" class="td_br"><input type="checkbox" name="chkAll" onclick="fnCheckValidAll(this.checked,frmSvArr.cksel);"></td>
	    -->
	    <td width="60" class="td_br">�Ǹż��θ�</td>
	    <td width="50" class="td_br">�̹���</td>
	    <td width="50" class="td_br">��ǰ�ڵ�</td>
	    <td width="50" class="td_br">�ɼ��ڵ�</td>
    	<td width="100" class="td_br">��ǰ��</td>
    	<td width="70" class="td_br">(��)�ǸŰ�</td>
    	<td width="70" class="td_br">(��)�Ǹſ���</td>
    	<td width="80" class="td_br">(����)��ǰ�ڵ�</td>
      	<td width="100" class="td_br">(����)��ǰ��</td>
      	<td width="70" class="td_br">(����)�ǸŰ�</td>
      	<td width="70" class="td_br">(����)�Ǹſ���</td>
      	<td width="70" class="td_br">�������</td>
    </tr>

    <% for i=0 to otmpItem.FresultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">

    	<td class="td_br"><%= otmpItem.FItemList(i).FmallID %></td>
    	<td class="td_br"><img src="<%= otmpItem.FItemList(i).FsmallImage %>" width="50"></td>
    	<td class="td_br"><%= otmpItem.FItemList(i).Fitemid %></td>
    	<td class="td_br"><%= otmpItem.FItemList(i).Fitemoption %></td>
    	<td class="td_br"><%= otmpItem.FItemList(i).Fitemname %>
    	<% if (otmpItem.FItemList(i).FitemOptionname<>"") then %>
      	<br><font color=blue><%= otmpItem.FItemList(i).FitemOptionname %></font>
      	<% end if %>
    	</td>
      	<td class="td_br" align="right">
      	<% if IsNULL(otmpItem.FItemList(i).Fsellcash+otmpItem.FItemList(i).FoptAddPrice) then %>
      	    <b><font color=red>null</font></b>
      	<% else %>
          	<% if otmpItem.FItemList(i).Fsellcash+otmpItem.FItemList(i).FoptAddPrice<>otmpItem.FItemList(i).FoutmallPrice then %>
          	<b><font color=red><%= FormatNumber(otmpItem.FItemList(i).Fsellcash+otmpItem.FItemList(i).FoptAddPrice,0) %></font></b>
          	<% else %>
          	<%= FormatNumber(otmpItem.FItemList(i).Fsellcash+otmpItem.FItemList(i).FoptAddPrice,0) %>
          	<% end if %>
        <% end if %>
      	</td>
      	<td class="td_br">
      	<% if otmpItem.FItemList(i).Fsellyn<>"Y" and otmpItem.FItemList(i).FoutmallSellYn="Y" then %>
      	<b><font color=red><%= otmpItem.FItemList(i).Fsellyn %></font></b>
      	<% else %>
      	<%= otmpItem.FItemList(i).Fsellyn %>
      	<% end if %>
      	<% if (otmpItem.FItemList(i).IsOptionSoldout) then %>
      	<br><font color=red>�ɼ�ǰ��</font>
      	<% end if %>

      	<% if (otmpItem.FItemList(i).IsLimitSell) then %>
      	<br><font color=blue>���� <%=otmpItem.FItemList(i).getLimitRemainNo%></font>
      	<% end if %>
      	</td>




      	<td class="td_br"><%= otmpItem.FItemList(i).Foutmallitemid %></td>
      	<td class="td_br"><%= otmpItem.FItemList(i).Foutmallitemname %>
      	<% if (otmpItem.FItemList(i).FoutmallitemOptionname<>"") then %>
      	<br><font color=blue><%= otmpItem.FItemList(i).FoutmallitemOptionname %></font>
      	<% end if %>
      	</td>
      	<td class="td_br" align="right"><%= FormatNumber(otmpItem.FItemList(i).FoutmallPrice,0) %></td>
      	<td class="td_br"><%= otmpItem.FItemList(i).FoutmallSellYn %></td>
      	<td class="td_br">
      	<input type="button" class="button" value="��ǰ���� ����" onclick="popMatchItem('<%= otmpItem.FItemList(i).FmallID %>', '<%= otmpItem.FItemList(i).Fitemid %>', '<%= otmpItem.FItemList(i).Fitemoption %>');">
        </td>
    </tr>

    <% next %>
    <tr height="25" align="center">
		<td class="td_br" colspan="16">
			<% if otmpItem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= otmpItem.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + otmpItem.StartScrollPage to otmpItem.FScrollCount + otmpItem.StartScrollPage - 1 %>
    			<% if i>otmpItem.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if otmpItem.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
	 </form>
</table>


<p>
<%
set otmpItem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->