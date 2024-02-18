<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ⱸ�� ��ǰ
' History : 2016.06.16 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/items/standing/item_standing_cls.asp"-->
<%
dim itemid, itemoption, i, menupos, sendkey, item_option_, identikey
	itemid = getNumeric(requestcheckvar(request("itemid"),10))
	itemoption = requestcheckvar(request("itemoption"),4)
	menupos = requestcheckvar(request("menupos"),10)
	item_option_ = requestcheckvar(request("item_option_"),4)

if itemid="" or isnull(itemid) or itemoption="" or isnull(itemoption) then
	response.write "<script type='text/javascript'>alert('��ǰ�ڵ尡 �ɼ��ڵ尡 �����ϴ�.');</script>"
	dbget.close() : response.end
end if

dim oitem
set oitem = new Citemstanding
	oitem.FRectItemID = itemid
	oitem.FRectitemoption = itemoption
	oitem.fitemstanding_item

if oitem.FTotalCount < 1 then
	response.write "<script type='text/javascript'>alert('�ش� ��ǰ�� �ɼ� ������ �����ϴ�.');</script>"
	dbget.close() : response.end
end if

dim ostanding
set ostanding = new Citemstanding
	ostanding.FPageSize = 1
	ostanding.FCurrPage = 500
	ostanding.FRectItemID = itemid
	ostanding.FRectitemoption = itemoption
	ostanding.fitemstanding_option
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript' src="/js/jsCal/js/jscal2.js"></script>
<script type='text/javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

function IsDouble(v){
	if (v.length<1) return false;

	for (var j=0; j < v.length; j++){
		if ("0123456789.".indexOf(v.charAt(j)) < 0) {
			return false;
		}
	}
	return true;
}

function pop_itemsearch(itemid, frmname, itemgubunfrm, itemidfrm, itemoptionfrm, itemnamefrm){
	var pop_itemsearch = window.open('<%= getSCMSSLURL %>/admin/itemmaster/standing/pop_item_option_search.asp?itemid='+itemid+'&frmname='+frmname+'&itemgubunfrm='+itemgubunfrm+'&itemidfrm='+itemidfrm+'&itemoptionfrm='+itemoptionfrm+'&itemnamefrm='+itemnamefrm+'&menupos=<%= menupos %>','pop_itemsearch','width=800,height=768,scrollbars=yes,resizable=yes');
	pop_itemsearch.focus();
}

function chkAllchartItem() {
	if($("input[name='identikey']:first").attr("checked")=="checked") {
		$("input[name='identikey']").attr("checked",false);
	} else {
		$("input[name='identikey']").attr("checked","checked");
	}
}

//����
function delstanding(itemgubun, itemid, itemoption, reserveidx){
	frmstandingedit.itemid.value=itemid;
	frmstandingedit.itemoption.value=itemoption;
	frmstandingedit.reserveidx.value=reserveidx;

	if(confirm("���� ���� �Ͻðڽ��ϱ�?")) {
		//frmstandingedit.target="ifproc";
		frmstandingedit.mode.value="standingdel";
		frmstandingedit.action="<%= getSCMSSLURL %>/admin/itemmaster/standing/standingIteminfo_process.asp";
		frmstandingedit.submit();
	}
}

//����Ʈ ��ü ����
function savestandingList() {
	var chk=0;
	$("form[name='frmstanding']").find("input[name='identikey']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� �׸��� �������ּ���.");
		return;
	}

	var identikey;
	for (i=0; i< frmstanding.identikey.length; i++){
		if (frmstanding.identikey[i].checked == true){
			identikey = frmstanding.identikey[i].value;

			if (eval("frmstanding.reserveidx_" + identikey).value!='<%= oitem.FOneItem.fstartreserveidx %>'){
				if (eval("frmstanding.reserveDlvDate_" + identikey).value==''){
					alert('�߼����� �ʼ��� �Դϴ�.');
					eval("frmstanding.reserveDlvDate_" + identikey).focus();
					return false;
				}
			}
			if (eval("frmstanding.reserveItemID_" + identikey).value!=''){
				if (!IsDouble(eval("frmstanding.reserveItemID_" + identikey).value)){
					alert('��Ī ��ǰ�ڵ�� ���ڸ� �����մϴ�.');
					eval("frmstanding.reserveItemID_" + identikey).focus();
					return;
				}
			}
	    }
	}

	if(confirm("�����Ͻ� ����Ʈ ������ ���� �Ͻðڽ��ϱ�?")) {
		frmstanding.mode.value="standinglistedit";
		frmstanding.action="<%= getSCMSSLURL %>/admin/itemmaster/standing/standingIteminfo_process.asp";
		frmstanding.submit();
	}
}

// ȣ���, Ŭ���� �ش� üũ�ڽ� üũ�� 
function CheckClick(identikey){
	for(var i=0; i<frmstanding.identikey.length; i++){
		if(frmstanding.identikey[i].value==identikey){
			frmstanding.identikey[i].checked=true;
			break;
		}
	}
}

// ���ⱸ�� ���� ����
function standingitemedit(itemid,itemoption){
	var vreservecount;
		vreservecount = 0;

	if (frmstandingedit.startreserveidx.value!=''){
		if (!IsDouble(frmstandingedit.startreserveidx.value)){
			alert('���ⱸ�� ���� ���� ȸ�� VOL(��ȣ) ���ڸ� �Է� �����մϴ�.');
			frmstandingedit.startreserveidx.focus();
			return;
		}
	}else{
		alert('���ⱸ�� ���� ���� ȸ�� VOL(��ȣ)�� ��ϵ��� �ʾҽ��ϴ�.');
		frmstandingedit.startreserveidx.focus();
		return false;
	}

	if (frmstandingedit.endreserveidx.value!=''){
		if (!IsDouble(frmstandingedit.endreserveidx.value)){
			alert('���ⱸ�� ���� ���� ȸ�� VOL(��ȣ) ���ڸ� �Է� �����մϴ�.');
			frmstandingedit.endreserveidx.focus();
			return;
		}
	}else{
		alert('���ⱸ�� ���� ���� ȸ�� VOL(��ȣ)�� ��ϵ��� �ʾҽ��ϴ�.');
		frmstandingedit.endreserveidx.focus();
		return false;
	}

	vreservecount = (Math.floor(frmstandingedit.endreserveidx.value)-Math.floor(frmstandingedit.startreserveidx.value))+1
	if (vreservecount < 2) {
		alert('���ⱸ�� ���� ȸ�� VOL(��ȣ)�� �߸� �����Ǿ����ϴ�.\n�� ���� Ƚ��(����ȸ��-����ȸ��)�� 2ȸ �̻����� �����ϼ���.');
		return;
	}

	if(confirm("���ⱸ���� �� " + vreservecount + " �� ����˴ϴ�. �����Ͻðڽ��ϱ�?")) {
		frmstandingedit.itemid.value=itemid;
		frmstandingedit.itemoption.value=itemoption;
		frmstandingedit.mode.value="standingitemedit";
		frmstandingedit.action="<%= getSCMSSLURL %>/admin/itemmaster/standing/standingIteminfo_process.asp";
		frmstandingedit.submit();
	}
}

</script>

<form name="frmstandingedit" method="POST" action="" style="margin:0;">
<input type="hidden" name="itemid">
<input type="hidden" name="itemoption">
<input type="hidden" name="reserveidx">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left" bgcolor="#FFFFFF">
	<td height="30" colspan="4">
		���ⱸ�� �⺻����
	</td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>" width="15%">��ǰ�ڵ� :</td>
	<td bgcolor="#FFFFFF" width="35%">
		<%= oitem.FOneItem.Fitemid %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="15%">�ɼ��ڵ� :</td>
	<td bgcolor="#FFFFFF" width="35%"><%=oitem.FOneItem.fitemoption %></td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�� :</td>
	<td bgcolor="#FFFFFF"><%= oitem.FOneItem.Fitemname %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�ɼǸ� :</td>
	<td bgcolor="#FFFFFF"><%= oitem.FOneItem.fitemoptionname %>
	</td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>" width="15%">�귣��ID :</td>
	<td bgcolor="#FFFFFF" width="35%"><%=oitem.FOneItem.FMakerid %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">�ɼǻ�뿩�� :</td>
	<td bgcolor="#FFFFFF"><%= oitem.FOneItem.fisusing %></td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>">���ⱸ�� ���� ȸ�� VOL(��ȣ) :</td>
	<td bgcolor="#FFFFFF" colspan=3>
		<% if (oitem.FOneItem.fstartreserveidx="" or isnull(oitem.FOneItem.fstartreserveidx) or oitem.FOneItem.fendreserveidx="" or isnull(oitem.FOneItem.fendreserveidx)) then %>
			���� : <input type="text" name="startreserveidx" size=6 maxlength=7 class="text" value="<%= oitem.FOneItem.fstartreserveidx %>" <% if not(oitem.FOneItem.fstartreserveidx="" or isnull(oitem.FOneItem.fstartreserveidx)) then response.write " readonly" %> />
			~
			���� : <input type="text" name="endreserveidx" size=6 maxlength=7 class="text" value="<%= oitem.FOneItem.fendreserveidx %>" <% if not(oitem.FOneItem.fendreserveidx="" or isnull(oitem.FOneItem.fendreserveidx)) then response.write " readonly" %> />
			<br>��) 74 ~ 79
		<% else %>
			���� : <%= oitem.FOneItem.fstartreserveidx %> ~ ���� : <%= oitem.FOneItem.fendreserveidx %>
			<input type="hidden" name="startreserveidx" value="<%= oitem.FOneItem.fstartreserveidx %>" >
			<input type="hidden" name="endreserveidx" value="<%= oitem.FOneItem.fendreserveidx %>">
		<% end if %>
	</td>
</tr>
<% if (oitem.FOneItem.fstartreserveidx="" or isnull(oitem.FOneItem.fstartreserveidx) or oitem.FOneItem.fendreserveidx="" or isnull(oitem.FOneItem.fendreserveidx)) then %>
	<tr align="center">
		<td bgcolor="#FFFFFF" colspan=4>
			<input type="button" onClick="standingitemedit('<%= oitem.FOneItem.Fitemid %>','<%= oitem.FOneItem.Fitemoption %>');" value="������������" class="button">
		</td>
	</tr>
<% end if %>

</table>
</form>
<br>
<form name="frmstanding" method="POST" action="" style="margin:0;">
<input type="hidden" name="chkAll" value="N">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" onClick="savestandingList();" value="��������" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= ostanding.FtotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width=30><input type="button" value="��ü" class="button" onClick="chkAllchartItem();"></td>
    <td>����ȸ��<br>Vol.(��ȣ)</td>
    <td>��Ī��ǰ<br>�˻�</td>
    <td>��Ī<br>��ǰ�ڵ�</td>
    <td>��Ī<br>�ɼ��ڵ�</td>
    <td>��Ī<br>��ǰ��</td>
    <td>�߼ۿ�����</td>
	<td>���ⱸ��<br>�����</td>
    <td width=60>���</td>
</tr>
<% if ostanding.FtotalCount>0 and not(oitem.FOneItem.fstartreserveidx="" or isnull(oitem.FOneItem.fstartreserveidx) or oitem.FOneItem.fendreserveidx="" or isnull(oitem.FOneItem.fendreserveidx)) then %>
<%
for i=0 to ostanding.FResultCount - 1
sendkey = sendkey + 1
identikey = ostanding.FItemList(i).fitemoption & "_" & ostanding.FItemList(i).freserveidx
%>

<tr bgcolor="<%=chkIIF(ostanding.FItemList(i).fisusing="Y","#FFFFFF","#DDDDDD")%>" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='<%=chkIIF(ostanding.FItemList(i).fisusing="Y","#FFFFFF","#DDDDDD")%>';>
    <td align="center">
		<input type="checkbox" name="identikey" value="<%= identikey %>"/>
    	<input type="hidden" name="itemid_<%= identikey %>" value="<%= ostanding.FItemList(i).fitemid %>" />
    	<input type="hidden" name="itemoption_<%= identikey %>" value="<%= ostanding.FItemList(i).fitemoption %>" />
	</td>
    <td align="center">
    	<%= ostanding.FItemList(i).freserveidx %>
		<input type="hidden" name="reserveidx_<%= identikey %>" size=6 maxlength=7 class="text" value="<%= ostanding.FItemList(i).freserveidx %>" onKeyup="CheckClick('<%= identikey %>');" />
    </td>
    <td align="center">
    	<input type="button" onclick="CheckClick('<%= identikey %>'); pop_itemsearch('','frmstanding','reserveitemgubun_<%= identikey %>','reserveItemID_<%= identikey %>','reserveItemOption_<%= identikey %>','reserveItemName_<%= identikey %>');" value="��ǰ�˻�" class="button">
    </td>
    <td align="left">
    	<input type="hidden" name="reserveitemgubun_<%= identikey %>" value="<%= ostanding.FItemList(i).freserveitemgubun %>" />
    	<input type="text" name="reserveItemID_<%= identikey %>" size=9 maxlength=10 readonly class="text_ro" value="<%= ostanding.FItemList(i).freserveItemID %>" onclick="alert('��ǰ�˻� ��ư���� �Է��ϼ���.');" />
    </td>
    <td align="left">
    	<input type="text" name="reserveItemOption_<%= identikey %>" size=4 maxlength=5 readonly class="text_ro" value="<%= ostanding.FItemList(i).freserveItemOption %>" onclick="alert('��ǰ�˻� ��ư���� �Է��ϼ���.');" />
    </td>
    <td align="left">
    	<input type="text" name="reserveItemName_<%= identikey %>" size=60 readonly class="text_ro" value="<%= ostanding.FItemList(i).freserveItemName %>" onclick="alert('��ǰ�˻� ��ư���� �Է��ϼ���.');" />
    </td>
	<td align="center">
		<%
		' ù��° �߼��� ��� ������� ����.
		if oitem.FOneItem.fstartreserveidx=ostanding.FItemList(i).freserveidx then
		%>
			�ù����
			<input type="hidden" id="reserveDlvDate_<%= identikey %>" name="reserveDlvDate_<%= identikey %>" value="<%= Left(ostanding.FItemList(i).freserveDlvDate,10) %>" />
		<% else %>
    		<input id="reserveDlvDate_<%= identikey %>" name="reserveDlvDate_<%= identikey %>" value="<%= Left(ostanding.FItemList(i).freserveDlvDate,10) %>" class="text" size="10" maxlength="10" onKeyup="CheckClick('<%= identikey %>');" />
    		<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="reserveDlvDate_<%= identikey %>_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<% end if %>
    </td>
    <td align="center">
    	<%= ostanding.FItemList(i).fstandingusercount %>
    </td>
    <td align="center">
    	<% if ostanding.FItemList(i).freserveItemOption<>"" then %>
    		<!--<input type="button" onClick="delstanding('10','<%'= ostanding.FItemList(i).fitemid %>','<%'= ostanding.FItemList(i).fitemoption %>','<%'= ostanding.FItemList(i).freserveidx %>');" value="����" class="button">-->
    	<% else %>
    		<font color="red">�̵��</font>
    	<% end if %>
    </td>
</tr>

<script type="text/javascript">

	var BKG_Start = new Calendar({
		inputField : "reserveDlvDate_<%= identikey %>", trigger    : "reserveDlvDate_<%= identikey %>_trigger",
		onSelect: function() {
			CheckClick('<%= identikey %>');
			var date = Calendar.intToDate(this.selection.get());
			BKG_End.args.min = date;
			BKG_End.redraw();
			setDefault(ticketreg.bookingStDtTime,'00:00:00');
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});

</script>
<%
Next
%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center">���ⱸ�� ���� �˻������ �����ϴ�. ��ܿ� ���ⱸ�� ������ �Է��� �ּ���.</td>
	</tr>
<% end if %>
</table>
</form>

<%
set oitem=nothing
set ostanding=nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->