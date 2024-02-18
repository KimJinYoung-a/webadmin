<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ⱸ�� ����� �߼�
' History : 2016.06.16 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/items/standing/item_standing_cls.asp"-->
<%
dim itemid, itemoption, i, menupos, page, orderserial, userid, sendstatus
dim reserveDlvDate, reserveidx, reserveItemID, reserveItemOption, reserveItemName, regadminid, regdate
dim lastadminid, lastupdate, username, isusing, reloading, jukyogubun
	itemid = getNumeric(requestcheckvar(request("itemid"),10))
	reserveitemid = getNumeric(requestcheckvar(request("reserveitemid"),10))
	menupos = getNumeric(requestcheckvar(request("menupos"),10))
	itemoption = requestcheckvar(request("itemoption"),4)
	page = getNumeric(requestcheckvar(request("page"),10))
	reserveidx = getNumeric(requestcheckvar(request("reserveidx"),10))
	orderserial = requestcheckvar(request("orderserial"),11)
	username = requestcheckvar(request("username"),32)
	userid = requestcheckvar(request("userid"),32)
	isusing = requestcheckvar(request("isusing"),1)
	reloading = requestcheckvar(request("reloading"),2)
	sendstatus = requestcheckvar(request("sendstatus"),10)
	jukyogubun = requestcheckvar(request("jukyogubun"),16)

if reloading="" and isusing="" then isusing="Y"
if page="" then page=1

dim ouser
set ouser = new Citemstanding
	ouser.FPageSize = 300
	ouser.FCurrPage = page
	ouser.FRectItemID = itemid
	ouser.FRectreserveitemid = reserveitemid
	ouser.FRectitemoption = itemoption
	ouser.FRectreserveidx = reserveidx
	ouser.FRectorderserial = orderserial
	ouser.FRectusername = username
	ouser.FRectuserid = userid
	ouser.FRectisusing = isusing
	ouser.FRectsendstatus = sendstatus
	ouser.FRectjukyogubun = jukyogubun
	ouser.fitemstanding_user
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function chkAllchartItem() {
	if($("input[name='uidx']:first").attr("checked")=="checked") {
		$("input[name='uidx']").attr("checked",false);
	} else {
		$("input[name='uidx']").attr("checked","checked");
	}
}

function frmsubmit(page){
	if (frmstanding.itemid.value!=""){
		if (!IsDouble(frmstanding.itemid.value)){
			alert('�Ǹſ� ��ǰ�ڵ�� ���ڸ� �Է� �����մϴ�.');
			frmstanding.itemid.focus();
			return;
		}
	}
	if (frmstanding.reserveitemid.value!=""){
		if (!IsDouble(frmstanding.reserveitemid.value)){
			alert('��� ��ǰ�ڵ�� ���ڸ� �Է� �����մϴ�.');
			frmstanding.reserveitemid.focus();
			return;
		}
	}
	frmstanding.itemoption.value=frmstanding.item_option_.value;

	frmstanding.page.value=page;
	frmstanding.submit();
}

// ��Ÿ��� ���
function editstandinguser(uidx, editmode, reserveidx, itemid, itemoption){
	if (editmode=='RE' || editmode=='EDIT'){
		if (uidx==''){
			alert('�ϷĹ�ȣ�� �����ϴ�.');
			return false;
		}
	}else{
		if (itemid=='' || itemoption==''){
			alert('�Ǹſ��ǰ�ڵ�� �Ǹſ�ɼ��ڵ带 �켱 �˻� �ϼž� ��Ÿ��� ��� �ϽǼ� �ֽ��ϴ�.');
			return false;
		}
	}

	var editstandinguser = window.open('<%= getSCMSSLURL %>/admin/itemmaster/standing/pop_standinguser_edit.asp?uidx='+ uidx +'&editmode='+ editmode + '&reserveidx='+ reserveidx + '&itemid='+ itemid + '&itemoption='+ itemoption +'&menupos=<%= menupos %>','editstandinguser','width=800,height=600,scrollbars=yes,resizable=yes');
	editstandinguser.focus();
}

function savestandingsend() {
	var smsyn='';
	smsyn = 'N';
	if (frmstanding.smsyn.value=='Y'){
		if(confirm("���Բ� ���ڹ߼��� ���� �ϼ̽��ϴ�. ���ڸ� �߼� �Ͻðڽ��ϱ�?") == true) {
			smsyn = 'Y';
		}else{
			return false;
		}
	}

	var chk=0;
	$("form[name='frmstandinglist']").find("input[name='uidx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�߼�ó�� �Ͻ� �׸��� �������ּ���.");
		return;
	}

	var uidx;
	for (i=0; i< frmstandinglist.uidx.length; i++){
		if (frmstandinglist.uidx[i].checked == true){
			uidx = frmstandinglist.uidx[i].value;

		    if ( !(eval("frmstandinglist.sendstatus_" + uidx).value=='0' || eval("frmstandinglist.sendstatus_" + uidx).value=='5') ){
		    	alert('�����Ͻ� �׸��߿� �߼۴�⳪ ��߼۴�� ���°� �ƴ� �׸��� �ֽ��ϴ�.');
				eval("frmstandinglist.sendstatus_" + uidx).focus();
				return false;
		    }
	    }
	}

	if(confirm("�����Ͻ� ���ⱸ���� �߼�ó�� �Ͻðڽ��ϱ�?")) {
		frmstandinglist.smsyn.value=smsyn;
		frmstandinglist.mode.value="savestandingsend";
		frmstandinglist.action="<%= getSCMSSLURL %>/admin/itemmaster/standing/standinguser_process.asp";
		frmstandinglist.target="";
		frmstandinglist.submit();
	}
}

// ���� �ٿ�ε�
function exceldownload(){
	frmstandinglist.action="<%= getSCMSSLURL %>/admin/itemmaster/standing/pop_standinguser_excel.asp";
	frmstandinglist.target="view";
	frmstandinglist.submit();
}

//���ⱸ�� �߼� ������ ��������
function regstandingusersudong(){
	var reserveidx='<%= reserveidx %>';

	if (frmstanding.item_option_.value==''){
		alert('�ɼ��ڵ带 ���� �ϼ���.');
		frmstanding.item_option_.focus();
		return false;
	}
	if (frmstanding.sendkey.value==''){
		alert('�߼������� ���� �ϼ���.');
		frmstanding.sendkey.focus();
		return false;
	}
	if (reserveidx==''){
		alert('����ȸ�� Vol.(��ȣ)�� ��ϵ��� �ʾҽ��ϴ�.');
		return false;
	}

	if(confirm("������ ����1 ������� ������(��߼�����) ������� ���� �ɴϴ�.\n�����Ͻðڽ��ϱ�?\n\n����)�ű��߰�,����,��߼��� ����1 �� �Է��ϼ���.")) {
		frmstandinguserreg.itemid.value='<%= itemid %>';
		frmstandinguserreg.itemoption.value=frmstanding.item_option_.value;
		frmstandinguserreg.sendkey.value=frmstanding.sendkey.value;
		frmstandinguserreg.reserveidx.value=reserveidx;
		frmstandinguserreg.mode.value="standingusersudonginsert";
		frmstandinguserreg.action="<%= getSCMSSLURL %>/admin/itemmaster/standing/standinguser_process.asp";
		frmstandinguserreg.submit();
	}
}

</script>

<form name="frmstanding" method="get" action="" style="margin:0;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reloading" value="ON">

<Br>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �Ǹſ��ǰ�ڵ� : <input type="text" name="itemid" value="<%= itemid %>" size=9 maxlength=10 >
		<% if itemid<>"" then %>
			<input type="hidden" name="itemoption" value="<%= itemoption %>">
			&nbsp;* �Ǹſ�ɼ��ڵ� : <%= getOptionBoxHTML_FrontTypenew_optionisusingN_standingitem(itemid, itemoption, " onchange='frmsubmit("""");'") %>

			<% if itemoption<>"" then %>
				&nbsp;* ȸ�� : <% drawSelectBoxsendkey "reserveidx", reserveidx, itemid, itemoption, " onchange='frmsubmit("""");'" %>
			<% end if %>
		<% else %>
			<input type="hidden" name="itemoption" value="<%= itemoption %>">
			<input type="hidden" name="item_option_" value="<%= itemoption %>">
			<input type="hidden" name="reserveidx" value="<%= reserveidx %>">
		<% end if %>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ��ۻ�ǰ�ڵ� : <input type="text" name="reserveitemid" value="<%= reserveitemid %>" size=9 maxlength=10 >
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* �ֹ���ȣ : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="14" maxlength="14">
		&nbsp;
		* �̸� : <input type="text" class="text" name="username" value="<%= username %>" size="7" maxlength="7">
		&nbsp;
		* ���̵� : <input type="text" class="text" name="userid" value="<%= userid %>" size="12" maxlength="20">
		&nbsp;
		* ��뿩�� : <% drawSelectBoxisusingYN "isusing", isusing, " onchange='frmsubmit("""");'" %>
		&nbsp;
		* ���� : <% drawSelectBoxsendstatus "sendstatus", sendstatus, " onchange='frmsubmit("""");'" %>
		&nbsp;
		* ���� : <% drawSelectBoxjukyo "jukyogubun", jukyogubun, " onchange='frmsubmit("""");'" %>
	</td>
</tr>
</table>
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" onClick="exceldownload();" value="�����ٿ�ε�" class="button">
		<input type="button" onclick="editstandinguser('','SUDONG','<%= reserveidx %>','<%= itemid %>','<%= itemoption %>');" value="��Ÿ�����" class="button">
		<!--<input type="button" onclick="regstandingusersudong();" value="1���������ڰ�������" class="button">-->
	</td>
	<td align="right">
		���ڹ߼� :
		<select name="smsyn">
			<option value="Y">Y</option>
			<option value="N">N</option>
		</select>
		<input type="button" onClick="savestandingsend();" value="���ù߼�ó��" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->
</form>

<form name="frmstandinglist" method="post" action="" style="margin:0;">
<input type="hidden" name="smsyn" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="<%= itemid %>" >
<input type="hidden" name="reserveitemid" value="<%= reserveitemid %>" >
<input type="hidden" name="itemoption" value="<%= itemoption %>">
<input type="hidden" name="reserveidx" value="<%= reserveidx %>">
<input type="hidden" name="orderserial" value="<%= orderserial %>">
<input type="hidden" name="username" value="<%= username %>">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="isusing" value="<%= isusing %>">
<input type="hidden" name="sendstatus" value="<%= sendstatus %>">
<input type="hidden" name="jukyogubun" value="<%= jukyogubun %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= ouser.FtotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= ouser.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=30><input type="button" value="��ü" class="button" onClick="chkAllchartItem();"></td>
    <!--<td width=60>idx</td>-->
    <td width=60>����ȸ��<br>Vol.(��ȣ)</td>
    <td width=60>���<br>��ǰ�ڵ�</td>
    <td width=50>���<br>�ɼ��ڵ�</td>
    <td>��ۻ�ǰ��</td>
	<td width=70>����</td>
    <td width=70>�ֹ���ȣ</td>
    <td width=40>����</td>
    <td width=70>���̵�</td>
    <td width=60>�̸�</td>
	<td width=60>����</td>
	<td width=70>�߼���</td>
	<td width=30>���<br>����</td>
    <td width=60>�Ǹſ�<br>��ǰ�ڵ�</td>
    <td width=50>�Ǹſ�<br>�ɼ��ڵ�</td>
	<td width=60>���</td>
</tr>

<% if ouser.FtotalCount>0 then %>
	<%
	for i=0 to ouser.FResultCount - 1
	%>
	<tr bgcolor="<%=chkIIF(ouser.FItemList(i).fisusing="Y","#FFFFFF","#DDDDDD")%>" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='<%=chkIIF(ouser.FItemList(i).fisusing="Y","#FFFFFF","#DDDDDD")%>'; align="center">
	    <td align="center"><input type="checkbox" name="uidx" value="<%= ouser.FItemList(i).fuidx %>"/></td>
	    <!--<td><%'= ouser.FItemList(i).fuidx %></td>-->
	    <td>
	    	<%= ouser.FItemList(i).freserveidx %>
	    </td>
	    <td>
	    	<%= ouser.FItemList(i).freserveItemID %>
	    </td>
	    <td>
	    	<%= ouser.FItemList(i).freserveItemOption %>
	    </td>
	    <td align="left">
	    	<%= ouser.FItemList(i).freserveItemname %>
	    </td>				
	    <td>
	    	<%= getjukyoname(ouser.FItemList(i).fjukyogubun) %>
	    </td>
	    <td>
	    	<%= ouser.FItemList(i).forderserial %>
	    </td>
	    <td align="center">
	    	<%= ouser.FItemList(i).fitemno %>
	    </td>
	    <td align="center">
	    	<%= ouser.FItemList(i).fuserid %>
	    </td>
	    <td align="center">
	    	<%= ouser.FItemList(i).fusername %>
	    </td>
	    <td align="center">
	    	<input type="hidden" name="sendstatus_<%= ouser.FItemList(i).fuidx %>" class="text_ro" value="<%= ouser.FItemList(i).fsendstatus %>" />
	    	<font color="red"><%= getsendstatusname(ouser.FItemList(i).fsendstatus) %></font>
	    </td>
	    <td align="center">
	    	<%= left(ouser.FItemList(i).fsenddate,10) %>
	    	<Br><%= mid(ouser.FItemList(i).fsenddate,12,11) %>
	    </td>
	    <td align="center">
	    	<%= ouser.FItemList(i).fisusing %>
	    </td>
	    <td>
	    	<%= ouser.FItemList(i).forgitemid %>
	    </td>
	    <td>
	    	<%= ouser.FItemList(i).forgitemoption %>
	    </td>
	    <td align="center">
	    	<% if ouser.FItemList(i).fsendstatus=0 or ouser.FItemList(i).fsendstatus=5 then %>
				<% if ouser.FItemList(i).fjukyogubun<>"ORDER" then %>
					<input type="button" onclick="editstandinguser('<%= ouser.FItemList(i).fuidx %>','EDIT','','','');" value="����" class="button">
				<% end if %>
	    	<% end if %>

			<% if ouser.FItemList(i).fsendstatus=3 or ouser.FItemList(i).fsendstatus=7 then %>
	    		<input type="button" onclick="editstandinguser('<%= ouser.FItemList(i).fuidx %>','RE','','','');" value="��߼�" class="button">
	    	<% end if %>
	    </td>
	</tr>
	<%
	Next
	%>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if ouser.HasPreScroll then %>
			<a href="javascript:frmsubmit('<%= ouser.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ouser.StartScrollPage to ouser.FScrollCount + ouser.StartScrollPage - 1 %>
				<% if i>ouser.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:frmsubmit('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ouser.HasNextScroll then %>
				<a href="javascript:frmsubmit('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center">�˻������ �����ϴ�.</td>
	</tr>
<% end if %>

</table>
</form>
<form name="frmstandinguserreg" method="POST" action="" style="margin:0;">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="reserveidx" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height="400" allowtransparency="true" frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" height="0" allowtransparency="true" frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
set ouser=nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
