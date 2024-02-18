<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���� ����� ����Ʈ
' History : 2009.04.07 ������ ����
'			2011.05.16 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->

<%
dim chargeid ,shopid ,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,yyyymmdd1,yyymmdd2
dim fromDate,toDate,tmpdate ,page,notipgo,datesearchtype ,oipchul , moveipchulyn
dim totcnt, totsum1, totsum2 ,i
	page = request("page")
	chargeid = request("chargeid")
	shopid = request("shopid")
	notipgo = request("notipgo")
	datesearchtype = request("datesearchtype")
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	moveipchulyn = request("moveipchulyn")

if page="" then page=1
if datesearchtype="" then datesearchtype="scheduledt"

if (C_IS_SHOP) then
	'����/������
	shopid = C_STREETSHOPID

else
	if (C_IS_Maker_Upche) then
		chargeid = session("ssBctId")
	else
		if not(C_ADMIN_USER) then
		else
		end if
	end if
end if

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()))-1, Cstr(day(now())))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then
    toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()))+1, Cstr(day(now())))
    tmpdate = dateAdd("d",toDate,-1)
    yyyy2 = Cstr(Year(tmpdate))
    mm2 = Cstr(Month(tmpdate))
    dd2 = Cstr(day(tmpdate))
else
    toDate = DateSerial(yyyy2, mm2, dd2+1)
end if

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

set oipchul = new CShopIpChul
	oipchul.FPageSize = 50
	oipchul.FCurrPage = page
	oipchul.FRectDatesearchtype = datesearchtype
	oipchul.FRectStartDay = CStr(fromDate)
	oipchul.FRectEndDay = CStr(toDate)
	oipchul.FRectChargeId = chargeid
	oipchul.FRectShopId = shopid
	oipchul.FRectNotIpgo = notipgo
	oipchul.FRectmoveipchulyn = moveipchulyn
	oipchul.GetIpChulMasterList
%>

<script language='javascript'>

function popsimpleBrandInfo(makerid){
	var popwin = window.open('/common/popsimpleBrandInfo.asp?makerid=' + makerid,'popsimpleBrandInfo','width=400,height=400,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function ReSearch(page){
	frm.page.value = page;
	frm.submit();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function IpgoStateChange(v){
	alert('�����ڸ� ���� ������ �޴��Դϴ�.');
	var popwin = window.open('/common/offshop/pop_offipgostatechange.asp?idx=' + v,'pop_offipgostatechange','width=480,height=300,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function DelThis(v,shopid,chargeid){
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		document.frmipchul.shopid.value=shopid;
		document.frmipchul.chargeid.value=chargeid;	
		document.frmipchul.idx.value=v;
		document.frmipchul.submit();
	}
}

function ReThis(v,comp){
	var ret = confirm('�� �԰�� ��ȯ �Ͻðڽ��ϱ�?');

	if (ret){
		document.frmipchul.mode.value="miipgo";
		document.frmipchul.idx.value=v;
		document.frmipchul.submit();
	}
}

function IpThis(v,comp,shopid,chargeid){

	if (!confirm('�԰� Ȯ�� �Ŀ��� ���������� �Ұ��� �մϴ�.\n������ ���̰� ������� ��ü�� �����Ͽ� ���� �� �����Ͻñ� �ٶ��ϴ�.\n\n - �����Ͻðڽ��ϱ�?(����â���� �԰����� �����ϼ���)')) return;
	if (!calendarOpen4(comp,'�԰���','')) return;

	var ret = confirm('�԰��� : ' + comp.value + '\n�԰� Ȯ�� �Ͻðڽ��ϱ�?');

	if (ret){
		document.frmipchul.shopid.value=shopid;
		document.frmipchul.chargeid.value=chargeid;
		document.frmipchul.mode.value="ipgook";
		document.frmipchul.idx.value=v;
		document.frmipchul.execdate.value = comp.value;

		document.frmipchul.submit();
	}
}

//�԰� ��û
function ReqIpChulInput(){
	var chargeid = frm.chargeid.value;
	var shopid = frm.shopid.value;
	if (chargeid==''){
		alert('����ó�� ���� ������ �ּ���');
		frm.chargeid.focus();
		return;
	}

	document.location = "/common/offshop/shop_ipchulinput.asp?menupos=<%= menupos %>&chargeid=" + chargeid + "&shopid=" + shopid + "&isreq=Y";
}

function ipChulInput(){
	var chargeid = frm.chargeid.value;
	var shopid = frm.shopid.value;
	if (chargeid==''){
		alert('����ó�� ���� ������ �ּ���');
		frm.chargeid.focus();
		return;
	}

	document.location = "/common/offshop/shop_ipchulinput.asp?menupos=<%= menupos %>&chargeid=" + chargeid + "&shopid=" + shopid ;
}

function PopIpgoSheet(v){
	var popwin;
	popwin = window.open('/common/offshop/pop_ipgosheet.asp?idx=' + v,'ipgosheet','width=680,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopIpgoSheetXL(v){
	var popwin;
	popwin = window.open('/common/offshop/pop_ipgosheet.asp?idx=' + v + '&xl=on','ipgosheet','width=680,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopBarCodePrint(v){
	document.iiframe.location.href = "/common/offshop/iframebarcode.asp?idxlist=" + v;
}

function SelBarCodePrt(){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� ������ �����ϴ�.');
		return;
	}

	var idxArr="";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				idxArr = idxArr + frm.idx.value + ","
			}
		}
	}

	if (idxArr.substr(idxArr.length-1,1)==","){
		idxArr = idxArr.substr(0,idxArr.length-1);
	}
	PopBarCodePrint(idxArr);
}

function SelImagePrt(){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� ������ �����ϴ�.');
		return;
	}

	var idxArr="";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				idxArr = idxArr + frm.idx.value + ","
			}
		}
	}

	if (idxArr.substr(idxArr.length-1,1)==","){
		idxArr = idxArr.substr(0,idxArr.length-1);
	}
	var popwin;
	popwin = window.open('popshopImagelist.asp?idx=' + idxArr,'shopitem','width=680,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function sendSMSEmail(idesigner,iidx){
	var popwin = window.open("/admin/offshop/popupchejumunsms_off.asp?designer=" + idesigner + "&idx=" + iidx,"popupchejumunsms","width=600 height=500,scrollbars=yes,resizabled=yes");
	popwin.focus();
}

//���� ��� �̵�
function ipchulmove(isreq){
	var makerid = frm.chargeid.value;
	var shopid = frm.shopid.value;
	if (makerid==''){
		alert('����ó�� ���� ������ �ּ���');
		frm.chargeid.focus();
		return;
	}

	var popwin = window.open('/common/offshop/shop_ipchuldetail_move.asp?menupos=<%= menupos %>&isreq='+isreq+'&firstshopid='+shopid+'&makerid='+makerid,'popwin','width=1024,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();		
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		����ó :<% drawSelectBoxDesignerwithName "chargeid",chargeid %>		
		���� : <% drawSelectBoxOffShop "shopid",shopid %>
		��������̵�:<% Call drawSelectBoxUsingYN("moveipchulyn",moveipchulyn) %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="notipgo" <% if notipgo="on" then response.write "checked" %> >�԰�����ü����
		<select name="datesearchtype">
		<option value="scheduledt" <% if datesearchtype="scheduledt" then response.write "selected" %> >�԰�����
		<option value="execdt" <% if datesearchtype="execdt" then response.write "selected" %> >�԰���
		</select>
		 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>	
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		�� ��ü���� �������� ���� ����� �ϴ°�� ����ϴ� �޴��Դϴ�. (��ü Ư���ΰ�츸 ��밡��)<br>
			&nbsp;&nbsp;&nbsp;&nbsp;- �ٹ����� �������ͷ� �԰��ϴ°�� �������Ϳ��� ����� ������ �Է��մϴ�.<br>
			<% 
			'/��ü�� �ƴҰ��
			if not(C_IS_Maker_Upche) then
				'/������ �̰ų� �����ϰ��
				if getoffshopdiv(shopid) = "1" or C_ADMIN_USER then
			%>
				&nbsp;&nbsp;&nbsp;&nbsp;- <font color="red">��������̵�</font>�ΰ�� ��߸���(<font color="red">���̳ʽ��ֹ�</font>)�� ��������(<font color="red">�԰��ֹ�</font>)�� �ֹ��� ���� �����˴ϴ�<br>
			<%
				end if
			end if
			%>
			&nbsp;&nbsp;&nbsp;&nbsp;- <font color="red">��ǰ</font>�ΰ�� ������ <font color="red">���̳ʽ�</font>�� �Է��մϴ�.<br>			
		�� �԰���� :<br>
			&nbsp;&nbsp;&nbsp;&nbsp;1. <b>�԰���</b> - ��ü���� �������� ��ǰ�� ���������Դϴ�.(������������)<br>
			&nbsp;&nbsp;&nbsp;&nbsp;2. <b>���� �԰�Ȯ��</b> - ���忡�� �԰� Ȯ���� �����Դϴ�.(���������Ұ�)<br>
			&nbsp;&nbsp;&nbsp;&nbsp;3. <b>�԰�Ȯ��(��ü �԰�Ȯ��)</b> - ���� �԰�Ȯ�� �� ��ü���� �԰� Ȯ���� �����Դϴ�.(���������Ұ�)<br>
	</td>
	<td align="right">
	    <input type="button" class="button" value="���ó������ڵ����" onclick="SelBarCodePrt()">
	    <!-- <input type="button" value="���ó����̹������" onclick="SelImagePrt()"> -->	
	    <input type="button" class="button" value="�԰� ��û �Է�  [���ּ� �ۼ�]" onclick="ReqIpChulInput()">
	    <input type="button" class="button" value="�԰�/��ǰ �Է�" onclick="ipChulInput()">
		<% 
		'/��ü�� �ƴҰ��
		if not(C_IS_Maker_Upche) then
			'/������ �̰ų� �����ϰ��
			if getoffshopdiv(shopid) = "1" or C_ADMIN_USER then
		%>
				<input type="button" onclick="ipchulmove('M');" class="button" value="����̵�">
		<% 
			end if
		end if
		%>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= oipchul.FResultCount %></b> <%= Page %>/<%= oipchul.FTotalPage %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>�����ڵ�</td>
	<td><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	<td>�Է�</td>
	<td>����</td>
	<td>����ó</td>
	<td>������</td>
	<td>�ǸŰ�</td>
	<td>���ް�</td>
	<td>����</td>
	<td>�԰�<br>������</td>
	<td>�԰���</td>
	<td>�԰����</td>
	<td>����<br>����</td>
	<td>�԰�<br>������</td>
	<td>���ڵ�<br>���</td>
	<td>����</td>
	<td>�԰�<br>Ȯ��</td>
	<td>���</td>
</tr>
<% if oipchul.FResultCount > 0 then %>
<% for i=0 to oipchul.FResultcount -1 %>
<%
totcnt = totcnt + 1
totsum1 = totsum1 + oipchul.FItemList(i).FTotalSellcash
totsum2 = totsum2 + oipchul.FItemList(i).FTotalSuplycash
%>
<form name="frmBuyPrc_<%= i %>" >
<input type="hidden" name="idx" value="<%= oipchul.FItemList(i).FIdx %>">
<tr bgcolor="#FFFFFF" align="center">
    <td ><%= oipchul.FItemList(i).FIdx %></td>
	<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td ><font color="<%= oipchul.FItemList(i).getInputAreaColor %>"><%= oipchul.FItemList(i).getInputAreaStr %></font></td>
	<td> 
		<% if (oipchul.FItemList(i).FisbaljuExists="Y") then %>
			����
		<% elseif (oipchul.FItemList(i).FisbaljuExists="M") then %>
			����̵�
		<% end if %>
	</td>
	<td ><a href="javascript:popsimpleBrandInfo('<%= oipchul.FItemList(i).FChargeID %>');"><%= oipchul.FItemList(i).FChargeID %></a></td>
	<td ><a href="/common/offshop/shop_ipchuldetail.asp?idx=<%= oipchul.FItemList(i).FIdx %>&menupos=<%= menupos %>"><%= oipchul.FItemList(i).FShopName %></a></td>
	<td align="right"><%= FormatNumber(oipchul.FItemList(i).FTotalSellcash,0) %></td>
	<td align="right"><%= FormatNumber(oipchul.FItemList(i).FTotalSuplycash,0) %></td>
	<td align=center>
	<% if oipchul.FItemList(i).FTotalSellcash<>0 then %>
	<%= 100-CLng(oipchul.FItemList(i).FTotalSuplycash/oipchul.FItemList(i).FTotalSellcash*100*100)/100 %> 
	<% end if %>
	</td>
	<td align=center><%= oipchul.FItemList(i).FScheduleDt %></td>
	<td align=center><%= oipchul.FItemList(i).FExecDt %></td>
	<td >
	<input type=hidden name=yyyymmdd>
    	<% if (C_ADMIN_AUTH) or (C_OFF_AUTH) or (session("ssBctId") = "sangmi")  then %>
    	<!-- �����ڸ� ���º��氡�� -->
		<a href="javascript:IpgoStateChange('<%= oipchul.FItemList(i).FIdx %>')"><font color="<%= oipchul.FItemList(i).GetStateColor %>"><%= oipchul.FItemList(i).GetStateName %></font></a>
		<% else %>
		<font color="<%= oipchul.FItemList(i).GetStateColor %>"><%= oipchul.FItemList(i).GetStateName %></font>
		<% end if %>
	</td>
	<td ><a href="/common/offshop/shop_ipchuldetail.asp?idx=<%= oipchul.FItemList(i).FIdx %>&menupos=<%= menupos %>"><img src="/images/icon_search.jpg" border="0" width="16"></a></td>
	<td align="center">
		<a href="javascript:PopIpgoSheet('<%= oipchul.FItemList(i).FIdx %>')"><img src="/images/iexplorer.gif" border="0" width="21"></a>
		<a href="javascript:PopIpgoSheetXL('<%= oipchul.FItemList(i).FIdx %>')"><img src="/images/iexcel.gif" border="0" width="21"></a>
	</td>
	<td align="center"><a href="javascript:PopBarCodePrint('<%= oipchul.FItemList(i).FIdx %>');"><img src="/images/icon_print02.gif" border="0" ></a></td>
	<td align="center">
		<% if (oipchul.FItemList(i).FStatecd>=7) then %>

		<% else %>
			<a href="javascript:DelThis('<%= oipchul.FItemList(i).FIdx %>','<%= oipchul.FItemList(i).fshopid %>','<%= oipchul.FItemList(i).FChargeID %>')">x</a>
		<% end if %>
	</td>
	<td>
		<% if (oipchul.FItemList(i).FStatecd>7) then %>
		<% else %>
			<input type="button" value="Ȯ��" onclick="javascript:IpThis('<%= oipchul.FItemList(i).FIdx %>',yyyymmdd,'<%= oipchul.FItemList(i).fshopid %>','<%= oipchul.FItemList(i).FChargeID %>')" class="button">
		<% end if %>
	</td>
	<td width=200>
		<%
		'/��ü�� �ƴҰ��
		if not(C_IS_Maker_Upche) then
			if (oipchul.FItemList(i).fsendsms="N") and (isnull(oipchul.FItemList(i).FExecDt)) and (oipchul.FItemList(i).Fstatecd < 0) then
		%>
				<input type="button" class="button" value="SMS" onclick="sendSMSEmail('<%= oipchul.FItemList(i).FChargeID %>','<%= oipchul.FItemList(i).Fidx %>')">
		<%
			end if
		end if

		if oipchul.FItemList(i).fipchulmoveidx <> "" then
		%>
			<br>��������̵������ڵ� : <%= oipchul.FItemList(i).fipchulmoveidx %>
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="6" align="center">�� <%= FormatNumber(totcnt,0) %>��</td>
	<td align="right"><%= FormatNumber(totsum1,0) %></td>
	<td align="right"><%= FormatNumber(totsum2,0) %></td>
	<td colspan="10"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">
		<% if oipchul.HasPreScroll then %>
			<a href="javascript:ReSearch('<%= oipchul.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for i=0 + oipchul.StartScrollPage to oipchul.FScrollCount + oipchul.StartScrollPage - 1 %>
			<% if i>oipchul.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:ReSearch('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
	
		<% if oipchul.HasNextScroll then %>
			<a href="javascript:ReSearch('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
<form name="frmipchul" method="post" action="/common/offshop/shopipchul_process.asp">
	<input type="hidden" name="mode" value="delmaster">
	<input type="hidden" name="idx" value="">
	<input type=hidden name="execdate" >
	<input type=hidden name="shopid" >
	<input type=hidden name="chargeid" >		
</form>
<iframe name="iiframe" src="" frameborder="0" width="0" height="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>
</table>

<%
set oipchul = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->