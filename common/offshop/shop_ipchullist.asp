<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� �ֹ�����(��ü)
' History : 2009.04.07 ������ ����
'			2010.05.12 �ѿ�� ����
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
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,chargeid ,shopid ,i ,totcnt, totsum1, totsum2 , tmptoDate ,idx
dim fromDate,toDate ,page,notipgo,datesearchtype ,yyyymmdd1,yyymmdd2 ,oipchul , moveipchulyn
dim edityn, popupyn
idx = requestCheckVar(request("idx"),10)
page = requestCheckVar(request("page"),10)
chargeid = requestCheckVar(request("chargeid"),32)
shopid = requestCheckVar(request("shopid"),32)
notipgo = requestCheckVar(request("notipgo"),2)
datesearchtype = requestCheckVar(request("datesearchtype"),32)
yyyy1 = requestCheckVar(request("yyyy1"),4)
mm1 = requestCheckVar(request("mm1"),2)
dd1 = requestCheckVar(request("dd1"),2)
yyyy2 = requestCheckVar(request("yyyy2"),4)
mm2 = requestCheckVar(request("mm2"),2)
dd2 = requestCheckVar(request("dd2"),2)
moveipchulyn = requestCheckVar(request("moveipchulyn"),1)
popupyn = request("popupyn")

edityn = FALSE
if page="" then page=1
if datesearchtype="" then datesearchtype="scheduledt"

'C_IS_SHOP = TRUE

'/����
if (C_IS_SHOP) then

	'//�������϶�
	if C_IS_OWN_SHOP then

		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/��ü
	if (C_IS_Maker_Upche) then
		chargeid = session("ssBctId")
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
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
    yyyy2 = Cstr(Year(toDate))
    mm2 = Cstr(Month(toDate))
    dd2 = Cstr(day(toDate))
    toDate = dateAdd("d",toDate,+1)
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
oipchul.frectidx = idx
oipchul.frect_IS_Maker_Upche = C_IS_Maker_Upche

'/����� ��ü�� ��쿡�� ��ü ����Ʈ
if C_ADMIN_USER or C_IS_Maker_Upche then
	oipchul.GetIpChulMasterList
else
	if (shopid<>"") then
		oipchul.GetIpChulMasterList
	else
		response.write "<script type='text/javascript'>"
		response.write "	alert('������ ������ �ּ���');"
		response.write "</script>"
	end if
end if

dim BasicMonth,  ThisMonth, lastYyyymm
BasicMonth = CStr(DateSerial(Year(now()),Month(now())-1,1))  '' �ִ� ���� 1��
ThisMonth  = CStr(DateSerial(Year(now()),Month(now())-0,1))  '' �̹��� 1�� �Ǵ� '' ��� �ڻ� �ۼ��� ��

''rw BasicMonth
dim sqlStr
sqlStr = " select top 1 yyyymm from db_summary.dbo.tbl_monthly_accumulated_shopstock_summary"
sqlStr = sqlStr & " where 1=1"
sqlStr = sqlStr & " order by yyyymm desc"

rsget.Open sqlStr, dbget, 1
if Not rsget.Eof then
	lastYyyymm = rsget("yyyymm")
end if
rsget.Close

if (lastYyyymm>=Left(BasicMonth,7)) then
    BasicMonth = ThisMonth
end if

''rw BasicMonth
%>

<script type='text/javascript'>

function popsimpleBrandInfo(makerid){
	var popwin = window.open('/common/popsimpleBrandInfo.asp?memupos=<%=menupos%>&makerid=' + makerid,'popsimpleBrandInfo','width=400,height=400,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function ReSearch(page){

	if(frm.idx.value!=''){
		if (!IsDouble(frm.idx.value)){
			alert('�����ڵ�� ���ڸ� �����մϴ�.');
			frm.idx.focus();
			return;
		}
	}

	frm.page.value = page;
	frm.submit();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function IpgoStateChange(v){
	alert('�����ڸ� ���� ������ �޴��Դϴ�.');
	var popwin = window.open('/common/offshop/pop_offipgostatechange.asp?memupos=<%=menupos%>&idx=' + v,'pop_offipgostatechange','width=480,height=300,scrollbars=yes,resizable=yes');
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

function NextStep(idx){
	var ret = confirm('�԰�Ȯ�� �Ͻðڽ��ϱ�?');

	if (ret){
		document.frmipchul.idx.value=idx;
		document.frmipchul.mode.value = "nextstep";
		document.frmipchul.submit();
	}
}

function IpThis(v, comp, shopid, chargeid, regdate, comm_cd){
    var validIpDate = '<%=BasicMonth%>';

	if (comp.value==''){
		alert('�԰�Ȯ������ �����Ͻ���, �԰�Ȯ����ư�� �ٽ� �����ּ���.');
		calendarOpen(comp);
	    return;
	}

	<% if not(C_ADMIN_AUTH) then %>
	if (comp.value<validIpDate){
		alert('�԰����� (' +validIpDate +') �������� ���� �Ұ��մϴ�.');
		return;
	}

	//�������걸���� ��ü��Ź�� �ƴѰ��
	if (comm_cd!='B012'){
		if (regdate<validIpDate){
			alert(validIpDate + ' ���� �ֹ��� �԰�Ȯ�� �ϽǼ� �����ϴ�.');
			return;
		}
	}
	<% end if %>

	var ret = confirm('�԰� Ȯ�� �Ŀ��� ���������� �Ұ��� �մϴ�.\n������ ���̰� ������� ��ü�� �����Ͽ� ���� �� �����Ͻñ� �ٶ��ϴ�.\n\n - �԰��� : ' + comp.value + '\n�԰� Ȯ�� �Ͻðڽ��ϱ�?');

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

function EidtIpgoDetail(v){
	location.href= "/common/offshop/shop_ipchuldetail.asp?menupos=<%= menupos %>&idx=" + v;
}

function PopIpgoDetail(v){
	var popwin;
	popwin = window.open('/common/offshop/pop_ipgodetail.asp?memupos=<%=menupos%>&idx=' + v,'exipgosheet','width=680,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopIpgoSheet(v){
	var popwin;
	popwin = window.open('/common/offshop/pop_ipgosheet.asp?memupos=<%=menupos%>&idx=' + v,'ipgosheet','width=680,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopIpgoSheetXL(v){
	var popwin;
	popwin = window.open('/common/offshop/pop_ipgosheet.asp?memupos=<%=menupos%>&idx=' + v + '&xl=on','ipgosheet','width=680,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopBarCodePrint(v){
	document.iiframe.location.href = "/common/offshop/iframebarcode.asp?memupos=<%=menupos%>&idxlist=" + v;
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
	popwin = window.open('popshopImagelist.asp?memupos=<%=menupos%>&idx=' + idxArr,'shopitem','width=680,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function sendSMSEmail(idesigner,iidx){
	var popwin = window.open("/admin/offshop/popupchejumunsms_off.asp?memupos=<%=menupos%>&designer=" + idesigner + "&idx=" + iidx,"popupchejumunsms","width=600 height=500,scrollbars=yes,resizabled=yes");
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

function jsAddSheet(idx) {

	var frm = opener.document.frmMaster;
	var shopid = frm.shopid.value;
	var chargeid = frm.chargeid.value;

	var url = "/common/offshop/popshopitem2.asp?shopid=" + shopid + "&chargeid=" + chargeid + "&cp_idx=" + idx;
	location.replace(url);
}

function jsDelMulti() {
	<% if (C_ADMIN_AUTH) or (C_OFF_AUTH) then %>
	<% end if %>
}

function jsIpgoStateChangeMulti() {
	<% if (C_ADMIN_AUTH) or (C_OFF_AUTH) then %>
	var frm;
	var idxArr="";

	for (var i=0; i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked) {
				if ((frm.statecd.value != '7') && (frm.statecd.value != '8')) {
					alert('���� �԰�Ȯ�� �Ǵ� �԰�Ȯ�������� �԰��� ��ȯ �����մϴ�.');
					return;
				}
				if (frm.ExecDt.value < '<%= Left(DateAdd("m", -1, Now()), 7) %>-1') {
					alert('������ ������ ������ �� �����ϴ�.');
					return;
				}
				idxArr = idxArr + frm.idx.value + ","
			}
		}
	}

	var ret;

	if (idxArr == '') {
		alert('���� ������ �����ϴ�.');
		return;
	}

	if (idxArr.substr(idxArr.length-1,1)==","){
		idxArr = idxArr.substr(0,idxArr.length-1);
	}

	var frmArr = document.frmipchul;
	if (confirm('[�����ڱ���]\n\n�԰�Ȯ�� ������ �԰��� ��ȯ�Ͻðڽ��ϱ�?')) {
		frmArr.mode.value = 'modistatemulti';
		frmArr.idx.value = idxArr;
		frmArr.submit();
	}
	<% end if %>
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="popupyn" value="<%= popupyn %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<%
		'����/������
		if (C_IS_SHOP) then
		%>
			����ó :
			<% if shopid = "" then %>
				<% drawSelectBoxDesignerwithName "chargeid",chargeid %>
			<%
			'/������ ���� �Ǿ� ������� ���� �귣�常
			else
			%>
				<% drawSelectBoxDesignerOffWitakContract "chargeid", chargeid, shopid, "'B012','B022','B023'", " ReSearch('');" %>
			<% end if %>
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				���� : <% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<%
		else
			''��ü�ΰ��
			if (C_IS_Maker_Upche) then
		%>
				����ó : <%= chargeid %><input type="hidden" name="chargeid" value="<%= chargeid %>">
				���� : <% drawBoxDirectIpchulOffShopByMakerchfg "shopid", shopid, chargeid ," ReSearch('');","'B012','B022','B023'" %>
		<%
			else
				if (C_ADMIN_USER) then
		%>
					����ó :
					<% if (popupyn <> "") then %>
					<input type="hidden" name="chargeid" value="<%= chargeid %>">
					<%= chargeid %>
					<% else %>
					<% drawSelectBoxDesignerwithName "chargeid",chargeid %>
					<% end if %>
					���� : <% drawSelectBoxOffShop "shopid",shopid %>
		<%
				end if
			end if
		end if
		%>
		��������̵�:<% Call drawSelectBoxUsingYN("moveipchulyn",moveipchulyn) %>
		�����ڵ� : <input type="text" name="idx" value="<%=idx%>" size=10>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:ReSearch('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="notipgo" <% if notipgo="on" then response.write "checked" %> >�԰����� ��ü����
		<select name="datesearchtype">
			<option value="scheduledt" <% if datesearchtype="scheduledt" then response.write "selected" %> >�԰�����
			<option value="execdt" <% if datesearchtype="execdt" then response.write "selected" %> >�԰���
		</select>
		 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
</tr>
</form>
</table>

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		�� ��ü���� �������� ���� ����� �ϴ°�� ����ϴ� �޴��Դϴ�. (��ü ��Ź�ΰ�츸 ��밡��)<br>
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
	<td align="right" valign="bottom">
		<% if (popupyn = "") then %>
		<input type="button" class="button" value="���ù��ڵ����" onclick="SelBarCodePrt()">
    	<!-- <input type="button" value="�����̹������" onclick="SelImagePrt()"> -->
		<% if Not (C_IS_Maker_Upche) then %>
	    	<input type="button" class="button" value="�԰��û�Է�[���ּ��ۼ�]" onclick="ReqIpChulInput()">
	    <% end if %>
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
		<% End If %>
		<% if (C_ADMIN_AUTH) or (C_OFF_AUTH) then %>
		<input type="button" onclick="jsIpgoStateChangeMulti();" class="button" value="�����԰�����ȯ(������)">
		<!--
		<input type="button" onclick="jsDelMulti();" class="button" value="�����԰������(������)">
		-->
		<% end if %>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= oipchul.FTotalCount %></b>
		&nbsp;
		������ : <b><%= Page %> / <%= oipchul.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
    <td>����<br>�ڵ�</td>
	<td>�Է�</td>
	<td>����</td>
	<td>����ó</td>
	<td>������</td>
	<td>�ǸŰ�</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td>���ް�</td>
		<td>����</td>
	<% end if %>

	<td>�԰�<br>������</td>
	<td>�԰���</td>
	<td>���<br>���</td>
	<td>�ֹ����<br>���걸��</td>
	<td>����<br>���걸��</td>
	<td>�԰�<br>����</td>
	<td>��<br>����</td>
	<td>�԰�<br>������</td>
	<td>���ڵ�<br>���</td>
	<td>���</td>
</tr>
<%
if oipchul.FresultCount>0 then

for i=0 to oipchul.FResultcount -1

if C_ADMIN_USER then
	edityn = TRUE

'//�����ϰ�� ���������� ���θ��常
elseif (C_IS_SHOP) then
	if C_STREETSHOPID = oipchul.FItemList(i).FShopid then
		edityn = TRUE
	else
		edityn = FALSE
	end if
else
	edityn = TRUE
end if

totcnt = totcnt + 1
totsum1 = totsum1 + oipchul.FItemList(i).FTotalSellcash
totsum2 = totsum2 + oipchul.FItemList(i).FTotalSuplycash
%>
<form name="frmBuyPrc_<%= i %>" >
<tr bgcolor="#FFFFFF" align="center">
	<td width=20><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
    <td width=60>
		<%= oipchul.FItemList(i).FIdx %>
		<input type="hidden" name="idx" value="<%= oipchul.FItemList(i).FIdx %>">
		<input type="hidden" name="statecd" value="<%= oipchul.FItemList(i).Fstatecd %>">
		<input type="hidden" name="ExecDt" value="<%= oipchul.FItemList(i).FExecDt %>">
	</td>
	<td width=30><font color="<%= oipchul.FItemList(i).getInputAreaColor %>"><%= oipchul.FItemList(i).getInputAreaStr %></font></td>
	<td width=30>
		<% if (oipchul.FItemList(i).FisbaljuExists="Y") then %>
			����
		<% elseif (oipchul.FItemList(i).FisbaljuExists="M") then %>
			���<br>�̵�
		<% end if %>
	</td>

	<% if Not (C_IS_Maker_Upche) then %>
		<td><a href="javascript:popsimpleBrandInfo('<%= oipchul.FItemList(i).FChargeID %>');"><%= oipchul.FItemList(i).FChargeID %></a></td>
	<% else %>
		<td><%= oipchul.FItemList(i).FChargeID %></td>
	<% end if %>

	<td><%= oipchul.FItemList(i).FShopName %></td>
	<td align="right"><font color="<%= oipchul.FItemList(i).getMinusColor(oipchul.FItemList(i).FTotalSellcash) %>"><%= FormatNumber(oipchul.FItemList(i).FTotalSellcash,0) %></font></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right">
			<font color="<%= oipchul.FItemList(i).getMinusColor(oipchul.FItemList(i).FTotalSuplycash) %>">
			<%= FormatNumber(oipchul.FItemList(i).FTotalSuplycash,0) %>
			</font>
		</td>
		<td width=35>
			<% if oipchul.FItemList(i).FTotalSellcash<>0 then %>
				<%= round(100-CLng(oipchul.FItemList(i).FTotalSuplycash/oipchul.FItemList(i).FTotalSellcash*100*100)/100,2) %>
			<% end if %>
		</td>
	<% end if %>

	<td width=70><%= oipchul.FItemList(i).FScheduleDt %></td>
	<td width=70><%= oipchul.FItemList(i).FExecDt %></td>
	<td>
		<%= oipchul.FItemList(i).Fsongjangname %><br>
		<%= oipchul.FItemList(i).Fsongjangno %>
    </td>
    <td <% if oipchul.FItemList(i).FComm_cd = oipchul.FItemList(i).fcomm_cd_jungsan then %> bgcolor="#FFFFFF"<% else %> bgcolor="#f1f1f1"<% end if %> width=80>
        <font color="<%= oipchul.FItemList(i).GetContractColor %>">
        	<%= oipchul.FItemList(i).GetContractName_jungsan %>
        </font>
    </td>
    <td <% if oipchul.FItemList(i).FComm_cd = oipchul.FItemList(i).fcomm_cd_jungsan then %> bgcolor="#FFFFFF"<% else %> bgcolor="#f1f1f1"<% end if %> width=80>
        <font color="<%= oipchul.FItemList(i).GetContractColor %>">
        	<%= oipchul.FItemList(i).GetContractName %>
        </font>
    </td>
	<td width=90>
		<input type="hidden" name="yyyymmdd">
		<% If oipchul.FItemList(i).FStatecd <> -5 Then %>
	    	<%
	    	'�����ڸ� ���º��氡��
	    	if (C_ADMIN_AUTH) or (C_OFF_AUTH) then
	    	%>
				<a href="javascript:IpgoStateChange('<%= oipchul.FItemList(i).FIdx %>')">
				<font color="<%= oipchul.FItemList(i).GetStateColor %>"><%= oipchul.FItemList(i).GetStateName %></font></a>
			<% else %>
				<font color="<%= oipchul.FItemList(i).GetStateColor %>"><%= oipchul.FItemList(i).GetStateName %></font>
			<% end if %>
		<% Else %>
			<%= oipchul.FItemList(i).GetStateName %>
		<% End If %>
	</td>
	<td width=35>
		<a href="javascript:EidtIpgoDetail('<%= oipchul.FItemList(i).FIdx %>')"><img src="/images/icon_search.jpg" border="0" width="16"></a>
	</td>
	<td width=55>
		<% If oipchul.FItemList(i).FStatecd <> -5 Then %>
		<a href="javascript:PopIpgoSheet('<%= oipchul.FItemList(i).FIdx %>')"><img src="/images/iexplorer.gif" border="0" width="21"></a>
		<a href="javascript:PopIpgoSheetXL('<%= oipchul.FItemList(i).FIdx %>')"><img src="/images/iexcel.gif" border="0" width="21"></a>
		<% End If %>
	</td>
	<td width=50>
		<a href="javascript:PopBarCodePrint('<%= oipchul.FItemList(i).FIdx %>');"><img src="/images/icon_print02.gif" border="0"></a>
		<a href="javascript:printbarcode_off('UPCHEJUMUN','<%= oipchul.FItemList(i).FIdx %>','','<%= oipchul.FItemList(i).FChargeID %>','<%= oipchul.FItemList(i).fshopid %>','','','','');">
		<img src="/images/icon_print_ttp.gif" border="0"></a>
	</td>
	<td width=200>
		<% if (popupyn <> "") then %>
		<input type="button" class="button" value="�߰�" onclick="jsAddSheet(<%= oipchul.FItemList(i).FIdx %>);">
		<% Else %>
		<% If oipchul.FItemList(i).FStatecd <> -5 Then %>
			<%
			if C_ADMIN_USER then
				if (oipchul.FItemList(i).FStatecd<7) then
			%>
				<input type="button" onclick="javascript:DelThis('<%= oipchul.FItemList(i).FIdx %>','<%= oipchul.FItemList(i).fshopid %>','<%= oipchul.FItemList(i).FChargeID %>');" value="����" class="button">
			<%
				end if
			else

				if (oipchul.FItemList(i).IsAvailDelete) then
			%>
					<input type="button" <% if not(edityn) then response.write " disabled" %> onclick="javascript:DelThis('<%= oipchul.FItemList(i).FIdx %>','<%= oipchul.FItemList(i).fshopid %>','<%= oipchul.FItemList(i).FChargeID %>');" value="����" class="button">
			<%
				end if
			end if
			%>
		    <% if (C_IS_Maker_Upche) then %>
		        <% if (oipchul.FItemList(i).IsUpcheStateChangeEnabled) then %>
	    			<% if (oipchul.FItemList(i).IsWaitState) then %>
	    				<input type="button" class="button" value="�԰�Ȯ��" onclick="javascript:IpThis('<%= oipchul.FItemList(i).FIdx %>',frmBuyPrc_<%= i %>.yyyymmdd,'<%= oipchul.FItemList(i).fshopid %>','<%= oipchul.FItemList(i).FChargeID %>','<%= left(oipchul.FItemList(i).fregdate,10) %>','<%= oipchul.FItemList(i).FComm_cd %>')">
	    			<% else %>
	    				<input type="button" class="button" value="�԰�Ȯ��" onclick="javascript:NextStep('<%= oipchul.FItemList(i).FIdx %>');">
	    			<% end if %>
	    		<% end if %>
		    <% else %>
		    	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		    		<% if (oipchul.FItemList(i).FStatecd<=7) then %>
		    			<input type="button" value="�԰�Ȯ��" onclick="javascript:IpThis('<%= oipchul.FItemList(i).FIdx %>',frmBuyPrc_<%= i %>.yyyymmdd,'<%= oipchul.FItemList(i).fshopid %>','<%= oipchul.FItemList(i).FChargeID %>','<%= left(oipchul.FItemList(i).fregdate,10) %>','<%= oipchul.FItemList(i).FComm_cd %>')" class="button">
		    		<% END IF %>
		    	<% else %>

		        	<!-- ���� -->
			        <% IF (oipchul.FItemList(i).FComm_cd="B023") THEN %>
		        	    <% if (oipchul.FItemList(i).Fstatecd<>"8") then %>
		        	    	<input type="button" class="button" <% if not(edityn) then response.write " disabled" %> value="�԰�Ȯ��" onclick="javascript:IpThis('<%= oipchul.FItemList(i).FIdx %>',frmBuyPrc_<%= i %>.yyyymmdd,'<%= oipchul.FItemList(i).fshopid %>','<%= oipchul.FItemList(i).FChargeID %>','<%= left(oipchul.FItemList(i).fregdate,10) %>','<%= oipchul.FItemList(i).FComm_cd %>')">
		        	    <% end if %>
		    	    <% ELSE %>
		        		<% if (oipchul.FItemList(i).IsShopStateChangeEnabled) then %>
		        			<% if (oipchul.FItemList(i).IsWaitState) then %>
		        				<input type="button" class="button" <% if not(edityn) then response.write " disabled" %> value="�԰�Ȯ��" onclick="javascript:IpThis('<%= oipchul.FItemList(i).FIdx %>',frmBuyPrc_<%= i %>.yyyymmdd,'<%= oipchul.FItemList(i).fshopid %>','<%= oipchul.FItemList(i).FChargeID %>','<%= left(oipchul.FItemList(i).fregdate,10) %>','<%= oipchul.FItemList(i).FComm_cd %>')">
		        			<% else %>
		        				<input type="button" class="button" <% if not(edityn) then response.write " disabled" %> value="�԰�Ȯ��" onclick="javascript:NextStep('<%= oipchul.FItemList(i).FIdx %>');">
		        			<% end if %>
		        		<% end if %>
		    		<% END IF %>
		    	<% end if %>
	    	<% end if %>
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
		<% Else %>
		<input type="button" onclick="javascript:DelThis('<%= oipchul.FItemList(i).FIdx %>','<%= oipchul.FItemList(i).fshopid %>','<%= oipchul.FItemList(i).FChargeID %>');" value="����" class="button">
		<% End If %>
		<% End If %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="6"></td>
	<td align="right"><%= FormatNumber(totsum1,0) %></td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
		<td align="right"><%= FormatNumber(totsum2,0) %></td>
	<% end if %>

	<td colspan=11></td>
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
	<input type="hidden" name="execdate">
	<input type="hidden" name="shopid">
	<input type="hidden" name="chargeid">
	<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<iframe name="iiframe" src="" frameborder="0" width="0" height="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>
</table>

<%
set oipchul = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
