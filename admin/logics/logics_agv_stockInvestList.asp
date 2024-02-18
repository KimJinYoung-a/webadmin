<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : agv
' History : �̻� ����
'           2020.05.12 ������ ����
'           2020.05.20 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_agvCls.asp"-->
<%
dim code, alinkcode, page,designer, statecd, research, itemid, tplgubun,pcuserdiv,rstate, Chargename
dim chulgocheck, yyyy1,yyyy2,mm1,mm2,dd1,dd2, fromDate, toDate, PrcGbn
dim totalsellcash,totalsuply,totalbuycash,totalsuply_plus,totalsuply_minus, totalitemno, i
	page = requestCheckvar(request("page"),32)
	designer = requestCheckvar(request("designer"),32)
	statecd = requestCheckvar(request("statecd"),32)
	code = requestCheckvar(request("code"),640)
	alinkcode = requestCheckvar(request("alinkcode"),640)
	research = requestCheckvar(request("research"),32)
	itemid = requestCheckvar(request("itemid"),32)
	tplgubun = requestCheckvar(request("tplgubun"),32)
	pcuserdiv = requestCheckvar(request("pcuserdiv"),32)
	rstate= requestCheckvar(request("rstate"),32)
	Chargename= requestCheckvar(request("Chargename"),32)
	chulgocheck = requestCheckvar(request("chulgocheck"),32)
	yyyy1 = requestCheckvar(request("yyyy1"),32)
	yyyy2 = requestCheckvar(request("yyyy2"),32)
	mm1	  = requestCheckvar(request("mm1"),32)
	mm2	  = requestCheckvar(request("mm2"),32)
	dd1	  = requestCheckvar(request("dd1"),32)
	dd2	  = requestCheckvar(request("dd2"),32)
	PrcGbn	  = requestCheckvar(request("PrcGbn"),32)

if page="" then page=1
if (research="") then chulgocheck="on"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), day(now()) - 7)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), day(now()))

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	yyyy2 = Cstr(Year(toDate))
	mm2 = Cstr(Month(toDate))
	dd2 = Cstr(day(toDate))
end if

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))


dim oAGV
Set oAGV = new CAGVItems
    oAGV.FPageSize = 1000
    oAGV.FCurrPage = page

    if (chulgocheck <> "") then
        oAGV.FRectStartDate = fromDate
		oAGV.FRectEndDate   = toDate
    end if

    oAGV.GetStockInvestMasterList

%>

<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script src="/js/jquery.placeholder.min.js"></script>
<script type="text/javascript">

function PopUpcheBrandInfoEdit(v){
	window.open("/admin/member/popupchebrandinfo.asp?designer=" + v,"PopUpcheBrandInfoEdit","width=640,height=580,scrollbars=yes,resizable=yes");
}

function jsStockInvestInput() {
	location.href="logics_agv_stockinvestinput.asp?menupos=<%= menupos %>";
}

function ChulgoEdit(masteridx) {
	location.href="logics_agv_stockinvestedit.asp?menupos=<%= menupos %>&idx=" + masteridx;
}

function PopChulgoSheet(v,itype){
	var popwin;
	popwin = window.open('popchulgosheetNew.asp?idx=' + v + '&itype=' + itype,'popchulgosheetNew','width=760,height=600,scrollbars=yes,status=no,resizable=yes');
	popwin.focus();
}

function ExcelSheet(v,itype){
	window.open('popchulgosheetNew.asp?idx=' + v + '&itype=' + itype + '&xl=on');
}

function EnDisabledDateBox(comp){
	document.frm.yyyy1.disabled = !comp.checked;
	document.frm.yyyy2.disabled = !comp.checked;
	document.frm.mm1.disabled = !comp.checked;
	document.frm.mm2.disabled = !comp.checked;
	document.frm.dd1.disabled = !comp.checked;
	document.frm.dd2.disabled = !comp.checked;
}

function NextPage(page){
	ClearPlaceHolder();
	document.frm.page.value = page;
	document.frm.submit();
}

function trim(value) {
	return value.replace(/^\s+|\s+$/g,"");
}

// ��ǰ�ڵ� üũ
function isUInt(val) {
	var re = /^[0-9]+$/;
	return re.test(val);
}

function SubmitFrm(frm) {
	frm.submit();
}

function popXL(fromDate, toDate) {
	<% if chulgocheck<>"on" then %>
	alert("���� ������� �����ϼ���.");
	return;
	<% end if %>

	var popwin = window.open("/admin/newstorage/pop_chulgolist_xl_download.asp?fromDate=" + fromDate + "&toDate=" + toDate,"popXL","width=100,height=100 scrollbars=yes resizable=yes");
	popwin.focus();
}

//���ڰ��� ǰ�Ǽ� ���
function jsRegEapp(scmidx,executedt){
	var BasicMonth ="<%= CStr(DateSerial(Year(now()),Month(now())-1,1))%>";
 	if ( executedt=="" ){
		alert("�̹� ���ó�� �Ͽ����ϴ�.");
		return;
	}

	if (executedt.length<1){
		alert('������� �Է��ϼ���.');
		calendarOpen(frm.executedt);
		return;
	}
	<% if Not (C_ADMIN_AUTH) then %>
		if ((executedt!='')&&(executedt< BasicMonth)){
			alert('������� �δ� ���� ��¥�δ� ���� �Ұ� �մϴ�.');
			return;
		}
	<% end if %>

	var winEapp = window.open("/admin/approval/eapp/regeapp.asp","popE","width=1000,height=600,scrollbars=yes,resizable=yes");
	document.frmEapp.iSL.value = scmidx;
	document.frmEapp.tC.value = document.all.divEapp.innerHTML.replace(/\r|\n/g,"");
	document.frmEapp.target = "popE";
	document.frmEapp.submit();
	winEapp.focus();
}

//���ڰ��� ǰ�Ǽ� ���뺸��
function jsViewEapp(reportidx,reportstate){
	var winEapp = window.open("/admin/approval/eapp/popIndex.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
	winEapp.focus();
}

function ClearPlaceHolder() {
	var frm = document.frm;
	frm.code.value = $('#code').val();
	frm.alinkcode.value = $('#alinkcode').val();
}

// �������
function uploadexcel(){
	document.domain = "10x10.co.kr";
	var popwin = window.open('/admin/logics/logics_agv_stock_invest_excel_upload.asp','adduploadexcel','width=1250,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

$( document ).ready(function() {
    $('textarea').placeholder();
});

</script>

<style>
textarea:-webkit-input-placeholder {color:#acacac;}
textarea:-moz-placeholder {color:#acacac;}
textarea:-ms-input-placeholder {color:#acacac;}
.placeholder { color: #acacac; }
</style>

<!-- ǥ ��ܹ� ����-->

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="3" width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
    <td align="left">
        <input type=checkbox name="chulgocheck" <% if chulgocheck="on" then  response.write "checked" %>> �����
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
    </td>
    <td rowspan="1" width="50" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:SubmitFrm(document.frm);">
	</td>
</tr>
</table>
</form>
<!-- ǥ ��ܹ� ��-->

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" class="button" value="�������" onclick="uploadexcel();">
		<input type="button" value="������" onclick="jsStockInvestInput();" class="button" >
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		�˻���� : <b><%= oAGV.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">IDX</td>
	<td width="60">�����̼�</td>
	<td width="300">����</td>
    <td width="300">�۾������ڵ�</td>
	<td width="100">�����</td>
	<td width="80">����</td>
    <td width="150">����������� ��ȣ</td>
	<td width="150">�����</td>
	<td>���</td>
</tr>
<% if oAGV.FResultCount >0 then %>
	<% for i=0 to oAGV.FResultcount-1 %>
	<tr bgcolor="#FFFFFF" height=24>
		<td align=center>
		  	<%= oAGV.FItemList(i).Fidx %>
		</td>
		<td align="center">
		  	<%= oAGV.FItemList(i).FstationCd %>
		</td>
		<td>
		  	<a href="javascript:ChulgoEdit(<%= oAGV.FItemList(i).Fidx %>)"><%= oAGV.FItemList(i).Ftitle %></a>
		</td>
		<td align="center">
		  	<%= oAGV.FItemList(i).FrequestNo %>
		</td>
		<td align="center">
		  	<%= oAGV.FItemList(i).Freguserid %>
		</td>
		<td align="center">
		  	<%= oAGV.FItemList(i).getStatusName %>
		</td>
		<td align="center">
		  	<%= oAGV.FItemList(i).FinventorySurveyOrderId %>
		</td>
		<td>
		  	<%= oAGV.FItemList(i).Fregdate %>
		</td>
		<td align=center>
	    </td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
<% end if %>
</table>

<%
set oAGV = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
