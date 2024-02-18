<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs���� ��ǰ����
' History : �̻� ����
'			2023.06.12 �ѿ�� ����(ǥ���ڵ����� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_couponcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_itemcouponcls.asp" -->
<!-- #include virtual="/cscenter/lib/csOrderFunction.asp" -->
<%
dim i, idx, orderserial, result
	idx = requestCheckVar(request("idx"),10)

'��ǰ�ɼǺ���
'result = CSOrderModifyItemOption("10032647343", 178977, "0000", "7029")

'��һ�ǰ ����ȭ
'result = CSOrderRestoreCanceledItem("10032647343", 178977, "0000")

'��ǰ���
'result = CSOrderCancelItem("10032647343", 178977, "0000")

'response.write "aaaaaaaaaaaaaaaa" & CS_ORDER_FUNCTION_RESULT

dim ojumunDetail
set ojumunDetail = new CJumunMaster
ojumunDetail.SearchOneJumunDetail idx

orderserial = ojumunDetail.FJumunDetail.FOrderSerial

dim ojumun
set ojumun = new COrderMaster

if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster
end if


if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
end if


If ojumunDetail.FJumunDetail.Fitemoption <> "0000" Then

	Dim sqlStr, rsOption, k, optionText, itemStatus
	dim sqlsub

	sqlsub = "select top 1 optaddprice "
	sqlsub = sqlsub + "from [db_item].[dbo].tbl_item_option "
	sqlsub = sqlsub + "where 1 = 1 "
	sqlsub = sqlsub + "and itemid = " & CStr(ojumunDetail.FJumunDetail.Fitemid) & " "
	sqlsub = sqlsub + "and itemoption = '" & CStr(ojumunDetail.FJumunDetail.Fitemoption) & "' "

	'* �ɼǺ����� <font color=red>�ɼǰ�</font>�� ������ �ɼǻ�ǰ�� �����մϴ�.<br>
	'* �ֹ���� �ɼǰ��ݿ� ������� ���� ��ǰ���� ���� �ɼǰ������� ���մϴ�.<br>
	'* ��ǰ��������(�ǸŰ�,���԰� ��)�� �ֹ���� ������ �����˴ϴ�.<br>
	' �ֹ��� ������ ó���� �Ǿ ǥ��
	sqlStr = " select "
	sqlStr = sqlStr + " v.itemoption, v.optionname "
	sqlStr = sqlStr + " , v.optsellyn, v.optlimityn, v.optlimitno, v.optlimitsold "
	sqlStr = sqlstr + " , 0 as notused "
	sqlStr = sqlStr + " , case when v.optaddprice=IsNULL((" & sqlsub & "),0) " & " then 'T' else 'F' end "
	sqlStr = sqlStr + " , v.isusing "
	sqlStr = sqlStr + " , v.optaddprice "
	sqlStr = sqlStr + " , IsNull(P.regno, 0) as prevregno "
	sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i "
	sqlStr = sqlStr + " left join [db_item].[dbo].tbl_item_option v "
	sqlStr = sqlStr + " on i.itemid=v.itemid "

	'���� CS��ǰ����(����+�Ϸ᳻��, ��ǰ�����������)
	sqlStr = sqlStr + "		LEFT JOIN (" + VbCrlf
	sqlStr = sqlStr + "		    select d.itemid, d.itemoption, sum(confirmitemno) as regno, max(a.id) asId " + VbCrlf
    sqlStr = sqlStr + "		    from" + VbCrlf
    sqlStr = sqlStr + "		    	[db_cs].[dbo].tbl_new_as_list a" + VbCrlf
    sqlStr = sqlStr + "		    	Join [db_cs].[dbo].tbl_new_as_detail d" + VbCrlf
    sqlStr = sqlStr + "		    on a.id=d.masterid" + VbCrlf
    sqlStr = sqlStr + "		    where a.orderserial='" + CStr(orderserial) + "'" + VbCrlf
    sqlStr = sqlStr + "		    and a.divcd in ('A004','A010', 'A111', 'A112')" + VbCrlf                ''��ǰ / ȸ�� / ��ǰ���� �±�ȯȸ��(�ٹ����ٹ��) / ��ǰ���� �±�ȯ��ǰ(��ü���).
    sqlStr = sqlStr + "		    and a.deleteyn='N'" + VbCrlf
    'sqlStr = sqlStr + "		    	and a.currstate='B007'" + VbCrlf					'����+�Ϸ� ��� ���
    sqlStr = sqlStr + "			group by d.itemid, d.itemoption" + VbCrlf
    sqlStr = sqlStr + " ) P " + VbCrlf
    sqlStr = sqlStr + "     ON i.itemid=P.itemid and v.itemoption=P.itemoption" + VbCrlf

	sqlStr = sqlStr + " WHERE 1=1 "
	sqlStr = sqlStr + " and i.itemid=" & ojumunDetail.FJumunDetail.Fitemid & ""
	sqlStr = sqlStr + " order by i.itemid desc, v.itemoption"

	rsget.Open sqlStr,dbget,1
	If Not rsget.EOF Then
		rsOption = rsget.getrows
	End If
	rsget.close()

	'response.write sqlStr
End If


'==============================================================================

dim oordermaster, oorderdetail, selecteditemindex

set oordermaster = new COrderMaster
oordermaster.FRectOrderSerial = orderserial
oordermaster.QuickSearchOrderMaster

if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if


set oorderdetail = new COrderMaster
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.QuickSearchOrderDetail

if (oorderdetail.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oorderdetail.FRectOldOrder = "on"
    oorderdetail.QuickSearchOrderDetail
end if


selecteditemindex = 0
for i = 0 to oorderdetail.FResultCount - 1
	if (CStr(oorderdetail.FItemList(i).Fidx) = CStr(ojumunDetail.FJumunDetail.Fdetailidx)) then
		selecteditemindex = i
	end if
next

dim currentitemoptionidx, currentitemoptionorgno
dim changedindex
dim prevregno

'==============================================================================
'// �ɼǺ��� �±�ȯ�� ��� ������ǰ����
prevregno = 0

For i = 0 To UBound(rsOption,2)
	if (rsOption(0,i) = ojumunDetail.FJumunDetail.Fitemoption) then
		if (rsOption(10,i) <> 0) then
			prevregno = rsOption(10,i)
		end if
	end if
Next

%>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script type="text/javascript">
window.resizeTo(1400,800);
var oldConfirmDate = "";
var oldBeasongDate = "";
function CheckConfirmDate(comp){
    if (comp.value==""){
        oldConfirmDate = comp.form.upcheconfirmdate.value;
        oldBeasongDate = comp.form.beasongdate.value;
        comp.form.upcheconfirmdate.value = "";
    }else{
        if (oldConfirmDate!=""){
            comp.form.upcheconfirmdate.value = oldConfirmDate;
        }

        if (oldBeasongDate!=""){
            comp.form.beasongdate.value = oldBeasongDate;
        }
    }
}

function EditDetail(detailidx,mode,comp){
    var frm = document.frm;

	if (mode=="buycash"){
		if (!IsDigit(comp.value)){
			alert('���԰��� ���ڸ� �����մϴ�.');
			comp.focus();
			return;
		}
	}else if(mode=="isupchebeasong"){
	    if (frm.isupchebeasong.value=="Y"){
	        if (frm.omwdiv.value!="U"){
	            alert('���Ա��а� ��۱����� ��ġ���� �ʽ��ϴ�.');
	            return;
	        }
	    }else{
	        if (frm.omwdiv.value=="U"){
	            alert('���Ա��а� ��۱����� ��ġ���� �ʽ��ϴ�.');
	            return;
	        }
	    }

        if (frm.omwdiv.value=="U"){
            if ((frm.odlvType.value=="1")||(frm.odlvType.value=="4")){
                alert('���Ա��а� ��۱����� ��ġ���� �ʽ��ϴ�.');
	            return;
            }
        }else{
            if ((frm.odlvType.value!="1")&&(frm.odlvType.value!="4")){
                alert('���Ա��а� ��۱����� ��ġ���� �ʽ��ϴ�.');
	            return;
            }
        }


    }else if(mode=="songjangdiv"){


	}else if(mode=="currstate"){


    }else if(mode=="songjangdiv"){
        if (frm.songjangdiv.value.length<1){
            alert('�ù�縦 �����ϼ���.');
			frm.songjangdiv.focus();
			return;
        }

        if (!IsDigit(frm.songjangno.value)){
			alert('������ȣ�� ���ڴ� �����մϴ�.');
			frm.songjangdiv.focus();
			return;
		}
	}else if (mode=="requiredetail"){

	}else if (mode=="itemno"){

	}else if (mode=="itemOption"){
		var arr = comp.value.split("|");
	    if (frm.preItemOption.value==arr[0])
	    {
			alert("��������۰� ������ �ɼ��Դϴ�. �����Ͻ� �� �����ϴ�.");
			return;
	    }
	}else{
		return;
	}

	frm.mode.value=mode;

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

// ============================================================================
function EditItemOption(){
    var frm = document.frm;

	if (frm.contents_jupsu.value == "") {
		alert("������ �ɼ��� �����ϼ���.");
		return;
	}

<% if ((ojumunDetail.FJumunDetail.Fisupchebeasong = "Y") and (ojumunDetail.FJumunDetail.FcurrState = "3")) then %>
	// ��ü���, ��ǰ�غ� ����
	if (confirm('��ü����̸鼭 ��ǰ�غ� �����Դϴ�.\n\n���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
<% elseif (ojumunDetail.FJumunDetail.FcurrState = "7") then %>
	// ��ǰ��� ����
	alert('��ǰ��� �����Դϴ�. ��Ƽ�忡�� �����ϼ���.');
<% else  %>
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
<% end if %>
}

function ForceEditItemOption(){
    var frm = document.frm;

	if (frm.contents_jupsu.value == "") {
		alert("������ �ɼ��� �����ϼ���.");
		return;
	}

	if (confirm('�����ɼǺ��� �Ͻðڽ��ϱ�?')){
		frm.forceedit.value="Y";
		frm.submit();
	}
}



// ============================================================================
function EditItemRestoreCancel(){
    var frm = document.frm;

<% if ((ojumunDetail.FJumunDetail.Fisupchebeasong = "Y") and (ojumunDetail.FJumunDetail.FcurrState = "3")) then %>
	// ��ü���, ��ǰ�غ� ����
	if (confirm('��ü����̸鼭 ��ǰ�غ� �����Դϴ�.\n\n���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
<% elseif (ojumunDetail.FJumunDetail.FcurrState = "7") then %>
	// ��ǰ��� ����
	alert('��ǰ��� �����Դϴ�. ��Ƽ�忡�� �����ϼ���.');
<% else  %>
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.title.value = "��ǰ�������ȭ";

		var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n�� �ɼ���\n";
		str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " �� ��һ��¸� ����ȭ ��û";

		frm.contents_jupsu.value = str;

		frm.mode.value="RestoreCancel";
		frm.submit();
	}
<% end if %>
}

function ForceEditItemRestoreCancel(){
    var frm = document.frm;

	if (confirm('��������ȭ �Ͻðڽ��ϱ�?')){
		frm.title.value = "��ǰ�������ȭ";

		var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n�� �ɼ���\n";
		str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " �� ��һ��¸� ����ȭ ��û";


		frm.mode.value="RestoreCancel";
		frm.forceedit.value="Y";
		frm.submit();
	}
}



// ============================================================================
function EditItemCancel(){
    var frm = document.frm;

<% if ((ojumunDetail.FJumunDetail.Fisupchebeasong = "Y") and (ojumunDetail.FJumunDetail.FcurrState = "3")) then %>
	// ��ü���, ��ǰ�غ� ����
	if (confirm('��ü����̸鼭 ��ǰ�غ� �����Դϴ�.\n\n���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
<% elseif (ojumunDetail.FJumunDetail.FcurrState = "7") then %>
	// ��ǰ��� ����
	alert('��ǰ��� �����Դϴ�. ��Ƽ�忡�� �����ϼ���.');
<% else  %>
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.title.value = "��ǰ���";

		var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n�� �ɼ���\n";
		str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " �� ��� ��û";

		frm.contents_jupsu.value = str;

		frm.mode.value="Cancel";
		frm.submit();
	}
<% end if %>
}

function ForceEditItemCancel(){
    var frm = document.frm;

	if (confirm('��������ȭ �Ͻðڽ��ϱ�?')){
		frm.title.value = "��ǰ���";

		var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n�� �ɼ���\n";
		str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " �� ��� ��û";


		frm.mode.value="Cancel";
		frm.forceedit.value="Y";
		frm.submit();
	}
}



// ============================================================================
function EditItemNo(){
    var frm = document.frm;

<% if ((ojumunDetail.FJumunDetail.Fisupchebeasong = "Y") and (ojumunDetail.FJumunDetail.FcurrState = "3")) then %>
	// ��ü���, ��ǰ�غ� ����
	if (confirm('��ü����̸鼭 ��ǰ�غ� �����Դϴ�.\n\n���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
<% elseif (ojumunDetail.FJumunDetail.FcurrState = "7") then %>
	// ��ǰ��� ����
	alert('��ǰ��� �����Դϴ�. ��Ƽ�忡�� �����ϼ���.');
<% else  %>
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.title.value = "��ǰ��������";

		var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n�� �ɼ���\n";
		str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " �� ������ " + frm.preItemNo.value + " ���� " + frm.itemno.value + " �� ���� ��û";

		frm.contents_jupsu.value = str;

		frm.mode.value="EditItemNo";
		frm.submit();
	}
<% end if %>
}

function ForceEditItemNo(){
    var frm = document.frm;

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.title.value = "��ǰ��������";

		var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n�� �ɼ���\n";
		str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " �� ������ " + frm.preItemNo.value + " ���� " + frm.itemno.value + " �� ���� ��û";

		frm.mode.value="EditItemNo";
		frm.forceedit.value="Y";
		frm.submit();
	}
}



// ============================================================================
function ChangeJupsucontents(){
    var frm = document.frm;

	var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= Replace(ojumunDetail.FJumunDetail.Fitemname,CHR(34),"") %>\n�� �ɼ���\n";
	str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " ����\n";


	str = str + eval("frm.itemOption" + frm.itemOption.value).value + " �� �����û";

	frm.contents_jupsu.value = str;
}



// ============================================================================
// �ɼǺ� ���� �ڵ�����(���̳ʽ� �ԷºҰ�)
// ============================================================================
function CheckItemOptionNoCount(changedindex){
    var frm = document.frm;
    var i;

	totalcount = 0;
	maxcount = parseInt(frm.currentitemoptionorgno.value);

	// for (i = 0; i < parseInt(frm.itemoptioncount.value); i++) {
	for (i = 0; i < parseInt(frm.itemoptionno.length); i++) {
		if ((frm.itemoptionno[i].value.length < 1) || (frm.itemoptionno[i].value*0 != 0)) {
			alert('������ ���ڸ� �Է��ϼ���.');
			return;
		}

		if (frm.itemoptionno[i].value*1 < 0) {
			alert('������ ���̳ʽ��� �Է��� �� �����ϴ�.');
			return;
		}

		if (i != changedindex) {
			maxcount = maxcount - parseInt(frm.itemoptionno[i].value);
		}

		if (i != parseInt(frm.currentitemoptionidx.value)) {
			totalcount = totalcount + parseInt(frm.itemoptionno[i].value);
		}
	}

	if ((parseInt(frm.currentitemoptionorgno.value) - totalcount) < 0) {
		alert('���氡���� ������ �ʰ��Ͽ����ϴ�.');
		frm.itemoptionno[changedindex].value = maxcount;
		return;
	}

	frm.itemoptionno[frm.currentitemoptionidx.value].value = parseInt(frm.currentitemoptionorgno.value) - totalcount;
}

// �ɼǺ��� �ֹ�����
function SaveItemOptionNo(){
    var frm = document.frm;

	if (parseInt(frm.currentitemoptionorgno.value) == parseInt(frm.itemoptionno[parseInt(frm.currentitemoptionidx.value)].value)) {
		alert('������ ������ 0�Դϴ�.');
		return;
	}

	if (frm.gubun01.value == "") {
		alert("���������� �����ϼ���.");
		return;
	}

<% if ((ojumunDetail.FJumunDetail.Fisupchebeasong = "Y") and (ojumunDetail.FJumunDetail.FcurrState = "3")) then %>
	// ��ü���, ��ǰ�غ� ����
	if (confirm('��ü����̸鼭 ��ǰ�غ� �����Դϴ�.\n\���� �Ͻðڽ��ϱ�?') != true){
		return;
	}
<% elseif (ojumunDetail.FJumunDetail.FcurrState = "7") then %>
	// ��ǰ��� ����
	alert('��ǰ��� �����Դϴ�. ��Ʈ�忡�� �����ϼ���.');
	return;
<% end if %>
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.title.value = "��ǰ�ɼǺ���";

		// var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= ojumunDetail.FJumunDetail.Fitemname %>\n�� �ɼ���\n";
		// str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " �� ������ " + frm.preItemNo.value + " ���� " + frm.itemno.value + " �� ���� ��û";

		frm.contents_jupsu.value = "str";

		frm.mode.value="EditItemNoPart";
		frm.submit();
	}
}

// �ɼǺ��� �ֹ�����
function ForceSaveItemOptionNo(){
    var frm = document.frm;

	if (parseInt(frm.currentitemoptionorgno.value) == parseInt(frm.itemoptionno[parseInt(frm.currentitemoptionidx.value)].value)) {
		alert('������ ������ 0�Դϴ�.');
		return;
	}

	if (frm.gubun01.value == "") {
		alert("���������� �����ϼ���.");
		return;
	}

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.title.value = "��ǰ�ɼǺ���";

		// var str = "[<%= ojumunDetail.FJumunDetail.Fitemid %>] <%= ojumunDetail.FJumunDetail.Fitemname %>\n�� �ɼ���\n";
		// str = str + "[" + frm.preItemOption.value + "] " + frm.preItemOptionName.value + " �� ������ " + frm.preItemNo.value + " ���� " + frm.itemno.value + " �� ���� ��û";

		frm.contents_jupsu.value = "str";

		frm.forceedit.value="Y";
		frm.mode.value="EditItemNoPart";
		frm.submit();
	}
}

var IsPossibleModifyCSMaster = true;
var IsPossibleModifyItemList = true;

// �ɼǺ��� �±�ȯ
function SaveChangeItemOptionNo(){
    var frm = document.frm;

	if (parseInt(frm.currentitemoptionorgno.value) == parseInt(frm.itemoptionno[parseInt(frm.currentitemoptionidx.value)].value)) {
		alert('������ ������ 0�Դϴ�.');
		return;
	}

	if (frm.gubun01.value == "") {
		alert("���������� �����ϼ���.");
		return;
	}

<% if (ojumunDetail.FJumunDetail.FcurrState < "7") then %>
	// ��ǰ��� ����
	alert('��ǰ��� ���� ��ǰ�Դϴ�. ��ȯ(�ɼǺ���)�� �� �����ϴ�.');
	return;
<% end if %>
	if (confirm('��ȯ ����(�ɼǺ���) �Ͻðڽ��ϱ�?')){
		frm.title.value = "��ȯ���(�ɼǺ���)";

		frm.contents_jupsu.value = "str";

		<% if (ojumunDetail.FJumunDetail.Fisupchebeasong = "Y") then %>
			frm.requiremakerid.value="<%= ojumunDetail.FJumunDetail.Fmakerid %>";
		<% end if %>

		frm.mode.value="ChangeEditItemNoPart";
		frm.submit();
	}
}

// itemoptioncount

</script>
<script language='javascript' SRC="/js/ajax.js"></script>
<script language='javascript' SRC="/cscenter/js/newcsas.js"></script>

<form name="frm" method="post" action="/cscenter/ordermaster/orderdetail_process.asp" style="margin:0px;">
<input type="hidden" name="detailidx" value="<%= ojumunDetail.FJumunDetail.Fdetailidx %>">
<input type="hidden" name="orderserial" value="<%= ojumunDetail.FJumunDetail.FOrderSerial %>">
<input type="hidden" name="mode" value="itemOption">
<input type="hidden" name="forceedit" value="N">
<input type="hidden" name="requiremakerid" value="">
<input type="hidden" name="itemId" value="<%= ojumunDetail.FJumunDetail.Fitemid %>">
<input type="hidden" name="preItemOption" value="<%= ojumunDetail.FJumunDetail.FitemOption %>">
<input type="hidden" name="preItemOptionName" value="<%= Replace(ojumunDetail.FJumunDetail.FitemOptionName, ",", "") %>">
<input type="hidden" name="preItemNo" value="<%= ojumunDetail.FJumunDetail.Fitemno - prevregno %>">
<input type="hidden" name="title" value="��ǰ�ɼǺ���">
<input type="hidden" name="contents_jupsu" value="">
<input type="hidden" name="contents_finish" value="���������� ó���Ǿ����ϴ�.">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ֹ����������� ����</b>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">IDX</td>
		<td><%= ojumunDetail.FJumunDetail.Fdetailidx %></td>
		<td width="110" rowspan="4"><img src="<%= ojumunDetail.FJumunDetail.FImageList %>"></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">�귣�� ID</td>
		<td><%= ojumunDetail.FJumunDetail.Fmakerid %></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemid %></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemname %></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">�����ɼ�</td>
		<td>[<%= ojumunDetail.FJumunDetail.Fitemoption %>] <%= ojumunDetail.FJumunDetail.Fitemoptionname %></td>
		<td></td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">��һ���</td>
		<td><%= ojumunDetail.FJumunDetail.Fcancelyn %></td>
		<td>

		</td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

<% if ojumunDetail.FJumunDetail.Fitemoption <> "0000" Then %>
	<%
	currentitemoptionidx = 0
	changedindex = 0
	%>
	<tr bgcolor="#FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">
			����ɼ�
		</td>
		<td>
			<%= "[" & ojumunDetail.FJumunDetail.FitemOption & "] " & ojumunDetail.FJumunDetail.FitemOptionName %>
		</td>
		<input type=hidden name=itemoptioncode value="<%= ojumunDetail.FJumunDetail.FitemOption %>">
		<input type=hidden name=ItemOptionName value="<%= Replace(ojumunDetail.FJumunDetail.FitemOptionName, ",", "") %>">
		<td>
			<input type="text" class="text_ro" name="itemoptionno" value="<%= (ojumunDetail.FJumunDetail.Fitemno - prevregno) %>" size="3" maxlength="9" readonly> ��(<%= ojumunDetail.FJumunDetail.Fitemno %>��)

			<% if (prevregno <> 0) then %>
				<font color=red>(������ǰ : <%= prevregno %> ��)</font>
			<% end if %>
		</td>
	</tr>
	<% For i = 0 To UBound(rsOption,2) %>
		<%
		If rsOption(2,i) = "N" Or ( (rsOption(3,i)="Y") and (rsOption(4,i) - rsOption(5,i) < 1) ) Then
			itemStatus = "�Ǹ�����"
		ElseIf rsOption(3,i)="Y" Then
			If ( rsOption(4,i) - rsOption(5,i) ) < 1 Then
				itemStatus = "����:0"
			Else
				itemStatus = "����:" & ( rsOption(4,i) - rsOption(5,i) )
			End If
		ElseIf rsOption(6,i) <> 0 Then
			itemStatus = "���ֹ�:" & rsOption(6,i)
		Else
			itemStatus = ""
		End If

		If rsOption(8,i) = "N" Then
			If itemStatus <> "" Then
				itemStatus = itemStatus & ", " & "������"
			else
				itemStatus = "������"
			end if
		End If

		If itemStatus <> "" Then
			itemStatus = " (" & itemStatus & ")"
		End If

		optionText = "[" & rsOption(0,i) & "] " & rsOption(1,i) & itemStatus

		%>




        <% ''rw rsOption(0,i) & ".." & ojumunDetail.FJumunDetail.Fitemoption %>
		<% if (rsOption(0,i) = ojumunDetail.FJumunDetail.Fitemoption) then %>
			<!-- �ɼǸ�Ͽ��� �����ϴ� ��� �ֹ������Ͽ��� ���� �ֹ����� �����´�. -->
		<% elseif (rsOption(7,i) = "F") then %>
			<% changedindex = changedindex + 1 %>
	<tr bgcolor="#FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">
			����Ұ��ɼ�(<%= i+1 %>)
		</td>
		<td>
			<%=optionText%><font color=red>(�ɼǰ� �ٸ�)</font>
		</td>
		<input type=hidden name=itemoptioncode value="<%= rsOption(0,i) %>">
		<input type=hidden name=ItemOptionName value="<%= Replace(rsOption(1,i), ",", "") %>">
		<td width="110">
			<input type="text" class="text_ro" name="itemoptionno" value="0" size="3" maxlength="9" onKeyUp="CheckItemOptionNoCount(<%= (changedindex) %>)" readonly> ��
		</td>
	</tr>
		<% else %>
			<% changedindex = changedindex + 1 %>
	<tr bgcolor="#FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">
			���氡�ɿɼ�(<%= i+1 %>)
		</td>
		<td>
			<%=optionText%>
		</td>
		<input type=hidden name=itemoptioncode value="<%= rsOption(0,i) %>">
		<input type=hidden name=ItemOptionName value="<%= Replace(Replace(rsOption(1,i), ",", ""), "Perl ", "Perl") %>">
		<td width="110">
			<input type="text" class="text" name="itemoptionno" value="0" size="3" maxlength="9" onKeyUp="CheckItemOptionNoCount(<%= (changedindex) %>)"> ��
		</td>
	</tr>
		<% end If %>
	<% Next %>
	<input type=hidden name=itemoptioncount value="<%= (UBound(rsOption,2) + 1) %>">
	<input type=hidden name=currentitemoptionidx value="<%= currentitemoptionidx %>">
	<input type=hidden name=currentitemoptionorgno value="<%= ojumunDetail.FJumunDetail.Fitemno - prevregno %>">
<% end If %>

	<tr bgcolor="#FFFFFF">
		<td width="100" height=35 bgcolor="<%= adminColor("tabletop") %>">
			��������
		</td>
		<td colspan=2>
                <input type="hidden" name="gubun01" value="">
                <input type="hidden" name="gubun02" value="">
                <input class="text_ro" type="text" name="gubun01name" value="" size="16" Readonly >
                &gt;
                <input class="text_ro" type="text" name="gubun02name" value="" size="16" Readonly >
                <input class="csbutton" type="button" value="����" onClick="divCsAsGubunSelect(frm.gubun01.value, frm.gubun02.value, frm.gubun01.name, frm.gubun02.name, frm.gubun01name.name, frm.gubun02name.name,'frm','causepop');">
                <div id="causepop" style="position:absolute;"></div>

                <!-- �Ϻ� ���� �̸� ǥ�� -->
                <%
                '��������
				'select top 100 m.comm_cd, m.comm_name, d.comm_cd, d.comm_name
				'from
				'	db_cs.dbo.tbl_cs_comm_code m
				'	left join db_cs.dbo.tbl_cs_comm_code d
				'	on
				'		m.comm_cd = d.comm_group
				'where
				'	1 = 1
				'	and m.comm_group = 'Z020'
				'	and m.comm_isdel <> 'Y'
				'	and d.comm_isdel <> 'Y'
				'order by m.comm_cd, d.comm_cd
                %>
                [<a href="javascript:selectGubun('C004','CD01','����','�ܼ�����','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">�ܼ�����</a>]
                [<a href="javascript:selectGubun('C004','CD05','����','ǰ��','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">ǰ��</a>]
                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">��ǰ�ҷ�</a>]
                [<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">��Ÿ</a>]
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height=35>
		<td colspan="3" align="center">
<% if ojumunDetail.FJumunDetail.Fcancelyn <> "Y" then %>
			<input type="button" class="button" value="�ɼǺ���" onclick="javascript:SaveItemOptionNo()">
			<% if (C_ADMIN_AUTH or C_CSPowerUser) then %>
			<!-- ��Ʈ�� �̻� -->
		    <input type="button" class="button" value="��������" onclick="javascript:ForceSaveItemOptionNo()">
			<% end if %>
			<input type="button" class="button" value="�ɼǺ��� �±�ȯ" onclick="javascript:SaveChangeItemOptionNo()">
<% else %>
			��ҵ� ��ǰ�� �������� �Ұ�
<% end if %>
		</td>
	</tr>
</table>
</form>
<div>
* �ɼǺ����� <font color=red>�ɼǰ�</font>�� ������ �ɼǻ�ǰ�� �����մϴ�.<br>
* �ֹ���� �ɼǰ��ݿ� ������� ���� ��ǰ���� ���� �ɼǰ������� ���մϴ�.<br>
* ��ǰ��������(�ǸŰ�,���԰� ��)�� �ֹ���� ������ �����˴ϴ�.<br>
</div>

<%
set ojumun       = Nothing
set ojumunDetail = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->