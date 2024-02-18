<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : cs����
' Hieditor : 2009.04.17 �̻� ����
'			 2016.07.19 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%
dim searchfield, userid, orderserial, username, userhp, etcfield, etcstring, yyyy1,yyyy2,mm1,mm2,dd1,dd2, jumundiv, jumunsite, jumunitem
dim checkYYYYMMDD, checkJumunDiv, checkJumunSite, checkJumunItem, checkSongjangno, research, AlertMsg, v6MonthAgo, outmallorderserial
dim songjangno, useAsterisk, nowdate, searchnextdate, page, ojumun, ix,iy, ResultOneOrderserial, ipkumdiv, checkIpkumdiv
	searchfield = request("searchfield")
	userid 		= requestCheckvar(request("userid"),32)
	orderserial = requestCheckvar(request("orderserial"),32)
	username 	= requestCheckvar(request("username"),32)
	userhp 		= requestCheckVarNoTrim(request("userhp"),32)
	etcfield 	= requestCheckvar(request("etcfield"),32)
	etcstring 	= requestCheckvar(request("etcstring"),50)
	songjangno 	= requestCheckvar(request("songjangno"),32)
	checkYYYYMMDD = request("checkYYYYMMDD")
	checkJumunDiv = request("checkJumunDiv")
	checkJumunSite = request("checkJumunSite")
	checkJumunItem = request("checkJumunItem")
	checkSongjangno = request("checkSongjangno")
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	jumundiv = request("jumundiv")
	jumunsite = request("jumunsite")
	jumunitem = requestCheckvar(request("jumunitem"), 32)
	ipkumdiv = requestCheckvar(request("ipkumdiv"), 1)
	checkIpkumdiv = requestCheckvar(request("checkIpkumdiv"), 1)
	research = request("research")
	v6MonthAgo = request("sixmonthago")
	page = request("page")

if (page="") then page=1
useAsterisk = True
if (research="") and (checkYYYYMMDD="") then checkYYYYMMDD="Y"

if (userid = "") then
	'// ���̵� �˻��ø� ��ǰ�˻� ����
	checkJumunItem = ""
end if

if (searchfield <> "etcfield") or (etcfield <> "02") or (etcstring = "") then
	'// �����θ� �ʼ�
	checkSongjangno = ""
end if

'2017-03-08 ������..Len(orderserial) < 10���� Len(orderserial) <= 10���� ���� | '2018-03-14 ������..orderserial) >= 12�� ����
'2020-04-09 ������..(LEFT(orderserial, 1)= "Y") ���� �߰�..yes24 ���޸� ����
if (Len(orderserial) >= 12) or (Len(orderserial) <= 10) OR (LEFT(orderserial, 1)= "Y") OR (LEFT(orderserial, 1)= "X") Then
	'// ���޸� �ֹ���ȣ -> �ֹ���ȣ
	outmallorderserial = orderserial
	Call GetOrderserialWithOutmallOrderserial(outmallorderserial, orderserial)
	if (orderserial = "") then
		orderserial = outmallorderserial
	end if
end if

''�⺻ N��. ����Ʈ üũ
if (yyyy1="") then
    nowdate = Left(CStr(dateadd("m",-1,now())),10)
	yyyy1   = Left(nowdate,4)
	mm1     = Mid(nowdate,6,2)
	dd1     = Mid(nowdate,9,2)

	nowdate = Left(CStr(now()),10)
	yyyy2   = Left(nowdate,4)
	mm2     = Mid(nowdate,6,2)
	dd2     = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2,mm2,dd2),1)),10)

set ojumun = new COrderMaster
ojumun.FPageSize = 10
ojumun.FCurrPage = page

if (checkYYYYMMDD="Y") and ((orderserial = "") or (searchfield <> "orderserial")) then
	'// �ֹ���ȣ ������ �Ⱓ �˻����� ����(2013-11-11 skyer9)
	ojumun.FRectRegStart = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
	ojumun.FRectRegEnd = searchnextdate
end if

if (checkJumunDiv = "Y") then
    if (jumundiv="flowers") then
        ojumun.FRectIsFlower = "Y"
    elseif (jumundiv="minus") then
        ojumun.FRectIsMinus = "Y"
    elseif (jumundiv="foreign") then
        ojumun.FRectIsForeign = "Y"
    elseif (jumundiv="foreigndirect") then
        ojumun.FRectIsForeignDirect = "Y"
    elseif (jumundiv="quick") then
        ojumun.FRectIsQuick = "Y"
    elseif (jumundiv="sendGift") then
        ojumun.FRectIsSendGift = "Y"
    end if
end if

if (checkJumunSite = "Y") then
	ojumun.FRectExtSiteName = jumunsite
end if

if (checkJumunItem = "Y") then
	ojumun.FRectJumunItem = jumunitem
end if

if (checkSongjangno = "Y") then
	ojumun.FRectSongjangno = songjangno
end if

if (checkIpkumdiv = "Y") then
	ojumun.FRectIpkumdiv = ipkumdiv
end if

if (searchfield = "orderserial") then
	'�ֹ���ȣ
	ojumun.FRectOrderSerial = orderserial
elseif (searchfield = "userid") then
	'�����̵�
	ojumun.FRectUserID = userid
	useAsterisk = False
elseif (searchfield = "username") then
	'�����ڸ�
	ojumun.FRectBuyname = username
elseif (searchfield = "userhp") then
	'�������ڵ���
	ojumun.FRectBuyHp = userhp
	useAsterisk = False
elseif (searchfield = "etcfield") then
	'��Ÿ����
	if etcfield="01" then
		ojumun.FRectBuyname = etcstring
	elseif etcfield="02" then
		ojumun.FRectReqName = etcstring
	elseif etcfield="03" then
		ojumun.FRectUserID = etcstring
	elseif etcfield="04" then
		ojumun.FRectIpkumName = etcstring
	elseif etcfield="06" then
		ojumun.FRectSubTotalPrice = etcstring
	elseif etcfield="07" then
		ojumun.FRectBuyPhone = etcstring
		useAsterisk = False
	elseif etcfield="08" then
		ojumun.FRectReqHp = etcstring
		useAsterisk = False
	elseif etcfield="09" then
		ojumun.FRectReqSongjangNo = etcstring
	elseif etcfield="10" then
		ojumun.FRectReqPhone = etcstring
		useAsterisk = False
	elseif etcfield="11" then
		ojumun.FRectbuyemail = etcstring
	elseif etcfield="12" then
		ojumun.FRectreqemail = etcstring
	elseif etcfield="20" then
		ojumun.FRectpaygatetid = etcstring
	end if
end if

If v6MonthAgo = "o" Then
	ojumun.FRectOldOrder = "on"
End If

''�˻����� ������ �ֱ� N�� �˻�
ojumun.QuickSearchOrderList

if (ojumun.FResultCount<1) and ((searchfield = "userhp") or ((searchfield = "etcfield") and (etcfield="08"))) then
    '// �˻������� ������ �ڵ��� or ������ �ڵ��� �� ���
    if (searchfield = "userhp") then
        '// ������ �ڵ���
        if (UBound(Split(userhp, "-")) = 0) then
            ojumun.FRectBuyHp = fnToPhoneNumber(userhp)
            ojumun.QuickSearchOrderList
            if (ojumun.FResultCount > 0) then
                userhp = ojumun.FRectBuyHp
            end if
        end if
    else
        '// ������ �ڵ���
        if (UBound(Split(etcstring, "-")) = 0) then
            ojumun.FRectReqHp = fnToPhoneNumber(etcstring)
            ojumun.QuickSearchOrderList
            if (ojumun.FResultCount > 0) then
                etcstring = ojumun.FRectReqHp
            end if
        end if
    end if
end if

'' ���� 6���� ���� ���� �˻�
if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderList

    if (ojumun.FResultCount>0) then
        AlertMsg = "6���� ���� �ֹ��Դϴ�."
    end if
end if

'' �˻������ 1���ϴ� ������ �ڵ����� �Ѹ�
ResultOneOrderserial = ""
if (ojumun.FResultCount=1) then
    ResultOneOrderserial = ojumun.FItemList(0).FOrderSerial
end if
%>
<link rel="stylesheet" href="/cscenter/css/cs.css" type="text/css">
<style>
.csH15 { line-height: 15px; }
</style>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>

function copyClipBoard(itxt) {
	//if( window.clipboardData && clipboardData.setData ){
	//	clipboardData.setData("Text", itxt);
	//}
	//if (itxt.length<1){ return; }

	var posSpliter = itxt.indexOf("|");

	try{
	    parent.callring.frm.orderserial.value=itxt.substring(0,posSpliter);
	    parent.callring.frm.userid.value=itxt.substring(posSpliter+1,255);
	}catch(ignore){

	}
}

function SearchByOrderserial(iorderserial){
	frm.searchfield[0].checked = true;
	frm.orderserial.value = iorderserial;
	frm.submit();
}

function SearchByUserID(iuserid){
	frm.searchfield[1].checked = true;
	frm.userid.value = iuserid;
	frm.submit();
}

function SearchByPhoneNumber(iphoneNumber){
    var isCell = false;
    var l3Str = iphoneNumber.substring(0,3);

    isCell = ((l3Str=="010")||(l3Str=="011")||(l3Str=="016")||(l3Str=="017")||(l3Str=="018")||(l3Str=="019"))?true:false;

    if (isCell){
        //frm.searchfield[3].checked = true;
	    //frm.userhp.value = iphoneNumber;
	    //frm.submit();


	    frm.searchfield[4].checked = true;
        frm.etcfield.value = "08";				//������ �ڵ���
	    frm.etcstring.value = iphoneNumber;
	    frm.submit();
    }else{
        frm.searchfield[4].checked = true;
        frm.etcfield.value = "10";				//������ ��ȭ
	    frm.etcstring.value = iphoneNumber;
	    frm.submit();
    }
}

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="/admin/ordermaster/viewordermaster.asp"
	frm.submit();

}

function jsGetFrameDocument(obj) {
	if (obj.contentWindow) {
		return obj.contentWindow;
	} else if (obj.contentDocument) {
		return obj.contentDocument.document;
	} else if (obj.location) {
		return obj;
	}
}

function GotoOrderDetail(orderserial) {
	var ifrm = jsGetFrameDocument(parent.detailFrame);
	ifrm.location.href = "ordermaster_detail.asp?orderserial=" + orderserial;
}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function EnDisabledDateBox(comp){
	document.frm.yyyy1.disabled = !comp.checked;
	document.frm.yyyy2.disabled = !comp.checked;
	document.frm.mm1.disabled = !comp.checked;
	document.frm.mm2.disabled = !comp.checked;
	document.frm.dd1.disabled = !comp.checked;
	document.frm.dd2.disabled = !comp.checked;
}

function ChangeCheckbox(frmname, frmvalue) {
    for (var i = 0; i < frm.elements.length; i++) {
        if (frm.elements[i].type == "radio") {
            if ((frm.elements[i].name == frmname) && (frm.elements[i].value == frmvalue)) {
                frm.elements[i].checked = true;
            }
        }
    }
}

function FocusAndSelect(frm, obj){
    ChangeFormBgColor(frm);

    obj.focus();
    obj.select();
}

function ChangeFormBgColor(frm) {
    // style='background-color:#DDDDFF'
    var radioselected = false;
    var checkboxchecked = false;
    var ischecked = false;

    for (var i = 0; i < frm.elements.length; i++) {
        if (frm.elements[i].type == "radio") {
            ischecked = frm.elements[i].checked;
        }

        if (frm.elements[i].type == "checkbox") {
            ischecked = frm.elements[i].checked;
        }

        if (frm.elements[i].type == "text") {
            if (ischecked == true) {
				$( frm.elements[i] ).removeClass("csDefBg").addClass("csSelBg");
            } else {
                $( frm.elements[i] ).removeClass("csSelBg").addClass("csDefBg");
            }
        }

        if (frm.elements[i].type == "select-one") {
            if (ischecked == true) {
                $( frm.elements[i] ).removeClass("csDefBg").addClass("csSelBg");
            } else {
                $( frm.elements[i] ).removeClass("csSelBg").addClass("csDefBg");
            }
        }
    }
}

// tr ���󺯰�
var pre_selected_row = null;
var pre_selected_row_color = null;

function ChangeColor(e, selcolor, defcolor){
	if (pre_selected_row_color != null) {
	        pre_selected_row.bgColor = pre_selected_row_color;
        }
        pre_selected_row = e;
        pre_selected_row_color = defcolor;
        e.bgColor = selcolor;
}

function MonthDiff(d1, d2) {
	d1 = d1.split("-");
	d2 = d2.split("-");

	d1 = new Date(d1[0], d1[1] - 1, d1[2]);
	d2 = new Date(d2[0], d2[1] - 1, d2[2]);

	var d1Y = d1.getFullYear();
	var d2Y = d2.getFullYear();
	var d1M = d1.getMonth();
	var d2M = d2.getMonth();

	return (d2M+12*d2Y)-(d1M+12*d1Y);
}

function CheckSubmit(frm) {
	if (frm.sixmonthago.checked == true) {
		// 6�������� �ֹ��� ���, �ֹ���ȣ&���̵� �̿��� �˻�(�����ڸ� ��)�� �ݵ�� �Ⱓ�� �����ؾ� �մϴ�.
		if ((frm.searchfield[0].checked == false) && (frm.searchfield[1].checked == false)) {
			if (frm.checkYYYYMMDD.checked != true) {
				alert("6�������� �ֹ��� �˻��� ���\n\n�����ڸ� ���� �˻��Ϸ��� �ݵ�� �Ⱓ(�ֹ���,�ִ� �Ѵ�)�� �����ؾ� �մϴ�.");
				return;
			}

			if ((CheckDateValid(frm.yyyy1.value, frm.mm1.value, frm.dd1.value) == true) && (CheckDateValid(frm.yyyy2.value, frm.mm2.value, frm.dd2.value) == true)) {
				if (MonthDiff(frm.yyyy1.value + "-" + frm.mm1.value + "-" + frm.dd1.value, frm.yyyy2.value + "-" + frm.mm2.value + "-" + frm.dd2.value) >= 1) {
					alert("�˻��Ⱓ�� �ִ� 1���������Դϴ�.");
					return;
				}
			} else {
				return;
			}
		}
	}

	frm.submit();
}

</script>


<!-- ǥ ��ܹ� ����-->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="F4F4F4">
	<tr height="50">
        <td>
    		<input type="radio" name="searchfield" value="orderserial" <% if searchfield="orderserial" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.orderserial)"> �ֹ���ȣ
    		<input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="13" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'orderserial'); FocusAndSelect(frm, frm.orderserial);">

    		<input type="radio" name="searchfield" value="userid" <% if searchfield="userid" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userid)"> ���̵�
    		<input type="text" class="text" name="userid" value="<%= userid %>" size="12" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'userid'); FocusAndSelect(frm, frm.userid);">

    		<input type="radio" name="searchfield" value="username" <% if searchfield="username" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.username)"> �����ڸ�
    		<input type="text" class="text" name="username" value="<%= username %>" size="8" maxlength="32" onKeyPress="if (event.keyCode == 13) CheckSubmit(document.frm);" onFocus="ChangeCheckbox('searchfield', 'username'); FocusAndSelect(frm, frm.username);">

    		<input type="radio" name="searchfield" value="userhp" <% if searchfield="userhp" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.userhp)"> �������ڵ���
    		<input type="text" class="text" name="userhp" value="<%= userhp %>" size="14" maxlength="14" onKeyPress="if (event.keyCode == 13) CheckSubmit(document.frm);" onFocus="ChangeCheckbox('searchfield', 'userhp'); FocusAndSelect(frm, frm.userhp);">

            <input type="radio" name="searchfield" value="etcfield" <% if searchfield="etcfield" then response.write "checked" %> onClick="FocusAndSelect(frm, frm.etcstring)"> ��Ÿ����

			<select name="etcfield" class="select">
				<option value="">����</option>
				<option value="02" <% if etcfield="02" then response.write "selected" %> >�����θ�</option>
				<option value="04" <% if etcfield="04" then response.write "selected" %> >�Ա��ڸ�</option>
				<option value="06" <% if etcfield="06" then response.write "selected" %> >�����ݾ�</option>
				<option value="07" <% if etcfield="07" then response.write "selected" %> >������ ��ȭ</option>
				<option value="10" <% if etcfield="10" then response.write "selected" %> >������ ��ȭ</option>
				<option value="08" <% if etcfield="08" then response.write "selected" %> >������ �ڵ���</option>
				<option value="09" <% if etcfield="09" then response.write "selected" %> >�����ȣ(�ٹ�)</option>
				<option value="11" <% if etcfield="11" then response.write "selected" %> >�������̸���</option>
				<option value="12" <% if etcfield="12" then response.write "selected" %> >�������̸���</option>
				<option value="20" <% if etcfield="20" then response.write "selected" %> >PG��TID</option>
			</select>
    		<input type="text" class="text" name="etcstring" value="<%= etcstring %>" size="30" maxlength="50" onKeyPress="if (event.keyCode == 13) document.frm.submit();" onFocus="ChangeCheckbox('searchfield', 'etcfield'); FocusAndSelect(frm, frm.etcstring);">
    		<br>
    		<input type="checkbox" name="checkYYYYMMDD" value="Y" <% if checkYYYYMMDD="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
    		�ֹ��� : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
                <input type="checkbox" name="checkJumunDiv" value="Y" <% if checkJumunDiv="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
    		�ֹ����� :
    		<select name="jumundiv" class="select">
                <option value="">����</option>
                <option value="flowers" <% if jumundiv="flowers" then response.write "selected" %> >�ö���ֹ�</option>
				<option value="sendGift"   <% if jumundiv="sendGift"   then response.write "selected" %> >�����ϱ�</option>
                <option value="minus"   <% if jumundiv="minus"   then response.write "selected" %> >���̳ʽ�</option>
                <option value="foreign"   <% if jumundiv="foreign"   then response.write "selected" %> >�ؿܹ��</option>
				<option value="foreigndirect"   <% if jumundiv="foreigndirect"   then response.write "selected" %> >�ؿ�����</option>
				<option value="quick"   <% if jumundiv="quick"   then response.write "selected" %> >��</option>
            </select>
            <input type="checkbox" name="checkJumunSite" value="Y" <% if checkJumunSite="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
    		Ư������Ʈ : <% DrawSelectExtSiteName "jumunsite", jumunsite %>
			&nbsp;
			<input type="checkbox" name="checkJumunItem" value="Y" <% if checkJumunItem="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
			��ǰ�ڵ�/��ǰ��(���̵� �ʼ�) :
			<input type="text" class="text" name="jumunitem" value="<%= jumunitem %>" size="8" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
			&nbsp;
			<input type="checkbox" name="checkSongjangno" value="Y" <% if checkSongjangno="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
			�����ȣ(�����θ� �ʼ�) :
			<input type="text" class="text" name="songjangno" value="<%= songjangno %>" size="15" maxlength="32" onKeyPress="if (event.keyCode == 13) document.frm.submit();">

			&nbsp;
			<input type="checkbox" name="checkIpkumdiv" value="Y" <% if checkIpkumdiv="Y" then response.write "checked" %> onClick="ChangeFormBgColor(frm)">
    		�ŷ����� : <% Call DrawIpkumDivName("ipkumdiv",ipkumdiv,"") %>
        </td>
        <td align="right" valign="top">
            <input type="button" class="button_s" value="���ΰ�ħ" onclick="document.location.reload();">
            &nbsp;
            <input type="button" class="button_s" value="�˻��ϱ�" onclick="CheckSubmit(document.frm);">
            <br>
            <input type="checkbox" name="sixmonthago" value="o" <% if v6MonthAgo="o" then response.write "checked" %>>6������������
        </td>
	</tr>
</table>
</form>
<!-- ǥ ��ܹ� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" class="csH15">
    	<td width="30">����</td>
    	<td width="50">�ֹ�����</td>
    	<td width="85">���ֹ���ȣ</td>
    	<td width="60">Site</td>
		<td width="80">rdSite</td>
    	<td>UserID</td>
    	<td width="70">������</td>
    	<td width="70">������</td>

		<% if (C_InspectorUser = False) then %>
    	<td width="60">�ǸŰ�</td>
    	<td width="60">��ǰ����</td>
    	<td width="60">���ʽ�����</td>
    	<td width="50">��Ÿ����</td>
		<% end if %>

    	<td width="60">�����Ѿ�</td>
		<td width="50">���ϸ���</td>
    	<td width="60">��������</td>
    	<td width="60"><b>�ǰ�����</b></td>

    	<td width="60">�������</td>
    	<td width="50">�ŷ�����</td>
    	<td width="80">�ֹ���</td>
    	<td width="80">�Ա�Ȯ����</td>
    	<!--<td width="70">������</td>-->
    	<!-- td width="70">�����</td -->
    	<!-- td>��ü��ۿ����</td> -->
    </tr>
    <% if ojumun.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF" class="csH15">
    	<td colspan="20" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
    <% else %>

	<% for ix=0 to ojumun.FresultCount-1 %>

	<% if ojumun.FItemList(ix).IsAvailJumun then %>
	<tr align="center" bgcolor="#FFFFFF" class="a csMp csH15" onclick="ChangeColor(this,'#AFEEEE','FFFFFF'); copyClipBoard('<%= ojumun.FItemList(ix).FOrderSerial %>|<%= ojumun.FItemList(ix).FUserID %>'); GotoOrderDetail('<%= ojumun.FItemList(ix).FOrderSerial %>');">
	<% else %>
	<tr align="center" bgcolor="#EEEEEE" class="gray csMp csH15" onclick="ChangeColor(this,'#AFEEEE','EEEEEE'); copyClipBoard('<%= ojumun.FItemList(ix).FOrderSerial %>|<%= ojumun.FItemList(ix).FUserID %>'); GotoOrderDetail('<%= ojumun.FItemList(ix).FOrderSerial %>');">
	<% end if %>
		<td><font color="<%= ojumun.FItemList(ix).CancelYnColor %>"><%= ojumun.FItemList(ix).CancelYnName %></font></td>
		<td>
		    <% if (ojumun.FItemList(ix).IsForeignDeliver) then %>
		    <strong>�ؿ�</strong>
		    <% elseif (ojumun.FItemList(ix).IsQuickDeliver) then %>
		    <strong>��</strong>
		    <% elseif (ojumun.FItemList(ix).IsArmiDeliver) then %>
		    <strong>���δ�</strong>
		    <% else %>
		    <%= ojumun.FItemList(ix).GetJumunDivName %>
		    <% end if %>
		</td>
		<td>
			<a href="?searchfield=orderserial&orderserial=<%= ojumun.FItemList(ix).Forgorderserial %>"><%= ojumun.FItemList(ix).FOrgOrderSerial %></a>
    		<% if (ojumun.FItemList(ix).Forderserial <> ojumun.FItemList(ix).Forgorderserial) then %>
    			+
    		<% end if %>
		</td>
		<td><font color="<%= ojumun.FItemList(ix).SiteNameColor %>"><%= ojumun.FItemList(ix).FSitename %></font></td>
		<td><acronym title="<%= ojumun.FItemList(ix).Frdsite %>"><%= ojumun.FItemList(ix).Frdsite %></acronym></td>
		<td align="left">
		    <% if ojumun.FItemList(ix).FSitename<>"10x10" then %>
		    (<%= ojumun.FItemList(ix).FAuthCode %>)
		    <% else %>
				<% if (C_InspectorUser = False) then %>
					<a href="?searchfield=userid&userid=<%= GetUseridWithAsterisk(ojumun.FItemList(ix).FUserID, useAsterisk) %>">
						<font color="<%= getUserLevelColorByDate(ojumun.FItemList(ix).fuserlevel, Left(ojumun.FItemList(ix).FRegDate,10)) %>">
							<b><%= GetUseridWithAsterisk(ojumun.FItemList(ix).FUserID, useAsterisk) %></b>
						</font>
					</a>
				<% else %>
					<%= GetUseridWithAsterisk(ojumun.FItemList(ix).FUserID, useAsterisk) %>
				<% end if %>
		    <% end if %>
		</td>
		<td><%= GetUsernameWithAsterisk(ojumun.FItemList(ix).FBuyName, useAsterisk) %></td>
		<td><%= GetUsernameWithAsterisk(ojumun.FItemList(ix).FReqName, useAsterisk) %></td>

		<% if (C_InspectorUser = False) then %>
		<td align="right">
			<% if (ojumun.FItemList(ix).IsOldJumun <> true) then %>
				<%= FormatNumber(ojumun.FItemList(ix).FsubtotalpriceCouponNotApplied,0) %>
			<% else %>
				----
			<% end if %>
		</td>
		<td align="right">
			<% if (ojumun.FItemList(ix).IsOldJumun <> true) then %>
				<%= FormatNumber((ojumun.FItemList(ix).FsubtotalpriceCouponNotApplied - ojumun.FItemList(ix).FTotalSum),0) %><!-- ��ǰ���� ���ξ� -->
			<% else %>
				<%= FormatNumber(ojumun.FItemList(ix).FTotalSum,0) %><!-- ��ǰ���� ���밡 -->
			<% end if %>
		</td>
		<td align="right"><%= FormatNumber(ojumun.FItemList(ix).Ftencardspend,0) %></td>
		<td align="right">
		    <% if ojumun.FItemList(ix).Fallatdiscountprice<>0 then %>
		    <acronym title="<%= CHKIIF(ojumun.FItemList(ix).FAccountDiv="80","�ÿ�����","����ī������") %>"><%= FormatNumber(ojumun.FItemList(ix).Fallatdiscountprice+ ojumun.FItemList(ix).Fspendmembership,0) %></acronym>
		    <% else %>
		    <%= FormatNumber(ojumun.FItemList(ix).Fallatdiscountprice+ ojumun.FItemList(ix).Fspendmembership,0) %>
		    <% end if %>
		</td>
		<% end if %>

		<!-- �����Ѿ׿� ���ϸ��� ����(2014-02-11 skyer9) -->
		<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>" ><%= FormatNumber((ojumun.FItemList(ix).FSubTotalPrice + ojumun.FItemList(ix).Fmiletotalprice),0) %></font></td>
		<td align="right"><%= FormatNumber(ojumun.FItemList(ix).Fmiletotalprice,0) %></td>
		<td align="right"><%= FormatNumber(ojumun.FItemList(ix).FsumPaymentEtc,0) %></td>
		<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>" ><b><%= FormatNumber((ojumun.FItemList(ix).FSubTotalPrice - ojumun.FItemList(ix).FsumPaymentEtc),0) %></b></font></td>


		<td><%= ojumun.FItemList(ix).JumunMethodName %></td>
		<% if ojumun.FItemList(ix).FIpkumdiv="1" then %>
		<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><acronym title="<%= ojumun.FItemList(ix).Fresultmsg %>"><%= ojumun.FItemList(ix).IpkumDivName %></acronym></font></td>
		<% else %>
		<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
		<% end if %>
		<td><acronym title="<%= ojumun.FItemList(ix).FRegDate %>"><%= Left(ojumun.FItemList(ix).FRegDate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FItemList(ix).Fipkumdate %>"><%= Left(ojumun.FItemList(ix).Fipkumdate,10) %></acronym></td>
		<!--<td><acronym title="<%= ojumun.FItemList(ix).Fbaljudate %>"><%= Left(ojumun.FItemList(ix).Fbaljudate,10) %></acronym></td>-->
		<!--td><acronym title="<%= ojumun.FItemList(ix).Fbeadaldate %>"><%= Left(ojumun.FItemList(ix).Fbeadaldate,10) %></acronym></td-->
		<!--td><%= ojumun.FItemList(ix).Fdeliverno %></td>-->
	</tr>
	<% next %>

<% end if %>

    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="20">
            <% if ojumun.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
    			<% if ix>ojumun.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(ix) then %>
    			<font color="red">[<%= ix %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
    			<% end if %>
    		<% next %>

    		<% if ojumun.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
        </td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->


<script language='javascript'>
    ChangeFormBgColor(frm);

    <% if ResultOneOrderserial<>"" then %>
    GotoOrderDetail('<%= ResultOneOrderserial %>')
    // top.detailFrame.location.href = "ordermaster_detail.asp?orderserial=<%= ResultOneOrderserial %>";
    <% end if %>

    <% if (AlertMsg<>"") then %>
        alert('<%= AlertMsg %>');
	<% end if %>
</script>
<%
set ojumun = Nothing
%>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
