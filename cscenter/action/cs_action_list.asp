<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : cs���� csó������Ʈ
' History	:  2007.06.01 �̻� ����
'              2022.08.16 �ѿ�� ����(isms������ġ)
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%
'' ���� ����� ��û���� �������� ��� ��ȸ�����ϰ� ����, skyer9, 2017-03-09
''if (session("ssAdminPsn") = 9) then
''	'// ����
''	if (session("ssBctId") <> "josin222") and (session("ssBctId") <> "jjh") and (session("ssBctId") <> "sunna0822") then
''		response.write "<br><br>������ �����ϴ�."
''		response.end
''	end if
''end if

Dim delYN		: delYN	 = requestCheckvar(request("delYN"),1)
Dim periodYN	: periodYN = requestCheckvar(request("periodYN"),1)
Dim notfinishYN	: notfinishYN = requestCheckvar(request("notfinishYN"),1)
Dim research	: research = requestCheckvar(request("research"),2)
dim i, userid, username, orderserial, makerid, searchfield, searchstring, asid, writeUser, extsitename, checkExtSite, finishuser
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, yyyymmdd1, fromDate, toDate, notfinishtype, divcd, currstate
Dim onlycustomerjupsu, onlycsservicerefund, searchtype, dateType, upreturnmifinishBaseDate, tmpSql, page
dim ResultOneCsID, ix
	userid      	= requestCheckvar(request("userid"),32)
	username    	= requestCheckvar(request("username"),32)
	orderserial 	= requestCheckvar(request("orderserial"),32)
	asid 			= requestCheckvar(request("asid"),32)
	searchfield 	= requestCheckvar(request("searchfield"),32)
	searchstring 	= requestCheckvar(request("searchstring"),32)
	notfinishtype  	= requestCheckvar(request("notfinishtype"),32)
	divcd       	= requestCheckvar(request("divcd"),32)
	currstate   	= requestCheckvar(request("currstate"),32)
	extsitename 	= requestCheckvar(request("extsitename"),32)
	checkExtSite	= requestCheckvar(request("checkExtSite"),32)
	onlycustomerjupsu	= requestCheckvar(request("onlycustomerjupsu"),32)
	onlycsservicerefund	= requestCheckvar(request("onlycsservicerefund"),32)
	searchtype		= requestCheckvar(request("searchtype"),32)			'// ȣȯ���� ���� ���ܵд�.(������� [CS]������>>[CS]���� ���� ���� ���)
	dateType		= requestCheckvar(request("dateType"),32)
	yyyy1   = requestcheckvar(request("yyyy1"),4)
	yyyy2   = requestcheckvar(request("yyyy2"),4)
	mm1     = requestcheckvar(request("mm1"),2)
	mm2     = requestcheckvar(request("mm2"),2)
	dd1     = requestcheckvar(request("dd1"),2)
	dd2     = requestcheckvar(request("dd2"),2)
	page = requestcheckvar(getNumeric(request("page")),10)

if page="" then page=1

if (searchtype <> "") then
	if (searchtype = "searchfield") then
		'
	else
		notfinishYN = "Y"
		notfinishtype = searchtype
	end if
end if

if (research = "") then

	delYN = "N"

	if (searchtype <> "upchefinish") then
		periodYN = "Y"
	end if

	'// userid/orderserial �Ķ���Ͱ� �������� �ش� �Ķ���ͷ� ����
	'// (�ٸ� ���������� ��ũ�� �ɾ� �˾��� ���������� ���� ó��.)
	if (userid <> "") then
	    searchfield = "userid"
	    searchstring = userid
	elseif (orderserial <> "") then
	    searchfield = "orderserial"
	    searchstring = orderserial
	end if

    if (notfinishtype = "confirm") then
        divcd = "A003"
        currstate = "B005"
    elseif (notfinishtype = "cardnocheckdp1") then
        divcd = "A007"
        currstate = "notfinish"
    elseif (notfinishtype = "norefund") then
        divcd = "A003"
        currstate = "B001"
    end if
end if

if (searchfield <> "") and (searchstring <> "") then

    if (searchfield = "userid") then

            userid = searchstring

    elseif (searchfield = "orderserial") then

            orderserial = searchstring

    elseif (searchfield = "username") then

            username = searchstring

    elseif (searchfield = "makerid") then

            makerid = searchstring

	elseif (searchfield = "writeUser") then

            writeUser = searchstring

	elseif (searchfield = "finishuser") then

            finishuser = searchstring

	elseif (searchfield = "asid") then

			asid = searchstring

    end If

end if

if (searchfield = "") and (searchstring <> "") then

	if IsNumeric(searchstring) and Len(searchstring) >= 11 then
		'// �ֹ���ȣ �˻�
		searchfield = "orderserial"
		orderserial = searchstring
	end if

end if

if (yyyy1="") then
    yyyymmdd1 = dateAdd("m",-3,now())			'// [CS]������>>[CS]���� ����
    yyyy1 = Cstr(Year(yyyymmdd1))
    mm1 = Cstr(Month(yyyymmdd1))
    dd1 = Cstr(day(yyyymmdd1))
end if

if (yyyy2 = "") then
	if (notfinishtype = "upreturnmifinish") then
		'// ��ü��ǰ��ó���� ��� �⺻�� = D+7 ��
		tmpSql = " exec [db_cs].[dbo].[usp_getDayMinusWorkday] '" & Left(now(), 10) & "', 7 " & VbCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open tmpSql, dbget, adOpenForwardOnly
		if Not rsget.Eof then
		    '// �ٹ��ϼ� ���� D+7 ��
		    upreturnmifinishBaseDate = rsget("minusworkday")

		    yyyy2 = Cstr(Year(upreturnmifinishBaseDate))
		    mm2 = Cstr(Month(upreturnmifinishBaseDate))
		    dd2 = Cstr(day(upreturnmifinishBaseDate))
		end if
		rsget.close
	end if

    if (notfinishtype = "cardnocheckdp1") then
        toDate = DateAdd("d", -1, Now())
        yyyy2 = Cstr(Year(toDate))
        mm2 = Cstr(Month(toDate))
        dd2 = Cstr(day(toDate))
        notfinishtype = "cardnocheck"
    end if

	if (yyyy2="")   then yyyy2 = Cstr(Year(now()))
	if (mm2="")     then mm2 = Cstr(Month(now()))
	if (dd2="")     then dd2 = Cstr(day(now()))
end if

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FPageSize = 10
ocsaslist.FCurrPage = page

if (searchfield <> "") and (searchstring <> "") then
    ocsaslist.FRectUserID = userid
    ocsaslist.FRectUserName = username
    ocsaslist.FRectOrderSerial = orderserial
    ocsaslist.FRectMakerid  = makerid
    ocsaslist.FRectWriteUser = writeUser
	ocsaslist.FRectFinishUser = finishuser
	ocsaslist.FRectCsAsID = asid
end if

ocsaslist.FRectDivcd = divcd
ocsaslist.FRectCurrstate = currstate

if (orderserial = "") and (userid = "") then
	'// �ֹ���ȣ �Ǵ� ���̵� �˻��ϸ� �������� ���� ǥ��
	ocsaslist.FRectDeleteYN	= delYN
end if

if (notfinishYN = "Y") then
	ocsaslist.FRectSearchType = notfinishtype
end if

If (periodYN = "Y") and (orderserial = "") Then
	'// �ֹ���ȣ �Է��ϸ� �Ⱓ���� ����
	ocsaslist.FRectDateType = dateType
	ocsaslist.FRectStartDate = fromDate
	ocsaslist.FRectEndDate = toDate
End If

IF (checkExtSite<>"") then                      '''2011-06 �߰�
    ocsaslist.FRectExtSitename = ExtSitename
ENd IF

ocsaslist.FRectOnlyCustomerJupsu = onlycustomerjupsu
ocsaslist.FRectOnlyCSServiceRefund = onlycsservicerefund
''ocsaslist.GetCSASMasterListNew
ocsaslist.GetCSASMasterListByProcedure

if ocsaslist.FResultCount=1 then
    ResultOneCsID = ocsaslist.FItemList(0).FId
end if

%>
<link rel="stylesheet" href="/cscenter/css/cs.css" type="text/css">
<style>
.csH15 { line-height: 15px; }
</style>
<script type='text/javascript'>
// tr ���󺯰�
var pre_selected_row = null;

function ChangeColor(e, selcolor, defcolor){
	if (pre_selected_row != null) {
	        pre_selected_row.bgColor = defcolor;
        }
        pre_selected_row = e;
        e.bgColor = selcolor;
}

function searchDetail(idx){
    buffrm.id.value = idx;
    buffrm.submit();
}

function NextPage(page){
	frm.target = "";
	frm.action = "cs_action_list.asp"
    frm.page.value = page;
    frm.submit();
}

function reSearch(){
	if (frm.searchfield.value=="asid"){
		if (frm.searchstring.value!=""){
			if (!IsDouble(frm.searchstring.value)){
				alert('cs��ȣ�� ���ڸ� �����մϴ�.');
				frm.searchstring.focus();
				return;
			}
		}
	}

	frm.target = "";
	frm.action = "cs_action_list.asp"
    frm.page.value="1";
    frm.submit();
}

<% ' ����� ������, isms ������ġ�� ���� %>
//function reSearchExcelDown(){
//	frm.target = "exceldown";
<% '	frm.action = "/cscenter/action/cs_action_list_excel.asp" %>
//    frm.submit();
//	frm.target = "";
//	frm.action = ""
//}

function reSearchByOrderserial(iorderserial){
    frm.searchfield.value = "orderserial";
    frm.searchstring.value = iorderserial;

    frm.divcd.value = "";
    frm.currstate.value = "";

	// frm.notfinishYN.checked = false;
	// frm.periodYN.checked = false;
	// frm.checkExtSite.checked = false;
	// frm.delYN.checked = false;

    frm.page.value="1";
	frm.target = "";
	frm.action = "cs_action_list.asp"
    frm.submit();
}

function reSearchByUserid(iuserid){
    frm.searchfield.value = "userid";
    frm.searchstring.value = iuserid;

    frm.divcd.value = "";
    frm.currstate.value = "";

	// frm.notfinishYN.checked = false;
	// frm.periodYN.checked = false;
	// frm.checkExtSite.checked = false;
	// frm.delYN.checked = false;

    frm.page.value="1";
	frm.target = "";
	frm.action = "cs_action_list.asp"
    frm.submit();
}

function reSearchByMakerid(imakerid){
    frm.searchfield.value = "makerid";
    frm.searchstring.value = imakerid;

    frm.divcd.value = "";
    frm.currstate.value = "";

	// frm.notfinishYN.checked = false;
	// frm.periodYN.checked = false;
	// frm.checkExtSite.checked = false;
	// frm.delYN.checked = false;

    frm.page.value="1";
	frm.target = "";
	frm.action = "cs_action_list.asp"
    frm.submit();
}

function SetComp(comp) {
	frm.notfinishYN.checked = true;
}

function SetExtCheck(comp) {
    if (comp.name=="checkExtSite"){
        if (comp.checked){
            frm.extsitename.style.background = "#FFFFFF";
        }else{
            frm.extsitename.style.background = "#EEEEEE";
        }
    }
}

function pop_modal_repay(id){
	if (id == "") {
	        alert("���� CS��û�� �����ϼ���.");
	        return;
        }
	var popwin = window.open("pop_modal_repay.asp?id=" + id,"pop_modal_repay","width=350 height=350 scrollbars=no resizable=no");
	popwin.focus();
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

</script>

<form name="frm" method="get" action="/cscenter/action/cs_action_list.asp" style="margin:0px;" >
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="F4F4F4">
	<tr>
    	<td>
			&nbsp;
            �˻� :
            <select class="select" name="searchfield">
            	<option value="" <% if (searchfield = "") then %>selected<% end if %>>��ü</option>
				<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %>>�ֹ���ȣ</option>
				<option value="username" <% if (searchfield = "username") then %>selected<% end if %>>����</option>
				<option value="userid" <% if (searchfield = "userid") then %>selected<% end if %>>���̵�</option>
				<option value="makerid" <% if (searchfield = "makerid") then %>selected<% end if %>>��üó�����̵�</option>
				<option value="writeUser" <% if (searchfield = "writeUser") then %>selected<% end if %>>�����ھ��̵�</option>
				<option value="finishuser" <% if (searchfield = "finishuser") then %>selected<% end if %>>ó���ھ��̵�</option>
				<option value="asid" <% if (searchfield = "asid") then %>selected<% end if %>>CSidx</option>
            </select>
            <input type="text" class="text" name="searchstring" value="<%= searchstring %>" size="14">
            &nbsp;&nbsp;
            ����:
            <select class="select" name="divcd">
            	<option value="">��ü</option>
            	<option value="">-------------------------</option>
				<option value="A000" <% if (divcd = "A000") then response.write "selected" end if %>>��ȯ���</option>
				<option value="A100" <% if (divcd = "A100") then response.write "selected" end if %>>��ȯ���(��ǰ����)</option>
				<option value="">-------------------------</option>
				<option value="A004" <% if (divcd = "A004") then response.write "selected" end if %>>��ǰ����(����)</option>
				<option value="">-------------------------</option>
				<option value="A001" <% if (divcd = "A001") then response.write "selected" end if %>>������߼�</option>
				<option value="A002" <% if (divcd = "A002") then response.write "selected" end if %>>���񽺹߼�</option>
				<option value="A200" <% if (divcd = "A200") then response.write "selected" end if %>>��Ÿȸ��</option>
				<option value="">-------------------------</option>
				<option value="A003" <% if (divcd = "A003") then response.write "selected" end if %>>ȯ�ҿ�û</option>
				<option value="A005" <% if (divcd = "A005") then response.write "selected" end if %>>�ܺθ�ȯ�ҿ�û</option>
				<option value="A007" <% if (divcd = "A007") then response.write "selected" end if %>>�ſ�ī��/��ü��ҿ�û</option>
				<option value="A700" <% if (divcd = "A700") then response.write "selected" end if %>>��ü��Ÿ����</option>
				<option value="A999" <% if (divcd = "A999") then response.write "selected" end if %>>���߰�����</option>
				<option value="">-------------------------</option>
				<option value="A060" <% if (divcd = "A060") then response.write "selected" end if %>>��ü��޹���</option>
				<option value="A006" <% if (divcd = "A006") then response.write "selected" end if %>>�������ǻ���</option>
				<option value="A008" <% if (divcd = "A008") then response.write "selected" end if %>>�ֹ����</option>
				<option value="A009" <% if (divcd = "A009") then response.write "selected" end if %>>��Ÿ����(�޸�)</option>
				<option value="A900" <% if (divcd = "A900") then response.write "selected" end if %>>�ֹ���������</option>
				<option value="">-------------------------</option>
				<option value="A010" <% if (divcd = "A010") then response.write "selected" end if %>>ȸ����û(�ٹ�)</option>
				<option value="">-------------------------</option>
				<option value="A011" <% if (divcd = "A011") then response.write "selected" end if %>>��ȯȸ��(�ٹ�)</option>
				<option value="A012" <% if (divcd = "A012") then response.write "selected" end if %>>��ȯȸ��(����)</option>
				<option value="A111" <% if (divcd = "A111") then response.write "selected" end if %>>��ȯȸ��(��ǰ����,�ٹ�)</option>
				<option value="A112" <% if (divcd = "A112") then response.write "selected" end if %>>��ȯȸ��(��ǰ����,����)</option>
            </select>
            &nbsp;&nbsp;
            ����:
            <select class="select" name="currstate">
            	<option value="">��ü</option>
				<option value="B001" <% if (currstate = "B001") then response.write "selected" end if %>>����</option>
				<option value="notfinish" <% if (currstate = "notfinish") then response.write "selected" end if %>>��ó����ü</option> <!-- 6�ܰ����� -->
				<option value="B003" <% if (currstate = "B003") then response.write "selected" end if %>>�ù������</option>
				<option value="B004" <% if (currstate = "B004") then response.write "selected" end if %>>������Է�</option>
				<option value="B005" <% if (currstate = "B005") then response.write "selected" end if %>>Ȯ�ο�û</option>
				<option value="B006" <% if (currstate = "B006") then response.write "selected" end if %>>��üó���Ϸ�</option>
				<option value="B007" <% if (currstate = "B007") then response.write "selected" end if %>>�Ϸ�</option>
            </select>
            &nbsp;&nbsp;
			<input type="checkbox" name="delYN" value="N" <%if (delYN="N") then %>checked<% end if %>>����(���)����
        </td>
        <td width="100" align="right" valign="top" rowspan="3">
			<% '<input type="button" class="button_s" value="�����ٿ�ε�" onclick="reSearchExcelDown();"><br /> %>
            <input type="button" class="button_s" value="���ΰ�ħ" onclick="document.location.reload();"><br />
			<div style="height: 5px;"></div>
            <input type="button" class="button_s" value="�˻��ϱ�" onclick="reSearch();">
        </td>
	</tr>
	<tr>
    	<td>
    		&nbsp;
    		<input type="checkbox" name="notfinishYN" value="Y" <%=CHKIIF(notfinishYN="Y","checked","")%>>
    		��ó��CS :
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="notfinish" <% if (notfinishtype = "notfinish") then %>checked<% end if %>> ��ü
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="norefundmile" <% if (notfinishtype = "norefundmile") then %>checked<% end if %>> ���ϸ���/��ġ�� ȯ��
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="cardnocheck" <% if (notfinishtype = "cardnocheck") then %>checked<% end if %>> ī�����
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="cancelnofinish" <% if (notfinishtype = "cancelnofinish") then %>checked<% end if %>> �ֹ����
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="beasongnocheck" <% if (notfinishtype = "beasongnocheck") then %>checked<% end if %>> �������ǻ���
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="upchemifinish" <% if (notfinishtype = "upchemifinish") then %>checked<% end if %>> ��ü��ó����ü
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="upreturnmifinish" <% if (notfinishtype = "upreturnmifinish") then %>checked<% end if %>> ��ü��ǰ
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="upchefinish" <% if (notfinishtype = "upchefinish") then %>checked<% end if %>> ��üó���Ϸ�
			<input type="radio" name="notfinishtype" onClick="SetComp(this)" value="logicsfinish" <% if (notfinishtype = "logicsfinish") then %>checked<% end if %>> ����ó���Ϸ�
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="chulgofinishnotreceive" <% if (notfinishtype = "chulgofinishnotreceive") then %>checked<% end if %>> ��ȯ����Ĺ�ȸ��
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="returnmifinish" <% if (notfinishtype = "returnmifinish") then %>checked<% end if %>> ȸ����û
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="confirm" <% if (notfinishtype = "confirm") then %>checked<% end if %>> Ȯ�ο�û(ȯ��)
			<input type="radio" name="notfinishtype" onClick="SetComp(this)" value="upcheconfirm" <% if (notfinishtype = "upcheconfirm") then %>checked<% end if %>> Ȯ�ο�û(��ü)
            <input type="radio" name="notfinishtype" onClick="SetComp(this)" value="norefundetc" <% if (notfinishtype = "norefundetc") then %>checked<% end if %>> �ܺθ�ȯ��
			<input type="radio" name="notfinishtype" onClick="SetComp(this)" value="customeraddpay" <% if (notfinishtype = "customeraddpay") then %>checked<% end if %>> ���߰��Ա�
        </td>
	</tr>
	<tr>
    	<td>
    		&nbsp;
            (Total : <%= ocsaslist.FTotalCount%> ��)
            &nbsp;
            <input type="checkbox" name="periodYN" value="Y" <%=CHKIIF(periodYN="Y","checked","")%>>
			<select class="select" name="dateType">
				<option value="regdate" <%= CHKIIF(dateType="regdate", "selected", "") %> >������</option>
				<option value="finishdate" <%= CHKIIF(dateType="finishdate", "selected", "") %> >ó����</option>
			</select>
             :
            <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;&nbsp;
            <input type="checkbox" name="checkExtSite" value="Y" <% if checkExtSite="Y" then response.write "checked" %> onClick="SetExtCheck(this)">
            Ư������Ʈ : <% DrawSelectExtSiteName "extsitename", extsitename %>
			&nbsp;&nbsp;
			<input type="checkbox" name="onlycustomerjupsu" value="Y" <%if (onlycustomerjupsu="Y") then %>checked<% end if %>>�� ����������
			&nbsp;&nbsp;
			<input type="checkbox" name="onlycsservicerefund" value="Y" <%if (onlycsservicerefund="Y") then %>checked<% end if %>>CS���� ȯ�Ҹ�
        </td>
	</tr>
</table>
</form>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a csH15" cellpadding="2" cellspacing="1" bgcolor="#BABABA" style="table-layout:fixed">
    <tr bgcolor="<%= adminColor("tabletop") %>">
        <td width="70" align="center">Idx</td>
        <td width="100" align="center">����</td>
        <td width="90" align="center">���ֹ���ȣ</td>
        <td width="90" align="center">Site</td>
        <td width="110" align="center">��üID</td>
        <td width="50" align="center">����</td>
        <td width="80" align="center">���̵�</td>
        <td align="center">����</td>
        <td width="75" align="center">����</td>
		<td width="75" align="center">������</td>
		<td width="75" align="center">ó����</td>
        <td width="70" align="center">ȯ�ұݾ�</td>
        <td width="80" align="center">�����</td>
        <td width="80" align="center">��üȮ��</td>
        <td width="80" align="center">ó����</td>
        <td width="30" align="center">����</td>
    </tr>

<% for i = 0 to (ocsaslist.FResultCount - 1) %>
    <% if (ocsaslist.FItemList(i).Fdeleteyn <> "N") then %>
    <tr bgcolor="#EEEEEE" class="csH15 csMp" style="color:gray" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= ocsaslist.FItemList(i).Fid %>');">
    <% else %>
	<tr bgcolor="#FFFFFF" class="csH15 csMp" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= ocsaslist.FItemList(i).Fid %>');">
    <% end if %>
        <td class="csNoWrap"><%= ocsaslist.FItemList(i).Fid %></td>
        <td class="csNoWrap" align="left"><font color="<%= ocsaslist.FItemList(i).GetAsDivCDColor %>"><%= ocsaslist.FItemList(i).GetAsDivCDName %></font></td>
        <td class="csNoWrap">
        	<a href="javascript:reSearchByOrderserial('<%= ocsaslist.FItemList(i).Forgorderserial %>');" >
        		<%= ocsaslist.FItemList(i).Forgorderserial %>
        		<% if (ocsaslist.FItemList(i).Forderserial <> ocsaslist.FItemList(i).Forgorderserial) then %>
        			+
        		<% end if %>
        	</a>
        </td>
        <td class="csNoWrap"><%= ocsaslist.FItemList(i).FExtsitename %></td>
        <td class="csNoWrap" align="left">
            <acronym title="<%= ocsaslist.FItemList(i).Fmakerid %>"><a href="javascript:reSearchByMakerid('<%= ocsaslist.FItemList(i).Fmakerid %>');" ><%= Left(ocsaslist.FItemList(i).Fmakerid,32) %></a></acronym>
		</td>
        <td class="csNoWrap">
			<%= AstarUserName(ocsaslist.FItemList(i).Fcustomername) %>
        </td>
        <td class="csNoWrap" align="left">
        	<!--<acronym title="<%'= ocsaslist.FItemList(i).Fuserid %>">-->
        	<!--<a href="javascript:reSearchByUserid('<%'= ocsaslist.FItemList(i).Fuserid %>');" >-->
			<%= AstarUserid(ocsaslist.FItemList(i).Fuserid) %>
        	<!--</a>-->
        	<!--</acronym>-->
        </td>
        <td class="csNoWrap" align="left">
			<%= ocsaslist.FItemList(i).Ftitle %>
			<% if ocsaslist.FItemList(i).FExtsitename<>"10x10" then %>(<%= ocsaslist.FItemList(i).FAuthCode %>)<% end if %>
		</td>
        <td class="csNoWrap"><font color="<%= ocsaslist.FItemList(i).GetCurrstateColor %>"><%= ocsaslist.FItemList(i).GetCurrstateName %></font></td>
		<td class="csNoWrap"><%= ocsaslist.FItemList(i).Fwriteuser %></td>
		<td class="csNoWrap"><%= ocsaslist.FItemList(i).Ffinishuser %></td>
        <td class="csNoWrap" align="right"><%= FormatNumber(ocsaslist.FItemList(i).Frefundrequire,0) %></td>
        <td class="csNoWrap"><acronym title="<%= ocsaslist.FItemList(i).Fregdate %>"><%= Left(ocsaslist.FItemList(i).Fregdate,10) %></acronym></td>
		<td class="csNoWrap"><acronym title="<%= ocsaslist.FItemList(i).Fconfirmdate %>"><%= Left(ocsaslist.FItemList(i).Fconfirmdate,10) %></acronym></td>
        <td class="csNoWrap"><acronym title="<%= ocsaslist.FItemList(i).Ffinishdate %>"><%= Left(ocsaslist.FItemList(i).Ffinishdate,10) %></acronym></td>
        <td class="csNoWrap">
        <% if ocsaslist.FItemList(i).Fdeleteyn="Y" then %>
        <font color="red">����</font>
        <% elseif ocsaslist.FItemList(i).Fdeleteyn="C" then %>
        <font color="red"><strong>���</strong></font>
        <% end if %>
        </td>
    </tr>
<% next %>
<% if (ocsaslist.FResultCount < 9) then %>
        <% for i = 0 to (9 - (ocsaslist.FResultCount mod 9)) %>
    <tr bgcolor="#FFFFFF" class="csH15 csMp" align="center">
        <td class="csNoWrap">-</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
		<td></td>
		<td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
        <% next %>
<% end if %>
    <tr bgcolor="#FFFFFF" >
        <td colspan="16" align="center">
            <% if ocsaslist.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ocsaslist.StarScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for ix=0 + ocsaslist.StarScrollPage to ocsaslist.FScrollCount + ocsaslist.StarScrollPage - 1 %>
    			<% if ix>ocsaslist.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(ix) then %>
    			<font color="red">[<%= ix %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
    			<% end if %>
    		<% next %>

    		<% if ocsaslist.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
        </td>
    </tr>

</table>

<form name="buffrm" method="get" target="detailFrame" action="/cscenter/action/cs_action_detail.asp" style="margin:0px;" >
<input type="hidden" name="id" value="">
</form>

<script type='text/javascript'>

    <% if ResultOneCsID<>"" then %>
    if (top.detailFrame!=undefined){
        top.detailFrame.location.href = "cs_action_detail.asp?id=<%= ResultOneCsID %>";
    }
    <% end if %>

</script>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="300"></iframe>
<% else %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="0"></iframe>
<% end if %>

<%
set ocsaslist = Nothing
%>
<script type='text/javascript'>

function getOnload(){
SetExtCheck(frm.checkExtSite);
}

window.onload=getOnload;

</script>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
