<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventWinner/event_confirmList.asp
' Description :  �̺�Ʈ ��÷�� ���� ������
' History : 2007.09.27 ������
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/eventWinner_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventWinnerManageCls.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/bct.css" type="text/css">
</head>
<body topmargin="0" >

<%

dim evtCode,Page,PageSize,ScrollCount,SortingOpt,i
evtCode =request("eC")
page = request("Page")
SortingOpt = request("srtOpt")
if Page="" then Page=1
PageSize =20
ScrollCount = 20


dim appList,appOne
dim arrOne,arrList,intLoop
dim evtName,totcnt

set appList = new ClsEventEntry
appList.FECode = evtCode
appList.FSortingOpt = SortingOpt
totcnt = appList.fnGetSelectedIdCount
arrList = appList.fnGetSelectedList

set appList = nothing

set appOne = new ClsEvent
appOne.FECode = evtCode
appOne.fnGetEventCont
evtName = appOne.FEName
set appOne = nothing
'	 0         1             2       3         4              5          6          7
'w.evt_code,w.evtcom_idx,w.userid,w.regdate,w.smsSended,w.mailSended ,g.userLevel,g.username
'      8          9           10        11        12     13
',n.usercell,n.userPhone,n.usermail,n.zipcode,address,ranking
%>

<script language="javascript">

function AnSelectAllChk(bool){
	var frm = document.getElementsByName('cksel');
	for (var i=0;i<frm.length;i++){
		if (frm[i].disabled!=true){
			frm[i].checked = bool;
			AnCheckClick(frm[i]);
		}
	}
}

function checkedValue(){
	var tgvalue="";
	var chkbx = document.getElementsByName('cksel');

	for (var i=0;i<chkbx.length;i++) {
		if (chkbx[i].checked){
			tgvalue=tgvalue  + chkbx[i].value + ",";
		}
	}

	if (tgvalue.length < 1){
		alert('�ϳ� �̻� ������ �ּ���');
		return '';
	}else{
		return tgvalue;
	}
}
function selEntry(strSel){

	var arridx = checkedValue();

	if (arridx.length < 1){
		return;
	} else {
		selFrm.arridx.value = arridx;
		selFrm.selStr.value=strSel;
		selFrm.target="selFrame";
		selFrm.action="event_entry_process.asp";
		selFrm.submit();
	}
}

// �������׵��
function fnNotice(){

	var arridx = checkedValue();

	if (arridx.length < 1){
		return;
	} else {
		window.open("", "pop", "width=10,height=10,menubar=no,toolbar=no,scrollbars=no,status=no,resizable=no,location=no");
		selFrm.arridx.value = arridx;
		selFrm.action="pop_event_winner.asp"
		selFrm.target="pop";
		selFrm.submit();
	}

}
// ������
function fnSongjang(){
	var arridx = checkedValue();

	if (arridx.length < 1){
		return;
	} else {
		window.open("", "pop", "width=10,height=10,menubar=no,toolbar=no,scrollbars=no,status=no,resizable=no,location=no");
		selFrm.arridx.value = arridx;
		selFrm.action="pop_event_winner.asp"
		selFrm.target="pop";
		selFrm.submit();
	}

}
//SMS ������
function fnSendSMS(){
	var arridx = checkedValue();

	if (arridx.length < 1){
		return;
	} else {

		window.open('','pop','width=350,height=350,top=150,left=300,menubar=no,toolbar=no,scrollbars=no,status=no,resizable=no,location=no');
		selFrm.arridx.value = arridx;
		selFrm.action="pop_evt_sms.asp"
		selFrm.target="pop";
		selFrm.submit();
	}
}
//���� ������
function fnSendMail(){
	var arridx = checkedValue();

	if (arridx.length < 1){
		return;
	} else {
		window.open('', 'pop', 'width=650,height=500,top=150,left=300,menubar=no,toolbar=no,scrollbars=no,status=no,resizable=no,location=no');
		selFrm.arridx.value = arridx;
		selFrm.action="pop_evt_mail.asp"
		selFrm.target="pop";
		selFrm.submit();
	}
}
// �����ڸ���Ʈ ����
function fnGoSelectList(){
	document.location.href="event_entryList.asp?eC=<%= evtCode %>";
}
</script>

<!-- ���̺� ��� �˻��� ���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="ListFrm" method="get" action="">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td align="left"><input type="text" name="eventName" size="50" value="<%= evtName %>">�� ��÷�ڼ�:<%= totcnt %></td>
        <td align="right">
			<input type="button" class="button" value="�����ڸ���Ʈ ����" onclick="fnGoSelectList()">
        </td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ���̺� ��� �˻��� �� -->
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td width="30" align="center"><input type="checkbox" name="ckselm" onClick="AnSelectAllChk(this.checked);"></td>
		<td width="70" align="center"><a href="?eC=<%= evtCode %>&srtOpt=DT">�����</a></td>
		<td width="110" align="center"><a href="?eC=<%= evtCode %>&srtOpt=UID">���̵�</a></td>
		<td width="60" align="center"><a href="?eC=<%= evtCode %>&srtOpt=UNM">����</a></td>
		<td width="90" align="center"><a href="?eC=<%= evtCode %>&srtOpt=HPNO">�ڵ���</a></td>
		<td width="80" align="center"><a href="?eC=<%= evtCode %>&srtOpt=TelNO">������ȭ</a></td>
		<td width="160" align="center"><a href="?eC=<%= evtCode %>&srtOpt=UMail">�̸���</a></td>
		<td width="50" align="center">�����ȣ</td>
		<td align="center">�ּ�</td>
		<td align="center" width="40">SMS</td>
		<td align="center" width="40">e-mail</td>
		<td align="center" width="40"><a href="?eC=<%= evtCode %>&srtOpt=Rank">���</a></td>
	</tr>
	<% if isArray(arrList) then %>
	<% for intLoop=0 to Ubound(arrList,2) %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><input type="checkbox" name="cksel" value="<%= arrList(2,intLoop) %>" onClick="AnCheckClick(this);"></td>
		<td align="center"><%= dateValue(arrList(3,intLoop)) %></td>
		<td align="center"><%= GetUserLevelColorStr(arrList(6,intLoop),arrList(2,intLoop)) %></td>
		<td align="center"><%= arrList(7,intLoop) %></td>
		<td align="center"><%= arrList(8,intLoop) %></td>
		<td align="right"><%= arrList(9,intLoop) %></td>
		<td align="center"><%= arrList(10,intLoop) %></td>
		<td align="center"><%= arrList(11,intLoop) %></td>
		<td align="left"><%= arrList(12,intLoop) %></td>
		<td align="center"><% if (arrList(4,intLoop)) then response.write "�߼�" else response.write "�̹߼�" end if %></td>
		<td align="center"><% if (arrList(5,intLoop)) then response.write "�߼�" else response.write "�̹߼�" end if %></td>
		<td align="center" width="40"><%= arrList(13,intLoop) %></td>
	</tr>

	<% next %>
	<% end if %>

</table>
<!-- �ϴ� ���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
				<tr>
					<td width="150" align="left">

						<input type="button" class="button" value="����" onclick="selEntry('N');">&nbsp;&nbsp;&nbsp;

					</td>
					<td align="right">
						<input type="button" class="button" value="�������� ���" onclick="fnNotice();">&nbsp;&nbsp;&nbsp;
						<input type="button" class="button" value="������" onclick="fnSongjang();">&nbsp;&nbsp;&nbsp;
						<input type="button" class="button" value="SMS ������" onclick="fnSendSMS();">&nbsp;&nbsp;&nbsp;
						<input type="button" class="button" value="���� �ۼ�" onclick="fnSendMail();">

					</td>
				</tr>
			</table>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<iframe name="selFrame" src="" frameborder="0" width="0" height="0"></iframe>
<form name="selFrm" method="post" action="event_entry_process.asp">
<input type="hidden" name="eC" value="<%= evtCode %>">
<input type="hidden" name="arridx" value="">
<input type="hidden" name="selStr" value="">
</form>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->