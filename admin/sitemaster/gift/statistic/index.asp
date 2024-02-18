<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ����Ʈ ���
' Hieditor : 2015.05.27 ���ر� ����
'			 2016.07.19 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)

'// ���ã��
dim IsMenuFavoriteAdded

IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery_common.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminPartnerDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<script language='javascript'>
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;

	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "���ã�⿡�� �����Ͻðڽ��ϱ�?";
	} else {
		msg = "���ã�⿡ �߰��Ͻðڽ��ϱ�?";
	}

	ret = confirm(msg);

	if (ret) {
		frm.submit();
	}
}
</script>
<% if session("sslgnMethod")<>"S" then %>
<!-- USBŰ ó�� ���� (2008.06.23;������) -->
<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
<script language="javascript" src="/js/check_USBToken.js"></script>
<!-- USBŰ ó�� �� -->
<% end if %>
</head>
<body <% if session("sslgnMethod")<>"S" then %>onload="checkUSBKey()"<% end if %>>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/gift/giftStatisticCls.asp" -->
<%
	Dim i, sDate, eDate, cStat, vTab, vArrTalk, vArrDay, vArrShop
	Dim vTotPC, vTotMob, vTotTalkPC, vTotTalkMob, vTotDayPC, vTotDayMob, vTotShopPC, vTotShopMob
	Dim vTW5, vTW0, vTW1, vTW2, vTW3, vTW4, vTW6, vTW7, vTM5, vTM0, vTM1, vTM2, vTM3, vTM4, vTM6, vTM7
	Dim vDW5, vDW0, vDW1, vDW2, vDW3, vDW4, vDW6, vDW7, vDM5, vDM0, vDM1, vDM2, vDM3, vDM4, vDM6, vDM7
	Dim vSW5, vSW0, vSW1, vSW2, vSW3, vSW4, vSW6, vSW7, vTot5, vTot0, vTot1, vTot2, vTot3, vTot4, vTot6, vTot7
	sDate = NullFillWith(request("sDate"),DateAdd("d",-10,date()))
	eDate = NullFillWith(request("eDate"),date())
	vTab = NullFillWith(request("tab"),1)
	
	
	SET cStat = New CgiftStat_list
	If vTab = "1" Then
		cStat.FRectSDate = sDate
		cStat.FRectEDate = eDate
		cStat.sbStatDaily
	ElseIf vTab = "2" Then
		cStat.FRectGubun = "talk"
		cStat.FRectSDate = sDate
		cStat.FRectEDate = eDate
		vArrTalk = cStat.fnStatUserLevel
		
		cStat.FRectGubun = "day"
		cStat.FRectSDate = sDate
		cStat.FRectEDate = eDate
		vArrDay = cStat.fnStatUserLevel
		
		cStat.FRectGubun = "shop"
		cStat.FRectSDate = sDate
		cStat.FRectEDate = eDate
		vArrShop = cStat.fnStatUserLevel
	End If
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function goTab(a){
	$('input[name="tab"]').val(a);
	frm1.submit();
}

function goExcelDown(){
	frm1.action = "exceldown.asp";
	frm1.submit();
	
	frm1.action = "";
}
</script>
<div class="contSectFix scrl">
	<div class="contHead">
		<table>
			<tr>
				<td width="90px"><a href="javascript:goTab('1');"><span style="font-size:11pt;<%=CHKIIF(vTab="1","font-weight:bold;text-decoration:underline;","")%>">[�Ϻ� ��ȸ]</span></a></td>
				<td style="padding:10px;"><a href="javascript:goTab('2');"><span style="font-size:11pt;<%=CHKIIF(vTab="2","font-weight:bold;text-decoration:underline;","")%>">[��޺� ��ȸ]</span></a></td>
			</tr>
		</table>
		<div class="simpleDesp">
			- <strong class="cBl1">��޺� ��ȸ ������</strong>�� ȸ��DB�� �����ϴ� �����ͷ� <strong class="cBl1">Ż���� ȸ���� �����ʹ� ��ȸ���� �ʽ��ϴ�.</strong> �׷��Ƿ� <strong>�ణ�� ��ġ ����</strong>�� ���� �� �ֽ��ϴ�.
		</div>
	</div>

	<!-- ��� �˻��� ���� -->
	<form name="frm1" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="tab" value="<%=vTab%>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<div class="searchWrap">
		<div class="search rowSum1">
			<ul>
				<li>
					<label class="formTit" for="term1">�Ⱓ :</label>
					<input type="text" class="formTxt" id="sDate" name="sDate" value="<%=sDate%>" style="width:100px" placeholder="������" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="sDate_trigger" alt="�޷����� �˻�" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "sDate", trigger    : "sDate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
					~
					<input type="text" class="formTxt" id="eDate" name="eDate" value="<%=eDate%>" style="width:100px" placeholder="������" maxlength="10" readonly />
					<img src="/images/admin_calendar.png" id="eDate_trigger" alt="�޷����� �˻�" />
					<script language="javascript">
						var CAL_Start = new Calendar({
							inputField : "eDate", trigger    : "eDate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
			</ul>
		</div>
		<input type="submit" class="schBtn" value="�˻�" />
	</div>
	</form>

	<% If vTab = "1" Then %>
	<div class="pad20">
		<div class="overHidden">
			<div class="ftRt">
				<p class="btn2 cBk1 ftLt"><a href="javascript:goExcelDown();"><span class="eIcon down"><em class="fIcon xls">����������</em></span></a></p>
			</div>
		</div>

		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div>�Ⱓ</div></th>
					<th><div>����</div></th>
					<th><div>����Ʈ ��</div></th>
					<th><div>����Ʈ ����</div></th>
					<th><div>����Ʈ ��</div></th>
					<th><div>�հ�</div></th>
				</tr>
				</thead>
				<tbody>
				<%
					for i=0 to cStat.FResultCount - 1
					
					vTotTalkPC	= vTotTalkPC + cStat.FItemList(i).FTalkWeb
					vTotTalkMob = vTotTalkMob + cStat.FItemList(i).FTalkMob
					vTotDayPC	= vTotDayPC + cStat.FItemList(i).FDayWeb
					vTotDayMob	= vTotDayMob + cStat.FItemList(i).FDayMob
					vTotShopPC	= vTotShopPC + cStat.FItemList(i).FShopWeb
					vTotShopMob	= ""
					
					vTotPC = cStat.FItemList(i).FTalkWeb + cStat.FItemList(i).FDayWeb + cStat.FItemList(i).FShopWeb
					vTotMob = cStat.FItemList(i).FTalkMob + cStat.FItemList(i).FDayMob
				%>
					<tr>
						<td rowspan="2"><%= cStat.FItemList(i).FDate %></td>
						<td>PC</td>
						<td><%= cStat.FItemList(i).FTalkWeb %></td>
						<td><%= cStat.FItemList(i).FDayWeb %></td>
						<td><%= cStat.FItemList(i).FShopWeb %></td>
						<td><%= vTotPC %></td>
					</tr>
					<tr>
						<td>�����</td>
						<td><%= cStat.FItemList(i).FTalkMob %></td>
						<td><%= cStat.FItemList(i).FDayMob %></td>
						<td><%= cStat.FItemList(i).FShopMob %></td>
						<td><%= vTotMob %></td>
					</tr>
				</tbody>
				<%
					next
				%>
				<tfoot>
					<tr>
						<td rowspan="2" class="bgBl1"><strong>�հ�</strong></td>
						<td class="bgBl1">PC</td>
						<td class="bgBl1"><strong><%= FormatNumber(vTotTalkPC,0) %></strong></td>
						<td class="bgBl1"><strong><%= FormatNumber(vTotDayPC,0) %></strong></td>
						<td class="bgBl1"><strong><%= FormatNumber(vTotShopPC,0) %></strong></td>
						<td class="bgBl1"><strong><%= FormatNumber((vTotTalkPC + vTotDayPC + vTotShopPC),0) %></strong></td>
					</tr>
					<tr>
						<td class="bgBl1">�����</td>
						<td class="bgBl1"><strong><%= FormatNumber(vTotTalkMob,0) %></strong></td>
						<td class="bgBl1"><strong><%= FormatNumber(vTotDayMob,0) %></strong></td>
						<td class="bgBl1"></td>
						<td class="bgBl1"><strong><%= FormatNumber((vTotTalkMob + vTotDayMob),0) %></strong></td>
					</tr>
					<tr>
						<td colspan="2" class="bgGy1"><strong>��������</strong></td>
						<td class="bgGy1"><strong><%= FormatNumber((vTotTalkPC + vTotTalkMob),0) %></strong></td>
						<td class="bgGy1"><strong><%= FormatNumber((vTotDayPC + vTotDayMob),0) %></strong></td>
						<td class="bgGy1"><strong><%= FormatNumber(vTotShopPC,0) %></strong></td>
						<td class="bgGy1"><strong><%= FormatNumber((vTotTalkPC + vTotTalkMob + vTotDayPC + vTotDayMob + vTotShopPC),0) %></strong></td>
					</tr>
				</tfoot>
			</table>
		</div>
	</div>
</div>
<% Else
	vTW5 = fnArrCount(vArrTalk,"w",5)
	vTW0 = fnArrCount(vArrTalk,"w",0)
	vTW1 = fnArrCount(vArrTalk,"w",1)
	vTW2 = fnArrCount(vArrTalk,"w",2)
	vTW3 = fnArrCount(vArrTalk,"w",3)
	vTW4 = fnArrCount(vArrTalk,"w",4)
	vTW6 = fnArrCount(vArrTalk,"w",6)
	vTW7 = fnArrCount(vArrTalk,"w",7)
	vTM5 = fnArrCount(vArrTalk,"m",5)
	vTM0 = fnArrCount(vArrTalk,"m",0)
	vTM1 = fnArrCount(vArrTalk,"m",1)
	vTM2 = fnArrCount(vArrTalk,"m",2)
	vTM3 = fnArrCount(vArrTalk,"m",3)
	vTM4 = fnArrCount(vArrTalk,"m",4)
	vTM6 = fnArrCount(vArrTalk,"m",6)
	vTM7 = fnArrCount(vArrTalk,"m",7)
	vDW5 = fnArrCount(vArrDay,"W",5)
	vDW0 = fnArrCount(vArrDay,"W",0)
	vDW1 = fnArrCount(vArrDay,"W",1)
	vDW2 = fnArrCount(vArrDay,"W",2)
	vDW3 = fnArrCount(vArrDay,"W",3)
	vDW4 = fnArrCount(vArrDay,"W",4)
	vDW6 = fnArrCount(vArrDay,"W",6)
	vDW7 = fnArrCount(vArrDay,"W",7)
	vDM5 = fnArrCount(vArrDay,"M",5)
	vDM0 = fnArrCount(vArrDay,"M",0)
	vDM1 = fnArrCount(vArrDay,"M",1)
	vDM2 = fnArrCount(vArrDay,"M",2)
	vDM3 = fnArrCount(vArrDay,"M",3)
	vDM4 = fnArrCount(vArrDay,"M",4)
	vDM6 = fnArrCount(vArrDay,"M",6)
	vDM7 = fnArrCount(vArrDay,"M",7)
	vSW5 = fnArrCount(vArrShop,"w",5)
	vSW0 = fnArrCount(vArrShop,"w",0)
	vSW1 = fnArrCount(vArrShop,"w",1)
	vSW2 = fnArrCount(vArrShop,"w",2)
	vSW3 = fnArrCount(vArrShop,"w",3)
	vSW4 = fnArrCount(vArrShop,"w",4)
	vSW6 = fnArrCount(vArrShop,"w",6)
	vSW7 = fnArrCount(vArrShop,"w",7)
	
	vTot5 = vTW5 + vTM5 + vDW5 + vDM5 + vSW5
	vTot0 = vTW0 + vTM0 + vDW0 + vDM0 + vSW0
	vTot1 = vTW1 + vTM1 + vDW1 + vDM1 + vSW1
	vTot2 = vTW2 + vTM2 + vDW2 + vDM2 + vSW2
	vTot3 = vTW3 + vTM3 + vDW3 + vDM3 + vSW3
	vTot4 = vTW4 + vTM4 + vDW4 + vDM4 + vSW4
	vTot6 = vTW6 + vTM6 + vDW6 + vDM6 + vSW6
	vTot7 = vTW7 + vTM7 + vDW7 + vDM7 + vSW7

	vTotTalkPC	= vTW5 + vTW0 + vTW1 + vTW2 + vTW3 + vTW4 + vTW6 + vTW7
	vTotTalkMob	= vTM5 + vTM0 + vTM1 + vTM2 + vTM3 + vTM4 + vTM6 + vTM7
	vTotDayPC	= vDW5 + vDW0 + vDW1 + vDW2 + vDW3 + vDW4 + vDW6 + vDW7
	vTotDayMob	= vDM5 + vDM0 + vDM1 + vDM2 + vDM3 + vDM4 + vDM6 + vDM7
	vTotShopPC	= vSW5 + vSW0 + vSW1 + vSW2 + vSW3 + vSW4 + vSW6 + vSW7
%>
<div class="pad20">
	<div class="overHidden">
		<div class="ftRt">
			<p class="btn2 cBk1 ftLt"><a href="javascript:goExcelDown();"><span class="eIcon down"><em class="fIcon xls">����������</em></span></a></p>
		</div>
	</div>

	<div class="tPad15">
		<table class="tbType1 listTb">
			<thead>
			<tr>
				<th><div>�Ⱓ</div></th>
				<th><div>����</div></th>
				<th><div>������</div></th>
				<th><div>���ο�</div></th>
				<th><div>�׸�</div></th>
				<th><div>���</div></th>
				<th><div>VIP<br />�ǹ�</div></th>
				<th><div>VIP<br />���</div></th>
				<th><div>VVIP</div></th>
				<th><div>����</div></th>
				<th><div>�հ�</div></th>
			</tr>
			</thead>
			<tbody>
			<tr>
				<td rowspan="2">����Ʈ��</td>
				<td>PC</td>
				<td><%= FormatNumber(vTW5,0) %></td>
				<td><%= FormatNumber(vTW0,0) %></td>
				<td><%= FormatNumber(vTW1,0) %></td>
				<td><%= FormatNumber(vTW2,0) %></td>
				<td><%= FormatNumber(vTW3,0) %></td>
				<td><%= FormatNumber(vTW4,0) %></td>
				<td><%= FormatNumber(vTW6,0) %></td>
				<td><%= FormatNumber(vTW7,0) %></td>
				<td><%= vTotTalkPC %></td>
			</tr>
			<tr>
				<td>�����</td>
				<td><%= FormatNumber(vTM5,0) %></td>
				<td><%= FormatNumber(vTM0,0) %></td>
				<td><%= FormatNumber(vTM1,0) %></td>
				<td><%= FormatNumber(vTM2,0) %></td>
				<td><%= FormatNumber(vTM3,0) %></td>
				<td><%= FormatNumber(vTM4,0) %></td>
				<td><%= FormatNumber(vTM6,0) %></td>
				<td><%= FormatNumber(vTM7,0) %></td>
				<td><%= vTotTalkMob %></td>
			</tr>
			<tr>
				<td rowspan="2">����Ʈ����</td>
				<td>PC</td>
				<td><%= FormatNumber(vDW5,0) %></td>
				<td><%= FormatNumber(vDW0,0) %></td>
				<td><%= FormatNumber(vDW1,0) %></td>
				<td><%= FormatNumber(vDW2,0) %></td>
				<td><%= FormatNumber(vDW3,0) %></td>
				<td><%= FormatNumber(vDW4,0) %></td>
				<td><%= FormatNumber(vDW6,0) %></td>
				<td><%= FormatNumber(vDW7,0) %></td>
				<td><%= vTotDayPC %></td>
			</tr>
			<tr>
				<td>�����</td>
				<td><%= FormatNumber(vDM5,0) %></td>
				<td><%= FormatNumber(vDM0,0) %></td>
				<td><%= FormatNumber(vDM1,0) %></td>
				<td><%= FormatNumber(vDM2,0) %></td>
				<td><%= FormatNumber(vDM3,0) %></td>
				<td><%= FormatNumber(vDM4,0) %></td>
				<td><%= FormatNumber(vDM6,0) %></td>
				<td><%= FormatNumber(vDM7,0) %></td>
				<td><%= vTotDayMob %></td>
			</tr>
			<tr>
				<td>����Ʈ��</td>
				<td>PC</td>
				<td><%= FormatNumber(vSW5,0) %></td>
				<td><%= FormatNumber(vSW0,0) %></td>
				<td><%= FormatNumber(vSW1,0) %></td>
				<td><%= FormatNumber(vSW2,0) %></td>
				<td><%= FormatNumber(vSW3,0) %></td>
				<td><%= FormatNumber(vSW4,0) %></td>
				<td><%= FormatNumber(vSW6,0) %></td>
				<td><%= FormatNumber(vSW7,0) %></td>
				<td><%= vTotShopPC %></td>
			</tr>
			</tbody>
			<tfoot>
			<tr>
				<td colspan="2" class="bgGy1"><strong>��������</strong></td>
				<td class="bgGy1"><strong><%= FormatNumber(vTot5,0) %></strong></td>
				<td class="bgGy1"><strong><%= FormatNumber(vTot0,0) %></strong></td>
				<td class="bgGy1"><strong><%= FormatNumber(vTot1,0) %></strong></td>
				<td class="bgGy1"><strong><%= FormatNumber(vTot2,0) %></strong></td>
				<td class="bgGy1"><strong><%= FormatNumber(vTot3,0) %></strong></td>
				<td class="bgGy1"><strong><%= FormatNumber(vTot4,0) %></strong></td>
				<td class="bgGy1"><strong><%= FormatNumber(vTot6,0) %></strong></td>
				<td class="bgGy1"><strong><%= FormatNumber(vTot7,0) %></strong></td>
				<td class="bgGy1"><strong><%= FormatNumber((vTotTalkPC + vTotTalkMob + vTotDayPC + vTotDayMob + vTotShopPC),0) %></strong></td>
			</tr>
			</tfoot>
		</table>
	</div>
</div>
<% End If %>

<% SET cStat = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->