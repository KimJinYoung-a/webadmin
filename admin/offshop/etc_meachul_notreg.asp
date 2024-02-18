<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ������ �������
' History : 2009.04.07 ������ ����
'			2010.05.13 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/etcmeachulcls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%

dim page
dim shopid, startdate, enddate, ctype, onlymifinish, research, gbybrand, makerid, exclude3pl
ctype = requestCheckVar(request("ctype"),10)
shopid = requestCheckVar(request("shopid"),32)
page = requestCheckVar(request("page"),10)
onlymifinish = requestCheckVar(request("onlymifinish"),2)
research = requestCheckVar(request("research"),2)
gbybrand = requestCheckVar(request("gbybrand"),2)
makerid  = requestCheckVar(request("makerid"),32)
exclude3pl = requestCheckVar(request("exclude3pl"),2)

if ctype="" then ctype = "M"
if page="" then page = 1
if (research="") and (onlymifinish="") then onlymifinish="on"

	dim nowdate, yyyy1,yyyy2,mm1,mm2,dd1,dd2
	dim tmpDate

yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1     = requestCheckVar(request("mm1"),2)
dd1     = requestCheckVar(request("dd1"),2)
yyyy2   = requestCheckVar(request("yyyy2"),4)
mm2     = requestCheckVar(request("mm2"),2)
dd2     = requestCheckVar(request("dd2"),2)

if (yyyy1="") then
	startdate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	enddate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)

	yyyy1 = Cstr(Year(startdate))
	mm1 = Cstr(Month(startdate))
	dd1 = Cstr(day(startdate))

	tmpDate = DateAdd("d", -1, enddate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
else
	startdate = DateSerial(yyyy1, mm1, dd1)
	enddate = DateSerial(yyyy2, mm2, dd2+1)
end if

startdate = Left(CStr(startdate), 10)
enddate = Left(CStr(enddate), 10)


'// ===========================================================================
dim oetcmeachul

set oetcmeachul = new CEtcMeachul
oetcmeachul.FPageSize=500
oetcmeachul.FCurrpage = page
oetcmeachul.FRectshopid = shopid
oetcmeachul.FRectStartDate = startdate
oetcmeachul.FRectEndDate = enddate
oetcmeachul.FRectonlymifinish = onlymifinish
oetcmeachul.FRectExclude3pl = exclude3pl
oetcmeachul.FRectGroupByBrand = gbybrand
oetcmeachul.FRectMakerid = makerid
oetcmeachul.FRectCType = ctype

if ctype="M" or ctype="M_ETC" then
''OFF ��������
	oetcmeachul.getChulgoJungsanTargetListNotReg
elseif ctype="W" then
''OFF �Ǹź�����(������)
	oetcmeachul.getWitakSellJungsanTargetListNotReg
elseif ctype="A" then
''OFF �Ǹź�����(������)
	oetcmeachul.getOfflineIpjumshopMaechulListNotReg
	''response.end
elseif ctype="B" then
''ON �Ǹź�����(������,��ۺ�����)
	oetcmeachul.FRectRemoveDlvPay ="on"
	oetcmeachul.getOnlineIpjumshopMaechulListNotReg
elseif ctype="P" then
''ON �Ǹź�����(������,��ۺ�����)
	oetcmeachul.FRectRemoveDlvPay =""
	oetcmeachul.getOnlineIpjumshopMaechulListNotReg
elseif ctype="C" then
''ON ��ۺ�����(������)
	oetcmeachul.getOnlineIpjumshopBeasongPayMaechulListNotReg
end if


'// ===========================================================================
dim opartner
dim bizsection_cd, selltype, papertype, b2bcharge, paperissuetype
dim shopdiv, shopname

b2bcharge = requestCheckVar(request("b2bcharge"),20)

set opartner = new CPartnerUser
opartner.FRectDesignerID = shopid

if (shopid <> "") then
	opartner.GetOnePartnerNUser

	selltype = opartner.FOneItem.Fselltype

	shopname = fnGetShopName(shopid, shopdiv, papertype)

	if (shopdiv = "7") then
		'������ : ����
		papertype = "200"
	elseif (shopdiv = "9") then
		'������ : ����
		papertype = "102"
	elseif (shopdiv <> "7") and (shopdiv <> "9") then
		papertype = "100"
	end if

	'// ����Ʈ ������
	paperissuetype = "1"
	if Not IsNull(opartner.FOneItem.Ftaxevaltype) then
		if (opartner.FOneItem.Ftaxevaltype = "1") then
			'// ������
			paperissuetype = "2"
		end if
	end if

	if (bizsection_cd = "") and (Not IsNull(opartner.FOneItem.FsellBizCd)) then
		bizsection_cd = opartner.FOneItem.FsellBizCd
	end if

	if (b2bcharge = "") and (Not IsNull(opartner.FOneItem.getCommissionPro)) then
		b2bcharge = opartner.FOneItem.getCommissionPro
		if (ctype = "C") then
			'// ��ۺ�����
			b2bcharge = 0
		end if
	end if
end if

'// ===========================================================================
dim IsRegAvail	: IsRegAvail = True
dim ErrMSG

if (shopid = "") then
	IsRegAvail = False
	ErrMSG = "���� ����ó�� �����ϼ���."
end if

if (IsRegAvail = True) and IsNull(selltype) and ctype <> "M_ETC" then
	IsRegAvail = False
	ErrMSG = "���� �귣���������� ���������� �����ϼ���."
else
	if (IsRegAvail = True) and (selltype = "") or (selltype = "0") then
		IsRegAvail = False
		ErrMSG = "���� �귣���������� ���������� �����ϼ���."
	end if
end if

if (IsRegAvail = True) and ((selltype = "20166") or (selltype = "20032")) then
	'// B2C(20166), ����B2C(20032)
	IsRegAvail = False
	ErrMSG = "B2C �����Դϴ�. ����� �� �����ϴ�."
end if

'// ===========================================================================
dim i
dim ttlorgsell, ttlsell,ttlsuply,ttlbuy
dim ttlorgsell_dlv,ttlsell_dlv,ttlsuply_dlv,ttlbuy_dlv

ttlorgsell = 0
ttlsell = 0
ttlsuply = 0
ttlbuy = 0

ttlorgsell_dlv = 0
ttlsell_dlv = 0
ttlsuply_dlv = 0
ttlbuy_dlv = 0

%>
<script language='javascript'>
function reCalcuSum(frm){
	var suplysum = 0;

	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.type=="checkbox") && (e.name != "chk")) {
			if (e.checked){
				suplysum = suplysum + eval("frm.val_" + e.value).value*1;
			}
		}
	}

	document.buffrm.totalsuply.value = suplysum;
}

function PopChulgoSheet(v){
	var popwin;
	popwin = window.open('/admin/newstorage/popchulgosheet.asp?idx=' + v ,'popchulgosheet','width=760,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopIpgoSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/popshopjumunsheet2.asp?idx=' + v ,'shopjumunsheet','width=740,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopExportSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/cartoonbox_modify.asp?menupos=1357&idx=' + v ,'PopExportSheet','width=740,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function editOffDesinger(shopid,designerid){
	var popwin = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + designerid,"popshopupcheinfo","width=640,height=540,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function PopDetailFranWitakSell(iidx,shopid){
	var popwin = window.open("/admin/offupchejungsan/off_jungsandetailsum.asp?gubuncd=B012&idx=" + iidx + '&shopid=' + shopid,"PopDetailFranWitakSell","width=1000, height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function popOfflineIpjumshopItemDetail(yyyy1,mm1,dd1,shopid){
	var popOfflineIpjumshopItemDetail = window.open('/admin/offshop/todayselldetail.asp?menupos=648&oldlist=&datefg=maechul&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy1+'&mm2='+mm1+'&dd2='+dd1+'&shopid='+shopid,'popOfflineIpjumshopItemDetail','width=1024,height=768,scrollbars=yes,resizable=yes');
	popOfflineIpjumshopItemDetail.focus();
}

function popOfflineIpjumshopJumunDetail(yyyy1,mm1,dd1,shopid){
	var popOfflineIpjumshopJumunDetail = window.open('/admin/offshop/todaysellmaster.asp?menupos=648&oldlist=&datefg=maechul&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy1+'&mm2='+mm1+'&dd2='+dd1+'&shopid='+shopid,'popOfflineIpjumshopJumunDetail','width=1024,height=768,scrollbars=yes,resizable=yes');
	popOfflineIpjumshopJumunDetail.focus();
}

function popOnlineIpjumshopItemDetail(yyyy1,mm1,dd1,shopid){
	var popOnlineIpjumshopItemDetail = window.open('/admin/upchejungsan/upcheselllist.asp?menupos=138&datetype=chulgoil&delivertype=all&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy1+'&mm2='+mm1+'&dd2='+dd1+'&sitename='+shopid,'popOnlineIpjumshopItemDetail','width=1100,height=768,scrollbars=yes,resizable=yes');
	popOnlineIpjumshopItemDetail.focus();
}

function SaveArr(frm){
	var ischecked = false;
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];

		if ((e.type=="checkbox")) {
			ischecked = (ischecked || e.checked);
			if (ischecked) break;
		}
	}

	<% if shopdiv = "7" then %>
		if ((frm.papertype.value != "200") && (frm.papertype.value != "102")) {
			alert("���ó������ ����(�ؿ�)�ΰ�� \n\n����Ű����� �Ǵ� ������꼭�� ���������� ����� �� �ֽ��ϴ�.");
			frm.papertype.focus();
			return;
		}

		if (frm.selltype.value != "20036") {
			alert("���ó������ ����(�ؿ�)�ΰ�� \n\n���������� ������ �����մϴ�.");
			frm.selltype.focus();
			return;
		}
	<% elseif shopdiv = "9" then %>
		if (frm.papertype.value != "102") {
			alert("���ó������ �����ΰ�� \n\n������꼭�� ���������� ����� �� �ֽ��ϴ�.");
			frm.papertype.focus();
			return;
		}

		if (frm.selltype.value != "20036") {
			alert("���ó������ �����ΰ�� \n\n���������� ������ �����մϴ�.");
			frm.selltype.focus();
			return;
		}
	<% else %>
		if ((frm.papertype.value == "200") || (frm.papertype.value == "102")) {
			alert("���ó������ ���� �Ǵ� �����ΰ�츸 ��� �����մϴ�.");
			frm.papertype.focus();
			return;
		}

		if (frm.selltype.value == "20036") {
			alert("���ó������ ���� �Ǵ� �����ΰ�츸 ��� �����մϴ�.");
			frm.selltype.focus();
			return;
		}
	<% end if %>

	if (!ischecked) {
		alert('���� ������ �����ϴ�.');
		return;
	}

	var val_workidx = "-";
	var is_multiworkidx = false;

	<% if (ctype="M") then %>
		for (var i=0;i<frm.elements.length;i++){
			var e = frm.elements[i];

			if ((e.type=="checkbox")) {
				if ((e.checked)&&(e.value!="on")){
					if (val_workidx == "-") {
						val_workidx = eval("frmArr.val_workidx_" + e.value).value;
					}

					if (eval("frmArr.val_workidx_" + e.value).value != val_workidx) {
						is_multiworkidx = true;
						val_workidx = eval("frmArr.val_workidx_" + e.value).value;
					}
				}
			}
		}

		if (is_multiworkidx == true) {
			if (confirm("�̹� �ٸ� �ؿ����� ������ �ֹ��� �ֽ��ϴ�.\n\n�ؿ����(IDX : " + val_workidx + ") �� �ϰ� �����Ͻðڽ��ϱ�?") != true) {
				return;
			} else {
				// val_workidx = "";
			}
		}
	<% end if %>

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		if ((val_workidx != "") && (val_workidx != "-")) {
			frm.workidx.value = val_workidx;
		}

		frm.submit();
	}
}

function SubmitSearch(frm) {
	frm.submit();
}

function CheckAll(obj) {
	var arrObj = document.getElementsByName("check");

	for (var i = 0; i < arrObj.length; i++) {
		if (arrObj[i].disabled != true) {
			arrObj[i].checked = obj.checked;
			AnCheckClick(arrObj[i]);
		}
	}

	<% if (ctype="M") then %>
		reCalcuSum(frmArr);
	<% end if %>
}

function popEtcMeachulReg(shopid) {
	var popwin = window.open("/admin/offshop/popetcmeachulreg.asp?ctype=<%= ctype %>&shopid=" + shopid + "&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>&dd1=<%= dd1 %>&yyyy2=<%= yyyy2 %>&mm2=<%= mm2 %>&dd2=<%= dd2 %>&onlymifinish=<%= onlymifinish %>","popEtcMeachulReg","width=1100,height=768,scrollbars=yes,resizable=yes");
	popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="b2bcharge" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left" height="30">
		<label><input type=radio name=ctype value="M" <% if ctype="M" then response.write "checked" %> onClick="SubmitSearch(frm)"> ��������</label>
		<label><input type=radio name=ctype value="M_ETC" <% if ctype="M_ETC" then response.write "checked" %> onClick="SubmitSearch(frm)"> ��������(��Ÿ)</label>
		<label><input type=radio name=ctype value="W" <% if ctype="W" then response.write "checked" %> onClick="SubmitSearch(frm)"> �Ǹź�����(������)</label>
		<label><input type=radio name=ctype value="A" <% if ctype="A" then response.write "checked" %> onClick="SubmitSearch(frm)"> �Ǹź�����(���� ������)</label>
		<label><input type=radio name=ctype value="B" <% if ctype="B" then response.write "checked" %> onClick="SubmitSearch(frm)"> �Ǹź�����(�� ������, ��ۺ� ����)</label>
		<label><input type=radio name=ctype value="C" <% if ctype="C" then response.write "checked" %> onClick="SubmitSearch(frm)"> ��ۺ�����(�� ������)</label>
		<label><input type=radio name=ctype value="P" <% if ctype="P" then response.write "checked" %> onClick="SubmitSearch(frm)"> <font color="gray">�Ǹź�����(�� ������)</font></label>
		&nbsp;&nbsp;
		<% if (ctype<>"B" and ctype<>"C" and ctype<>"P") then %>
		<label><input type="checkbox" name="gbybrand" disabled >�귣�庰</label>
		<% else %>
		<label><input type="checkbox" name="gbybrand" <%=CHKIIF(gbybrand="on","checked","") %> >�귣�庰</label>
		<% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		<% end if %>
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitSearch(frm)">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="30">
		����� / �Ǹ��� :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		<label><input type="checkbox" name="onlymifinish" <% if onlymifinish="on" then response.write "checked" %> > �̵�� ������</label>
		<% if ctype="M" or ctype="M_ETC" then %>
		&nbsp;
		<label><input type="checkbox" name="exclude3pl" <% if exclude3pl="on" then response.write "checked" %> > 3PL ����</label>
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p>

<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
<% if (ctype="M") or (ctype="M_ETC") then %>
	<input type=hidden name="mode" value="chulgo">
	<input type=hidden name="ctype" value="<%= ctype %>">
	<input type=hidden name="makerid" value="<%= makerid %>">
	<tr bgcolor="#DDDDFF" align=center>
		<td width=90>���ó</td>
		<td width=90>���ó��</td>
		<td width=90>����μ�</td>
		<td width=70>�����</td>
		<td width=70>����ڵ�<br>�ֹ��ڵ�</td>
		<td width=70>�ֹ���/<br>������</td>
		<td width=70>�Һ��ڰ�</td>
		<td width=70>���ǸŰ�</td>
		<td width=70><b>�����</b></td>
		<td width=70>�Ѹ��԰�</td>
		<td width=40>����</td>
		<td width=64>�����</td>
		<td>���</td>
	</tr>
	<% for i=0 to oetcmeachul.FResultCount-1 %>
	<input type="hidden" name="val_<%= oetcmeachul.FItemList(i).Fid %>" value="<%= oetcmeachul.FItemList(i).Fjumunrealsuplycash %>">
	<%
	ttlsell = ttlsell + oetcmeachul.FItemList(i).Ftotalsellcash
	ttlsuply = ttlsuply + oetcmeachul.FItemList(i).Ftotalsuplycash
	ttlbuy = ttlbuy + oetcmeachul.FItemList(i).Ftotalbuycash
	%>
	<tr bgcolor="#FFFFFF">
		<td ><a href="javascript:popEtcMeachulReg('<%= oetcmeachul.FItemList(i).FSocid %>')"><%= oetcmeachul.FItemList(i).FSocid %></a></td>
		<td ><%= oetcmeachul.FItemList(i).Fshopname %></td>
		<td ><%= oetcmeachul.FItemList(i).Fbizsection_nm %></td>
		<td align=center><%= oetcmeachul.FItemList(i).Fexecutedt %>
			<% if oetcmeachul.FItemList(i).Fexecutedt<>oetcmeachul.FItemList(i).FIpgodate then %>
			<br><font color=red>(<%= oetcmeachul.FItemList(i).FIpgodate %>)</font>
			<% end if %>
		</td>
		<td align=center>
			<a href="javascript:PopChulgoSheet('<%= oetcmeachul.FItemList(i).Fid %>')"><%= oetcmeachul.FItemList(i).Fcode %></a><br>
			<a href="javascript:PopIpgoSheet('<%= oetcmeachul.FItemList(i).Fbaljuidx %>')">
				<font color="<% if (oetcmeachul.FItemList(i).FOrderStateCD < "7") then %>red<% else %>gray<% end if %>"><%= oetcmeachul.FItemList(i).Fbaljucode %><%= oetcmeachul.FItemList(i).FOrderStateCD %></font>
			</a>
		</td>
		<td align=center><%= Left(oetcmeachul.FItemList(i).FjumunRegDate,10) %><br><%= oetcmeachul.FItemList(i).Fscheduledate %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotalsellcash,0) %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotalsellcash,0) %>
			<% if oetcmeachul.FItemList(i).Ftotalsellcash<>oetcmeachul.FItemList(i).Fjumunrealsellcash then %>
			<br><font color=red>(<%= FormatNumber(oetcmeachul.FItemList(i).Fjumunrealsellcash,0) %>)</font>
			<% end if %>
		</td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotalsuplycash,0) %>
			<% if oetcmeachul.FItemList(i).Ftotalsuplycash<>oetcmeachul.FItemList(i).Fjumunrealsuplycash then %>
			<br><font color=red>(<%= FormatNumber(oetcmeachul.FItemList(i).Fjumunrealsuplycash,0) %>)</font>
			<% end if %>
		</td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotalbuycash,0) %>
			<% if oetcmeachul.FItemList(i).Ftotalbuycash<>oetcmeachul.FItemList(i).Fjumunrealbuycash then %>
			<br><font color=red>(<%= FormatNumber(oetcmeachul.FItemList(i).Fjumunrealbuycash,0) %>)</font>
			<% end if %>
		</td>
		<td align=center>
			<% if not IsNULL(oetcmeachul.FItemList(i).Fprecheckidx) then %>
			<%= oetcmeachul.FItemList(i).Fprecheckmasteridx %>
			<% end if %>
		</td>
		<td align=center>
			<input type="button" class="button" value="��ȸ" onClick="PopChulgoSheet('<%= oetcmeachul.FItemList(i).Fid %>')">
		</td>
		<td>
			<input type="hidden" name="val_workidx_<%= oetcmeachul.FItemList(i).Fid %>" value="<%= oetcmeachul.FItemList(i).Fworkidx %>">
			<% if (oetcmeachul.FItemList(i).Fworkidx <> "") then %>
				�ؿ� IDX : <a href="javascript:PopExportSheet(<%= oetcmeachul.FItemList(i).Fworkidx %>)"><%= oetcmeachul.FItemList(i).Fworkidx %></a>
			<% end if %>
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td align=right><%= formatnumber(ttlsell,0) %></td>
		<td align=right><%= formatnumber(ttlsell,0) %></td>
		<td align=right><%= formatnumber(ttlsuply,0) %></td>
		<td align=right><%= formatnumber(ttlbuy,0) %></td>
		<td></td>
		<td></td>
		<td ></td>
	</tr>
<% elseif (ctype="W") then %>

	<input type=hidden name="mode" value="witsksell">
	<input type=hidden name="ctype" value="<%= ctype %>">
	<input type=hidden name="makerid" value="<%= makerid %>">
	<tr bgcolor="#DDDDFF" align=center height="30">
		<td width=90>���óID</td>
		<td width=90>���ó��</td>
		<td width=90>����μ�</td>
		<td width=60>�����<br>(����)</td>
		<td width=90>�귣��</td>
		<td width=40>�ѰǼ�</td>
		<td width=70>�ѼҺ��ڰ�</td>
		<td width=70>���ǸŰ�</td>
		<td width=70><b>�����</b></td>
		<td width=70>�Ѹ��԰�</td>
		<td width=40>��ó��</td>
		<td width=60>�󼼳���</td>
		<td>���</td>
	</tr>
	<% for i=0 to oetcmeachul.FResultCount-1 %>
	<%
	ttlorgsell = ttlorgsell + oetcmeachul.FItemList(i).Ftotorgsum
	ttlsell = ttlsell + oetcmeachul.FItemList(i).Ftotsum
	ttlbuy = ttlbuy + oetcmeachul.FItemList(i).Frealjungsansum
	ttlsuply = ttlsuply + 0
	%>
	<tr bgcolor="#FFFFFF" height="30">
		<td ><a href="javascript:popEtcMeachulReg('<%= oetcmeachul.FItemList(i).Fshopid %>')"><%= oetcmeachul.FItemList(i).Fshopid %></a></td>
		<td ><%= oetcmeachul.FItemList(i).Fshopname %></td>
		<td ><%= oetcmeachul.FItemList(i).Fbizsection_nm %></td>
		<td align=center><a href="javascript:PopDetailFranWitakSell('<%= oetcmeachul.FItemList(i).Fidx %>','<%= oetcmeachul.FItemList(i).Fshopid %>');"><%= oetcmeachul.FItemList(i).FYYYYMM %></a></td>
		<td ><a href="javascript:editOffDesinger('<%= oetcmeachul.FItemList(i).Fshopid %>','<%= oetcmeachul.FItemList(i).Fjungsanid %>');"><%= oetcmeachul.FItemList(i).Fjungsanid %></a></td>

		<td align=center><%= oetcmeachul.FItemList(i).Ftotitemcnt %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotorgsum,0) %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotsum,0) %></td>
		<td align=right></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Frealjungsansum,0) %> </td>
		<td ><%= oetcmeachul.FItemList(i).Fprecheckidx %></td>
		<td align=center><input type="button" class="button" value="��ȸ" onClick="PopDetailFranWitakSell('<%= oetcmeachul.FItemList(i).Fidx %>','<%= oetcmeachul.FItemList(i).Fshopid %>')"></td>
		<td>
			<input type="hidden" name="val_workidx_<%= oetcmeachul.FItemList(i).Fidx %>" value="">
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF" height="30">
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td align=right><%= formatnumber(ttlorgsell,0) %></td>
		<td align=right><%= formatnumber(ttlsell,0) %></td>
		<td align=right><%= formatnumber(ttlsuply,0) %></td>
		<td align=right><%= formatnumber(ttlbuy,0) %></td>
		<td ></td>
		<td></td>
		<td></td>
	</tr>
<% elseif (ctype="A") then %>
	<input type=hidden name="mode" value="offipjumshop">
	<input type=hidden name="ctype" value="<%= ctype %>">
	<input type=hidden name="makerid" value="<%= makerid %>">
	<tr bgcolor="#DDDDFF" align=center height="30">
		<td width=90>���óID</td>
		<td>���ó��</td>
		<td>����μ�</td>
		<td width=80>������</td>
		<td width=40>�ѰǼ�</td>
		<td width=70>�ѼҺ��ڰ�</td>
		<td width=70>���ǸŰ�</td>
		<td width=70><b>�����</b></td>
		<td width=70>�Ѹ��԰�</td>
		<td width=40>��ó��</td>
		<td width=125>�󼼳���</td>
		<td>���</td>
	</tr>
	<% for i=0 to oetcmeachul.FResultCount-1 %>
	<%
	ttlorgsell = ttlorgsell + oetcmeachul.FItemList(i).Ftotorgsum
	ttlsell = ttlsell + oetcmeachul.FItemList(i).Ftotsum
	ttlbuy = ttlbuy + oetcmeachul.FItemList(i).Frealjungsansum
	''ttlsuply = ttlsuply + CLng(oetcmeachul.FItemList(i).Ftotsum * (100 - b2bcharge) / 100)
	%>
	<tr bgcolor="#FFFFFF" height="30">
		<td ><a href="javascript:popEtcMeachulReg('<%= oetcmeachul.FItemList(i).Fshopid %>')"><%= oetcmeachul.FItemList(i).Fshopid %></a></td>
		<td ><%= oetcmeachul.FItemList(i).Fshopname %></td>
		<td ><%= oetcmeachul.FItemList(i).Fbizsection_nm %></td>
		<td align=center><%= oetcmeachul.FItemList(i).Fyyyymmdd %></td>

		<td align=center><%= oetcmeachul.FItemList(i).Ftotitemcnt %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotorgsum,0) %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotsum,0) %></td>
		<td align=right>

		</td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Frealjungsansum,0) %> </td>
		<td ><%= oetcmeachul.FItemList(i).Fprecheckidx %></td>
		<td align=center>
			<input type="button" class="button" value="�ֹ���" onClick="popOfflineIpjumshopJumunDetail('<%= Left(oetcmeachul.FItemList(i).Fyyyymmdd, 4) %>','<%= Mid(oetcmeachul.FItemList(i).Fyyyymmdd, 6, 2) %>','<%= Right(oetcmeachul.FItemList(i).Fyyyymmdd, 2) %>','<%= oetcmeachul.FItemList(i).Fshopid %>');">
			<input type="button" class="button" value="��ǰ��" onClick="popOfflineIpjumshopItemDetail('<%= Left(oetcmeachul.FItemList(i).Fyyyymmdd, 4) %>','<%= Mid(oetcmeachul.FItemList(i).Fyyyymmdd, 6, 2) %>','<%= Right(oetcmeachul.FItemList(i).Fyyyymmdd, 2) %>','<%= oetcmeachul.FItemList(i).Fshopid %>');">
		</td>
		<td>
			<input type="hidden" name="val_workidx_<%= oetcmeachul.FItemList(i).Fidx %>" value="">
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF" height="30">
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td align=right><%= formatnumber(ttlorgsell,0) %></td>
		<td align=right><%= formatnumber(ttlsell,0) %></td>
		<td align=right><%= formatnumber(ttlsuply,0) %></td>
		<td align=right><%= formatnumber(ttlbuy,0) %></td>
		<td ></td>
		<td></td>
		<td></td>
	</tr>
<% elseif (ctype="B") or (ctype="P") then %>
	<input type=hidden name="mode" value="onipjumshop">
	<input type=hidden name="ctype" value="<%= ctype %>">
	<input type=hidden name="makerid" value="<%= makerid %>">
	<tr bgcolor="#DDDDFF" align=center height="30">
		<td width=90>���óID</td>
		<td>���ó��</td>
		<td>����μ�</td>
		<% if (gbybrand<>"") then %>
		<td width=100>�귣��</td>
		<% end if %>
		<td width=80>������</td>

		<td width=40>�ѰǼ�</td>
		<td width=70>�ѼҺ��ڰ�</td>
		<td width=70>���ǸŰ�</td>
		<td width=70><b>�����</b></td>
		<td width=70>�Ѹ��԰�</td>
		<td width=40>��ó��</td>
		<td width=80>�󼼳���</td>
		<td>���</td>
	</tr>
	<% for i=0 to oetcmeachul.FResultCount-1 %>
	<%
	ttlorgsell = ttlorgsell + oetcmeachul.FItemList(i).Ftotorgsum
	ttlsell = ttlsell + oetcmeachul.FItemList(i).Ftotsum
	ttlbuy = ttlbuy + oetcmeachul.FItemList(i).fbuyprice
	''ttlsuply = ttlsuply + CLng(oetcmeachul.FItemList(i).Ftotsum * (100 - b2bcharge) / 100)

	ttlorgsell_dlv  = ttlorgsell_dlv + oetcmeachul.FItemList(i).Ftotdeliverorgsum
    ttlsell_dlv     = ttlsell_dlv + oetcmeachul.FItemList(i).Ftotdeliversum
    ttlsuply_dlv    = ttlsuply_dlv + CLng(oetcmeachul.FItemList(i).Ftotdeliversum * (100 - 0) / 100)
    ttlbuy_dlv      = ttlbuy_dlv + oetcmeachul.FItemList(i).Fbuydeliverprice

	%>
	<tr bgcolor="#FFFFFF" height="30">
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ><a href="javascript:popEtcMeachulReg('<%= oetcmeachul.FItemList(i).Fshopid %>')"><%= oetcmeachul.FItemList(i).Fshopid %></a></td>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ><%= oetcmeachul.FItemList(i).Fshopname %></td>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ><%= oetcmeachul.FItemList(i).Fbizsection_nm %></td>
		<% if (gbybrand<>"") then %>
		    <td <%= CHKIIF(ctype="P","rowspan=2","") %> align=center><%= oetcmeachul.FItemList(i).Fmakerid %></td>
		<% end if %>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> align=center><%= oetcmeachul.FItemList(i).Fyyyymmdd %></td>
		<td align=center><%= oetcmeachul.FItemList(i).Ftotitemcnt %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotorgsum,0) %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotsum,0) %></td>
		<td align=right>

		</td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).fbuyprice,0) %> </td>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> align=center><%= oetcmeachul.FItemList(i).Fprecheckidx %></td>
		<td align=center>
			<input type="button" class="button" value="��ǰ��" onClick="popOnlineIpjumshopItemDetail('<%= Left(oetcmeachul.FItemList(i).Fyyyymmdd, 4) %>','<%= Mid(oetcmeachul.FItemList(i).Fyyyymmdd, 6, 2) %>','<%= Right(oetcmeachul.FItemList(i).Fyyyymmdd, 2) %>','<%= oetcmeachul.FItemList(i).Fshopid %>');">
		</td>
		<td >
			<input type="hidden" name="val_workidx_<%= oetcmeachul.FItemList(i).Fidx %>" value="">
		</td>
	</tr>
	<% if (ctype="P") then %>
	<tr bgcolor="#FFFFFF" height="30">
	    <td align=center><%= oetcmeachul.FItemList(i).Ftotdeliveritemcnt %></td>
	    <td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotdeliverorgsum,0) %></td>
	    <td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotdeliversum,0) %></td>
	    <td align=right><%= FormatNumber(CLng(oetcmeachul.FItemList(i).Ftotdeliversum * (100 - 0) / 100),0) %></td>
	    <td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Fbuydeliverprice,0) %></td>
	    <td align=center></td>
	    <td align=center>��ۺ�</td>
	</tr>
	<% end if %>
	<% next %>
	<tr bgcolor="#FFFFFF" height="30">
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ></td>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ></td>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ></td>
		<% if (gbybrand<>"") then %>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %>></td>
		<% end if %>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ></td>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ></td>
		<td align=right><%= formatnumber(ttlorgsell,0) %></td>
		<td align=right><%= formatnumber(ttlsell,0) %></td>
		<td align=right><%= formatnumber(ttlsuply,0) %></td>
		<td align=right><%= formatnumber(ttlbuy,0) %></td>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ></td>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ></td>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ></td>
	</tr>
	<% if (ctype="P") then %>
	<tr bgcolor="#FFFFFF" height="30">
	    <td align=right><%= formatnumber(ttlorgsell_dlv,0) %></td>
		<td align=right><%= formatnumber(ttlsell_dlv,0) %></td>
		<td align=right><%= formatnumber(ttlsuply_dlv,0) %></td>
		<td align=right><%= formatnumber(ttlbuy_dlv,0) %></td>
	</tr>
	<% end if %>
<% elseif (ctype="C") then %>
	<input type=hidden name="mode" value="onipjumshopbeasongpay">
	<input type=hidden name="ctype" value="<%= ctype %>">
	<input type=hidden name="makerid" value="<%= makerid %>">
	<tr bgcolor="#DDDDFF" align=center height="30">
		<td width=90>���óID</td>
		<td>���ó��</td>
		<td>����μ�</td>
		<td width=80>������</td>
		<td width=40>�ѰǼ�</td>
		<td width=70>�ѼҺ��ڰ�</td>
		<td width=70>���ǸŰ�</td>
		<td width=70><b>�����</b></td>
		<td width=70>�Ѹ��԰�</td>
		<td width=40>��ó��</td>
		<td width=80>�󼼳���</td>
		<td>���</td>
	</tr>
	<% for i=0 to oetcmeachul.FResultCount-1 %>
	<%
	ttlorgsell = ttlorgsell + oetcmeachul.FItemList(i).Ftotdeliverorgsum
	ttlsell = ttlsell + oetcmeachul.FItemList(i).Ftotdeliversum
	ttlbuy = ttlbuy + oetcmeachul.FItemList(i).Fbuydeliverprice
	''ttlsuply = ttlsuply + CLng(oetcmeachul.FItemList(i).Ftotdeliversum * (100 - b2bcharge) / 100)
	%>
	<tr bgcolor="#FFFFFF" height="30">
		<td ><a href="javascript:popEtcMeachulReg('<%= oetcmeachul.FItemList(i).Fshopid %>')"><%= oetcmeachul.FItemList(i).Fshopid %></a></td>
		<td ><%= oetcmeachul.FItemList(i).Fshopname %></td>
		<td ><%= oetcmeachul.FItemList(i).Fbizsection_nm %></td>
		<td align=center><%= oetcmeachul.FItemList(i).Fyyyymmdd %></td>

		<td align=center><%= oetcmeachul.FItemList(i).Ftotdeliveritemcnt %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotdeliverorgsum,0) %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotdeliversum,0) %></td>
		<td align=right>

		</td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Fbuydeliverprice,0) %> </td>
		<td ><%= oetcmeachul.FItemList(i).Fprecheckidx %></td>
		<td align=center>
			<!--
			<input type="button" class="button" value="��ǰ��" onClick="popOnlineIpjumshopItemDetail('<%= Left(oetcmeachul.FItemList(i).Fyyyymmdd, 4) %>','<%= Mid(oetcmeachul.FItemList(i).Fyyyymmdd, 6, 2) %>','<%= Right(oetcmeachul.FItemList(i).Fyyyymmdd, 2) %>','<%= oetcmeachul.FItemList(i).Fshopid %>');">
			-->
		</td>
		<td>
			<input type="hidden" name="val_workidx_<%= oetcmeachul.FItemList(i).Fidx %>" value="">
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF" height="30">
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td ></td>
		<td align=right><%= formatnumber(ttlorgsell,0) %></td>
		<td align=right><%= formatnumber(ttlsell,0) %></td>
		<td align=right><%= formatnumber(ttlsuply,0) %></td>
		<td align=right><%= formatnumber(ttlbuy,0) %></td>
		<td ></td>
		<td></td>
		<td></td>
	</tr>
<% else %>
<tr><td>1111</td></tr>
<% end if %>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
