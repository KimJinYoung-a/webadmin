<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ������
' History : ������ ����
'			2017.04.12 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/etcmeachulcls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%

dim page
dim shopid, startdate, enddate, ctype, onlymifinish, research, gbybrand, makerid
ctype = requestCheckVar(request("ctype"),10)
shopid = requestCheckVar(request("shopid"),32)
page = requestCheckVar(request("page"),10)
onlymifinish = requestCheckVar(request("onlymifinish"),2)
research = requestCheckVar(request("research"),2)
gbybrand = requestCheckVar(request("gbybrand"),2)
makerid  = requestCheckVar(request("makerid"),32)

if ctype="" then ctype = "M"
if page="" then page = 1
if (research="") and (onlymifinish="") then onlymifinish="on"

dim nowdate, yyyy1,yyyy2,mm1,mm2,dd1,dd2

yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1     = requestCheckVar(request("mm1"),2)
dd1     = requestCheckVar(request("dd1"),2)
yyyy2   = requestCheckVar(request("yyyy2"),4)
mm2     = requestCheckVar(request("mm2"),2)
dd2     = requestCheckVar(request("dd2"),2)

if (yyyy1="") then
	nowdate = Left(CStr(now()),10)

	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2) - 3
	dd1   = "01" ''Mid(nowdate,9,2)

	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

startdate   = CStr(DateSerial(yyyy1 , mm1 , dd1))
enddate     = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)


'// ===========================================================================
dim oetcmeachul

set oetcmeachul = new CEtcMeachul
oetcmeachul.FPageSize=500
oetcmeachul.FCurrpage = page
oetcmeachul.FRectshopid = shopid
oetcmeachul.FRectStartDate = startdate
oetcmeachul.FRectEndDate = enddate
oetcmeachul.FRectonlymifinish = onlymifinish
oetcmeachul.FRectGroupByBrand = gbybrand
oetcmeachul.FRectMakerid = makerid

if shopid<>"" then
	if ctype="M" or ctype="M_ETC" then
	''OFF ��������
		oetcmeachul.getChulgoJungsanTargetList
	elseif ctype="W" then
	''OFF �Ǹź�����(������)
		oetcmeachul.getWitakSellJungsanTargetList
	elseif ctype="A" then
	''OFF �Ǹź�����(������)
		oetcmeachul.getOfflineIpjumshopMaechulList
	elseif ctype="B" then
	''ON �Ǹź�����(������,��ۺ�����)
	    oetcmeachul.FRectRemoveDlvPay ="on"
		oetcmeachul.getOnlineIpjumshopMaechulList
	elseif ctype="P" then
	''ON �Ǹź�����(������,��ۺ�����)
	    oetcmeachul.FRectRemoveDlvPay =""
		oetcmeachul.getOnlineIpjumshopMaechulList
	elseif ctype="C" then
	''ON ��ۺ�����(������)
		oetcmeachul.getOnlineIpjumshopBeasongPayMaechulList
	end if
end if


'// ===========================================================================
'���ͺμ����
Dim clsBS, arrBizList
Set clsBS = new CBizSection
	clsBS.FUSE_YN = "Y"
	clsBS.FOnlySub = "Y"
	clsBS.FSale = "N"
	arrBizList = clsBS.fnGetBizSectionList
Set clsBS = nothing


'// ===========================================================================
dim opartner
dim bizsection_cd, selltype, papertype, b2bcharge, paperissuetype
dim shopdiv, shopname

b2bcharge = request("b2bcharge")

set opartner = new CPartnerUser
opartner.FRectDesignerID = shopid

if (shopid <> "") then
	opartner.GetOnePartnerNUser

	if opartner.FResultCount>0 then
		selltype = opartner.FOneItem.Fselltype

		'// ����Ʈ ������
		paperissuetype = "1"
		if Not IsNull(opartner.FOneItem.Ftaxevaltype) then
			if (opartner.FOneItem.Ftaxevaltype = "1") then
				'// ������
				paperissuetype = "2"
			end if
		end if

		if (ctype = "C") then
			'// ��ۺ�����
			b2bcharge = 0
		else
			if (b2bcharge = "" or b2bcharge = "0") and (Not IsNull(opartner.FOneItem.getCommissionPro)) then
				b2bcharge = opartner.FOneItem.getCommissionPro
			end if
		end if

		if (bizsection_cd = "") and (Not IsNull(opartner.FOneItem.FsellBizCd)) then
			bizsection_cd = opartner.FOneItem.FsellBizCd
		end if
	end if

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
		ErrMSG = "���� �귣���������� �⺻ ��������� �����ϼ���.(selltype)"
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
<script type='text/javascript'>

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
        // 20036 => 4010005  //2016/05/17 ERP �ڵ� ����
		if (frm.selltype.value != "4010005") {
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

		if (frm.selltype.value != "4010005") {
			alert("���ó������ �����ΰ�� \n\n���������� ������ �����մϴ�.");
			frm.selltype.focus();
			return;
		}
	<% elseif (shopdiv <> "11") then %>
		if ((frm.papertype.value == "200") || (frm.papertype.value == "102")) {
			alert("���ó������ ���� �Ǵ� �����ΰ�츸 ��� �����մϴ�.");
			frm.papertype.focus();
			return;
		}

		if (frm.selltype.value == "4010005") {
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
	if (frmArr.b2bcharge) {
		if ((frmArr.b2bcharge.value != "") && (frmArr.b2bcharge.value*0 == 0)) {
			frm.b2bcharge.value = frmArr.b2bcharge.value;
		}
	}

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

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<input type="hidden" name="b2bcharge" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left" height="30">
		<input type=radio name=ctype value="M" <% if ctype="M" then response.write "checked" %> onClick="SubmitSearch(frm)"> ��������
		<input type=radio name=ctype value="M_ETC" <% if ctype="M_ETC" then response.write "checked" %> onClick="SubmitSearch(frm)"> ��������(��Ÿ)
		<input type=radio name=ctype value="W" <% if ctype="W" then response.write "checked" %> onClick="SubmitSearch(frm)"> �Ǹź�����(������)
		<input type=radio name=ctype value="A" <% if ctype="A" then response.write "checked" %> onClick="SubmitSearch(frm)"> �Ǹź�����(���� ������)
		<input type=radio name=ctype value="B" <% if ctype="B" then response.write "checked" %> onClick="SubmitSearch(frm)"> �Ǹź�����(�� ������, ��ۺ� ����)
		<input type=radio name=ctype value="C" <% if ctype="C" then response.write "checked" %> onClick="SubmitSearch(frm)"> ��ۺ�����(�� ������)
		<input type=radio name=ctype value="P" <% if ctype="P" then response.write "checked" %> onClick="SubmitSearch(frm)"> <font color="gray">�Ǹź�����(�� ������)</font>
		&nbsp;&nbsp;
		<% if (ctype<>"B" and ctype<>"C" and ctype<>"P") then %>
		<input type="checkbox" name="gbybrand" disabled >�귣�庰
		<% else %>
		<input type="checkbox" name="gbybrand" <%=CHKIIF(gbybrand="on","checked","") %> >�귣�庰
		<% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		<% end if %>
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitSearch(frm)">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="30">
		����ó :
		<%
		if (ctype = "M") or (ctype = "W") or (ctype = "A") then
			'drawSelectBoxByUserDiv "'503','501'", "", "'21'", "", "shopid",shopid
			NewDrawSelectBoxDesignerwithNameAndUserDIV "shopid",shopid, "21"
		elseif (ctype = "M_ETC") then
			'rawSelectBoxByUserDiv "'900'", "", "'21'", "", "shopid",shopid
			NewDrawSelectBoxDesignerwithNameAndUserDIV "shopid",shopid, "21"
		else
			'drawSelectBoxByUserDiv "'999'", "", "'50'", "", "shopid",shopid
			NewDrawSelectBoxDesignerwithNameAndUserDIV "shopid",shopid, "50"
		end if
		''if (ctype <> "B") and (ctype <> "C") then
		''	if (ctype <> "M_ETC") then
		''		drawSelectBoxOffShopNot000 "shopid",shopid
		''		''NewdrawSelectBoxShopAll "shopid", shopid
		''	else
		''		drawSelectBoxEtcMeachul "shopid",shopid
		''	end if
		''else
		''	drawSelectBoxOnIpjumShop "shopid",shopid
		''end if
		%>
		&nbsp;
		����� / �Ǹ��� :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		<input type=checkbox name=onlymifinish <% if onlymifinish="on" then response.write "checked" %> >�̵�� ������
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class=a>
<form name="buffrm">
<tr>
	<td align=right><input type="text" name="totalsuply" value="" size=10 maxlength=10 style="border:1px #999999 solid; text-align=right"></td>
</tr>
</form>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmArr" method=post action="etc_meachul_process.asp">
<input type=hidden name="shopid" value="<%= shopid %>">
<input type=hidden name="workidx" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="10%" bgcolor="<%= adminColor("gray") %>" height="30">����μ�</td>
	<td width="40%" align="left">
        <select class="select" name="bizsection_cd">
        <option value="">--����--</option>
        <% For i = 0 To UBound(arrBizList,2)	%>
    		<option value="<%=arrBizList(0,i)%>" <%IF Cstr(bizsection_cd) = Cstr(arrBizList(0,i)) THEN%> selected <%END IF%>><%=arrBizList(1,i)%></option>
    	<% Next %>
        </select>
	</td>
	<td width="10%" bgcolor="<%= adminColor("gray") %>" height="30">��������</td>
	<td align="left">
		<% drawPartnerCommCodeBox true,"sellacccd","selltype",selltype,"" %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>" height="30">��������</td>
	<td align="left">
        <select class="select" name="papertype">
        	<option value="">--����--</option>
			<option value="100" <%IF (Cstr(papertype) = "100") THEN%> selected <%END IF%>>���� ��꼭</option>
			<option value="101" <%IF (Cstr(papertype) = "101") THEN%> selected <%END IF%>>�鼼 ��꼭</option>
			<option value="102" <%IF (Cstr(papertype) = "102") THEN%> selected <%END IF%>>������ ��꼭</option>
			<option value="200" <%IF (Cstr(papertype) = "200") THEN%> selected <%END IF%>>����Ű�����</option>
        </select>

        <select class="select" name="paperissuetype">
        	<option value="">--����--</option>
			<option value="1" <%IF (Cstr(paperissuetype) = "1") THEN%> selected <%END IF%>>������</option>
			<option value="2" <%IF (Cstr(paperissuetype) = "2") THEN%> selected <%END IF%>>������</option>
        </select>
        *������ = ������ ����
	</td>
	<td width="100" bgcolor="<%= adminColor("gray") %>" height="30">������</td>
	<td align="left">
		<% if ctype="A" or ctype="B" or ctype="C" or ctype="P" then %>
		    <!--
		    <select name="AsignMgTp">
		        <option value="">���ǸŰ� ���
		        <option value="B">���԰� ���
		    </select>
		    -->
			���ǸŰ� ��� <input type="text" class="text" name="b2bcharge" value="<%= b2bcharge %>" size=4 maxlength=5>%
		<% end if %>
	</td>
</tr>
</table>

<br>

<% if (IsRegAvail = False) then %>
	* <font color="red"><%= ErrMSG %></font>
<% end if %>

<p>

<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
<% if (ctype="M") or (ctype="M_ETC") then %>
	<input type=hidden name="mode" value="chulgo">
	<input type=hidden name="ctype" value="<%= ctype %>">
	<input type=hidden name="makerid" value="<%= makerid %>">
	<tr bgcolor="#DDDDFF" align=center>
		<td width=20><input type="checkbox" name="chk" onClick="CheckAll(this)" <% if (IsRegAvail = False) then %>disabled<% end if %> ></td>
		<td width=90>���ó</td>
		<td width=90>���ó��</td>
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
		<td ><input type="checkbox" name="check" <% if not IsNULL(oetcmeachul.FItemList(i).Fprecheckidx) then response.write "disabled" %> value="<%= oetcmeachul.FItemList(i).Fid %>" onClick="AnCheckClick(this); reCalcuSum(frmArr);" <% if (IsRegAvail = False) then %>disabled<% end if %> ></td>
		<td ><%= oetcmeachul.FItemList(i).FSocid %></td>
		<td ><%= oetcmeachul.FItemList(i).Fshopname %></td>
		<td align=center><%= oetcmeachul.FItemList(i).Fexecutedt %>
			<% if oetcmeachul.FItemList(i).Fexecutedt<>oetcmeachul.FItemList(i).FIpgodate then %>
			<br><font color=red>(<%= oetcmeachul.FItemList(i).FIpgodate %>)</font>
			<% end if %>
		</td>
		<td align=center>
			<a href="javascript:PopChulgoSheet('<%= oetcmeachul.FItemList(i).Fid %>')"><%= oetcmeachul.FItemList(i).Fcode %></a><br>
			<a href="javascript:PopIpgoSheet('<%= oetcmeachul.FItemList(i).Fbaljuidx %>')">
				<font color="<% if (oetcmeachul.FItemList(i).FOrderStateCD < "7") then %>red<% else %>gray<% end if %>"><%= oetcmeachul.FItemList(i).Fbaljucode %></font>
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
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align=center><input type=button class=button value="���� ���� ����" onclick="SaveArr(frmArr)"></td>
	</tr>
<% elseif (ctype="W") then %>

	<input type=hidden name="mode" value="witsksell">
	<input type=hidden name="ctype" value="<%= ctype %>">
	<input type=hidden name="makerid" value="<%= makerid %>">
	<tr bgcolor="#DDDDFF" align=center height="30">
		<td width=20><input type="checkbox" name="chk" onClick="CheckAll(this)" <% if (IsRegAvail = False) then %>disabled<% end if %> ></td>
		<td width=90>���óID</td>
		<td width=90>���ó��</td>
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
		<td ><input type="checkbox" name="check" value="<%= oetcmeachul.FItemList(i).Fidx %>" <% if not IsNULL(oetcmeachul.FItemList(i).Fprecheckidx) then response.write "disabled" %> onClick="AnCheckClick(this);" <% if (IsRegAvail = False) then %>disabled<% end if %> ></td>
		<td ><%= oetcmeachul.FItemList(i).Fshopid %></td>
		<td ><%= oetcmeachul.FItemList(i).Fshopname %></td>
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
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="13" align=center><input type=button class=button value="���� ���� ����" onclick="SaveArr(frmArr)"></td>
	</tr>
<% elseif (ctype="A") then %>
	<input type=hidden name="mode" value="offipjumshop">
	<input type=hidden name="ctype" value="<%= ctype %>">
	<input type=hidden name="makerid" value="<%= makerid %>">
	<tr bgcolor="#DDDDFF" align=center height="30">
		<td width=20><input type="checkbox" name="chk" onClick="CheckAll(this)" <% if (IsRegAvail = False) then %>disabled<% end if %> ></td>
		<td width=90>���óID</td>
		<td>���ó��</td>
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
	ttlsuply = ttlsuply + CLng(oetcmeachul.FItemList(i).Ftotsum * (100 - b2bcharge) / 100)
	%>
	<tr bgcolor="#FFFFFF" height="30">
		<td ><input type="checkbox" name="check" value="'<%= oetcmeachul.FItemList(i).Fyyyymmdd %>'" <% if not IsNULL(oetcmeachul.FItemList(i).Fprecheckidx) and (onlymifinish = "") then response.write "disabled" %> onClick="AnCheckClick(this);" <% if (IsRegAvail = False) then %>disabled<% end if %> ></td>
		<td ><%= oetcmeachul.FItemList(i).Fshopid %></td>
		<td ><%= oetcmeachul.FItemList(i).Fshopname %></td>
		<td align=center><%= oetcmeachul.FItemList(i).Fyyyymmdd %></td>

		<td align=center><%= oetcmeachul.FItemList(i).Ftotitemcnt %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotorgsum,0) %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotsum,0) %></td>
		<td align=right>
			<%= FormatNumber(CLng(oetcmeachul.FItemList(i).Ftotsum * (100 - b2bcharge) / 100),0) %>
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
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="12" align=center><input type=button class=button value="���� ���� ����" onclick="SaveArr(frmArr)"></td>
	</tr>
<% elseif (ctype="B") or (ctype="P") then %>
	<input type=hidden name="mode" value="onipjumshop">
	<input type=hidden name="ctype" value="<%= ctype %>">
	<input type=hidden name="makerid" value="<%= makerid %>">
	<tr bgcolor="#DDDDFF" align=center height="30">
		<td width=20><input type="checkbox" name="chk" onClick="CheckAll(this)" <% if (IsRegAvail = False) then %>disabled<% end if %> ></td>
		<td width=90>���óID</td>
		<td>���ó��</td>
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
	<input type="hidden" name="val_<%= oetcmeachul.FItemList(i).Fyyyymmdd %>" value="<%= CLng(oetcmeachul.FItemList(i).Ftotsum * (100 - b2bcharge) / 100)+oetcmeachul.FItemList(i).Ftotdeliversum %>">
	<%
	ttlorgsell = ttlorgsell + oetcmeachul.FItemList(i).Ftotorgsum
	ttlsell = ttlsell + oetcmeachul.FItemList(i).Ftotsum
	ttlbuy = ttlbuy + oetcmeachul.FItemList(i).fbuyprice
	ttlsuply = ttlsuply + CLng(oetcmeachul.FItemList(i).Ftotsum * (100 - b2bcharge) / 100)

	ttlorgsell_dlv  = ttlorgsell_dlv + oetcmeachul.FItemList(i).Ftotdeliverorgsum
    ttlsell_dlv     = ttlsell_dlv + oetcmeachul.FItemList(i).Ftotdeliversum
    ttlsuply_dlv    = ttlsuply_dlv + CLng(oetcmeachul.FItemList(i).Ftotdeliversum * (100 - 0) / 100)
    ttlbuy_dlv      = ttlbuy_dlv + oetcmeachul.FItemList(i).Fbuydeliverprice

	%>
	<tr bgcolor="#FFFFFF" height="30">
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ><input type="checkbox" name="check" value="'<%= oetcmeachul.FItemList(i).Fyyyymmdd %>'" <% if not IsNULL(oetcmeachul.FItemList(i).Fprecheckidx) then response.write "disabled" %> onClick="AnCheckClick(this);" <% if (IsRegAvail = False) then %>disabled<% end if %> ></td>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ><%= oetcmeachul.FItemList(i).Fshopid %></td>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> ><%= oetcmeachul.FItemList(i).Fshopname %></td>
		<% if (gbybrand<>"") then %>
		    <td <%= CHKIIF(ctype="P","rowspan=2","") %> align=center><%= oetcmeachul.FItemList(i).Fmakerid %></td>
		<% end if %>
		<td <%= CHKIIF(ctype="P","rowspan=2","") %> align=center><%= oetcmeachul.FItemList(i).Fyyyymmdd %></td>
		<td align=center><%= oetcmeachul.FItemList(i).Ftotitemcnt %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotorgsum,0) %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotsum,0) %></td>
		<td align=right>
			<%= FormatNumber(CLng(oetcmeachul.FItemList(i).Ftotsum * (100 - b2bcharge) / 100),0) %>
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
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="13" align=center><input type=button class=button value="���� ���� ����" onclick="SaveArr(frmArr)"></td>
	</tr>
<% elseif (ctype="C") then %>
	<input type=hidden name="mode" value="onipjumshopbeasongpay">
	<input type=hidden name="ctype" value="<%= ctype %>">
	<input type=hidden name="makerid" value="<%= makerid %>">
	<tr bgcolor="#DDDDFF" align=center height="30">
		<td width=20><input type="checkbox" name="chk" onClick="CheckAll(this)" <% if (IsRegAvail = False) then %>disabled<% end if %> ></td>
		<td width=90>���óID</td>
		<td>���ó��</td>
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
	<input type="hidden" name="val_<%= oetcmeachul.FItemList(i).Fyyyymmdd %>" value="<%= CLng(oetcmeachul.FItemList(i).Ftotsum * (100 - b2bcharge) / 100)+oetcmeachul.FItemList(i).Ftotdeliversum %>">
	<%
	ttlorgsell = ttlorgsell + oetcmeachul.FItemList(i).Ftotdeliverorgsum
	ttlsell = ttlsell + oetcmeachul.FItemList(i).Ftotdeliversum
	ttlbuy = ttlbuy + oetcmeachul.FItemList(i).Fbuydeliverprice
	ttlsuply = ttlsuply + CLng(oetcmeachul.FItemList(i).Ftotdeliversum * (100 - b2bcharge) / 100)
	%>
	<tr bgcolor="#FFFFFF" height="30">
		<td ><input type="checkbox" name="check" value="'<%= oetcmeachul.FItemList(i).Fyyyymmdd %>'" <% if not IsNULL(oetcmeachul.FItemList(i).Fprecheckidx) then response.write "disabled" %> onClick="AnCheckClick(this);" <% if (IsRegAvail = False) then %>disabled<% end if %> ></td>
		<td ><%= oetcmeachul.FItemList(i).Fshopid %></td>
		<td ><%= oetcmeachul.FItemList(i).Fshopname %></td>
		<td align=center><%= oetcmeachul.FItemList(i).Fyyyymmdd %></td>

		<td align=center><%= oetcmeachul.FItemList(i).Ftotdeliveritemcnt %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotdeliverorgsum,0) %></td>
		<td align=right><%= FormatNumber(oetcmeachul.FItemList(i).Ftotdeliversum,0) %></td>
		<td align=right>
			<%= FormatNumber(CLng(oetcmeachul.FItemList(i).Ftotdeliversum * (100 - b2bcharge) / 100),0) %>
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
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="12" align=center><input type=button class=button value="���� ���� ����" onclick="SaveArr(frmArr)"></td>
	</tr>
<% else %>
<tr><td>1111</td></tr>
<% end if %>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
