<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��ü�����ֹ����ۼ�
' History : �̻� ����
'			2018.05.21 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/summaryupdatelib.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%
''��ü�����ֹ����ۼ�
dim designerid, statecd, itemgubunarr, itemidarr, itemoptionarr, ordernoarr, orgbaljucode
dim includepreorderno, shortyn, orderno, realstock, research, centermwdiv, shopid
dim yyyy1,mm1,dd1,nowdate, oupchejumun, iidx, DefaultItemMwDiv
	designerid  		= requestcheckvar(request("designerid"),32)
	centermwdiv  		= requestcheckvar(request("centermwdiv"),1)
	itemgubunarr  		= request("itemgubunarr")
	itemidarr  			= request("itemidarr")
	itemoptionarr  		= request("itemoptionarr")
	ordernoarr  		= request("ordernoarr")
	orgbaljucode  		= request("orgbaljucode")
	statecd     		= requestcheckvar(request("statecd"),1)
	shortyn   			= requestcheckvar(request("shortyn"),10)
	includepreorderno   = requestcheckvar(request("includepreorderno"),10)
	research   			= requestcheckvar(request("research"),2)
	shopid   			= requestcheckvar(request("shopid"),32)
	yyyy1 = requestCheckVar(request("yyyy1"), 4)
	mm1 = requestCheckVar(request("mm1"), 2)
	dd1 = requestCheckVar(request("dd1"), 2)

if (research = "") then
	'shortyn = "Y"
	''includepreorderno = "Y"
	shortyn = "Y"
end if

if (includepreorderno = "Y") then
	shortyn = "Y"
end if

if yyyy1="" then
	nowdate = CStr(Now)
	nowdate = DateSerial(Left(nowdate,4), CLng(Mid(nowdate,6,2))-1,Mid(nowdate,9,2))
	yyyy1 = Left(nowdate,4)
	mm1 = Mid(nowdate,6,2)
	dd1 = Mid(nowdate,9,2)
end if

if (designerid<>"") then
	DefaultItemMwDiv = GetDefaultItemMwdivByBrand(designerid)
end if

iidx =0
if (itemgubunarr<>"") and (designerid<>"") then
	set oupchejumun = new COrderSheet
		oupchejumun.FRectMakerid = designerid
		oupchejumun.FRectTargetid = designerid
		oupchejumun.FRectBaljuId = "10x10"
		oupchejumun.FRectBaljuname = "�ٹ�����"
		oupchejumun.FRectReguser = session("ssBctId")
		oupchejumun.FRectRegname = session("ssBctCname")
		oupchejumun.FRectITemGubunArr = itemgubunarr
		oupchejumun.FRectItemIdArr = itemidarr
		oupchejumun.FRectItemOptionArr = itemoptionarr
		oupchejumun.FRectOrderNoArr = ordernoarr
		oupchejumun.FRectOrgBaljuCode = orgbaljucode
		oupchejumun.FRectScheduledate = Left(now(), 10)

		if (centermwdiv="M") then
			oupchejumun.FRectdivcode = "101"
		else
			oupchejumun.FRectdivcode = "111"
		end if

		iidx = oupchejumun.MakeUpcheJumunNew

		'�ֹ��� ���� ���ֹ� ������Ʈ
		PreOrderUpdateBySheetIdx(iidx)
	set oupchejumun = Nothing

	response.write "<script>alert('�ۼ��Ǿ����ϴ�.');</script>"
	response.write "<script>location.href = 'upchejumuninputNew.asp?menupos=" & menupos & "&designerid=" & designerid & "&statecd=" & statecd & "&shortyn=" & shortyn & "&includepreorderno=" & includepreorderno & "';</script>"
	dbget.close : response.end
end if

'// �����ڵ����
dim oordersheet1
set oordersheet1 = new COrderSheet
	oordersheet1.FRectMakerid = designerid
	oordersheet1.FRectBaljuId = shopid
	oordersheet1.FRectTargetid = "10x10"
	oordersheet1.FRectStatecd = statecd
	oordersheet1.FRectShortYN = shortyn
	oordersheet1.FRectIncludePreOrderNo = includepreorderno
	oordersheet1.FRectStartDate = yyyy1 + "-" + mm1 + "-" + dd1
	oordersheet1.FGroupByBaljuCode = "N"
	oordersheet1.GetFranBalju2UpcheBaljuBrandlistNewProcNEW

'//�����ڵ�����
dim oordersheet2
set oordersheet2 = new COrderSheet
	oordersheet2.FRectMakerid = designerid
	oordersheet2.FRectBaljuId = shopid
	oordersheet2.FRectTargetid = "10x10"
	oordersheet2.FRectStatecd = statecd
	oordersheet2.FRectShortYN = shortyn
	oordersheet2.FRectIncludePreOrderNo = includepreorderno
	oordersheet2.FRectStartDate = yyyy1 + "-" + mm1 + "-" + dd1
	oordersheet2.FGroupByBaljuCode = "Y"

dim MultiBaljuCodeExist : MultiBaljuCodeExist = False
for i=0 to oordersheet1.FResultCount - 1
	if oordersheet1.FItemList(i).Fbaljucodecnt > 1 then
		MultiBaljuCodeExist = True
		exit for
	end if
next

if (MultiBaljuCodeExist) AND (designerid<>"") then  ''(designerid<>"") �߰� 2017/01/12
	oordersheet2.GetFranBalju2UpcheBaljuBrandlistNewProcNEW
	''rw oordersheet2.FResultCount
end if
rw oordersheet1.FResultCount

dim i, j, tmpcolor
dim priceErrItemID : priceErrItemID = 0

%>

<script type='text/javascript'>

function SearchMakerid(makerid) {
	var frm = document.searchfrm;

	frm.designerid.value = makerid;
	frm.submit();
}


function MakeJumunByMakerid(designerid){
	//alert(idxarr);
	//alert(designerid);
	document.dumifrm.idxarr.value=idxarr;
	document.dumifrm.designerid.value=designerid;
	document.dumifrm.etcstr.value=etcstr;
	document.dumifrm.submit();
}

function PopFranBalju2Upchebalju(frm){
	var designerid,baljuid,popwin;
	designerid = frm.designerid.value;
	baljuid = frm.baljuid.value;
	popwin = window.open('popfranbalju2upchebalju.asp?designerid=' + designerid + '&baljuid=' + baljuid  ,'franbalju2upchebalju','width=800,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function PopFranBalju2UpchebaljuByID(designerid){
    var baljuid,popwin;
	baljuid = "10x10";
	popwin = window.open('popfranbalju2upchebalju.asp?designerid=' + designerid + '&baljuid=' + baljuid  ,'franbalju2upchebalju','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function MakeJumun(designerid){
	var frm;
	var pass = false;
	var orgbaljucode;

	if (designerid == "") {
		alert("���� �귣�带 �˻��ϼ���.");
		return;
	}

	/*
	if ("<%= shopid %>" == "") {
		alert("���� ������ �˻��ϼ���.");
		return;
	}
	*/

	if (priceErrItemID !== 0) {
		if (confirm("�ݾ׿���!!\n\n���Ը��� ������ ��ǰ����(��ǰ�ڵ� : " + priceErrItemID + ")\n\n��� �����Ͻðڽ��ϱ�?") != true) {
			return;
		}
	}

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� ��ǰ�� �����ϴ�.');
		return;
	}

	var itemgubunarr, itemidarr, itemoptionarr, ordernoarr;

	itemgubunarr = "";
	itemidarr = "";
	itemoptionarr = "";
	ordernoarr = "";
	orgbaljucode = "";

	for (var i = 0; i < document.forms.length; i++) {
		frm = document.forms[i];

		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if (orgbaljucode == "") {
					orgbaljucode = "'" + frm.baljucode.value + "'";
				} else {
					orgbaljucode = orgbaljucode + ",'" + frm.baljucode.value + "'";
				}
				itemgubunarr = itemgubunarr + frm.itemgubun.value + ",";
				itemidarr = itemidarr + frm.itemid.value + ",";
				itemoptionarr = itemoptionarr + frm.itemoption.value + ",";
				ordernoarr = ordernoarr + frm.orderno.value + ",";

				if ((frm.orderno.value == "") || (frm.orderno.value*0 != 0)) {
					alert("�ֹ������� ���ڷ� �Է��ؾ� �մϴ�.");
					frm.orderno.focus();
					return;
				}
			}
		}
	}

	var DefaultItemMwDiv = "<%= DefaultItemMwDiv %>";

    //alert(document.dumifrm.centermwdiv.value);
	if (confirm('�ֹ����� �ۼ��Ͻðڽ��ϱ�?')) {
		if (frm.centermwdiv.value == "") {
			frm.centermwdiv.value = DefaultItemMwDiv;
		}
		document.dumifrm.designerid.value=designerid;
		document.dumifrm.orgbaljucode.value=orgbaljucode;

		document.dumifrm.itemgubunarr.value=itemgubunarr;
		document.dumifrm.itemidarr.value=itemidarr;
		document.dumifrm.itemoptionarr.value=itemoptionarr;
		document.dumifrm.ordernoarr.value=ordernoarr;

		document.dumifrm.submit();
	}
}

function EnableDisableAll(chk, centermwdiv) {
	var frm;
	var isselect = chk.checked;
	var pass;
	var checkeditemcount = 0;

	if (searchfrm.designerid.value == "") {
		alert("���� �귣�带 �˻��ϼ���.");
		chk.checked = false;
		return;
	}

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked) {
				checkeditemcount = checkeditemcount + 1;
			}
		}
	}

	if ((isselect == true) && (dumifrm.centermwdiv.value == "") && (centermwdiv != "")) {
		// ù��° ����/��Ź ������ ��ǰ ���ý�
		if (confirm("���� ����/��Ź ��ǰ�� �ϰ����� �Ͻðڽ��ϱ�?") == true) {
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.centermwdiv.value == centermwdiv) {
						frm.cksel.checked = true;
					}
				}
			}
		}
		dumifrm.centermwdiv.value = centermwdiv;
	} else if ((isselect == true) && (dumifrm.centermwdiv.value != "") && (centermwdiv != dumifrm.centermwdiv.value) && (centermwdiv != "")) {
		// ù��° �̿� ����/��Ź ������ ��ǰ ���ý�
		alert("���Ի�ǰ�� ��Ź��ǰ�� ���ÿ� �ֹ��� �� �����ϴ�.\n\n�и��ؼ� �ֹ����� �ۼ��ϼ���.");
		chk.checked = false;
	} else if (checkeditemcount == 0) {
		dumifrm.centermwdiv.value = "";
	}

	for (var i = 0; i < document.forms.length; i++) {
		frm = document.forms[i];

		if (frm.name.substr(0,9)=="frmBuyPrc") {
			AnCheckClick(frm.cksel);
		}
	}
}

function jsPopCurrentItemStock(itemgubun, itemid, itemoption) {
	var popwin;
	popwin = window.open('/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption, 'jsPopCurrentItemStock','width=1200,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function poporderlist(designerid,shopid,yyyy1,mm1,dd1){
	var popwin = window.open('/admin/fran/jumunlist.asp?designer='+designerid+'&shopid='+shopid+'&yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1,'addreg','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!--
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#FFFFFF">
	<td>
		* 2�� 1�� �ֹ� ���� <br>
		���Ծ�ü��� ������ �¶��� ����� ��������.<br>
		��Ź��ü�� ���԰�->�¶������ ������ ���� ���<br>
		��Ź��ü�� ��Ź��->�¶������ ������ ���� ���<br>
		<br>
		�̰����� ���� �ֹ��ؾ� �ϴ°��<br>
		- �����ε� ������ �ٸ����(���� ���� prixe, multiple_choice, nanishow)<br>
		- ��ü����ֹ���.(�������� ��������, �������� ������Ź)
	</td>
</tr>
</table>
-->

<form name="searchfrm" method="get">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="statecd" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
	    �귣�� : <% drawSelectBoxDesignerwithName "designerid", designerid %>
	    &nbsp;
	    �� : 
		<% 'drawSelectBoxOffShop "shopid",shopid %>
		<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
	    <!--
	    &nbsp;
		<input type=radio name="statecd" value="" <% if statecd="" then response.write "checked" %> >�ֹ����� + ��ǰ�غ�
		<input type=radio name="statecd" value="0" <% if statecd="0" then response.write "checked" %> >�ֹ�����
		<input type=radio name="statecd" value="1" <% if statecd="1" then response.write "checked" %> >��ǰ�غ�
		-->
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.searchfrm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!--

		-->
		<b>�ۼ��� : <% DrawOneDateBox yyyy1,mm1,dd1 %> ~ ����</b>
		&nbsp;
		<input type=checkbox name="shortyn" value="Y" <% if shortyn = "Y" then response.write "checked" %>> ��������
	</td>
</tr>
</table>
</form>

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="���û�ǰ���� �ֹ��� �ۼ�(<%= designerid %>, <%= shopid %>)" onclick="MakeJumun('<%= designerid %>');">
	</td>
	<td align="right">
	    * �귣�� ���̵� Ŭ���� �ۼ� ����
		/ ��ü ��ǰ �ֹ����� �̰����� �ۼ� �Ͻ� �� �����ϴ�.
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=20></td>
	<td width="120">�귣��ID</td>
	<td width="50">�̹���</td>
	<td width="120">
		��ǰ�ڵ�<br>
		<font color="#FF0000">�ٹ�</font>/<font color="#000000">����</font>/<font color="#0000FF">��������</font>
	</td>
	<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
	<td width="50">����<br>���Ա���</td>
	<td width="50">�ǻ�<br>���</td>
	<td width="50">����<br />�ֹ�</td>
	<td width="50"><b>����<br />����</b></td>
	<td width="50">����<br>����</td>
	<td width="80"><b>�ֹ�<br>����</b></td>
	<td width="200">���</td>
</tr>
<% for i=0 to oordersheet1.FResultCount -1 %>
	<%

	if (Not oordersheet1.FItemList(i).IsOnLineItem) then
		tmpcolor = "#0000FF"
	else
		if (oordersheet1.FItemList(i).IsUpchebeasong = True) then
			tmpcolor = "#000000"
		else
			tmpcolor = "#FF0000"
		end if
	end if

	'// �����ڵ� ����� ���Ѵ�. 2016-05-24, skyer9
	if oordersheet2.FResultCount > 0 then
		oordersheet1.FItemList(i).Fbaljucode = ""
		for j = 0 to oordersheet2.FResultCount - 1
			if (oordersheet1.FItemList(i).FItemGubun = oordersheet2.FItemList(j).FItemGubun) and (oordersheet1.FItemList(i).FItemId = oordersheet2.FItemList(j).FItemId) and (oordersheet1.FItemList(i).FItemoption = oordersheet2.FItemList(j).FItemoption) then
				if oordersheet1.FItemList(i).Fbaljucode = "" then
					oordersheet1.FItemList(i).Fbaljucode = oordersheet2.FItemList(j).Fbaljucode
				else
					oordersheet1.FItemList(i).Fbaljucode = oordersheet1.FItemList(i).Fbaljucode + "','" + oordersheet2.FItemList(j).Fbaljucode
				end if
			end if
		next
	end if

	if (oordersheet1.FItemList(i).FpriceCnt > 1) then
		priceErrItemID = oordersheet1.FItemList(i).FItemId
	end if

	%>
<form name="frmBuyPrc" method=get action="">
<input type="hidden" name="baljucode" value="<%= oordersheet1.FItemList(i).Fbaljucode %>">
<input type="hidden" name="itemgubun" value="<%= oordersheet1.FItemList(i).FItemGubun %>">
<input type="hidden" name="itemid" value="<%= oordersheet1.FItemList(i).FItemId %>">
<input type="hidden" name="itemoption" value="<%= oordersheet1.FItemList(i).FItemoption %>">
<input type="hidden" name="centermwdiv" value="<%= oordersheet1.FItemList(i).Fcentermwdiv %>">
<tr align="center" bgcolor="#FFFFFF">
	<td><input type=checkbox name="cksel" onClick="EnableDisableAll(this, '<%= oordersheet1.FItemList(i).Fcentermwdiv %>');"></td>
	<input type=hidden name="idx" value="">
	<td>
		<a href="javascript:SearchMakerid('<%= oordersheet1.FItemList(i).FMakerid %>');"><%= oordersheet1.FItemList(i).FMakerid %></a><br>
	</td>
	<td height=55>
		<% if (oordersheet1.FItemList(i).GetImageSmall <> "") then %>
		<img src="<%= oordersheet1.FItemList(i).GetImageSmall %>" width=50 height=50>
		<% end if %>
	</td>
	<td>
		<a href="javascript:jsPopCurrentItemStock('<%= oordersheet1.FItemList(i).FItemGubun %>', <%= oordersheet1.FItemList(i).FItemId %>, '<%= oordersheet1.FItemList(i).FItemoption %>')">
			<font color="<%= tmpcolor %>">
				<%= oordersheet1.FItemList(i).FItemGubun %>-<%= CHKIIF(oordersheet1.FItemList(i).FItemId>=1000000,Format00(8,oordersheet1.FItemList(i).FItemId),Format00(6,oordersheet1.FItemList(i).FItemId)) %>-<%= oordersheet1.FItemList(i).FItemoption %>
			</font>
		</a>
	</td>
	<td align="left">
		<%= oordersheet1.FItemList(i).FItemName %>
			&nbsp;
		<% if oordersheet1.FItemList(i).FItemoption<>"0000" then %>
			<br>[<font color="blue"><%= oordersheet1.FItemList(i).FItemOptionname %></font>]
		<% end if %>
	</td>
	<td><%= oordersheet1.FItemList(i).GetCenterMWDivString %></td>
	<td>
		<% if (oordersheet1.FItemList(i).Frealstock <> 0) then %><font color=red><b><% end if %>
		<%= oordersheet1.FItemList(i).Frealstock %>
	</td>

	<td><%= oordersheet1.FItemList(i).Ftotbaljuitemno %></td>
	<td><b><%= oordersheet1.FItemList(i).FJupsuCount %></b></td>

	<%
	if (oordersheet1.FItemList(i).Frealstock - oordersheet1.FItemList(i).FpreUnderCnt - oordersheet1.FItemList(i).FJupsuCount + oordersheet1.FItemList(i).Fonbaljuitemno) < 0 then
		orderno = (oordersheet1.FItemList(i).Frealstock - oordersheet1.FItemList(i).FpreUnderCnt - oordersheet1.FItemList(i).FJupsuCount + oordersheet1.FItemList(i).Fonbaljuitemno) * -1
	else
		orderno = 0
	end if
	%>
	<td>
		<%= -1*orderno %>
	</td>
	<td>
		<input type="text" name="orderno" value="<%= orderno %>" size="4">
	</td>
	<td>
		<%= oordersheet1.FItemList(i).Fshopname %>
		<% if (Not IsNull(oordersheet1.FItemList(i).Fupcheorderlinkcode)) then %>
			<!--
			<br><%= oordersheet1.FItemList(i).Fupcheorderlinkcode %><br>
			-->
		<% end if %>
		<% if ((Not IsNull(oordersheet1.FItemList(i).FreipgoMayDate)) and (Left(oordersheet1.FItemList(i).FreipgoMayDate, 10) >= Left(DateAdd("m", -3, now()), 10) ) ) then %>
			<br><%= Left(oordersheet1.FItemList(i).FreipgoMayDate, 10) %><br>
		<% end if %>
		<% if oordersheet1.FItemList(i).Fpreorderno<>0 then %>
			<br>���ֹ�: <%=  oordersheet1.FItemList(i).Fpreorderno %>
			<% if oordersheet1.FItemList(i).Fpreorderno<>oordersheet1.FItemList(i).Fpreordernofix then %>
				-> <%= oordersheet1.FItemList(i).Fpreordernofix %>
			<% end if %>
		<% end if %>
		<% if (oordersheet1.FItemList(i).FpriceCnt > 1) then %>
			<br><font color="red"><b>�ݾ׿���</b></font>
		<% end if %>
		<br><input type="button" value="�ֹ���������" onclick="poporderlist('<%= oordersheet1.FItemList(i).FMakerid %>','<%=shopid%>','<%=yyyy1%>','<%=mm1%>','<%=dd1%>');" class="button">
	</td>
</tr>
</form>
<% next %>
</table>

<form name="dumifrm" method=post action="">
<input type="hidden" name="designerid" value="">
<input type="hidden" name="centermwdiv" value="">
<input type="hidden" name="orgbaljucode" value=""><!-- max(m.baljucode) -->
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="ordernoarr" value="">
</form>

<script type='text/javascript'>
var priceErrItemID = <%= priceErrItemID %>;
</script>

<%
set oordersheet1 = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
