<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������λ�ǰ ���
' Hieditor : 2009.04.07 ������ ����
'			 2010.06.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim itemgubun,itemid, itemoption, barcode ,i
	barcode	  = requestCheckVar(Trim(request("barcode")),32)

if BF_IsMaybeTenBarcode(barcode) then
    itemgubun 	= BF_GetItemGubun(barcode)
	itemid 		= BF_GetItemId(barcode)
	itemoption 	= BF_GetItemOption(barcode)
end if

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FRectItemgubun = itemgubun
ioffitem.FRectItemId = itemid
ioffitem.FRectItemOption = itemoption
ioffitem.GetOffOneItem

dim IsOnlineItem
	IsOnlineItem = (itemgubun="10")

dim opartner
set opartner = new CPartnerUser
if (ioffitem.FResultCount>0) then
    opartner.FRectDesignerID = ioffitem.FOneItem.Fmakerid
    opartner.GetOnePartnerNUser
end if

dim ooffontract
set ooffontract = new COffContractInfo
if (ioffitem.FResultCount>0) then
    ooffontract.FRectDesignerID = ioffitem.FOneItem.Fmakerid
    ooffontract.GetPartnerOffContractInfo
end if
%>
<script type='text/javascript'>

function AttachImage(comp,filecomp){
	comp.src=filecomp.value;
}

function popOffImageEdit(ibarcode){
	var popwin = window.open('popoffimageedit.asp?barcode=' + ibarcode,'popoffimageedit','width=500,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function EditItem(frm){
<% if (itemgubun<>"00") then %>
	if (frm.cd1.value.length<1){
		alert('ī�װ��� �����ϼ���.');
		return;
	}
<% end if %>
	if (frm.shopitemname.value.length<1){
		alert('��ǰ���� �Է��ϼ���.');
		frm.shopitemname.focus();
		return;
	}

    if (frm.orgsellprice.value.length<1){
		alert('�Һ��ڰ��� �Է��ϼ���.');
		frm.orgsellprice.focus();
		return;
	}

	if (frm.shopitemprice.value.length<1){
		alert('�ǸŰ��� �Է��ϼ���.');
		frm.shopitemprice.focus();
		return;
	}

	if (frm.shopsuplycash.value.length<1){
		alert('���԰��� �Է��ϼ���.');
		frm.shopsuplycash.focus();
		return;
	}

<% if (itemgubun="60") then %>
    if (frm.orgsellprice.value.substr(0,1) != '-'){
		frm.orgsellprice.value = "-"+frm.orgsellprice.value
	}
    if (frm.shopitemprice.value.substr(0,1) != '-'){
		frm.shopitemprice.value = "-"+frm.shopitemprice.value
	}
<% elseif (itemgubun="80") then %>
    if (frm.shopitemprice.value > 0){
		alert("����ǰ�� �ǸŰ��� 0���Ͽ��� �մϴ�.");
		frm.shopitemprice.focus();
		return;
	}
    if (frm.orgsellprice.value > 0){
		alert("����ǰ�� �Һ��ڰ� 0���Ͽ��� �մϴ�.");
		frm.orgsellprice.focus();
		return;
	}
	if (frm.shopitemname.value.match(/^\[����ǰ\] /) == null) {
		alert("����ǰ ������ ������ �� �����ϴ�.");
		return;
	}
<% elseif (itemgubun<>"00") then %>
    if (!IsDigit(frm.shopitemprice.value)){
		alert('�ǸŰ��� ���ڸ� �����մϴ�.');
		frm.shopitemprice.focus();
		return;
	}

    if (!IsDigit(frm.orgsellprice.value)){
		alert('�Һ��ڰ��� ���ڸ� �����մϴ�.');
		frm.orgsellprice.focus();
		return;
	}

<% else %>
	if (!IsInteger(frm.shopitemprice.value)){
		alert('�ǸŰ��� ���ڸ� �����մϴ�.');
		frm.shopitemprice.focus();
		return;
	}

    if (!IsInteger(frm.orgsellprice.value)){
		alert('�Һ��ڰ��� ���ڸ� �����մϴ�.');
		frm.orgsellprice.focus();
		return;
	}
<% end if %>

<% if (itemgubun<>"80") then %>
    if (frm.orgsellprice.value*1<frm.shopitemprice.value*1){
        alert('�Һ��ڰ����� �� �ǸŰ��� Ŭ �� �����ϴ�. �ٽ� �Է��ϼ���.');
		frm.shopitemprice.focus();
		return;
    }
<% end if %>

    if ((!frm.centermwdiv[0].checked)&&(!frm.centermwdiv[1].checked)){
        alert('���� ���� ������ ���� �ϼ���.');
		frm.centermwdiv[0].focus();
		return;
    }

    if ((!frm.vatinclude[0].checked)&&(!frm.vatinclude[1].checked)){
        alert('���� ������ ���� �ϼ���.');
		frm.vatinclude[0].focus();
		return;
    }

<% if Not IsOnlineItem then %>
//	if (frm.ioffimgmain.fileSize<1){
//		alert('�̹����� �Է��� �ּ��� - �ʼ� �����Դϴ�.');
//		frm.file1.focus();
//		return;
//	}
<% end if %>
	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function PopUpcheInfo(v){
	window.open("/admin/lib/popbrandinfoonly.asp?designer=" + v,"popupcheinfo","width=640 height=580 scrollbars=yes resizable=yes");
}

function editOffDesinger(shopid,designerid){
	var popwin = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + designerid,"popshopupcheinfo","width=640 height=540");
	popwin.focus();
}

// ============================================================================
// ī�װ����
function editCategory(cdl,cdm,cdn){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cdn=" + cdn ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.frmedit;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;
}


</script>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		��ǰ�ڵ� : <input type="text" class="text" name="barcode" value="<%= barcode %>" size="20">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
<% if (ioffitem.FResultCount<1) then %>
<tr height="30" bgcolor="FFFFFF">
	<td align="center">[�˻� ����� �����ϴ�.]</td>
</tr>
<% else %>
<tr height="1" bgcolor="FFFFFF">
	<td colspan="15"></td>
</tr>
<form name="frmedit" method=post action="offitemedit_process.asp" >
<input type=hidden name=itemgubun value="<%= itemgubun %>">
<input type=hidden name=itemid value="<%= itemid %>">
<input type=hidden name=itemoption value="<%= itemoption %>">

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">��ǰ�ڵ�</td>
	<td bgcolor="#FFFFFF" colspan="2"><%= ioffitem.FOneItem.GetBarcode %>

	<%if left(ioffitem.FOneItem.GetBarcode,2) = "10" then %>
		�¶��ΰ����ǰ
	<% elseif left(ioffitem.FOneItem.GetBarcode,2) = "90" then %>
		�������������ǰ
	<% elseif left(ioffitem.FOneItem.GetBarcode,2) = "95" then %>
		���������������ǸŻ�ǰ
	<% elseif left(ioffitem.FOneItem.GetBarcode,2) = "80" then %>
		����ǰ
	<% elseif left(ioffitem.FOneItem.GetBarcode,2) = "70" then %>
		�Ҹ�ǰ
	<% end if %>
	<br><font color="#AAAAAA">(90������������, 80����ǰ, 70�Ҹ�ǰ, 95���������������Ǹ�)</font>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>��ǰ��</td>
	<td bgcolor="#FFFFFF" colspan="2">
	<input type="text" class="text" name="shopitemname" value="<%= ioffitem.FOneItem.Fshopitemname %>" size="40" maxlength="40">
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>�ɼǸ�</td>
	<% if (IsOnlineItem) and (ioffitem.FOneItem.Fitemoption<>"0000") then %>
	<td bgcolor="#FFFFFF" colspan="2">
		<input type="text" class="text" name="shopitemoptionname" value="<%= ioffitem.FOneItem.Fshopitemoptionname %>" size="20" maxlength="20" class="input_01" >
	</td>
	<% else %>
		<td bgcolor="#FFFFFF" colspan="2">
			<input type="text" name="shopitemoptionname" value="<%= ioffitem.FOneItem.Fshopitemoptionname %>" size="40" maxlength="40" class="input_01">
		</td>
	<% end if %>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td>ī�װ�</td>
	<td bgcolor="#FFFFFF" colspan="2">
	  <input type="hidden" name="cd1" value="<%= ioffitem.FOneItem.FCateCDL %>">
	  <input type="hidden" name="cd2" value="<%= ioffitem.FOneItem.FCateCDM %>">
	  <input type="hidden" name="cd3" value="<%= ioffitem.FOneItem.FCateCDS %>">

      <input type="text" class="text" name="cd1_name" value="<%= ioffitem.FOneItem.FCateCDLName %>" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" class="text" name="cd2_name" value="<%= ioffitem.FOneItem.FCateCDMName %>" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" class="text" name="cd3_name" value="<%= ioffitem.FOneItem.FCateCDSName %>" size="12" readonly style="background-color:#E6E6E6">

      <input type="button" class="button" value="����" onclick="editCategory(frmedit.cd1.value,frmedit.cd2.value,frmedit.cd3.value);">
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>���ݼ���</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td bgcolor="#FFFFFF" >�Һ��ڰ�</td>
				<td bgcolor="#FFFFFF" >�� �ǸŰ�</td>
				<% if not(C_IS_SHOP) then %>
					<td bgcolor="#FFFFFF" >���԰�</td>
				<% end if %>
				<td bgcolor="#FFFFFF" >���ް�</td>
			</tr>
			<tr bgcolor="#DDDDFF" align="center">
			    <td bgcolor="#FFFFFF"><input type=text name="orgsellprice" value="<%= ioffitem.FOneItem.FShopItemOrgprice %>" size=8 maxlength=9 class="input_right" ></td>
				<td bgcolor="#FFFFFF"><input type=text name="shopitemprice" value="<%= ioffitem.FOneItem.Fshopitemprice %>" size=8 maxlength=9 class="input_right" ></td>
				<% if not(C_IS_SHOP) then %>
					<td bgcolor="#FFFFFF"><input type=text name="shopsuplycash" value="<%= ioffitem.FOneItem.Fshopsuplycash %>" size=8 maxlength=9 class="input_right" ></td>
				<% end if %>
				<td bgcolor="#FFFFFF" ><input type=text name="shopbuyprice" value="<%= ioffitem.FOneItem.Fshopbuyprice %>" size=8 maxlength=9 class="input_right" ></td>
			</tr>
			<tr bgcolor="#DDDDFF" align="center">
				<td bgcolor="#FFFFFF" colspan="2"></td>
				<td bgcolor="#FFFFFF" colspan="2">(0 �ΰ�� �⺻���� �ڵ� ����)</td>
			</tr>
			<tr bgcolor="#DDDDFF" align="center">
			    <td bgcolor="#FFFFFF" colspan="4">
			        <% if (ioffitem.FOneItem.FItemGubun="10") then %>
			            <b>�¶��� �Ǹ� ��ǰ�� ��� ���� ������ �¶��� �ǸŰ���<br>�����ϰ� �����˴ϴ�.</b>
			        <% end if %>
			    </td>
			</tr>
		</table>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>�������</td>
	<td bgcolor="#FFFFFF">
		<% if ioffitem.FOneItem.Fisusing="Y" then %>
		<input type=radio name=isusing value="Y" checked >�����
		<input type=radio name=isusing value="N">������
		<% else %>
		<input type=radio name=isusing value="Y"  >�����
		<input type=radio name=isusing value="N" checked >������
		<% end if %>
	</td>
	<td rowspan="4" bgcolor="#FFFFFF" align="center">
		<% if IsOnlineItem then %>
		<img src="<%= ioffitem.FOneItem.FimageList %>" width="100" height="100">
		<% else %>
		<a href="javascript:popOffImageEdit('<%= ioffitem.FOneItem.GetBarcode %>');"><img src="<%= ioffitem.FOneItem.FOffImgList %>" width="100" height="100" border="0"></a>
        <br>
        <a href="javascript:popOffImageEdit('<%= ioffitem.FOneItem.GetBarcode %>');">[�̹�������]</a>
		<% end if %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>���͸��Ա���</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="centermwdiv" value="W" <%= ChkIIF(ioffitem.FOneItem.FCenterMwDiv="W","checked","") %> >��Ź
		<input type="radio" name="centermwdiv" value="M" <%= ChkIIF(ioffitem.FOneItem.FCenterMwDiv="M","checked","") %> >����
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>��������</td>
	<td bgcolor="#FFFFFF">
	<input type="radio" name="vatinclude" value="Y" <%= ChkIIF(ioffitem.FOneItem.Fvatinclude="Y","checked","") %>  >����
	<input type="radio" name="vatinclude" value="N" <%= ChkIIF(ioffitem.FOneItem.Fvatinclude="N","checked","") %> > <font color="<%= ChkIIF(ioffitem.FOneItem.Fvatinclude="N","blue","#000000") %>">�鼼</font>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>������ڵ�</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="extbarcode" value="<%= ioffitem.FOneItem.Fextbarcode %>" size="20" maxlength="20" class="input_01" >
	</td>
</tr>

<tr height="1" bgcolor="FFFFFF">
	<td colspan="15"></td>
</tr>

<tr bgcolor="<%= adminColor("pink") %>">
	<td width=100>�귣��������</td>
	<td bgcolor="#FFFFFF" colspan="2"><a href="javascript:PopUpcheInfo('<%= ioffitem.FOneItem.Fmakerid %>');"><%= ioffitem.FOneItem.Fmakerid %></a> (<%= opartner.FOneItem.Fsocname_kor %>,<%= opartner.FOneItem.FCompany_name %>)
	<br><font color="#AAAAAA">(�귣�� ����� �����ڿ��� ����)</font>
	</td>
</tr>
<tr bgcolor="<%= adminColor("pink") %>">
	<td>�¶���</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<%= FormatNumber(ioffitem.FOneItem.FOnlineOrgprice,0) %> / <%= FormatNumber(ioffitem.FOneItem.FOnlineBuycash,0) %>
		&nbsp;&nbsp;
		<font color="<%= mwdivColor(ioffitem.FOneItem.FmwDiv) %>"><%= mwdivName(ioffitem.FOneItem.FmwDiv) %></font>
		&nbsp;
		<% if ioffitem.FOneItem.FOnlineSellcash<>0 then %>
		<%= CLng((1- ioffitem.FOneItem.FOnlineBuycash/ioffitem.FOneItem.FOnlineOrgprice)*100) %> %
		<% end if %>

		<% if (ioffitem.FOneItem.FOnlineSailYn="Y") then %>
		<br>
		<font color="red">
		<%= FormatNumber(ioffitem.FOneItem.FOnlineSellcash,0) %> / <%= FormatNumber(ioffitem.FOneItem.FOnlineBuycash,0) %>
		&nbsp;&nbsp;
			<% if (ioffitem.FOneItem.FOnlineOrgprice<>0) then %>
		        <%= CLng((ioffitem.FOneItem.FOnlineOrgprice - ioffitem.FOneItem.FOnlineSellcash)/ioffitem.FOneItem.FOnlineOrgprice*100) %>%
		    <% end if %>
		    ����
		</font>
		&nbsp;&nbsp;
		<font color="<%= mwdivColor(ioffitem.FOneItem.FmwDiv) %>"><%= mwdivName(ioffitem.FOneItem.FmwDiv) %></font>
		&nbsp;
			<% if ioffitem.FOneItem.FOnlineSellcash<>0 then %>
				<%= CLng((1- ioffitem.FOneItem.FOnlineBuycash/ioffitem.FOneItem.FOnlineSellcash)*100) %> %
			<% end if %>

		<% end if %>

	</td>
</tr>


<tr bgcolor="<%= adminColor("pink") %>">
	<td width=100>��������<br>[������]</td>
	<td bgcolor="#FFFFFF" colspan=5>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=80%>
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop000','<%= ioffitem.FOneItem.Fmakerid %>')"><b>��������ǥ</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop000") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop000") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="1") then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= ioffitem.FOneItem.Fmakerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td width=60><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td width=60><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
</tr>
<tr bgcolor="<%= adminColor("pink") %>">
	<td width=100>��������<br>[������]</td>
	<td bgcolor="#FFFFFF" colspan=5>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=80%>
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop800','<%= ioffitem.FOneItem.Fmakerid %>')"><b>����������ǥ</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop800") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop800") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="3")  then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= ioffitem.FOneItem.Fmakerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td ><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>

		</table>
	</td>
</tr>
<tr bgcolor="<%= adminColor("pink") %>">
	<td width=100>��������<br>[�ؿܰ���]</td>
	<td bgcolor="#FFFFFF" colspan=5>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=80%>
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop870','<%= ioffitem.FOneItem.Fmakerid %>')"><b>�ؿܰ��޴�ǥ</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop870") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop870") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="5")  then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= ioffitem.FOneItem.Fmakerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td ><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
</tr>

<tr height="1" bgcolor="FFFFFF">
	<td colspan="15"></td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>�����</td>
	<td bgcolor="#FFFFFF" colspan=5><%= ioffitem.FOneItem.Fregdate %></td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>����������</td>
	<td bgcolor="#FFFFFF" colspan=5><%= ioffitem.FOneItem.Fupdt %></td>
</tr>

</form>
<tr bgcolor="#FFFFFF">
	<td colspan="3" align=center>
		<% if not(C_IS_SHOP) then %>
			<input type="button" class="button" value=" ���� " onclick="EditItem(frmedit)">
		<% end if %>
	</td>
</tr>
<% end if %>
</table>

<%
set opartner = Nothing
set ooffontract = Nothing
set ioffitem = Nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->