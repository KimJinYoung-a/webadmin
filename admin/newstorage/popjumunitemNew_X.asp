<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/stock/newshortagestockcls.asp"-->

<%
const C_STOCK_DAY=7

dim page, mode, makerid, shopid,itemid, research
dim onlynotupchebeasong, onlyusingitem, onlyusingitemoption, onlynotdanjong, soldoutover7days, onlysell, onlynottempdanjong
dim onlynotmddanjong, includepreorder, skiplimitsoldout
dim onoffgubun, idx, shortagetype, onlystockminus
dim changemakerid

shopid = request("shopid")
page = request("page")
mode = request("mode")
itemid = request("itemid")
research = request("research")
onlynotupchebeasong = request("onlynotupchebeasong")
onlyusingitem = request("onlyusingitem")
onlyusingitemoption = request("onlyusingitemoption")
onlynotdanjong = request("onlynotdanjong")
soldoutover7days = request("soldoutover7days")
onoffgubun = request("onoffgubun")
idx = request("idx")
shortagetype = request("shortagetype")
onlysell = request("onlysell")
onlynottempdanjong = request("onlynottempdanjong")
onlynotmddanjong = request("onlynotmddanjong")
includepreorder = request("includepreorder")
skiplimitsoldout = request("skiplimitsoldout")
onlystockminus = request("onlystockminus")


changemakerid = request("changesuplyer")
if (changemakerid = "") then
	changemakerid = request("changemakerid")
end if

makerid = request("makerid")
if (makerid = "") then
	makerid = request("suplyer")
end if



if (research<>"on") and (onlynotupchebeasong = "") then
	onlynotupchebeasong = "on"
end if

if (research<>"on") and (onlyusingitem = "") then
	onlyusingitem = "on"
end if

if (research<>"on") and (onlyusingitemoption="") then
	onlyusingitemoption = "on"
end if

if (research<>"on") and (onlynotdanjong = "") then
	onlynotdanjong = "on"
end if

if (research<>"on") and (onoffgubun="") then
	onoffgubun = "online"
end if

if page="" then page=1
if mode="" then mode="bybrand"

'��ǰ�ڵ� ��ȿ�� �˻�(2008.07.31;������)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim oshortagestock
set oshortagestock  = new CShortageStock
oshortagestock.FPageSize = 50
oshortagestock.FCurrPage = page

oshortagestock.FRectOnlySell			= onlysell
oshortagestock.FRectOnlyUsingItem		= onlyusingitem
oshortagestock.FRectOnlyUsingItemOption	= onlyusingitemoption
oshortagestock.FRectOnlyNotUpcheBeasong	= onlynotupchebeasong

oshortagestock.FRectShortage7days		= chkIIF(shortagetype="7day","on","")
oshortagestock.FRectShortage14days		= chkIIF(shortagetype="14day","on","")
oshortagestock.FRectShortageRealStock	= chkIIF(shortagetype="5under","on","")
oshortagestock.FRectOnlyNotDanjong		= onlynotdanjong
oshortagestock.FRectOnlyNotTempDanjong	= onlynottempdanjong
oshortagestock.FRectOnlyNotMDDanjong	= onlynotmddanjong
oshortagestock.FRectIncludePreOrder		= includepreorder
oshortagestock.FRectSkipLimitSoldOut	= skiplimitsoldout
oshortagestock.FRectOnlyStockMinus		= onlystockminus

oshortagestock.FRectMakerid				= makerid
'oshortagestock.FRectItemGubun			= makerid
oshortagestock.FRectItemId				= itemid
'oshortagestock.FRectItemOption			= makerid

if onoffgubun = "offline" then
	oshortagestock.GetShortageItemListOffline
else
	if (makerid<>"") or (itemid<>"") then
		oshortagestock.GetShortageItemListOnline
	end if
end if

dim i, shopsuplycash, buycash
%>
<script language='javascript'>
function popOffItemEdit(ibarcode){
	var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}


function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'adminitemselledit','width=500,height=600,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function enablebrand(bool){
	//document.frm.designer.disabled = bool;
}

function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

function search(frm){
	/*
	if ((frm.makerid.value.length<1)){
		if ((frm.mode[0].checked)&&(frm.designer.value.length<1)){
			alert('�귣�带 ���� �ϼ���.');
			frm.designer.focus();
			return;
		}
	}
	*/

	frm.submit();
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function AddArr(){
	var upfrm = document.frmArrupdate;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.itemnoarr.value = "";
	upfrm.itemnamearr.value = "";
	upfrm.itemoptionnamearr.value = "";
	upfrm.designerarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (!IsInteger(frm.itemno.value)){
					alert('������ ������ �����մϴ�.');
					frm.itemno.focus();
					return;
				}

				if (frm.itemno.value=="0"){
					alert('������ �Է��ϼ���.');
					frm.itemno.focus();
					return;
				}

				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
				upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
				upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
				upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
				upfrm.itemnamearr.value = upfrm.itemnamearr.value + frm.itemname.value + "|";
				upfrm.itemoptionnamearr.value = upfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
				upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
				upfrm.mwdivarr.value = upfrm.mwdivarr.value + frm.mwdiv.value + "|";

			}
		}
	}


	opener.ReActItems('<%= idx %>', upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.suplycasharr.value,upfrm.buycasharr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value,upfrm.mwdivarr.value);


	//�ʱ�ȭ
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				frm.cksel.checked = false;
				frm.itemno.value="0"


			}
		}
	}

}
</script>


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="page" value="<%= page %>">
	<% if (changemakerid <> "Y") then %>
	<input type="hidden" name="makerid" value="<%= makerid %>" >
	<% else %>
	<input type="hidden" name="changemakerid" value="Y" >
	<% end if %>
	<input type="hidden" name="shopid" value="<%= shopid %>" >
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<% if (changemakerid <> "Y") then %>
			�귣�� : <b><%= makerid %></b>
			<% else %>
			�귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
			<% end if %>
			&nbsp;
			|
			&nbsp;
			���� : <input type="radio" name="onoffgubun" value="online" <% if onoffgubun="online" then response.write "checked" %> >�¶���
			<input type="radio" name="onoffgubun" value="offline" <% if onoffgubun="offline" then response.write "checked" %> >��������
			&nbsp;
			|
			&nbsp;
			<input type=checkbox name="onlyusingitem" <% if onlyusingitem = "on" then response.write "checked" %> >����ǰ��
			<input type=checkbox name="onlyusingitemoption" <% if onlyusingitemoption = "on" then response.write "checked" %> >���ɼǸ�
			<input type=checkbox name="onlysell" <% if onlysell = "on" then response.write "checked" %> >�ǸŻ�ǰ��
			<input type=checkbox name="onlynotupchebeasong" <% if onlynotupchebeasong = "on" then response.write "checked" %> >��ü�������
		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:search(frm);">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
            ��������:
            <input type="radio" name="shortagetype" value="" <% if shortagetype="" then response.write "checked" %> >��ü
            <input type="radio" name="shortagetype" value="7day" <% if shortagetype="7day" then response.write "checked" %> ><%= C_STOCK_DAY %>����������
			<input type="radio" name="shortagetype" value="14day" <% if shortagetype="14day" then response.write "checked" %> ><%= C_STOCK_DAY*2 %>����������
            <input type="radio" name="shortagetype" value="5under" <% if shortagetype="5under" then response.write "checked" %> >�ǻ���� 5����
			&nbsp;
			|
			&nbsp;
			<input type=checkbox name="onlynotdanjong" <% if onlynotdanjong = "on" then response.write "checked" %> >��������
			<input type=checkbox name="onlynottempdanjong" <% if onlynottempdanjong = "on" then response.write "checked" %> >�Ͻ�ǰ������
			<input type=checkbox name="onlynotmddanjong" <% if onlynotmddanjong = "on" then response.write "checked" %> >MD��������


		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size=8 maxlength=7>
			&nbsp;
			|
			&nbsp;
            <input type=checkbox name="includepreorder" <% if includepreorder = "on" then response.write "checked" %> >���ֹ����Ժ�����
            <input type=checkbox name="skiplimitsoldout" <% if skiplimitsoldout = "on" then response.write "checked" %> >�����Ǹ���������
            <input type=checkbox name="onlystockminus" <% if onlystockminus = "on" then response.write "checked" %> >���̳ʽ���
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

	<!-- ��ܹ� ���� -->
	<% if oshortagestock.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						�˻���� : <b><%= oshortagestock.FTotalCount %></b>
						&nbsp;
						������ : <b><%= Page %> / <%= oshortagestock.FTotalPage %></b>
					</td>
					<td align="right">
						<input type="button" class="button" value="��ü����" onClick="AnSelectAllFrame(true)">
        				<input type="button" class="button" value="���� ������ �߰�" onclick="AddArr()">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<% end if %>

	<!-- ��ܹ� �� -->
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
		<td width="50">�̹���</td>
		<td width="80">�귣��ID</td>
		<td width="90">��ǰ�ڵ�</td>
		<td>��ǰ��</td>
		<td>�ɼǸ�</td>
		<td width="45">�ǸŰ�</td>
		<td width="45">���԰�</td>
		<td width="45">����</td>
		<td width="45">����</td>
		<td>���</td>
	</tr>
	<% for i=0 to oshortagestock.FResultCount -1 %>

	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= oshortagestock.FItemList(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= oshortagestock.FItemList(i).Fitemid %>">
	<input type="hidden" name="itemoption" value="<%= oshortagestock.FItemList(i).Fitemoption %>">
	<input type="hidden" name="itemname" value="<%= oshortagestock.FItemList(i).FItemName %>">
	<input type="hidden" name="itemoptionname" value="<%= oshortagestock.FItemList(i).FItemOptionName %>">
	<input type="hidden" name="desingerid" value="<%= oshortagestock.FItemList(i).FMakerid %>">
	<input type="hidden" name="sellcash" value="<%= oshortagestock.FItemList(i).Fsellcash %>">
	<input type="hidden" name="suplycash" value="<%= oshortagestock.FItemList(i).FBuycash %>">
	<input type="hidden" name="buycash" value="<%= oshortagestock.FItemList(i).FBuycash %>">
	<input type="hidden" name="mwdiv" value="<%= oshortagestock.FItemList(i).Fmwdiv %>">

	<% if (oshortagestock.FItemList(i).Foptionusing="N") or (oshortagestock.FItemList(i).Fisusing="N") then %>
	<tr bgcolor="<%= adminColor("gray") %>">
	<% else %>
	<tr bgcolor="#FFFFFF">
	<% end if %>
		<td rowspan=2><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td rowspan=2><img src="<%= oshortagestock.FItemList(i).FimageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
		<td ><%= oshortagestock.FItemList(i).FMakerid %></td>
		<% if oshortagestock.FItemList(i).FItemGubun<>"10" then %>
		<td ><a href="javascript:popOffItemEdit('<%= oshortagestock.FItemList(i).GetBarCode %>')"><%= oshortagestock.FItemList(i).GetBarCodeBoldStr %></a></td>
		<% else %>
		<td ><a href="javascript:PopItemSellEdit('<%= oshortagestock.FItemList(i).FItemID %>');"><%= oshortagestock.FItemList(i).GetBarCodeBoldStr %></a></td>
		<% end if %>
		<td ><a href="/admin/stock/itemcurrentstock.asp?itemid=<%= oshortagestock.FItemList(i).FItemID %>&itemoption=<%= oshortagestock.FItemList(i).FItemOption %>" target=_blank ><%= oshortagestock.FItemList(i).FItemName %></a></td>
		<td ><%= oshortagestock.FItemList(i).FItemOptionName %></td>
		<td rowspan=2 align=right><%= FormatNumber(oshortagestock.FItemList(i).FSellcash,0) %></td>
		<td rowspan=2 align=right><%= FormatNumber(oshortagestock.FItemList(i).FBuycash,0) %></td>
		<td rowspan=2 align=center>
		<font color="<%= oshortagestock.FItemList(i).getMwDivColor %>"><%= oshortagestock.FItemList(i).getMwDivName %></font><br>
		<% if oshortagestock.FItemList(i).FSellcash<>0 then %>
		<%= 100-(CLng(oshortagestock.FItemList(i).FBuycash/oshortagestock.FItemList(i).FSellcash*10000)/100) %> %
		<% end if %>
		</td>
		<td rowspan=2>
			<% if oshortagestock.FItemList(i).Frealstock<0 and oshortagestock.FItemList(i).Fsell7days=0 then %>
			<input type="text" class="text" name="itemno" value="0" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);">
		    <% elseif oshortagestock.FItemList(i).GetNdayShortageNo(14) < 0 then %>
		    <input type="text" class="text" name="itemno" value="<%= (oshortagestock.FItemList(i).GetNdayShortageNo(14))*-1 %>" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);">
		    <% else %>
		    <input type="text" class="text" name="itemno" value="0" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);">
		    <% end if %>
		</td>
		<td rowspan=2 >
			<%= fnColor(oshortagestock.FItemList(i).Fdanjongyn,"dj") %>
			<br>
			<% if oshortagestock.FItemList(i).Foptionusing="N" then %>
			<font color="red">�ɼ�x</font><br>
			<% end if %>
			<% if oshortagestock.FItemList(i).IsSoldOut then %>
			<font color="red">�Ǹ�����</font><br>
			<% end if %>
			<% if oshortagestock.FItemList(i).Flimityn="Y" then %>
			<font color="blue">����(<%= oshortagestock.FItemList(i).getOptionLimitNo %>)</font><br>
			<% end if %>
			<% if oshortagestock.FItemList(i).Fpreorderno<>0 then %>
				���ֹ�:
				<% if oshortagestock.FItemList(i).Fpreorderno<>oshortagestock.FItemList(i).Fpreordernofix then response.write "</br>" + CStr(oshortagestock.FItemList(i).Fpreorderno) + "->" %>
					<%= oshortagestock.FItemList(i).Fpreordernofix %>
			<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=4>
			<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td>�԰�</td>
					<td>�Ǹ�</td>
					<td>���</td>
					<td>��Ÿ</td>
					<td>CS</td>
					<td>�ҷ�</td>
					<td>����</td>

					<td>�ǻ����</td>
					<td bgcolor="<%= adminColor("green") %>">�������</td>
					<td>�������</td>

					<% if oshortagestock.FItemList(i).Fmaxsellday<>7 then %>
					<td bgcolor="<%= adminColor("green") %>">On<font color="#CC1111"><%= oshortagestock.FItemList(i).Fmaxsellday %></font>��</td>
					<td bgcolor="<%= adminColor("green") %>">Off<font color="#CC1111"><%= oshortagestock.FItemList(i).Fmaxsellday %></font>��</td>
					<% else %>
					<td bgcolor="<%= adminColor("green") %>">On<%= oshortagestock.FItemList(i).Fmaxsellday %>��</td>
					<td bgcolor="<%= adminColor("green") %>">Off<%= oshortagestock.FItemList(i).Fmaxsellday %>��</td>
					<% end if %>

					<td><%= C_STOCK_DAY %>��</td>
					<td><%= C_STOCK_DAY*2 %>��</td>
					<!--
					<td>OFF�غ�</td>
					-->
				</tr>
				<tr bgcolor="#FFFFFF" align=center>
					<td><%= oshortagestock.FItemList(i).Ftotipgono %></td>
					<td><%= oshortagestock.FItemList(i).Ftotsellno %></td>
					<td><%= oshortagestock.FItemList(i).Ftotchulgono %></td>
					<td></td>
					<td></td>
					<td><%= oshortagestock.FItemList(i).Ferrbaditemno %></td>
					<td><%= oshortagestock.FItemList(i).Ferrrealcheckno %></td>

					<td>
						<b>
						<% if oshortagestock.FItemList(i).Frealstock<1 then %>
						<font color="#CC1111"><b><%= oshortagestock.FItemList(i).GetCheckStockNo %></b></font>
						<% else %>
						<%= oshortagestock.FItemList(i).Frealstock %>
						<% end if %>
						</b>
					</td>

					<td>
					    <!-- ������� -->
					    <%= oshortagestock.FItemList(i).GetReqNotChulgoNo %></td>
					</td>
					<td>
						<b>
						<% if oshortagestock.FItemList(i).Frealstock + oshortagestock.FItemList(i).GetReqNotChulgoNo < 1 then %>
						<font color="#CC1111"><%= oshortagestock.FItemList(i).Frealstock + oshortagestock.FItemList(i).GetReqNotChulgoNo %></b></font>
						<% else %>
						<%= oshortagestock.FItemList(i).Frealstock + oshortagestock.FItemList(i).GetReqNotChulgoNo %>
						<% end if %>
						</b>
					</td>
					<td><%= oshortagestock.FItemList(i).Fsell7days %></td>
					<td><%= oshortagestock.FItemList(i).Foffchulgo7days %></td>


					<td>
					    <!-- 7�� -->
						<% if oshortagestock.FItemList(i).Fshortageno< 1 then %>
						<font color="#CC1111"><b><%= oshortagestock.FItemList(i).Fshortageno %></b></font>
						<% else %>
						<%= oshortagestock.FItemList(i).Fshortageno %>
						<% end if %>
					</td>
					<td>
					    <!-- N�� �ʿ� -->
						<% if (oshortagestock.FItemList(i).GetNdayShortageNo(14))< 1 then %>
						<font color="#CC1111"><b><%= oshortagestock.FItemList(i).GetNdayShortageNo(14) %></b></font>
						<% else %>
						<%= oshortagestock.FItemList(i).GetNdayShortageNo(14) %>
						<% end if %>
					</td>
					<!--
					<td><%= oshortagestock.FItemList(i).Foffconfirmno %></td>
					-->
				</tr>
			</table>
		</td>
	</tr>
	</form>
	<% next %>

	<!-- �ϴܹ� ���� -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if oshortagestock.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oshortagestock.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + oshortagestock.StartScrollPage to oshortagestock.FScrollCount + oshortagestock.StartScrollPage - 1 %>
				<% if i>oshortagestock.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if oshortagestock.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
	<!-- �ϴܹ� �� -->
</table>



<form name="frmArrupdate" method="post" action="">
<input type="hidden" name="mode" value="arrins">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="itemnamearr" value="">
<input type="hidden" name="itemoptionnamearr" value="">
<input type="hidden" name="designerarr" value="">
<input type="hidden" name="mwdivarr" value="">
</form>
<script language='javascript'>
//alert('������');
</script>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->