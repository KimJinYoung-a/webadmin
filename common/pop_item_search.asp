<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��ǰ�˻�
' History : ���� ������ ��
'			2017.04.10 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<!--

���������� ������� ��ǰ�� �����ϴ� ��� ���������� ��밡���ϵ��� �����Ǿ����ϴ�.
���ѻ��� :

 - ���������� window.open() �Լ��� �̿��� ������ �Ѵ�.

 window.open("/common/pop_item_search.asp");

 - �θ� â�� ������ �Լ��� �����ؾ� �Ѵ�.

function ReActItems(itemgubunarr,
                    itemarr,
                    itemoptionarr,
                    sellcasharr,
                    suplycasharr,
                    buycasharr,
                    itemnoarr,
                    itemnamearr,
                    itemoptionnamearr,
                    designerarr,
                    mwdivarr);

 - ����Ÿ�� ���´� ������ ����.

 "11111|22222|33333|32323|"

-->
<%
dim page, mode, suplyer,shopid, itemgubun, itemid, research
dim nubeasong, nuitem, nuitemoption
dim onoffgubun, idx

shopid = requestCheckVar(request("shopid"),32)
page = requestCheckVar(request("page"),10)
mode = requestCheckVar(request("mode"),32)
suplyer = requestCheckVar(request("suplyer"),32)

itemgubun = requestCheckVar(request("itemgubun"),2)
itemid = requestCheckVar(request("itemid"),10)
research = requestCheckVar(request("research"),2)
nubeasong = requestCheckVar(request("nubeasong"),2)
nuitem = requestCheckVar(request("nuitem"),2)
nuitemoption = requestCheckVar(request("nuitemoption"),2)
onoffgubun = requestCheckVar(request("onoffgubun"),32)

idx = requestCheckVar(request("idx"),10)
if (research<>"on") and (nubeasong="") then
	nubeasong = "on"
end if

if (research<>"on") and (nuitem="") then
	nuitem = "on"
end if

if (research<>"on") and (onoffgubun="") then
	onoffgubun = "online"
end if

if (onoffgubun="online") then
	itemgubun = "10"
end if

if page="" then page=1
if mode="" then mode="bybrand"

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FPageSize = 50
ioffitem.FCurrPage = page
ioffitem.FRectDesigner = suplyer
ioffitem.FRectNoSearchUpcheBeasong = nubeasong
ioffitem.FRectNoSearchNotusingItem = nuitem

ioffitem.FRectItemgubun = itemgubun
ioffitem.FRectNoSearchNotusingItemOption = nuitemoption
ioffitem.FRectItemid = itemid
if onoffgubun="offline" then
	ioffitem.GetOffShopItemList
else
	if (suplyer="") and (itemid="") then

	else
		ioffitem.GetOnLineJumunByBrand
	end if
end if

dim i, shopsuplycash, buycash
%>
<script type='text/javascript'>

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600')
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
	if ((frm.suplyer.value.length<1)){
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


	opener.ReActItems(upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.suplycasharr.value,upfrm.buycasharr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value,upfrm.mwdivarr.value);

}
</script>


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="page" value="<%= page %>">
	<input type="hidden" name="shopid" value="<%= shopid %>" >
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣��:<% drawSelectBoxDesignerwithName "suplyer", suplyer %>
			&nbsp;
			��ǰ�ڵ�ΰ˻� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size=6 maxlength=7>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="radio" name="onoffgubun" value="online" <% if onoffgubun="online" then response.write "checked" %> >ON
			<input type="radio" name="onoffgubun" value="offline" <% if onoffgubun="offline" then response.write "checked" %> >OFF
			&nbsp;
			��ǰ���� :<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
			&nbsp;
			<input type=checkbox name="nubeasong" <% if nubeasong="on" then response.write "checked" %> >��ü��۰˻�����
			<input type=checkbox name="nuitem" <% if nuitem="on" then response.write "checked" %> >����ǰ��
			<input type=checkbox name="nuitemoption" <% if nuitemoption="on" then response.write "checked" %> >���ɼǸ�

		</td>
	</tr>
	</form>
</table>

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="���� ������ �߰�" onclick="AddArr()">
		</td>
		<td align="right">
			
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if ioffitem.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= ioffitem.FTotalCount %></b>
			&nbsp;
			������ : <b><%= Page %> / <%= ioffitem.FTotalPage %></b>
		</td>
	</tr>
	<% end if %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20" align="center">-</td>
		<td width="50">�̹���</td>
		<td width="50">�귣��ID</td>
		<td width="80">��ǰ�ڵ�</td>
		<td width="100">��ǰ��</td>
		<td width="70">�ɼǸ�/<br>���</td>
		<td width="45">�ǸŰ�</td>
		<td width="45">���԰�</td>
		<td width="45">����<br>����</td>
		<td width="45">����</td>
		<td width="50">���</td>
	</tr>
	<% for i=0 to ioffitem.FResultCount -1 %>

	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemList(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemList(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemList(i).Fitemoption %>">
	<input type="hidden" name="itemname" value="<%= ioffitem.FItemList(i).FShopItemName %>">
	<input type="hidden" name="itemoptionname" value="<%= ioffitem.FItemList(i).FShopItemOptionName %>">
	<input type="hidden" name="desingerid" value="<%= ioffitem.FItemList(i).FMakerid %>">
	<input type="hidden" name="sellcash" value="<%= ioffitem.FItemList(i).Fshopitemprice %>">
	<input type="hidden" name="suplycash" value="<%= ioffitem.FItemList(i).Fshopsuplycash %>">
	<input type="hidden" name="buycash" value="<%= ioffitem.FItemList(i).Fshopsuplycash %>">
	<input type="hidden" name="mwdiv" value="<%= ioffitem.FItemList(i).Fmwdiv %>">

	<tr bgcolor="#FFFFFF">
		<td rowspan=2><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td rowspan=2><img src="<%= ioffitem.FItemList(i).FimageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
		<td ><%= ioffitem.FItemList(i).FMakerid %></td>
		<td ><a href="javascript:PopItemSellEdit('<%= ioffitem.FItemList(i).FShopItemID %>');"><%= ioffitem.FItemList(i).GetBarCodeBoldStr %></a></td>
		<td ><%= ioffitem.FItemList(i).FShopItemName %></td>
		<td ><%= ioffitem.FItemList(i).FShopItemOptionName %></td>
		<td rowspan=2 align=right><%= FormatNumber(ioffitem.FItemList(i).Fshopitemprice,0) %></td>
		<td rowspan=2 align=right><%= FormatNumber(ioffitem.FItemList(i).Fshopsuplycash,0) %></td>
		<td rowspan=2 align=center>
		<font color="<%= ioffitem.FItemList(i).getMwDivColor %>"><%= ioffitem.FItemList(i).getMwDivName %></font><br>
		<% if ioffitem.FItemList(i).Fshopitemprice<>0 then %>
		<%= 100-(CLng(ioffitem.FItemList(i).Fshopsuplycash/ioffitem.FItemList(i).Fshopitemprice*10000)/100) %> %
		<% end if %>
		</td>
		<td rowspan=2 ><input type="text" class="text" name="itemno" value="0" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);"></td>
		<td rowspan=2 >

		<% if ioffitem.FItemList(i).Foptusing="N" then %>
		<font color="red">�ɼ�x</font><br>
		<% end if %>
		<% if ioffitem.FItemList(i).IsSoldOut then %>
		<font color="red">�Ǹ�����</font><br>
		<% end if %>
		<% if ioffitem.FItemList(i).Flimityn="Y" then %>
		<font color="blue">����(<%= ioffitem.FItemList(i).getLimitNo %>)</font><br>
		<% end if %>
		<% if ioffitem.FItemList(i).Fpreorderno<>0 then %>
		���ֹ�:<%= ioffitem.FItemList(i).Fpreorderno %>
		<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=3>
			<font color="#444444">
			[<%= Left(ioffitem.FItemList(i).Flastrealdate,10) %>]
			<%= ioffitem.FItemList(i).Flastrealno %>
			+
			<%= ioffitem.FItemList(i).Fipno %>
			<% if ioffitem.FItemList(i).Fchulno<0 then %>
			-
			<% else %>
			+
			<% end if %>
			<%= Abs(ioffitem.FItemList(i).Fchulno) %>
			-
			<%= ioffitem.FItemList(i).Fsellno %>
			-
			<%= ioffitem.FItemList(i).Fipkumdiv4 %>
			-
			<%= ioffitem.FItemList(i).Fipkumdiv2 %>
			</font>
			<br>
			<%= ioffitem.FItemList(i).Fmaxsellday %>��[<%= ioffitem.FItemList(i).Fsell7days %>]
			<% if ioffitem.FItemList(i).Fmaxsellday<>0 then %>
			�����[<%= CLng(ioffitem.FItemList(i).Fsell7days/ioffitem.FItemList(i).Fmaxsellday*10)/10 %>]
			<% else %>
			�����[-]
			<% end if %>
			����[<%= ioffitem.FItemList(i).Frequireno %>]
			����[<%= ioffitem.FItemList(i).Foffjupno+ioffitem.FItemList(i).Foffconfirmno %>]

			����[<%= ioffitem.FItemList(i).Fipkumdiv4 %>]
			����[<%= ioffitem.FItemList(i).Fipkumdiv2 %>]

			<% if ioffitem.FItemList(i).Getshortageno<0 then %>
			��������[<font color="#CC1111"><b><%= ioffitem.FItemList(i).Getshortageno %></b></font>]
			<% else %>
			��������[+<%= ioffitem.FItemList(i).Getshortageno %>]
			<% end if %>

		</td>
		<td align=center>
			<% if ioffitem.FItemList(i).Fcurrno<1 then %>
			<font color="#CC1111"><b><%= ioffitem.FItemList(i).Fcurrno %></b></font>
			<% else %>
			<%= ioffitem.FItemList(i).Fcurrno %>
			<% end if %>
		</td>
	</tr>
	</form>
	<% next %>
</table>



<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if ioffitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ioffitem.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ioffitem.StartScrollPage to ioffitem.FScrollCount + ioffitem.StartScrollPage - 1 %>
				<% if i>ioffitem.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ioffitem.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->




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

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->