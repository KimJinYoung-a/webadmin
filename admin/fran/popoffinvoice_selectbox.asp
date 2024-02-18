<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
<%

dim masteridx, i
dim excnoweight, research

research = requestCheckVar(request("research"), 32)
masteridx = requestCheckVar(request("masteridx"), 32)
excnoweight = requestCheckVar(request("excnoweight"), 32)

if (research = "") then excnoweight = "N"

'================================================================================
dim ocartoonboxmaster

set ocartoonboxmaster = new CCartoonBox

ocartoonboxmaster.FRectMasterIdx = masteridx

ocartoonboxmaster.GetMasterOne

'================================================================================
dim oinnerboxlist

set oinnerboxlist = new CCartoonBox

oinnerboxlist.FRectMasterIdx = -1
oinnerboxlist.FRectExcNoWeight = excnoweight
oinnerboxlist.FRectShopid = ocartoonboxmaster.FOneItem.Fshopid

oinnerboxlist.GetInnerBoxList

%>
<script language="javascript">

function CheckBox(frm) {
	if (frm.cartoonboxno.value == "") {
		alert("Cartoon박스번호를 입력하세요.");
		frm.cartoonboxno.focus();
		return false;
	}

	if (frm.innerboxno.value == "") {
		alert("Inner박스번호를 입력하세요.");
		frm.innerboxno.focus();
		return false;
	}

	if (frm.cartoonboxno.value*0 != 0) {
		alert("Cartoon박스번호는 숫자만 가능합니다.");
		frm.cartoonboxno.focus();
		return false;
	}

	if (frm.innerboxno.value*0 != 0) {
		alert("Inner박스번호는 숫자만 가능합니다.");
		frm.innerboxno.focus();
		return false;
	}

	if (frm.cartoonboxweight.value == "") {
		frm.cartoonboxweight.value = 0;
	}

	if (frm.cartoonboxweight.value*0 != 0) {
		alert("Cartoon박스 무게는 숫자만 가능합니다.");
		frm.cartoonboxweight.focus();
		return false;
	}

	if (frm.innerboxweight.value == "") {
		frm.innerboxweight.value = 0;
	}

	if (frm.innerboxweight.value*0 != 0) {
		alert("Inner박스 무게는 숫자만 가능합니다.");
		frm.innerboxweight.focus();
		return false;
	}

	return true;
}

function SaveSelectArr() {
	var upfrm = document.frmadd;
	var frm;
	var pass = false;
	var ret;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,12)=="frmSelectPrc") {
			pass = ((pass) || (frm.cksel.checked == true));
		}
	}

	if (pass != true) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	upfrm.detailidxarr.value = "";
	upfrm.cartoonboxnoarr.value = "";
	upfrm.innerboxnoarr.value = "";
	upfrm.innerboxweightarr.value = "";

	upfrm.baljudatearr.value = "";
	upfrm.shopidarr.value = "";


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,12)=="frmSelectPrc") {
			if (frm.cksel.checked == true) {
				if (CheckBox(frm) != true) {
					return;
				}

				upfrm.detailidxarr.value = upfrm.detailidxarr.value + "|" + frm.detailidx.value;
				upfrm.cartoonboxnoarr.value = upfrm.cartoonboxnoarr.value + "|" + frm.cartoonboxno.value;
				upfrm.innerboxnoarr.value = upfrm.innerboxnoarr.value + "|" + frm.innerboxno.value;
				upfrm.innerboxweightarr.value = upfrm.innerboxweightarr.value + "|" + frm.innerboxweight.value;

				upfrm.baljudatearr.value = upfrm.baljudatearr.value + "|" + frm.baljudate.value;
				upfrm.shopidarr.value = upfrm.shopidarr.value + "|" + frm.shopid.value;
			}
		}
	}

	if (confirm('저장 하시겠습니까?')){
		upfrm.mode.value = "saveselectedbox";
		upfrm.submit();
	}
}

function PopBoxItemList(shopid, yyyy, mm, dd, boxno) {
	var popurl = "/admin/fran/jumunbyboxitemlist.asp?research=on&shopid=" + shopid + "&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=" + dd + "&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + dd + "&boxno=" + boxno;

	var w = window.open(popurl);
	w.focus();
}

function chkAllitem(frmname, frmlength) {
    var frm;
    for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,frmlength)==frmname) {
		    frm.cksel.checked = true;
		    AnCheckClick(frm.cksel);
		}
	}
}

</script>
<form name="frmadd" method=post action="cartoonbox_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="submode" value="popup">
	<input type="hidden" name="shopid" value="<%= ocartoonboxmaster.FOneItem.Fshopid %>">
	<input type="hidden" name="masteridx" value="<%= masteridx %>">
	<input type="hidden" name="detailidxarr" value="">
	<input type="hidden" name="cartoonboxnoarr" value="">
	<input type="hidden" name="cartoonboxweightarr" value="">
	<input type="hidden" name="innerboxnoarr" value="">
	<input type="hidden" name="innerboxweightarr" value="">
	<input type="hidden" name="baljudatearr" value="">
	<input type="hidden" name="shopidarr" value="">
</form>
</table>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="research" value="on">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="right">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="#FFFFFF" height=20>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>미지정박스</strong></font>
                    <input type="checkbox" name="excnoweight" value="N" <%= CHKIIF(excnoweight="N", "checked", "") %> onClick="document.frm.submit();"> 무게 미입력포함
				</td>
				<td align="right">
					총건수:  <%= oinnerboxlist.FResultCount %>
				</td>
			</tr>
		</table>
	</td>
</tr>
</form>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="cksel" onClick="chkAllitem('frmSelectPrc', 12)"></td>
    <td width="110">샵아이디</td>
	<td width="250">샵이름</td>
    <td width="120">발주일</td>
	<td width="80">Inner<br>박스번호</td>
	<td width="80">Inner<br>무게(KG)</td>
    <td width="80">Carton<br>박스번호</td>
	<td width="80">출고일</td>
	<td>비고</td>
</tr>
<% for i=0 to oinnerboxlist.FResultCount-1 %>
<form name="frmSelectPrc_<%= i %>" method="post" action="cartoonbox_process.asp">
<input type="hidden" name="detailidx" value="<%= oinnerboxlist.FItemList(i).Fidx %>">
<input type="hidden" name="baljudate" value="<%= oinnerboxlist.FItemList(i).Fbaljudate %>">
<input type="hidden" name="shopid" value="<%= oinnerboxlist.FItemList(i).Fshopid %>">
<input type="hidden" name="innerboxweight" value="<%= oinnerboxlist.FItemList(i).Finnerboxweight %>">
<input type="hidden" name="cartoonboxno" value="<%= oinnerboxlist.FItemList(i).Fcartoonboxno %>">
<tr align="center" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td><%= oinnerboxlist.FItemList(i).Fshopid %></td>
	<td><%= oinnerboxlist.FItemList(i).Fshopname %></td>
	<td><%= oinnerboxlist.FItemList(i).Fbaljudate %></td>
	<td>
		<%= oinnerboxlist.FItemList(i).Finnerboxno %>
		<input type="hidden" name="innerboxno" value="<%= oinnerboxlist.FItemList(i).Finnerboxno %>">
	</td>
	<td>
		<%
		if (oinnerboxlist.FItemList(i).Finnerboxweight <> "") then
			oinnerboxlist.FItemList(i).Finnerboxweight = FormatNumber(oinnerboxlist.FItemList(i).Finnerboxweight, 2)
		end if
		%>
		<%= oinnerboxlist.FItemList(i).Finnerboxweight %>
	</td>
	<td>
		<%= oinnerboxlist.FItemList(i).Fcartoonboxno %>
	</td>
	<td>
		<%= oinnerboxlist.FItemList(i).Fbeasongdate %>
		<input type="hidden" name="cartoonboxweight" value="0">
	</td>
	<td>
		<input type="button" class="button" value=" 상품보기 " onClick="PopBoxItemList('<%= oinnerboxlist.FItemList(i).Fshopid %>', '<%= Left(oinnerboxlist.FItemList(i).Fbaljudate, 4) %>', '<%= Right(Left(oinnerboxlist.FItemList(i).Fbaljudate, 7), 2) %>', '<%= Right(oinnerboxlist.FItemList(i).Fbaljudate, 2) %>', <%= oinnerboxlist.FItemList(i).Finnerboxno %>)">
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="9" align=center height=30>
		<input type="button" class="button" value=" 선택박스지정 " onclick="SaveSelectArr()">
		<!--
		&nbsp;
		<input type="button" class="button" value=" 선택박스[삭제] " onclick="DeleteSelectArr()">
		-->
	</td>
</tr>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
