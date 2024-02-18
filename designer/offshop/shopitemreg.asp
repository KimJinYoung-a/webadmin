<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%
dim designer, page
designer = session("ssBctID")
page = session("page")

if page="" then page=1

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FPageSize = 1000
ioffitem.FCurrPage = page
ioffitem.FRectDesigner = designer
ioffitem.FRectUpchebeasongInclude = "on"

ioffitem.GetLinkNotRegList3

dim i

dim IsDirectIpchulContractExistsBrand
IsDirectIpchulContractExistsBrand = fnIsDirectIpchulContractExistsBrand(designer)
%>
<script language='javascript'>


function AddArr(){
    <% if Not (IsDirectIpchulContractExistsBrand) then %>
        alert('권한이 없습니다. - 매장 직접 입고 브랜드만 등록 가능합니다.');
        return;
    <% end if %>

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
		alert('선택 아이템이 없습니다.');
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";


	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";

			}
		}
	}
	var ret = confirm('선택 상품을 오프라인 상품으로 등록 하시겠습니까?');

	if (ret){
		upfrm.mode.value = "arradd";
		upfrm.submit();
	}
}
</script>


<!-- 차후에 메뉴설명부분에 넣어야 합니다. -->
<table width="100%" border="0" valign="top" cellpadding="0" cellspacing="0" class="a">
	<tr bgcolor="#FFFFFF">
		<td style="padding:5; border:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">
			* 온라인에서 판매되고 있는 상품 중 오프라인 상품으로 등록되지 않은 상품 리스트 입니다.<br>
			* 등록하시면 [오프상품관리] 메뉴에 상품이 나타나며 바코드 등록 하실 수 있습니다.<br>
			* 오프라인에서 판매되는 상품만 등록하세요.
		</td>
	</tr>
</table>
<!-- 차후에 메뉴설명부분에 넣어야 합니다. -->

<p>

<!--
<table width="100%" cellspacing="1" class="a" >
<tr>
	<td align="right"><a href="javascript:OffItemReg('<%=designer%>')">[오프라인전용 상품등록]</a></td>
</tr>
</table>

<br>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

-->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="선택 상품 오프라인 상품으로 등록" onclick="AddArr()">
			<% if Not (IsDirectIpchulContractExistsBrand) then %>
            (매장 직접 입고 브랜드만 둥록 가능합니다.)
            <% end if %>
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ioffitem.FResultCount %></b>
			&nbsp;
			페이지 : <b><%= Page %> / <%= ioffitem.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
    	<td width="70">브랜드</td>
    	<td width="100">상품코드</td>
    	<td>상품명</td>
    	<td>옵션명</td>
    	<td width="50">온라인<br>계약구분</td>
    	<td width="70">판매가</td>
	</tr>
	<% if ioffitem.FresultCount>0 then %>
	<% for i=0 to ioffitem.FresultCount -1 %>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemlist(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemlist(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemlist(i).Fitemoption %>">
	<input type="hidden" name="makerid" value="<%= ioffitem.FItemlist(i).FMakerID %>">
	<tr align="center" bgcolor="#FFFFFF">
  		<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
  		<td ><%= ioffitem.FItemlist(i).FMakerID %></td>
  		<td><%= ioffitem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(ioffitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,ioffitem.FItemlist(i).Fshopitemid),Format00(6,ioffitem.FItemlist(i).Fshopitemid)) %>-<%= ioffitem.FItemlist(i).Fitemoption %></td>
  		<td align="left"><%= ioffitem.FItemlist(i).FShopItemName %></td>
  		<td><%= ioffitem.FItemlist(i).FShopitemOptionname %></td>
  		<td><font color="<%= ioffitem.FItemlist(i).getMwDivColor %>"><%= ioffitem.FItemlist(i).getMwDivName %></font></td>
  		<td align="right" ><%= FormatNumber(ioffitem.FItemlist(i).FShopItemprice,0) %></td>
  	</tr>
  	</form>
  	<% next %>

  	<% else %>
  	<tr bgcolor="#FFFFFF">
		<td colspan="10" align="center" > [검색결과가 없습니다.] - 등록할 상품이 없습니다. </td>
	</tr>
	<% end if %>
</table>

<br>
<form name="frmArrupdate" method="post" action="shopitem_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemarr" value="">
<input type="hidden" name="itemoptionarr" value="">
</form>
<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->