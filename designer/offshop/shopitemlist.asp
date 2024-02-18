<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%
dim designer,page,ckonlyoff,ckonlyusing,research
designer    = reQuestCheckVar(session("ssBctID"),32)
page        = reQuestCheckVar(request("page"),10)
ckonlyoff   = reQuestCheckVar(request("ckonlyoff"),10)
ckonlyusing = reQuestCheckVar(request("ckonlyusing"),10)
research    = reQuestCheckVar(request("research"),10)

if page="" then page=1
if research<>"on" then ckonlyusing="Y"

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FPageSize = 50
ioffitem.FCurrPage = page
ioffitem.FRectDesigner = designer
ioffitem.FRectOnlyOffLine = ckonlyoff
ioffitem.FRectOnlyUsing = ckonlyusing

if designer<>"" then
	ioffitem.GetOffShopItemList
end if

dim i

dim IsDirectIpchulContractExistsBrand
IsDirectIpchulContractExistsBrand = fnIsDirectIpchulContractExistsBrand(designer)

%>
<script language='javascript'>
function NextPage(page){
	frm.page.value=page;
	frm.submit();
}
function popOffItemEdit(ibarcode){
    <% if Not (IsDirectIpchulContractExistsBrand) then %>
        alert('권한이 없습니다. - 매장 직접 입고 브랜드만 수정 가능합니다.');
        return;
    <% end if %>
	var popwin = window.open('popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}


function OffItemReg(idesigner){
    <% if Not (IsDirectIpchulContractExistsBrand) then %>
        alert('권한이 없습니다. - 매장 직접 입고 브랜드만 수정 가능합니다.');
        return;
    <% end if %>
	var subwin;
	subwin = window.open('popoffitemreg.asp?designer=' + idesigner,'window_reg','width=500,height=600,scrollbars=yes,resizable=yes');
	subwin.focus();
}

function AnSearch(frm){
	frm.submit();
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function ChargeIdAvail(ichargeid){
	var comp = document.frm.designer;

	if (ichargeid=="10x10"){
		return true
	}

	for (var i=0;i<comp.length;i++){
		if (comp[i].value==ichargeid){
			return true
		}
	}

	return false;
}

function ModiArr(){
    <% if Not (IsDirectIpchulContractExistsBrand) then %>
        alert('권한이 없습니다. - 매장 직접 입고 브랜드만 수정 가능합니다.');
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
	upfrm.isusingarr.value = "";
	upfrm.extbarcodearr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.extbarcodearr.value = upfrm.extbarcodearr.value + frm.extbarcode.value + "|";

				if (frm.isusing[0].checked){
					upfrm.isusingarr.value = upfrm.isusingarr.value + "Y" + "|";
				}else{
					upfrm.isusingarr.value = upfrm.isusingarr.value + "N" + "|";
				}
			}
		}
	}

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		upfrm.mode.value = "arrmodi";
		upfrm.submit();
	}
}

</script>

<!-- 차후에 메뉴설명부분에 넣어야 합니다. -->
<table width="100%" border="0" valign="top" cellpadding="0" cellspacing="0" class="a">
	<tr bgcolor="#FFFFFF">
		<td style="padding:5; border:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">
			* 오프샾 전용상품에 대해 이미지 등록이 필수로 변경되었습니다.<br>
			* 원활한 주문 공급처리를 위해 이미지 없는 상품에 대해 <b>이미지를 등록</b>해 주세요<br>
			* 상품상세정보를 수정하려면 상품번호를 눌러주세요.
		</td>
	</tr>
</table>
<!-- 차후에 메뉴설명부분에 넣어야 합니다. -->

<p>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			사용:<% drawSelectBoxUsingYN "ckonlyusing", ckonlyusing %>
			&nbsp;
			상품구분:
			<select class="select" name="ckonlyoff">
		     	<option value='' selected>전체</option>
		     	<option value='10' <% if ckonlyoff="10" then response.write "selected" %>>온라인등록상품(10)</option>
		     	<option value='90' <% if ckonlyoff="90" then response.write "selected" %>>오프전용상품(90)</option>
	     	</select>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>


<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="오프라인전용 상품등록" onClick="OffItemReg('<%=designer%>')">
			<% if ioffitem.FresultCount>0 then %>
			<input type="button" class="button" value="선택아이템저장" onclick="ModiArr()">
			<% end if %>

			<% if Not (IsDirectIpchulContractExistsBrand) then %>
            (매장 직접 입고 브랜드만 수정 가능합니다.)
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
			검색결과 : <b><%= ioffitem.FTotalCount%></b>
			&nbsp;
			페이지 : <b><%= page %> / <%=  ioffitem.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
    	<td width="50">이미지</td>
    	<td width="100">상품코드</td>
    	<td>상품명</td>
    	<td>옵션명</td>
    	<td width="60">소비자가</td>
    	<td width="60">판매가</td>
    	<td width="110">범용바코드</td>
    	<td width="80">사용여부</td>
	</tr>
	<% for i=0 to ioffitem.FresultCount -1 %>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemlist(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemlist(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemlist(i).Fitemoption %>">
	<input type="hidden" name="tx_charge" value="<%= ioffitem.FItemlist(i).FMakerID %>">
	<tr align="center" bgcolor="#FFFFFF">
  		<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
  		<td ><a href="javascript:popOffItemEdit('<%= ioffitem.FItemlist(i).GetBarCode %>');"><img src="<%= ioffitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" border=0></a></td>
  		<td><a href="javascript:popOffItemEdit('<%= ioffitem.FItemlist(i).GetBarCode %>');"><%= ioffitem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(ioffitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,ioffitem.FItemlist(i).Fshopitemid),Format00(6,ioffitem.FItemlist(i).Fshopitemid)) %>-<%= ioffitem.FItemlist(i).Fitemoption %></a></td>
  		<td align="left"><%= ioffitem.FItemlist(i).FShopItemName %></td>
  		<td><%= ioffitem.FItemlist(i).FShopitemOptionname %></td>
  		<td align="right" ><%= FormatNumber(ioffitem.FItemlist(i).FShopItemOrgprice,0) %></td>
  		<td align="right" ><%= FormatNumber(ioffitem.FItemlist(i).FShopItemprice,0) %></td>
  		<td><input type="text" name="extbarcode" value="<%= ioffitem.FItemlist(i).FextBarcode %>" size="13" maxlength="32" style="border:1px #999999 solid;" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)"></td>
  		<td align="center" >
  		<% if ioffitem.FItemlist(i).Fisusing="Y" then %>
  		<input type="radio" name="isusing" value="Y" checked onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
  		<input type="radio" name="isusing" value="N" onclick="CheckThis(frmBuyPrc_<%= i %>)">N
  		<% else %>
  		<input type="radio" name="isusing" value="Y" onclick="CheckThis(frmBuyPrc_<%= i %>)">Y
  		<input type="radio" name="isusing" value="N" checked onclick="CheckThis(frmBuyPrc_<%= i %>)"><font color="red">N</font>
  		<% end if %>
  		</td>
  	</tr>
  	</form>
  	<% next %>
  	<tr bgcolor="#FFFFFF">
		<td colspan="10" align="center">
		<% if ioffitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ioffitem.StartScrollPage-1 %>');">[pre]</a>
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
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<form name="frmArrupdate" method="post" action="doshopitem.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="isusingarr" value="">
<input type="hidden" name="extbarcodearr" value="">
</form>
<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->