<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 상품 등록
' History : 2009.04.07 서동석 생성
'			2010.12.13 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<%
dim designer,page ,cdl, cdm, cds ,itemid , i ,IsDirectIpchulContractExistsBrand, itemname
dim research, mwdiv, sellyn, usingyn, itemlinktype, itemoption, prdcode, contractyn
	designer    = RequestCheckVar(request("designer"),32)
	page        = RequestCheckVar(request("page"),9)
	research    = RequestCheckVar(request("research"),9)
	mwdiv       = RequestCheckVar(request("mwdiv"),9)
	usingyn     = RequestCheckVar(request("usingyn"),9)
	sellyn      = RequestCheckVar(request("sellyn"),9)
	cdl         = RequestCheckVar(request("cdl"),3)
	cdm         = RequestCheckVar(request("cdm"),3)
	cds         = RequestCheckVar(request("cds"),3)
	itemid      = RequestCheckVar(request("itemid"),1500)
	itemname    = RequestCheckVar(request("itemname"),32)
	itemoption = RequestCheckVar(request("itemoption"),4)
	prdcode 		= requestCheckVar(request("prdcode"),32)
    contractyn 		= requestCheckVar(request("contractyn"),32)

if page="" then page=1

'/업체일 경우 아이디 박아넣음
if (C_IS_Maker_Upche) then
	designer = session("ssBctId")
	IsDirectIpchulContractExistsBrand = fnIsDirectIpchulContractExistsBrand(designer)
	usingyn ="Y" ''업체인경우 사용여부 무조건 Y
else
    if (research="") and (mwdiv="") then mwdiv="MW"  ''기본값. MW (업체가 아닌경우)
    if (research="") and (usingyn="") then usingyn="Y"  ''기본값. Y
end if

if (research="") and (sellyn="") then sellyn="Y"  ''기본값. Y
if (research="") and (contractyn="") then contractyn="Y"

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 200
	ioffitem.FCurrPage = page
	ioffitem.FRectDesigner = designer
	ioffitem.FRectitemid = itemid
    ioffitem.FRectOnlineMWdiv = mwdiv
    ioffitem.FRectIsusing = usingyn
    ioffitem.FRectSellYN  = sellyn
    ioffitem.FRectitemname  = itemname
    ioffitem.FRectcdl  = cdl
    ioffitem.FRectcdm  = cdm
    ioffitem.FRectcds  = cds
    ioffitem.frectitemoption = itemoption
	ioffitem.FRectPrdCode = prdcode
    ioffitem.FRectContractYN = contractyn
	ioffitem.GetLinkReqList()

'if itemlinktype	= "" then itemlinktype = "O"
%>
<script type="text/javascript">

function frmsubmit(){
	if (frm.itemname.value!=''){
		if (frm.designer.value==''){
			alert('상품명 검색시 브랜드를 반드시 넣어 주세요.(부하문제)');
			return;
		}
	}

	frm.submit();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}


function AddArr(){
    <% if C_IS_Maker_Upche and Not(IsDirectIpchulContractExistsBrand) then %>
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

				if (frm.itemlinktype[0].checked == true){
					upfrm.itemlinktypearr.value = upfrm.itemlinktypearr.value + frm.itemlinktype[0].value + "|";
				} else if (frm.itemlinktype[1].checked == true){
					upfrm.itemlinktypearr.value = upfrm.itemlinktypearr.value + frm.itemlinktype[1].value + "|";
				}
			}
		}
	}
	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		upfrm.mode.value = "arradd";
		upfrm.submit();
	}
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function gotoPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 상품코드 : <textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;
		* 옵션코드 : <input type="text" class="text" name="itemoption" value="<%= itemoption %>" size="4" maxlength="4">
		&nbsp;
		* 물류코드 :
		<input type="text" class="text" name="prdcode" value="<%= prdcode %>" size="16" maxlength="32" onKeyPress="if (event.keyCode == 13) frmsubmit();">
		&nbsp;
		* 브랜드 :
		<% if (C_IS_Maker_Upche) then %>
			<%= designer %>
			<input type="hidden" name="designer" value="<%= designer %>">
		<% else %>
			<% drawSelectBoxDesignerwithName "designer",designer  %>
		<% end if %>

		&nbsp;
		* 상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	    <span style="white-space:nowrap;">* ON 매입구분:<% drawSelectBoxMWU "mwdiv", mwdiv %></span>
	    &nbsp;
	    <span style="white-space:nowrap;">* ON 판매여부:<% drawSelectBoxSellYN "sellyn", sellyn %></span>
	    &nbsp;
        <span style="white-space:nowrap;">* 계약여부:<% drawSelectBoxUsingYN "contractyn", contractyn %></span>
	    &nbsp;
	    <% if (C_IS_Maker_Upche) then %>
	        <span style="white-space:nowrap;">* ON 사용여부:<%= usingyn %></span>
	    <% else %>
	        <span style="white-space:nowrap;">* ON 사용여부:<% drawSelectBoxUsingYN "usingyn", usingyn %></span>
        <% end if %>

	    <% if (FALSE) then %>
		<input type="radio" name="umwdiv" value="ALL" <% if umwdiv="ALL" then response.write "checked" %> <% if designer = "" then response.write " disabled" %>>(판매중)모든상품
		<input type="radio" name="umwdiv" value="Y" <% if umwdiv="Y" then response.write "checked" %> <% if designer = "" then response.write " disabled" %>>(판매중)업체배송상품
		<input type="radio" name="umwdiv" value="N" <% if umwdiv="N" then response.write "checked" %> <% if designer = "" then response.write " disabled" %>>판매중지 상품 검색
	    <% end if %>
		<br><!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<% if (designer<>"") and (mwdiv="U") and (ioffitem.FresultCount>0) then %>
	<script language='javascript'>alert('업체 배송 상품의 경우 판매 할 상품만 등록 하시기 바랍니다.');</script>
<% end if %>

<!-- 액션 시작 -->
<!--※ 온라인에서 판매되고 있는 상품 중 오프라인 상품으로 등록되지 않은 상품 리스트 입니다.<br>
※ 등록하시면 [오프상품관리] 메뉴에 상품이 나타나며 바코드 등록 하실 수 있습니다.<br>
※ 오프라인에서 판매되는 상품만 등록하세요.<br>
※ 기본값 소비자가 등록.-->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		* <font color="red">구매제한상품</font>, Present상품, 티켓상품, 여행상품 등은 표시되지 않습니다.
	</td>
	<td align="right">
		<% if ioffitem.FresultCount>0 then %>
		<input type="button" class="button" value="선택 상품 오프라인 상품으로 등록" onclick="AddArr()">
		<% end if %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ioffitem.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= Page %> / <%= ioffitem.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	<td width="100">상품코드</td>
	<td>브랜드ID</td>
	<td>상품명<font color="blue">[옵션명]</font></td>
	<td width="50">ON<br>매입구분</td>
	<!-- <td width="50">Center<br>매입구분</td> -->
	<td width="90">소비자가</td>
	<td width="90">판매가</td>
	<td width="250">등록구분</td>
</tr>

<% if ioffitem.FresultCount > 0 then %>
	<%
	for i=0 to ioffitem.FresultCount -1

	''할인중이면서 기간할인이 아니면 판매가로 등록(상시할인)
	if ioffitem.FItemlist(i).Ftermsale ="N" and ioffitem.FItemlist(i).FOnlineitemorgprice>ioffitem.FItemlist(i).FShopItemprice then
		itemlinktype = "S"
	else
		itemlinktype = "O"
	end if
	%>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemlist(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemlist(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemlist(i).Fitemoption %>">
	<input type="hidden" name="makerid" value="<%= ioffitem.FItemlist(i).FMakerID %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td><%= ioffitem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(ioffitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,ioffitem.FItemlist(i).Fshopitemid),Format00(6,ioffitem.FItemlist(i).Fshopitemid)) %>-<%= ioffitem.FItemlist(i).Fitemoption %></td>
		<td>
			<%= ioffitem.FItemlist(i).FMakerID %>
		</td>
		<td align="left">
			<%= ioffitem.FItemlist(i).FShopItemName %>
			<% if ioffitem.FItemlist(i).Fitemoption<>"0000" then %>
				<font color="blue">[<%= ioffitem.FItemlist(i).FShopitemOptionname %>]</font>
			<% end if %>
		</td>
		<td><font color="<%= ioffitem.FItemlist(i).getMwDivColor %>"><%= ioffitem.FItemlist(i).getMwDivName %></font></td>
		<!-- <td></td> -->
		<td align="right"><%= FormatNumber(ioffitem.FItemlist(i).FOnlineitemorgprice,0) %></td>
		<td align="right">
			<% if ioffitem.FItemlist(i).Ftermsale ="Y" then %>
				<font color="red">기간할인</font>
		    <% elseif (ioffitem.FItemlist(i).FOnlineitemorgprice>ioffitem.FItemlist(i).FShopItemprice) then %>
		    	<font color="red"><!--상시-->할인</font>
		    <% end if %>
		    <%= FormatNumber(ioffitem.FItemlist(i).FShopItemprice,0) %>
		</td>
		<td>
			<input type="radio" name="itemlinktype" value="S" <% if itemlinktype = "S" then response.write " checked" %> onclick="CheckThis(frmBuyPrc_<%= i %>)">판매가등록
			<input type="radio" name="itemlinktype" value="O" <% if itemlinktype = "O" then response.write " checked" %> onclick="CheckThis(frmBuyPrc_<%= i %>)">소비자가등록
		</td>
	</tr>
	</form>
	<% next %>

    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="8">
	   	<% if ioffitem.HasPreScroll then %>
			<span class="list_link"><a href="javascript:gotoPage(<%= ioffitem.StartScrollPage-1 %>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ioffitem.StartScrollPage to ioffitem.StartScrollPage + ioffitem.FScrollCount - 1 %>
			<% if (i > ioffitem.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ioffitem.FCurrPage) then %>
			<span class="page_link"><font color="red"><b>[<%= i %>]</b></font></span>
			<% else %>
			<a href="javascript:gotoPage(<%= i %>)" class="list_link"><font color="#000000">[<%= i %>]</font></a>
			<% end if %>
		<% next %>
		<% if ioffitem.HasNextScroll then %>
			<span class="list_link"><a href="javascript:gotoPage(<%= i %>)">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
      	</td>
    </tr>
<% else %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan=20>검색결과가 없습니다</td>
	</tr>
<% end if %>

<form name="frmArrupdate" method="post" action="shopitem_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="itemlinktypearr" value="">
</form>
</table>

<!--
<div class="a">- 상품정보 변경 상품(온라인과 오프라인 상품정보가 맞지 않는경우)</div>
-->

<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
