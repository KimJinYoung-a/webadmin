<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 개별입출고리스트
' Hieditor : 2009.04.07 서동석 생성
'			 2011.04.20 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False
dim EditEnabled , yyyymmdd,yyyy1,mm1,dd1 ,i ,oipchulmaster, oipchul
dim PriceEditEnable ,idx ,DispReqNo ,edityn
	idx = requestCheckVar(request("idx"),10)

edityn = FALSE

if idx = "" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('idx 값이 없습니다');"
	response.write "	history.back();"
	response.write "</script>"
	dbget.close() : response.end
end if

'if Not (C_IS_SHOP) and Not (C_IS_Maker_Upche) then PriceEditEnable = true

set oipchulmaster = new CShopIpChul
	oipchulmaster.FRectIdx = idx
	oipchulmaster.GetIpChulMasterList

set oipchul = new CShopIpChul
	oipchul.FRectIdx = idx
	oipchul.GetIpChulDetail

if oipchulmaster.ftotalcount < 1 then
	response.write "<script type='text/javascript'>"
	response.write "	alert('해당되는 입출고내역이 없습니다');"
	response.write "	history.back();"
	response.write "</script>"
	dbget.close() : response.end
end if

if C_ADMIN_USER or C_IS_OWN_SHOP then
	edityn = TRUE

'//매장일경우 수정권한은 본인매장만
elseif (C_IS_SHOP) then
	IS_HIDE_BUYCASH = True

	if C_STREETSHOPID = oipchulmaster.FItemList(0).FShopid then
		edityn = TRUE
	else
		edityn = FALSE
	end if
else
	edityn = TRUE
end if

yyyymmdd = Left(CStr(oipchulmaster.FItemList(0).FScheduleDt),10)
yyyy1 = left(yyyymmdd,4)
mm1 = mid(yyyymmdd,6,2)
dd1 = mid(yyyymmdd,9,2)

''입고요청 상태인 경우 확인으로 돌림.
if (C_IS_Maker_Upche) and (oipchulmaster.FItemList(0).IsRequireConfirm) then
    oipchulmaster.FItemList(0).UpcheConfirmProcess
    response.write "<script type='text/javascript'>alert('입고요청 확인되었습니다. 내역을 확인후 발송 처리 해 주세요.');</script>"
end if

''입고대기 상태 && 자기가 등록한내역만 수정 가능
EditEnabled = oipchulmaster.FItemList(0).IsEditEnabled
PriceEditEnable = oipchulmaster.FItemList(0).IsPriceEditEnabled
DispReqNo   = oipchulmaster.FItemList(0).IsDispReqNo

'/본사, 매장일 경우 수정 삭제 가능
if C_ADMIN_USER or C_IS_SHOP then EditEnabled = true

%>

<script type='text/javascript'>

<% if (EditEnabled) then %>
	var ipgowait = true;
<% else %>
	var ipgowait = false;
<% end if %>

function ReActItems(igubun,iitemid,iitemoption,isellcash,isuplycash,ishopbuyprice,iitemno,iitemname,iitemoptionname,iitemdesigner){
	frmArrupdate.itemgubunarr.value = igubun;
	frmArrupdate.itemarr.value = iitemid;
	frmArrupdate.itemoptionarr.value = iitemoption;
	frmArrupdate.sellcasharr.value = isellcash;
	frmArrupdate.suplycasharr.value = isuplycash;
	frmArrupdate.shopbuypricearr.value = ishopbuyprice
	frmArrupdate.itemnoarr.value = iitemno;
	frmArrupdate.designerarr.value = iitemdesigner;
	frmArrupdate.submit();
}

function ReAct(){
	location.reload();
}

function UpcheChulgoProc(frm){
    if (frm.songjangdiv.value.length<1){
		alert('택배사를 선택 하세요');
		frm.songjangdiv.focus();
		return;
	}


	if (frm.songjangno.value.length<1){
		alert('송장 번호를 입력 하세요');
		frm.songjangno.focus();
		return;
	}

    var ret= confirm('발송 처리 하시겠습니까?');
	if (ret){
	    frm.mode.value = "upchechulgoproc";
		frm.submit();
	}
}


function ModiMaster(frm,scd){
	if (!ipgowait){
		alert('입고대기 상태가 아니면 수정할 수 없습니다.');
		return;
	}

	if (frm.chargeid.value.length<1){
		alert('공급처를 선택하세요.');
		return;
	}

	if (frm.shopid.value.length<1){
		alert('샵ID를 선택하세요.');
		return;
	}

	var ret= confirm('수정 하시겠습니까?');
	if (ret){
		if(scd != "")
		{
			frm.statecd.value = scd;
		}
		frm.submit();
	}
}

function AddItems(chargeid, idx){
	if (!ipgowait){
		alert('입고대기 상태가 아니면 수정할 수 없습니다.');
		return;
	}

	var popwin;
	popwin = window.open('popshopitem2.asp?shopid=' + frmMaster.shopid.value + '&chargeid=' + chargeid + '&idx=' + idx,'shopitem','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ModiDetail(frm){
	if (!ipgowait){
		<% if Not(C_ADMIN_AUTH) then %>
			alert('입고대기 상태가 아니면 수정할 수 없습니다.');
			return;
		<% else %>
			alert("관리자 권한!!");
		<% end if %>
	}

	if (!IsDigit(frm.sellcash.value)){
		alert('판매가는 숫자로 입력하세요.');
		frm.sellcash.focus();
		return;
	}

	if (frm.suplycash.value*0 != 0) {
		alert('공급가는 숫자로 입력하세요.');
		frm.suplycash.focus();
		return;
	}

	if (!IsInteger(frm.itemno.value)){
		alert('갯수는 정수로 입력하세요.');
		frm.itemno.focus();
		return;
	}

	var ret = confirm('수정 하시겠습니까?');

	if (ret){
		frm.mode.value="detailmodi";
		frm.submit();
	}
}

<% if (idx <> "") and (edityn = True) or (C_ADMIN_AUTH) then %>
	function ModiDetailArr() {
		var frm;

		var mode = "detailmodiarr";
		var midx = <%= idx %>;

		var didxarr = "";
		var sellcasharr = "";
		var suplycasharr = "";
		var shopbuypricearr = "";
		var itemnoarr = "";

		for (var i = 0;i < document.forms.length; i++) {
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (!IsDigit(frm.sellcash.value)) {
					alert('판매가는 숫자로 입력하세요.');
					frm.sellcash.focus();
					return;
				}

				if (frm.suplycash.value*0 != 0) {
					alert('공급가는 숫자로 입력하세요.');
					frm.suplycash.focus();
					return;
				}

				if (!IsInteger(frm.itemno.value)) {
					alert('갯수는 정수로 입력하세요.');
					frm.itemno.focus();
					return;
				}

				didxarr = didxarr + "|" + frm.idx.value;
				sellcasharr = sellcasharr + "|" + frm.sellcash.value;
				suplycasharr = suplycasharr + "|" + frm.suplycash.value;
				shopbuypricearr = shopbuypricearr + "|" + frm.shopbuyprice.value;
				itemnoarr = itemnoarr + "|" + frm.itemno.value;
			}
		}

		if (confirm('저장 하시겠습니까?')) {
			frm = document.frmArrupdate;

			frm.mode.value = mode;
			frm.midx.value = midx;
			frm.didxarr.value = didxarr;
			frm.sellcasharr.value = sellcasharr;
			frm.suplycasharr.value = suplycasharr;
			frm.shopbuypricearr.value = shopbuypricearr;
			frm.itemnoarr.value = itemnoarr;

			frm.action = "do_shopipchul.asp";

			frm.submit();
		}
	}
<% end if %>

function DelDetail(frm){
	if (!ipgowait){
		alert('입고대기 상태가 아니면 수정할 수 없습니다.');
		return;
	}

	var ret = confirm('삭제 하시겠습니까?');

	if (ret){
		frm.mode.value="detaildel";
		frm.submit();
	}
}
function AddItemsBarCode(frm, digitflag){
	if (frm.shopid.value.length<1){
		alert('가맹점을 먼저 선택하세요');
		frm.shopid.focus();
		return;
	}

	var popwin;
	popwin = window.open('popshopitemBybarcode.asp?shopid=' + frmMaster.shopid.value + '&chargeid=' + frmMaster.chargeid.value + '&digitflag=' + digitflag,'popshopitemBybarcode','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>

<!-- 차후에 메뉴설명부분에 넣어야 합니다. -->
<table width="100%" border="0" valign="top" cellpadding="0" cellspacing="0" class="a">
<tr bgcolor="#FFFFFF">
	<td style="padding:5; border:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>매장 개별 입출고 수정</strong></font><br>
		* 입고 확정후 매일 새벽 1시에 재고에 반영됩니다.<br>
		* 반품시 갯수를 마이너스로 잡아주세요
	</td>
</tr>
</table>
<!-- 차후에 메뉴설명부분에 넣어야 합니다. -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmMaster" method="post" action="do_shopipchul.asp">
<input type="hidden" name="mode" value="modimaster">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="divcode" value="006">
<input type="hidden" name="vatcode" value="008">
<input type="hidden" name="statecd" value="">
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">공급처</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="chargeid" value="<%= oipchulmaster.FItemList(0).FChargeid %>">
		<%= oipchulmaster.FItemList(0).FChargeid %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">매장 </td>
	<td bgcolor="#FFFFFF">
	<input type="hidden" name="shopid" value="<%= oipchulmaster.FItemList(0).FShopid %>">
		<%= oipchulmaster.FItemList(0).FShopid %> (<%= oipchulmaster.FItemList(0).FShopname %>)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">총판매가</td>
	<td bgcolor="#FFFFFF"><%= FormatNumber(oipchulmaster.FItemList(0).FTotalSellCash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">총공급가</td>
	<td bgcolor="#FFFFFF"><%= FormatNumber(oipchulmaster.FItemList(0).FTotalSuplyCash,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">입고예정일</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="scheduledt" value="<%= oipchulmaster.FItemList(0).FScheduleDt %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.scheduledt);">
		<img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>

		&nbsp;
		택배사:<% drawSelectBoxDeliverCompany "songjangdiv", oipchulmaster.FItemList(0).Fsongjangdiv %>
		&nbsp;
		송장번호:<input type="text" class="text" name="songjangno" size=14 maxlength=16 value="<%= oipchulmaster.FItemList(0).Fsongjangno %>" >

		<% IF (C_IS_Maker_Upche) and (oipchulmaster.FItemList(0).FisbaljuExists="Y") and (oipchulmaster.FItemList(0).Fstatecd=-1) then %>
		    <input type="button" class="button" value="발송처리" onClick="UpcheChulgoProc(frmMaster);">
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">입고일</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).FexecDt %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">매장확인일</td>
	<td bgcolor="#FFFFFF">
		<%= oipchulmaster.FItemList(0).Fshopconfirmdate %>
		<% if Not IsNULL(oipchulmaster.FItemList(0).Fshopconfirmuserid) then %>
			(확인ID : <%= oipchulmaster.FItemList(0).Fshopconfirmuserid %>)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">업체확인일</td>
	<td bgcolor="#FFFFFF">
		<%= oipchulmaster.FItemList(0).Fupcheconfirmdate %>
		<% if Not IsNULL(oipchulmaster.FItemList(0).Fupcheconfirmuserid) then %>
			(확인ID : <%= oipchulmaster.FItemList(0).Fupcheconfirmuserid %>)
		<% end if %>
	</td>
</tr>
<% if Not IsNULL(oipchulmaster.FItemList(0).Fbaljuconfirmdate) then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">입고요청확인일</td>
	<td bgcolor="#FFFFFF">
		<%= oipchulmaster.FItemList(0).Fbaljuconfirmdate %>
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">등록일</td>
	<td bgcolor="#FFFFFF">
		<%= oipchulmaster.FItemList(0).FRegDate %>
		<% if Not IsNULL(oipchulmaster.FItemList(0).Freguserid) then %>
			(등록ID : <%= oipchulmaster.FItemList(0).Freguserid %>)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">상태</td>
	<td bgcolor="#FFFFFF"><font color="<%= oipchulmaster.FItemList(0).getStateColor %>"><%= oipchulmaster.FItemList(0).getStateName %></font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">기타요청사항</td>
	<td>
		<textarea name="comment" class="textarea" cols="80" rows="6"><%= oipchulmaster.FItemList(0).fcomment %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" colspan="2" align="center">
	<% if (C_IS_Maker_Upche) and (oipchulmaster.FItemList(0).FStatecd<0) then %>

	<% else %>
		<% if not(edityn) or not(EditEnabled) then %>

	    <% else %>
	    	<input type="button" value="수정" onClick="ModiMaster(frmMaster,'')" class="button">
	    <% end if %>
	<% end if %>
	<%
	'//매장에서 입력한 업체에 발주요청
	if oipchulmaster.FItemList(i).FisbaljuExists="Y" then
	%>
		<% if oipchulmaster.FItemList(0).FStatecd = -5 then %>
			&nbsp;<input type="button" value="입고요청으로변경" onClick="ModiMaster(frmMaster,'-2')" class="button">
		<% end if %>
	<%
	else
	%>
		<% if oipchulmaster.FItemList(0).FStatecd = -5 then %>
			&nbsp;<input type="button" value="입고대기로변경" onClick="ModiMaster(frmMaster,'0')" class="button">
		<% end if %>
	<% end if %>

	</td>
</tr>
</form>
</table>

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="상품추가" onclick="AddItems('<%= oipchulmaster.FItemList(0).FChargeid %>','<%= oipchulmaster.FItemList(0).FIdx %>')" <% if not EditEnabled then response.write "disabled" %>>

		<%' If C_IS_SHOP or C_ADMIN_AUTH or C_OFF_AUTH or C_logics_Part then %>
			<input type="button" class="button" value="발주(바코드)" onclick="AddItemsBarCode(frmMaster,'P')">
			<input type="button" class="button" value="반품(바코드)" onclick="AddItemsBarCode(frmMaster,'M')">
		<%' End If %>
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oipchul.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= oipchul.FTotalCount %></b>
	</td>
</tr>
<% end if %>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">바코드</td>
	<td width="80">브랜드ID</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td width="50">판매가</td>

	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
	    <td width="60">텐바이텐<br>매입가</td>
	    <td width="60">매장<br>공급가</td>
	<% elseif (C_IS_Maker_Upche) then %>
		<td width="60">텐바이텐<br>공급가</td>
	<% else %>
		<td width="60">매장<br>공급가</td>
	<% end if %>

	<td width="50">수량</td>

	<% if (DispReqNo) then %>
		<td width="50">요청<br>수량</td>
	<% end if %>

	<td width="60">판매가합계</td>
	<td width="40">수정</td>
	<td width="40">삭제</td>
</tr>
<% for i=0 to oipchul.FResultCount-1 %>
<form name="frmBuyPrc_<%= i %>" method="post" action="do_shopipchul.asp">
<input type="hidden" name="chargeid" value="<%= oipchulmaster.FItemList(0).FChargeid %>">
<input type="hidden" name="shopid" value="<%= oipchulmaster.FItemList(0).FShopid %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="midx" value="<%= idx %>">
<input type="hidden" name="idx" value="<%= oipchul.FItemList(i).FIdx %>">

<% if Not PriceEditEnable then %>
	<input type="hidden" name="sellcash" value="<%= oipchul.FItemList(i).FSellCash %>">
	<% if IS_HIDE_BUYCASH = True then %>
	<input type="hidden" name="suplycash" value="-1">
	<% else %>
	<input type="hidden" name="suplycash" value="<%= oipchul.FItemList(i).FSuplyCash %>">
	<% end if %>
	<input type="hidden" name="shopbuyprice" value="<%= oipchul.FItemList(i).Fshopbuyprice %>">
<% end if %>

<tr align="center" bgcolor="#FFFFFF">
	<td><%= oipchul.FItemList(i).GetBarCode %></td>
	<td>
		<%= oipchul.FItemList(i).Fdesignerid %>
		<% if (C_ADMIN_AUTH) then %>
		    <% if (oipchul.FItemList(i).Fdesignerid<>oipchul.FItemList(i).FCurrMakerid) then %>
		    <br>(<%=oipchul.FItemList(i).FCurrMakerid%>)
		    <% end if %>
	    <% end if %>
	</td>
	<td align="left"><%= oipchul.FItemList(i).FItemName %></td>
	<td><%= oipchul.FItemList(i).FItemOptionName %></td>

	<% if Not (PriceEditEnable) then %>
		<td align="right"><%= FormatNumber(oipchul.FItemList(i).FSellCash,0) %></td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right"><%= FormatNumber(oipchul.FItemList(i).FSuplyCash,0) %></td><!--텐바이텐 매입가-->
			<td align="right"><%= FormatNumber(oipchul.FItemList(i).Fshopbuyprice,0) %></td><!--매장 공급가-->
		<% elseif (C_IS_Maker_Upche) then %>
			<td align="right"><%= FormatNumber(oipchul.FItemList(i).FSuplyCash,0) %></td><!--텐바이텐 공급가-->
		<% else %>
			<td align="right"><%= FormatNumber(oipchul.FItemList(i).Fshopbuyprice,0) %></td><!--매장 공급가-->
		<% end if %>
	<% else %>
		<td align="right"><input type="text" name="sellcash" value="<%= oipchul.FItemList(i).FSellCash %>" size="7" maxlength="9"></td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right">
				<input type="text" name="suplycash" value="<%= oipchul.FItemList(i).FSuplyCash %>" size="7" maxlength="9"><!--텐바이텐 매입가-->
				<% if oipchul.FItemList(i).FSellCash<>0 then response.write FormatNumber(100-CLNG(oipchul.FItemList(i).FSuplyCash/oipchul.FItemList(i).FSellCash*100*100)/100,2) end if %>
			</td>
			<td align="right">
				<input type="text" name="shopbuyprice" value="<%= oipchul.FItemList(i).Fshopbuyprice %>" size="7" maxlength="9"><!--매장 공급가-->
				<% if oipchul.FItemList(i).FSellCash<>0 then response.write FormatNumber(100-CLNG(oipchul.FItemList(i).Fshopbuyprice/oipchul.FItemList(i).FSellCash*100*100)/100,2) end if %>
			</td>
		<% elseif (C_IS_Maker_Upche) then %>
			<td align="right">
				<input type="text" name="suplycash" value="<%= oipchul.FItemList(i).FSuplyCash %>" size="7" maxlength="9"><!--텐바이텐 공급가-->
				<% if oipchul.FItemList(i).FSellCash<>0 then response.write FormatNumber(100-CLNG(oipchul.FItemList(i).FSuplyCash/oipchul.FItemList(i).FSellCash*100*100)/100,2) end if %>
			</td>
		<% else %>
			<td align="right">
				<input type="text" name="shopbuyprice" value="<%= oipchul.FItemList(i).Fshopbuyprice %>" size="7" maxlength="9"><!--매장 공급가-->
				<% if oipchul.FItemList(i).FSellCash<>0 then response.write FormatNumber(100-CLNG(oipchul.FItemList(i).Fshopbuyprice/oipchul.FItemList(i).FSellCash*100*100)/100,2) end if %>
			</td>
		<% end if %>
	<% end if %>

	<td><input type="text" class="text" name="itemno" value="<%= oipchul.FItemList(i).Fitemno %>" size="3" maxlength="4"></td>

	<% if (DispReqNo) then %>
		<td><%= oipchul.FItemList(i).Freqno %></td>
	<% end if %>

	<td align="right">
		<%= ForMatNumber(oipchul.FItemList(i).Fitemno*oipchul.FItemList(i).FSellCash,0) %>
	</td>
	<td>
		<input type="button" class="button" value="수정" <% if not(edityn) and Not(C_ADMIN_AUTH) then response.write " disabled" %> onclick="ModiDetail(frmBuyPrc_<%= i %>)" <% if not EditEnabled and Not(C_ADMIN_AUTH) then response.write "disabled" %>>
	</td>
	<td>
		<input type="button" class="button" value="삭제" <% if not(edityn) then response.write " disabled" %> onclick="DelDetail(frmBuyPrc_<%= i %>)" <% if not EditEnabled then response.write "disabled" %>>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="12" align=center>
		<% if (idx <> "") and (edityn = True) or (C_ADMIN_AUTH) then %>
			<input type="button" class="button" value=" 전체저장 " onclick="ModiDetailArr()" >
		<% end if %>
	</td>
</tr>
</table>

<form name="frmArrupdate" method="post" action="shopipchulitem_process.asp">
	<input type="hidden" name="mode" value="arrins">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="midx" value="">
	<input type="hidden" name="didxarr" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="shopbuypricearr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="designerarr" value="">
	<input type="hidden" name="chargeid" value="<%= oipchulmaster.FItemList(0).FChargeid %>">
	<input type="hidden" name="shopid" value="<%= oipchulmaster.FItemList(0).FShopid %>">
</form>

<%
set oipchulmaster = Nothing
set oipchul = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
