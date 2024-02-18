<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 매장 타 사이트 매칭
' History : 2012.05.15 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<%
dim oother ,i, shopdiv, isusing , menupos ,siteseq
	menupos = requestCheckvar(request("menupos"),10)
	shopdiv = requestCheckvar(request("shopdiv"),32)
	isusing = requestCheckvar(request("isusing"),10)
	siteseq = requestCheckvar(request("siteseq"),10)

if siteseq = "" then siteseq = "99"
if isusing = "" then isusing = "Y"
if shopdiv = "" then shopdiv = "1"
		
set oother = new COffShopChargeUser
    oother.FRectShopDiv2 = shopdiv
    oother.FRectIsUsing = isusing
	oother.FPageSize = 500
	oother.FCurrPage = 1    
	oother.getShoplinkothersitelist
	
function getShoplinkothersite(siteseq)
	if siteseq = "99" then
		Shoplinkothersite = "ITHINKSO[99]"
	else
		Shoplinkothersite = siteseq
	end if
end function

function drawShoplinkothersite(boxname,selectid,chflg)
%>
	<select name="<%= boxname %>" <%= chflg %>>
		<option value="" <% if selectid = "" then response.write " selected" %>>선택</option>
		<option value="99" <% if selectid = "99" then response.write " selected" %>>ITHINKSO[99]</option>	
	</select>
<%
end function
%>

<script language='javascript'>

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

//전체 선택
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

//선택상품 저장
function saveArr(){
	var frmmaster = document.frm;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (frmmaster.siteseq.value<1){
		alert('외부매장구분을 선택하세요');
		frmmaster.siteseq.focus();
		return;
	}
	
	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}
				
	frmarr.shopid.value = "";
	frmarr.othershopid.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (frm.shopid.value<1){
					alert('텐바이텐매장ID 이 지정되지 않았습니다');
					frm.shopid.focus();
					return;
				}

				if (frm.othershopid.value<1){
					alert('외부매장ID를 입력 하세요');
					frm.othershopid.focus();
					return;
				}
								
				frmarr.shopid.value = frmarr.shopid.value + frm.shopid.value + ","
				frmarr.othershopid.value = frmarr.othershopid.value + frm.othershopid.value + ","

			}
		}
	}

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		frmarr.mode.value = 'shopotherreg';
		frmarr.siteseq.value = frmmaster.siteseq.value;
		frmarr.action = 'Shoplinkothersite_process.asp';
		frmarr.submit();
	}
}

function savedel(shopid){
	var frmmaster = document.frm;

	if (frmmaster.siteseq.value<1){
		alert('외부매장구분을 선택하세요');
		frmmaster.siteseq.focus();
		return;
	}

	if (shopid ==''){
		alert('매장이 지정되지 않았습니다');
		return;
	}
	
	var ret = confirm('삭제 하시겠습니까?');

	if (ret){
		frmarr.mode.value = 'shopotherdel';
		frmarr.siteseq.value = frmmaster.siteseq.value;
		frmarr.shopid.value = shopid;
		frmarr.action = 'Shoplinkothersite_process.asp';
		frmarr.submit();
	}
}

function frmsubmit(){
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		외부매장구분 : <% drawShoplinkothersite "siteseq" ,siteseq ," onchange='frmsubmit();'" %>
		매장구분 : <% Call DrawShopDivCombo("shopdiv",shopdiv) %>
		사용여부 : <% Call drawSelectBoxUsingYN("isusing",isusing) %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">

	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" class="button_s" value="선택수정" onClick="saveArr();">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oother.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oother.fresultcount %></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td >
		매장
	</td>
	<td>구분</td>
	<td>국가</td>
	<td>화폐단위<br>해외마진배수</td>
	<td>사용여부</td>
	<td>외부매장ID</td>
	<td>최종수정</td>
	<td>비고</td>
</tr>
<%
for i=0 to oother.FresultCount - 1
%>
<form action="" name="frmBuyPrc<%=i%>" method="get">
<input type="hidden" name="shopid" value="<%=oother.FItemList(i).fshopid%>">
<% if oother.FItemList(i).FIsUsing="N" then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
<% else %>
	<tr align="center" bgcolor="#FFFFFF">
<% end if %>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td>
		<%= oother.FItemList(i).Fshopname %>
		<br><%= oother.FItemList(i).fshopid %>
	</td>
	<td><%= oother.FItemList(i).GetShopdivName %></td>
	<td><%= oother.FItemList(i).FcountryNamekr %></td>
	<td>
		<%= oother.FItemList(i).fcurrencyUnit %>
		<br><%= oother.FItemList(i).fmultipleRate %>
		*<%= oother.FItemList(i).FexchangeRate %>
	</td>
	<td><%= oother.FItemList(i).FIsUsing %></td>
	<td>
		<input type="text" name="othershopid" onKeyup="CheckThis(frmBuyPrc<%=i%>)" size="12" maxlength="13" value="<%=oother.FItemList(i).fothershopid%>" style="text-align:right;">
	</td>
	<td>
		<%= oother.FItemList(i).flastadminuserid %>
		<% if oother.FItemList(i).flastupdate <> "" then %>
			<Br><%= oother.FItemList(i).flastupdate %>
		<% end if %>
	</td>
	<td width=120>
		<% if oother.FItemList(i).fsiteseq <> "" and not isnull(oother.FItemList(i).fsiteseq) then %>
			<% if C_ADMIN_AUTH then %>
				<input type="button" onclick="savedel('<%=oother.FItemList(i).fshopid%>');" value="삭제[관리자]" class="button_s">
			<% end if %>
		<% end if %>
	</td>	
</tr>
</form>
<%
next
else
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=10>검색 결과가 없습니다</td>
</tr>
<%
end if
%>
</table>
<form name="frmarr" method="post" action="">
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="othershopid">
	<input type="hidden" name="siteseq">
	<input type="hidden" name="shopid">
</form>
<%
set oother = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->