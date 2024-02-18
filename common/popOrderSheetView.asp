<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ipchulbarcodecls.asp"-->
<!-- #include virtual="/lib/classes/stock/ipchullocationcls.asp"-->
<%

dim i,page,research, masteridx
dim divcd
dim companyid
dim userid
dim defaultlocationid

dim ordersheettype
dim mode
dim titlestring

dim locationidfrom				'매입처
dim locationnamefrom
dim locationidto				'이동처
dim locationnameto

dim executedatestring

dim totalproductno, totalsupplyprice



divcd = requestCheckVar(request("divcd"),32)

'companyid = requestCheckVar(trim(request("companyid")),32)
companyid = requestCheckVar(session("ssBctID"), 32)

masteridx = requestCheckVar(request("masteridx"),32)

ordersheettype = requestCheckVar(request("ordersheettype"),32)
mode = requestCheckVar(request("mode"),32)

if (masteridx = "") then
	masteridx = 0
end if



'==============================================================================
dim ocstoragemaster

set ocstoragemaster = new CStorageMaster

ocstoragemaster.FRectCompanyId = companyid
ocstoragemaster.FRectMasterIdx = masteridx



if (ordersheettype = "offlineorder") then

	ocstoragemaster.GetOneStorageMaster
	titlestring = "오프라인 주문 - " + CStr(ocstoragemaster.FOneItem.Flocationnameto) + "(" + ocstoragemaster.FOneItem.Flocationidto + ")"


else

	ocstoragemaster.GetOneStorageMaster
	titlestring = "오프라인 주문 - " + CStr(ocstoragemaster.FOneItem.Flocationnameto) + "(" + ocstoragemaster.FOneItem.Flocationidto + ")"


end if

executedatestring = "입고일 : " & Left(ocstoragemaster.FOneItem.Ffinishdt, 10)


if C_ADMIN_USER then
elseif (C_IS_SHOP = true) then

	if ((ocstoragemaster.FOneItem.Flocationidfrom <> C_STREETSHOPID) and (ocstoragemaster.FOneItem.Flocationidto <> C_STREETSHOPID)) then
		response.write "<script>alert('잘못된 접근입니다.');</script>"
		response.end
	end if

end if



'==============================================================================
dim ocstoragedetail

set ocstoragedetail = new CStorageDetail

ocstoragedetail.FRectCompanyId = companyid
ocstoragedetail.FRectMasterIdx = masteridx
ocstoragedetail.FRectIsForeignOrder = ocstoragemaster.FOneItem.Fisforeignorder
ocstoragedetail.FRectForeignOrderShopid = ocstoragemaster.FOneItem.Fforeignordershopid


ocstoragedetail.FPageSize = 2000

'상품종류가 300 가지를 넘기면 문제가 생긴다.  //??

if (ordersheettype = "offlineorder") then

	ocstoragedetail.FRectShowSupplyCash = "Y"
	ocstoragedetail.GetStorageDetailList

else

	ocstoragedetail.GetStorageDetailList

end if



'==============================================================================
dim olocationfrom
set olocationfrom = new CLocation
olocationfrom.FRectCompanyId = companyid
olocationfrom.FRectlocationid = ocstoragemaster.FOneItem.Flocationidfrom

olocationfrom.GetOneLocation



'==============================================================================
dim olocationto
set olocationto = new CLocation
olocationto.FRectCompanyId = companyid
olocationto.FRectlocationid = ocstoragemaster.FOneItem.Flocationidto

olocationto.GetOneLocation



'==============================================================================
divcd = ocstoragemaster.FOneItem.Fdivcd

locationidfrom = ocstoragemaster.FOneItem.Flocationidfrom
locationnamefrom = ocstoragemaster.FOneItem.Flocationnamefrom

locationidto = ocstoragemaster.FOneItem.Flocationidto
locationnameto = ocstoragemaster.FOneItem.Flocationnameto


Function AddSpace(byval str)
	if ((str = "") or (IsNull(str))) then
		AddSpace = "&nbsp;"
	else
		AddSpace = str
	end if
End Function

function ConvertCurrencyUnit(str)
	if (str = "USD") then
		ConvertCurrencyUnit = "$"
	else
		ConvertCurrencyUnit = "￦"
	end if
End Function

%>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<script language='javascript'>

// ============================================================================
function SubmitCheckAll() {

	var frm;

	if (document.frmBuyPrc.checkall.checked == true) {
		SubmitSelectAll();
	} else {
		SubmitDeSelectAll();
	}
}

// ============================================================================
function SubmitSelectAll() {

	var frm;

	for (var i = 0;i < document.forms.length; i++) {
		frm = document.forms[i];

		if (frm.name.substr(0, 10)=="frmBuyPrc_") {
			SubmitSelectThis(frm);
		}
	}
}

function SubmitDeSelectAll() {

	var frm;

	for (var i = 0;i < document.forms.length; i++) {
		frm = document.forms[i];

		if (frm.name.substr(0, 10)=="frmBuyPrc_") {
			SubmitDeSelectThis(frm);
		}
	}
}

// ============================================================================
function SubmitCheckThis(frm) {

	if (frm.checkthis.checked == true) {
		SubmitSelectThis(frm);
	} else {
		SubmitDeSelectThis(frm);
	}
}

// ============================================================================
function SubmitSelectThis(frm) {

	frm.checkthis.checked = true;
	hL(frm.checkthis);
}

function SubmitDeSelectThis(frm) {

	document.frmBuyPrc.checkall.checked = false;
	frm.checkthis.checked = false;
	dL(frm.checkthis);
}

</script>






<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr height="20">
		<td align="left" colspan="6">
			<font size="3"><b><%= titlestring %></b></font>
		</td>
		<td align="right" colspan="3">
			<b>주문코드 (<%= ocstoragemaster.FOneItem.Fordercode %>)</b>
		</td>
	</tr>
	<tr height="1">
		<td colspan="9"></td>
	</tr>
</table>

<p>

	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
		<tr valign="top">
			<td width="48%">
        		<!-- 공급자정보 시작 -->
        		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="0" class="table_tl" bgcolor="<%= adminColor("tablebg") %>">
    				<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
        				<td class="td_br" colspan="4"><b>공급자 정보</b></td>
        			</tr>
        			<tr height="25" align="center" bgcolor="#FFFFFF">
        				<td class="td_br">등록번호</td>
        				<td class="td_br" colspan="3"><%= olocationfrom.FOneItem.Fsocno %></td>
        			</tr>
        			<tr height="25" align="center" bgcolor="#FFFFFF">
        				<td class="td_br" width="60">상호</td>
        				<td class="td_br" width="135"><b><%= olocationfrom.FOneItem.Fsocname %></b></td>
        				<td class="td_br" width="60">대표자</td>
        				<td class="td_br" width="90"><%= olocationfrom.FOneItem.Fceoname %></td>
        			</tr>
        			<tr height="25" align="center" bgcolor="#FFFFFF">
        				<td class="td_br">소재지</td>
        				<td class="td_br" colspan="3"><%= olocationfrom.FOneItem.Faddress %>&nbsp;<%= olocationfrom.FOneItem.fmanager_address %></td>
        			</tr>
        			<tr height="25" align="center" bgcolor="#FFFFFF">
        				<td class="td_br">업태</td>
        				<td class="td_br"><%= olocationfrom.FOneItem.Fbisstatus %></td>
        				<td class="td_br">업종</td>
        				<td class="td_br"><%= olocationfrom.FOneItem.Fbistype %></td>
        			</tr>
        			<tr height="25" align="center" bgcolor="#FFFFFF">
        				<td class="td_br">담당자</td>
        				<td class="td_br"><%= olocationfrom.FOneItem.Fdeliver_name %></td>
        				<td class="td_br">연락처</td>
        				<td class="td_br"><%= olocationfrom.FOneItem.Fdeliver_phone %></td>
        			</tr>
        		</table>
        		<!-- 공급자정보 끝 -->
			</td>
			<td bgcolor="#FFFFFF">&nbsp;</td>
			<td width="48%">
        		<!-- 공급받는자정보 시작 -->
        		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="0" class="table_tl" bgcolor="<%= adminColor("tablebg") %>">
    				<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
        				<td class="td_br" colspan="4"><b>공급받는자 정보</b></td>
        			</tr>
        			<tr height="25" align="center" bgcolor="#FFFFFF">
        				<td class="td_br">등록번호</td>
        				<td class="td_br" colspan="3"><%= olocationto.FOneItem.Fsocno %></td>
        			</tr>
        			<tr height="25" align="center" bgcolor="#FFFFFF">
        				<td class="td_br" width="60">상호</td>
        				<td class="td_br" width="135"><b><%= olocationto.FOneItem.Fsocname %></b></td>
        				<td class="td_br" width="60">대표자</td>
        				<td class="td_br" width="90"><%= olocationto.FOneItem.Fceoname %></td>
        			</tr>
        			<tr height="25" align="center" bgcolor="#FFFFFF">
        				<td class="td_br">소재지</td>
        				<td class="td_br" colspan="3"><%= olocationto.FOneItem.Faddress %>&nbsp;<%= olocationto.FOneItem.fmanager_address %></td>
        			</tr>
        			<tr height="25" align="center" bgcolor="#FFFFFF">
        				<td class="td_br">업태</td>
        				<td class="td_br"><%= olocationto.FOneItem.Fbisstatus %></td>
        				<td class="td_br">업종</td>
        				<td class="td_br"><%= olocationto.FOneItem.Fbistype %></td>
        			</tr>
        			<tr height="25" align="center" bgcolor="#FFFFFF">
        				<td class="td_br">담당자</td>
        				<td class="td_br"><%= olocationto.FOneItem.Fmanager_name %></td>
        				<td class="td_br">연락처</td>
        				<td class="td_br"><%= olocationto.FOneItem.Fmanager_hp %></td>
        			</tr>
        		</table>
        		<!-- 공급받는자정보 끝 -->
			</td>
		</tr>
	</table>

	<p>

		<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td class="td_br" colspan="10">
					<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
						<tr>
							<td colspan="6">&nbsp;&nbsp;<strong>상세내역</strong></td>
							<td colspan="3" align="right"><b><%= executedatestring %></b>&nbsp;&nbsp;</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    			<td class="td_br" width="90">물류코드</td>
    			<td class="td_br">매입처</td>
    			<td class="td_br" colspan=2>상품명</td>
    			<td class="td_br">옵션명</td>
    			<td class="td_br" width="60">소비자가</td>
    			<td class="td_br" width="60">공급가</td>
    			<td class="td_br" width="60">주문수량</td>
				<td class="td_br" width="60">확정수량</td>
    			<td class="td_br" width="70">공급가합계</td>
			</tr>
			<%

			totalproductno = 0
			totalsupplyprice = 0

			%>
			<% for i=0 to ocstoragedetail.FresultCount-1 %>
			<%

			totalproductno = totalproductno + ocstoragedetail.FItemList(i).Ffixedno
			totalsupplyprice = totalsupplyprice + (ocstoragedetail.FItemList(i).Fsupplyprice * ocstoragedetail.FItemList(i).Ffixedno)

			%>
			<tr align="center" bgcolor="#FFFFFF">
				<td class="td_br"><%= ocstoragedetail.FItemList(i).Fitemgubun %>-<%= CHKIIF(ocstoragedetail.FItemList(i).Fitemid>=1000000,Format00(8,ocstoragedetail.FItemList(i).Fitemid),Format00(6,ocstoragedetail.FItemList(i).Fitemid)) %>-<%= ocstoragedetail.FItemList(i).Fitemoption %></td>
				<td class="td_br" align="left"><%= ocstoragedetail.FItemList(i).Flocationid %></td>
				<td class="td_br" align="left" colspan=2><%= ocstoragedetail.FItemList(i).Fprdname %></td>
				<td class="td_br"><%= ocstoragedetail.FItemList(i).Fitemoptionname %></td>
				<td class="td_br" align="right"><%= FormatNumber(ocstoragedetail.FItemList(i).Fcustomerprice, 0) %></td>
				<td class="td_br" align="right"><%= FormatNumber(ocstoragedetail.FItemList(i).Fsupplyprice, 0) %></td>
				<td class="td_br">
					<%= ocstoragedetail.FItemList(i).Frequestedno %>
				</td>
				<td class="td_br">
					<%= ocstoragedetail.FItemList(i).Ffixedno %>
				</td>
				<td class="td_br" align="right">
					<%= FormatNumber(ocstoragedetail.FItemList(i).Ffixedno * ocstoragedetail.FItemList(i).Fsupplyprice, 0) %>
				</td>
				<% next %>
				<tr height="35" align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td class="td_br" bgcolor="#FFFFFF">비고</td>
					<td class="td_br" colspan="6" align="left" bgcolor="#FFFFFF"><%= nl2br(ocstoragemaster.FOneItem.Fregistermemo) %></td>
					<td class="td_br"><b>총계</b></td>
					<td class="td_br"><b><%= totalproductno %></b></td>
					<td class="td_br" align="right"><b><%= ForMatNumber(totalsupplyprice,0) %></b></td>
				</tr>
		</table>

		<p>
			<br>
			<p>

				<table width="100%" align="center" cellpadding="3" cellspacing="0" class="table_tl" bgcolor="<%= adminColor("tablebg") %>">
					<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    					<td class="td_br" width="90" height="35">인계자</td>
    					<td class="td_br" colspan=3 align="right" bgcolor="#FFFFFF">(인)</td>
    					<td class="td_br" width="90" height="35">인수자</td>
    					<td class="td_br" colspan=4 align="right" bgcolor="#FFFFFF">(인)</td>
					</tr>
				</table>








				<!-- #include virtual="/lib/db/dbclose.asp" -->
