<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%

'사용않함

1

'// [OFF]오프_가맹관리>>가맹점정산관리(매출) 에서 발행요청 하면 나오는 페이지

	'// 변수 선언 //
	dim mode

	dim taxIdx, account_idx
	dim sdate, edate, chkTerm
	dim page, searchDiv, searchKey, searchString, param

	dim ofranchulgomaster
	dim ofranchulgojungsan
	dim opartner
	dim ogroup

	dim oTax, i, lp

	dim Ftenten_manager_name
	dim Ftenten_manager_phone
	dim Ftenten_manager_email

	dim Fetcstring

	dim taxtype



	'==========================================================================
	account_idx = request("idx")
	mode = request("mode")
	taxtype = request("taxtype")


	if (mode = "") then
		mode = "02"
	end if

	if (taxtype = "") then
		taxtype = "Y"
	end if

	if (mode = "01") then
		Ftenten_manager_name = "신희영"
		Ftenten_manager_phone = "02-1644-6030"
		Ftenten_manager_email = "accounts@10x10.co.kr"
	elseif (mode = "02") then
		Ftenten_manager_name = "신희영"
		Ftenten_manager_phone = "02-554-2033"
		Ftenten_manager_email = "accounts@10x10.co.kr"
	else
		Ftenten_manager_name = "신희영"
		Ftenten_manager_phone = "02-554-2033"
		Ftenten_manager_email = "accounts@10x10.co.kr"
	end if



	'==========================================================================
	'정산정보
	set ofranchulgomaster = new CFranjungsan
	ofranchulgomaster.FRectidx = account_idx

	ofranchulgomaster.getOneFranJungsan

	'ofranchulgomaster.FOneItem.Ftotalsum '총 발행금액을 총공급가로 함.(부가세포함금액)



	'==========================================================================
	'삽아이디에서 그룹코드 추출
	set opartner = new CPartnerUser

	opartner.FCurrpage = 1
	opartner.FPageSize = 100
	opartner.FRectDesignerID = ofranchulgomaster.FOneItem.Fshopid

	opartner.GetPartnerNUserCList



	'==========================================================================
	'그룹코드에서 세금계산서/정산담당자 정보 추출
	set ogroup = new CPartnerGroup

	ogroup.FRectGroupid = opartner.FPartnerList(0).FGroupID

	ogroup.GetOneGroupInfo



	'==========================================================================
	Fetcstring = CStr(account_idx)



	'==========================================================================
	''기발행 세금계산서인지 체크

	set oTax = new CTax

	oTax.FRectsearchKey = " t1.orderidx "
	oTax.FRectsearchString = CStr(account_idx)

	oTax.GetTaxList

	if oTax.FResultCount > 0 then
		if oTax.FTaxList(0).FisueYn="Y" then
			response.write "<script>alert('이미 발행된 세금계산서가 있습니다.\n\n재발행 하시려면 거래처에 [취소요청]후 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다');</script>"
		else
			response.write "<script>alert('발행대기중인 세금계산서가 있습니다.\n\n재발행 하시려면 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다.');</script>"
		end if
	end if

%>
<script language="javascript">

function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

function doRegisterSheet(){
<% if (ogroup.FResultCount = 1) then %>
	<% if oTax.FResultCount > 0 then %>
		<% if oTax.FTaxList(0).FisueYn="Y" then %>
			alert('이미 발행된 세금계산서가 있습니다.\n\n재발행 하시려면 거래처에 [취소요청]후 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다');
		<% else %>
			alert('발행대기중인 세금계산서가 있습니다.\n\n재발행 하시려면 매출 세금계산서 목록에서 [삭제]후 발행하셔야 합니다.');
		<% end if %>
	<% else %>
    if (document.frm.yyyymmdd_register.value == "") {
    	alert("작성일을 입력하세요.");
    	return;
    }

    if (confirm('세금계산서를 작성하시겠습니까?')){
        document.frm.submit();
    }
	<% end if %>
<% else %>
	alert("그룹코드가 지정되어 있지 않은 업체입니다.");
<% end if %>
}

function ChangePage(frm){
    location.href = "?mode=" + frm.mode.value + "&idx=" + <%= account_idx %> + "&taxtype=" + frm.taxtype.value;
}

function CalcPriceWithPrice111()
{
	if (frm.totalpricesum.value == "") { return; }

	if (frm.taxtype.value.length<1){
		alert('과세구분을 입력하세요.');
		return;
	}

	if (frm.totalpricesum.value*0 != 0) { alert("잘못된 값을 입력했습니다."); return; }

	frm.totalpricesum2.value = frm.totalpricesum.value;
	frm.totalpricesum3.value = frm.totalpricesum.value;

	if (frm.taxtype.value == "Y") {
		// 세액은 공급가를 구하고 0.1 후 반올림 해주면 된다.
		frm.totaltax.value = Math.round(1.0 * frm.totalpricesum.value / 1.1 / 10.0);
		frm.totaltaxsum.value = frm.totaltax.value;
	} else {
		frm.totaltax.value = 0;
		frm.totaltaxsum.value = 0;
	}

	frm.totalsuply.value = frm.totalpricesum.value - frm.totaltax.value;
	frm.totalsuply2.value = frm.totalsuply.value;
	frm.totalsuplysum.value = frm.totalsuply.value;
}
</script>

<!-- 세금계산 요청서 정보 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<form name="frm" method="post" action="doTaxOrder.asp">
	<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<input type="hidden" name="idx" value="<%= account_idx %>">
		<td colspan="4" align="left">
			<b>가맹점 세금계산서 발행</b>
		</td>
	</tr>
</table>

<br>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr valign="top">
        <td width="49%">
        	<!-- 공급자정보 시작 -->
        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4" height="25"><b>공급자 정보</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">사업자번호</td>
        			<td colspan="3"><b>211-87-00620</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="70" bgcolor="#F0F0FD" height="25">상호</td>
        			<td><b>(주)텐바이텐</b></td>
        			<td width="70" bgcolor="#F0F0FD">대표자</td>
        			<td><b>이문재</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">사업장주소</td>
        			<td colspan="3">서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">업태</td>
        			<td>서비스,도소매 등</td>
        			<td bgcolor="#F0F0FD">종목</td>
        			<td>전자상거래 등</td>
        		</tr>
        		<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">담당자</td>
        			<td><%= Ftenten_manager_name %></td>
        			<td bgcolor="#F0F0FD">연락처</td>
        			<td><%= Ftenten_manager_phone %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">이메일</td>
        			<td><%= Ftenten_manager_email %></td>
        			<td bgcolor="#F0F0FD">BILL아이디</td>
        			<td>
        			 	<select class="select" name="mode" onchange="ChangePage(frm)">
							<option value="02" <% if (mode = "02") then %>selected<% end if %>>가맹점(ACCOUNTS)</option>
							<option value="03" <% if (mode = "03") then %>selected<% end if %>>프로모션(PROMOTION)</option>
						</select>
        			</td>
        		</tr>
        	</table>
        	<!-- 공급자정보 끝 -->
        </td>
        <td>&nbsp;</td>
        <td width="49%">
        	<!-- 공급받는자정보 시작 -->
        	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4" height="25"><b>공급받는자 정보</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">사업자번호</td>
        			<td colspan="3"><b><%= ogroup.FOneItem.Fcompany_no %></b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="70" bgcolor="#F0F0FD" height="25">상호</td>
        			<td><b><%= ogroup.FOneItem.FCompany_name %></b></td>
        			<td width="70" bgcolor="#F0F0FD">대표자</td>
        			<td><b><%= ogroup.FOneItem.Fceoname %></b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">사업장주소</td>
        			<td colspan="3"><%= ogroup.FOneItem.Fcompany_address %> <%= ogroup.FOneItem.Fcompany_address2 %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">업태</td>
        			<td><%= ogroup.FOneItem.Fcompany_uptae %></td>
        			<td bgcolor="#F0F0FD">종목</td>
        			<td><%= ogroup.FOneItem.Fcompany_upjong %></td>
        		</tr>
        		<tr><td height="1" colspan="4" bgcolor="#FFFFFF"></td></tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">담당자</td>
        			<td><%= ogroup.FOneItem.Fjungsan_name %></td>
        			<td bgcolor="#F0F0FD">연락처</td>
        			<td><%= ogroup.FOneItem.Fjungsan_hp %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td bgcolor="#F0F0FD" height="25">이메일</td>
        			<td colspan="3"><%= ogroup.FOneItem.Fjungsan_email %></td>
        		</tr>
        	</table>
        	<!-- 공급받는자정보 끝 -->
        </td>
	</tr>
</table>

<p>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#F0F0FD">
		<td width="120" height="25">작성일</td>
		<td width="100">공급가액</td>
		<td width="100">과세구분</td>
		<td width="100">세액</td>
		<td width="100">합계금액</td>
		<td>비고</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td height="25"><input type="text" size="10" name="yyyymmdd_register" value="" onClick="jsPopCal('frm','yyyymmdd_register');" style="cursor:hand;"></td>
<% if (taxtype = "Y") then %>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum/1.1),0) %></td>
<% else %>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></td>
<% end if %>
		<td>
			<select name=taxtype class="writebox" onchange="ChangePage(frm)">
			<option value="Y" <% if (taxtype = "Y") then %>selected<% end if %>>과세</option>
			<option value="N" <% if (taxtype = "N") then %>selected<% end if %>>면세</option>
			<option value="0" <% if (taxtype = "0") then %>selected<% end if %>>영세</option>
			</select>
		</td>
<% if (taxtype = "Y") then %>
		<td><%= FormatNumber(((ofranchulgomaster.FOneItem.Ftotalsum/1.1)*0.1),0) %></td>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></td>
<% else %>
		<td>0</td>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></td>
<% end if %>
		<td>인덱스코드 : <%= Fetcstring %></td>
	</tr>
</table>

<p>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#F0F0FD">
		<td width="30" height="25">월</td>
		<td width="30">일</td>
		<td>품목</td>
		<td width="50">규격</td>
		<td width="50">수량</td>
		<td width="100">단가</td>
		<td width="100">공급가액</td>
		<td width="100">세액</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td height="25"></td>
		<td></td>
		<td><%= ofranchulgomaster.FOneItem.Ftitle %></td>
		<td></td>
		<td>1</td>
<% if (taxtype = "Y") then %>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum/1.1),0) %></td>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum/1.1),0) %></td>
		<td><%= FormatNumber(((ofranchulgomaster.FOneItem.Ftotalsum/1.1)*0.1),0) %></td>
<% else %>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></td>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></td>
		<td>0</td>
<% end if %>
	</tr>
</table>

<p>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#F0F0FD">
		<td height="25"><b>합계금액</b></td>
		<td width="100">현금</td>
		<td width="100">수표</td>
		<td width="100">어음</td>
		<td width="100">외상미수금</td>
	</tr>
    <tr align="center" bgcolor="#FFFFFF">
		<td height="25"><b><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></b></td>
		<td></td>
		<td></td>
		<td></td>
		<td><%= FormatNumber((ofranchulgomaster.FOneItem.Ftotalsum),0) %></td>
	</tr>
	</form>
</table>
<br>

<p>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="#FFFFFF">
		<td align="center" height="25">
		  <input type="button" class="button" value="작성" onClick="doRegisterSheet()">
		  &nbsp;
		  <input type="button" class="button" value="목록" onClick="self.location='Tax_list.asp'">
		</td>
	</tr>
</table>





<!-- 세금계산 요청서 정보 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->