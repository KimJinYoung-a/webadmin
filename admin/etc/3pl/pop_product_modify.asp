<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/3pl/productCls.asp" -->
<!-- #include virtual="/lib/classes/3pl/companyCls.asp" -->
<!-- #include virtual="/lib/classes/3pl/common.asp" -->
<%

dim companyid, prdcode, itemgubun

companyid 	= requestCheckVar(request("companyid"),32)
prdcode 	= requestCheckVar(request("prdcode"),32)


dim oCTPLProduct
set oCTPLProduct = New CTPLProduct
	oCTPLProduct.FRectCompanyID	= companyid
	oCTPLProduct.FRectPrdCode	= prdcode

oCTPLProduct.GetTPLProductOne

if (prdcode = "") then
	oCTPLProduct.FOneItem.Fuseyn = "Y"
	oCTPLProduct.FOneItem.Fregdate = Now()
	oCTPLProduct.FOneItem.Flastupdt = Now()
else
	companyid = oCTPLProduct.FOneItem.Fcompanyid
	itemgubun = oCTPLProduct.FOneItem.Fitemgubun
end if

dim oCTPLCompany
set oCTPLCompany = New CTPLCompany
	oCTPLCompany.FRectCompanyID					= companyid

if (companyid <> "") and (prdcode = "") then
	oCTPLCompany.GetTPLCompanyOne
	itemgubun = oCTPLCompany.FOneItem.Fcompanygubun
end if

%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript">

function SubmitForm() {
	var frm = document.frm;

	if (validate(frm)==false) {
		return;
	}

	if (frm.companyid.value == '') {
		alert('고객사를 선택하세요.');
		return;
	}

	if (frm.brandid.value == '') {
		alert('브랜드를 선택하세요.');
		return;
	}

	if ((frm.itemoption.value == '') && (frm.itemoptionname.value == '')) {
		alert('고객사 옵션코드/옵션명 둘중 하나는 반드시 등록해야 합니다.');
		return;
	}
	if (frm.useyn.value == '') {
		alert('사용여부를 선택하세요.');
		return;
	}

    if (confirm("저장하시겠습니까?") == true) {
        frm.submit();
    }
}

function SubmitNext() {
	var frm = document.frm;

	if (frm.companyid.value == '') {
		alert('고객사를 선택하세요.');
		return;
	}

	frm.action = '';
	frm.submit();
}

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="product_process.asp">
<input type="hidden" name="mode" value="<%= CHKIIF(prdcode<>"", "modi", "ins") %>">

<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="300">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>상품 정보</b>
			    </td>
			    <td align="right">
			    	<input type="button" value="저장하기" class="csbutton" onclick="javascript:SubmitForm();">
			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>" width="150">고객사</td>
    <td>
		<% if (companyid = "") then %>
		<% Call SelectBoxCompanyID("companyid", oCTPLProduct.FOneItem.Fcompanyid, CHKIIF(companyid<>"", "", "Y")) %>
		<input type="button" value="다음단계" class="csbutton" onclick="javascript:SubmitNext();">
		<% else %>
		<%= companyid %>
		<input type="hidden" name="companyid" value="<%= companyid %>">
		<% end if %>
	</td>
</tr>
<% if (companyid <> "") then %>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">브랜드</td>
    <td>
		<% Call SelectBoxEngBrandID(companyid, "brandid", oCTPLProduct.FOneItem.Fbrandid, CHKIIF(oCTPLProduct.FOneItem.Fbrandid<>"", "", "Y")) %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">물류코드</td>
    <td>
		<% if (prdcode = "") then %>
		--
		<% else %>
		<%= FormatPrdCode(oCTPLProduct.FOneItem.Fprdcode) %>
		<input type="hidden" name="prdcode" value="<%= prdcode %>">
		<% end if %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">상품명</td>
    <td>
		<input type="text" class="text" name="prdname" id="[on,off,1,32][상품명]" value="<%= oCTPLProduct.FOneItem.Fprdname %>" size="64">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">옵션명</td>
    <td>
		<input type="text" class="text" name="prdoptionname" id="[off,off,0,32][옵션명]" value="<%= oCTPLProduct.FOneItem.Fprdoptionname %>" size="32">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">소비자가</td>
    <td>
		<input type="text" class="text" name="customerprice" id="[off,on,1,off][소비자가]" value="<%= oCTPLProduct.FOneItem.Fcustomerprice %>">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">범용바코드</td>
    <td>
		<input type="text" class="text" name="generalbarcode" id="[off,off,0,32][범용바코드]" value="<%= oCTPLProduct.FOneItem.Fgeneralbarcode %>">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">고객사코드</td>
    <td>
		<%= itemgubun %>
		<input type="hidden" name="itemgubun" value="<%= itemgubun %>" id="[on,off,2,4][고객사코드]">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">고객사상품코드</td>
    <td>
		<input type="text" class="text" name="itemid" id="[on,off,1,32][고객사상품코드]" value="<%= oCTPLProduct.FOneItem.Fitemid %>">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">고객사옵션코드</td>
    <td>
		<input type="text" class="text" name="itemoption" id="[off,off,0,32][고객사옵션코드]" value="<%= oCTPLProduct.FOneItem.Fitemoption %>">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">고객사옵션명</td>
    <td>
		<input type="text" class="text" name="itemoptionname" id="[off,off,0,32][고객사옵션명]" value="<%= oCTPLProduct.FOneItem.Fitemoptionname %>" size="32">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">사용여부</td>
    <td>
		<% Call drawSelectBoxUsingYN("useyn", oCTPLProduct.FOneItem.Fuseyn) %>
	</td>
</tr>
<% end if %>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">등록일</td>
    <td>
		<%= oCTPLProduct.FOneItem.Fregdate %>
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">최종수정</td>
    <td>
		<%= oCTPLProduct.FOneItem.Flastupdt %>
	</td>
</tr>
</table>

<%
set oCTPLProduct = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
