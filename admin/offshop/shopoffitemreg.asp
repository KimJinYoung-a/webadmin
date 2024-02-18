<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 상품등록
' History : 서동석 생성
'			2017.04.12 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim designer,react
	designer = requestCheckVar(request("designer"),32)
	react = requestCheckVar(request("react"),10)

response.write "<script type='text/javascript'>location.href='/admin/offshop/popoffitemreg.asp?makerid=" + designer + "';</script>"
dbget.close()	:	response.End

%>
<script type='text/javascript'>

function refreshParent(){
	opener.frm.submit();
}

function AddOffItem(frm){
	if ((frm.itemgubun[0].checked==false)&&(frm.itemgubun[1].checked==false)){
		alert('상품구분을 선택하세요.');
		return;
	}

	if (frm.designer.value.length<1){
		alert('브랜드를 선택하세요.');
		return;
	}

	if (frm.itemname.value.length<1){
		alert('아이템명을 하세요.');
		frm.itemname.focus();
		return;
	}

	if (!IsDigit(frm.sellcash.value)){
		alert('판매가는 숫자만 가능합니다.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDigit(frm.suplycash.value)){
		alert('업체 매입가는 숫자만 가능합니다.');
		frm.suplycash.focus();
		return;
	}

	if (!IsDigit(frm.shopbuyprice.value)){
		alert('샾 공급가는 숫자만 가능합니다.');
		frm.shopbuyprice.focus();
		return;
	}

	if (((frm.suplycash.value!=0)||(frm.shopbuyprice.value!=0))){
		if (!confirm('!! 기본 계약 마진과 다를 경우에만 매입가 공급가를 입력 하셔야 합니다. \n\n계속 하시겠습니까?')){
			return;
		}
	}

	var ret = confirm('추가하시겠습니까?');

	if (ret) {
		frm.submit();
	}
}
</script>

<div align="center">
<br>
<table width="400" cellspacing="1" class="a" bgcolor=#3d3d3d>
<form name="frmadd" method="post" action="shopitem_process.asp">
<input type="hidden" name="mode" value="offitemreg">
<tr bgcolor="#FFFFFF">
	<td colspan=2>
	<table border=0 cellspacing=0 cellpadding=0 class="a" >
	<tr>
		<td width=110>○오프샾 전용상품 </td>
		<td>:온라인 상품과 별개로 진행 할 경우.</td>
	</tr>
	<tr>
		<td>○이벤트상품 </td>
		<td>:업체 기본 공급마진과 공급마진이 다른경우.<br><b>(공급가 필히 입력)</b></td>
	</tr>
	<tr>
		<td>○소모품 </td>
		<td>:기타 소모품.</td>
	</tr>
	<tr>
		<td>○가맹점전용상품 </td>
		<td>:가맹점에서 직접 판매하는상품.</td>
	</tr>
	</table>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">구분</td>
	<td bgcolor="#FFFFFF">
	<input type="radio" name="itemgubun" value="90">오프샾 전용상품(90)<br>
	<input type="radio" name="itemgubun" value="70">소모품(70)<br>
	<!--
	<input type="radio" name="itemgubun" value="80" disabled >이벤트상품(80) : 사용안함<br>
	<input type="radio" name="itemgubun" value="95" disabled >가맹점전용상품(95) : 사용안함
	-->
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">브랜드 ID</td>
	<td bgcolor="#FFFFFF"><% drawSelectBoxDesignerwithName "designer",designer  %></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">상품명</td>
	<td bgcolor="#FFFFFF"><input type="text" name="itemname" value="" maxlength="32"></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">판매가격</td>
	<td bgcolor="#FFFFFF"><input type="text" name="sellcash" value="" maxlength="9"></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">매입가</td>
	<td bgcolor="#FFFFFF"><input type="text" name="suplycash" value="0" size=6 maxlength="9"><br><b>(0일경우 계약 마진에 의해 자동설정됩니다.)</b></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">샾공급가</td>
	<td bgcolor="#FFFFFF"><input type="text" name="shopbuyprice" value="0" size=6 maxlength="9"><br><b>(0일경우 계약 마진에 의해 자동설정됩니다.)</b></td>
</tr>
<tr>
	<td colspan="2" align="center" bgcolor="#FFFFFF"><input type="button" value="추가" onclick="AddOffItem(frmadd)"></td>
</tr>
</form>
</table>
</div>

<% if react="true" then %>
<!-- <script type='text/javascript'>refreshParent();</script> -->
<% end if %>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->