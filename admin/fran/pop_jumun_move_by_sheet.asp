<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  매장대 매장 재고이동
' History : 2018.02.07 이상구 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%

dim i, j, k
dim idx

idx  = requestCheckVar(request("idx"), 3200)

%>
<script>
function jsChkForm(frm) {
	if (frm.moveshopid.value == "") {
		alert("에러!!\n\n도착매장 지정안됨.");
		return false;
	}

	if (frm.scheduledt.value.length<1){
		alert('재고이동일을 입력하세요');
		calendarOpen3(frm.scheduledt,'재고이동일을 입력하세요','');
		return false;
	}

	if (frm.songjangdiv.value.length<1){
		alert('택배사를 선택 하세요');
		frm.songjangdiv.focus();
		return false;
	}

	if (frm.songjangno.value.length<1){
		alert('송장 번호를 입력 하세요');
		msfrm.songjangno.focus();
		return false;
	}

	return true;
}

function jsStockMove() {
	var frm = document.frm;

	if (jsChkForm(frm) != true) { return; }

	var ret = confirm('입력하신대로 재고 이동처리 하시겠습니까?');
	if (ret) {
		frm.mode.value = "saveorderbysheet";
		frm.method = "post";
		frm.action = "pop_jumun_move_process.asp";
		frm.submit();
	}
}
</script>
<form name="frm" method="get" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idx" value="<%= idx %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
	    <td height="25" bgcolor="<%= adminColor("tabletop") %>">도착매장</td>
	    <td>
		    <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("moveshopid", "", "21") %>
	    </td>
    </tr>
    <tr bgcolor="#FFFFFF">
	    <td bgcolor="<%= adminColor("tabletop") %>">
		    재고이동일
	    </td>
	    <td>
		    <input type="text" class="text" name="scheduledt" value="" size=10 readonly ><a href="javascript:calendarOpen(frm.scheduledt);">
		        <img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>

			    택배사 :<% drawSelectBoxDeliverCompany "songjangdiv", "" %>
			    송장번호:<input type="text" class="text" name="songjangno" size=14 maxlength=16 value="" >
			    <br>
			    (택배로 보내지 않을경우 택배사:기타선택 송장번호:퀵배송, 직접배송 등을 입력 하시면 됩니다.)
	    </td>
    </tr>
    <tr bgcolor="#FFFFFF">
	    <td bgcolor="<%= adminColor("tabletop") %>" colspan="2" align="center">
	        <input type="button" value="재고이동처리" onClick="jsStockMove()" class="button" id="btnMove">
	    </td>
    </tr>
</table>
</form>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
