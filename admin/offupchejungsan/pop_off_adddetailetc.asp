<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim makerid, shopid, gubuncd, masteridx
makerid     = request("makerid")
shopid      = request("shopid")
gubuncd     = request("gubuncd")
masteridx   = request("masteridx")
%>
<script language='javascript'>
function AddThis(frm){
	if (frm.sellprice.value.length<1){
		alert('판매가를 입력하세요.');
		frm.sellprice.focus();
		return;
	}

	if (frm.suplyprice.value.length<1){
		alert('매입가를 입력하세요.');
		frm.suplyprice.focus();
		return;
	}

	if (frm.commission.value.length<1){
		alert('수수료를 입력하세요.');
		frm.commission.focus();
		return;
	}


	if (frm.itemno.value.length<1){
		alert('갯수를 입력하세요.');
		frm.itemno.focus();
		return;
	}

	if (confirm('기타 내역을 추가하시겠습니까?')){
		frm.submit();
	}
}

function calCommission() {
	var frm = document.frm;
	var sellprc = frm.sellprice.value;
	var suplyprc = frm.suplyprice.value;
	if(!sellprc){sellprc=0;}
	if(!suplyprc){suplyprc=0;}
	frm.commission.value = parseInt(sellprc)-parseInt(suplyprc);
	frm.sellprice.value = parseInt(sellprc);
	frm.suplyprice.value = parseInt(suplyprc);
}
</script>
<table border=0 cellspacing="1" class="a"  width=500 bgcolor=#3d3d3d>
<form name=frm method=post action="off_jungsan_process.asp">
<input type=hidden name=mode value="addetcdetail">
<input type=hidden name=gubuncd value="B999">
<input type=hidden name=shopid value="<%= shopid %>">
<input type=hidden name=masteridx value="<%= masteridx %>">
<tr>
	<td width=120 bgcolor="#DDDDFF">상품코드</td>
	<td bgcolor="#FFFFFF">
	<input type=text name="itemgubun" value="00" size=2 maxlength=2 >
	<input type=text name="itemid" value="000000" size=9 maxlength=9 >
	<input type=text name="itemoption" value="0000" size=4 maxlength=4 >
	</td>
</tr>
<tr>
	<td width=120 bgcolor="#DDDDFF">상품명</td>
	<td bgcolor="#FFFFFF"><input type=text name="itemname" value="" size=26 maxlength=40></td>
</tr>
<tr>
	<td width=120 bgcolor="#DDDDFF">옵션명</td>
	<td bgcolor="#FFFFFF"><input type=text name="itemoptionname" value="" size=26 maxlength=40></td>
</tr>
<tr>
	<td width=120 bgcolor="#DDDDFF">정산아이디</td>
	<td bgcolor="#FFFFFF"><input type=text name="makerid" value="<%= makerid %>" size=26 maxlength=32></td>
</tr>

<tr>
	<td width=120 bgcolor="#DDDDFF">판매가</td>
	<td bgcolor="#FFFFFF">
	<input type=text name="sellprice" value="" size=9 maxlength=9 style="text-align:right" onkeyup="calCommission();" />(판매상품이 아닌경우 0원)
	</td>
</tr>
<tr>
	<td width=120 bgcolor="#DDDDFF">수수료</td>
	<td bgcolor="#FFFFFF">
	<input type=text name="commission" value="0" size=9 maxlength=9 style="text-align:right" readOnly >
	</td>
</tr>
<tr>
	<td width=120 bgcolor="#DDDDFF">매입가(정산액)</td>
	<td bgcolor="#FFFFFF"><input type=text name="suplyprice" value="" size=9 maxlength=9 style="text-align:right" onkeyup="calCommission();" /></td>
</tr>
<tr>
	<td width=120 bgcolor="#DDDDFF">갯수</td>
	<td bgcolor="#FFFFFF"><input type=text name="itemno" value="" size=3 maxlength=5 ></td>
</tr>
<tr>
	<td colspan=2 align=center bgcolor="#FFFFFF"><input type=button value=" 저 장 " onclick="AddThis(frm)"></td>
</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->