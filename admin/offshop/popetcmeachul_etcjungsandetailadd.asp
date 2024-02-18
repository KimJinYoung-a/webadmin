<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 매출등록
' History : 서동석 생성
'			2017.04.12 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<%
dim idx, topidx, makerid, shopid

idx  = requestCheckVar(request("idx"),10)
topidx = requestCheckVar(request("topidx"),10)
makerid = requestCheckVar(request("makerid"),32)
shopid = requestCheckVar(request("shopid"),32)

%>
<script language='javascript'>
function AddValue(frm){
<% if (C_ADMIN_AUTH) then %>

<% else %>
	if (!ajaxBrandItem())
	{
		return;
	}
<% end if %>

	if (frm.itemno.value.length<1){
		alert('갯수를 입력 하세요.');
		frm.itemno.focus();
		return;
	}

	if (frm.orgsellcash.value.length<1){
		alert('소비가를 입력 하세요.');
		frm.orgsellcash.focus();
		return;
	}


	if (frm.sellcash.value.length<1){
		alert('판매가를 입력 하세요.');
		frm.sellcash.focus();
		return;
	}

	if (frm.buycash.value.length<1){
		alert('매입가를 입력 하세요.');
		frm.buycash.focus();
		return;
	}

	if (frm.suplycash.value.length<1){
		alert('공급가를 입력 하세요.');
		frm.suplycash.focus();
		return;
	}

	frm.submit();
}

function checkBrandItem()
{
	if (event.keyCode==13)
		ajaxBrandItem("GET");
}

function ajaxBrandItem(mode)
{
	var f = document.frm;
	if (f.itemgubun.value.length!=2){
		alert('상품구분을 입력 하세요.');
		f.itemgubun.focus();
		return false;
	}

	if (frm.itemid.value.length<1){
		alert('상품번호를 입력 하세요.');
		f.itemid.focus();
		return false;
	}

	if (frm.itemoption.value.length!=4){
		alert('상품옵션코드를 입력 하세요.');
		f.itemoption.value='0000';
		f.itemoption.focus();
		return false;
	}

	var url = "ajaxBrandItem.asp?shopid=<%=shopid%>&makerid=" + f.makerid.value + "&itemgubun=" + f.itemgubun.value + "&itemid=" + f.itemid.value + "&itemoption=" + f.itemoption.value;
	var xmlHttp = createXMLHttpRequest();
	xmlHttp.open("GET", url, false);
	xmlHttp.send('');
	if(xmlHttp.status == 200) {
		var arr = xmlHttp.responseText.split("|");
		//alert(xmlHttp.responseText);
		if (arr[0]=="Y")
		{
			if (mode=="GET")
			{
				f.itemname.value		= arr[1];
				f.itemoptionname.value	= arr[2];
				f.orgsellcash.value		= arr[3];
				f.sellcash.value		= arr[4];
				f.buycash.value		    = arr[5];
				f.suplycash.value		= arr[6];

			}
			return true;
		}
		else if (arr[0]=="N")
		{
			alert("삭제된 상품코드입니다.");
			f.itemid.focus();
			return false;
		}
		else
		{
			alert("브랜드ID와 일치하지 않는 상품코드입니다.");
			f.itemid.focus();
			return false;
		}
	}
}



// ajax용 객체 함수
function createXMLHttpRequest() {
    if (window.ActiveXObject) {
        xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
    }
    else if (window.XMLHttpRequest) {
        xmlHttp = new XMLHttpRequest();
    }
	return xmlHttp;
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name=frm method=post action="etc_meachul_process.asp">
	<input type=hidden name="mode" value="etcsubdetailadd">
	<input type=hidden name="topidx" value="<%= topidx %>">
	<input type=hidden name="idx" value="<%= idx %>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width="80">구분</td>
		<td bgcolor="#FFFFFF" >
			<select class="select" name="linkbaljucode">
				<option value="">일반매입
				<option value="witak">위탁
				<option value="bojung" selected >할인보정
				<option value="etc2">기타
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
		<td bgcolor="#FFFFFF" >
		<input type="text" class="text_ro" name='makerid' value="<%= makerid %>" size=32 maxlength=30 readonly>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
		<td bgcolor="#FFFFFF" >
		<input type="text" class="text" name='itemgubun' value="" size=2 maxlength=2 onkeydown="checkBrandItem();">-
		<input type="text" class="text" name='itemid' value="" size=9 maxlength=9 onkeydown="checkBrandItem();">-
		<input type="text" class="text" name='itemoption' value="" size=4 maxlength=4 onkeydown="checkBrandItem();">
		<input type="button" class="button" value="브랜드상품체크" onclick="ajaxBrandItem('GET');">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='itemname' value="" size=32 maxlength=32></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">옵션명</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='itemoptionname' value="" size=32 maxlength=32></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">수량</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='itemno' value="" size=4 maxlength=9></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">소비가</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='orgsellcash' value="" size=10 maxlength=9></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">판매가</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='sellcash' value="" size=10 maxlength=9></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">매입가</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='buycash' value="" size=10 maxlength=9></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">공급가</td>
		<td bgcolor="#FFFFFF" ><input type="text" class="text" name='suplycash' value="" size=10 maxlength=9></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF" colspan=2 align=center>
			<input type="button" class="button" value="내역추가" onclick="AddValue(frm)">
		</td>
	</tr>
	</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
