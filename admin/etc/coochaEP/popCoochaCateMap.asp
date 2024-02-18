<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/coochaEp/epShopCls.asp"-->
<%
Dim oCoocha, i, idx
idx = request("idx")

'// 카테고리 내용 접수
Set oCoocha = new epShop
	oCoocha.FRectIdx = idx
	oCoocha.getCoochaMapList
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function popDispCateSelect(){
	$.ajax({
		url: "/admin/etc/coochaEP/act_DispCategorySelect.asp?isDft=0",
		cache: false,
		success: function(message) {
			$("#lyrDispCateAdd").empty().append(message).fadeIn();
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

// 레이어에서 전시카테고리 추가
function addDispCateItem(dcd,cnm,div,dpt) {

	
	// 행추가
	var oRow = tbl_DispCate.insertRow();
	oRow.onmouseover=function(){tbl_DispCate.clickedRowIndex=this.rowIndex};

	// 셀추가 (구분,카테고리,삭제버튼)
	var oCell2 = oRow.insertCell();
	var oCell3 = oRow.insertCell();

	$(cnm).each(function(i){
		if(dpt>i) {
			if(i>0) oCell2.innerHTML += " >> ";
			oCell2.innerHTML += $(this).text();
		}
	});
	oCell2.innerHTML += "<input type='hidden' name='catecode' value='" + dcd + "'>";
	oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' align=absmiddle>";
	$("#lyrDispCateAdd").fadeOut();
}

// 선택 전시카테고리 삭제
function delDispCateItem() {
	if(confirm("선택한 카테고리를 삭제하시겠습니까?")) {
		tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);
	}
}

// 매칭 저장하기
function fnSaveForm() {
	var frm = document.srcFrm;

    // 카테고리 지정여부 검사
	if(tbl_DispCate.rows.length < 1)	{
		alert("카테고리를 선택해주세요.");
		return;
	}

	if(confirm("선택하신 카테고리로 매칭하시겠습니까?")) {
		frm.mode.value="saveCate";
		frm.action="procCoocha.asp";
		frm.submit();
	}
}

//매칭 전체 삭제하기
function fnDelForm() {
	var frm = document.srcFrm;

	if(confirm("선택하신 카테고리의 매칭을 전부 삭제하시겠습니까?")) {
		frm.mode.value="delCate";
		frm.action="procCoocha.asp";
		frm.submit();
	}
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>쿠차 카테고리 매칭</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<p>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 상단바 끝-->
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 쿠차 카테고리 정보</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">중분류</td>
	<td bgcolor="#FFFFFF"><%= oCoocha.FOneItem.FDEPTH1NM %></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">소분류</td>
	<td bgcolor="#FFFFFF"><%= oCoocha.FOneItem.FDEPTH2NM %></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">3DEPTH</td>
	<td bgcolor="#FFFFFF"><%= oCoocha.FOneItem.FDEPTH3NM %></td>
</tr>
</table>
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 텐바이텐 전시 카테고리 매칭 정보 </td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<form name="srcFrm" method="post" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="mode" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="2" >카테고리</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td>
				<td><%=getDispCategory(idx)%></td>
			</td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
</table>
</form>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="28">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left"></td>
    <td valign="bottom" align="right">
		<img src="/images/icon_save.gif" width="45" height="20" border="0" onclick="fnSaveForm()" style="cursor:pointer" align="absmiddle">
		<img src="/images/icon_delete.gif" width="45" height="20" border="0" onclick="fnDelForm()" style="cursor:pointer" align="absmiddle">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td colspan="2" background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->