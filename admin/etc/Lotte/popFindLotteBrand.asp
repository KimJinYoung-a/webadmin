<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<%
	dim oLotte
	dim TenMakerid, lotteBrandCd, lotteBrandName

	TenMakerid		= request("mkid")

	if TenMakerid<>"" then
		'// 목록 접수
		Set oLotte = new cLotte
		oLotte.FPageSize = 20
		oLotte.FCurrPage = 1
		oLotte.FRectMakerid = TenMakerid
		oLotte.getLotteBrandList
		if oLotte.FResultCount>0 then
			lotteBrandCd = oLotte.FItemList(0).FlotteBrandCd
			lotteBrandName = oLotte.FItemList(0).FlotteBrandName
		end if
		Set oLotte = Nothing
	end if
%>
<script language="javascript">
<!--
	// 롯데닷컴 브랜드 검색
	function fnSearchLotteBrand() {
		if(!fsrch.brnNm.value) {
			alert("검색어를 입력해주세요.(ex.브랜드명)");
			fsrch.brnNm.focus();
			return;
		}
		var pFBL = window.open("","popLotteBrand","width=400,height=500,scrollbars=yes,resizable=yes");
		pFBL.focus();
		fsrch.target="popLotteBrand";
		fsrch.action="actFindLotteBrand.asp";
		fsrch.submit();
	}

	// 매칭 저장하기
	function fnSaveForm() {
		var frm = document.frm;
		if(frm.TenMakerid.value=="") {
			alert("매칭할 텐바이텐 브랜드를 선택해주세요.");
			frm.TenMakerid.focus();
			return;
		}

		if(frm.lotteBrandCd.value=="") {
			alert("매칭할 롯데닷컴 브랜드를 선택해주세요.");
			return;
		}

		if(confirm("선택하신 브랜드를 서로 매칭하시겠습니까?")) {
			frm.mode.value="save";
			frm.action="procLotteBrandMap.asp";
			frm.submit();
		}
	}

	// 매칭정보 삭제
	function fnDelForm() {
		var frm = document.frm;
		<% if TenMakerid="" then %>
		alert("매칭된 브랜드가 없습니다.");
		return;
		<% else %>
		if(confirm("브랜드를 매칭 해제하시겠습니까?")) {
			frm.mode.value="del";
			frm.action="procLotteBrandMap.asp";
			frm.submit();
		}
		<% end if %>
	}

	// 창닫기
	function fnCancel() {
		if(confirm("작업을 취소하고 창을 닫으시겠습니까?")) {
			self.close();
		}
	}
//-->
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
	<font color="red"><strong>롯데닷컴 브랜드 매칭</strong></font></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 텐바이텐 브랜드 정보</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<form name="frm" method="get" action="" target="xLink" style="margin:0px;">
<input type="hidden" name="mode" value="save">
<input type="hidden" name="lotteBrandCd" value="<%=lotteBrandCd%>">
<input type="hidden" name="lotteBrandNm" value="<%=lotteBrandName%>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">브랜드ID</td>
	<td bgcolor="#FFFFFF"><% drawSelectBoxDesignerwithName "TenMakerid",TenMakerid %></td>
</tr>
</table>
</form>
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 롯데닷컴 브랜드 정보</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<form name="fsrch" method="get" action="actFindLotteBrand.asp" style="margin:0px;" onsubmit="return false;">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td id="brTT" <%=chkIIF(TenMakerid<>"","rowspan=2","")%> width="80" align="center" bgcolor="#DDDDFF">브랜드 검색</td>
	<td bgcolor="#FFFFFF">브랜드명 <input type="text" name="brnNm"  size="12" class="text">
		<input type="button" value="검색" class="button" onclick="fnSearchLotteBrand()">
	</td>
</tr>
<tr id="BrRow" height="25" <%=chkIIF(TenMakerid<>"","","style='display:none;'")%>>
	<td bgcolor="#FFFFFF">
		선택 브랜드 : <span id="selBr">[<%=lotteBrandCd%>]<%=lotteBrandName%></span>
	</td>
</tr>
</table>
</form>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="28">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left"><img src="http://testwebadmin.10x10.co.kr/images/icon_delete.gif" width="45" height="20" border="0" onclick="fnDelForm()" style="cursor:pointer" align="absmiddle"></td>
    <td valign="bottom" align="right">
		<img src="http://testwebadmin.10x10.co.kr/images/icon_cancel.gif" width="45" height="20" border="0" onclick="fnCancel()" style="cursor:pointer" align="absmiddle"> &nbsp;&nbsp;&nbsp;
		<img src="http://testwebadmin.10x10.co.kr/images/icon_save.gif" width="45" height="20" border="0" onclick="fnSaveForm()" style="cursor:pointer" align="absmiddle">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td colspan="2" background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 하단바 끝-->
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
</p>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
