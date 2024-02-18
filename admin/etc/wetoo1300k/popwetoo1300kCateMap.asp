<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/wetoo1300k/wetoo1300kcls.asp"-->
<%
Dim o1300k, i
Dim cdl, cdm, cds, large_category, middle_category, small_category, detail_category
cdl		= request("cdl")
cdm		= request("cdm")
cds		= request("cds")
large_category	= requestCheckVar(request("large_category"),10)
middle_category	= requestCheckVar(request("middle_category"),10)
small_category	= requestCheckVar(request("small_category"),10)
detail_category	= requestCheckVar(request("detail_category"),10)

If cdl = "" Then
	Call Alert_Close("카테고리 코드가 없습니다.")
	dbget.Close: Response.End
End IF

'// 카테고리 내용 접수
Set o1300k = new C1300k
	o1300k.FPageSize = 20
	o1300k.FCurrPage = 1
	o1300k.FRectCDL = cdl
	o1300k.FRectCDM = cdm
	o1300k.FRectCDS = cds
	o1300k.getTen1300kCateList

If o1300k.FResultCount <= 0 Then
	Call Alert_Close("해당 카테고리 정보가 없습니다.")
	dbget.Close: Response.End
End If
%>
<script language="javascript">
<!--
	// 매칭 저장하기
	function fnSaveForm() {
		var frm = document.frmAct;

		if(frm.large_category.value=="") {
			alert("매칭할 1300k 카테고리를 선택해주세요.");
			return;
		}

		if(confirm("선택하신 카테고리로 매칭하시겠습니까?")) {
			frm.mode.value="saveCate";
			frm.action="procwetoo1300k.asp";
			frm.submit();
		}
	}

    function fnDelForm(cdl, cdm, cds) {
		var frm = document.frmAct;
		if (cdl=="") {
		    alert("삭제할 1300k 카테고리를 선택해주세요.");
			return;
		}

		if(confirm("현재 매칭된 카테고리를 연결해제 하시겠습니까?\n\n※ 상품 또는 카테고리가 삭제되는 것은 아니며, 연결된 정보만 삭제됩니다.")) {
			frm.mode.value="delCate";
			frm.cdl.value=cdl;
			frm.cdm.value=cdm;
			frm.cds.value=cds;
			frm.action="procwetoo1300k.asp";
			frm.submit();
		}
	}

	// 창닫기
	function fnCancel() {
		if(confirm("작업을 취소하고 창을 닫으시겠습니까?")) {
			self.close();
		}
	}

	// 1300k 카테고리 검색
	function fnSearchwetoo1300kCate() {
		var kwd;
		kwd = document.getElementById("srcKwd").value;
		var pFCL = window.open("popFindwetoo1300kCate.asp?srcKwd="+kwd,"popwetoo1300kCate","width=1000,height=700,scrollbars=yes,resizable=yes");
		pFCL.focus();
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
	<font color="red"><strong>wetoo1300k 카테고리 매칭</strong></font></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 텐바이텐 카테고리 정보</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">대분류</td>
	<td bgcolor="#FFFFFF">[<%=cdl%>] <%=o1300k.FItemList(0).FtenCDLName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">중분류</td>
	<td bgcolor="#FFFFFF">[<%=cdm%>] <%=o1300k.FItemList(0).FtenCDMName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">소분류</td>
	<td bgcolor="#FFFFFF">[<%=cds%>] <%=o1300k.FItemList(0).FtenCDSName%></td>
</tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> wetoo1300k 카테고리 매칭 정보</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<form name="srcFrm" method="GET" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="srcDiv" value="CNM">
<input type="hidden" name="disptpcd" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="2" >검색</td>
	<td bgcolor="#FFFFFF">
		카테고리명 <input type="text" id="srcKwd" name="srcKwd" class="text">
		<input type="button" value="검색" class="button" onClick="fnSearchwetoo1300kCate();">
	</td>
</tr>
<tr id="BrRow" style="display:">
	<td bgcolor="#F2F2F2">추가 : <b><span id="selBr"></span></b></td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="<%= o1300k.FResultCount + 1 %>" >등록된<br>카테고리</td>
	<td bgcolor="#FFFFFF" height="1"></td>
</tr>
<% For i = 0 to o1300k.FResultCount - 1 %>
<% If Not IsNULL(o1300k.FItemList(i).FLarge_category) Then %>
<tr>
    <td bgcolor="#F2F2F2"><b><span id="selBr"><%=o1300k.FItemList(i).FCategory_name%> [<%= o1300k.FItemList(i).FLarge_category & o1300k.FItemList(i).FMiddle_category & o1300k.FItemList(i).FSmall_category & o1300k.FItemList(i).FDetail_category %>]</span></b>
    &nbsp;&nbsp;&nbsp;&nbsp;<img src="/images/icon_delete.gif" width="45" height="20" border="0" onclick="fnDelForm('<%= cdl %>', '<%= cdm %>', '<%= cds %>')" style="cursor:pointer" align="absmiddle">
    </td>
</tr>
<% End If %>
<% Next %>
</table>
</form>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="28">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left"></td>
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
<form name="frmAct" method="POST" target="xLink" style="margin:0px;">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="cdm" value="<%=cdm%>">
<input type="hidden" name="cds" value="<%=cds%>">
<input type="hidden" name="large_category" value="<%= large_category %>">
<input type="hidden" name="middle_category" value="<%= middle_category %>">
<input type="hidden" name="small_category" value="<%= small_category %>">
<input type="hidden" name="detail_category" value="<%= detail_category %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="categbn" value="cate">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="110" height="110"></iframe>
</p>
<% Set o1300k = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
