<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/skstoa/skstoaCls.asp"-->
<%
Dim oSkstoa, i
Dim cdl, cdm, cds, lgroup, mgroup, sgroup, dgroup
cdl		= requestCheckVar(request("cdl"),3)
cdm		= requestCheckVar(request("cdm"),3)
cds		= requestCheckVar(request("cds"),3)
lgroup	= requestCheckVar(request("lgroup"),10)
mgroup	= requestCheckVar(request("mgroup"),10)
sgroup	= requestCheckVar(request("sgroup"),10)
dgroup	= requestCheckVar(request("dgroup"),10)

If cdl = "" Then
	Call Alert_Close("카테고리 코드가 없습니다.")
	dbget.Close: Response.End
End IF

'// 카테고리 내용 접수
Set oSkstoa = new CSkstoa
	oSkstoa.FPageSize = 20
	oSkstoa.FCurrPage = 1
	oSkstoa.FRectCDL = cdl
	oSkstoa.FRectCDM = cdm
	oSkstoa.FRectCDS = cds
	oSkstoa.getTenSkstoaCateList

If oSkstoa.FResultCount <= 0 Then
	Call Alert_Close("해당 카테고리 정보가 없습니다.")
	dbget.Close: Response.End
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
<!--
	// 매칭 저장하기
	function fnSaveForm() {
		var frm = document.frmAct;

		if(frm.lgroup.value=="") {
			alert("매칭할 Skstoa 카테고리를 선택해주세요.");
			return;
		}

		if(confirm("선택하신 카테고리로 매칭하시겠습니까?")) {
			frm.mode.value="saveCate";
			frm.action="procskstoa.asp";
			frm.submit();
		}
	}

    function fnDelForm(cdl, cdm, cds) {
		var frm = document.frmAct;
		if (cdl=="") {
		    alert("삭제할 Skstoa 카테고리를 선택해주세요.");
			return;
		}

		if(confirm("현재 매칭된 카테고리를 연결해제 하시겠습니까?\n\n※ 상품 또는 카테고리가 삭제되는 것은 아니며, 연결된 정보만 삭제됩니다.")) {
			frm.mode.value="delCate";
			frm.cdl.value=cdl;
			frm.cdm.value=cdm;
			frm.cds.value=cds;
			frm.action="procskstoa.asp";
			frm.submit();
		}
	}

	// 창닫기
	function fnCancel() {
		if(confirm("작업을 취소하고 창을 닫으시겠습니까?")) {
			self.close();
		}
	}

	// Skstoa 카테고리 검색
	function fnSearchSkstoaCate() {
		var kwd;
		kwd = document.getElementById("srcKwd").value;
		var pFCL = window.open("popFindskstoaCate.asp?srcKwd="+kwd,"popSkstoaCate","width=1000,height=700,scrollbars=yes,resizable=yes");
		pFCL.focus();
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr valign="top">
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>Skstoa 카테고리 매칭</strong></font></td>
</tr>
</table>
<p>
<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 텐바이텐 카테고리 정보</td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">대분류</td>
	<td bgcolor="#FFFFFF">[<%=cdl%>] <%=oSkstoa.FItemList(0).FtenCDLName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">중분류</td>
	<td bgcolor="#FFFFFF">[<%=cdm%>] <%=oSkstoa.FItemList(0).FtenCDMName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">소분류</td>
	<td bgcolor="#FFFFFF">[<%=cds%>] <%=oSkstoa.FItemList(0).FtenCDSName%></td>
</tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> Skstoa 카테고리 매칭 정보</td>
</tr>
</table>
<!-- 표 중간바 끝-->
<form name="srcFrm" method="GET" onsubmit="fnSearchSkstoaCate();return false;" style="margin:0px;">
<input type="hidden" name="srcDiv" value="CNM">
<input type="hidden" name="disptpcd" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="2" >검색</td>
	<td bgcolor="#FFFFFF">
		카테고리명 <input type="text" id="srcKwd" name="srcKwd" class="text">
		<input type="button" value="검색" class="button" onClick="fnSearchSkstoaCate();">
	</td>
</tr>
<tr id="BrRow" style="display:">
	<td bgcolor="#F2F2F2">추가 : <b><span id="selBr"></span></b></td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="<%= oSkstoa.FResultCount + 1 %>" >등록된<br>카테고리</td>
	<td bgcolor="#FFFFFF" height="1"></td>
</tr>
<% For i = 0 to oSkstoa.FResultCount - 1 %>
<% If Not IsNULL(oSkstoa.FItemList(i).FLgroup) Then %>
<tr>
    <td bgcolor="#F2F2F2"><b><span id="selBr"><%=oSkstoa.FItemList(i).FLastDepthNm%> [<%= oSkstoa.FItemList(i).FLgroup & oSkstoa.FItemList(i).FMgroup & oSkstoa.FItemList(i).FSgroup & oSkstoa.FItemList(i).FDgroup %>]</span></b>
    &nbsp;&nbsp;&nbsp;&nbsp;
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
		<img src="/images/icon_cancel.gif" width="45" height="20" border="0" onclick="fnCancel()" style="cursor:pointer" align="absmiddle"> &nbsp;&nbsp;&nbsp;
		<img src="/images/icon_save.gif" width="45" height="20" border="0" onclick="fnSaveForm()" style="cursor:pointer" align="absmiddle"> &nbsp;&nbsp;&nbsp;
		<% If lgroup <> "" Then %>
		<img src="/images/icon_delete.gif" width="45" height="20" border="0" onclick="fnDelForm('<%= cdl %>', '<%= cdm %>', '<%= cds %>');" style="cursor:pointer" align="absmiddle">
		<% End If %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 하단바 끝-->
<form name="frmAct" method="POST" style="margin:0px;">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="cdm" value="<%=cdm%>">
<input type="hidden" name="cds" value="<%=cds%>">
<input type="hidden" name="lgroup" value="<%= lgroup %>">
<input type="hidden" name="mgroup" value="<%= mgroup %>">
<input type="hidden" name="sgroup" value="<%= sgroup %>">
<input type="hidden" name="dgroup" value="<%= dgroup %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="categbn" value="cate">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="1110" height="110"></iframe>
</p>
<% Set oSkstoa = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
