<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/homeplus/homepluscls.asp"-->
<%
Dim oHomeplus, i, page
Dim cdd_NAME, infodiv, schCateID

page		= request("page")
cdd_NAME	= request("cdd_NAME")
infodiv		= request("infodiv")

Select Case infodiv
	Case "01"	schCateID = "129"
	Case "02"	schCateID = "134"
	Case "03"	schCateID = "133"
	Case "04"	schCateID = "130"
	Case "05"	schCateID = "113"
	Case "06"	schCateID = "101"
	Case "07"	schCateID = "102"
	Case "08"	schCateID = "107"
	Case "09"	schCateID = "108"
	Case "10"	schCateID = "103"
	Case "11"	schCateID = "104"
	Case "12"	schCateID = "105"
	Case "14"	schCateID = "106"
	Case "15"	schCateID = "135"
	Case "16"	schCateID = "131"
	Case "17"	schCateID = "112"
	Case "18"	schCateID = "125"
	Case "19"	schCateID = "132"
	Case "20"	schCateID = "114"
	Case "21"	schCateID = "115"
	Case "23"	schCateID = "116"
	Case "25"	schCateID = "111"
	Case "26"	schCateID = "118"
	Case "31"	schCateID = "110"
	Case "35"	schCateID = "126"
End Select

If page = ""	Then page = 1
'// 목록 접수
Set oHomeplus = new CHomeplus
	oHomeplus.FPageSize = 5000
	oHomeplus.FCurrPage = page
	oHomeplus.FsearchName = cdd_NAME
	oHomeplus.FsearchCateId = schCateID
	oHomeplus.getHomeplusPrdDivList
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}

	// 상품분류 선택
	function fnSelPrddiv(divsioncode, groupcode, deptcode, classcode, subclasscode, categoryid, subclassnm) {
		opener.document.frmAct.divsioncode.value=divsioncode;
		opener.document.frmAct.groupcode.value=groupcode;
		opener.document.frmAct.deptcode.value=deptcode;
		opener.document.frmAct.classcode.value=classcode;
		opener.document.frmAct.subclasscode.value=subclasscode;
		opener.document.frmAct.categoryid.value=categoryid;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML= subclassnm;
		self.close();
	}
//-->
</script>
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="cdd_NAME" value="<%=cdd_NAME%>">
</form>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>Homeplus 기준카테고리 검색</strong></font></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> 검색결과 : <strong><%=oHomeplus.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>DIVISION_CODE</td>
	<td>DIVISION명</td>
	<td>GROUP_CODE</td>
	<td>GROUP명</td>
	<td>DEPT_CODE</td>
	<td>DEPT명</td>
	<td>CLASS_CODE</td>
	<td>CLASS명</td>
	<td>SUBCLASS_CODE</td>
	<td>SUBCLASS명</td>
	<td>정보고시_ID</td>
	<td>정보고시명</td>
</tr>
<% If oHomeplus.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
	Else
		For i = 0 to oHomeplus.FresultCount - 1
%>
<tr align="center" height="25" onClick="fnSelPrddiv('<%= oHomeplus.FItemList(i).FhDIVISION %>', '<%= oHomeplus.FItemList(i).FhGROUP %>', '<%= oHomeplus.FItemList(i).FhDEPT %>', '<%= oHomeplus.FItemList(i).FhCLASS %>', '<%= oHomeplus.FItemList(i).FhSUBCLASS %>', '<%= oHomeplus.FItemList(i).FhCATEGORY_ID %>', '<%= oHomeplus.FItemList(i).FhSUB_NAME %>')" style="cursor:pointer" title="카테고리 선택" bgcolor="#FFFFFF">
	<td><%= oHomeplus.FItemList(i).FhDIVISION %></td>
	<td><%= oHomeplus.FItemList(i).FhDIV_NAME %></td>
	<td><%= oHomeplus.FItemList(i).FhGROUP %></td>
	<td><%= oHomeplus.FItemList(i).FhGROUP_NAME %></td>
	<td><%= oHomeplus.FItemList(i).FhDEPT %></td>
	<td><%= oHomeplus.FItemList(i).FhDEPT_NAME %></td>
	<td><%= oHomeplus.FItemList(i).FhCLASS %></td>
	<td><%= oHomeplus.FItemList(i).FhCLASS_NAME %></td>
	<td><%= oHomeplus.FItemList(i).FhSUBCLASS %></td>
	<td><%= oHomeplus.FItemList(i).FhSUB_NAME %></td>
	<td><%= oHomeplus.FItemList(i).FhCATEGORY_ID %></td>
	<td><%= oHomeplus.FItemList(i).FhCATEGORY_NAME %></td>
</tr>
<%
		Next
	End If
%>
</table>
<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% If oHomeplus.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oHomeplus.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>

		<% For i = 0 + oHomeplus.StartScrollPage to oHomeplus.FScrollCount + oHomeplus.StartScrollPage - 1 %>
			<% If i>oHomeplus.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<foNt color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% next %>

		<% If oHomeplus.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% Else %>
			[next]
		<% End If %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 하단바 끝-->
<iframe name="xLink" id="xLink" frameborder="1" width="10" height="10"></iframe>
<% Set oHomeplus = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
