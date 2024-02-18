<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/admin/etc/between/projectcls.asp"-->
<script language="javascript">
 function jsGroupSubmit(frm){
 	if(!frm.pjtgroup_desc.value){
	 	alert("그룹명을 입력해주세요");
	 	return false;
 	}
 }
</script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
</script>
<%
Dim pjt_code : pjt_code = Request("pjt_code")
Dim pjtgroup_code : pjtgroup_code = Request("pjtgroup_code")
Dim cPGroup, arrP, intP, sM
Dim pjtgroup_desc, pjtgroup_sort, pjtgroup_depth, pjtgroup_pdesc, pjtgroup_pcode, pjtgroup_BGColor, pjtgroup_FontColor
SET cPGroup = new cProject
	cPGroup.FRectpjt_code = pjt_code
	arrP = cPGroup.fnGetRootGroup()
	sM = "GI"
	If (pjtgroup_code <> "" and pjtgroup_code <> "0" and not isnull(pjtgroup_code)) Then
		cPGroup.FRectPjtgroup_code = pjtgroup_code
		cPGroup.GetProjectItemGroupCont
	  	pjtgroup_code	= cPGroup.FItemList(0).FPjtgroup_code
	  	pjtgroup_desc  	= cPGroup.FItemList(0).FPjtgroup_desc
	  	pjtgroup_sort	= cPGroup.FItemList(0).FPjtgroup_sort
	  	pjtgroup_pcode	= cPGroup.FItemList(0).FPjtgroup_pcode
	  	pjtgroup_depth	= cPGroup.FItemList(0).FPjtgroup_depth
	  	pjtgroup_pdesc	= cPGroup.FItemList(0).FPjtgroup_pdesc
	  	pjtgroup_BGColor= cPGroup.FItemList(0).FPjtgroup_BGColor
	  	pjtgroup_FontColor= cPGroup.FItemList(0).FPjtgroup_FontColor
		sM = "GU"
	End If

	If pjtgroup_sort = "" Then pjtgroup_sort = 0
%>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 이벤트 그룹 등록</div>
<table width="580" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<form name="frmG" method="post" action="project_process.asp" onSubmit="return jsGroupSubmit(this);">
<input type="hidden" name="pjt_code" value="<%=pjt_code%>">
<input type="hidden" name="pjtgroup_code" value="<%=pjtgroup_code%>">
<input type="hidden" name="mode" value="<%=sM%>">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">상위그룹</td>
			<td bgcolor="#FFFFFF">
			<% If pjtgroup_depth = "" Then %>
				<select name="selPC">
					<option value="0">최상위</option>
			<%
				If isArray(arrP) Then
					For intP =0 To UBound(arrP,2)
			%>
					<option value="<%=arrP(0,intP)%>" <%= Chkiif(Cstr(pjtgroup_code) = CStr(arrP(0,intP)), "selected", "") %>><%= arrP(1,intP) %></option>
			<%
					Next
				End If
			%>
				</select>
			<% Else %>
				<input type="hidden" name="selPC" value="<%=pjtgroup_pcode%>">
			<%= pjtgroup_pdesc %>
			<% End If %>
			</td>
		</tr>
		<tr>
			<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">그룹명</td>
			<td bgcolor="#FFFFFF"><input type="text" name="pjtgroup_desc" size="20" value="<%=db2html(pjtgroup_desc)%>"></td>
		</tr>
		<tr>
			<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">그룹배경색</td>
			<td bgcolor="#FFFFFF"><input type="text" name="pjtgroup_BGColor" style="width:80px;" maxlength="7" value="<%=db2html(pjtgroup_BGColor)%>"></td>
		</tr>
		<tr>
			<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">그룹폰트색</td>
			<td bgcolor="#FFFFFF"><input type="text" name="pjtgroup_FontColor" style="width:80px;" maxlength="7" value="<%=db2html(pjtgroup_FontColor)%>"></td>
		</tr>
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">정렬순서</td>
			<td bgcolor="#FFFFFF"><input type="text" size="2" name="pjtgroup_sort"  value="<%=pjtgroup_sort%>"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right" height="40">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>
<% set cPGroup = nothing %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->