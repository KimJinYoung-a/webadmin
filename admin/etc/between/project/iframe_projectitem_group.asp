<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/projectcls.asp"-->
<%
Dim pjt_code : pjt_code = Request("pjt_code")
Dim sType : sType = Request("T")
Dim cPjtGroup, arrList,intg, i

SET cPjtGroup = new cProject
	cPjtGroup.FRectPjt_code = pjt_code
	cPjtGroup.getProjectItemGroup()
%>
<script language="javascript" defer>
function jsAddGroup(pjt_code, gCode){
	var winG
	winG = window.open('pop_projectitem_group.asp?pjt_code='+pjt_code+'&pjtgroup_code='+gCode,'popG','width=600, height=500');
	winG.focus();
}
function jsDelGroup(pjt_code,gCode){
	if(confirm("삭제시 하위그룹들 모두 삭제됩니다. 삭제하시겠습니까? ")){
		document.frmD.pGC.value = gCode;
		document.frmD.submit();
	}
}
</script>
<% IF sType="1" THEN %><body onunload="opener.location.href='project_regist.asp?pjt_code=<%=pjt_code%>';"><% END IF %>
<form name="frmD" method="post" action="project_process.asp">
<input type="hidden" name="pjt_code" value="<%=pjt_code%>">
<input type="hidden" name="pGC" value="">
<input type="hidden" name="mode" value="GD">
</form>
<table width="650" border="0" cellpadding="1" cellspacing="3" class="a">
<tr>
	<td>
		<input type="button" value="그룹추가" onClick="jsAddGroup('<%=pjt_code%>','');" class="input_b">
	</td>
</tr>
<tr>
	<td>
		<% If cPjtGroup.FResultCount > 0 Then %>
			<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">그룹코드</td>
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">상위그룹</td>
				<td bgcolor="<%= adminColor("tabletop") %>">그룹명</td>
				<td bgcolor="<%= adminColor("tabletop") %>">실제표시BG/FontColor</td>
				<td width="50" bgcolor="<%= adminColor("tabletop") %>">정렬순서</td>
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">관리</td>
			</tr>
			<% For i = 0 to cPjtGroup.FResultCount - 1 %>
			<tr>
				<td  align="center" bgcolor="#FFFFFF"><% IF cPjtGroup.FItemList(i).FPjtgroup_pcode <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%= cPjtGroup.FItemList(i).FPjtgroup_code %></td>
				<td  align="center" bgcolor="#FFFFFF"><% IF isnull(cPjtGroup.FItemList(i).FIstop)THEN%>최상위<%ELSE%>[<%=cPjtGroup.FItemList(i).FPjtgroup_pcode%>]<%=db2html(cPjtGroup.FItemList(i).FIstop)%><%END IF%></td>
				<td  align="center" bgcolor="#FFFFFF"><%= db2html(cPjtGroup.FItemList(i).FPjtgroup_desc)%></td>
				<td  align="center" bgcolor="<%= cPjtGroup.FItemList(i).FPjtgroup_BGColor %>">
					<font color="<%= cPjtGroup.FItemList(i).FPjtgroup_FontColor %>"><%= cPjtGroup.FItemList(i).FPjtgroup_FontColor %></font>
				</td>
				<td  align="center" bgcolor="#FFFFFF"><%= cPjtGroup.FItemList(i).FPjtgroup_sort %></td>
				<td  align="center" bgcolor="#FFFFFF">
					<input type="button" name="btnU" value="수정" onclick="jsAddGroup('<%=pjt_code%>','<%= cPjtGroup.FItemList(i).FPjtgroup_code %>')" class="button">
					<input type="button" name="btnD" value="삭제" onclick="jsDelGroup('<%=pjt_code%>','<%= cPjtGroup.FItemList(i).FPjtgroup_code %>')"  class="button">
				</td>
			</tr>
			<% Next %>
			</table>
		<% End If %>
	</td>
</tr>
</table>
<% SET cPjtGroup = nothing %>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->