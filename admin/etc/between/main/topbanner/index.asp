<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����������
' History : 2014.04.18 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/between/mainCls.asp"-->
<!-- #include virtual="/admin/etc/between/main/inc_mainhead.asp"-->
<%
Dim page, i
Dim otopban, isusing, gender

page	= request("page")
isusing	= request("isusing")
gender	= request("gender")

If page = "" Then page=1
SET otopban = new cMain
	otopban.FPageSize		= 20
	otopban.FCurrPage		= page
	otopban.FRectIsusing	= isusing
	otopban.FRectGender		= gender
	otopban.getTopBannerList()
%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script type='text/javascript'>
<!--
//����
function jsmodify(v){
	location.href = "topban_insert.asp?menupos=<%=menupos%>&idx="+v;
}

function RefreshCaFavKeyWordRec(term){
	if(confirm("���� TopBanner �����Ͻðڽ��ϱ�?")) {
		var popwin = window.open('','refreshFrm','');
		popwin.focus();
		refreshFrm.target = "refreshFrm";
		refreshFrm.action = "<%=mobileUrl%>/chtml/between/make_topbanner_xml.asp"
		refreshFrm.submit();
	}
}
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
-->
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<div style="padding-bottom:10px;">
		* ���� :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<select name="gender" class="select">
			<option value="">-Choice-</option>
			<option value="M" <%= Chkiif(gender="M", "selected", "") %> >����</option>
			<option value="F" <%= Chkiif(gender="F", "selected", "") %> >����</option>
		</select>
		* ��뿩�� :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<select name="isusing" class="select">
			<option value="">-Choice-</option>
			<option value="Y" <%= Chkiif(isusing="Y", "selected", "") %> >Y</option>
			<option value="N" <%= Chkiif(isusing="N", "selected", "") %> >N</option>
		</select>
		</div>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:submit();">
	</td>
</tr>
</form>	
</table>
<!-- �˻� �� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<% If gender <> "" Then %>
	<td><a href="javascript:RefreshCaFavKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html�����"></a>XML Real ����</a></td>
	<% Else %>
	<td>&nbsp;</td>
	<% End If %>
    <td align="right">
		<!-- �űԵ�� -->
    	<a href="topban_insert.asp?menupos=<%=menupos%>&prevDate="><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		�� ��ϼ� : <b><%=otopban.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=otopban.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>idx</td>
    <td>������ <br/>real ����ð�</td>
    <td>����</td>
    <td>��ȹ������</td>
    <td>����̹���</td>
    <td>�����ɻ���</td>
    <td>�����</td>
    <td>�����</td>
    <td>����������</td>
    <td>��뿩��</td>
</tr>
<% 
	For i = 0 To otopban.FResultCount - 1
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(otopban.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=otopban.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=otopban.FItemList(i).Fidx%></td>
	<td>
		<%
			If otopban.FItemList(i).Fxmlregdate <> "" then
			Response.Write replace(left(otopban.FItemList(i).Fxmlregdate,10),"-",".") & " <br/> " & Num2Str(hour(otopban.FItemList(i).Fxmlregdate),2,"0","R") & ":" &Num2Str(minute(otopban.FItemList(i).Fxmlregdate),2,"0","R")
			End If 
		%>
	</td>
	<td>
	<%
		If otopban.FItemList(i).FGender = "M" Then
			response.write "<font Color='BLUE'>����</font>"
		Else
			response.write "<font Color='PINK'>����</font>"
		End If
	%>
	</td>
	<td><%= getDBcodeByName(otopban.FItemList(i).FPjt_kind) %></td>
	<td><img src="<%=otopban.FItemList(i).FImgurl%>" width="100" /></td>
	<td bgcolor="<%= otopban.FItemList(i).FBanBgColor %>">
		<font Color="<%= otopban.FItemList(i).FPartnerNmColor %>"><%= otopban.FItemList(i).FBanText1 %></font><br>
		<font Color="<%= otopban.FItemList(i).FBanTxtColor %>"><%= otopban.FItemList(i).FBanText2 %></font>&nbsp;
	</td>
	<td><%=left(otopban.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(otopban.FItemList(i).Fadminid)%></td>
	<td>
	<%
		If Not(otopban.FItemList(i).Flastupdate="" or isNull(otopban.FItemList(i).Flastupdate)) then
			Response.Write getStaffUserName(otopban.FItemList(i).Flastadminid) & "<br />"
			Response.Write left(otopban.FItemList(i).Flastupdate,10)
		End If
	%>
	</td>
    <td><%=chkiif(otopban.FItemList(i).Fisusing="N","������","�����")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="20" bgcolor="#FFFFFF">
	<td colspan="18" align="center" bgcolor="#FFFFFF">
	    <% if otopban.HasPreScroll then %>
		<a href="javascript:goPage('<%= otopban.StartScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for i=0 + otopban.StartScrollPage to otopban.FScrollCount + otopban.StartScrollPage - 1 %>
			<% if i>otopban.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>
	
		<% if otopban.HasNextScroll then %>
			<a href="javascript:goPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<%
set otopban = Nothing
%>
<form name="refreshFrm" method="post">
<input type="hidden" name="gender" value="<%= gender %>">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->