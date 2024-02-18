<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����������
' History : 2014.04.01 ������ ����
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
Dim o3ban, isusing, gender

page	= request("page")
isusing	= request("isusing")
gender	= request("gender")

If page = "" Then page=1
SET o3ban = new cMain
	o3ban.FPageSize		= 20
	o3ban.FCurrPage		= page
	o3ban.FRectIsusing	= isusing
	o3ban.FRectGender	= gender
	o3ban.Get3BannerList()
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
	location.href = "3ban_insert.asp?menupos=<%=menupos%>&idx="+v;
}

function RefreshCaFavKeyWordRec(term){
	if(confirm("���� 3Banner�� �����Ͻðڽ��ϱ�?")) {
			var popwin = window.open('','refreshFrm_main','');
			popwin.focus();
			refreshFrm.target = "refreshFrm_main";
			refreshFrm.action = "<%=mobileUrl%>/chtml/between/make_3banner_xml.asp?term=" + term;
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
	<td>������ �����Ͽ� <input type="text" name="vTerm" value="1" size="1" class="text" style="text-align:right;">�ϰ�<a href="javascript:RefreshCaFavKeyWordRec(document.all.vTerm.value);"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html�����"></a>XML Real ����(����)</a></td>
	<% Else %>
	<td>&nbsp;</td>
	<% End If %>
    <td align="right">
		<!-- �űԵ�� -->
    	<a href="3ban_insert.asp?menupos=<%=menupos%>&prevDate="><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!--  ����Ʈ -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="10">
		�� ��ϼ� : <b><%=o3ban.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=o3ban.FtotalPage%></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
    <td width="10%">������ <br/>real ����ð�</td>
    <td width="5%">����</td>
	<td width="15%">����̹���</td>
    <td width="15%">������/������</td>
    <td width="10%">�����</td>
    <td width="10%">�����</td>
    <td width="10%">����������</td>
    <td width="10%">���Ĺ�ȣ</td>	
    <td width="10%">��뿩��</td>
</tr>
<% 
	for i=0 to o3ban.FResultCount-1 
%>
<tr  height="30" align="center" bgcolor="<%=chkIIF(o3ban.FItemList(i).Fisusing="Y","#FFFFFF","#F0F0F0")%>">
    <td onclick="jsmodify('<%=o3ban.FItemList(i).Fidx%>');" style="cursor:pointer;"><%=o3ban.FItemList(i).Fidx%></td>
	<td>
		<%
			If o3ban.FItemList(i).Fxmlregdate <> "" then
			Response.Write replace(left(o3ban.FItemList(i).Fxmlregdate,10),"-",".") & " <br/> " & Num2Str(hour(o3ban.FItemList(i).Fxmlregdate),2,"0","R") & ":" &Num2Str(minute(o3ban.FItemList(i).Fxmlregdate),2,"0","R")
			End If 
		%>
	</td>
	<td>
	<%
		If o3ban.FItemList(i).FGender = "M" Then
			response.write "<font Color='BLUE'>����</font>"
		Else
			response.write "<font Color='PINK'>����</font>"
		End If
	%>
	</td>
    <td><img src="<%=o3ban.FItemList(i).FImgurl%>" width="100" /></td>
	<td>
		<% 
			Response.Write "����: "
			Response.Write replace(left(o3ban.FItemList(i).Fstartdate,10),"-",".") & " / " & Num2Str(hour(o3ban.FItemList(i).Fstartdate),2,"0","R") & ":" &Num2Str(minute(o3ban.FItemList(i).Fstartdate),2,"0","R")
			Response.Write "<br />����: "
			Response.Write replace(left(o3ban.FItemList(i).Fenddate,10),"-",".") & " / " & Num2Str(hour(o3ban.FItemList(i).Fenddate),2,"0","R") & ":" & Num2Str(minute(o3ban.FItemList(i).Fenddate),2,"0","R")
		%>
	</td>
	<td><%=left(o3ban.FItemList(i).Fregdate,10)%></td>
	<td><%=getStaffUserName(o3ban.FItemList(i).Fadminid)%></td>
	<td>
		<%
			if Not(o3ban.FItemList(i).Flastupdate="" or isNull(o3ban.FItemList(i).Flastupdate)) then
					Response.Write getStaffUserName(o3ban.FItemList(i).Flastadminid) & "<br />"
					Response.Write left(o3ban.FItemList(i).Flastupdate,10)
			end if
		%>
	</td>
	<td><%=o3ban.FItemList(i).Fsortno%></td>
    <td><%=chkiif(o3ban.FItemList(i).Fisusing="N","������","�����")%></td>
</tr>
<% Next %>
</table>
<!-- paging -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="20" bgcolor="#FFFFFF">
	<td colspan="18" align="center" bgcolor="#FFFFFF">
	    <% if o3ban.HasPreScroll then %>
		<a href="javascript:goPage('<%= o3ban.StartScrollPage-1 %>');">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for i=0 + o3ban.StartScrollPage to o3ban.FScrollCount + o3ban.StartScrollPage - 1 %>
			<% if i>o3ban.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>
	
		<% if o3ban.HasNextScroll then %>
			<a href="javascript:goPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<%
set o3ban = Nothing
%>
<form name="refreshFrm" method="post">
<input type="hidden" name="gender" value="<%= gender %>">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->