<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ��ġ����ĿApp ����Ʈ
' History : 2013.02.28 ������ ����
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/appmanage/hitchhiker.asp" -->
<%
Dim g_MenuPos
IF application("Svr_Info")="Dev" THEN
	g_MenuPos   = "1561"		'### �޴���ȣ ����.
Else
	g_MenuPos   = "1304"		'### �޴���ȣ ����.
End If

Dim hitch, page, i
Dim search1, searchVol
search1 = request("search1")
searchVol = request("searchVol")
page = request("page")

If page = "" Then page = 1

Set hitch = new Hitchhiker
	hitch.FPageSize = 10
	hitch.Fsearch1 = search1
	hitch.FsearchVol = searchVol
	hitch.FCurrPage = page
	hitch.HitchList
%>
<script language="javascript">
function jszip(mode,rev,vol){
	var winImg;
	winImg = window.open('/admin/appmanage/pop_hitch_zip.asp?mode='+mode+'&rev='+rev+'&vol='+vol+'','popImg','width=650,height=250, status=yes');
	winImg.focus();
}
function HitchModi(idx, vol){
	var winImg2;
	winImg2 = window.open('/admin/appmanage/pop_hitch_modify.asp?idx='+idx+'&vol='+vol+'','popImg2','width=650,height=200, status=yes');
	winImg2.focus();
}
function hitchNotice(){
	var winImg3;
	winImg3 = window.open('/admin/appmanage/pop_hitch_NoticeList.asp','popImg3','width=850,height=700, status=yes');
	winImg3.focus();
}
function CommonBanner(){
	var winImg4;
	winImg4 = window.open('/admin/appmanage/pop_CommonBannerList.asp','popImg3','width=950,height=700, status=yes');
	winImg4.focus();
}
function gosubmit(page){
    frm.page.value=page;
	frm.submit();
}
function seach_check(){
	var fsearch = document.fsearch;
	fsearch.submit();
}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fsearch" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		Vol : <input type="text" name="searchVol" value="<%=searchVol%>" size="5" maxlength="3">&nbsp;&nbsp;
		<select name = "search1" class="select">
			<option value="all" <%=chkiif(search1="all","selected","")%> >--��ü--
			<option value="lastRev" <%=chkiif(search1="lastRev","selected","")%> >����Rev
			<option value="open" <%=chkiif(search1="open","selected","")%> >����(��������)
		</select>&nbsp;&nbsp;
		<img src="/admin/images/search2.gif" border="0" align="absmiddle" style="cursor:hand" onclick="seach_check()">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%=g_MenuPos%>">
<input type="hidden" name="page">
<tr>
	<td align="left">
		<input type = "button" value="�ű�Vol���ε�" class="button" onclick="javascript:jszip('I','','');">
	</td>
	<td align="right">
		<input type = "button" value="������" class="button" onclick="javascript:CommonBanner();">
		<input type = "button" value="��������" class="button" onclick="javascript:hitchNotice();">
	</td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30"><td><img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>��ġ����Ŀ ����Ʈ</b></td></tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="30">Vol</td>
	<td width="40">Rev</td>
	<td>����</td>
	<td width="120">�����̹���1</td>
	<td width="120">�����̹���2</td>
	<td>�������ϰ��</td>
	<td width="70">�����ID</td>
	<td>�����</td>
	<td width="70">������</td>
	<td>����</td>
	<td>����</td>
	<td>�� Rev���</td>
	<td>�󼼵��</td>
</tr>
<%
	For i = 0 to hitch.FResultCount -1
%>
<tr height="25" bgcolor="FFFFFF">
	<td align="center"><%=hitch.FhitchList(i).Fvol%></td>
	<td align="center"><%=hitch.FhitchList(i).Frev%></td>
	<td align="center"><%=hitch.FhitchList(i).FmTitleName%></td>
	<td align="center"><img src="<%=hitch.FhitchList(i).FmImgURL%>" width="100" height="120"></td>
	<td align="center"><img src="<%=hitch.FhitchList(i).FmImgURL2%>" width="100" height="120"></td>
	<td align="center"><a href="<%=hitch.FhitchList(i).FzipUrl%>"><%=hitch.FhitchList(i).FzipUrl%></a></td>
	<td align="center"><%=hitch.FhitchList(i).FregUserID%></td>
	<td align="center"><%=hitch.FhitchList(i).Fregdate%></td>
	<td align="center"><%=hitch.FhitchList(i).Fopendate%></td>
	<td align="center">
	<%
		Select Case	hitch.FhitchList(i).FopenState
			Case "0"	response.write "������"
			Case "3"	response.write "DevOpen"
			Case "9"	response.write "����"
		End Select

		If hitch.FhitchList(i).FopenState = "7" AND DateDiff("d", Now, hitch.FhitchList(i).Fopendate) > 0 Then
			response.write "���±���</br><font color='RED'><strong>"&DateDiff("d", Now, hitch.FhitchList(i).Fopendate)&"</strong></font>�� ����"
		ElseIf hitch.FhitchList(i).FopenState = "7" AND DateDiff("d", Now, hitch.FhitchList(i).Fopendate) <= 0 Then
			response.write "<font color='BLUE'><strong>����</strong></font>"
		End If
	%>
	</td>
	<td align="center">
		<%
		' �ָ��Ƴ��?? 		2018.05.08 �ѿ��
		'If (hitch.FhitchList(i).FopenState <> "7") OR (hitch.FhitchList(i).FopenState = "7" AND DateDiff("d", Now, hitch.FhitchList(i).Fopendate) > 0) Then 
		%>
			<input type="button" class="button" value="����" onclick="javascript:HitchModi('<%=hitch.FhitchList(i).Fidx%>','<%=hitch.FhitchList(i).Fvol%>');">
		<% 'Else %>
			<!--�����Ұ�-->
		<% 'End If %>
	</td>
	<td align="center"><input type="button" class="button" value="���" onclick="javascript:jszip('R','<%=hitch.FhitchList(i).Frev%>','<%=hitch.FhitchList(i).Fvol%>');"></td>
	<td align="center"><input type="button" class="button" value="���" onclick="location.href='hitchDetail.asp?midx=<%=hitch.FhitchList(i).Fidx%>&vol=<%=hitch.FhitchList(i).Fvol%>&rev=<%=hitch.FhitchList(i).Frev%>';"></td>
</tr>
<%
	Next
%>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If hitch.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= hitch.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + hitch.StartScrollPage to hitch.StartScrollPage + hitch.FScrollCount - 1 %>
			<% If (i > hitch.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(hitch.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If hitch.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
</form>
</table>
<% Set hitch = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->