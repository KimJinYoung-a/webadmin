<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��Ƽ��Ʈ ���� ������ ����Ʈ
' History : 2012.03.29 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/artist/artist_class.asp"-->
<%
'// ���� ����
Dim page, isusing, designerid, i, mm
	mm = request("mm")
	page = request("page")
	isusing = request("isusing")
	designerid = request("designerid")
	
	if page="" then page=1

'// ��� ����
Dim oGallery
	set oGallery = New cposcode_list
	oGallery.FCurrPage = page
	oGallery.FPageSize=10
	oGallery.FArtistMonthItemList

%>
<script>
function goView(ii){
	location.href = "artist_MonthItemWrite.asp?mode=edit&idx="+ii;
}
function gosubmit(page){
    frm.page.value=page;
	frm.submit();
}
function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}
function AssignReal(upfrm,imagecount){
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				upfrm.fidx.value = upfrm.fidx.value + frm.idx.value + "," ;
			}
		}
	}
	var tot;
	tot = upfrm.fidx.value;
	upfrm.fidx.value = ""

	var AssignimageReal;
	AssignimageReal = window.open("<%=wwwUrl%>/chtml/make_artist_shop2item.asp?idx=" +tot + '&imagecount='+imagecount, "AssignimageReal","width=800,height=600,scrollbars=yes,resizable=yes");
	AssignimageReal.focus();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="40">��ġ : 
		<select onchange="location.href=this.value;" class="select">
			<option value="artist_weekmonth.asp?menupos=<%=menupos%>&mm=1">��Ƽ��Ʈ �� ��ܹ��
			<option value="artist_MonthItemList.asp?menupos=<%=menupos%>&mm=2" <% If mm = 2 Then response.write "selected"%>>��Ƽ��Ʈ �� ���� ��ǰ
		</select>		
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="fidx">
<tr><td align="left"><a href="javascript:AssignReal(frm,2);"><img src="/images/refreshcpage.gif" border="0"> Real ����</a></td></tr>
<tr><td align="left"><input type="button" class="button" value="���" onclick="javascript:location.href='artist_MonthItemWrite.asp';"></td></tr>
</form>
</table>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30"><td><img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>��Ƽ��Ʈ �� ��ǰ ����Ʈ</b></td></tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td width="50">��ȣ</td>
	<td width="70">��ǰ�ڵ�</td>
	<td width="60">�̹���</td>
	<td>�ڸ�Ʈ</td>
	<td width="60">����</td>
	<td width="60">���</td>
	<td width="150">�����</td>
</tr>

<% If oGallery.FTotalCount = 0 Then %>
<tr height="25" bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'">
	<td align="center" colspan="6">[�����Ͱ� �����ϴ�.]</td>
</tr>
<% End If %>

<% For i=0 to oGallery.FResultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">
<input type="hidden" name="idx" value="<%=oGallery.FItemList(i).fidx%>">
<tr height="25" bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" >
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>		
	<td align="center" width="50"><%=oGallery.FItemList(i).fidx%></td>
	<td align="center" width="70"><%=oGallery.FItemList(i).fitemid%></td>
	<td align="center" width="60"><img src="<%=oGallery.FItemList(i).ficon2image%>"></td>
	<td onClick="goView('<%=oGallery.FItemList(i).fidx%>')" style="cursor:pointer" ><%=db2html(oGallery.FItemList(i).fcomment)%></td>
	<td align="center" width="60"><%=oGallery.FItemList(i).fsortNo%></td>
	<td align="center" width="60"><%=oGallery.FItemList(i).fisusing%></td>
	<td align="center" width="150"><%=oGallery.FItemList(i).fregdate%></td>
</tr>
</form>	
<% Next %>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="8" align="center">
       	<% If oGallery.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= ohistory.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + oGallery.StartScrollPage to oGallery.StartScrollPage + oGallery.FScrollCount - 1 %>
			<% If (i > oGallery.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(oGallery.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If oGallery.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->