<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2011.01.11 ������ ����
'			   2022.07.04 �ѿ�� ����(isms�������������, �ҽ�ǥ��ȭ)
'	Description : QR�ڵ� ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/qrCodeCls.asp"-->
<%
dim page, QRDiv, CntYn, isusing, keyWd
	page		= requestCheckvar(getNumeric(request("page")),10)
	QRDiv		= request("QRDiv")
	CntYn		= request("CntYn")
	isusing		= requestCheckvar(request("isusing"),1)
	keyWd		= request("keyWd")
	
	if page="" then page=1
	if isusing="" then isusing="Y"

dim oQR
	set oQR = New CQRCode
	oQR.FCurrPage = page
	oQR.FPageSize=20
	oQR.FRectQRDiv = QRDiv
	oQR.FRectCntYn = CntYn
	oQR.FRectIsUsing = isusing
	oQR.FRectkeyWd = keyWd
	oQR.GetQRCode
dim i
%>
<script type='text/javascript'>
	document.domain = "10x10.co.kr";
	function popNewCode(){
		var popup_New = window.open("pop_QRCodeReg.asp", "popup_New", "width=800,height=600,scrollbars=yes,status=no");
		popup_New.focus();
	}

	function popModiCode(sn){
		var popup_New = window.open("pop_QRCodeReg.asp?qrSn="+sn, "popup_New", "width=800,height=600,scrollbars=yes,status=no");
		popup_New.focus();
	}

	function gotoPage(pg) {
		document.Listfrm.page.value=pg;
		document.Listfrm.submit();
	}
</script>	
<!-- �˻��� ���� -->
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		���м��� :
		<% DrawSelectBoxQRDiv "QRDiv", QRDiv %>&nbsp;/&nbsp;
		�α׻�� :
		<select name="CntYn" class="select">
			<option value=""  <% if CntYn="" then response.write "selected" %>>��ü</option>
			<option value="Y" <% if CntYn="Y" then response.write "selected" %>>���</option>
			<option value="N" <% if CntYn="N" then response.write "selected" %>>������</option>
		</select>&nbsp;/&nbsp;
		������� :
		<select name="isusing" class="select">
			<option value="A" <% if isusing="A" then response.write "selected" %>>��ü</option>
			<option value="Y" <% if isusing="Y" then response.write "selected" %>>���</option>
			<option value="N" <% if isusing="N" then response.write "selected" %>>������</option>
		</select>&nbsp;/&nbsp;
		���� :
		<input type="text" name="keyWd" size="25" class="text" value="<%=keyWd%>">
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right"><input type="button" value="���ڵ� �߰�" onclick="popNewCode()" class="button"></td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="9">&nbsp;�˻��� �ڵ�� : <%=oQR.FTotalCount%> ��</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center">��ȣ</td>
	<td align="center">QR�ڵ�</td>
	<td align="center">����</td>
	<td align="center">�ڵ��</td>
	<td align="center">�����</td>
	<td align="center">��뿩��</td>
	<td align="center">ī��Ʈ</td>
</tr>
<% for i=0 to oQR.FResultCount-1 %>
<tr bgcolor="<%=chkIIF(oQR.FItemList(i).FisUsing="Y","#FFFFFF","#E0E0E0")%>" onclick="popModiCode(<%=oQR.FItemList(i).FqrSn%>)" style="cursor:pointer">
	<td align="center"><%= oQR.FItemList(i).FqrSn %></td>
	<td align="center"><img src="<%= oQR.FItemList(i).FqrImage %>" width="50" height="50"></td>
	<td align="center">
	<%
		Select Case oQR.FItemList(i).FqrDiv
			Case 1
				response.write "URL"
			Case 2
				response.write "�ؽ�Ʈ"
			Case 3
				response.write "�̹���"
			Case 4
				response.write "������"
			Case 5
				response.write "APP URL"
		End Select
	%>
	</td>
	<td align="center"><%= ReplaceBracket(oQR.FItemList(i).FqrTitle) %></td>
	<td align="center"><%= left(oQR.FItemList(i).Fregdate,10) %></td>
	<td align="center"><%= oQR.FItemList(i).FisUsing %></td>
	<td align="center"><% if oQR.FItemList(i).FcountYn="Y" then Response.Write FormatNumber(oQR.FItemList(i).FqrHitCount,0) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="center">
	<% if oQR.HasPreScroll then %>
		<a href="javascript:gotoPage(<%= oQR.StarScrollPage-1 %>)">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oQR.StarScrollPage to oQR.FScrollCount + oQR.StarScrollPage - 1 %>
		<% if i>oQR.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:gotoPage(<%= i %>)">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oQR.HasNextScroll then %>
		<a href="javascript:gotoPage(<%= i %>)">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<% set oQR = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->