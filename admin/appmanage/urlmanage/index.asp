<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2014-08-18 ����ȭ ����
'	Description : app URL ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/appURLCls.asp"-->
<%
dim page, urldiv, isusing, keyWd
	page		= request("page")
	urldiv		= request("urldiv")
	isusing		= request("isusing")
	keyWd		= request("keyWd")
	
	if page="" then page=1
	if isusing="" then isusing="Y"

dim oAppurl
	set oAppurl = New APPURL
	oAppurl.FCurrPage = page
	oAppurl.FPageSize = 20
	oAppurl.FRectkeyWd = keyWd
	oAppurl.FRecturldiv = urldiv
	oAppurl.getappurl
dim i
%>
<script language='javascript'>
	document.domain = "10x10.co.kr";
	function popNewCode(){
		var popup_New = window.open("pop_URLReg.asp", "popup_New", "width=800,height=300,scrollbars=yes,status=no");
		popup_New.focus();
	}

	function popModiCode(sn){
		var popup_New = window.open("pop_URLReg.asp?idx="+sn, "popup_New", "width=800,height=300,scrollbars=yes,status=no");
		popup_New.focus();
	}

	function gotoPage(pg) {
		document.Listfrm.page.value=pg;
		document.Listfrm.submit();
	}

	function popQrNewCode(i,t,u){
		if(confirm("QR�ڵ带 �����Ͻðڽ��ϱ�?")) {
			var frm = document.frmReg;
			frm.appidx.value = i;
			frm.QRTitle.value = t;
			frm.QRContent.value = u;
			frm.target = "prociframe";
			frm.submit();
		}
	}

	function popReadCode(v){
		var popup_New = window.open("/admin/sitemaster/QRCode/pop_QRCodeReg.asp?qrSn="+v, "popup_New", "width=600,height=500,scrollbars=yes,status=no");
		popup_New.focus();
	}
</script>	
<!-- �˻��� ���� -->
<!-- qr�ڵ���� -->
<form name="frmReg" method="post" action="<%=staticUploadUrl%>/linkweb/mobile/captureQRcode_proc.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" value=""  name="QRTitle">
<input type="hidden" value=""  name="QRContent">
<input type="hidden" value="Y" name="countYN">
<input type="hidden" value="5" name="QRDiv"><!-- 5 APPURL���� -->
<input type="hidden" value="M" name="qrQuality">
<input type="hidden" value="Y" name="isUsing">
<input type="hidden" value=""  name="appidx">
</form>
<iframe name="prociframe" id="prociframe" frameborder="0" width="0" height="0" src="about:blank"></iframe>
<!-- qr�ڵ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="Listfrm" method="get" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		���м��� :
		<select name="urldiv" class="select">
			<option value="" <% if urldiv="" then response.write "selected" %>>��ü</option>
			<option value="1" <% if urldiv="1" then response.write "selected" %>>��ǰ��</option>
			<option value="2" <% if urldiv="2" then response.write "selected" %>>�̺�Ʈ</option>
			<option value="3" <% if urldiv="3" then response.write "selected" %>>�귣��</option>
			<option value="4" <% if urldiv="4" then response.write "selected" %>>ī�װ�</option>
			<option value="8" <% if urldiv="8" then response.write "selected" %>>�ܺ�URL</option>
			<option value="9" <% if urldiv="9" then response.write "selected" %>>Today</option>
			<option value="10" <% if urldiv="10" then response.write "selected" %>>����Ʈ</option>
			<option value="11" <% if urldiv="11" then response.write "selected" %>>��ٱ���</option>
		</select>&nbsp;/&nbsp;
		������� :
		<select name="isusing" class="select">
			<option value="" <% if isusing="A" then response.write "selected" %>>��ü</option>
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
</form>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right"><input type="button" value="URL �߰�" onclick="popNewCode()" class="button"></td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FAFAFA" height="22">
	<td colspan="9">&nbsp;�˻��� URL�� : <%=oAppurl.FTotalCount%> ��</td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center">��ȣ</td>
	<td align="center">����</td>
	<td align="center">����</td>
	<td align="center">APPURL��</td>
	<td align="center">�����</td>
	<td align="center">��뿩��</td>
	<td align="center">ī��Ʈ</td>
	<td align="center">qr�ڵ�</td>
</tr>
<% for i=0 to oAppurl.FResultCount-1 %>
<tr bgcolor="<%=chkIIF(oAppurl.FItemList(i).FisUsing="Y","#FFFFFF","#E0E0E0")%>" height="30">
	<td align="center" onclick="popModiCode(<%=oAppurl.FItemList(i).Fidx%>)" style="cursor:pointer" ><%= oAppurl.FItemList(i).Fidx %></td>
	<td align="center">
	<%
		Select Case oAppurl.FItemList(i).Furldiv
			Case 1
				response.write "��ǰ"
			Case 2
				response.write "�̺�Ʈ"
			Case 3
				response.write "�귣��"
			Case 4
				response.write "ī�װ�"
			Case 8
				response.write "�ܺ�URL"
			Case 9
				response.write "Today"
			Case 10
				response.write "����Ʈ"
			Case 11
				response.write "��ٱ���"
		End Select
	%>
	</td>
	<td align="center"><%= oAppurl.FItemList(i).Furltitle %></td>
	<td align="center"><%= oAppurl.FItemList(i).Furlcomplete %></td>
	<td align="center"><%= left(oAppurl.FItemList(i).Fregdate,10) %></td>
	<td align="center"><%= oAppurl.FItemList(i).FisUsing %></td>
	<td align="center"><%= FormatNumber(oAppurl.FItemList(i).Furlhitcount,0) %></td>
	<td align="center"><input type="button" value="<%=chkiif(oAppurl.FItemList(i).Fqrsn<>"","QR�̸�����","QR����")%>" onclick="<%=chkiif(oAppurl.FItemList(i).Fqrsn<>"","popReadCode('" & oAppurl.FItemList(i).Fqrsn& "')","popQrNewCode('"& oAppurl.FItemList(i).Fidx &"','"& oAppurl.FItemList(i).Furltitle &"','"& oAppurl.FItemList(i).Furlcomplete &"')")%>" class="button"></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="center">
	<% if oAppurl.HasPreScroll then %>
		<a href="javascript:gotoPage(<%= oAppurl.StarScrollPage-1 %>)">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oAppurl.StarScrollPage to oAppurl.FScrollCount + oAppurl.StarScrollPage - 1 %>
		<% if i>oAppurl.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:gotoPage(<%= i %>)">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oAppurl.HasNextScroll then %>
		<a href="javascript:gotoPage(<%= i %>)">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<% set oAppurl = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->