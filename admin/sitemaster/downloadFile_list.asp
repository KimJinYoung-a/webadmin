<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/downloadFileCls.asp"-->
<%
'###############################################
' PageName : downloadFile_List.asp
' Discription : ���� �ٿ�ε� ���� ���
'           2012.04.04 ������ �̺�Ʈ�ڵ� �߰�
'           2014.05.09 ������ ������ũ ������ �߰�
'###############################################

dim page, i, lp

page = requestCheckvar(request("page"),10)
if page = "" then page=1

dim oFile
set oFile = New cDownFile
oFile.FCurrPage = page
oFile.FPageSize=20
oFile.FRectUsing = "Y"
oFile.GetfileList

%>
 
<script type="text/javascript">
// ������ �̵�
function goPage(pg)
{
	document.refreshFrm.page.value=pg;
	document.refreshFrm.action="downloadFile_list.asp";
	document.refreshFrm.submit();
}

 //���� �ٿ�ε� ��ũ��Ʈ ����
function copyScrt(vSn) {
	var doc = "javascript:fileDownload(" + vSn + ");";
	copyStringToClipboard(doc);
	alert('�����Ͻ� ������ �ٿ�ε� ��ũ��Ʈ�� ����Ǿ����ϴ�. ����Ͻ� ���� Ctrl+V �Ͻø�˴ϴ�.\n\n���ڹٽ�ũ��Ʈ�̹Ƿ� ��ũ�� �ٹ����� ����Ʈ �������� �� �� �ֽ��ϴ�.');
} 

 //���� �ٿ�ε� ������ũ ����
function copyLink(vSn) {
	var doc = "http://upload.10x10.co.kr/linkweb/download/fileDownload.asp?fn=" + vSn;
	copyStringToClipboard(doc);
	alert('�����Ͻ� ������ �ٿ�ε� ��ũ�� ����Ǿ����ϴ�. ����Ͻ� ���� Ctrl+V �Ͻø�˴ϴ�.\n\n�ش� �������δ� ��ũ��Ʈ�� �̿��Ͻð� �� ��ũ�� ������� ������.');
} 
</script> 
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right"><input type="button" value="���� �߰�" onclick="self.location='downloadFile_Write.asp?mode=add&menupos=<%= menupos %>'" class="button"></td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		�˻���� : <b><%=oFile.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oFile.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ȣ</td>
	<td>�̺�Ʈ�ڵ�</td>
	<td>����</td>
	<td>���ϸ�</td>
	<td>ũ��</td>
	<td>�ٿ�ε�</td>
	<td>�����</td>
	<td>&nbsp;</td>
</tr>
<%	if oFile.FResultCount < 1 then %>
<tr>
	<td colspan="8" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� ������ �����ϴ�.</td>
</tr>
<%
	else
		for i=0 to oFile.FResultCount-1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><a href="downloadFile_Write.asp?mode=edit&menupos=<%= menupos %>&fileSn=<%= oFile.FItemList(i).FfileSn %>"><%= oFile.FItemList(i).FfileSn %></a></td> 
	<td align="center"><a href="downloadFile_Write.asp?mode=edit&menupos=<%= menupos %>&fileSn=<%= oFile.FItemList(i).FfileSn %>"><%= oFile.FItemList(i).Fevt_code %></a></td> 
	<td align="center"><a href="downloadFile_Write.asp?mode=edit&menupos=<%= menupos %>&fileSn=<%= oFile.FItemList(i).FfileSn %>"><%= oFile.FItemList(i).FfileTitle %></a></td>
	<td align="center"><%= oFile.FItemList(i).FfileDownNm & "<br>(" & oFile.FItemList(i).FfileName & ")"%></td>
	<td align="center">
	<%
		if oFile.FItemList(i).FfileSize >= 1048576 then
			Response.Write FormatNumber(oFile.FItemList(i).FfileSize/1024/1024,1) & "MBytes"
		elseif oFile.FItemList(i).FfileSize >= 1024 then
			Response.Write FormatNumber(oFile.FItemList(i).FfileSize/1024,0) & "KBytes"
		else
			Response.Write FormatNumber(oFile.FItemList(i).FfileSize,0) & "Bytes"
		end if
	%>
	</td>
	<td align="center"><%= oFile.FItemList(i).FdownCount & "ȸ<br>" & left(oFile.FItemList(i).FlastDownDate,10) %></td>
	<td align="center"><%= left(oFile.FItemList(i).Fregdate,10) %></td>
	<td align="center"> 
		<input type="button"  id="btnLink" class="button" value="��ũ��Ʈ ����" title="�ٹ����� ����Ʈ�� �ٿ�ε� ��ũ��Ʈ ����" onClick="copyScrt('<%=oFile.FItemList(i).FfileSn %>')"><br>
		<input type="button"  id="btnLink" class="button" value="������ũ ����" title="���� ������ ���� �ٿ�ε� ��ũ ����" onClick="copyLink('<%=oFile.FItemList(i).FfileSn %>')">
	</td>
</tr>
<%
		next
	end if
%>
<!-- ���� ��� �� -->
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<!-- ������ ���� -->
	<%
		if oFile.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oFile.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oFile.StartScrollPage to oFile.FScrollCount + oFile.StartScrollPage - 1

			if lp>oFile.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if oFile.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- ������ �� -->
	</td>
</tr>
</table>
<form name="refreshFrm" method="get" action="downloadFile_list.asp">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
</form>
<%
set oFile = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->