<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ���޸� [����������]����
' �ش� �Ŵ� ������ �����Ұ�� �ݵ�� ��ü ���� ������ ���� �ϼž� �մϴ�. �ΰ��� �ݾ��� ��ġ�ؾ� �մϴ�.
' History : 2017.07.17 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/commissionjungsan_cls.asp"-->

<%
dim yyyy, mm, stdate, arrlist, i, page, orderserial, itemnname
	yyyy = requestcheckvar(getNumeric(request("yyyy")),4)
	mm = requestcheckvar(getNumeric(request("mm")),2)
	page = requestcheckvar(getNumeric(request("page")),10)
	orderserial = requestcheckvar(getNumeric(request("orderserial")),11)
	itemnname = requestcheckvar(request("itemnname"),10)

if page="" then page=1
if yyyy="" then
	stdate = dateadd("m", -1, date())
	stdate = DateSerial(Left(stdate,4), CLng(Mid(stdate,6,2)),1)
	yyyy = Left(stdate,4)
	mm = Mid(stdate,6,2)
end if

dim cjungsan
Set cjungsan = New Ccommission
	cjungsan.FRectyyyymm = yyyy + "-" + mm
	cjungsan.FPageSize = 500
	cjungsan.FCurrPage = page
	cjungsan.frectorderserial = orderserial
	cjungsan.frectitemname = itemnname
	cjungsan.frectrdsite = "daumshop"
	cjungsan.Getcommissionjungsan_daum_paging()
%>

<script type='text/javascript'>

function searchSubmit(page){
	frm.page.value=page;
	frm.submit();
}

function regcommissionjungsan(vmode){
	frm.action='/admin/jungsan/commission/commissionjungsan_process.asp';
	frm.target='view';
	frm.mode.value=vmode;
	frm.submit();
	frm.action='';
	frm.target='';
	frm.mode.value='';
}

function downloadfile(vmode){
	frm.action='/admin/jungsan/commission/commissionjungsan_process.asp';
	frm.target='view';
	frm.mode.value=vmode;
	frm.submit();
	frm.action='';
	frm.target='';
	frm.mode.value='';
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="reload" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="rdsite" value="daumshop">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="30" bgcolor="<%= adminColor("gray") %>">�˻�<Br>����</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				*�Ⱓ :
				<% DrawYMBoxdynamic "yyyy", yyyy, "mm", mm, "" %>
				&nbsp;
				*�ֹ���ȣ : <input type="text" name="orderserial" value="<%= orderserial %>" size="15" maxlength=15>
				&nbsp;
				*��ǰ�� : <input type="text" name="itemnname" value="<%= itemnname %>" size="25" maxlength=64>
			</td>
		</tr>
	    </table>
	</td>	
	<td width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit('');"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
<!-- ǥ �߰��� ����-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
    	<input type="button" onclick="regcommissionjungsan('regdaum');" value="<%= yyyy %>��<%= mm %>�������ۼ�" class="button">
    </td>
    <td align="right">
    	<input type="button" onclick="downloadfile('csvdaum');" value="CSV�ٿ�" class="button">
    </td>        
</tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= cjungsan.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= cjungsan.FTotalPage %></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td width=70>�ֹ�����</td>
    <td width=70>�������<br>/Ȯ������</td>
    <td width=120>�����ڵ�</td>
    <td width=60>����ϱ���</td>
    <td width=80>�ֹ���ȣ</td>
    <td>��ǰ��</td>
    <td width=50>�ֹ�����</td>
    <td width=70>�ֹ��ݾ�<br>(V.A.T����)</td>
    <td width=60>��������</td>
    <td width=70>������</td>
    <td width=70>�ֹ�����</td>
    <td width=70>��ҳ�¥</td>
</tr>					  		  	

<%
if cjungsan.FResultCount > 0 then
	
For i = 0 To cjungsan.FResultCount-1
%>
	<tr bgcolor="#FFFFFF" align="center">
		<td>
			<%= cjungsan.FItemList(i).frDate %>
		</td>
		<td>
			<%= cjungsan.FItemList(i).ffixedDate %>
		</td>
		<td>
			<%= cjungsan.FItemList(i).frdsite %>
		</td>
		<td>
			<%= cjungsan.FItemList(i).fdevice %>
		</td>
		<td>
			<%= cjungsan.FItemList(i).forderserial %>
		</td>
		<td align="left">
			<%= cjungsan.FItemList(i).fitemname %>
		</td>
		<td align="right">
			<%= cjungsan.FItemList(i).fitemno %>
		</td>
		<td align="right">
			<%= FormatNumber(cjungsan.FItemList(i).fsuppPrc,0) %>
		</td>
		<td align="right">
			<%= cjungsan.FItemList(i).fcommpro %>
		</td>
		<td align="right">
			<%= cjungsan.FItemList(i).fcommissoin %>
		</td>
		<td>
			<%= cjungsan.FItemList(i).fordStatName %>
		</td>
		<td>
			<%= cjungsan.FItemList(i).fcancelDT %>
		</td>
	</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if cjungsan.HasPreScroll then %>
			<span class="list_link"><a href="#" onclick="searchSubmit('<%= cjungsan.StartScrollPage-1 %>'); return false;">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + cjungsan.StartScrollPage to cjungsan.StartScrollPage + cjungsan.FScrollCount - 1 %>
			<% if (i > cjungsan.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(cjungsan.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="#" onclick="searchSubmit('<%= i %>'); return false;" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if cjungsan.HasNextScroll then %>
			<span class="list_link"><a href="#" onclick="searchSubmit('<%= i %>'); return false;">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="25">��ϵ� ������ �����ϴ�.</td>
	</tr>
<% end if %>

</table>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>
<%
set cjungsan = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->