<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/company/nv/incGlobalVariable.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/company/nv/navepCls_T.asp" -->
<%
Dim searchtype, searchrect, meCode, mtype
Dim orderserial, yyyy1, yyyy2, mm1, mm2, dd1, dd2
Dim nowdate, searchnextdate
nowdate = Left(CStr(now()),10)

orderserial = requestCheckvar(request("orderserial"),16)
searchtype	= requestCheckvar(request("searchtype"),16)
meCode		= requestCheckvar(request("meCode"),16)
searchrect	= requestCheckvar(request("searchrect"),32)
yyyy1		= requestCheckvar(request("yyyy1"),4)
mm1			= requestCheckvar(request("mm1"),2)
dd1			= requestCheckvar(request("dd1"),2)
yyyy2		= requestCheckvar(request("yyyy2"),4)
mm2			= requestCheckvar(request("mm2"),2)
dd2			= requestCheckvar(request("dd2"),2)
mtype       = requestCheckvar(request("mtype"),2)
If (yyyy1 = "") Then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
End If

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)
Dim page
Dim ojumun
page = request("page")
If (page = "") Then page = 1
if (mtype="") then mtype="rg"

Set ojumun = new CJumunMaster
ojumun.FPageSize = 30
ojumun.FCurrPage = page
ojumun.FRectRegStart = yyyy1 & "-" & mm1 & "-" & dd1
ojumun.FRectRegEnd = searchnextdate
ojumun.FRectMType = mtype

'If searchtype="01" Then
'	ojumun.FRectBuyname = searchrect
'ElseIf searchtype="02" Then
'	ojumun.FRectReqName = searchrect
'ElseIf searchtype="03" Then
'	ojumun.FRectUserID = searchrect
'ElseIf searchtype="04" Then
'	ojumun.FRectIpkumName = searchrect
'ElseIf searchtype="06" Then
'	ojumun.FRectSubTotalPrice = searchrect
'End If

If session("ssBctDiv")="999" then
	ojumun.FRectRdSite = session("ssBctID")
Else
	ojumun.FRectSiteName = session("ssBctID")
End If

ojumun.FRectOrderSerial = orderserial
ojumun.FRectMeCode = meCode

if (session("ssBctID")<>"") then
    ojumun.navEpJumunList()
end if

Dim ix,iy
%>
<script language='javascript'>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
function ViewOrderDetail(os){
    var frm = document.frmDtl;
    frm.target = '_ViewOrderDetail';
    frm.orderserial.value=os;
    frm.action="viewordermaster.asp"
	frm.submit();
}
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<body>
<table width="700" border="0" class="a">
<tr>
	<td>&gt;&gt;��������</td>
</tr>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr>
	<td class="a" >
	�ֹ���ȣ :
	<input type="text" name="orderserial" value="<%= orderserial %>" size="11" maxlength="16">
	&nbsp;
	<select name="mtype" class="select">
	<option value="rg" <%= ChkIIF(mtype = "rg", "selected", "") %> >�ֹ���
	<option value="ip" <%= ChkIIF(mtype = "ip", "selected", "") %> >������
	<option value="fx" <%= ChkIIF(mtype = "fx", "selected", "") %> >������
	</select>
	&nbsp;
	�˻��Ⱓ :<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	<br>

	�����ڵ� :
	<select name="meCode" class="select">
		<option value="">--����--</option>
		<option value="nvshop_boxlogo"				<%= ChkIIF(meCode = "nvshop_boxlogo", "selected", "") %> >���� ��ũ</option>
		<option value="nvshop_boxA1" 				<%= ChkIIF(meCode = "nvshop_boxA1", "selected", "") %> >�׸�����(���̽� ������) ����-1</option>
		<option value="nvshop_boxA2"				<%= ChkIIF(meCode = "nvshop_boxA2", "selected", "") %> >�׸�����(���̽� ������) ����-2</option>
		<option value="nvshop_castleft"				<%= ChkIIF(meCode = "nvshop_castleft", "selected", "") %> >(����) ���θ� �ΰ�</option>
		<option value="nvshop_castright" 			<%= ChkIIF(meCode = "nvshop_castright", "selected", "") %> >(����) ���θ� �ٷΰ���</option>
		<option value="nvshop_cast1"				<%= ChkIIF(meCode = "nvshop_cast1", "selected", "") %> >����ĳ��Ʈ ����-1</option>
		<option value="nvshop_cast2"				<%= ChkIIF(meCode = "nvshop_cast2", "selected", "") %> >����ĳ��Ʈ ����-2</option>
		<option value="nvshop_mens"					<%= ChkIIF(meCode = "nvshop_mens", "selected", "") %> >������</option>
		<option value="nvshop_luckmain"				<%= ChkIIF(meCode = "nvshop_luckmain", "selected", "") %> >���� ��ǰ</option>
		<option value="nvshop_lucksub"				<%= ChkIIF(meCode = "nvshop_lucksub", "selected", "") %> >���� ��ǰ</option>
		<option value="nvshop_sp"					<%= ChkIIF(meCode = "nvshop_sp", "selected", "") %> >��ǰEP</option>
		<option value="nvshop_logo"					<%= ChkIIF(meCode = "nvshop_logo", "selected", "") %> >���ļ��� ���� �̹��� �ΰ�</option>
		<option value="nvshop_logo2"				<%= ChkIIF(meCode = "nvshop_logo2", "selected", "") %> >���ļ��� ���� ����</option>
		<option value="nvshop_sticb"				<%= ChkIIF(meCode = "nvshop_sticb", "selected", "") %> >��ƽ ���</option>
		<option value="nvshop_mainb"				<%= ChkIIF(meCode = "nvshop_mainb", "selected", "") %> >�ָ��ȹ��</option>
		<option value="nvshop_pb"					<%= ChkIIF(meCode = "nvshop_pb", "selected", "") %> >�÷��� ���</option>
		<option value="nvshop_exhb"					<%= ChkIIF(meCode = "nvshop_exhb", "selected", "") %> >��ȹ�� ����(ī�װ� ����)</option>
		<option value="WEB_ALL"						<%= ChkIIF(meCode = "WEB_ALL", "selected", "") %> >==== �� ���� ��ü ====</option>
		<option value="mobile_nvshop_boxlogo"		<%= ChkIIF(meCode = "mobile_nvshop_boxlogo", "selected", "") %> >���� ��ũ[�����]</option>
		<option value="mobile_nvshop_boxA1"			<%= ChkIIF(meCode = "mobile_nvshop_boxA1", "selected", "") %> >�׸�����(���̽� ������) ����-1[�����]</option>
		<option value="mobile_nvshop_boxA2"			<%= ChkIIF(meCode = "mobile_nvshop_boxA2", "selected", "") %> >�׸�����(���̽� ������) ����-2[�����]</option>
		<option value="mobile_nvshop_castleft"		<%= ChkIIF(meCode = "mobile_nvshop_castleft", "selected", "") %> >(����) ���θ� �ΰ� [�����]</option>
		<option value="mobile_nvshop_castright"		<%= ChkIIF(meCode = "mobile_nvshop_castright", "selected", "") %> >(����) ���θ� �ٷΰ��� [�����]</option>
		<option value="mobile_nvshop_cast1"			<%= ChkIIF(meCode = "mobile_nvshop_cast1", "selected", "") %> >����ĳ��Ʈ ����-1 [�����]</option>
		<option value="mobile_nvshop_cast2"			<%= ChkIIF(meCode = "mobile_nvshop_cast2", "selected", "") %> >����ĳ��Ʈ ����-2[�����]</option>
		<option value="mobile_nvshop_mens"			<%= ChkIIF(meCode = "mobile_nvshop_mens", "selected", "") %> >������[�����]</option>
		<option value="mobile_nvshop_luckmain"		<%= ChkIIF(meCode = "mobile_nvshop_luckmain", "selected", "") %> >���� ��ǰ [�����]</option>
		<option value="mobile_nvshop_lucksub"		<%= ChkIIF(meCode = "mobile_nvshop_lucksub", "selected", "") %> >���� ��ǰ [�����]</option>
		<option value="mobile_nvshop_sp"			<%= ChkIIF(meCode = "mobile_nvshop_sp", "selected", "") %> >��ǰEP  [�����]</option>
		<option value="mobile_nvshop_logo"			<%= ChkIIF(meCode = "mobile_nvshop_logo", "selected", "") %> >���ļ��� ���� �̹��� �ΰ� [�����]</option>
		<option value="mobile_nvshop_logo2"			<%= ChkIIF(meCode = "mobile_nvshop_logo2", "selected", "") %> >���ļ��� ���� ����[�����]</option>
		<option value="mobile_nvshop_sticb"			<%= ChkIIF(meCode = "mobile_nvshop_sticb", "selected", "") %> >��ƽ ��� [�����]</option>	
		<option value="mobile_nvshop_mainb"			<%= ChkIIF(meCode = "mobile_nvshop_mainb", "selected", "") %> >�ָ��ȹ�� [�����]</option>
		<option value="mobile_nvshop_pb"			<%= ChkIIF(meCode = "mobile_nvshop_pb", "selected", "") %> >�÷��� ��� [�����]</option>
		<option value="mobile_nvshop_exhb"			<%= ChkIIF(meCode = "mobile_nvshop_exhb", "selected", "") %> >��ȹ�� ����(ī�װ� ����) [�����]</option>
		<option value="MOBILE_ALL"					<%= ChkIIF(meCode = "MOBILE_ALL", "selected", "") %> >==== ����� ���� ��ü ====</option>
	</select>

	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</form>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr height="20" bgcolor="#FFFFFF">
	<td colspan="15" align="right">
		�� �Ǽ� : <Font color="#3333FF"><%= FormatNumber(ojumun.FTotalCount,0) %></font>
		&nbsp; page : <%= ojumun.FCurrPage %>/<%=ojumun.FTotalPage %>&nbsp;
    </td>
</tr>

<% if (mtype="fx") then %>
    <% If ojumun.FTotalCount>0 then %>
    <tr height="30" bgcolor="#FFFFFF" align="center">
    	<td >�հ�</td>
    	<td ></td>
    	<td ></td>
    	<td ></td>
    	<td ></td>
    	<td><%= FormatNumber(ojumun.FOneItem.getJungsanTargetNoVatSum,0) %></td>
    	<td></td>
    	<td><%= FormatNumber(ojumun.FOneItem.FcommiSum,0) %></td>
    	<td></td>
    	<td ></td>
    	<td ></td>
    </tr>
    <% end if %>
    <tr height="30" bgcolor="#FFD8D8" align="center">
    	<td width="100" >�ֹ���ȣ</td>
    	<td width="100" >�ֹ�����</td>
    	<td width="100" >Ȯ������</td>
    	<td width="100" >��ǰ��</td>
    	<td width="100" >�ֹ�����</td>
    	<td width="100" >�ֹ��ݾ�(vat����)</td>
    	<td width="100" >��������</td>
    	<td width="100" >������</td>
    	<td width="100" >�ֹ�����</td>
    	<td width="40">�����<br>����</td>
    	<td >�����ڵ�</td>
    </tr>    
    <% If ojumun.FresultCount < 1 Then %>
    <tr height="60" bgcolor="#FFFFFF">
    	<td colspan="14" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
    <% Else %>
    <% For ix = 0 To ojumun.FresultCount - 1 %>
    <tr class="a"  height="30" bgcolor="#FFFFFF" align="center">
    	<td><a href="#" onclick="ViewOrderDetail('<%= ojumun.FMasterItemList(ix).FOrderSerial %>')" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
    	<td><%= ojumun.FMasterItemList(ix).GetRegDate %></td>
    	<td><%= Left(ojumun.FMasterItemList(ix).Fbeadaldate,10) %></td>
    	<td><%= ojumun.FMasterItemList(ix).FitemNOptionName %></td>
    	<td><%= ojumun.FMasterItemList(ix).Fitemno %></td>
    	<td><%= FormatNumber(ojumun.FMasterItemList(ix).getJungsanTargetNoVatSum,0) %></td>
    	<td><%= ojumun.FMasterItemList(ix).Fcommpro %></td>
    	<td><%= FormatNumber(ojumun.FMasterItemList(ix).FcommiSum,0) %></td>
    	<td><%= ojumun.FMasterItemList(ix).FordStatName %>
    	    
    	    <%= ojumun.FMasterItemList(ix).getCanceldate %>
    	    </td>
    	<td><%= CHKIIF(ojumun.FMasterItemList(ix).Fismobile=1,"Y","") %></td>
    	<td ><%= ojumun.FMasterItemList(ix).getRdSiteName %></td>
    </tr>
    <% Next %>
    <tr bgcolor="#FFFFFF">
	<td colspan="14" height="30" align="center">
	<% If ojumun.HasPreScroll Then %>
		<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
	<% Else %>
		[pre]
	<%
	   End If
		For ix = 0 + ojumun.StartScrollPage To ojumun.FScrollCount + ojumun.StartScrollPage - 1
			If ix>ojumun.FTotalpage Then Exit For
			If CStr(page) = CStr(ix) Then
	%>
		<font color="red">[<%= ix %>]</font>
	<%		Else %>
		<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
	<%
			End If
		Next
		If ojumun.HasNextScroll Then
	%>
		<a href="javascript:NextPage('<%= ix %>')">[next]</a>
	<%	Else %>
		[next]
	<%	End If %>
	</td>
</tr>
    <% end if %>
<% else %>
    <% If ojumun.FTotalCount>0 then %>
    <tr height="30" bgcolor="#FFFFFF" align="center">
    	<td >�հ�</td>
    	<td ></td>
    	<td ></td>
    	<td ></td>
    	<td ></td>
    	<td><%= FormatNumber(ojumun.FOneItem.FTotalSum,0) %></td>
    	<td><%= FormatNumber(ojumun.FOneItem.getEnuiSum,0) %></td>
    	<td><%= FormatNumber(ojumun.FOneItem.getDlvPaySum,0) %></td>
    	<td><%= FormatNumber(ojumun.FOneItem.getJungsanTargetNoVatSum,0) %></td>
    	<td ></td>
    	<td ></td>
    </tr>
    <% end if %>
    <tr height="30" bgcolor="#FFD8D8" align="center">
    	<td width="100" >�ֹ���ȣ</td>
    	<td width="100" >�ֹ���</td>
    	<td width="100" >������</td>
    	<td width="100" >�����</td>
    	<td width="100" >������</td>
    	<td width="100" >�ֹ��ݾ�</td>
    	<td width="100" >�������ݾ�</td>
    	<td width="100" >��ۺ�</td>
    	<td width="100" >����ݾ�</td>
    	<td width="40">�����<br>����</td>
    	<td >�����ڵ�</td>
    </tr>
    <% If ojumun.FresultCount < 1 Then %>
    <tr height="60" bgcolor="#FFFFFF">
    	<td colspan="14" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
    <% Else %>
    <% For ix = 0 To ojumun.FresultCount - 1 %>
    <tr class="a"  height="30" bgcolor="#FFFFFF" align="center">
    	<td><a href="#" onclick="ViewOrderDetail('<%= ojumun.FMasterItemList(ix).FOrderSerial %>')" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
    	<td><%= ojumun.FMasterItemList(ix).GetRegDate %></td>
    	<td><%= Left(ojumun.FMasterItemList(ix).Fipkumdate,10) %></td>
    	<td><%= ojumun.FMasterItemList(ix).getCanceldate %></td>
    	<td><%= ojumun.FMasterItemList(ix).getJungsanFixdate %></td>
    	<td><%= FormatNumber(ojumun.FMasterItemList(ix).FTotalSum,0) %></td>
    	<td><%= FormatNumber(ojumun.FMasterItemList(ix).getEnuiSum,0) %></td>
    	<td><%= FormatNumber(ojumun.FMasterItemList(ix).getDlvPaySum,0) %></td>
    	<td><%= FormatNumber(ojumun.FMasterItemList(ix).getJungsanTargetNoVatSum,0) %></td>
    	<td><%= CHKIIF(ojumun.FMasterItemList(ix).isMobileOrder,"Y","") %></td>
    	<td ><%= ojumun.FMasterItemList(ix).getRdSiteName %></td>
    </tr>
    <% Next %>
    <tr bgcolor="#FFFFFF">
	<td colspan="14" height="30" align="center">
	<% If ojumun.HasPreScroll Then %>
		<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
	<% Else %>
		[pre]
	<%
	   End If
		For ix = 0 + ojumun.StartScrollPage To ojumun.FScrollCount + ojumun.StartScrollPage - 1
			If ix>ojumun.FTotalpage Then Exit For
			If CStr(page) = CStr(ix) Then
	%>
		<font color="red">[<%= ix %>]</font>
	<%		Else %>
		<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
	<%
			End If
		Next
		If ojumun.HasNextScroll Then
	%>
		<a href="javascript:NextPage('<%= ix %>')">[next]</a>
	<%	Else %>
		[next]
	<%	End If %>
	</td>
</tr>
    <% end if %>


<% End If %>
</table>
<form name="frmDtl" method="post">
<input type="hidden" name="orderserial">
</form>
</body>
</html>
<% Set ojumun = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->