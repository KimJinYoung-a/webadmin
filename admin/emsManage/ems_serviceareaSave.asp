<%@ Language=VBScript %>
<%
'==========================================================================
'	Description: EMS �������� �����, ������
'	History: 2009.04.07
'==========================================================================
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #Include Virtual="/lib/classes/order/clsEms_serviceArea.asp" -->

<%
Dim conListURL	: conListURL = "ems_serviceareaList.asp"
Dim conSaveURL	: conSaveURL = "ems_serviceareaSave.asp"
Dim conProcURL	: conProcURL = "ems_serviceareaProc.asp"

Dim page		: page			= requestCheckVar(request("page"),10)
Dim scCountryCode	: scCountryCode	= requestCheckVar(request("scCountryCode"),2)
Dim CountryCode	: CountryCode	= requestCheckVar(request("CountryCode"),2)
Dim IsUsing		: IsUsing		= requestCheckVar(request("IsUsing"),1)
Dim CompanyCode		: CompanyCode		= requestCheckVar(request("CompanyCode"),3)


'referer�� ��ü
'Dim qString: qString = "scCountryCode=" & scCountryCode & "&IsUsing=" & IsUsing & "&menupos=" & menupos
'conProcURL = conProcURL & "?" & qString & "&page=" & page
Dim retUrl : retUrl = request.ServerVariables("HTTP_REFERER")


' ���̺�Ŭ����
Dim obj	: Set obj = new CEms
obj.FRectCompanyCode = html2db(CompanyCode)
obj.FRectCountryCode = CountryCode
obj.GetServiceAreaData


' ȭ��ǥ������
Dim pageInfo1, pageInfo2, pageInfo3
If countryCode = "" Then
	pageInfo1 = "���"
	pageInfo2 = "�� ��"
Else
	pageInfo1 = "����"
	pageInfo2 = "�� ��"
End If
%>

<script language='javascript'>
<!--

// ���,����,���� ó��
function jsSubmit(mode)
{
	var f = document.frmWrite;
	if (!mode)
		if (f.OcountryCode.value=="")
			f.mode.value = "INS";
		else
			f.mode.value = "UPD";
	else
		f.mode.value = mode;

    if (f.companyCode.value.length!=3){
        alert('��ü�ڵ� 3�ڸ� �Է��ϼ���.');
        return;
    }

    if (f.countryCode.value.length!=2){
        alert('�����ڵ� 2�ڸ� �Է��ϼ���.');
        return;
    }

	if (!validField(f.countryNameKr, "������(�ѱ�)��"))return ;
	if (!validField(f.countryNameEn, "������(����)��"))return ;
	if (!validField(f.emsAreaCode, "EMS�������������"))	return ;
	if (!validField(f.emsMaxWeight, "EMS�ִ��߷���"))return ;
	if (!validField(f.receiverPay, "�����κδ㿩�θ�"))	return ;
	if (!validField(f.isusing, "��뿩�θ�"))		return ;

	f.submit();

}

//-->
</script>
</head>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmWrite" method="post" action="<%=conProcURL%>" />
<input type="hidden" name="mode" />
<input type="hidden" name="OcompanyCode" value="<%= obj.FoneItem.FCompanyCode %>" />
<input type="hidden" name="OcountryCode" value="<%= obj.FoneItem.FcountryCode %>" />
<input type="hidden" name="retUrl" value="<%=retUrl%>" />

	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ü�ڵ�</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<% if obj.FoneItem.FCompanyCode="" then %>
			<input type="text" class="text" name="companyCode" value="<%=obj.FoneItem.FCompanyCode%>" size="3" maxlength="3">
		<% else %>
		    <input type="hidden" name="companyCode" value="<%=obj.FoneItem.FCompanyCode%>" size="3" maxlength="3">
		    <%=obj.FoneItem.FCompanyCode%>
	    <% end if %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����ڵ�</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<% if obj.FoneItem.FcountryCode="" then %>
			<input type="text" class="text" name="countryCode" value="<%=obj.FoneItem.FcountryCode%>" size="2" maxlength="2">
		(����2)
		<% else %>
		    <input type="hidden" name="countryCode" value="<%=obj.FoneItem.FcountryCode%>" size="2" maxlength="2">
		    <%=obj.FoneItem.FcountryCode%>
	    <% end if %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">������(�ѱ�)</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="countryNameKr" value="<%=doubleQuote(obj.FoneItem.FcountryNameKr)%>" size="50" maxlength="50"> (50 byte)
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">������(����)</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="countryNameEn" value="<%=doubleQuote(obj.FoneItem.FcountryNameEn)%>" size="50" maxlength="50"> (50 byte)
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">EMS�����������</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="emsAreaCode" value="<%=doubleQuote(obj.FoneItem.FemsAreaCode)%>" size="1" maxlength="1"> (0=������,1~5,A~E)
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">EMS�ִ��߷�</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="emsMaxWeight" value="<%=obj.FoneItem.FemsMaxWeight%>" size="8" maxlength="8" style="ime-mode:disabled;" onkeydown="onlyNumber(this,event);"> (g)
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����κδ㿩��</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<label><input type="radio" class="radio" name="receiverPay" value="Y" <%=chkIIF(obj.FoneItem.FreceiverPay="Y","checked","")%> />���</label>
			<label><input type="radio" class="radio" name="receiverPay" value="N" <%=chkIIF(obj.FoneItem.FreceiverPay="N","checked","")%> />������</label>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<label><input type="radio" class="radio" name="isusing" value="Y" <%=chkIIF(obj.FoneItem.Fisusing="Y","checked","")%> />���</label>
			<label><input type="radio" class="radio" name="isusing" value="N" <%=chkIIF(obj.FoneItem.Fisusing="N","checked","")%> />������</label>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">��Ÿ����</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text" name="etcContents" value="<%=doubleQuote(obj.FoneItem.FetcContents)%>" size="100" maxlength="100">
		</td>
	</tr>

	<tr>
		<td align="center" colspan="4" bgcolor="#FFFFFF">
			<input type="button" class="button" value=" <%=pageInfo2%> " onClick="jsSubmit();">
			&nbsp;&nbsp;
			<input type="button" class="button" value=" �� �� " onClick="history.back();">
		</td>
	</tr>
</table>
</form>
<%
Set obj = Nothing
%>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
