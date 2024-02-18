<%@ Language=VBScript %>
<%
'==========================================================================
'	Description: EMS �߷� ������ ��� ����Ʈ, ������
'	History: 2009.04.07
'==========================================================================
	Option Explicit
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #Include Virtual="/lib/classes/order/clsEms_serviceArea.asp" -->



<%
Dim conListURL	: conListURL = "ems_weightPrice.asp"
Dim conSaveURL	: conSaveURL = "ems_weightPriceSave.asp"
Dim conProcURL	: conProcURL = "ems_weightPriceProc.asp"

Dim PageSize, PerPage, iTotCnt
Dim i

Dim page		: page			    = requestCheckVar(request("page"),10)
Dim scEmsAreaCode	: scEmsAreaCode	= requestCheckVar(request("scEmsAreaCode"),2)
Dim scWeightLimit	: scWeightLimit	= requestCheckVar(request("scWeightLimit"),10)
Dim scCompanyCode	: scCompanyCode	= requestCheckVar(request("scCompanyCode"),3)

if (page="") then page=1
if (scCompanyCode="") then scCompanyCode="EMS"

Dim qString
qString = "scEmsAreaCode=" & scEmsAreaCode & "&scWeightLimit=" & scWeightLimit & "&menupos=" & menupos
conSaveURL = conSaveURL & "?" & qString & "&page=" & page
conListURL = conListURL & "?" & qString

PageSize	= 100		' ������ ������
PerPage		= 10		' ������ ����

' ���̺�Ŭ����
Dim obj	: Set obj = new CEms
obj.FRectPageSize = PageSize
obj.FRectCurrPage = Cint(page)
obj.FRectCompanyCode = scCompanyCode
obj.FRectEmsAreaCode = scEmsAreaCode
obj.FRectWeightLimit = scWeightLimit
obj.GetWeightPriceList

'rw page
'rw scCompanyCode&"scCompanyCode"
'rw IsUsing


%>

<script language="javascript">

// �˻�
function jsSearch(){
    document.frmSearch.submit();
}

</script>
</head>


<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmSearch" method="get" action="<%=conListUrl%>">
	<input type="hidden" name="menupos" value="<%=menupos%>">
		<tr height="40">
			<td width="80" align="center"><font color="#FFFFFF">�˻�����</font></td>
			<td bgcolor="#FFFFFF">
				<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a">
					<tr>
						<td>
                            * <label>��ü�ڵ� :
                                <select name="scCompanyCode" class="select">
								    <option value="EMS" <%=chkIIF(scCompanyCode="EMS","selected","")%>>EMS</option>
								    <option value="UPS" <%=chkIIF(scCompanyCode="UPS","selected","")%>>UPS</option>
							    </select>
                            </label>
							* ����������� :  <input type="text" name="scEmsAreaCode" value="<%=scEmsAreaCode%>" size="2" maxlength="2"> (0,1~5,A~E)
							* �߷� : <input type="text" name="scWeightLimit" value="<%=scWeightLimit%>" size="8" maxlength="10"> (g, �׶�)
						</td>
					</tr>
				</table>
			</td>
			<td width="60" bgcolor="#FFFFFF" align="center"><a href="javascript:jsSearch();"><font class="text_blue">�˻�</font></a></td>
		</tr>
	</form>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="100%" class="a">
	<tr>
		<td>
			<a href="ems_weightPriceSave.asp?menupos=<%= menupos %>&CompanyCode=<%= scCompanyCode %>&scCompanyCode=<%= scCompanyCode %>"><font class="text_blue">+ EMS �߷���������ݵ��</font></a>
		</td>
		<td align="right">�� <%=obj.FTotalCount%> �� <%=page%>/<%=obj.FTotalPage%></td>
	</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center">��ü�ڵ�</td>
        <td align="center">�����������</td>
		<td align="center">�߷�</td>
		<td align="center">����</td>

	</tr>
<%For i = 0 To obj.FResultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
	    <td><a href="<%=conSaveURL%>&emsAreaCode=<%=obj.FItemList(i).FemsAreaCode%>&WeightLimit=<%=obj.FItemList(i).FweightLimit%>&companyCode=<%=obj.FItemList(i).FcompanyCode%>"><%=obj.FItemList(i).FcompanyCode%></a></td>
	    <td><%= obj.FItemList(i).FemsAreaCode %></td>
        <td><%=FormatNumber(obj.FItemList(i).FWeightLimit,0)%></td>
	    <td><%=FormatNumber(obj.FItemList(i).FemsPrice,0)%></td>

	</tr>
<%Next%>
    <tr bgcolor="#FFFFFF">
		<td align="center" colspan="20">
		 <% sbDisplayPaging "page="&page, obj.FTotalCount, PageSize, PerPage%>
		</td>
	</tr>
</table>
<%
Set obj = Nothing
%>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
