<%@ Language=VBScript %>
<%
'==========================================================================
'	Description: EMS �������� ����Ʈ
'	History: 2009.04.07 ������ ����
'			 2017.11.02 �ѿ�� ����
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

Dim conListURL	: conListURL = "ems_serviceareaList.asp"
Dim conSaveURL	: conSaveURL = "ems_serviceareaSave.asp"
Dim conProcURL	: conProcURL = "ems_serviceareaProc.asp"

Dim PageSize, PerPage, iTotCnt
Dim i

Dim page			: page			    = requestCheckVar(request("page"),10)
Dim scCountryCode	: scCountryCode		= requestCheckVar(request("scCountryCode"),2)
Dim IsUsing			: IsUsing		    = requestCheckVar(request("IsUsing"),1)
Dim CountryNameKr	: CountryNameKr		= requestCheckVar(request("CountryNameKr"),50)
Dim CountryNameEn	: CountryNameEn		= requestCheckVar(request("CountryNameEn"),50)
Dim EmsAreaCode		: EmsAreaCode		= requestCheckVar(request("EmsAreaCode"),2)
Dim CompanyCode		: CompanyCode		= requestCheckVar(request("CompanyCode"),3)

if (page="") then page=1
if (CompanyCode="") then CompanyCode="EMS"

Dim qString
qString = "scCountryCode=" & scCountryCode & "&IsUsing=" & IsUsing & "&menupos=" & menupos
conSaveURL = conSaveURL & "?" & qString & "&page=" & page
conListURL = conListURL & "?" & qString

PageSize	= 200		' ������ ������
PerPage		= 10		' ������ ����

' ���̺�Ŭ����
Dim obj	: Set obj = new CEMS
obj.FRectPageSize = PageSize
obj.FRectCurrPage = Cint(page)
obj.FRectCountryCode = scCountryCode
obj.FRectIsUsing = IsUsing
obj.FRectCountryNameKr = html2db(CountryNameKr)
obj.FRectCountryNameEn = html2db(CountryNameEn)
obj.FRectEmsAreaCode = html2db(EmsAreaCode)
obj.FRectCompanyCode = html2db(CompanyCode)
obj.GetServiceAreaList

'rw page
'rw scCountryCode&"scCountryCode"
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
                                <select name="CompanyCode" class="select">
								    <option value="EMS" <%=chkIIF(CompanyCode="EMS","selected","")%>>EMS</option>
								    <option value="UPS" <%=chkIIF(CompanyCode="UPS","selected","")%>>UPS</option>
							    </select>
                            </label>
                            * <label>�����ڵ� :  <input type="text" name="scCountryCode" value="<%=scCountryCode%>" size="4" maxlength="2" /></label>
							* <label>������(�ѱ�) :  <input type="text" name="CountryNameKr" value="<%=CountryNameKr%>" size="16" maxlength="16" /></label>
							* <label>������(����) :  <input type="text" name="CountryNameEn" value="<%=CountryNameEn%>" size="16" maxlength="16" /></label>
							* <label>EMS ����������� :  <input type="text" name="EmsAreaCode" value="<%=EmsAreaCode%>" size="2" maxlength="2" /></label>
							* <label>��뿩�� :
								<select name="IsUsing" class="select">
								<option value="">::��ü::</option>
								<option value="Y" <%=chkIIF(IsUsing="Y","selected","")%>>���</option>
								<option value="N" <%=chkIIF(IsUsing="N","selected","")%>>������</option>
								</select>
							  </label>
						</td>
					</tr>
				</table>
			</td>
			<td width="60" bgcolor="#FFFFFF" align="center">
				<input type="button" class="button" value="�˻�" onClick="jsSearch()">
			</td>
		</tr>
	</form>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="100%" class="a">
	<tr>
		<td>
			<a href="ems_serviceareaSave.asp?menupos=<%= menupos %>"><font class="text_blue">+ EMS �����������</font></a>
		</td>
		<td align="right">�� <%=obj.FTotalCount%> �� <%=page%>/<%=obj.FTotalPage%></td>
	</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td class="line01" align="center">��ü�ڵ�</td>
        <td class="line01" align="center">�����ڵ�</td>
		<td class="line01" align="center">������(�ѱ�)</td>
		<td class="line01" align="center">������(����)</td>
		<td class="line01" align="center">EMS�����������</td>
		<td class="line01" align="center">EMS�ִ��߷�</td>
		<td class="line01" align="center">�����κδ㿩��</td>
		<td class="line01" align="center">��뿩��</td>
		<td class="line01" align="center">��Ÿ����</td>
        <td class="line01" align="center">��</td>
	</tr>
<%For i = 0 To obj.FResultCount-1 %>
	<tr align="center" bgcolor="<%=chkIIF(obj.FItemList(i).Fisusing="Y","#FFFFFF","#EFEFEF")%>">
	    <td><a href="<%=conSaveURL%>&companyCode=<%=obj.FItemList(i).FcompanyCode%>&countryCode=<%=obj.FItemList(i).FcountryCode%>"><%= obj.FItemList(i).FcompanyCode %></a></td>
	    <td><%=obj.FItemList(i).FcountryCode%></td>
        <td><%=obj.FItemList(i).FcountryNameKr%></td>
	    <td><%=obj.FItemList(i).FcountryNameEn%></td>
	    <td><%=obj.FItemList(i).FemsAreaCode%></td>
	    <td><%=obj.FItemList(i).FemsMaxWeight%></td>
	    <td><%=obj.FItemList(i).FreceiverPay%></td>
	    <td><%=obj.FItemList(i).Fisusing%></td>
	    <td><%=obj.FItemList(i).FetcContents%></td>
        <td><a href="javascript:popForeignDeliverInfo('<%=obj.FItemList(i).FcountryCode%>', '<%= obj.FItemList(i).FcompanyCode %>');">����</a></td>
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
