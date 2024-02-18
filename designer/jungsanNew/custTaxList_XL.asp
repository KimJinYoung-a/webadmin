<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<%

'// ============================================================================
dim makerid, yyyy1,mm1
makerid = session("ssBctID")
yyyy1   = requestCheckvar(request("yyyy1"),10)
mm1     = requestCheckvar(request("mm1"),10)

if (yyyy1="") then
    yyyy1 = LEFT(dateadd("m",-1,now()),4)
    mm1 = MID(dateadd("m",-1,now()),6,2)
end if

dim startDate, endDate
startDate = yyyy1 & "-" & mm1 & "-01"
endDate = Left(DateAdd("m", 1, DateSerial(yyyy1, mm1, 1)), 10)


'// ============================================================================
dim opartner, i, page, groupid
set opartner = new CPartnerUser
opartner.FCurrpage = 1
opartner.FRectDesignerID = makerid
opartner.FPageSize = 1
opartner.GetOnePartnerNUser

groupid = opartner.FOneItem.FGroupid

dim ogroup
''set ogroup = new CPartnerGroup
''ogroup.FRectGroupid = groupid
''ogroup.GetOneGroupInfo


'// ============================================================================
page   = requestCheckvar(request("page"),10)

if (page = "") then
	page = "1"
end if


dim oTax
set oTax = new CTax
oTax.FCurrPage = 1
oTax.FPageSize = 200						'// �ִ� 200��
oTax.FRectSdate = startDate
oTax.FRectEdate = endDate
oTax.FRectSupplyGroupID = groupid			'// �׷���̵� �ؿ� ��� �귣�� ���೻�� ǥ��
oTax.GetTaxListUpche

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"

%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<title>�� ���ݰ�꼭 ���೻��(<%= CStr(yyyy1) %>-<% CStr(mm1) %>)</title>
<table>
	<tr>
		<td>��ȣ</td>
		<td>��������</td>
		<td>�����ڻ���ڵ�Ϲ�ȣ</td>
		<td>��������ȣ</td>
		<td>��ȣ</td>
		<td>��ǥ�ڸ�</td>
		<td>���޹޴��ڻ���ڵ�Ϲ�ȣ</td>
		<td>��������ȣ</td>
		<td>��ȣ</td>
		<td>��ǥ�ڸ�</td>
		<td>�հ�ݾ�</td>
		<td>���ް���</td>
		<td>����</td>
		<td>���ڼ��ݰ�꼭�з�</td>
		<td>���ڼ��ݰ�꼭����</td>
		<td>������ �̸���</td>
		<td>���޹޴��� �̸���</td>
		<td>ǰ���</td>
	</tr>
	<% for i=0 to oTax.FResultCount - 1 %>
	<tr>
		<td><%= oTax.FTaxList(i).FtaxIdx %></td>
		<td><%= FormatDate(oTax.FTaxList(i).FisueDate,"0000-00-00") %></td>
		<td><%= oTax.FTaxList(i).FsupplyBusiNo %></td>
		<td><%= oTax.FTaxList(i).FsupplyBusiSubNo %></td>
		<td><%= oTax.FTaxList(i).FsupplyBusiName %></td>
		<td><%= oTax.FTaxList(i).FsupplyBusiCEOName %></td>
		<td><%= oTax.FTaxList(i).FBusiNo %></td>
		<td><%= oTax.FTaxList(i).FbusiSubNo %></td>
		<td><%= oTax.FTaxList(i).FBusiName %></td>
		<td><%= oTax.FTaxList(i).FbusiCEOName %></td>
		<td><%= oTax.FTaxList(i).FtotalPrice %></td>
		<td><%= oTax.FTaxList(i).FtotalPrice - oTax.FTaxList(i).FtotalTax %></td>
		<td><%= oTax.FTaxList(i).FtotalTax %></td>
		<td>���ݰ�꼭</td>
		<td>����Ź</td>
		<td><%= oTax.FTaxList(i).FsupplyRepEmail %></td>
		<td><%= oTax.FTaxList(i).FrepEmail %></td>
		<td><%= oTax.FTaxList(i).Fitemname %></td>
	</tr>
	<% next %>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
