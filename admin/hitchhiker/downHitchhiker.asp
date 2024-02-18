<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	PageName 	: /admin/hitchhiker/downHitchhiker.asp
'	Description : ��ġ����Ŀ ��ûȸ������Ʈ �ٿ�
'	History		: 2006.11.30 ������ ����
'				  2019.11.11 �ѿ�� ����
'#############################################################

'// MS Word���� ���� ����
Response.Expires=-1440
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-disposition","attachment;filename=hitchhiker_V"&Request.Querystring("iHV")&"_"&Request.Querystring("iAV")& ".xls"
Response.Buffer = true    '���ۻ�뿩��
%>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhikerCls.asp"-->
<%
Dim iHVol, iAVol, iType
Dim clsHList, arrList, intLoop
Dim startDate, endDate
	iHVol = Request.Querystring("iHV")
	iAVol = Request.Querystring("iAV")
	iType = Request.Querystring("iType")
	startDate	= Request("startDate")
	endDate	= Request("endDate")

set clsHList =  new Chitchhiker
	clsHList.FCPage = 1
	clsHList.FPSize = 30000
	clsHList.FHVol = iHVol
	clsHList.FAVol = iAVol
	clsHList.FisSend = iType
	clsHList.FSDate = startDate		'�˻�������
	clsHList.FEDate = endDate		'�˻�������
	arrList = clsHList.fnGetList()
set clsHList = nothing
%>
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
	<style type="text/css">
		td {font-family:����;font-size:9pt; mso-number-format:\@;}
	</style>
</head>
<body>
	<table border=1 cellpadding=3 cellspacing=0>
	<tr bgcolor="#f4f4f4">
		<td align="center">��ȣ</td>
		<td align="center">���̵�</td>
		<td align="center">�̸�</td>
		<td align="center">������</td>
		<td align="center">�����ȣ</td>
		<td align="center">�ּ�</td>
		<td align="center">���ּ�</td>
		<td align="center">��ȭ��ȣ</td>
		<td align="center">�ڵ�����ȣ</td>
		<td align="center">��û��</td>
		<td align="center">�߼���</td>
		<td align="center">LV</td>
		<td align="center">��û����</td>
	</tr>
	<%
	IF isArray(arrList) THEN
		FOR intLoop =0 TO UBound(arrList,2)
	%>
		<tr>
			<td  align="center"><%=intLoop+1%></td>
			<td><%=arrList(3,intLoop)%></td>
			<td><%=arrList(4,intLoop)%></td>
			<td><%=arrList(12,intLoop)%></td>
			<td><%=arrList(5,intLoop)%></td>
			<td><%=arrList(6,intLoop)%></td>
			<td><%=db2html(arrList(7,intLoop))%></td>
			<td><%=arrList(8,intLoop)%></td>
			<td><%=arrList(9,intLoop)%></td>
			<td><%=arrList(2,intLoop)%></td>
			<td><%=arrList(10,intLoop)%></td>
			<td><%= getUserLevelStr(arrList(11,intLoop)) %></td>
			<td><%=arrList(14,intLoop)%></td>
		</tr>
	<%
		if intLoop mod 3000 = 0 then
			Response.Flush		' ���۸��÷���
		end if
		NEXT
	ELSE
	%>
		<tr>
			<td colspan="9">��ϵ� ������ �����ϴ�.</td>
		</tr>
	<%END IF%>
	</table>
</body>
</html>

<!-- #include virtual="/lib/db/dbclose.asp" -->