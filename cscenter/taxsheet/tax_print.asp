<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<%
	'// MS Word���� ���� ����
	Response.Buffer=true
	Response.Expires=0
	Response.ContentType = "application/msword" 

	'// ���� ���� //
	dim chkPrint
	dim oTax, i
	dim curName, curPosit, curPart

	'// �Ķ���� ���� //
	chkPrint = request("chkSelect")

	'// ���� ����
	set oTax = new CTaxPrint
	oTax.FRectChkPrint = chkPrint

	oTax.GetTaxPrint

%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Word.Document>
<!--[if gte mso 9]><xml>
 <w:WordDocument>
  <w:View>Print</w:View>
  <w:GrammarState>Clean</w:GrammarState>
  <w:ValidateAgainstSchemas/>
  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
  <w:Compatibility>
   <w:UseFELayout/>
  </w:Compatibility>
  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>
 </w:WordDocument>
</xml><![endif]-->
<style>
<!--
@page Section1
	{size:595.3pt 841.9pt;
	margin:2.0cm 2.0cm 2.0cm 2.0cm;
	mso-header-margin:1.0cm;
	mso-footer-margin:1.0cm;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}

.border_1_All	{border:1px solid #060606;}
.border_1_TL	{border-top:1px solid #060606;border-left:1px solid #060606;}
.border_1_TLR	{border-top:1px solid #060606;border-left:1px solid #060606;border-right:1px solid #060606;}
.border_1_TLB	{border-top:1px solid #060606;border-left:1px solid #060606;border-bottom:1px solid #060606;}

.border_2_TL	{border-top:2px solid #060606;border-left:1px solid #060606;}
.border_2_T		{border-top:2px solid #060606;}
.border_2_TB	{border-top:1px solid #060606;border-bottom:2px solid #060606;}
.border_2_TLB	{border-top:1px solid #060606;border-left:1px solid #060606;border-bottom:2px solid #060606;}
-->
</style>
</head>
<body>
<div class=Section1>
<%
	'// ���ڵ�� ��ŭ ��� //
	for i=0 to (oTax.FTotalCount-1)

		'����� ���� ǥ��
		Select Case oTax.FTaxList(i).FcurUserId
			Case "kobula"
				curName		= "������"
				curPosit	= "�븮"
				curPart		= "�������"
			Case "icommang"
				curName		= "������"
				curPosit	= "����"
				curPart		= "�������"
			Case Else
				curName		= "������"
				curPosit	= "����"
				curPart		= "�¶��λ����"
		End Select
%>
<table width="640" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td>
		<table width="640" cellpadding="0" cellspacing="0" border="0">
		<tr height="24">
			<td width="400" rowspan="2"><img src="http://fiximage.10x10.co.kr/topimg/top_logo.gif" width="167"><br><span style="font-size:24pt;"><b>���ݰ�꼭�����û��</b></span></td>
			<td width="80" align="center" class="border_1_TL">�� �� ��</td>
			<td width="80" align="center" class="border_1_TL">�� ��</td>
			<td width="80" align="center" class="border_1_TLR" bgcolor="#F6F6F6">�� ��</td>
		</tr>
		<tr height="90" bgcolor="#FFFFFF">
			<td class="border_1_TLB">&nbsp;</td>
			<td class="border_1_TLB">&nbsp;</td>
			<td class="border_1_All" bgcolor="#F6F6F6">&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
	<td>
		<table width="640" cellpadding="4" cellspacing="0" border="0">
		<tr bgcolor="#FFFFFF">
			<td width="120" align="center" class="border_2_T">�� �� ��</td>
			<td width="200" align="center" class="border_2_TL"><%=Year(date) & "�� " & Month(date) & "�� " & Day(date) & "��"%></td>
			<td width="120" align="center" class="border_2_TL">�� ��</td>
			<td width="200" align="center" class="border_2_TL"><%=curPosit%></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td class="border_2_TB" align="center">�� �� ��</td>
			<td class="border_2_TLB" align="center"><%=curName%></td>
			<td class="border_2_TLB" align="center">�ҼӺμ�</td>
			<td class="border_2_TLB" align="center"><%=curPart%></td>
		</tr>
		</table>
	</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
	<td align="center">
		<table width="600" cellpadding="3" cellspacing="0" border="0" style="font-size:10pt;">
		<tr><td colspan="4"><b>&lt;���೻��&gt;</b></td></tr>
		<tr>
			<td width="120" class="border_1_TL" align="center">������(�Ա���)</td>
			<td width="180" class="border_1_TL" align="left"><%=Year(oTax.FTaxList(i).Fipkumdate) & "�� " & Month(oTax.FTaxList(i).Fipkumdate) & "�� " & Day(oTax.FTaxList(i).Fipkumdate) & "��"%></td>
			<td width="120" class="border_1_TL" align="center">����߼ۿ���</td>
			<td width="180" class="border_1_TLR" align="left"><%=db2html(oTax.FTaxList(i).FrepEmail)%></td>
		</tr>
		<tr>
			<td class="border_1_TL" align="center">�ŷ�ó��</td>
			<td class="border_1_TL" align="left"><%=db2html(oTax.FTaxList(i).FbusiName)%></td>
			<td class="border_1_TL" align="center">����ڵ�Ϲ�ȣ</td>
			<td class="border_1_TLR" align="left"><%=db2html(oTax.FTaxList(i).FbusiNo)%></td>
		</tr>
		<tr>
			<td class="border_1_TL" align="center">�� �� ��</td>
			<td class="border_1_TL" align="left"><%=db2html(oTax.FTaxList(i).FrepName)%></td>
			<td class="border_1_TL" align="center">�� �� ó</td>
			<td class="border_1_TLR" align="left"><%=db2html(oTax.FTaxList(i).FrepTel)%></td>
		</tr>
		<tr>
			<td class="border_1_TL" align="center">�� ��</td>
			<td class="border_1_TL" align="left"><%=FormatNumber(oTax.FTaxList(i).FtotalPrice,0)%>��</td>
			<td class="border_1_TL" align="center">ǰ ��</td>
			<td class="border_1_TLR" align="left"><%=db2html(oTax.FTaxList(i).Fitemname)%></td>
		</tr>
		<tr>
			<td class="border_1_TL" align="center">�� �� ��</td>
			<td class="border_1_TL" align="left"><%=db2html(oTax.FTaxList(i).Fbuyname)%></td>
			<td class="border_1_TL" align="center">�ֹ���ȣ</td>
			<td class="border_1_TLR" align="left"><%=db2html(oTax.FTaxList(i).Forderserial)%></td>
		</tr>
		<tr>
			<td width="120" class="border_1_TLB" align="center">�� ��</td>
			<td width="480" class="border_1_All" colspan="3" align="left"><%=db2html(oTax.FTaxList(i).FbusiAddr)%></td>
		</tr>
		</table>
	</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
	<td align="center">
		<table width="600" cellpadding="3" cellspacing="0" border="0" style="font-size:10pt;">
		<tr><td><b>&lt;÷�μ���&gt;</b></td></tr>
		<tr>
			<td height="70" class="border_1_All" valign="top">
				1. ����ڵ���� �纻<br>
				2. ��Ÿ
			</td>
		</tr>
		<tr>
			<td>
				1) �ݾ��� ���ް��װ� ������ ������ �ݾ��� ����<br>
				2) ���� �� ����߼��� �ʿ�ÿ� �ݵ�� ����ڿ� ����ó ����<br>
				3) �¶��� �ǸŽô� �ֹ��ڿ� �ֹ���ȣ �ݵ�� ����
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
	<td align="center">
		<table width="600" cellpadding="3" cellspacing="0" border="0" style="font-size:10pt;">
		<tr><td><b>&lt;ó�� ���&gt;</b></td></tr>
		<tr>
			<td height="70" class="border_1_All">&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
	<td>
		<b>(��)�ٹ����� <u>www.10x10.co.kr</u></b><br>
		(03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ����� Tel.02-554-2033 | Fax.02-2179-9245
	</td>
</tr>
</table>
<%
		'# ���� �������� �ִٸ� ������ ���� ���
		if i<(oTax.FTotalCount-1) then
%>
<!-- ������ ���� ���� -->
<span lang=EN-US style='font-size:12.0pt;font-family:����;mso-bidi-font-family:
����;mso-ansi-language:EN-US;mso-fareast-language:KO;mso-bidi-language:AR-SA'><br
clear=all style='page-break-before:always'>
</span>
<%
		end if
	next

	set oTax = Nothing
%>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->

