<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<script language="javascript">
<!--
function putSelBrnCd(bcd,bnm) {
	opener.frm.lotteBrandCd.value=bcd;
	opener.frm.lotteBrandNm.value=bnm;
	opener.document.getElementById("brTT").rowSpan=2;
	opener.document.getElementById("BrRow").style.display="";
	opener.document.getElementById("selBr").innerHTML="[" + bcd + "] " + bnm;
	self.close();
}
//-->
</script>
<%
	'// ��������
	dim lottenBrandCD, lotteBrandName
	dim srcStr, rstCnt, BrnInfo
	srcStr = Trim(Request("brnNm"))

	if srcStr="" then
		Call Alert_Close("�˻�� �����ϴ�.")
		Response.End
	end if

	'// �Ե����� �귥�� ��ȸ
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", lotteAPIURL & "/openapi/searchBrandListOpenApi.lotte?subscriptionId=" & lotteAuthNo & "&brnd_nm=" & Server.URLEncode(srcStr), false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()
	If objXML.Status = "200" Then

		'//���޹��� ���� Ȯ��
		'Response.contentType = "text/xml; charset=euc-kr"
		'response.write BinaryToText(objXML.ResponseBody, "euc-kr")

		'XML�� ���� DOM ��ü ����
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		
		on Error Resume Next
			rstCnt = xmlDOM.getElementsByTagName("BrandCount").item(0).text		'�����
			if Err<>0 then
				Call Alert_Close("�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����.")
				Response.End
			end if

			if rstCnt>0 then
				'//���â ǥ��
				Response.Write	"<table width='100%' border='0' cellpadding='2' cellspacing='1' class='a' bgcolor='#BABABA'>"
				Response.Write	"<tr align='center'>"
				Response.Write	"	<td bgcolor='#DDDDFF'>�귣���ڵ�</td>"
				Response.Write	"	<td bgcolor='#DDDDFF'>�귣���</td>"
				Response.Write	"	<td bgcolor='#DDDDFF'>����</td>"
				Response.Write	"</tr>"

				'// BrnInfo Loop
				Set BrnInfo = xmlDOM.getElementsByTagName("BrandInfo")
				for each SubNodes in BrnInfo
					lottenBrandCD	= Trim(SubNodes.getElementsByTagName("BrandCode").item(0).text)		'�귣���ڵ�
					lotteBrandName	= Trim(SubNodes.getElementsByTagName("BrandName").item(0).text)		'�귣���(�ѱ�)

					Response.Write	"<tr align='center'>"
					Response.Write	"	<td bgcolor='#FFFFFF'>" & lottenBrandCD & "</td>"
					Response.Write	"	<td bgcolor='#FFFFFF'>" & lotteBrandName & "</td>"
					Response.Write	"	<td bgcolor='#FFFFFF'><input type='button' value='����' onClick=""putSelBrnCd('" & lottenBrandCD & "','" & lotteBrandName & "')"" class='button'></td>"
					Response.Write	"</tr>"
				Next
				Set BrnInfo = Nothing
				
				Response.Write	"</table>"
			else
				Call Alert_Close("�˻� ����� �����ϴ�.\�˻�� Ȯ���Ͻ� �� �ٽ� �˻����ּ���.")
				Response.End
			end if
		on Error Goto 0

		Set xmlDOM = Nothing
	else
		Call Alert_Close("�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����.")
		Response.End
	end if
	Set objXML = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->