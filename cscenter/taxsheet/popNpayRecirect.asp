<%@ language=vbscript %>
<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : cs���� ���̹����� �ſ�ī�� ������ǥ ��ȸ
' History : �̻� ����
'			2023.08.08 �ѿ�� ����(�ش� ��� ����.���̹����� �ſ�ī�� ������ǥ ��ȸ api 2021��5��7�Ϻη� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/cscenter/action/incNaverpayCommon.asp"-->
<%
Dim orderserial, tid, i, rURi, isNaverPayReceiptSvcValid
	orderserial=requestCheckvar(request("orderserial"),20)
	tid=requestCheckvar(request("tid"),100)

isNaverPayReceiptSvcValid=FALSE

response.write "<script type='text/javascript'>"
response.write "	alert('[���̹�����]�ſ�ī�� ������ǥ ��ȸ api�� 2021��5��7�Ϻη� ���� �Ǿ����ϴ�.\n���Բ��� ���� ���̹����� Ȩ�������� ���� �ϼż� ��ȸ�ϼž� �մϴ�.');"
response.write "	self.close();"
response.write "</script>"
dbget.close() : response.end

''if (isNaverPayReceiptSvcValid) then
    rURi = fnCallNaverPayReceipt(tid)
    
    if left(rURi,4)="ERR:" then
    	response.write "<script>alert('ó���� ������ �߻��߽��ϴ�.\n(" & right(rURi,len(rURi)-4) & ")');</script>"
    	response.end
    end if

    ''### 2. ��ȸ ȭ�� ȣ��
    ''Response.Redirect URLDecodeUTF8(rURi)
''end if

'// URL Decoding
Public Function URLDecodeUTF8(byVal pURL)
Dim i, s1, s2, s3, u1, u2, result
pURL = Replace(pURL,"+"," ")

For i = 1 to Len(pURL)
	if Mid(pURL, i, 1) = "%" then
		s1 = CLng("&H" & Mid(pURL, i + 1, 2))

        '1����Ʈ�� ���
        If CInt("&H" & Mid(pURL, i + 1, 2)) < 128 Then
            result = result & Chr(CInt("&H" & Mid(pURL, i + 1, 2)))
            i = i + 2 ' �߶� ��ŭ �ڷ� �̵�

		'2����Ʈ�� ���
		elseif ((s1 AND &HC0) = &HC0) AND ((s1 AND &HE0) <> &HE0) then
			s2 = CLng("&H" & Mid(pURL, i + 4, 2))

			u1 = (s1 AND &H1C) / &H04
			u2 = ((s1 AND &H03) * &H04 + ((s2 AND &H30) / &H10)) * &H10
			u2 = u2 + (s2 AND &H0F)
			result = result & ChrW((u1 * &H100) + u2)
			i = i + 5

		'3����Ʈ�� ���
		elseif (s1 AND &HE0 = &HE0) then
			s2 = CLng("&H" & Mid(pURL, i + 4, 2))
			s3 = CLng("&H" & Mid(pURL, i + 7, 2))

			u1 = ((s1 AND &H0F) * &H10)
			u1 = u1 + ((s2 AND &H3C) / &H04)
			u2 = ((s2 AND &H03) * &H04 +  (s3 AND &H30) / &H10) * &H10
			u2 = u2 + (s3 AND &H0F)
			result = result & ChrW((u1 * &H100) + u2)
			i = i + 8
		end if
	else
		result = result & Mid(pURL, i, 1)
	end if
Next
URLDecodeUTF8 = result
End Function

%>
<style>
body, tr, td {font-size:9pt; font-family:����,verdana; color:#433F37; line-height:19px;}
table, img {border:none}

/* Padding ******/
.pl_01 {padding:1 10 0 10; line-height:19px;}
.pl_03 {font-size:20pt; font-family:����,verdana; color:#FFFFFF; line-height:29px;}

/* Link ******/
.a:link  {font-size:9pt; color:#333333; text-decoration:none}
.a:visited { font-size:9pt; color:#333333; text-decoration:none}
.a:hover  {font-size:9pt; color:#0174CD; text-decoration:underline}

.txt_03a:link  {font-size: 8pt;line-height:18px;color:#333333; text-decoration:none}
.txt_03a:visited {font-size: 8pt;line-height:18px;color:#333333; text-decoration:none}
.txt_03a:hover  {font-size: 8pt;line-height:18px;color:#EC5900; text-decoration:underline}

.buttoncss {
	font-family: "Verdana", "����";
	font-size: 9pt;
	background-color: #E6E6E6;
	border: 1px outset #BABABA;
	color: #000000;
	height: 20px;
	cursor:hand;
}

</style>
<body bgcolor="#FFFFFF" text="#242424" leftmargin=0 topmargin=0 marginwidth=0 marginheight=0 bottommargin=0 rightmargin=0><center>

<table width="650" border="0" cellspacing="0" cellpadding="0">
<tr>
	<!---- �˾����� ���� ---->
	<td valign="top" bgcolor="#af1414"></td>
	<!---- �˾����� �� ---->
</tr>
<tr>
	<td style="padding:0px 15px">

		<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td style="padding:25px 0 7px 0;">
					<b>Notics</b>
				</td>
			</tr>
			<tr>
				<td>
					���̹� ���� �̽��� ���� �������� ����� �ȵɰ�� ���̹����� ������ �ȳ��� �ּ���.
					<br>
					<a href="https://pay.naver.com" target="_blank">https://pay.naver.com</a>
					<br>
					<br>
					<a href="<%=URLDecodeUTF8(rURi)%>" target="_blank">��ǥ��ȸ</a>
				</td>
			</tr>
		</table>
	</td>
</tr>
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp" -->
