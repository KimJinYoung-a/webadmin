<%@ language=vbscript %>
<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : cs센터 네이버페이 신용카드 매출전표 조회
' History : 이상구 생성
'			2023.08.08 한용민 수정(해당 기능 막음.네이버페이 신용카드 매출전표 조회 api 2021년5월7일부로 종료)
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
response.write "	alert('[네이버페이]신용카드 매출전표 조회 api는 2021년5월7일부로 종료 되었습니다.\n고객님께서 직접 네이버페이 홈페이지에 접속 하셔서 조회하셔야 합니다.');"
response.write "	self.close();"
response.write "</script>"
dbget.close() : response.end

''if (isNaverPayReceiptSvcValid) then
    rURi = fnCallNaverPayReceipt(tid)
    
    if left(rURi,4)="ERR:" then
    	response.write "<script>alert('처리중 오류가 발생했습니다.\n(" & right(rURi,len(rURi)-4) & ")');</script>"
    	response.end
    end if

    ''### 2. 조회 화면 호출
    ''Response.Redirect URLDecodeUTF8(rURi)
''end if

'// URL Decoding
Public Function URLDecodeUTF8(byVal pURL)
Dim i, s1, s2, s3, u1, u2, result
pURL = Replace(pURL,"+"," ")

For i = 1 to Len(pURL)
	if Mid(pURL, i, 1) = "%" then
		s1 = CLng("&H" & Mid(pURL, i + 1, 2))

        '1바이트일 경우
        If CInt("&H" & Mid(pURL, i + 1, 2)) < 128 Then
            result = result & Chr(CInt("&H" & Mid(pURL, i + 1, 2)))
            i = i + 2 ' 잘라낸 만큼 뒤로 이동

		'2바이트일 경우
		elseif ((s1 AND &HC0) = &HC0) AND ((s1 AND &HE0) <> &HE0) then
			s2 = CLng("&H" & Mid(pURL, i + 4, 2))

			u1 = (s1 AND &H1C) / &H04
			u2 = ((s1 AND &H03) * &H04 + ((s2 AND &H30) / &H10)) * &H10
			u2 = u2 + (s2 AND &H0F)
			result = result & ChrW((u1 * &H100) + u2)
			i = i + 5

		'3바이트일 경우
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
body, tr, td {font-size:9pt; font-family:굴림,verdana; color:#433F37; line-height:19px;}
table, img {border:none}

/* Padding ******/
.pl_01 {padding:1 10 0 10; line-height:19px;}
.pl_03 {font-size:20pt; font-family:굴림,verdana; color:#FFFFFF; line-height:29px;}

/* Link ******/
.a:link  {font-size:9pt; color:#333333; text-decoration:none}
.a:visited { font-size:9pt; color:#333333; text-decoration:none}
.a:hover  {font-size:9pt; color:#0174CD; text-decoration:underline}

.txt_03a:link  {font-size: 8pt;line-height:18px;color:#333333; text-decoration:none}
.txt_03a:visited {font-size: 8pt;line-height:18px;color:#333333; text-decoration:none}
.txt_03a:hover  {font-size: 8pt;line-height:18px;color:#EC5900; text-decoration:underline}

.buttoncss {
	font-family: "Verdana", "돋움";
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
	<!---- 팝업제목 시작 ---->
	<td valign="top" bgcolor="#af1414"></td>
	<!---- 팝업제목 끝 ---->
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
					네이버 페이 이슈로 인해 증빙서류 출력이 안될경우 네이버페이 쪽으로 안내해 주세요.
					<br>
					<a href="https://pay.naver.com" target="_blank">https://pay.naver.com</a>
					<br>
					<br>
					<a href="<%=URLDecodeUTF8(rURi)%>" target="_blank">전표조회</a>
				</td>
			</tr>
		</table>
	</td>
</tr>
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp" -->
