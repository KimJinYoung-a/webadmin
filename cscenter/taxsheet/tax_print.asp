<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_taxsheetcls.asp"-->
<%
	'// MS Word파일 구성 선언
	Response.Buffer=true
	Response.Expires=0
	Response.ContentType = "application/msword" 

	'// 변수 선언 //
	dim chkPrint
	dim oTax, i
	dim curName, curPosit, curPart

	'// 파라메터 접수 //
	chkPrint = request("chkSelect")

	'// 내용 접수
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
	'// 레코드수 만큼 출력 //
	for i=0 to (oTax.FTotalCount-1)

		'담당자 정보 표시
		Select Case oTax.FTaxList(i).FcurUserId
			Case "kobula"
				curName		= "허진원"
				curPosit	= "대리"
				curPart		= "운영관리팀"
			Case "icommang"
				curName		= "서동석"
				curPosit	= "팀장"
				curPart		= "운영관리팀"
			Case Else
				curName		= "성민정"
				curPosit	= "주임"
				curPart		= "온라인사업팀"
		End Select
%>
<table width="640" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td>
		<table width="640" cellpadding="0" cellspacing="0" border="0">
		<tr height="24">
			<td width="400" rowspan="2"><img src="http://fiximage.10x10.co.kr/topimg/top_logo.gif" width="167"><br><span style="font-size:24pt;"><b>세금계산서발행요청서</b></span></td>
			<td width="80" align="center" class="border_1_TL">작 성 자</td>
			<td width="80" align="center" class="border_1_TL">팀 장</td>
			<td width="80" align="center" class="border_1_TLR" bgcolor="#F6F6F6">전 결</td>
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
			<td width="120" align="center" class="border_2_T">작 성 일</td>
			<td width="200" align="center" class="border_2_TL"><%=Year(date) & "년 " & Month(date) & "월 " & Day(date) & "일"%></td>
			<td width="120" align="center" class="border_2_TL">직 급</td>
			<td width="200" align="center" class="border_2_TL"><%=curPosit%></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td class="border_2_TB" align="center">작 성 자</td>
			<td class="border_2_TLB" align="center"><%=curName%></td>
			<td class="border_2_TLB" align="center">소속부서</td>
			<td class="border_2_TLB" align="center"><%=curPart%></td>
		</tr>
		</table>
	</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
	<td align="center">
		<table width="600" cellpadding="3" cellspacing="0" border="0" style="font-size:10pt;">
		<tr><td colspan="4"><b>&lt;발행내역&gt;</b></td></tr>
		<tr>
			<td width="120" class="border_1_TL" align="center">발행일(입금일)</td>
			<td width="180" class="border_1_TL" align="left"><%=Year(oTax.FTaxList(i).Fipkumdate) & "년 " & Month(oTax.FTaxList(i).Fipkumdate) & "월 " & Day(oTax.FTaxList(i).Fipkumdate) & "일"%></td>
			<td width="120" class="border_1_TL" align="center">우편발송여부</td>
			<td width="180" class="border_1_TLR" align="left"><%=db2html(oTax.FTaxList(i).FrepEmail)%></td>
		</tr>
		<tr>
			<td class="border_1_TL" align="center">거래처명</td>
			<td class="border_1_TL" align="left"><%=db2html(oTax.FTaxList(i).FbusiName)%></td>
			<td class="border_1_TL" align="center">사업자등록번호</td>
			<td class="border_1_TLR" align="left"><%=db2html(oTax.FTaxList(i).FbusiNo)%></td>
		</tr>
		<tr>
			<td class="border_1_TL" align="center">담 당 자</td>
			<td class="border_1_TL" align="left"><%=db2html(oTax.FTaxList(i).FrepName)%></td>
			<td class="border_1_TL" align="center">연 락 처</td>
			<td class="border_1_TLR" align="left"><%=db2html(oTax.FTaxList(i).FrepTel)%></td>
		</tr>
		<tr>
			<td class="border_1_TL" align="center">금 액</td>
			<td class="border_1_TL" align="left"><%=FormatNumber(oTax.FTaxList(i).FtotalPrice,0)%>원</td>
			<td class="border_1_TL" align="center">품 목</td>
			<td class="border_1_TLR" align="left"><%=db2html(oTax.FTaxList(i).Fitemname)%></td>
		</tr>
		<tr>
			<td class="border_1_TL" align="center">주 문 자</td>
			<td class="border_1_TL" align="left"><%=db2html(oTax.FTaxList(i).Fbuyname)%></td>
			<td class="border_1_TL" align="center">주문번호</td>
			<td class="border_1_TLR" align="left"><%=db2html(oTax.FTaxList(i).Forderserial)%></td>
		</tr>
		<tr>
			<td width="120" class="border_1_TLB" align="center">주 소</td>
			<td width="480" class="border_1_All" colspan="3" align="left"><%=db2html(oTax.FTaxList(i).FbusiAddr)%></td>
		</tr>
		</table>
	</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
	<td align="center">
		<table width="600" cellpadding="3" cellspacing="0" border="0" style="font-size:10pt;">
		<tr><td><b>&lt;첨부서류&gt;</b></td></tr>
		<tr>
			<td height="70" class="border_1_All" valign="top">
				1. 사업자등록증 사본<br>
				2. 기타
			</td>
		</tr>
		<tr>
			<td>
				1) 금액은 공급가액과 세액을 포함한 금액을 기재<br>
				2) 발행 후 우편발송이 필요시엔 반드시 담당자와 연락처 기재<br>
				3) 온라인 판매시는 주문자와 주문번호 반드시 기재
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
	<td align="center">
		<table width="600" cellpadding="3" cellspacing="0" border="0" style="font-size:10pt;">
		<tr><td><b>&lt;처리 결과&gt;</b></td></tr>
		<tr>
			<td height="70" class="border_1_All">&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr>
	<td>
		<b>(주)텐바이텐 <u>www.10x10.co.kr</u></b><br>
		(03082) 서울시 종로구 대학로 57 홍익대학교 대학로캠퍼스 교육동 14층 텐바이텐 Tel.02-554-2033 | Fax.02-2179-9245
	</td>
</tr>
</table>
<%
		'# 다음 페이지가 있다면 페이지 나눔 출력
		if i<(oTax.FTotalCount-1) then
%>
<!-- 페이지 강제 나눔 -->
<span lang=EN-US style='font-size:12.0pt;font-family:굴림;mso-bidi-font-family:
굴림;mso-ansi-language:EN-US;mso-fareast-language:KO;mso-bidi-language:AR-SA'><br
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

