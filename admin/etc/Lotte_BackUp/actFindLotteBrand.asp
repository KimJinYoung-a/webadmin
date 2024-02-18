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
	'// 변수선언
	dim lottenBrandCD, lotteBrandName
	dim srcStr, rstCnt, BrnInfo
	srcStr = Trim(Request("brnNm"))

	if srcStr="" then
		Call Alert_Close("검색어가 없습니다.")
		Response.End
	end if

	'// 롯데닷컴 브램드 조회
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", lotteAPIURL & "/openapi/searchBrandListOpenApi.lotte?subscriptionId=" & lotteAuthNo & "&brnd_nm=" & Server.URLEncode(srcStr), false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()
	If objXML.Status = "200" Then

		'//전달받은 내용 확인
		'Response.contentType = "text/xml; charset=euc-kr"
		'response.write BinaryToText(objXML.ResponseBody, "euc-kr")

		'XML을 담을 DOM 객체 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		
		on Error Resume Next
			rstCnt = xmlDOM.getElementsByTagName("BrandCount").item(0).text		'결과수
			if Err<>0 then
				Call Alert_Close("롯데닷컴 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요.")
				Response.End
			end if

			if rstCnt>0 then
				'//결과창 표시
				Response.Write	"<table width='100%' border='0' cellpadding='2' cellspacing='1' class='a' bgcolor='#BABABA'>"
				Response.Write	"<tr align='center'>"
				Response.Write	"	<td bgcolor='#DDDDFF'>브랜드코드</td>"
				Response.Write	"	<td bgcolor='#DDDDFF'>브랜드명</td>"
				Response.Write	"	<td bgcolor='#DDDDFF'>선택</td>"
				Response.Write	"</tr>"

				'// BrnInfo Loop
				Set BrnInfo = xmlDOM.getElementsByTagName("BrandInfo")
				for each SubNodes in BrnInfo
					lottenBrandCD	= Trim(SubNodes.getElementsByTagName("BrandCode").item(0).text)		'브랜드코드
					lotteBrandName	= Trim(SubNodes.getElementsByTagName("BrandName").item(0).text)		'브랜드명(한글)

					Response.Write	"<tr align='center'>"
					Response.Write	"	<td bgcolor='#FFFFFF'>" & lottenBrandCD & "</td>"
					Response.Write	"	<td bgcolor='#FFFFFF'>" & lotteBrandName & "</td>"
					Response.Write	"	<td bgcolor='#FFFFFF'><input type='button' value='선택' onClick=""putSelBrnCd('" & lottenBrandCD & "','" & lotteBrandName & "')"" class='button'></td>"
					Response.Write	"</tr>"
				Next
				Set BrnInfo = Nothing
				
				Response.Write	"</table>"
			else
				Call Alert_Close("검색 결과가 없습니다.\검색어를 확인하시 후 다시 검색해주세요.")
				Response.End
			end if
		on Error Goto 0

		Set xmlDOM = Nothing
	else
		Call Alert_Close("롯데닷컴과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요.")
		Response.End
	end if
	Set objXML = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->