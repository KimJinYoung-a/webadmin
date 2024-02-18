<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('세션이 종료되었습니다.');</script>"
    dbget.close()	:	response.End
end if

dim searchtype, thismonth, filetitle

searchtype =  request("searchtype")

thismonth = Left(CStr(DateSerial(year(now()),month(now())-1,1)),7)

if (searchtype = "") then
	searchtype = "thismonth"
end if



'==============================================================================
dim ojungsan

set ojungsan = new CUpcheJungsan

if (searchtype = "thismonth") then

	filetitle = "금월(" & thismonth & ") 세금계산서"
	ojungsan.FRectNotIncludeWonChon = "on"
	ojungsan.FRectYYYYMM = thismonth
	ojungsan.FRectbankingupflag = "Y"

elseif (searchtype = "prevmonth") then

	filetitle = "전월 세금계산서"
	ojungsan.FRectYYYYMM = ""
	ojungsan.FRectNotIncludeWonChon = "on"
	ojungsan.FRectNotYYYYMM = thismonth
	ojungsan.FRectbankingupflag = "Y"

elseif (searchtype = "withholding") then

	filetitle = "원천징수 대상자"
	ojungsan.FRectYYYYMM = ""
	ojungsan.FRectNotYYYYMM = ""
	ojungsan.FRectNotIncludeWonChon = ""
	ojungsan.FRectOnlyIncludeWonChon = "on"
	ojungsan.FRectbankingupflag = "Y"

else

    response.write "<script language='javascript'>alert('잘못된 접속입니다.');</script>"
    dbget.close()	:	response.End

end if


ojungsan.JungsanFixedList

dim ipsum,i
ipsum =0



'Response.ContentType = "application/unknown"
'Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=EUC-KR'>")

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" & filetitle & ".csv"
Response.CacheControl = "public"



dim bufStr, tmpS
bufStr = ""

bufStr = "은행,계좌,정산금액,업체명,사업자등록번호,(주)텐바이텐"
response.write bufStr & VbCrlf



for i=0 to ojungsan.FresultCount-1
	bufStr = ""

	if ojungsan.FItemList(i).Fipkum_bank = "홍콩샹하이" then
		bufStr = bufStr & "HSBC"
	elseif ojungsan.FItemList(i).Fipkum_bank = "단위농협" then
		bufStr = bufStr & "농협"
	elseif ojungsan.FItemList(i).Fipkum_bank = "제일" then
		bufStr = bufStr & "SC제일"
	elseif ojungsan.FItemList(i).Fipkum_bank = "시티" then
		bufStr = bufStr & "한국씨티"
	else
		bufStr = bufStr & Left(ojungsan.FItemList(i).Fipkum_bank,9)
	end if

	bufStr = bufStr & "," & Trim(ojungsan.FItemList(i).Fipkum_acctno)

	if (searchtype = "withholding") then
		bufStr = bufStr & "," & ojungsan.FItemList(i).GetTotalWithHoldingJungSanSum
	else
		bufStr = bufStr & "," & ojungsan.FItemList(i).GetTotalSuplycash
	end if
	bufStr = bufStr & "," & Replace(Left(Trim(ojungsan.FItemList(i).Fcompany_name),9), ",", " ")
	bufStr = bufStr & ","&ojungsan.FItemList(i).Fcompany_no
	bufStr = bufStr & ",(주)텐바이텐"

	response.write bufStr & VbCrlf
next

%>
<%
set ojungsan = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->