<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<!-- #include virtual="/admin/upchejungsan/upchejungsan_function.asp"-->
<%
If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('세션이 종료되었습니다.');</script>"
    dbget.close()	:	response.End
end if

dim ipFileNo, xltype
ipFileNo =  requestCheckvar(request("ipFileNo"),10)
xltype =  requestCheckvar(request("xltype"),10)

Dim ipFileRegdate,ipFileName, isWonChonFile : isWonChonFile=false
Dim sqlStr
sqlStr = "select M.ipFileNo,M.ipFileName,M.ipFileRegdate, M.ipFileGbn"
sqlStr = sqlStr & " From db_jungsan.dbo.tbl_jungsan_ipkumFile_Master M"
sqlStr = sqlStr & " where ipFileNo="&ipFileNo
rsget.Open sqlStr,dbget,1
IF Not rsget.Eof THEN
    ipFileName = rsget("ipFileName")
    ipFileRegdate = rsget("ipFileRegdate")
    isWonChonFile = (rsget("ipFileGbn")="WN")
ENd IF
rsget.Close


Dim arrDetailList
arrDetailList = fnGetJFixIpkumListSum(ipFileNo)

if (xltype="1") then
    ipFileName = ipFileNo &"_이체_"&ipFileName
ELSEif (xltype="2") then
    ipFileName = ipFileNo &"_예금주_"&ipFileName
END IF
'Response.ContentType = "application/unknown"
'Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=EUC-KR'>")

''Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" & ipFileName & ".xls"
Response.CacheControl = "public"

response.write "<style type='text/css'>" & VbCrlf
response.write ".txt {mso-number-format:'\@'}" & VbCrlf
response.write "</style>" & VbCrlf

dim bufStr
bufStr = ""

bufStr = "<table>" ''<tr><td>거래처명</td><td>입금은행</td><td>입금계좌</td><td>이체금액</td><td>출금통장인쇄내용</td><td>입금통장인쇄내용</td></tr>"

if (xltype="2") then
    bufStr = "<table>" ''<tr><td>예금주</td><td>입금은행</td><td>입금계좌</td><td>이체금액</td><td>출금통장인쇄내용</td><td>입금통장인쇄내용</td></tr>"
end if

response.write bufStr & VbCrlf
dim intLoop
For intLoop = 0 To UBound(arrDetailList,2)
	bufStr = ""

    if (xltype="2") then
	bufStr = bufStr &"<tr><td class='txt'>"& arrDetailList(6,intLoop) 
	else
	bufStr = bufStr &"</td><td class='txt'>"& arrDetailList(5,intLoop) 
    end if

    bufStr = bufStr &"</td><td class='txt'>"& arrDetailList(1,intLoop) 
    bufStr = bufStr &"</td><td class='txt'>"& arrDetailList(2,intLoop)
    IF (isWonChonFile) then
        bufStr = bufStr &"</td><td >"& GetHoldingJungSanSum(arrDetailList(3,intLoop) )
    ELSE 
        bufStr = bufStr &"</td><td >"& arrDetailList(3,intLoop) 
    ENd IF
    bufStr = bufStr &"</td><td class='txt'>"& arrDetailList(5,intLoop) 
    bufStr = bufStr &"</td><td class='txt'>"& "(주)텐바이텐</td></tr>"
    		

	response.write bufStr & VbCrlf
next
response.write "</table>" & VbCrlf
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->