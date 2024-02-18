<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 40
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_refundcheckcls.asp"-->
<%

Const MaxPage   = 1
Const PageSize = 5000

dim research, page, i
dim divcd, returnmethod, orderserial, chkGubun, refundMin, refundMax
dim yyyy1, yyyy2, mm1, mm2, dd1, dd2
dim fromDate, toDate
dim exCheckFinish
dim returnmethodIN, retR007, retR910, retR900, dategbn

'===============================================================================
research 		= requestCheckVar(request("research"),32)
page 			= requestCheckVar(request("page"),32)
divcd 			= requestCheckVar(request("divcd"),32)
returnmethod 	= requestCheckVar(request("returnmethod"),32)
orderserial 	= requestCheckVar(request("orderserial"),32)
chkGubun 		= requestCheckVar(request("chkGubun"),32)
refundMin 		= requestCheckVar(request("refundMin"),32)
refundMax 		= requestCheckVar(request("refundMax"),32)
exCheckFinish 	= requestCheckVar(request("exCheckFinish"),32)
retR007 		= requestCheckVar(request("retR007"),32)
retR910 		= requestCheckVar(request("retR910"),32)
retR900 		= requestCheckVar(request("retR900"),32)
dategbn     = requestCheckvar(request("dategbn"),32)
'===============================================================================
yyyy1   = request("yyyy1")
yyyy2   = request("yyyy2")
mm1     = request("mm1")
mm2     = request("mm2")
dd1     = request("dd1")
dd2     = request("dd2")

if (yyyy1="") then
	fromDate = CStr(DateSerial(Year(Now()), (Month(Now()) - 1), 1))
	toDate = CStr(DateSerial(Year(Now()), Month(Now()), 0))

    yyyy1 = CStr(Year(fromDate))
    mm1 = CStr(Month(fromDate))
    dd1 =  CStr(day(fromDate))

    yyyy2 = CStr(Year(toDate))
    mm2 = CStr(Month(toDate))
    dd2 =  CStr(day(toDate))
end if

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if dategbn="" then dategbn="finishdate"

if (retR007 <> "") or (retR910 <> "") or (retR900 <> "") then
	returnmethodIN = "'XXXX'"
	if (retR007 <> "") then
		returnmethodIN = returnmethodIN + ",'R007'"
	end if
	if (retR910 <> "") then
		returnmethodIN = returnmethodIN + ",'R910'"
	end if
	if (retR900 <> "") then
		returnmethodIN = returnmethodIN + ",'R900'"
	end if
end if


'===============================================================================
if (page="") then page = 1
if (research="") then
	''divcd = "A003"
	chkGubun = "err"
	''exCheckFinish = "Y"
end if


'===============================================================================
dim oCCSRefundCheck

set oCCSRefundCheck = new CCSRefundCheck


oCCSRefundCheck.FPageSize = PageSize
oCCSRefundCheck.FCurrPage = page

oCCSRefundCheck.FRectOrderSerial = orderserial
oCCSRefundCheck.FRectDivCD = divcd
oCCSRefundCheck.FRectReturnMethod = returnmethod
oCCSRefundCheck.FRectStartDate = fromDate
oCCSRefundCheck.FRectEndDate = toDate
oCCSRefundCheck.FRectChkGubun = chkGubun
oCCSRefundCheck.FRectRefundMin = refundMin
oCCSRefundCheck.FRectRefundMax = refundMax

oCCSRefundCheck.FRectExCheckFinish = exCheckFinish
oCCSRefundCheck.FRectReturnMethodIN = returnmethodIN
oCCSRefundCheck.FRectDategbn = dategbn
oCCSRefundCheck.GetRefundCheckList


dim yyyymm : yyyymm = yyyy1 & "-" & mm1
Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & Replace(yyyymm, "-", "")
Dim appPath : appPath = server.mappath(AdmPath) + "/"

Dim sNow, sY, sM, sD, sH, sMi, sS, sDateName
sNow = now()
sY= Year(sNow)
sM = Format00(2,Month(sNow))
sD = Format00(2,Day(sNow))
sH = Format00(2,Hour(sNow))
sMi = Format00(2,Minute(sNow))
sS = Format00(2,Second(sNow))
sDateName = sY&sM&sD&sH&sMi&sS

Dim FileName: FileName = "refund_check_"&sDateName&".csv"
dim fso, tFile

Dim ArrRows
Dim headLine, bufstr
IF oCCSRefundCheck.FResultCount > 0 THEN
    Set fso = CreateObject("Scripting.FileSystemObject")
	If NOT fso.FolderExists(appPath) THEN
		fso.CreateFolder(appPath)
	END If
	Set tFile = fso.CreateTextFile(appPath & FileName )

	headLine = ",ASID,주문번호,구분,사유01,사유02,제목,환불방식,취소/반품,환불액,업체정산,정산사유,관련입금,접수일,완료일"
	tFile.WriteLine headLine

	for i = 0 to oCCSRefundCheck.FResultCount - 1
		bufstr = "," & oCCSRefundCheck.FItemList(i).Fasid
		bufstr = bufstr & "," & oCCSRefundCheck.FItemList(i).FOrderserial
		bufstr = bufstr & "," & oCCSRefundCheck.FItemList(i).Fdivcdname
		bufstr = bufstr & "," & oCCSRefundCheck.FItemList(i).Fgubun01name
		bufstr = bufstr & "," & oCCSRefundCheck.FItemList(i).Fgubun02name
		bufstr = bufstr & "," & replace(oCCSRefundCheck.FItemList(i).Ftitle,","," ")
		bufstr = bufstr & "," & oCCSRefundCheck.FItemList(i).FreturnmethodName
		bufstr = bufstr & "," & oCCSRefundCheck.FItemList(i).FOrgRefundRequire
		bufstr = bufstr & "," & oCCSRefundCheck.FItemList(i).Frefundresult
		bufstr = bufstr & "," & oCCSRefundCheck.FItemList(i).Fadd_upchejungsandeliverypay
		bufstr = bufstr & "," & oCCSRefundCheck.FItemList(i).Fadd_upchejungsancause
		bufstr = bufstr & "," & oCCSRefundCheck.FItemList(i).FappPrice
		bufstr = bufstr & "," & Left(oCCSRefundCheck.FItemList(i).Fregdate,10)
		bufstr = bufstr & "," & Left(oCCSRefundCheck.FItemList(i).Ffinishdate,10)

		tFile.WriteLine bufstr
	next

	tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
end if

response.redirect AdmPath&"/"&FileName

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
