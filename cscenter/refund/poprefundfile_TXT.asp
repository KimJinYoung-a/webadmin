<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_refundcls.asp" -->

<%


dim upfiledate
upfiledate  = request("upfiledate")

dim OrefundList
set OrefundList = new CCSRefund
OrefundList.FCurrPage           = 1
OrefundList.FPageSize           = 1000
OrefundList.FRectCurrstate      = "B001"
OrefundList.FRectReturnmethod   = "R007"
OrefundList.FRectUploadState    = "uploaded"
OrefundList.FRectUpfiledate     = upfiledate

OrefundList.GetRefundRequireList

dim i
dim xlFilename

if (upfiledate="") then
    xlFilename = "환불목록_미처리전체_" &  Replace(Replace(Replace(FormatDate(now(),"0000.00.00-00:00:00")," ",""),"-",""),":","")
else
    xlFilename = "환불목록_" & Replace(Replace(Replace(upfiledate," ",""),"-",""),":","")
end if


Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" & xlFilename & ".txt"

for i=0 to OrefundList.FResultCount - 1
    response.write  OrefundList.FItemList(i).getUploadbankname & "," & OrefundList.FItemList(i).getUploadrebankaccount & "," & OrefundList.FItemList(i).Frefundrequire & "," & OrefundList.FItemList(i).getUploadrebankownername  & "," & "(주)텐바이텐" & Vbcrlf
next

set OrefundList = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->