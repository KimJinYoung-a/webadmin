<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 40
%>
<%
'###########################################################
' Description : 불량오차상품관리 엑셀다운로드
' History : 이상구 생성
'           2021.04.06 한용민 수정(검색조건수정)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%

dim makerid,mode, searchtype, purchasetype, mwdiv, sellyn, onlyisusing, makeruseyn, itemgubun
dim datetype, centermwdiv, monthlymwdiv, yyyy, mm
	makerid 		= requestCheckVar(request("makerid"),32)
	mode 			= requestCheckVar(request("mode"),32)
	searchtype 		= requestCheckVar(request("searchtype"),32)
	purchasetype 	= requestCheckVar(request("purchasetype"),32)
	mwdiv 			= requestCheckVar(request("mwdiv"),32)
	sellyn 			= requestCheckVar(request("sellyn"),32)
	onlyisusing 	= requestCheckVar(request("onlyisusing"),32)
	makeruseyn	 	= requestCheckVar(request("makeruseyn"),32)
	itemgubun 		= requestCheckVar(request("itemgubun"),32)
	datetype 		= requestCheckVar(request("datetype"),32)
	yyyy 			= requestCheckVar(request("yyyy1"),32)
	mm 				= requestCheckVar(request("mm1"),32)
	centermwdiv		= requestcheckvar(request("centermwdiv"),1)
	monthlymwdiv	= requestcheckvar(request("monthlymwdiv"),1)

if (searchtype = "") then
	searchtype = "bad"
	'datetype = "curr"
	yyyy = Left(now(),4)
	mm   = mid(now(),6,2)
end if

if (itemgubun = "") then
	itemgubun = "10"
end if
datetype = "yyyymm"
' 현재월일경우
'if yyyy = Left(now(),4) and mm = mid(now(),6,2) then
'	datetype = "curr"
'end if

dim osummarystock
set osummarystock = new CSummaryItemStock
	osummarystock.FRectmakerid = "all"
	osummarystock.FRectSearchType = searchtype
	osummarystock.FRectDatetype   = datetype
	osummarystock.FRectYYYYMM = yyyy+"-"+mm
	osummarystock.FRectMWDiv = mwdiv
	osummarystock.FRectlastmwdiv = monthlymwdiv
	osummarystock.FRectCenterMWDiv = centermwdiv
	osummarystock.FRectSellYN = sellyn
	osummarystock.FRectOnlyIsUsing = onlyisusing
	osummarystock.FRectItemGubun = itemgubun
	osummarystock.FRectPurchaseType = purchasetype
	osummarystock.FRectMakerUseYN = makeruseyn
	osummarystock.FPageSize = 10000
	osummarystock.GetBadOrErrItemListByBrand

Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & Replace(Left(Now,7), "-", "")
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

Dim FileName: FileName = "badorerritem_"&sDateName&".csv"
dim fso, tFile
dim i, j

if osummarystock.FResultCount > 0 then
    Set fso = CreateObject("Scripting.FileSystemObject")
	If NOT fso.FolderExists(appPath) THEN
		fso.CreateFolder(appPath)
	END If
	Set tFile = fso.CreateTextFile(appPath & FileName )

	tFile.WriteLine "브랜드ID,거래구분,상품구분,상품코드,옵션,물류바코드,상품명,옵션명,소비자가,매입가,판매여부,사용여부,불량수량,매입가합,실사유효재고,최종입고월"

	for i=0 to osummarystock.FResultCount - 1
		tFile.WriteLine """" & osummarystock.FItemList(i).Fmakerid & """,""" & osummarystock.FItemList(i).GetMwDivName & """," & osummarystock.FItemList(i).FItemgubun & "," & osummarystock.FItemList(i).FItemid & ",=""" & osummarystock.FItemList(i).FItemoption & """," &_
			"=""" & BF_MakeTenBarcode(osummarystock.FItemList(i).FItemgubun,osummarystock.FItemList(i).FItemID,osummarystock.FItemList(i).FItemoption) & """,""" & osummarystock.FItemList(i).FItemname & """,""" & osummarystock.FItemList(i).FItemOptionName & """," &_
			osummarystock.FItemList(i).Fsellcash & "," & osummarystock.FItemList(i).Fbuycash & "," & osummarystock.FItemList(i).Fsellyn & "," & osummarystock.FItemList(i).Fisusing & "," & osummarystock.FItemList(i).Fregitemno & "," & osummarystock.FItemList(i).Fbuycash*osummarystock.FItemList(i).Fregitemno & "," & osummarystock.FItemList(i).Frealstock & "," & osummarystock.FItemList(i).FlastIpgoDate
	next
    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing

	response.write osummarystock.FResultCount&"건 생성 ["&FileName&"]"
	response.redirect AdmPath&"/"&FileName
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
