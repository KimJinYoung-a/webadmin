<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/DataMartItemsalecls.asp"-->

<%
Dim iColorList :iColorList = Array("F60925","F2F84A","0611F9","B22222","7CFC00","808000","006400","BA55D3","008B8B","FF1493","000080","DCDCDC","A52A2A","F08080","D2691E","B8860B","7CFC00")

Dim mode : mode = RequestCheckVar(request("mode"),10)
Dim yyyy1 : yyyy1=RequestCheckVar(request("yyyy1"),4)
Dim yyyy2 : yyyy2=RequestCheckVar(request("yyyy2"),4)
Dim mm1 : mm1=RequestCheckVar(request("mm1"),2)
Dim mm2 : mm2=RequestCheckVar(request("mm2"),2)

Dim ckMinus : ckMinus=RequestCheckVar(request("ckMinus"),10)

Dim cdl : cdl=RequestCheckVar(request("cdl"),3)
Dim cdm : cdm=RequestCheckVar(request("cdm"),3)
Dim cds : cds=RequestCheckVar(request("cds"),3)
Dim cdx : cdx=RequestCheckVar(request("cdx"),3)
Dim catebase : catebase=RequestCheckVar(request("catebase"),10)

Dim fromDate : fromDate = yyyy1 + "-" + mm1 + "-" + "01"
Dim toDate : toDate = Left(CStr(DateAdd("m",1,DateSerial(yyyy2,mm2,"01"))),10)
Dim DateGubun : DateGubun="M"

IF (yyyy1 + "-" + mm1=yyyy2 + "-" + mm2) then DateGubun="D"

dim oReport, i, j
Dim Buf : Buf=""
Dim Categories,DataSerise

Buf = Buf & "<?xml version='1.0' encoding='EUC-KR' ?>"&VbCRLF
Buf = Buf & "<chart chartBottomMargin='2' formatNumberScale='0' drawAnchors='1' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' canvasBorderColor='CBCBCB' animation='1' baseFontColor='666666' bgColor='FFFFFF' formatNumber='1' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' divLineIsDashed='1' showBorder='0' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' anchorRadius='5' plotBorderAlpha='20' >"&VbCRLF

IF (mode="C1") then
    set oReport = new CDatamartItemSale
    oReport.FRectStartDate = fromDate
    oReport.FRectEndDate = toDate
    oReport.FRectDateGubun = DateGubun
    oReport.FRectIncludeMinus = ckMinus
    oReport.FRectCD1 = cdL
    oReport.FRectCD2 = cdM
    oReport.FRectCD3 = cdS
    oReport.FRectCD4 = cdx

    if (catebase="V") then
        oReport.getCateSellTrandByCurrentDispCate
    elseif (catebase="C") then
        oReport.getCateSellTrandByCurrentCate
    else
        oReport.getCateSellTrand
    end if

    Categories = oReport.getDPartList
    DataSerise = oReport.getCateList

    for j=LBound(Categories) to UBound(Categories)
    Buf = Buf & "<categories>"&VbCRLF
        Buf = Buf & "<category name='" & Categories(j) & "' showName='1' showLine='1' />"&VbCRLF
    Buf = Buf & "</categories>"&VbCRLF
    next

    for j=LBound(DataSerise) to UBound(DataSerise)
        Buf = Buf & "<dataset seriesName='"&DataSerise(j)&"' color='"&iColorList(j)&"' showValues='0' parentYAxis='P'>"&VbCRLF
        for i=0 to oReport.FResultCount - 1
            if (DataSerise(j)=oReport.FItemList(i).FcateName) then
            	Buf = Buf & "<set value='" & oReport.FItemList(i).Fpro & "' />"&VbCRLF
            end if
        next
        Buf = Buf & "</dataset>"&VbCRLF
    next

    set oReport = Nothing
END IF
Buf = Buf & "</chart>"&VbCRLF

response.write Buf
%>



<!-- #include virtual="/lib/db/db3close.asp" -->