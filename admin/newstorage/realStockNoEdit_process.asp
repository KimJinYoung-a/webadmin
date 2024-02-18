<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/summaryupdatelib.asp"-->
<!-- #include virtual="/lib/classes/items/limit_item_cls.asp"-->
<%

dim ArrIteminfo, ArrStockNo, ix, sqlStr

ArrIteminfo = request("arrCheckinfo")
ArrStockNo = request("arrStockNo")


dim ArrCount, ArrItemGubun, ArrItemID, ArrItemOption
dim SplitStockNo, ArrContents, SplitContents, realstockNo
dim realstock, offconfirmno, ipkumdiv2, ipkumdiv4, ipkumdiv5
ArrContents = Split(ArrIteminfo,"|")
SplitStockNo = Split(ArrStockNo,"|")
ArrCount = UBound(ArrContents)

dim refer, result
refer = request.ServerVariables("HTTP_REFERER")
    for ix=1 to ArrCount
        SplitContents=Split(ArrContents(ix),"_")

        ArrItemGubun=SplitContents(0)
        ArrItemID=SplitContents(1)
        ArrItemOption=SplitContents(2)

        result = UpdateItemLimitCount(ArrItemID, ArrItemOption, SplitStockNo(ix), 0)

        sqlStr = "exec db_summary.dbo.sp_Ten_SellYnSetByLimitNo " & CStr(ArrItemID)
        dbget.Execute sqlStr

    next

	response.write "<script language='javascript'>"
	response.write "alert('수정 되었습니다.');"
	response.write "location.replace('" + refer + "');"
	response.write "</script>"
	dbget.close()	:	response.End
%>