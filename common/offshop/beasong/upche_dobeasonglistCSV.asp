<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.26 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim ojumun ,ix,sql ,detailidxarr ,iSall, SheetType ,i, j
dim bufStr, tmpS
	detailidxarr =  request("detailidxarr")
	iSall   =  request("isall")
	SheetType  =  request("SheetType")

	bufStr = ""

If session("ssBctId") = "" then
    response.write "<script language='javascript'>alert('세션이 종료되었습니다.');</script>"
    dbget.close()	:	response.End
end if

function ReplaceSCVStr(oStr)
    ReplaceSCVStr = ""
    if IsNULL(oStr) then Exit function
    ReplaceSCVStr = Replace(oStr, chr(34),"'")

end function

set ojumun = new cupchebeasong_list
	ojumun.FRectdetailidxarr = detailidxarr
	ojumun.FRectIsAll       = iSall
	ojumun.FRectDesignerID = session("ssBctID")
	ojumun.fReDesignerSelectBaljuList()

response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".csv"
Response.CacheControl = "public"

bufStr = "주문번호,주문일,구매자명,구매자전화,구매자핸드폰,수령인,수령인전화,수령인핸드폰,우편번호,배송지주소1,배송지주소2,배송유의사항,택배번호,상품아이디,상품명,옵션,판매가,수량"

response.write bufStr & VbCrlf

for ix=0 to ojumun.FResultCount - 1
    bufStr = ""
    bufStr = bufStr & Chr(34) & ojumun.FItemList(ix).Forderno & Chr(34)
    bufStr = bufStr & "," & Chr(34) & Left(CStr(ojumun.FItemList(ix).FRegDate),10) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FBuyName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FBuyPhone) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FBuyHp) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FReqName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FReqPhone) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FReqHp) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FReqZipCode) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FReqZipAddr) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FReqAddress) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(db2html(ojumun.FItemList(ix).FComment)) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).Fsongjangno) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).Fitemid) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FItemName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FItemoptionName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ojumun.FItemList(ix).Fsellprice & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ojumun.FItemList(ix).FItemNo & Chr(34)

    response.write bufStr & VbCrlf
next %>
<%
set ojumun = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->