<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<script language="javascript" runat="server">
var confirmDt = (new Date()).valueOf();
</script>
<style>
body {
  font-size: small;
}
</style>
</head>
<body bgcolor="#F4F4F4" >
<%
function getEzWelDlvCode2Name(idlvCd)
    if isNULL(idlvCd) then Exit function

    SELECT CASE idlvCd
        CASE "1007" : getEzWelDlvCode2Name = "CJ대한통운"
        CASE "1017" : getEzWelDlvCode2Name = "롯데택배"
        CASE "1016" : getEzWelDlvCode2Name = "한진택배"
        CASE "1008" : getEzWelDlvCode2Name = "로젠택배"
        CASE "1161" : getEzWelDlvCode2Name = "우편등기"

        CASE "1180" : getEzWelDlvCode2Name = "일양로지스"
        CASE "1163" : getEzWelDlvCode2Name = "이노지스"
        CASE "1200" : getEzWelDlvCode2Name = "대신택배"
        CASE "1082" : getEzWelDlvCode2Name = "기타택배"
        CASE "1001" : getEzWelDlvCode2Name = "DHL"
        CASE "1002" : getEzWelDlvCode2Name = "KGB택배"
        CASE "1005" : getEzWelDlvCode2Name = "경동택배"
        CASE "1011" : getEzWelDlvCode2Name = "옐로우캡"
        CASE "1012" : getEzWelDlvCode2Name = "우체국택배EMS"
        CASE "1014" : getEzWelDlvCode2Name = "천일택배"
        CASE "1080" : getEzWelDlvCode2Name = "KG로지스택배"
        CASE "1081" : getEzWelDlvCode2Name = "업체직배송"
        CASE "1260" : getEzWelDlvCode2Name = "GTX로지스"
        
        CASE "1102" : getEzWelDlvCode2Name = "합동택배"
        CASE "1103" : getEzWelDlvCode2Name = "한의사랑택배"
        CASE "1104" : getEzWelDlvCode2Name = "다드림"
        CASE "1105" : getEzWelDlvCode2Name = "굿투럭"
        CASE "1106" : getEzWelDlvCode2Name = "건영택배"
        CASE "1107" : getEzWelDlvCode2Name = "호남택배"
        CASE "1108" : getEzWelDlvCode2Name = "CJ대한통운국제특송"
        CASE "1109" : getEzWelDlvCode2Name = "EMS"
        CASE "1110" : getEzWelDlvCode2Name = "한덱스"
        CASE "1111" : getEzWelDlvCode2Name = "FedEx"
        CASE "1112" : getEzWelDlvCode2Name = "UPS"
        CASE "1113" : getEzWelDlvCode2Name = "TNT"
        CASE "1114" : getEzWelDlvCode2Name = "USPS"
        CASE "1115" : getEzWelDlvCode2Name = "i-parcel"
        CASE "1116" : getEzWelDlvCode2Name = "GSM NtoN"
        CASE "1117" : getEzWelDlvCode2Name = "성원글로벌"
        CASE "1118" : getEzWelDlvCode2Name = "범한판토스"
        CASE "1119" : getEzWelDlvCode2Name = "ACI Express"
        CASE "1121" : getEzWelDlvCode2Name = "대운글로벌"
        CASE "1122" : getEzWelDlvCode2Name = "에어보이익스프레스"
        CASE "1123" : getEzWelDlvCode2Name = "KGL네트웍스"
        CASE "1124" : getEzWelDlvCode2Name = "LineExpress"
        CASE "1125" : getEzWelDlvCode2Name = "2fast익스프레스"
        CASE "1126" : getEzWelDlvCode2Name = "GSI익스프레스"
        CASE "1240" : getEzWelDlvCode2Name = "편의점택배"
        CASE ELSE : getEzWelDlvCode2Name =""
    END SELECT 
end function

Dim sqlStr
Dim isOnlyTodayBaljuView : isOnlyTodayBaljuView=false  ''발주된내역만 (주문수신 view)
Dim isDlvConfirmProc 	 : isDlvConfirmProc=false  ''주문확인 Proc
Dim isDlvInputProc 	 	 : isDlvInputProc=false    ''입력 Proc
Dim eddt : eddt=requestCheckvar(request("eddt"),10)

Dim IS_TEST_MODE : IS_TEST_MODE=FALSE 

Dim istyyyymmdd, iedyyyymmdd
    iedyyyymmdd = LEFT(dateadd("d",-2,now()),10)

    if eddt<>"" then 
        if isDate(eddt) then
            iedyyyymmdd=eddt  '''yyyy-mm-dd
        end if
    end if
    istyyyymmdd = LEFT(dateadd("d",-14,iedyyyymmdd),10)
 
sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'ezwel','"&confirmDt&"'"
dbget.Execute sqlStr
rw "초기화작업"

'' 1001:주문완료 / 1002:출고준비중 / 1003:배송중 / 1004:수취완료 / 1005:주문취소 / 1007:반품요청 ....
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1001","주문미확인")
rw "---------------------"
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1002","주문확인")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1003","배송중")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-14,iedyyyymmdd),10)

call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1001","주문미확인")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1002","주문확인")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1003","배송중")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)

call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1001","주문미확인")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1002","주문확인")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1003","배송중")
response.flush

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'ezwel','"&confirmDt&"'"
dbget.Execute sqlStr
rw "주문매핑"

rw "완료"

' 일별정산시 이용할 수 있을듯함.
' call Get_ezwelOrderListByStatus("2019-10-25","2019-10-25","1004","수취완료")
' response.flush
'response.write("<script>setTimeout(alert('완료'),1000);self.close();</script>")

function Get_ezwelOrderListByStatus(stdate,eddate,iorderStatus,istatusName)
	Dim sellsite : sellsite = "ezwel"
	Dim jUrl, xmlSelldate, iRbody
	Dim objXML, objData, strObj
	Dim masterCnt, detailCnt, resultcode, obj
	Dim arrOrderList, arrOrderGoods
	Dim objDetailListXML, objDetailOneXML
	Dim i, j, k
	Dim getParam
    Dim strSql, bufStr
    Dim ordNo, ordItemSeq, shppNo, shppSeq, reOrderYn, delayNts
    Dim cspGoodsCd, goodsCd, uitemId, orderQty, shppDivDtlNm
    Dim optionContent, shppRsvtDt, whoutCritnDt, autoShortgYn
    Dim orderStatus, dlvrCd, dlvrNo, dlvrDt, dlvrFinishDt, cancelDt
    Dim paramInfo, retParamInfo, RetErr
    Dim shppTypeDtlNm, delicoVenId, delicoVenNm, wblNo
    Get_ezwelOrderListByStatus = False

	getParam = ""
	getParam = getParam & "mallId=ezwel&sellDate="&Replace(stdate, "-", "")&"&code="&iorderStatus
    If application("Svr_Info")="Dev" Then
        jUrl = "http://gateway.10x10.co.kr/external/apis/order?" & getParam
    Else
        jUrl = "http://gateway.10x10.co.kr/external/apis/order?" & getParam
    End If

    rw "기간검색:"&stdate&"~"&eddate&" 상태:"&iorderStatus&"("&istatusName&")"
	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
        objXML.Open "GET", jUrl, false
        objXML.setRequestHeader "Content-Type", "application/json"
        objXML.setTimeouts 5000,80000,80000,80000
        objXML.Send()
        iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
        Set strObj = JSON.parse(iRbody)
            If objXML.Status <> "200" AND objXML.Status <> "201" Then
                response.write "ERROR : 통신오류" & objXML.Status
                response.write "<script>alert('ERROR : 통신오류.');</script>"
                dbget.close : response.end
            Else
                Set arrOrderList = strObj.result.ezwelResponseModel.arrOrderList
                    If arrOrderList.length < 1 Then
                        rw "내역없음 : 종료"
                        rw "resultMsg:"&strObj.result.ezwelResponseModel.resultMsg
                        Get_ezwelOrderListByStatus = True
                        Set objXML = Nothing
                        exit function
                    Else
                        rw "건수(" & arrOrderList.length & ")"
                        For i=0 to arrOrderList.length-1
                            Set arrOrderGoods = arrOrderList.get(i).arrOrderGoods
                                ordNo = arrOrderList.get(i).orderNum
                                For j=0 to arrOrderGoods.length-1
                                    ordItemSeq = arrOrderGoods.get(j).orderGoodsNum     ''주문순번
                                    shppNo = "" ''배송번호
                                    shppSeq ="" ''배송Seq
                                    reOrderYn ="N" ''재주문여부 
                                    delayNts  =""  ''지연일수
                                    cspGoodsCd = arrOrderGoods.get(j).cspGoodsCd        ''업체상품코드
                                    goodsCd = arrOrderGoods.get(j).goodsCd              ''(제휴)상품코드
                                    uitemId = "" ''제휴 단품ID
                                    orderQty = arrOrderGoods.get(j).orderQty              ''주문수량
                                    shppDivDtlNm = "" ''배송구분상세명 (출고/교환출고..)
                                    optionContent = arrOrderGoods.get(j).optionContent              ''옵션명 */arrOrderGoods/optionContent
                                    If Left(optionContent, 9) = "<![CDATA[" Then
                                        optionContent = Trim(Replace(optionContent, "<![CDATA[", ""))
                                    End If

                                    If Right(optionContent,3) = "]]>" Then
                                        optionContent = Trim(Left(optionContent, Len(optionContent) - 3))
                                    End If

                                    shppRsvtDt = ""  ''예정일?
                                    whoutCritnDt ="" ''출고기준일
                                    autoShortgYn ="" ''자동결품여부
                                    orderStatus = arrOrderGoods.get(j).orderStatus        ''주문상태

                                    On Error Resume Next
                                        dlvrCd = ""&arrOrderGoods.get(j).dlvrCd&""        ''택배사
                                        If Err.number <> 0 Then
                                            dlvrCd = ""
                                        End If
                                    On Error Goto 0

                                    On Error Resume Next
                                        dlvrNo = arrOrderGoods.get(j).dlvrNo        ''송장번호
                                        If Err.number <> 0 Then
                                            dlvrNo = ""
                                        End If
                                    On Error Goto 0

                                    On Error Resume Next
                                        dlvrDt = arrOrderGoods.get(j).dlvrDt        ''배송일
                                        If Err.number <> 0 Then
                                            dlvrDt = ""
                                        End If
                                    On Error Goto 0
                                    
                                    On Error Resume Next
                                        dlvrFinishDt = arrOrderGoods.get(j).dlvrFinishDt        ''배송완료일
                                        If Err.number <> 0 Then
                                            dlvrFinishDt = ""
                                        End If
                                    On Error Goto 0

                                    On Error Resume Next
                                        cancelDt = arrOrderGoods.get(j).cancelDt        ''취소일
                                        If Err.number <> 0 Then
                                            cancelDt = ""
                                        End If
                                    On Error Goto 0

                                    bufStr = ""
                                    bufStr = sellsite&"|"&ordNo
                                    bufStr = bufStr &"|"&ordItemSeq
                                    bufStr = bufStr &"|"&cspGoodsCd
                                    bufStr = bufStr &"|"&goodsCd
                                    
                                    bufStr = bufStr &"|"&orderQty

                                    bufStr = bufStr &"|"&orderStatus
                                    bufStr = bufStr &"|"&dlvrCd
                                    bufStr = bufStr &"|"&dlvrNo

                                    bufStr = bufStr &"|"&dlvrDt
                                    bufStr = bufStr &"|"&dlvrFinishDt
                                    bufStr = bufStr &"|"&cancelDt
                                    shppTypeDtlNm = ""
                                    delicoVenId   = dlvrCd
                                    delicoVenNm   = getEzWelDlvCode2Name(dlvrCd)
                                    wblNo         = dlvrNo

                                    sqlStr = "db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Input]"
                                    paramInfo = Array(Array("@RETURN_VALUE",adInteger	,adParamReturnValue	,,0) _
                                        ,Array("@SellSite" 				, adVarchar		, adParamInput		, 	32, sellsite)	_
                                        ,Array("@OutMallOrderSerial"	, adVarchar		, adParamInput		,	32, Trim(ordNo)) _
                                        ,Array("@OrgDetailKey"			, adVarchar		, adParamInput		,   32, Trim(ordItemSeq)) _

                                        ,Array("@confirmDt"				, adVarchar     , adParamInput		,	16, Trim(confirmDt)) _
                                        ,Array("@shppNo"				, adVarchar		, adParamInput		,   32, Trim(shppNo)) _
                                        ,Array("@shppSeq"				, adVarchar		, adParamInput		,   10, Trim(shppSeq)) _
                                        ,Array("@reOrderYn"				, adVarchar		, adParamInput		,    1, Trim(reOrderYn)) _
                                        ,Array("@delayNts"			    , adInteger		, adParamInput		,     , Trim(delayNts)) _
                                        ,Array("@splVenItemId"			, adInteger		, adParamInput		,     , Trim(cspGoodsCd)) _
                                        ,Array("@outMallGoodsNo"		, adVarchar		, adParamInput		,   16, Trim(goodsCd)) _
                                        ,Array("@outMallOptionNo"		, adVarchar		, adParamInput		,   32, Trim(uitemId)) _
                                        ,Array("@ordQty"			    , adInteger		, adParamInput		,     , Trim(orderQty)) _
                                        ,Array("@shppDivDtlNm"		    , adVarchar		, adParamInput		,   20, Trim(shppDivDtlNm)) _
                                        ,Array("@uitemNm"		        , adVarchar		, adParamInput		,   128, Trim(optionContent)) _
                                        ,Array("@shppRsvtDt"			, adDate		, adParamInput		,	  , Trim(shppRsvtDt)) _
                                        ,Array("@whoutCritnDt"			, adDate		, adParamInput		,	  , Trim(whoutCritnDt)) _
                                        ,Array("@autoShortgYn"			, adVarchar		, adParamInput		,    1, Trim(autoShortgYn)) _
                                        ,Array("@outorderstatus"		, adVarchar		, adParamInput		,   30, Trim(orderStatus)) _

                                        ,Array("@shppTypeDtlNm"		, adVarchar		, adParamInput		,   16, Trim(shppTypeDtlNm)) _
                                        ,Array("@delicoVenId"		, adVarchar		, adParamInput		,   16, Trim(delicoVenId)) _
                                        ,Array("@delicoVenNm"		, adVarchar		, adParamInput		,   32, Trim(delicoVenNm)) _
                                        ,Array("@wblNo"		        , adVarchar		, adParamInput		,   32, Trim(wblNo)) _

                                        ,Array("@invoiceUpDt"	    , adVarchar		, adParamInput		,   19, "") _
                                        ,Array("@outjFixedDt"		, adVarchar		, adParamInput		,   19, Trim(dlvrFinishDt)) _
                                    )

                                    'On Error RESUME Next
                                    retParamInfo = fnExecSPOutput(sqlStr, paramInfo)
                                    RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
                                Next
                            Set arrOrderGoods = nothing
                        Next
                    End If
                Set arrOrderList = nothing
            End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If

        Set strObj = nothing
	Get_ezwelOrderListByStatus = True
	Set objXML = Nothing
end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->