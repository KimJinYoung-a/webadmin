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
        CASE "1007" : getEzWelDlvCode2Name = "CJ�������"
        CASE "1017" : getEzWelDlvCode2Name = "�Ե��ù�"
        CASE "1016" : getEzWelDlvCode2Name = "�����ù�"
        CASE "1008" : getEzWelDlvCode2Name = "�����ù�"
        CASE "1161" : getEzWelDlvCode2Name = "������"

        CASE "1180" : getEzWelDlvCode2Name = "�Ͼ������"
        CASE "1163" : getEzWelDlvCode2Name = "�̳�����"
        CASE "1200" : getEzWelDlvCode2Name = "����ù�"
        CASE "1082" : getEzWelDlvCode2Name = "��Ÿ�ù�"
        CASE "1001" : getEzWelDlvCode2Name = "DHL"
        CASE "1002" : getEzWelDlvCode2Name = "KGB�ù�"
        CASE "1005" : getEzWelDlvCode2Name = "�浿�ù�"
        CASE "1011" : getEzWelDlvCode2Name = "���ο�ĸ"
        CASE "1012" : getEzWelDlvCode2Name = "��ü���ù�EMS"
        CASE "1014" : getEzWelDlvCode2Name = "õ���ù�"
        CASE "1080" : getEzWelDlvCode2Name = "KG�������ù�"
        CASE "1081" : getEzWelDlvCode2Name = "��ü�����"
        CASE "1260" : getEzWelDlvCode2Name = "GTX������"
        
        CASE "1102" : getEzWelDlvCode2Name = "�յ��ù�"
        CASE "1103" : getEzWelDlvCode2Name = "���ǻ���ù�"
        CASE "1104" : getEzWelDlvCode2Name = "�ٵ帲"
        CASE "1105" : getEzWelDlvCode2Name = "������"
        CASE "1106" : getEzWelDlvCode2Name = "�ǿ��ù�"
        CASE "1107" : getEzWelDlvCode2Name = "ȣ���ù�"
        CASE "1108" : getEzWelDlvCode2Name = "CJ��������Ư��"
        CASE "1109" : getEzWelDlvCode2Name = "EMS"
        CASE "1110" : getEzWelDlvCode2Name = "�ѵ���"
        CASE "1111" : getEzWelDlvCode2Name = "FedEx"
        CASE "1112" : getEzWelDlvCode2Name = "UPS"
        CASE "1113" : getEzWelDlvCode2Name = "TNT"
        CASE "1114" : getEzWelDlvCode2Name = "USPS"
        CASE "1115" : getEzWelDlvCode2Name = "i-parcel"
        CASE "1116" : getEzWelDlvCode2Name = "GSM NtoN"
        CASE "1117" : getEzWelDlvCode2Name = "�����۷ι�"
        CASE "1118" : getEzWelDlvCode2Name = "�������佺"
        CASE "1119" : getEzWelDlvCode2Name = "ACI Express"
        CASE "1121" : getEzWelDlvCode2Name = "���۷ι�"
        CASE "1122" : getEzWelDlvCode2Name = "������ͽ�������"
        CASE "1123" : getEzWelDlvCode2Name = "KGL��Ʈ����"
        CASE "1124" : getEzWelDlvCode2Name = "LineExpress"
        CASE "1125" : getEzWelDlvCode2Name = "2fast�ͽ�������"
        CASE "1126" : getEzWelDlvCode2Name = "GSI�ͽ�������"
        CASE "1240" : getEzWelDlvCode2Name = "�������ù�"
        CASE ELSE : getEzWelDlvCode2Name =""
    END SELECT 
end function

Dim sqlStr
Dim isOnlyTodayBaljuView : isOnlyTodayBaljuView=false  ''���ֵȳ����� (�ֹ����� view)
Dim isDlvConfirmProc 	 : isDlvConfirmProc=false  ''�ֹ�Ȯ�� Proc
Dim isDlvInputProc 	 	 : isDlvInputProc=false    ''�Է� Proc
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
rw "�ʱ�ȭ�۾�"

'' 1001:�ֹ��Ϸ� / 1002:����غ��� / 1003:����� / 1004:����Ϸ� / 1005:�ֹ���� / 1007:��ǰ��û ....
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1001","�ֹ���Ȯ��")
rw "---------------------"
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1002","�ֹ�Ȯ��")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1003","�����")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-14,iedyyyymmdd),10)

call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1001","�ֹ���Ȯ��")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1002","�ֹ�Ȯ��")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1003","�����")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)

call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1001","�ֹ���Ȯ��")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1002","�ֹ�Ȯ��")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1003","�����")
response.flush

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'ezwel','"&confirmDt&"'"
dbget.Execute sqlStr
rw "�ֹ�����"

rw "�Ϸ�"

' �Ϻ������ �̿��� �� ��������.
' call Get_ezwelOrderListByStatus("2019-10-25","2019-10-25","1004","����Ϸ�")
' response.flush
'response.write("<script>setTimeout(alert('�Ϸ�'),1000);self.close();</script>")

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

    rw "�Ⱓ�˻�:"&stdate&"~"&eddate&" ����:"&iorderStatus&"("&istatusName&")"
	'// =======================================================================
	'// ����Ÿ ��������
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
        objXML.Open "GET", jUrl, false
        objXML.setRequestHeader "Content-Type", "application/json"
        objXML.setTimeouts 5000,80000,80000,80000
        objXML.Send()
        iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
        Set strObj = JSON.parse(iRbody)
            If objXML.Status <> "200" AND objXML.Status <> "201" Then
                response.write "ERROR : ��ſ���" & objXML.Status
                response.write "<script>alert('ERROR : ��ſ���.');</script>"
                dbget.close : response.end
            Else
                Set arrOrderList = strObj.result.ezwelResponseModel.arrOrderList
                    If arrOrderList.length < 1 Then
                        rw "�������� : ����"
                        rw "resultMsg:"&strObj.result.ezwelResponseModel.resultMsg
                        Get_ezwelOrderListByStatus = True
                        Set objXML = Nothing
                        exit function
                    Else
                        rw "�Ǽ�(" & arrOrderList.length & ")"
                        For i=0 to arrOrderList.length-1
                            Set arrOrderGoods = arrOrderList.get(i).arrOrderGoods
                                ordNo = arrOrderList.get(i).orderNum
                                For j=0 to arrOrderGoods.length-1
                                    ordItemSeq = arrOrderGoods.get(j).orderGoodsNum     ''�ֹ�����
                                    shppNo = "" ''��۹�ȣ
                                    shppSeq ="" ''���Seq
                                    reOrderYn ="N" ''���ֹ����� 
                                    delayNts  =""  ''�����ϼ�
                                    cspGoodsCd = arrOrderGoods.get(j).cspGoodsCd        ''��ü��ǰ�ڵ�
                                    goodsCd = arrOrderGoods.get(j).goodsCd              ''(����)��ǰ�ڵ�
                                    uitemId = "" ''���� ��ǰID
                                    orderQty = arrOrderGoods.get(j).orderQty              ''�ֹ�����
                                    shppDivDtlNm = "" ''��۱��л󼼸� (���/��ȯ���..)
                                    optionContent = arrOrderGoods.get(j).optionContent              ''�ɼǸ� */arrOrderGoods/optionContent
                                    If Left(optionContent, 9) = "<![CDATA[" Then
                                        optionContent = Trim(Replace(optionContent, "<![CDATA[", ""))
                                    End If

                                    If Right(optionContent,3) = "]]>" Then
                                        optionContent = Trim(Left(optionContent, Len(optionContent) - 3))
                                    End If

                                    shppRsvtDt = ""  ''������?
                                    whoutCritnDt ="" ''��������
                                    autoShortgYn ="" ''�ڵ���ǰ����
                                    orderStatus = arrOrderGoods.get(j).orderStatus        ''�ֹ�����

                                    On Error Resume Next
                                        dlvrCd = ""&arrOrderGoods.get(j).dlvrCd&""        ''�ù��
                                        If Err.number <> 0 Then
                                            dlvrCd = ""
                                        End If
                                    On Error Goto 0

                                    On Error Resume Next
                                        dlvrNo = arrOrderGoods.get(j).dlvrNo        ''�����ȣ
                                        If Err.number <> 0 Then
                                            dlvrNo = ""
                                        End If
                                    On Error Goto 0

                                    On Error Resume Next
                                        dlvrDt = arrOrderGoods.get(j).dlvrDt        ''�����
                                        If Err.number <> 0 Then
                                            dlvrDt = ""
                                        End If
                                    On Error Goto 0
                                    
                                    On Error Resume Next
                                        dlvrFinishDt = arrOrderGoods.get(j).dlvrFinishDt        ''��ۿϷ���
                                        If Err.number <> 0 Then
                                            dlvrFinishDt = ""
                                        End If
                                    On Error Goto 0

                                    On Error Resume Next
                                        cancelDt = arrOrderGoods.get(j).cancelDt        ''�����
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
                                    RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ�
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