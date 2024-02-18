<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 출고지시주문리스트 처리
' History : 이상구 생성
'           2020.11.11 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
' EMS 송장번호 입력 후 주문상태 변경

Dim songjangDiv	: songjangDiv	= req("songjangDiv","")
Dim idx			: idx			= req("idx","")

Dim orderSerial	: orderSerial	= req("orderSerial","")
Dim songjangNo	: songjangNo	= req("songjangNo","")
Dim realweight  : realweight	= req("realweight","") ''2016/05/26
Dim mode        : mode	= req("mode","") ''2016/05/26

dim boxSizeX, boxSizeY, boxSizeZ
boxSizeX = req("boxSizeX","")
boxSizeY = req("boxSizeY","")
boxSizeZ = req("boxSizeZ","")

''dbget.BeginTrans  ''트랜잭션 제거. 불필요

Dim ErrCode, ErrMsg, strSql, paramInfo, url, refer
  url = "popBaljuList.asp?idx=" & idx & "&songjangDiv=" & songjangDiv
  refer = request.ServerVariables("HTTP_REFERER")

if (mode="svsongjang") then
    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
    	,Array("@orderSerial"		, adVarchar	, adParamInput	, 11	, orderSerial) _
    	,Array("@songjangNo"		, adVarchar	, adParamInput	, 32	, songjangNo) _
    )

    if (songjangDiv = "8") then
    	'우체국배송
    	strSql = "db_order.dbo.sp_Ten_EpostSongjangInput"
    else
    	'EMS
    	strSql = "db_order.dbo.sp_Ten_EmsSongjangInput"
    end if
    Call fnExecSP(strSql, paramInfo)

    ErrCode  = CInt(GetValue(paramInfo, "@RETURN_VALUE"))			' 에러코드

    If ErrCode = -1 Then
    	''dbget.RollBackTrans
    	ErrMsg = "취소된 주문입니다."
    ElseIf Err Or ErrCode <> 0 Then
    	''dbget.RollBackTrans
    	ErrMsg = "오류발생 : " & Err.Number & " : " & Err.Source & " : " & Replace(Err.Description,"'","") & " : "
    Else
    	''dbget.CommitTrans
    End If

    If ErrMsg <> "" Then
    	response.write "<script type='text/javascript'>alert('" & ErrMsg & "');</script>"
    	response.write "<script type='text/javascript'>history.back();</script>"
    Else

    	response.write "<script type='text/javascript'>alert('입력되었습니다.');</script>"
    	response.write "<script type='text/javascript'>location.replace('"& refer &"');</script>"
    End If

elseif (mode="svttlwight") then
	'// 무게입력 : /admin/ordermaster/popbaljulist.asp, /v2/online/chulgo.asp?orderserial=주문번호
    strSql = " update [db_order].[dbo].tbl_ems_orderInfo "&vbCRLF
    strSql = strSql&" SET realWeight="&realweight&vbCRLF
    strSql = strSql&" where orderserial='"&orderserial&"'"&vbCRLF

    dbget.Execute strSql

    response.write "<script type='text/javascript'>alert('입력되었습니다.');</script>"
    response.write "<script type='text/javascript'>location.replace('"& refer &"');</script>"

elseif (mode="saveBoxSize") then
	strSql = " update [db_order].[dbo].tbl_ems_orderInfo "&vbCRLF
	strSql = strSql&" SET  boxSizeX = " & boxSizeX & ", boxSizeY = " & boxSizeY & ", boxSizeZ = " & boxSizeZ & vbCRLF
	strSql = strSql&" where orderserial='"&orderserial&"'"&vbCRLF

    dbget.Execute strSql

    response.write "<script type='text/javascript'>alert('입력되었습니다.');</script>"
    response.write "<script type='text/javascript'>location.replace('"& refer &"');</script>"

else
    response.write "<script type='text/javascript'>alert('미지정:mode:"&mode&"');</script>"
    response.write "<script type='text/javascript'>location.href = '" & url & "';</script>"
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
