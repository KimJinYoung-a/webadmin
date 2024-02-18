<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   
Session.codepage="949"
Response.codepage="949"
'###########################################################
' Page : /admin/itemmaster/deal/dodealitemreg.asp
' Description :  딜 상품 - 등록, 삭제
' History : 2022.06.29 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%
Response.ContentType = "application/json"
response.charset = "euc-kr"
Dim oDealitem, oDealItem2, arrList, arrList2, arrList3
Dim intLoop, oJson, objRst, oList, SalePer, MinPrice
Dim idx : idx = requestCheckVar(Request("idx"),9)

'등록 상품 리스트 정보
set oDealitem = new CDealItem
oDealitem.FRectMasterIDX = idx
arrList = oDealitem.fnGetDealEventItem
Set oJson = jsObject()

set oList = jsArray()
If isArray(arrList) Then
    For intLoop = 0 To UBound(arrList,2)
        Set objRst = jsObject()
        objRst("optionValue") = arrList(1,intLoop)
        objRst("optionName") = arrList(2,intLoop)
        set oList(null) = objRst
        Set objRst = Nothing
    Next
End If

'할인율 최저가 정보 가져오기
Set oDealItem2 = New ClsDeal
oDealItem2.FRectMasterIDX = idx
arrList2 = oDealItem2.fnGetMAXDealSalePer
arrList3 = oDealItem2.fnGetDealItemMinPrice
Set oDealItem2 = Nothing


If isArray(arrList2) Then
    If arrList2(2,0)="Y" Then
        SalePer = Cint(((arrList2(0,0)-arrList2(1,0))/arrList2(0,0))*100)
        If SalePer>0 Then
            SalePer = SalePer
        else
            SalePer = 0
        End If
    else
        SalePer = 0
    End If
End If

If isArray(arrList3) Then
    MinPrice = arrList3(0,0)
End If

set oJson("option") = oList
oJson("salePer") = SalePer
oJson("minPrice") = MinPrice
oJson.flush
Set oJson = Nothing
Set oList = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->