<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
'###########################################################################
'DIY 상품 불러오기 처리 페이지
'2016-12-05 이종화
'###########################################################################
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<%
'//작품 타입 , 작품 코드 불러온후
Dim itemstate , itemid
Dim objCmd , returnValue

itemstate = RequestCheckVar(request("itemstate"),1)  '//복제DB선택용 Y/N (판매DB) , W (대기DB)
itemid	  = RequestCheckVar(request("itemid"),32) '//복제할 itemid
'//작품별 DB 저장

'//2016-12-05 이종화 등록 대기 상품 복제 -- 상세 썸네일 이외에 전부 복사
If itemstate = "W" then
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbACADEMYget
				.CommandType = adCmdText
				.CommandText = "{?= call db_academy.[dbo].[sp_academy_diy_waititem_dummy] ('"& itemid &"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
		Set objCmd = nothing

ElseIf itemstate = "Y" Or itemstate = "N" Then
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
			With objCmd
				.ActiveConnection = dbACADEMYget
				.CommandType = adCmdText
				.CommandText = "{?= call db_academy.[dbo].[sp_academy_diy_item_dummy] ('"& itemid &"')}"
				.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
				.Execute, , adExecuteNoRecords
				End With
			    returnValue = objCmd(0).Value
		Set objCmd = nothing
End If 

'Response.write returnValue
'Response.end

If returnValue = 0 then
	Call Alert_return ("데이터 처리에 문제가 발생하였습니다.")
Else
%>
<script>
<%'작품 복제후 창 닫고 모창에 작품 수정 페이지 이동%>
<!--
	//fnAPPpopupWaitItemEdit('<%=g_AdminURL%>/apps/academy/itemmaster/artRegistEdit.asp?waititemid=<%=returnValue%>');
	parent.fnAPPopenerJsCallClose('fnMoveWaitItemEdit(\'<%=g_AdminURL%>/apps/academy/itemmaster/artRegistEdit.asp?waititemid=<%=returnValue%>\')');
//-->
</script>
<%
End If 
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->