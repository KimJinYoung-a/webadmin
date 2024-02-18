<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 핑거스 업체 배송
' Hieditor : 2016.07.21 한용민 텐바이텐 업체 배송 이전/수정 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/jumun/baljucls.asp"-->
<%
dim idxArr, ix,sqlStr, listitemlist,listitem,listitemcount, iSall, SheetType
	idxArr = request.Form("idxArr")
	listitem =  request("orderserial")
	iSall   =  RequestCheckvar(request("isall"),10)
  	if idxArr <> "" then
		if checkNotValidHTML(idxArr) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end if
If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('세션이 종료되었습니다.');</script>"
    dbget.close()	:	response.End
end if

''기본택배사
dim defaultSongjangdiv
sqlStr = "select defaultSongjangdiv from [db_partner].[dbo].tbl_partner"
sqlStr = sqlStr + " where id='" & session("ssBctID") & "'"
rsget.Open sqlStr,dbget,1
if Not Rsget.Eof then
    defaultSongjangdiv = rsget("defaultSongjangdiv")

    if IsNULL(defaultSongjangdiv) then defaultSongjangdiv=""
end if
rsget.close

dim ojumun
set ojumun = new CJumunMaster
	ojumun.FRectOrderSerial = idxArr
	ojumun.FRectDesignerID = session("ssBctID")
	ojumun.FRectIsAll       = iSall
	ojumun.ReDesignerSelectBaljuList

function ReplaceSCVStr(oStr)
    ReplaceSCVStr = ""
    if IsNULL(oStr) then Exit function
    ReplaceSCVStr = Replace(oStr, chr(34),"'")
    ReplaceSCVStr = Replace(oStr, ",","")
end function

Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENDLV_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<title>송장파일</title>
<style>
    <!--
	br
	    {mso-data-placement:same-cell;}
	tr
	    {mso-height-source:auto;
	    mso-ruby-visibility:none;}
	td
	    {white-space:normal;}
	-->
</style>
</head>

<body leftmargin="10">
<table border=1>
<tr>
    <td align="center" x:str >일련번호</td>
	<td align="center" height="25" x:str >주문번호</td>
	<td align="center" x:str >구매자명</td>
	<td align="center" x:str >수령인</td>
	<td align="center" x:str >상품명</td>
	<td align="center" x:str >옵션명</td>
	<td align="center" x:str >택배사코드</td>
	<td align="center" x:str >송장번호</td>
</tr>

<% if ojumun.FResultCount >0 then %>
	<% for ix=0 to ojumun.FResultCount - 1 %>
	<tr>
	    <td align="center" x:str><%= ojumun.FMasterItemList(ix).FDetailidx %></td>
		<td align="center" x:str><%= ojumun.FMasterItemList(ix).FOrderSerial %></td>
		<td align="center" x:str><%= ReplaceSCVStr(ojumun.FMasterItemList(ix).FBuyName) %></td>
		<td align="center" x:str><%= ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqName) %></td>
		<td align="center" x:str><%= ReplaceSCVStr(ojumun.FMasterItemList(ix).FItemName) %></td>
		<td align="center" x:str><%= ReplaceSCVStr(ojumun.FMasterItemList(ix).FItemoptionName) %></td>
		<td align="center" x:str>
			<% if IsNULL(ojumun.FMasterItemList(ix).FSongjangdiv) then %>
				<%= defaultSongjangdiv %>
			<% end if %>
		</td>
		<td align="center" x:str><%= ojumun.FMasterItemList(ix).Fsongjangno %></td>
	</tr>
	<% next %>
<% else %>
	<tr colspan=8>
	    <td align="center" x:str >검색 결과가 없습니다.</td>
	</tr>
<% end if %>

</table>
</body>
</html>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->