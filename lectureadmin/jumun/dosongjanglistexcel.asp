<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ΰŽ� ��ü ���
' Hieditor : 2016.07.21 �ѿ�� �ٹ����� ��ü ��� ����/���� ����
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
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end if
If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('������ ����Ǿ����ϴ�.');</script>"
    dbget.close()	:	response.End
end if

''�⺻�ù��
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
<title>��������</title>
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
    <td align="center" x:str >�Ϸù�ȣ</td>
	<td align="center" height="25" x:str >�ֹ���ȣ</td>
	<td align="center" x:str >�����ڸ�</td>
	<td align="center" x:str >������</td>
	<td align="center" x:str >��ǰ��</td>
	<td align="center" x:str >�ɼǸ�</td>
	<td align="center" x:str >�ù���ڵ�</td>
	<td align="center" x:str >�����ȣ</td>
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
	    <td align="center" x:str >�˻� ����� �����ϴ�.</td>
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