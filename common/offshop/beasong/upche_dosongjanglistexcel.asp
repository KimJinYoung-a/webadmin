<%@ language=vbscript %>
<% option explicit %>

<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.02.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
If session("ssBctId") = "" then
    response.write "<script language='javascript'>alert('������ ����Ǿ����ϴ�.');</script>"
    dbget.close()	:	response.End
end if

dim detailidxArr , ojumun ,ix,sqlStr ,listitem
dim iSall, SheetType
	detailidxArr = request.Form("detailidxArr")
	listitem =  request("orderserial")
	iSall   =  request("isall")


''�⺻�ù��
dim defaultSongjangdiv
sqlStr = "select defaultSongjangdiv from [db_partner].[dbo].tbl_partner"
sqlStr = sqlStr + " where id='" & session("ssBctID") & "'"

'response.write sqlStr &"<Br>"
rsget.Open sqlStr,dbget,1
if Not Rsget.Eof then
    defaultSongjangdiv = rsget("defaultSongjangdiv")

    if IsNULL(defaultSongjangdiv) then defaultSongjangdiv=""
end if
rsget.close

set ojumun = new cupchebeasong_list
	ojumun.frectdetailidxarr = detailidxArr
	ojumun.FRectDesignerID = session("ssBctID")
	ojumun.FRectIsAll       = iSall
	ojumun.fReDesignerSelectBaljuList()

function ReplaceSCVStr(oStr)
    ReplaceSCVStr = ""
    if IsNULL(oStr) then Exit function
    ReplaceSCVStr = Replace(oStr, chr(34),"'")
    ReplaceSCVStr = Replace(oStr, ",","")
end function

'Response.Buffer=False
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

	br
	    {mso-data-placement:same-cell;}
	tr
	    {mso-height-source:auto;
	    mso-ruby-visibility:none;}
	td
	    {white-space:normal;}

</style>
</head>

<body leftmargin="10">
<table border=1>
<tr>
    <td align="center" x:str >�Ϸù�ȣ</td>
	<td align="center" height="25" x:str >�ֹ���ȣ</td>
	<td align="center" x:str >������</td>
	<td align="center" x:str >��ǰ��</td>
	<td align="center" x:str >�ɼǸ�</td>
	<td align="center" x:str >�ù���ڵ�</td>
	<td align="center" x:str >�����ȣ</td>

</tr>
<% for ix=0 to ojumun.FResultCount - 1 %>
<tr>
    <td align="center" x:str><%= ojumun.FItemList(ix).FDetailidx %></td>
	<td align="center" x:str><%= ojumun.FItemList(ix).forderno %></td>
	<td align="center" x:str><%= ReplaceSCVStr(ojumun.FItemList(ix).FReqName) %></td>
	<td align="center" x:str><%= ReplaceSCVStr(ojumun.FItemList(ix).FItemName) %></td>
	<td align="center" x:str><%= ReplaceSCVStr(ojumun.FItemList(ix).FItemoptionName) %></td>
	<td align="center" x:str>
	<% if IsNULL(ojumun.FItemList(ix).FSongjangdiv) then %>
	<%= defaultSongjangdiv %>
	<% end if %>
	</td>
	<td align="center" x:str><%= ojumun.FItemList(ix).Fsongjangno %></td>
</tr>
<% next %>
</table>
</body>
</html>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->