<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.02.26 �ѿ�� ����
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
	detailidxarr =  request("detailidxarr")
	iSall   =  request("isall")
	SheetType  =  request("SheetType")

If session("ssBctId") = "" then
    response.write "<script language='javascript'>alert('������ ����Ǿ����ϴ�.');</script>"
    dbget.close()	:	response.End
end if

function replaceXlText(org)
    dim reText
    reText = replace(org,"<","&lt;")
    replaceXlText = replace(reText,">","&gt;")
end function

set ojumun = new cupchebeasong_list
	ojumun.FRectdetailidxarr = detailidxarr
	ojumun.FRectIsAll       = iSall
	ojumun.FRectDesignerID = session("ssBctID")
	ojumun.fReDesignerSelectBaljuList()

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
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
<table width=1200 cellspacing=0 cellpadding=1 border=0>
<tr>
    <% if (SheetType="V4") then %>
    <td align="center" height="25" x:str >�Ϸù�ȣ</td>
    <% end if %>
	<td align="center" x:str >�ֹ���ȣ</td>
	<td align="center" x:str >�ֹ���</td>
	<td align="center" x:str >������</td>
	<td align="center" x:str >��������ȭ</td>
	<td align="center" x:str >�������ڵ���</td>
	<td align="center" x:str >�����ȣ</td>

	<% if (SheetType="V2") then %>
		<td align="center" x:str >������ּ�</td>
	<% else %>
		<td align="center" x:str >������ּ�1</td>
		<td align="center" x:str >������ּ�2</td>
	<% end if %>

	<td align="center" x:str >������ǻ���</td>
	<td align="center" x:str >�ù��ȣ</td>
	<td align="center" x:str >��ǰ���̵�</td>
	<td align="center" x:str >��ǰ��</td>
	<td align="center" x:str >�ɼ�</td>
	<td align="center" x:str >�ǸŰ�</td>
	<td align="center" x:str >����</td>
</tr>
<% for ix=0 to ojumun.FResultCount - 1 %>
<tr>
    <% if (SheetType="V4") then %>
    <td align="center" x:str><%= ojumun.FItemList(ix).Fdetailidx %></td>
    <% end if %>
	<td align="center" x:str><%= ojumun.FItemList(ix).Forderno %></td>
	<td align="center" x:str><%= Left(CStr(ojumun.FItemList(ix).FRegDate),10) %></td>
	<td align="center" x:str><%= ojumun.FItemList(ix).FReqName %></td>
	<td align="center" x:str><%= ojumun.FItemList(ix).FReqPhone %></td>
	<td align="center" x:str><%= ojumun.FItemList(ix).FReqHp %></td>
	<td align="center" x:str><%= ojumun.FItemList(ix).FReqZipCode %></td>

	<% if (SheetType="V2") then %>
		<td align="center" x:str><%= ojumun.FItemList(ix).FReqZipAddr %><%=chr(32)%><%= ojumun.FItemList(ix).FReqAddress %></td>
	<% else %>
		<td align="center" x:str><%= ojumun.FItemList(ix).FReqZipAddr %></td>
		<td align="center" x:str><%= ojumun.FItemList(ix).FReqAddress %></td>
	<% end if %>

	<td align="center" x:str><%= db2html(ojumun.FItemList(ix).FComment) %></td>
	<td align="center" x:str><%= ojumun.FItemList(ix).Fsongjangno %></td>
	<td align="center" x:str><%= ojumun.FItemList(ix).Fitemid %></td>
	<td align="center" x:str><%= ojumun.FItemList(ix).FItemName %></td>
	<td align="center" x:str><%= ojumun.FItemList(ix).FItemoptionName %></td>
	<td align="center" x:num="<%= ojumun.FItemList(ix).fsellprice %>" ><%= ojumun.FItemList(ix).fsellprice %></td>
	<td align="center" x:num="<%= ojumun.FItemList(ix).FItemNo %>" ><%= ojumun.FItemList(ix).FItemNo %></td>
</tr>
<% next %>
</table>
</body>
</html>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->