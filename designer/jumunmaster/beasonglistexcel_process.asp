<%@ language=vbscript %>
<% option explicit %>
<%

'response.expires = -1
'response.AddHeader "Pragma", "no-cache"
'response.AddHeader "cache-control", "no-store"

response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" + Left(CStr(now()),10) + ".xls"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_baljucls.asp"-->
<%

''Not Using File
1 Raize Error

If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('������ ����Ǿ����ϴ�.');</script>"
    dbget.close()	:	response.End
end if
dim ojumun
dim ix,sql
Dim listitemlist,listitem,listitemcount
dim iSall

listitem =  Replace(request("orderserial"), " ", "")
iSall   =  requestCheckVar(request("isall"), 32)

set ojumun = new CJumunMaster

ojumun.FRectOrderSerial = listitem
ojumun.FRectIsAll       = iSall
ojumun.FRectDesignerID = session("ssBctID")
ojumun.ReDesignerSelectBaljuList

%>
<html xmlns:x="urn:schemas-microsoft-com:office:excel">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<style>
 br
	{mso-data-placement:same-cell;}
  .big_title
    {
    mso-style-parent:style0;
	white-space:normal;
    font-size:18.0pt;
    font-weight:700;
    }
  .mid_title
    {
    mso-style-parent:style0;
	white-space:normal;
    font-size:12.0pt;
    font-weight:700;
    }
  .title_center
	{mso-style-parent:style0;
	text-align:center;
	vertical-align:middle;
	white-space:normal;}
  .normal
	{
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-style-parent:style0;
	vertical-align:middle;
	white-space:normal;
	font-size:8.0pt;
	}
  .normal_b
	{
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-style-parent:style0;
	vertical-align:middle;
	white-space:normal;
	font-size:8.0pt;
	font-weight:700;
	}
  .currency
	{mso-style-parent:style0;
 	mso-number-format:"\#\,\#\#0\.00_\)\;\[Red\]\\\(\#\,\#\#0\.00\\\)";
	border:0.5pt solid black;
	white-space:normal;}
   .Format_Y1
	{mso-style-parent:style0;
	mso-number-format:"yyyy\0022\/\0022m\0022\/\0022d\;\@";
 	white-space:normal;}
   .Format_Y2
	{mso-style-parent:style0;
	mso-number-format:"yyyy\/mm\;\@";
	text-align:center;
	border:0.5pt solid black;
 	white-space:normal;}
   .Format_number
	{
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-style-parent:style0;
	vertical-align:middle;
	mso-number-format:"\#\,\#\#0";
	white-space:normal;
	font-size:8.0pt;
	}
   .Format_number_L
	{
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-style-parent:style0;
	vertical-align:middle;
	mso-number-format:"\#\,\#\#0";
	white-space:normal;
	font-size:12.0pt;
	}
  .Format_T1
	{mso-style-parent:style0;
	mso-number-format:"hh\:mm\:ss\;\@";
	text-align:center;
 	white-space:normal;}  </style>
</head>

<body leftmargin="10">
<table width=1200 cellspacing=0 cellpadding=1 border=0>
<tr>
	<td align="center" height="25" class=normal >�ֹ���ȣ</td>
	<td align="center" class=normal >�ֹ���</td>
	<td align="center" class=normal >�����ڸ�</td>
	<td align="center" class=normal >��������ȭ</td>
	<td align="center" class=normal >�������ڵ���</td>
	<td align="center" class=normal >�������̸���</td>
	<td align="center" class=normal >������</td>
	<td align="center" class=normal >��������ȭ</td>
	<td align="center" class=normal >�������ڵ���</td>
	<td align="center" class=normal >�����ȣ</td>
	<td align="center" class=normal >������ּ�1</td>
	<td align="center" class=normal >������ּ�2</td>
	<td align="center" class=normal >������ǻ���</td>
	<td align="center" class=normal >�ù��ȣ</td>
	<td align="center" class=normal >��ǰ���̵�</td>
	<td align="center" class=normal >��ǰ��</td>
	<td align="center" class=normal >�ɼ�</td>
	<td align="center" class=normal >�ǸŰ�</td>
	<td align="center" class=normal >����</td>
	<td align="center" class=normal >�ֹ����۸޼���</td>
<% if Not IsNULL(ojumun.FMasterItemList(0).Freqdate) then %>
	<td align="center" class=normal >��������</td>
	<td align="center" class=normal >ī�帮��</td>
	<td align="center" class=normal >�޼���</td>
	<td align="center" class=normal >�����»��</td>
<% end if %>
</tr>
<% for ix=0 to ojumun.FResultCount - 1 %>
<tr>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).FOrderSerial %></td>
	<td align="center" class=normal><%= Left(CStr(ojumun.FMasterItemList(ix).FRegDate),10) %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).FBuyName %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).FBuyPhone %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).FBuyHp %></td>
	<td align="center" class=normal></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).FReqName %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).FReqPhone %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).FReqHp %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).FReqZipCode %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).FReqZipAddr %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).FReqAddress %></td>
	<td align="center" class=normal><%= db2html(ojumun.FMasterItemList(ix).FComment) %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).Fsongjangno %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).Fitemid %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).FItemName %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).FItemoptionName %></td>
	<td align="center" class=Format_number ><%= FormatNumber(ojumun.FMasterItemList(ix).FItemCost,0) %></td>
	<td align="center" class=Format_number ><%= ojumun.FMasterItemList(ix).FItemNo %></td>
	<td align="center" class=normal ><%= nl2br(ojumun.FMasterItemList(ix).Frequiredetail) %></td>
<% if Not IsNULL(ojumun.FMasterItemList(ix).Freqdate) then %>
	<td align="center" class=normal><%= Left(CStr(ojumun.FMasterItemList(ix).Freqdate),10) %>�� <%= (ojumun.FMasterItemList(ix).GetReqTimeText) %></td>
	<td align="center" class=normal><%= ojumun.FMasterItemList(ix).getCardribbonName %></td>
	<td align="center" class=normal><%= nl2br(db2html(ojumun.FMasterItemList(ix).Fmessage)) %></td>
	<td align="center" class=normal><%= db2html(ojumun.FMasterItemList(ix).Ffromname) %></td>
<% end if %>
</tr>
<% next %>
</table>
</body>
</html>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->