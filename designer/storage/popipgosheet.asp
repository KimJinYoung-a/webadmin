<%
'Response.AddHeader "Cache-Control","no-cache"
'Response.AddHeader "Expires","0"
'Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  �ֹ�������
' History : 		   �̻� ����
'			2016.08.17 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim idx,itype, ibrandname
idx = requestCheckVar(request("idx"),20)
itype = requestCheckVar(request("itype"),50)
ibrandname = requestCheckVar(request("ibrandname"),100)

'==============================================================================
dim oipchul, oipchuldetail
set oipchul = new CIpChulStorage
oipchul.FRectId = idx
oipchul.GetIpChulMaster

set oipchuldetail = new CIpChulStorage
oipchuldetail.FRectStoragecode = oipchul.FOneItem.Fcode
oipchuldetail.GetIpChulDetail

'==============================================================================
dim obrand
set obrand = new CPartnerUser
obrand.FRectDesignerID = oipchul.FOneItem.Fsocid
obrand.GetOnePartnerNUser



dim i

dim executedate

if (oipchul.FOneItem.Fexecutedt <> "") then
	executedate = replace(Left(CstR(oipchul.FOneItem.Fexecutedt),10),"-","/")
else
	executedate = "<font color='red'>���԰�</font>"
end if

dim ttlsellcash, ttlsuplycash, ttlcount
ttlsellcash = 0
ttlsuplycash  = 0
ttlcount    = 0
%>
<%
if request("xl")<>"" then
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" + oipchul.FOneItem.Fsocid + Left(CStr(now()),10) + ".xls"
end if
%>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
	    <td colspan="3">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a">
				<tr>
			    	<td style="font-size:12pt; font-family:����, arial;"><b>�԰�����(<%= obrand.FOneItem.FSocName_Kor %>)</b></td>
					<td align="right">
			    		<b>�԰��ڵ� (<%= oipchul.FOneItem.Fcode %>)</b>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr height="1">
		<td colspan="3" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
	<tr valign="top" style="padding:10 0 0 0">
        <td width="49%">
        	<!-- ���������� ���� -->
        	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4"><b>������ ����</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>��Ϲ�ȣ</td>
        			<td colspan="3"><%= obrand.FOneItem.Fcompany_no %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="60">��ȣ</td>
        			<td width="135"><b><%= obrand.FOneItem.Fcompany_name %></b></td>
        			<td width="60">��ǥ��</td>
        			<td width="90"><%= obrand.FOneItem.FCeoname %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>������</td>
        			<td colspan="3"><%= obrand.FOneItem.Faddress %>&nbsp;<%= obrand.FOneItem.Fmanager_address %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>����</td>
        			<td><%= obrand.FOneItem.Fcompany_uptae %></td>
        			<td>����</td>
        			<td><%= obrand.FOneItem.Fcompany_upjong %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>�����</td>
        			<td><%= obrand.FOneItem.Fmanager_name %></td>
        			<td>����ó</td>
        			<td><%= obrand.FOneItem.Fmanager_hp %></td>
        		</tr>
        	</table>
        	<!-- ���������� �� -->
        </td>
        <td>&nbsp;</td>
        <td width="49%">
        	<!-- ���޹޴������� ���� -->
        	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        			<td colspan="4"><b>���޹޴��� ����</b></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>��Ϲ�ȣ</td>
        			<td colspan="3">211-87-00620</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="60">��ȣ</td>
        			<td width="135">(��)�ٹ�����</td>
        			<td width="60">��ǥ��</td>
        			<td width="90">������</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>������</td>
        			<td colspan="3">����� ���α� ������ 1-45 �������� 2��</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>����</td>
        			<td>����,���Ҹ� ��</td>
        			<td>����</td>
        			<td>���ڻ�ŷ� ��</td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>&nbsp;</td>
        			<td></td>
        			<td></td>
        			<td></td>
        		</tr>
        	</table>
        	<!-- ���޹޴������� �� -->
        </td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
		<td colspan="15">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">&nbsp;<strong>�԰�󼼳���</strong>
			        </td>
			       	<td align="right">
			       		<b>�԰����� : <%= executedate %></b>
			    	</td>
			    </tr>
			</table>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="90">��ǰ�ڵ�</td>
    	<td>��ǰ��</td>
    	<td>�ɼǸ�</td>
    	<td width="60">�Һ��ڰ�</td>
    	<td width="60">���ް�</td>
    	<td width="50">����</td>
    	<td width="70">���ް��հ�</td>
    </tr>


	 <% for i=0 to oipchuldetail.FResultCount -1 %>
	 <%
	 	ttlsellcash = ttlsellcash + oipchuldetail.FItemList(i).Fitemno*oipchuldetail.FItemList(i).Fsellcash
	 	ttlsuplycash = ttlsuplycash + oipchuldetail.FItemList(i).Fitemno*oipchuldetail.FItemList(i).Fsuplycash
	 	ttlcount = ttlcount + oipchuldetail.FItemList(i).Fitemno
	 %>

	<tr align="center" bgcolor="#FFFFFF">
<!--	<td><%= oipchuldetail.FItemList(i).FIMakerid %></td>	-->
		<td><%= oipchuldetail.FItemList(i).Fiitemgubun %>-<b><%= CHKIIF(oipchuldetail.FItemList(i).FItemId>=1000000,Format00(8,oipchuldetail.FItemList(i).FItemId),Format00(6,oipchuldetail.FItemList(i).FItemId)) %></b>-<%= oipchuldetail.FItemList(i).FItemOption %></td>
		<td align="left"><%= oipchuldetail.FItemList(i).FIItemName %></td>
		<td><%= oipchuldetail.FItemList(i).FIItemoptionName %></td>
		<td align="right"><%= FormatNumber(oipchuldetail.FItemList(i).Fsellcash,0) %></td>
		<td align="right"><%= FormatNumber(oipchuldetail.FItemList(i).Fsuplycash,0) %></td>
		<td><%= oipchuldetail.FItemList(i).Fitemno %></td>
		<td align="right"><%= FormatNumber(oipchuldetail.FItemList(i).Fitemno*oipchuldetail.FItemList(i).Fsuplycash,0) %></td>
	<% next %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="#FFFFFF">���</td>
		<td colspan="3" align="left" bgcolor="#FFFFFF"><%= nl2br(oipchul.FOneItem.Fcomment) %></td>
		<td><b>�Ѱ�</b></td>
		<td><b><%= ttlcount %></b></td>
		<td align="right"><b><%= ForMatNumber(ttlsuplycash,0) %></b></td>
	</tr>
</table>


<%
set obrand = Nothing
set oipchul = Nothing
set oipchuldetail = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
