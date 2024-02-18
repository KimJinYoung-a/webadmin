<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �ֹ�������
' History : 		   �̻� ����
'			2016.08.17 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim idx,itype
idx = requestCheckVar(request("idx"),20)
itype = requestCheckVar(request("itype"),50)


dim oordersheetmaster, oordersheet
set oordersheetmaster = new COrderSheet
oordersheetmaster.FRectIdx = idx
oordersheetmaster.GetOneOrderSheetMaster

dim isFixed
isFixed = oordersheetmaster.FOneItem.IsFixed


set oordersheet = new COrderSheet
oordersheet.FrectisFixed = isFixed
oordersheet.FRectIdx = idx
oordersheet.GetOrderSheetDetail


dim obrand
set obrand = new CBrandShopInfoItem

obrand.FRectChargeId = oordersheetmaster.FOneItem.Ftargetid
obrand.GetBrandShopInFo


dim i

dim scheduleorexedate
if not IsNULL(oordersheetmaster.FOneItem.FScheduleDate) then
scheduleorexedate = replace(Left(CstR(oordersheetmaster.FOneItem.FScheduleDate),10),"-","/")
end if

dim ttlsellcash, ttlbuycash, ttlcount
ttlsellcash = 0
ttlbuycash  = 0
ttlcount    = 0

function getObjStr(v)
	dim reStr
	reStr = "<OBJECT" + vbCrlf
	reStr = reStr + "id=iaxobject" + vbCrlf
	reStr = reStr + "classid='clsid:A4F3A486-2537-478C-B023-F8CCC41BF29D'" + vbCrlf
	reStr = reStr + "codebase='http://partner.10x10.co.kr/cab/tenbarShow.cab#version=1,0,0,3'" + vbCrlf
	reStr = reStr + "width=100" + vbCrlf
	reStr = reStr + "height=20" + vbCrlf
	reStr = reStr + "align=bottom" + vbCrlf
	reStr = reStr + "hspace=0" + vbCrlf
	reStr = reStr + "vspace=0" + vbCrlf
	reStr = reStr + ">" + vbCrlf
	reStr = reStr + "</OBJECT>" + vbCrlf

	getObjStr = reStr
end function

%>

<%
if request("xl")<>"" then
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" + oordersheetmaster.FOneItem.Ftargetid + Left(CStr(now()),10) + ".xls"
end if
%>



<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
	    <td colspan="3">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a">
				<tr>
			    	<td style="font-size:12pt; font-family:����, arial;"><b>�ŷ�����(<%= oordersheetmaster.FOneItem.Ftargetid %>)</b></td>
					<td align="right">
			    		<img src="http://company.10x10.co.kr/barcode/barcode.asp?image=3&type=23&data=<%= oordersheetmaster.FOneItem.FBaljuCode %>&height=20&barwidth=1&TextAlign=2" <%=CHKIIF(LCASE(session("ssBctId"))="smlgroup","onClick='this.remove()'","") %>>
			    		&nbsp;&nbsp;&nbsp;&nbsp;
			    		<b>�ֹ��ڵ� (<%= oordersheetmaster.FOneItem.FBaljuCode %>)</b>
			<!--       	&nbsp;<%= getObjStr("oordersheetmaster.FOneItem.FBaljuCode") %>	-->
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
        			<td colspan="3"><%= obrand.FSocNo %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td width="60">��ȣ</td>
        			<td width="135"><b><%= obrand.FChargeName %></b></td>
        			<td width="60">��ǥ��</td>
        			<td width="90"><%= obrand.FCeoName %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>������</td>
        			<td colspan="3"><%= obrand.FAddress %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>����</td>
        			<td><%= obrand.FUptae %></td>
        			<td>����</td>
        			<td><%= obrand.FUpjong %></td>
        		</tr>
        		<tr align="center" bgcolor="#FFFFFF">
        			<td>�����</td>
        			<td><%= obrand.FManagerName %></td>
        			<td>����ó</td>
        			<td><%= obrand.FManagerHp %></td>
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
        			<td>�ֹ���</td>
        			<td><%= oordersheetmaster.FOneItem.Fregname %></td>
        			<td>����ó</td>
        			<td>1644-1851</td>
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
						<img src="/images/icon_arrow_down.gif" align="absbottom">&nbsp;<strong>�󼼳���</strong>
			        	<b>(�Ѿ� : \<%= ForMatNumber(oordersheetmaster.FOneItem.FTotalBuycash,0) %>)</b>
			        </td>
			       	<td align="right">
			       		<b>�ֹ����� : <%= scheduleorexedate %></b>
			    	</td>
			    </tr>
			</table>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="90">��ǰ�ڵ�</td>
    	<td>��ǰ��</td>
    	<td>�ɼǸ�</td>
    	<td width="55">�Һ��ڰ�</td>
    	<td width="55">���ް�</td>
    	<td width="50">����</td>
    	<td width="70">���ް��հ�</td>
    </tr>

	<% for i=0 to oordersheet.FResultCount -1 %>
	<%
		ttlsellcash = ttlsellcash + oordersheet.FItemList(i).Frealitemno*oordersheet.FItemList(i).FSellcash
		ttlbuycash = ttlbuycash + oordersheet.FItemList(i).Frealitemno*oordersheet.FItemList(i).FBuycash
		ttlcount = ttlcount + oordersheet.FItemList(i).Frealitemno
	%>

	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oordersheet.FItemList(i).FItemGubun %>-<b><%= CHKIIF(oordersheet.FItemList(i).FItemId>=1000000,Format00(8,oordersheet.FItemList(i).FItemId),Format00(6,oordersheet.FItemList(i).FItemId)) %></b>-<%= oordersheet.FItemList(i).FItemOption %></td>
		<td align="left"><%= left(oordersheet.FItemList(i).FItemName,35) %></td>
		<td><%= left(oordersheet.FItemList(i).FItemOptionName,15) %></td>
		<td align="right"><%= FormatNumber(oordersheet.FItemList(i).Fsellcash,0) %></td>
		<td align="right"><%= FormatNumber(oordersheet.FItemList(i).FBuycash,0) %></td>
		<td><%= oordersheet.FItemList(i).FRealItemno %></td>
		<td align="right"><%= FormatNumber(oordersheet.FItemList(i).FRealItemno*oordersheet.FItemList(i).FBuycash,0) %></td>
	</tr>
	<% next %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td bgcolor="#FFFFFF">���</td>
		<td colspan="3" align="left" bgcolor="#FFFFFF"><%= nl2br(oordersheetmaster.FoneItem.FComment) %></td>
		<td><b>�Ѱ�</b></td>
		<td><b><%= ttlcount %></b></td>
		<td align="right"><b><%= ForMatNumber(ttlbuycash,0) %></b></td>
	</tr>
	<tr height="25" bgcolor="<%= adminColor("topbar") %>">
		<td colspan="15">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td width="50%" align="left"><b>�ΰ��� :</b></td>
			       	<td><b>�μ��� :</b></td>
			    </tr>
			</table>
		</td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			* �ٹ����� �������� �ּ� : <b>[11154] ��⵵ ��õ�� ������ ����������2�� 83 �ٹ����� ��������</b> (����ó : 1644-1851)</b>
		</td>
	</tr>
</table>


<script language='javascript'>
//iaxobject.ShowBarCode(30,'<%= oordersheetmaster.FOneItem.FBaljuCode %>',2);

function getOnLoad(){
   window.print();
}
window.onload=getOnLoad;
</script>

<%
set obrand = Nothing
set oordersheetmaster = Nothing
set oordersheet = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
