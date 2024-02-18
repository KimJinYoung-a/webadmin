<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
 Response.AddHeader "Pragma","no-cache"   
 Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###########################################################
' Page : /admin/itemmaster/deal/dodealitemreg.asp
' Description :  �� ��ǰ - ���, ����
' History : 2017.08.28 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/dealManageCls.asp"-->
<%
'--------------------------------------------------------
' �������� & �Ķ���� �� �ޱ�
'--------------------------------------------------------
Dim idx : idx = requestCheckVar(Request("idx"),9)

Dim oDealitem, arrList, iTotCnt, intLoop
set oDealitem = new CDealItem
oDealitem.FRectMasterIDX = idx
arrList = oDealitem.fnGetDealEventItem	
%>
<div id="divIpG">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>����</td>
	<td>��ǰ�ڵ�</td>
	<td>��ǰ��</td>
	<td>�ǸŰ�</td>
	<td>���԰�</td>
	<td>������</td>
</tr>
<% If isArray(arrList) Then %>
<% For intLoop = 0 To UBound(arrList,2) %>
<tr bgcolor="#FFFFFF" align="center">
	<td><%=arrList(0,intLoop)%></td>
	<td><a href="javascript:editItemBasicInfo('<%=arrList(1,intLoop)%>')"><%=arrList(1,intLoop)%></a></td>
	<td><%=arrList(2,intLoop)%></td>
	<td>
		<%
			Response.Write FormatNumber(arrList(5,intLoop),0)
			'���ΰ�
			if arrList(9,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>(��)" & FormatNumber(arrList(7,intLoop),0) & "</font>"
			end if
			'������
			if arrList(10,intLoop)="Y" then
				Select Case arrList(11,intLoop)
					Case "1"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(4,intLoop)*((100-arrList(12,intLoop))/100),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(4,intLoop)-arrList(12,intLoop),0) & "</font>"
				end Select
			end if
		%>
	</td>
	<td>
		<%
			Response.Write FormatNumber(arrList(6,intLoop),0)
			'���ΰ�
			if arrList(9,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(8,intLoop),0) & "</font>"
			end if
			'������
			if arrList(10,intLoop)="Y" then
				if arrList(12,intLoop)="1" or arrList(12,intLoop)="2" then
					if arrList(13,intLoop)=0 or isNull(arrList(13,intLoop)) then
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(6,intLoop),0) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(13,intLoop),0) & "</font>"
					end if
				end if
			end if
		%>
	</td>
	<td>
		<a href="javascript:fnSaleInfo();"><%if arrList(9,intLoop)="Y" then%>
		<font color="#F08050"><%=CLng(((arrList(5,intLoop)-arrList(7,intLoop))/arrList(5,intLoop))*100)%>%</font>		
		<%end if%>
		<%if arrList(10,intLoop)="Y" then 
		if arrList(12,intLoop)="1" or arrList(12,intLoop)="2" then
			if arrList(13,intLoop)=0 or isNull(arrList(13,intLoop)) then
				 Response.Write "<br><font color=#5080F0>" & FormatNumber( arrList(6,intLoop),0) & "</font>"
			else
				Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(12,intLoop),0) 
				 if arrList(12,intLoop)="1" then 
				 Response.Write "%"
				else
				 Response.Write "��"
				end if
				 Response.Write "</font>"
			end if
		end if
		end if%></a>
	</td>
</tr>
<% Next %>
<% End If %>
</table>
</div>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
$("#itemcode", opener.document).attr("disabled",false);
<% If isArray(arrList) Then %>
<% For intLoop = 0 To UBound(arrList,2) %>
$("#itemcode option:eq(<%=intLoop+1%>)", opener.document).remove();
$("#itemcode", opener.document).append("<option value='<%=arrList(1,intLoop)%>'><%=arrList(2,intLoop)%></option>");
<% Next %>
<% End If %>
</script>
<script type="text/javascript">
opener.document.all.divForm.style.display = "none";
$("#divFrm3", opener.document).html($("#divIpG").html()); 
opener.document.all.saleper1.style.display = "";
opener.document.all.saleper2.style.display = "";
$("#saleper1", opener.document).css("display",true);
$("#saleper2", opener.document).css("display",true);
self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->