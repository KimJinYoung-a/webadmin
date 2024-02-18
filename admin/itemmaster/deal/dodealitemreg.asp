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
Dim k, sqlStr, i
Dim vCnt : vCnt = Request.Form("cksel").count
Dim idx : idx = requestCheckVar(Request.Form("idx"),9)
Dim stype : stype = requestCheckVar(Request.Form("stype"),1)
Dim upback : upback = requestCheckVar(Request.Form("upback"),1)

if Request.Form("cksel") <> "" then
	if checkNotValidHTML(Request.Form("cksel")) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
if Request.Form("sitemname") <> "" then
	if checkNotValidHTML(Request.Form("sitemname")) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if

'�迭�� ó��
redim arritemcode(vCnt)
redim arritemname(vCnt)
for i=1 to vCnt
	arritemcode(i) = Request.Form("cksel")(i)
	arritemname(i) = Request.Form("sitemname")(i)
next

If vCnt >= 1 Then
dbget.beginTrans
		sqlStr = " delete FROM [db_event].[dbo].[tbl_deal_event_item] WHERE dealcode=" & idx
		dbget.execute sqlStr
	For k=1 To vCnt
		sqlStr = " IF Not Exists(SELECT IDX FROM [db_event].[dbo].[tbl_deal_event_item] WHERE itemid='" & arritemcode(k) & "' and dealcode="&idx&")"			
		sqlStr = sqlStr + "	BEGIN "
		sqlStr = sqlStr+ " 			INSERT INTO [db_event].[dbo].[tbl_deal_event_item] (dealcode, itemid, itemname, viewidx)"
		sqlStr = sqlStr + "     	VALUES (" & idx & ", " & arritemcode(k) &",'" & arritemname(k) & "'," & k & ")"
		sqlStr = sqlStr + " 	END "
		sqlStr = sqlStr + " ELSE "
		sqlStr = sqlStr + " 	BEGIN "			
		sqlStr = sqlStr + "			UPDATE [db_event].[dbo].[tbl_deal_event_item]"
		sqlStr = sqlStr + " 		SET viewidx ='" & k & "'"
		sqlStr = sqlStr + " 		WHERE dealcode = '" & idx & "' "
		sqlStr = sqlStr + " 		and itemid ="&arritemcode(k)&""
		sqlStr = sqlStr + " 	END "
		dbget.execute sqlStr
	IF Err.Number <> 0 THEN
		dbget.RollBackTrans 
		Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.", "back", "")
		response.End 
	END IF
	Next
	dbget.CommitTrans
End If

'������Ű�� ����?? ����
if idx="" then
	Response.End
end if

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
	<td><%=arrList(1,intLoop)%></td>
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
		<%if arrList(9,intLoop)="Y" then%>
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
		end if%>
	</td>
</tr>
<% Next %>
<% End If %>
</table>
</div>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript">
$("#divFrm3", opener.document).html($("#divIpG").html()); 
opener.document.all.divForm.style.display = "none"; 
self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->