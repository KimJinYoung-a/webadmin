<%
'###########################################################
' Description : ��÷�� ���ó��
' History : 2010.03.22 �ѿ�� ����
'###########################################################

Dim arrPrize , arrPrizeType, arrPrizeStatus , cEvtPrize  ,page ,i
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)

if page = "" then page = 1
		
set cEvtPrize = new cevent_list
cEvtPrize.FPageSize = 100
cEvtPrize.FCurrPage = page	
cEvtPrize.frectevt_code	= evt_code			'�̺�Ʈ �ڵ�
cEvtPrize.fnGetPrize_off()

arrPrizeType = fnSetCommonCodeArr_off("evtprize_type",False)
arrPrizeStatus= fnSetCommonCodeArr_off("evtprize_status",False)
		
%>
<script language="javascript">

	//��÷�� ���
	function jsSetWinner(evt_code,epC){
		var winW, popURL;
		if (epC > 0){
			popURL ="/admin/eventmanage/event/pop_event_changewinner.asp?epC="+epC;  		
		}else{
			popURL="/admin/offshop/event_off/pop_event_winner.asp?evt_code="+evt_code;
		}
		winW = window.open(popURL,'popW','width=630, height=500, scrollbars=yes');
		winW.focus();
	}
  
</script>

<table width="100%" border="0" align="left" class="a" cellpadding="0" cellspacing="1">
	<tr>
		<td>
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1">
			<tr>
				<td>
					<input type="button" name="btnadd"  value="�� ��÷���" onClick="javascript:jsSetWinner(<%=evt_code%>,0);" class="button">
					<input type="button" value="�̺�Ʈ�������ε��ư���" onClick="location.href='index.asp?evt_code=<%=evt_code%>&menupos=<%=menupos%>';" class="button">     				
				</td>
			</tr>	
			</table>
		</td>	
	<tr>
		<td>
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frmPrize" method="post" >
			<input type="hidden" name="menupos" value="<%=menupos%>">			
			<input type="hidden" name="evt_code" value="<%=evt_code%>">							
			<tr bgcolor="#FFFFFF" height="25">
				<td colspan="9">�˻���� : <b><%=cEvtPrize.FTotalCount%></b>&nbsp;&nbsp;������ : <b><%=page%> / <%=cEvtPrize.FTotalPage%></b></td>
			</tr>		
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">��÷�ڵ�</td>							
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">���</td>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����Ī</td>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">����ǰ��</td>							
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">��÷��</td>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">��÷Ȯ�αⰣ</td>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
			</tr>
			<%IF cEvtPrize.fresultcount >0 THEN%>	
				<%For i = 0 To cEvtPrize.fresultcount - 1 %>
				<tr>
					<td bgcolor="#FFFFFF" align="center"><%=cEvtPrize.fitemlist(i).fevtprize_code%></td>
					<td bgcolor="#FFFFFF" align="center"><%=cEvtPrize.fitemlist(i).fevt_ranking%></td>
					<td bgcolor="#FFFFFF" align="center"><%=cEvtPrize.fitemlist(i).fevt_rankname%></td>
					<td bgcolor="#FFFFFF" align="center"><%=fnGetCommCodeArrDesc_off(arrPrizeType,cEvtPrize.fitemlist(i).fevtprize_type)%></td>
					<td bgcolor="#FFFFFF"  align="left">&nbsp;<%=cEvtPrize.fitemlist(i).fevt_giftname%></td>
					<td bgcolor="#FFFFFF"  align="center"><%=cEvtPrize.fitemlist(i).fevt_winner%><%=cEvtPrize.fitemlist(i).fevt_winner_name%></td>
					<td bgcolor="#FFFFFF" align="left">
						&nbsp;<%if cEvtPrize.fitemlist(i).fevtprize_startdate <> "1900-01-01" then%><%=cEvtPrize.fitemlist(i).fevtprize_startdate%> 
							~ <%if cEvtPrize.fitemlist(i).fevtprize_enddate <> "1900-01-01" then%><%=cEvtPrize.fitemlist(i).fevtprize_enddate%>
							<%end if%>
							<%end if%>
					</td>
					<td bgcolor="#FFFFFF" align="center">
						<%=fnGetCommCodeArrDesc_off(arrPrizeStatus,cEvtPrize.fitemlist(i).fevtprize_status)%>	
					</td>						
				</tr>	
				<%Next%>				
				<tr height="25" bgcolor="FFFFFF">
					<td colspan="15" align="center">
				       	<% if cEvtPrize.HasPreScroll then %>
							<span class="list_link"><a href="?evt_code=<%=evt_code%>&page=<%=cEvtPrize.StartScrollPage-1%>&menupos=<%=menupos%>">[pre]</a></span>
						<% else %>
						[pre]
						<% end if %>
						<% for i = 0 + cEvtPrize.StartScrollPage to cEvtPrize.StartScrollPage + cEvtPrize.FScrollCount - 1 %>
							<% if (i > cEvtPrize.FTotalpage) then Exit for %>
							<% if CStr(i) = CStr(cEvtPrize.FCurrPage) then %>
							<span class="page_link"><font color="red"><b><%= i %></b></font></span>
							<% else %>
							<a href="?evt_code=<%=evt_code%>&page=<%=i%>&menupos=<%=menupos%>" class="list_link"><font color="#000000"><%= i %></font></a>
							<% end if %>
						<% next %>
						<% if cEvtPrize.HasNextScroll then %>
							<span class="list_link"><a href="?evt_code=<%=evt_code%>&page=<%=i%>&menupos=<%=menupos%>">[next]</a></span>
						<% else %>
						[next]
						<% end if %>
					</td>
				</tr>			
			<%else%>	
				<tr>
					<td bgcolor="#FFFFFF" colspan="9" align="center">��÷������ �����ϴ�.</td>
				</tr>
			<%END IF%>	
			</table>	
		</td>			
	</tr>		
</table>	
<%
	set cEvtPrize = nothing
%>