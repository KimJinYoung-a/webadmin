<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����ڵ���� ����Ʈ
' History : 2011.03.09 ������  ����
'			2022.07.11 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/commCls.asp"-->
<%
Dim clsComm, arrList, intLoop 
Dim iparentkey 
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
 
	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1
	
	iparentkey = requestCheckvar(Request("selPK"),10)  
	if iparentkey = "" then iparentkey = 0 
Set clsComm = new CcommCode
	clsComm.Fparentkey 	= iparentkey 
	clsComm.FCurrPage 	= iCurrPage
	clsComm.FPageSize 	= iPageSize
	arrList = clsComm.fnGetCommCDList 	
	iTotCnt = clsComm.FTotCnt
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%> 
 
<script type='text/javascript'>
<!--
// ������ �̵�
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}
	  
function jsChangeGroup(){
	document.frm.submit();
}
	  
//���ε��
function jsNewReg(){
	var winC = window.open("popCommCodeConts.asp","popC","width=1200, height=768, resizable=yes, scrollbars=yes");
	winC.focus();
} 
//����
function jsModReg(commCD){
	var winC = window.open("popCommCodeConts.asp?icc="+commCD,"popC","width=1200, height=768, resizable=yes, scrollbars=yes");
	winC.focus();
}

//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a"> 
<tr>
	<td>
		<form name="frm" method="get" action="">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="iCP" value="">
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
					�׷��  :
					<select name="selPK" onChange="jsChangeGroup();">
					<option value="0">--�׷�--</option> 
					<%clsComm.FRectParentKey = iparentkey
					clsComm.sbOptCommCDGroup%>
					</select> 
				</td> 
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
		</table>
		</form>
	</td>
</tr>
<%Set clsComm = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" --> 
<tr>
	<td><input type="button" class="button" value="�űԵ��" onClick="jsNewReg();"></td>
</tr>
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td colspan="15">
					�˻���� : <b><%=iTotCnt%></b> &nbsp;
					������ : <b><%= iCurrPage %> / <%=iTotalPage%></b>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
				<td>�׷��</td>
				<td>IDX</td>
				<td>�߰��ڵ�</td>
				<td>�ڵ��</td> 
				<td>����</td> 	   
				<td>ǥ�ü���</td> 	
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">	
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(5,intLoop)%></td>
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(0,intLoop)%></td>			
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(3,intLoop)%></a></td>	
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%= ReplaceBracket(arrList(1,intLoop)) %></td>	
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(2,intLoop)%></td>  
				<td><a href="javascript:jsModReg(<%=arrList(0,intLoop)%>);"><%=arrList(6,intLoop)%></td>  
			</tr>
		<%	Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="12">��ϵ� ������ �����ϴ�.</td>	
			</tr>
			<%END IF%>
		</table>	
	</td> 
</tr> 	
<!-- ������ ���� -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10
		
		iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1
		
		If (iCurrPage mod iPerCnt) = 0 Then
			iEndPage = iCurrPage
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
					    <tr valign="bottom" height="25">        
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(iCurrPage) then
							%>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
							<%		else %>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
							<%
									end if
								next
							%>
					    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
							<% else %>[next]<% end if %>
					        </td>        
					    </tr>        
					</table>
				</td>
			</tr>
</table>
<!-- ������ �� -->
</body>
</html>
 



	