<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��� �������� ����Ʈ
' History : 2011.05.30 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/classes/expenses/OpExpArapCls.asp"-->
<%
Dim clsAccount, arrList, intLoop  	
Dim sarap_nm
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
 
	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1
		 
 	sarap_nm =  requestCheckvar(Request("sAN"),30)
 	
Set clsAccount = new COpExpAccount
	clsAccount.Farap_nm  	= sarap_nm 
	clsAccount.FCurrPage 	= iCurrPage
	clsAccount.FPageSize 	= iPageSize
	arrList = clsAccount.fnGetOpExpAccountList 	
	iTotCnt = clsAccount.FTotCnt
Set clsAccount = nothing
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%> 
 
<script language="javascript">
<!--
// ������ �̵�
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}
	   
//���ε��
function jsNewReg(){
	var winC = window.open("popAccount.asp","popC","width=800, height=800, resizable=yes, scrollbars=yes");
	winC.focus();
}  

//����
function jsDelete(iOEA){
	if(confirm("�����Ͻðڽ��ϱ�?")){
		document.frmDel.hidOEA.value = iOEA;
		document.frmDel.submit();
	}
}

//����
function jsMod(iOEA, iValue,strMsg){
	if(confirm(strMsg)){
		document.frmMod.hidOEA.value = iOEA;
		document.frmMod.hidInOut.value = iValue;
		document.frmMod.submit();
	}else{ 
	window.location.reload();
}
}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a"> 
<form name="frmDel" method="post" action="procAccount.asp">
<input type="hidden" name="hidM" value="D">
<input type="hidden" name="hidOEA" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<form name="frmMod" method="post" action="procAccount.asp">
<input type="hidden" name="hidM" value="U">
<input type="hidden" name="hidOEA" value="">
<input type="hidden" name="hidInOut" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
					��������: 
					 <input type="text" name="sAN" size="20" maxlenght="30" value="<%=sarap_nm%>">
				</td>
				<td  width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr> 
<!-- #include virtual="/lib/db/dbclose.asp" --> 
<tr>
	<td><input type="button" class="button" value="�����׸� �߰�" onClick="jsNewReg();"></td>
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
				<td>����ڵ�</td>  
				<td>�����׸�</td>  
				<td>�����������</td> 
				<td>erpcode</td>  
				<td>���/���� �ݾױ���</td>
				<td>ó��</td>
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 bgcolor="#FFFFFF">	
				<td align="center" > <%=arrList(0,intLoop)%></td>
				<td>[<%=arrList(1,intLoop)%>] <%=arrList(2,intLoop)%></td>
				<td>[<%=arrList(3,intLoop)%>] <%=arrList(4,intLoop)%></td>			
				<td align="center"><%=arrList(5,intLoop)%></td>	 
				<td align="center"><input type="radio" name="rdoInOut<%=arrList(0,intLoop)%>" value="1" <%IF arrList(6,intLoop) THEN %>checked<%END IF%> onClick="jsMod('<%=arrList(0,intLoop)%>',1,'���ݾ����� �����Ͻðڽ��ϱ�?');">���ݾ� <input type="radio" name="rdoInOut<%=arrList(0,intLoop)%>" value="0" <%IF not arrList(6,intLoop) THEN %>checked<%END IF%> onClick="jsMod('<%=arrList(0,intLoop)%>',0,'���ޱݾ����� �����Ͻðڽ��ϱ�?');">���ޱݾ�</td>	 
				<td align="center"><input type="button" class="button" value="����" onClick="jsDelete('<%=arrList(0,intLoop)%>');"> </td>
			</tr>
		<%	Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="6">��ϵ� ������ �����ϴ�.</td>	
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
 



	