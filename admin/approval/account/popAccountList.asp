<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : ��Ź������ �ش��ϴ� �������� ����  ����Ʈ
' History : 2011.03.09 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/commCls.asp"-->
<!-- #include virtual="/lib/classes/approval/accountCls.asp"-->
<%
Dim clsComm, clsAccount, arrList, intLoop 
Dim iaccountkind, iedmsidx, saccountname, iparentkey  
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage

	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1
	
	iedmsidx	 =	  requestCheckvar(Request("ieidx"),10)
	iaccountkind =  requestCheckvar(Request("selAK"),10)
 	saccountname =  requestCheckvar(Request("sAN"),30)
 	
Set clsAccount = new CAccount
	clsAccount.FedmsIdx		= iedmsidx
	clsAccount.Faccountkind = iaccountkind 
	clsAccount.Faccountname = saccountname 
	clsAccount.FCurrPage 	= iCurrPage
	clsAccount.FPageSize 	= iPageSize
	arrList = clsAccount.fnGetEdmsAccountList 	
	iTotCnt = clsAccount.FTotCnt
Set clsAccount = nothing
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>  
<script language="javascript">
<!--
	function jsSelectEApp(iaccountidx){
	opener.location.href= "/admin/approval/eapp/regeapp.asp?iAidx="+iaccountidx; 
	self.close();
	}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF"> 
<tr>
	<td><strong>����������</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action=""> 
			<input type="hidden" name="iCP" value="">
			<input type="hidden" name="ieidx" value="<%=iedmsidx%>">    
			<input type="hidden" name="iAidx" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
				<td align="left">
					��������:
					<select name="selAK">
					<option value="0">--��ü--</option> 
					<%
					set clsComm = new CcommCode
					clsComm.Fparentkey = 1
					clsComm.Fcomm_cd = iaccountkind
					clsComm.sbOptCommCD
					Set clsComm = nothing
					%>
					</select>  
					&nbsp;&nbsp;
					�������񳻿�:
					 <input type="text" name="sAN" size="20" maxlenght="30" value="<%=saccountname%>">
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
				<td>IDX</td>
				<td>�������񳻿�</td> 
				<td>��������</td>
				<td>������</td>  
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">	
				<td><a href="javascript:jsSelectEApp('<%=arrList(0,intLoop)%>');"><%=arrList(0,intLoop)%></td>
				<td><a href="javascript:jsSelectEApp('<%=arrList(0,intLoop)%>');"><%=arrList(3,intLoop)%></td>			
				<td><a href="javascript:jsSelectEApp('<%=arrList(0,intLoop)%>');"><%=arrList(5,intLoop)%></a></td>	
				<td><a href="javascript:jsSelectEApp('<%=arrList(0,intLoop)%>');"><%=arrList(7,intLoop)%></td>	 
			</tr>
		<%	Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="4">��ϵ� ������ �����ϴ�.</td>	
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
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" >
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
 



	