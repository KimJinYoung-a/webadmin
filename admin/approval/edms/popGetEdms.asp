<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������  ����Ʈ - ������
' History : 2011.11.15 ������  ����
'	jsSetEdms ��ũ��Ʈ �Լ� opener���� �����ؼ� ����ó��
'########################################################### 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/lib/classes/approval/edmsCls.asp"-->
<%
Dim clsEdms
Dim arrList, intLoop
Dim icateidx1, icateidx2, sEdmsName
Dim iTotCnt,iPageSize, iTotalPage,iCurrentPage
Dim blnUsing
  blnUsing = 1 '--������� ������ ������
	iPageSize = 20
	iCurrentPage = requestCheckvar(Request("page"),10)
	if iCurrentPage="" then iCurrentPage=1
	icateidx1 = requestCheckvar(Request("selC1"),10)
	icateidx2 = requestCheckvar(Request("hidC2"),10) 
 
	if icateidx1 = "" then icateidx1 = 0
	if icateidx2 = "" then icateidx2= 0
	sEdmsName = requestCheckvar(Request("sEN"),60)   

Set clsEdms = new Cedms
	 clsEdms.FCateIdx1	=icateidx1 	
	 clsEdms.FCateIdx2 	=icateidx2 
	 clsEdms.FEdmsName  =sEdmsName 
	 clsEdms.FisUsing 	= blnUsing
	 clsedms.FCurrPage 	= iCurrentPage
	 clsedms.FPageSize 	= iPageSize
	 arrList = clsEdms.fnGetEdmsList 
	 iTotCnt = clsedms.FTotCnt

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script> 
<script language="javascript">
	// ī�װ� ajax =========================================================================================================
	$(document).ready(function(){
	$("#selC1").change(function(){
		var iValue = $("#selC1").val();
		var url="/admin/approval/edms/ajaxCategory.asp";
		 var params = "sMode=CL&ipcidx="+iValue;  
		  	 
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){   
		 		$("#divCL").html(args);   
		 	},
		 	 
		 	error:function(e){ 
		 		alert("�����ͷε��� ������ ������ϴ�. �ý������� �������ּ���");
		 		//alert(e.responseText);
		 	}
		 }); 
	}); 
});

 function jsSearch(){    
		document.frm.hidC2.value = $("#selC2").val(); 
		document.frm.submit();
	}	
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF"> 
<tr>
	<td><strong>���� ����</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="post" action="popGetEdms.asp">   
			<input type="hidden" name="hidC2" value="">
			<tr align="center" bgcolor="#FFFFFF">
				<td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
				<td align="left">
						�� ī�װ� :
					<select name="selC1">
					<option value="0">��ü</option>
					<%clsedms.sbGetOptedmsCategory 1,0,icateidx1 %>
					</select>
					
					�� ī�װ� :
					<span id="divCL">
					<select name="selC2" id="selC2">
					<option value="0">��ü</option>
				<% 	IF icateidx1 > 0 THEN	'��ī�װ� ���� �� ��ī�װ� ���ð����ϰ�
						clsedms.sbGetOptedmsCategory 2,icateidx1,icateidx2 
					END IF
				%>
					</select>
					</span>
					<%Set clsEdms = nothing%>
					</td> 
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>������: <input type="text" name="sEN" value="<%=sEdmsName%>" size="20"></td>
			</tr>				
		</form>
		</table>
	</td>
</tr> 
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">  
		<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
				<td>idx</td>
				<td>�����ڵ�</td>
				<td>��ī�װ�</td>
				<td>��ī�װ�</td>  
				<td>������</td>
			<td>����</td>  
		</tr> 
		<%IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
			%>
		<tr bgcolor="#FFFFFF"  align="center">
			<td><%=arrList(0,intLoop)%></td>
		 	<td><%=arrList(7,intLoop)%></td> 
		 	<td><%=arrList(2,intLoop)%></td> 
		 	<td><%=arrList(4,intLoop)%></td> 
		 	<td><%=arrList(6,intLoop)%></td> 
		 	<td><input type="button" class="button" value="����" onClick="opener.jsSetEdms('<%=arrList(0,intLoop)%>','<%=arrList(6,intLoop)%>');self.close();"> </td>
		</tr>  
	<%	Next
		END IF%>
		</table>	
	</td> 
</tr>  
</table>
<!-- ������ ���� -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10
		
		iStartPage = (Int((iCurrentpage-1)/iPerCnt)*iPerCnt) + 1
		
		If (iCurrentpage mod iPerCnt) = 0 Then
			iEndPage = iCurrentpage
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
					    <tr valign="bottom" height="25">        
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(iCurrentpage) then
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
 