<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfingersCls.asp"-->
<%
'##############################################
' History: 2008.03.12 modify - 2008 ������ �߰� ��� ����
' Description: ������ �ΰŽ�
'##############################################
 Dim clsDF,clsDFCode
 Dim arrList, intLoop
 Dim iDFSeq, sTitle
 Dim iTotCnt
 Dim iPageSize, iCurrpage ,iDelCnt
 Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
 Dim arrCode, edid, emktid
  
	iDFSeq 		= requestCheckVar(Request("iDFS"),10)	'�ΰŽ�  id
	sTitle 		= requestCheckVar(Request("sT"),10)		'����
	iCurrpage 	= requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
	edid  		= requestCheckVar(Request("selDId"),32)		'��� �����̳�
	emktid 		= requestCheckVar(Request("selMKTId"),32)		'��� MD
 
	IF iCurrpage = "" THEN	iCurrpage = 1
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	
'//����Ʈ ��������	
 set clsDF = new CDesignFingers
 	clsDF.FCPage = iCurrpage	'����������
	clsDF.FPSize = iPageSize '���������� ���̴� ���ڵ尹��
 	clsDF.FDFSeq = iDFSeq
 	clsDF.FTitle = sTitle
 	clsDF.FEDId = edid
 	clsDF.FEMKTId = emktid
 	arrList = clsDF.fnGetDFList
 	iTotCnt = clsDF.FTotCnt	'��ü ������  ��
 set clsDF = nothing
 
 '//�ΰŽ�����(10)�� �ش��ϴ� �ڵ峻�� �迭�� �ֱ�
 set clsDFCode = new CDesignFingersCode
 	arrCode = clsDFCode.fnGetCommCode(10)	
 set clsDFCode =nothing
 iTotalPage 	=  Int(iTotCnt/iPageSize)	'��ü ������ ��
 IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1
 	
%>
<script language="javascript">
<!--
	function jsSearch(){
		document.frmSearch.submit();
	}
	
	function jsGoPage(iP){
		document.frmPage.iC.value = iP;
		document.frmPage.submit();
	}
	
	function jsPopCode(){
		var winCode;
		winCode = window.open('popManageCode.asp','popCode','width=400,height=600');
		winCode.focus();
	}
	
 	function jsSetFile(iDFS){   
 	 var winfile = window.open('','setfile','width=1,height=1');	
 	 	 document.frmFile.iDFS.value = iDFS;
		 document.frmFile.target 	= "setfile";
		 document.frmFile.action 	= "<%=staticUploadUrl%>/chtml/make_designfingers_FlashText.asp";
		 document.frmFile.submit(); 
		
	 winfile.focus();			 
	}
	function onlyNumberInput() 
	{ 
		var code = window.event.keyCode; 
		if ((code > 34 && code < 41) || (code > 47 && code < 58) || (code > 95 && code < 106) || code == 8 || code == 9 || code == 13 || code == 46) { 
			window.event.returnValue = true; 
			return; 
		} 
		window.event.returnValue = false; 
	}
	function workerlist()
	{
		var openWorker = null;
		var worker = frmSearch.selMKTId.value;
		openWorker = window.open('PopWorkerList.asp?worker='+worker+'&team=11&frm=frmSearch','openWorker','width=570,height=570,scrollbars=yes');
		openWorker.focus();
	}
//-->
</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" >
<form name="frmFile" method="post">
<input type="hidden" name="iDFS" value="">
</form>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">			
			<form name="frmSearch" method="get" action="listDF.asp">	
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
				<td align="left">
					<table cellpadding="0" cellspacing="0" border="0" class="a">
					<tr height="25">
						<td>�������ΰ� ID: <input type="text" name="iDFS" value="<%= iDFSeq %>" size="10" maxlength="10"  onKeyDown = "javascript:onlyNumberInput()" style="IME-MODE: disabled" />
						&nbsp;�������ΰ� ����:<input type="text" name="sT" value="<%= sTitle %>" size="32" maxlength="32"></td>
					</tr>
					<tr height="25">
						<td>�������: <%sbGetDesignerid "selDId",edid, "onChange='javascript:document.frmSearch.submit();'"%>&nbsp;&nbsp;
						����ȹ��: <%sbGetwork "selMKTId",emktid,""%></td>
					</tr>
					</table>
				</td>
				<td  width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSearch();">
				</td>
			</tr>
			</form>	
		</table>	
	</td>	
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >	
	    <tr height="40" valign="bottom">       
	        <td align="left">
				<input type="button" class="button" value="���ε��" onClick="location.href='regDF.asp?menupos=<%= menupos %>'">
				<input type="button" class="button" value="��õ����Ʈ" onClick="window.open('recommend_list.asp','','width=800,height=600,scrollbars=yes');">
				<input type="button" class="button" value="��õ�˻���" onClick="window.open('<%=Replace(wwwUrl,"2010","2011")%>/chtml/designfingers/taglist.asp','','width=350,height=130,scrollbars=no');">
			</td>
			<td align="right">	
				<input type="button" class="button" value="Best ����" onClick="location.href='listBest.asp?menupos=<%= menupos %>'">
				<% if C_ADMIN_AUTH then %><input type="button" class="button" value="�ڵ����" onclick="javascript:jsPopCode();"><%END IF%>				
			</td>
		</tr>
			
		</table>
	</td>
</tr>
<tr>
	<td> 
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a"  bgcolor="#CCCCCC">		
		<tr bgcolor="#EFEFEF">
			<td width="40" align="center" nowrap>ID	</td>
			<td width="60" align="center" nowrap>����</td>
			<td width="50" align="center" nowrap>Image</td>
			<td align="center">����</td>
			<td width="60" align="center" nowrap>������</td>
			<td width="60" align="center" nowrap>��÷��ǥ��</td>
			<td width="60" align="center" nowrap>�����</td>
			<td width="30" align="center" nowrap>����</td>
			<td  align="center" nowrap>����</td>			
		</tr>
		<%IF isArray(arrList) THEN%>
		<% For intLoop =0 To UBound(arrList,2) %>	
		<tr bgcolor="#FFFFFF">
			<td align="center"><%=arrList(0,intLoop)%></td>
			<td align="center"><%=fnGetCodeArrDesc(arrCode,arrList(1,intLoop))%></td>
			<td align="center"><%IF arrList(6,intLoop) <> "" THEN%><img src="<%=arrList(6,intLoop)%>"><%END IF%></td>
			<td align="left" ><a href="regDF.asp?iDFS=<%=arrList(0,intLoop)%>&menupos=<%= menupos %>&iC=<%=iCurrpage%>"><%=arrList(2,intLoop)%></a></td>
			<td align="center" ><%=arrList(7,intLoop)%></td>
			<td align="center" ><%=arrList(3,intLoop)%></td>
			<td align="center"><%=FormatDate(arrList(5,intLoop),"0000.00.00")%></td>
			<td align="center" ><%IF arrList(4,intLoop) THEN%>Y<%ELSE%>N<%END IF%></td>
			<td align="center">
			<!--<input type="button" value="�÷������ϻ���" class="button" onClick="javascript:jsSetFile('<%=arrList(0,intLoop)%>');">//-->
			<input type="button" value="��÷" class="button" onClick="location.href='regPrizeDF.asp?iDFS=<%=arrList(0,intLoop)%>';"></td>
		</tr> 
		<% Next%>
		<%ELSE%>
		<tr bgcolor="#FFFFFF">
			<td colspan="8" align="center">��ϵ� ������ �����ϴ�.</td>
		</tr>
		<%END IF%>	
		</table>
	</td>
		
</tr>
<tr>
	<td>
		<!-- ����¡ó�� -->
<%
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1

If (iCurrpage mod iPerCnt) = 0 Then
	iEndPage = iCurrpage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
		<form name="frmPage" method="get" action="listDF.asp">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="iC" value="">
	    <tr valign="bottom" height="25">	       
	        <td valign="bottom" align="center">
	         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
	        <%
				for ix = iStartPage  to iEndPage
					if (ix > iTotalPage) then Exit for
					if Cint(ix) = Cint(iCurrpage) then
			%>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
			<%		else %>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
			<%
					end if
				next
			%>
	    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
			<% else %>[next]<% end if %>
	        </td>	        
	    </tr> 
	    </form>
		</table>
	</td>
</tr>
</table>	
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->