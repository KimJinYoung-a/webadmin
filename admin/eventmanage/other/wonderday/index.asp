<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventOtherCls_wonderday.asp"-->
<%
 Dim clsWonderday
 Dim iTotCnt, arrList,intLoop
 Dim iPageSize, iCurrentpage ,iDelCnt
 Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
 	
 	'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
	iCurrentpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
	IF iCurrentpage = "" THEN		iCurrentpage = 1
	iPageSize = 16		'�� �������� �������� ���� ��, front�� �����ϰ�
	iPerCnt = 10		'�������� ������ ����
	
 set clsWonderday = new CWonderday
 	clsWonderday.FCPage = iCurrentpage	'����������
	clsWonderday.FPSize = iPageSize '���������� ���̴� ���ڵ尹��
	arrList = clsWonderday.fnGetImgList	'�����͸�� ��������
 	iTotCnt = clsWonderday.FTotCnt	'��ü ������  ��
 set clsWonderday = nothing
 
 iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>
<script language="javascript">
<!--
	function jsGoURL(sURL){
		location.href = sURL+"?menupos=<%=menupos%>";
	}
	
	function jsGoPage(iP){
		document.frmpage.iC.value = iP;
		document.frmpage.submit();
	}

//-->
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="5" class="a">
<tr>
	<td><!-- �׼� ���� -->
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
			<tr>
				<td align="left">
					<input type="button" class="button" value="���ε��" onClick="jsGoURL('regConts.asp');">
					&nbsp;
				</td>		
			</tr>
		</table>
		<!-- �׼� �� -->
	</td>
</tr>	
<tr>
	<td>�˻���� : <b><%=iTotCnt%></b>
		&nbsp;
		������ : <b><%=iCurrentpage%> / <%=iTotalPage%></b>
		<!-- ����Ʈ ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">			
			<tr align="center" bgcolor="<%= adminColor("gray") %>">
				<td>ȸ��</td>
		    	<td width="80">ID</td>
		    	<td>����Ʈ�̹���</td>
		    	<td>���ÿ���</td>
		    	<td>������</td>
		      	<td>�����</td>      	
		    </tr>   
			<%IF isArray(arrList) THEN 
				For intLoop =0 To UBound(arrList,2)
				%>
			<tr align="center" bgcolor="#FFFFFF">
				<td><%=arrList(5,intLoop)%></td>
				<td><%=arrList(0,intLoop)%></td>
				<td><a href="regConts.asp?idx=<%=arrList(0,intLoop)%>&iC=<%=iCurrentpage%>&menupos=<%=menupos%>"><img src="<%=arrList(1,intLoop)%>" border="0"></a></td>
				<td><% IF arrList(2,intLoop) THEN%><font color="blue">Y</font><%ELSE%><font color="gray">N</font><%END IF%></td>
				<td><%=arrList(4,intLoop)%></td>
				<td><%=arrList(3,intLoop)%></td>
		    </tr>   
		   	<%
		   		Next
		   	ELSE	
		   	%>
		   	<tr>
		   		<td colspan="4"  bgcolor="#FFFFFF" align="center">��ϵ� ������ �����ϴ�.</td>
			</tr>
			<%
			END IF
			%>
		</table>
	</td>		
</tr>
<tr>
	<tD>
	<!-- ����¡ó�� -->
	<%
	iStartPage = (Int((iCurrentpage-1)/iPerCnt)*iPerCnt) + 1
	
	If (iCurrentpage mod iPerCnt) = 0 Then
		iEndPage = iCurrentpage
	Else
		iEndPage = iStartPage + (iPerCnt-1)
	End If
	%>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frmpage" method="post">	
	    <tr valign="bottom" height="25">        
	        <td valign="bottom" align="center">
	         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
	        <%
				for ix = iStartPage  to iEndPage
					if (ix > iTotalPage) then Exit for
					if Cint(ix) = Cint(iCurrentpage) then
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
