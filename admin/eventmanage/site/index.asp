<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/site/index.asp
' Description :  �̺�Ʈ Static �̹��� ����
' History : 2007.03.27 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventSiteCls.asp"-->
<%
	Call fnSetEventCommonCode '�����ڵ� ���ø����̼� ������ ����
	
	'��������
	Dim cEvtList
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
	Dim slocation, stype, limitCnt
	
	'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
	slocation = Request("sitelocation")
	
	iCurrpage = Request("iC")	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF	  
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	
	
	'������ ��������
	set cEvtList = new ClsEvtSite	
		cEvtList.FCPage = iCurrpage	'����������
		cEvtList.FPSize = iPageSize '���������� ���̴� ���ڵ尹�� 
		cEvtList.FSLocation = slocation '�˻� : ��ġ
		
 		arrList = cEvtList.fnGetList	'�����͸�� ��������
 		iTotCnt = cEvtList.FTotCnt	'��ü ������  ��
 	set cEvtList = nothing
 	
	iTotalPage 	=  Int(iTotCnt/iPageSize)	'��ü ������ ��
	IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1		
	IF isArray(arrList) THEN stype = arrList(2,0) 	
%>
<script language="javascript">
<!--
	function jsGoPage(iP){
		document.frmEvt.iC.value = iP;
		document.frmEvt.action = "index.asp";
		document.frmEvt.submit();	
	}
	
	
	function AssignTest(slocation,stype){	
	 	var popwin = window.open('','refreshFrm_Test','');
		popwin.focus();
		 frmEvt.target = "refreshFrm_Test";
		 frmEvt.action = "<%=staticImgUrl%>/flash/link/make_event_test_JS.asp?sl=" + slocation+"&st="+stype;
		 frmEvt.submit();			 
	}
	
	function AssignReal(slocation,stype){	  
		 var popwin = window.open('','refreshFrm_Main','');
		 popwin.focus();
		 frmEvt.target = "refreshFrm_Main";
		 frmEvt.action = "<%=staticImgUrl%>/flash/link/make_event_JS.asp?sl=" + slocation+"&st="+stype;
		 frmEvt.submit();
	}
	
	function jsChangeFrm(){		
	 var sl;
	 sl= document.frmSearch.sitelocation.options[document.frmSearch.sitelocation.selectedIndex].value ;	 	
	 self.location.href = "index.asp?sitelocation=" + sl;
	 
	}
//-->
</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frmSearch" method="post">
	<input type="hidden" name="iC">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td  >
        	��ġ : <%sbGetOptEventCodeValue "sitelocation",slocation,True,"onChange='javascript:jsChangeFrm();'"%>
        	&nbsp;
        	<%IF slocation <> "" THEN %>
        	<!--<a href="javascript:AssignTest(<%=slocation%>,'<%=stype%>');"><img src="/images/icon_search.jpg" border="0" align="absmiddle">�̸�����</a>
        	/ --><a href="javascript:AssignReal(<%=slocation%>,'<%=stype%>');"><img src="/images/refreshcpage.gif" align="absmiddle" border="0"> ��������</a>        	
            <%END IF%>
        </td>
        <td  align="left" valign="bottom">        	
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>		
	<tr valign="top" style="padding : 0 0 10 0">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td colspan="2">
        	+ ��ġ�˻��� �ؼ� ���밡���� �̹��� Ȯ�� �Ŀ��� �̸�����/���������� �����մϴ�.<br>
        	+ ����κ��� ���밡���� �̹������Դϴ�.        	
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<form name="frmEvt" method="post">
<input type="hidden" name="iC">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="sitelocation" value="<%=slocation%>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="35">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">       	
       	<a href="evtsite_regist.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0"></a>
       	<% If stype = "flash" Then %>* �÷��� ��� ���� �ֱٿ� �ø����� 1���Դϴ�.<% End If %>
    	</td>
    	<td align="right">
       
       <!--	<input type="button" value="���" onclick=" ">  -->
       <!--	����: <select name="selSort">
       	<option value="1">�̺�Ʈ�ڵ峻������</option>
       	
       	</select>-->
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>idx</td>
    	<td>��ġ</td>
    	<td>����</td>
    	<td>�̹���</td>
    	<td>������</td>
      	<td>link����</td>
      	<td>����</td>
      	<td>��뿩��</td>      	
      	<td>�����</td>      	
    </tr>    
    <%IF isArray(arrList) THEN 
    	For intLoop = 0 To UBound(arrList,2)
    	 IF arrList(2,intLoop) = "flash" THEN
    	   limitCnt = 3
    	 ELSE
    	   limitCnt = 1
    	END IF    
    %>
    <tr align="center"  <%IF slocation<> "" and Cint(intLoop) < Cint(limitCnt) THEN%>bgcolor="#FFFFF4"<%ELSE%>bgcolor="#FFFFFF"<%END IF%>>        	
    	<td><%=arrList(0,intLoop)%></td>
    	<td><a href="evtsite_regist.asp?menupos=<%=menupos%>&idx=<%=arrList(0,intLoop)%>"><%=fnGetEventCodeDesc("sitelocation",arrList(1,intLoop))%></a></td>
    	<td><%=arrList(2,intLoop)%></td>
    	<td><a href="evtsite_regist.asp?menupos=<%=menupos%>&idx=<%=arrList(0,intLoop)%>"><img src="<%=arrList(3,intLoop)%>" width="100" border="0"></a></td>
    	<td><%=arrList(6,intLoop)%> X <%=arrList(7,intLoop)%></td>
    	<td><%=arrList(4,intLoop)%></td>
    	<td><%=arrList(8,intLoop)%></td>
    	<td><%=arrList(10,intLoop)%></td>
    	<td><%=arrList(9,intLoop)%></td>    	
    </tr>   
   <%	Next
   	ELSE
   %>
   	<tr  align="center" bgcolor="#FFFFFF">
   		<td colspan="9">��ϵ� ������ �����ϴ�.</td>
   	</tr>	
   <%END IF%>
</table>
<!-- ����¡ó�� -->
<%		
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	

If (iCurrpage mod iPerCnt) = 0 Then																
	iEndPage = iCurrpage
Else								
	iEndPage = iStartPage + (iPerCnt-1)
End If	
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
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
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->