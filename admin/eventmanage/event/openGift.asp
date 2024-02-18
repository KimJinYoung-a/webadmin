<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/openGift.asp
' Description :  ��ü�����̺�Ʈ ���� 369��.
' History : 2010.04 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/openGiftCls.asp"-->

<%


dim evtCode, frontOpen
dim i, page
page = requestCheckVar(request("page"),10)
evtCode = requestCheckVar(request("evtCode"),10)
frontOpen = requestCheckVar(request("frontOpen"),10)

if page="" then page=1

Dim oOpenGift
set oOpenGift=new CopenGift
oOpenGift.FCurrPage = page
oOpenGift.FPageSize = 30
oOpenGift.getOpenGiftList

Dim urlPara : urlPara = Server.UrlEnCode("&evtCode="&evtCode&"&page="&page&"&frontOpen="&frontOpen)
%>

<script language="javascript">
function jsLastEvent(){
	  var winLast,eKind;
	  eKind = 1;
	  var pTarget = '<%= Server.URLEncode("openGift_Reg.asp?menupos=1184") %>';
	  winLast = window.open('pop_event_lastlist.asp?menupos=<%=menupos%>&eventkind='+eKind+'&pTarget='+pTarget,'pLast','width=550,height=600, scrollbars=yes')
	  winLast.focus();
	}
	
function changeGiftScope(eCode,ogiftType){
    var frm = document.frmSm;
    var confirmStr ='����ǰ ���� ����� ��ü�������� �����Ͻðڽ��ϱ�?';
    if (ogiftType==9){
        confirmStr ='����ǰ ���� ����� ���̾����Ÿ������ �����Ͻðڽ��ϱ�?';
    }
    
    if (confirm(confirmStr)){
        frm.eCode.value=eCode;
        frm.opengiftType.value=ogiftType;
        frm.submit();
    }
}
</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmEvt" method="get"  >
	<input type="hidden" name="menupos" value="<%=menupos%>">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">  
			�̺�Ʈ �ڵ� : <input type="text" name="evtCode" value="<%=evtCode%>" maxlength="10" size="9">
			
			����Ʈ���� : 
			<select name="frontOpen">
			<option value="">��ü
			<option value="Y" <%= CHKIIF(frontOpen="Y","selected","") %> >����
			</select>
		</td>
        </td>
    		<td  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:frmEvt.submit();">
		</td>
	</tr>	
	</form>
</table>
<!-- ǥ ��ܹ� ��-->
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >	
    <tr height="40" valign="bottom">       
        <td align="left">
        	<input type="button" value="�̺�Ʈ ����" onclick="jsLastEvent();" class="button">
	    </td>
	    <td align="right">
	       <!-- input type="button" value="������" onclick="jsSchedule();"  class="button" -->
	       <!-- <% if C_ADMIN_AUTH then %><input type="button" value="�ڵ����" onclick="jsCodeManage();"  class="button"><%END IF%> -->
        </td>        
	</tr>	
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="5" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="12">�˻���� : <b><%=oOpenGift.FTotalCount%></b>&nbsp;&nbsp;������ : <b><%=page%> / <%=oOpenGift.FtotalPage%></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td nowrap>�̺�Ʈ�ڵ�</td>
    	<td nowrap>����</td>
    	<td nowrap>����</td>
    	<td nowrap>�̺�Ʈ��</td>
    	<td nowrap>�̺�Ʈ�Ⱓ</td>
    	<td nowrap>�̺�Ʈ����</td>
    	<td nowrap>����Ʈ����</td>
    	<td nowrap>����ǰ</td>
    	<td nowrap>�����</td>
    	<td nowrap>���</td>
    </tr>
    <% for i= 0 to oOpenGift.FREsultCount -1 %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%= oOpenGift.FItemList(i).FEvent_Code %></td>
        <td><%= oOpenGift.FItemList(i).getOpengiftTypeName %></td>
        <td><%= oOpenGift.FItemList(i).getOpengiftScopeName %></td>
        <td align="left"><a href="openGift_Reg.asp?eC=<%= oOpenGift.FItemList(i).FEvent_Code %>&menupos=<%=menupos%>"><%= oOpenGift.FItemList(i).FEvent_Name %></a></td>
        <td><%= oOpenGift.FItemList(i).Fevt_startdate%>~<%= oOpenGift.FItemList(i).Fevt_enddate %></td>
        <td><%= oOpenGift.FItemList(i).getEventStateName %></td>
        <td><%= oOpenGift.FItemList(i).FfrontOpen %></td>
        <td><a href="/admin/shopmaster/gift/giftlist.asp?eC=<%= oOpenGift.FItemList(i).FEvent_Code %>&menupos=1045&fcSc=1"><%= oOpenGift.FItemList(i).FGiftCNT %></a>
        <% if oOpenGift.FItemList(i).FGiftCNT<>oOpenGift.FItemList(i).FALLGiftCNT then %>
        <strong><font color="red">(<%= oOpenGift.FItemList(i).FALLGiftCNT %>)</font></strong>
        <% end if %>
        </td>
        <td><%= oOpenGift.FItemList(i).Freguser %></td>
        <td>
        <% if oOpenGift.FItemList(i).FGiftCNT<>oOpenGift.FItemList(i).FALLGiftCNT then %>
        <img src="/images/icon_arrow_link.gif" onClick="changeGiftScope('<%= oOpenGift.FItemList(i).FEvent_Code %>,'<%=oOpenGift.FItemList(i).FopengiftType%>'')" style="cursor:pointer">
        <% end if %>
        </td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
		<td colspan="12" height="30" align="center">
		<% if oOpenGift.HasPreScroll then %>
			<a href="?page=<%= oOpenGift.StarScrollPage-1 %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oOpenGift.StarScrollPage to oOpenGift.FScrollCount + oOpenGift.StarScrollPage - 1 %>
			<% if i>oOpenGift.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oOpenGift.HasNextScroll then %>
			<a href="?page=<%= i %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>

<form name="frmSm" method="post" action="openGift_Process.asp" >
<input type="hidden" name="imod" value="chgScope"> 
<input type="hidden" name="eCode" value=""> 
<input type="hidden" name="opengiftType" value=""> 
<input type="hidden" name="opengiftScope" value=""> 
</form>
<%
set oOpenGift=Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->