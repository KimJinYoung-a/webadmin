<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp"-->

<%

dim oldmisend,  inputyn, itemid, itemoption, vSiteName, designer
dim lackItemOnly
itemid  = RequestCheckVar(request("itemid"),10)
itemoption  = RequestCheckVar(request("itemoption"),10)
inputyn = request("inputyn")
vSiteName		= requestCheckVar(request("sitename"),10)
designer		= requestCheckVar(request("designer"),32)
lackItemOnly	= requestCheckVar(request("lackItemOnly"),32)

if inputyn="" then inputyn="N"

set oldmisend = New COldMiSend
oldmisend.FPageSize = 500
oldmisend.FRectDelayDate = 0
'oldmisend.FRectNotInCludeUpcheCheck = notincludeupchecheck
oldmisend.FRectInCludeAlreadyInputed = inputyn
oldmisend.FRectSiteName = vSiteName

oldmisend.FRectMakerid = designer
oldmisend.FRectItemID = itemid
oldmisend.FRectItemOption = itemoption

oldmisend.FRectLackItemOnly = lackItemOnly
oldmisend.GetOldMisendListMasterCS

dim i, tmp

%>
<script language='javascript'>

function misendmaster(v){
	var popwin = window.open("/admin/ordermaster/misendmaster_main.asp?orderserial=" + v,"misendmaster","width=1200 height=700 scrollbars=yes resizable=yes");
	popwin.focus();
}

function cOrderFin(detailidx){
    if (confirm('��� ó�� Ȯ�� �Ͻðڽ��ϱ�?')){
        var popwin = window.open("/admin/ordermaster/misendmaster_main_process.asp?detailidx=" + detailidx + "&mode=cancelFin","misendmaster_process","width=100 height=100 scrollbars=yes resizable=yes");
	    popwin.focus();
    }
}
</script>


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" >
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� : <% drawSelectBoxDesigner "designer", designer %>
			&nbsp;
			��ǰ�ڵ� :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="8" maxlength="10">
            &nbsp;
			��ǰ�ڵ� :
			<input type="text" class="text" name="itemoption" value="<%= itemoption %>" size="8" maxlength="10">
            &nbsp;
			Site :
			<select name="sitename" class="select">
				<option value="">-��ü-</option>
				<option value="10x10" <%=CHKIIF(vSiteName="10x10","selected","")%>>�ٹ�����</option>
				<option value="NOTTEN" <%=CHKIIF(vSiteName="NOTTEN","selected","")%>>���޻���ü</option>
				<option value="interpark" <%=CHKIIF(vSiteName="interpark","selected","")%>>������ũ</option>
				<option value="lotteCom" <%=CHKIIF(vSiteName="lotteCom","selected","")%>>�Ե�����</option>
				<option value="lotteimall" <%=CHKIIF(vSiteName="lotteimall","selected","")%>>�Ե�iMall</option>
				<option value="wizwid" <%=CHKIIF(vSiteName="wizwid","selected","")%>>��������</option>
				<option value="wconcept" <%=CHKIIF(vSiteName="wconcept","selected","")%>>����������</option>
				<option value="bandinlunis" <%=CHKIIF(vSiteName="bandinlunis","selected","")%>>�ݵ�ط��̽�</option>
			</select>
		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="radio" name="inputyn" value="Y" <% if (inputyn = "Y") then response.write "checked" end if %>> ��ü���
			<input type="radio" name="inputyn" value="N" <% if (inputyn = "N") then response.write "checked" end if %>> ��ó�����
			<!--
			<input type="radio" name="inputyn" value="1" <% if (inputyn = "1") then response.write "checked" end if %>> SMS�Ϸ�
			<input type="radio" name="inputyn" value="2" <% if (inputyn = "2") then response.write "checked" end if %>> �ȳ�Mail�Ϸ�
			<input type="radio" name="inputyn" value="3" <% if (inputyn = "3") then response.write "checked" end if %>> ��ȭ�Ϸ�
			-->
			<input type="radio" name="inputyn" value="4" <% if (inputyn = "4") then response.write "checked" end if %>> ���ȳ�
			<input type="radio" name="inputyn" value="6" <% if (inputyn = "6") then response.write "checked" end if %>> CSó���Ϸ�
			&nbsp;
			<input type="radio" name="inputyn" value="C" <% if (inputyn = "C") then response.write "checked" end if %>> ����ֹ� (������ = �Ϻ���Ұǵ� ����.)
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<input type="checkbox" name="lackItemOnly" value="Y" <%= CHKIIF(lackItemOnly="Y", "checked", "")%>> �̹��� ��ǰ��
		</td>
	</tr>
	</form>
</table>

<p>

ó������ ��ó�� / ���ȳ� / CSó���Ϸ�

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmview" method="get">
	<input type="hidden" name="iid" value="">
	</form>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			�˻���� : <b><%= oldmisend.FResultCount %></b> / �ֹ��Ǽ� : <b><%= oldmisend.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="70">�ֹ���ȣ</td>
        <td width="70">Site</td>
	    <td width="60">�ֹ���</td>
	    <td width="60">������</td>
		<td width="50">��ǰ�ڵ�</td>
		<td width="100">�귣��</td>
		<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
		<td width="40">�ֹ�<br>����</td>
		<td width="40">����<br>����</td>
		<td width="40">�ҿ�<br>�ϼ�</td>
		<td width="40">���<br><%= CHKIIF(inputyn = "C","Ȯ��","����") %></td>
	    <td width="60">�������</td>
		<td width="80">��������</td>
	    <td width="70">�������</td>

	    <td width="80">ó������</td>
	    <td width="40">��<br>����</td>
	</tr>
	<% if oldmisend.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF">
	  	<td colspan="16" align="center">�˻������ �����ϴ�.</td>
	</tr>
	<% else %>

	<% for i=0 to oldmisend.FResultCount -1 %>
	<tr align="center" bgcolor="#FFFFFF">
	    <td align="center">
	    <%
	    if (tmp <> oldmisend.FItemList(i).FOrderSerial) then
	      tmp = oldmisend.FItemList(i).FOrderSerial
	    %>
			<a href="javascript:misendmaster('<%= oldmisend.FItemList(i).FOrderSerial %>');"><%= oldmisend.FItemList(i).FOrderSerial %></a>
	    <% end if %>
	    </td>
        <td><%= oldmisend.FItemList(i).FSiteName %></td>
		<td><%= oldmisend.FItemList(i).FBuyName %></td>
    	<td><%= oldmisend.FItemList(i).FReqName %></td>
	    <td><%= oldmisend.FItemList(i).FItemId %></td>
		<td><%= oldmisend.FItemList(i).Fmakerid %></td>
		<td align="left">
			<%= oldmisend.FItemList(i).FItemname %>
			<% if oldmisend.FItemList(i).FItemOptionName<>"" then %>
			<font color="blue">[<%= oldmisend.FItemList(i).FItemOptionName %>]</font>
			<% end if %>
		</td>
		<td><%= oldmisend.FItemList(i).FItemNo %></td>
		<td><b><font color="red"><%= oldmisend.FItemList(i).FItemLackNo %></font></b></td>
		<td>
		<%
			'response.write oldmisend.FItemList(i).getBeasongDPlusDateStr
			response.write oldmisend.FItemList(i).getNewBeasongDPlusDateStr
		%>
		</td>
		<td>
		    <% IF (inputyn="C") then %>
		        <img src="/images/icon_arrow_link.gif" onClick="cOrderFin('<%= oldmisend.FItemList(i).FDetailIdx %>');" style="cursor:pointer">
		    <% else %>
    		    <% if (oldmisend.FItemList(i).FDetailCancelYn="Y") or (oldmisend.FItemList(i).FCancelYn="Y") then %>
    		    <strong><font color="red">���</font></strong>
    		    <% end if %>
		    <% end if %>
		</td>
	    <td>
	        <font color="<%= oldmisend.FItemList(i).getUpcheDeliverStateColor %>"><%= oldmisend.FItemList(i).getUpcheDeliverStateName %></font>
		</td>
		<td><%= oldmisend.FItemList(i).getMiSendCodeName %></td>
	    <td><%= oldmisend.FItemList(i).getIpgoMayDay %></td>
	    <td><%= oldmisend.FItemList(i).GetStateString %></td>

	    <td>
			<a href="javascript:misendmaster('<%= oldmisend.FItemList(i).FOrderSerial %>');"><img src="/images/icon_search.jpg" border="0"></a>
		</td>
	</tr>
  <% next %>
  <% end if %>
</table>


<%
set oldmisend = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
