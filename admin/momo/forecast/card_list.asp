<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������
' Hieditor : 2010.11.15 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim oforecast,i,page , cardidx , isusing
	menupos = request("menupos")
	page = request("page")
	cardidx = request("cardidxsearch")	
	isusing = request("isusing")			
	if page = "" then page = 1

'// ����Ʈ
set oforecast = new cforecast_list
	oforecast.FPageSize = 20
	oforecast.FCurrPage = page
	oforecast.frectcardidx = cardidx	
	oforecast.frectisusing = isusing			
	oforecast.fcard_list()
%>

<script language="javascript">

	//�űԵ�� & ����
	function reg(cardidx){
		var reg = window.open('/admin/momo/forecast/card_reg.asp?cardidx='+cardidx,'reg','width=600,height=400,scrollbars=yes,resizable=yes');
		reg.focus();
	}
	
	//��ǥ���
	function card_reg(cardidx){
		var card_reg = window.open('/admin/momo/forecast/card_detail.asp?cardidx='+cardidx,'card_reg','width=600,height=768,scrollbars=yes,resizable=yes');
		card_reg.focus();
	}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get action="">
<input type="hidden" name="cardidx">	
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>		
	<td align="left">
		cardidx : <input type="text" name="cardidxsearch" value="<%=cardidx%>" size=10>		
		&nbsp; ��뿩�� : 
		<select name="isusing">
			<option value="" <% if isusing = "" then response.write " selected" %>>��뿩��</option>
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>		
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">			
		</td>
		<td align="right">		
			<input type="button" onclick="reg('');" value="�űԵ��" class="button">					
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oforecast.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oforecast.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= oforecast.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">��ȣ</td>
	<td align="center">����</td>	
	<td align="center">�Ⱓ</td>
	<td align="center">�����</td>
	<td align="center">��뿩��</td>
	<td align="center">���</td>
</tr>
<% for i=0 to oforecast.FresultCount-1 %>			

<% if oforecast.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>
	<td align="center">
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	</td>
	<td align="center">
		<%= oforecast.FItemList(i).fcardidx %><input type="hidden" name="cardidx" value="<%= oforecast.FItemList(i).fcardidx %>">
	</td>
	<td align="center">
		<%= statsgubun(oforecast.FItemList(i).fstats) %>
	</td>	
	<td align="center">
		<%= formatdate(oforecast.FItemList(i).fstartdate,"0000.00.00") %> ~ <%=formatdate(oforecast.FItemList(i).fenddate,"0000.00.00")%>
	</td>			
	<td align="center">
		<%= formatdate(oforecast.FItemList(i).fregdate,"0000.00.00") %>
	</td>		
	<td align="center">
		<%= oforecast.FItemList(i).fisusing %>
	</td>			
	<td align="center">
		<input type="button" onclick="reg(<%= oforecast.FItemList(i).fcardidx %>);" class="button" value="����">
		<input type="button" onclick="card_reg(<%= oforecast.FItemList(i).fcardidx %>);" class="button" value="ī����(<%= oforecast.FItemList(i).fcardcount %>)">
	</td>			
</tr>   

<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if oforecast.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= oforecast.StartScrollPage-1 %>&isusing=<%=isusing%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oforecast.StartScrollPage to oforecast.StartScrollPage + oforecast.FScrollCount - 1 %>
			<% if (i > oforecast.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oforecast.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>&isusing=<%=isusing%>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oforecast.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
	set oforecast = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->