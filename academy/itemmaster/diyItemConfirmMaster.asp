<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���� ��ǰ ��� ��� ��ǰ 
' Hieditor : 2010.10.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/diyshopitem/waitDIYitemCls.asp"-->
<%
dim waititembysort, sorttype, sortkey ,currstate, research ,i,j
	sorttype  = RequestCheckvar(request("sorttype"),10)
	sortkey = RequestCheckvar(request("sortkey"),10)
	currstate = RequestCheckvar(request("currstate"),10)
	research = RequestCheckvar(request("research"),10)
	
	if sorttype="" then sorttype="B"
	
	if research="" then
		currstate="W"
	end if

set waititembysort = new CWaitItemlist
	waititembysort.FRectcurrstate = currstate
	
	'/ī�װ���
	if sorttype="C" then
		waititembysort.getWaitSummaryListByCategory
	
	'/�귣�庰
	elseif sorttype="B" then		
		waititembysort.getWaitSummaryListByBrand
	end if
%>

<script language='javascript'>

function PopItemconfirm(sorttype,sortkey,currstate){
	var popwin=window.open('diyitem_confirm.asp?sorttype=' + sorttype + '&sortkey=' + sortkey +'&currstate='+ currstate,'_blank','width=900,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function PopUpcheBrandInfoEdit(v){
	window.open("/admin/member/popupchebrandinfo.asp?designer=" + v,"PopUpcheBrandInfoEdit","width=640,height=580,scrollbars=yes,resizabled=yes");
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<input type="radio" name="currstate" value="W" <% if currstate="W" then response.write "checked" %>>��ϴ���ǰ��
		<input type="radio" name="currstate" value="WR" <% if currstate="WR" then response.write "checked" %>>��ϴ��+��Ϻ���				
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		���� :
		<select name="sorttype">
			<option value="C" <% if sorttype="C" then response.write "selected" %> >ī�װ���</option>
			<option value="B" <% if sorttype="B" then response.write "selected" %> >�귣�庰</option>
		</select>		
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>
		
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">			
		</td>
		<td align="right">			
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if waititembysort.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= waititembysort.FResultCount %></b>		
	</td>
</tr>
	<% if sorttype="B" then %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>�귣��ID</td>
			<td>�귣���</td>
			<td>��ϴ��</td>
			<td>��Ϻ���</td>
			<td>ī�װ�</td>
			<td>���������</td>
			<td>���</td>
	    </tr>
		<% for i=0 to waititembysort.FresultCount-1 %>
	
	    <tr align="center" bgcolor="#FFFFFF" >
			<td><a href="javascript:PopUpcheBrandInfoEdit('<%= html2db(waititembysort.FItemList(i).FSortKey) %>')"><%= db2html(waititembysort.FItemList(i).FSortKey) %></a></td>
			<td><%= waititembysort.FItemList(i).FSortname %></td>
			<td align=center><%= waititembysort.FItemList(i).FSortCount %></td>
			<td align=center><%= waititembysort.FItemList(i).FRejCount %></td>
			<td align=center><%= waititembysort.FItemList(i).Fcdl_nm %></td>
			<td align=center><%= waititembysort.FItemList(i).Flastregdate %></td>
			<td><a href="javascript:PopItemconfirm('<%= sorttype %>','<%= waititembysort.FItemList(i).FSortKey %>','<%=currstate%>');">�󼼳�������>></a></td>
	    </tr>
		<% next %>
	
	<% else %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>�ڵ�</td>
			<td>ī�װ�</td>
			<td>��ϴ��</td>
			<td>��Ϻ���</td>
			<td>���������</td>
			<td>���</td>
	    </tr>
		<% for i=0 to waititembysort.FResultCount-1 %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
			<td align=center><%= waititembysort.FItemList(i).FSortKey %></td>
			<td align=center><%= waititembysort.FItemList(i).FSortname %></td>
			<td align=center><%= waititembysort.FItemList(i).FSortCount %></td>
			<td align=center><%= waititembysort.FItemList(i).FRejCount %></td>
			<td align=center><%= waititembysort.FItemList(i).Flastregdate %></td>
			<td align=center><a href="javascript:PopItemconfirm('<%= sorttype %>','<%= waititembysort.FItemList(i).FSortKey %>','<%=currstate%>');">�󼼳�������>></a></td>
		</tr>
		<% next %>	    	
	<% end if %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
	set waititembysort = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->