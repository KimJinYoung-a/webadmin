<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : ��ġ����Ŀ ���� ���ι�� ����Ʈ ������
'	History		: 2014.07.24 ���¿� ����
'#############################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhiker_mainbannerCls.asp"-->

<%
Dim i
Dim FResultCount, iCurrentpage, iTotCnt
Dim Searchgubun, SearchUsing, validdate, research
	research= request("research")
	validdate= request("validdate")
	SearchUsing = request("SearchUsing")
	Searchgubun = request("Searchgubun")

	iCurrentpage = NullFillWith(requestCheckVar(Request("IC"),10),1)
if iCurrentpage="" then iCurrentpage=1

if ((research="") and (SearchUsing="")) then 
    SearchUsing = "Y"
    validdate = "on"
end if

Dim opart
set opart = new CAbouthitchhiker
	opart.FCurrPage = iCurrentpage
	opart.FPageSize = 15
	opart.FIsusing = SearchUsing
	opart.Fgubun = Searchgubun
	opart.FValiddate = validdate
	opart.fnGetHitchhikerList
iTotCnt = opart.FTotalCount
%>

<script type="text/javascript">
function conwrite(idx){
	var conwrite = window.open('/admin/hitchhiker/mainbanner/hitchhiker_mainbanner_write.asp?idx='+idx,'hitchhiker_mainbanner_write','width=800,height=768,scrollbars=yes,resizable=yes');
	conwrite.focus();
}
function searchFrm(p){
	frm.iC.value = p;
	frm.submit();
}
</script>
<!-- #include virtual="/admin/hitchhiker/inc_HichHead.asp"-->
<img src="/images/icon_arrow_link.gif"> <b>���ι��</b>
<% '�˻�---------------------------------------------------------------------------------------------------------- %>
<form name="frm" action="index.asp" method="get">
<input type="hidden" name="iC" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=admincolor("tablebg")%>">
	<tr align="center" bgcolor="#FFFFFF">
		<td lowsapn="2" width="100" bgcolor="<%=admincolor("gray")%>"> <b>�˻�����</b> </td>
		<td align="left">
			<select name="SearchGubun">
				<option value ="" style="color:blue">�� ��</option>
				<option value="1" <% If "1" = cstr(SearchGubun) Then%> selected <%End if%>>���λ�ܷѸ����_��ũ</option>
				<option value="2" <% If "2" = cstr(SearchGubun) Then%> selected <%End if%>>���λ�ܷѸ����_���̾��˾�</option>
				<option value="3" <% If "3" = cstr(SearchGubun) Then%> selected <%End if%>>���λ�ܷѸ����_OnlyView</option>
				<option value="4" <% If "4" = cstr(SearchGubun) Then%> selected <%End if%>>���λ�ܷѸ����_����&�߰�</option>
			</select>&nbsp;&nbsp;
			
			<b> �� �� : </b>
			<select name="SearchUsing">
				<option value ="" style="color:blue">�� ü</option>
				<option value="Y" <% If "Y" = cstr(SearchUsing) Then%> selected <%End if%>>Y</option>
				<option value="N" <% If "N" = cstr(SearchUsing) Then%> selected <%End if%>>N</option>
			</select>&nbsp;&nbsp;&nbsp;
			
			<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������
		</td>
		<td lowsapn="2" width=100 bgcolor="<%=admincolor("gray")%>">
			<input type="button" class="button" value="�˻�" onclick="searchFrm('');">
		</td>
	</tr>
</table>
</form>
<% '�˻� ��------------------------------------------------------------------------------------------------------- %>
<br>
<% '�ű��Է¹�ư-------------------------------------------------------------------------------------------------- %>
<table width="100%" align="center">
	<tr>
		<td align="right"><input type="button" name="newBT" class="button" value="�ű��Է�" onclick="conwrite('');"></td>
	</tr>
</table>
<% '�ű��Է¹�ư ��----------------------------------------------------------------------------------------------- %>

<% '����Ʈ-------------------------------------------------------------------------------------------------------- %>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="7" > <% '���պ�(colspan)7�� %>
			�˻���� : <b><%= iTotCnt %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
		<td width="5%"><b>��ȣ</b></td>
		<td width="5%"><b>����</b></td>
		<td width="25%"><b>�̹���</b></td>
		<td width="5%"><b>��뿩��</b></td>
		<td width="5%"><b>�켱����</b></td>
		<td width="15%"><b>������/������</b></td>
		<td width="15%"><b>�����</b></td>
	</tr>
	
	<% if opart.FResultCount > 0 then %>
	
		<% for i = 0 to opart.FResultCount - 1 %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF'; height="30"> 
			<td style="cursor:hand"  onclick="conwrite('<%= opart.FItemList(i).Fidx %>');"><%= opart.FItemList(i).Fidx %></td> <% '��ȣ(idx) %>
			
			<td><%= getHitchhikerGubun(opart.FItemList(i).FReqgubun) %></td> <% '����(PC,MOBILE,MOVIE %>
			
			<td><img src="<%= opart.FItemList(i).FReqcon_viewthumbimg %>" width="300" height="100"></td> <% '����� %>
	
			<td><%= opart.FItemList(i).FReqIsusing %></td> <% '��뿩�� %>
			
			<td><%= opart.FItemList(i).FReqsortnum %></td> <% '�켱���� %>
	
			<td align="left"> <% '������,������ %>
				<% 
					Response.Write "����: "
					Response.Write replace(left(opart.FItemList(i).FReqSdate,10),"-",".") & " / " & Num2Str(hour(opart.FItemList(i).FReqSdate),2,"0","R") & ":" &Num2Str(minute(opart.FItemList(i).FReqSdate),2,"0","R")
					Response.Write "<br />����: "
					Response.Write replace(left(opart.FItemList(i).FReqEdate,10),"-",".") & " / " & Num2Str(hour(opart.FItemList(i).FReqEdate),2,"0","R") & ":" & Num2Str(minute(opart.FItemList(i).FReqEdate),2,"0","R")
					Response.Write "<br />"
		
					if now() >=  opart.FItemList(i).FReqSdate AND NOW() <= opart.FItemList(i).FReqEdate then
						Response.write " <span style=""color:blue"">(����)</span>"
					elseif now() < opart.FItemList(i).FReqSdate then
						Response.write " <span style=""color:green"">(���¿���)</span>"
					else
						Response.write " <span style=""color:red"">(����)</span>"
					end if
					Response.Write "<br />"
				%>
			</td>
			<td><% Response.Write left(opart.FItemList(i).FreqRegdate,22) %></td> <% '����� %>
		</tr>
		<% next %>
		
		<% '����¡ó��----------------------------------------- %>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="15" align="center">
		       	<% if opart.HasPreScroll then %>
					<span class="list_link"><a href="javascript:searchFrm('<%= opart.StartScrollPage-1 %>')">[pre]</a></span> '&menupos=<%=menupos%>
				<% else %>
				[pre]
				<% end if %>
					<% for i = 0 + opart.StartScrollPage to opart.StartScrollPage + opart.FScrollCount - 1 %>
						<% if (i > opart.FTotalpage) then Exit for %>
						<% if CStr(i) = CStr(iCurrentpage) then %>
						<span class="page_link"><font color="red"><b><%= i %></b></font></span>
						<% else %>
						<a href="javascript:searchFrm('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
						<% end if %>
					<% next %>
				<% if opart.HasNextScroll then %>
					<span class="list_link"><a href="javascript:searchFrm('<%= i %>')">[next]</a></span>
				<% else %>
				[next]
				<% end if %>
			</td>
		</tr>
		<% '����¡ó�� ��------------------------------------------ %>
	<% else %>	
		<tr>
			<td colspan=7 align="center">
				�˻������ �����ϴ�.
			</td>
		</tr>
	<% end if %>
</table>
<% '����Ʈ ��----------------------------------------------------------------------------------------------- %>
<%
set opart = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->