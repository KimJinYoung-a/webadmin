<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : ��ǰ�� ��� �귣�� ���� ����Ʈ
'	History		: 2017.01.20 ���¿� ����
'				  2022.07.12 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'#############################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/board/brand_noticeCls.asp"-->

<%
Dim i, FResultCount, iCurrentpage, iTotCnt, Searchgubun, SearchUsing, validdate, research, brandidtext, opart
	research= requestCheckvar(request("research"),10)
	validdate= requestCheckvar(request("validdate"),10)
	SearchUsing = requestCheckvar(request("SearchUsing"),10)
	Searchgubun = requestCheckvar(request("Searchgubun"),10)
	brandidtext = request("brandidtext")
	iCurrentpage = NullFillWith(requestCheckVar(Request("IC"),10),1)

if iCurrentpage="" then iCurrentpage=1
if ((research="") and (SearchUsing="")) then 
    SearchUsing = "Y"
    validdate = "on"
end if

set opart = new CBrandNotice
	opart.FCurrPage = iCurrentpage
	opart.FPageSize = 15
	opart.FIsusing = SearchUsing
	opart.Fgubun = Searchgubun
	opart.Fbrandidtext = brandidtext
	opart.FValiddate = validdate
	opart.fnGetBrandNoticeList
iTotCnt = opart.FTotalCount

%>

<script type="text/javascript">

function conwrite(idx){
	var conwrite = window.open('/admin/board/brand_notice/brand_notice_write.asp?idx='+idx,'brand_notice_write','width=1400,height=800,scrollbars=yes,resizable=yes');
	conwrite.focus();
}
function searchFrm(p){
	frm.iC.value = p;
	frm.submit();
}

</script>
<% '�˻�---------------------------------------------------------------------------------------------------------- %>
<form name="frm" action="index.asp" method="get" style="margin:0px;">
<input type="hidden" name="iC" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ���� <select name="SearchGubun">
			<option value ="" style="color:blue">�� ��</option>
			<option value="1" <% If "1" = cstr(SearchGubun) Then%> selected <%End if%>>�Ϲݰ���</option>
			<option value="2" <% If "2" = cstr(SearchGubun) Then%> selected <%End if%>>��۰���</option>
			<option value="3" <% If "3" = cstr(SearchGubun) Then%> selected <%End if%>>��Ÿ����</option>
		</select>
		&nbsp;
		* ��뿩�� :
		<select name="SearchUsing">
			<option value ="" style="color:blue">�� ü</option>
			<option value="Y" <% If "Y" = cstr(SearchUsing) Then%> selected <%End if%>>Y</option>
			<option value="N" <% If "N" = cstr(SearchUsing) Then%> selected <%End if%>>N</option>
		</select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="searchFrm('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������
		&nbsp;
		�귣��ID : <input type="text" name="brandidtext" value="<%= brandidtext %>">
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
		<td colspan="15" >
			�˻���� : <b><%= iTotCnt %></b>
		</td>
	</tr>

	<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
		<td width="5%"><b>��ȣ</b></td>
		<td width="5%"><b>�귣�� ID</b></td>
		<td width="5%"><b>��������</b></td>
		<td width="15%"><b>��������</b></td>
		<td width="35%"><b>��������</b></td>
		<td width="10%"><b>������/������</b></td>
		<td width="5%"><b>����</b></td>
		<td width="5%"><b>��뿩��</b></td>
		<td width="7%"><b>����� ID</b></td>
		<td width="5%"><b>�����</b></td>
		<td width="5%"><b>���</b></td>
	</tr>
	
	<% if opart.FResultCount > 0 then %>
	
		<% for i = 0 to opart.FResultCount - 1 %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF'; height="30"> 
			<td style="cursor:hand"  onclick="conwrite('<%= opart.FItemList(i).Fidx %>');"><b><%= opart.FItemList(i).Fidx %></b></td> <% '��ȣ(idx) %>

			<td><a href="http://www.10x10.co.kr/street/street_brand.asp?makerid=<%= opart.FItemList(i).FReqbrandid %>" target="blank"><%= opart.FItemList(i).FReqbrandid %></a></td>
			
			<td><%= getBrandNoticeGubun(opart.FItemList(i).FReqgubun) %></td> <% '����(�Ϲݰ���,��۰���) %>
			
			<td><%= opart.FItemList(i).FReqnotice_title %></td> <% '�������� %>

			<td align="left"><%= nl2br(opart.FItemList(i).FReqnotice_text) %></td> <% '�������� %>

			<td align="center"> <% '������,������ %>
				<% 
					Response.Write "����: "
					Response.Write replace(left(opart.FItemList(i).FReqSdate,10),"-",".") & " / " & Num2Str(hour(opart.FItemList(i).FReqSdate),2,"0","R") & ":" &Num2Str(minute(opart.FItemList(i).FReqSdate),2,"0","R")
					Response.Write "<br />����: "
					Response.Write replace(left(opart.FItemList(i).FReqEdate,10),"-",".") & " / " & Num2Str(hour(opart.FItemList(i).FReqEdate),2,"0","R") & ":" & Num2Str(minute(opart.FItemList(i).FReqEdate),2,"0","R")
'					Response.Write "<br />"
				%>
			</td>
			<td>
				<%
				if now() >=  opart.FItemList(i).FReqSdate AND NOW() <= opart.FItemList(i).FReqEdate or opart.FItemList(i).Freqinfiniteregyn="Y" then
					if opart.FItemList(i).Freqinfiniteregyn = "Y" then
						if opart.FItemList(i).Frank = "1" and opart.FItemList(i).FReqisusing = "Y" then
							Response.write " <span style=""color:blue"">������(���)</span>"
						else
							Response.write " <span style=""color:green"">�����(���)</span>"
						end if
					else
						if opart.FItemList(i).Frank = "1" and opart.FItemList(i).FReqisusing = "Y" then
							Response.write " <span style=""color:blue"">������</span>"
						else
							Response.write " <span style=""color:green"">�����</span>"
						end if
					end if
				elseif now() < opart.FItemList(i).FReqSdate then
					if opart.FItemList(i).Freqinfiniteregyn = "Y" then
						Response.write " <span style=""color:green"">�����(���)</span>"
					else
						Response.write " <span style=""color:green"">�����</span>"
					end if
				else
					Response.write " <span style=""color:red"">����</span>"
				end if

'				Response.Write "<br />"
				%>
			</td> <% '���� %>

			<td><%= opart.FItemList(i).FReqisusing %></td>
			<td><%= opart.FItemList(i).FReqmakerid %></td>
			<td><%= opart.FItemList(i).FReqregdate %></td>
			<td><input type="button" class="button" value="����" onclick="conwrite('<%= opart.FItemList(i).Fidx %>');" /></td>
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
			<td colspan=15 align="center">
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
