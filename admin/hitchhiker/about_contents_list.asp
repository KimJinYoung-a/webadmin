<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : HITCHHIKER ADMIN
'	History		: 2014.07.09 ���¿� ����
'			 	  2022.07.07 �ѿ�� ����(isms�����������ġ)
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/about_hitchhiker_contentsCls.asp"-->

<%
Dim i
Dim FResultCount, iCurrentpage, iTotCnt
Dim Searchgubun, SearchTitle, SearchUsing, SearchEvtCode
	SearchTitle = request("evt_title")
	SearchUsing = request("SearchUsing")
	Searchgubun = request("Searchgubun")
		
	iCurrentpage = NullFillWith(requestCheckVar(Request("IC"),10),1)
if iCurrentpage="" then iCurrentpage=1
	
Dim opart
set opart = new CAbouthitchhiker
	opart.FCurrPage = iCurrentpage
	opart.FPageSize = 15
	opart.Frectcon_title = SearchTitle
	opart.FrectIsusing = SearchUsing
	opart.Frectgubun = Searchgubun
	opart.fnGetHitchhikerList
iTotCnt = opart.FTotalCount
%>

<script type="text/javascript">
	function conwrite(contentsidx){
		var conwrite = window.open('/admin/hitchhiker/about_contents_write.asp?contentsidx='+contentsidx,'about_contents_write','width=1400,height=768,scrollbars=yes,resizable=yes');
		conwrite.focus();
	}
	function sizewrite(){
		var sizewrite = window.open('/admin/hitchhiker/about_size_write.asp','about_size_write','width=800,height=768,scrollbars=yes,resizable=yes');
		sizewrite.focus();
	}
	function searchFrm(p){
		frm.iC.value = p;
		frm.submit();
	}
</script>

<!--�˻�------------------------------------------------------------------------------------------------->
<form name="frm" action="about_contents_list.asp" method="get">
<input type="hidden" name="iC" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=admincolor("tablebg")%>">
	<tr align="center" bgcolor="#FFFFFF">
		<td lowsapn="2" with="50" bgcolor="<%=admincolor("gray")%>"> <b>�˻�����</b> </td>
		<td align="left">
			<select name="Searchgubun">
				<option value ="" style="color:blue">�� ��</option>
				<option value="1" <% If "1" = cstr(SearchGubun) Then%> selected <%End if%>>PC</option>
				<option value="2" <% If "2" = cstr(SearchGubun) Then%> selected <%End if%>>MOBILE</option>
				<option value="3" <% If "3" = cstr(SearchGubun) Then%> selected <%End if%>>MOVIE</option>
				<option value="4" <% If "4" = cstr(SearchGubun) Then%> selected <%End if%>>MOBILE���</option>
			</select>
			&nbsp;&nbsp;
			<b> �� �� : </b>
			<select name="SearchUsing">
				<option value ="" style="color:blue">�� ü</option>
				<option value="Y" <% If "Y" = cstr(SearchUsing) Then%> selected <%End if%>>Y</option>
				<option value="N" <% If "N" = cstr(SearchUsing) Then%> selected <%End if%>>N</option>
			</select>
			&nbsp;&nbsp;
			<b> Ÿ��Ʋ : </b>
			<input type=text value ="<%= SearchTitle %>" name="evt_title" onKeydown="javascript:if(event.keyCode == 13) form.submit();">
			<input type="button" class="button" value="�˻�����Reset" onClick="location.href='about_contents_list.asp?reload=on&menupos=<%=menupos%>';">
		</td>
		<td lowsapn="2" with="50" bgcolor="<%=admincolor("gray")%>">
			<input type="button" class="button" value="�˻�" onclick="searchFrm('');">&nbsp;
		</td>
	</tr>
</table>
</form>
<!--�˻���----------------------------------------------------------------------------------------------->
<br>
<!--�ű��Է¹�ư---------------------------------------------------------------------------------------->
<table width="100%" align="center">
	<tr>
		<td align="right"><input type="button" name="sizeBT" class="button" value="���������" onclick="sizewrite();"> <input type="button" name="newBT" class="button" value="�ű��Է�" onclick="conwrite('');"></td>
	</tr>
</table>
<!--�ű��Է¹�ư��------------------------------------------------------------------------------------->

<!--����Ʈ----------------------------------------------------------------------------------------------->
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="7" > <!--���պ�(colspan)7��-->
			�˻���� : <b><%= iTotCnt %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
		<td width="10%"><b>��ȣ</b></td>
		<td width="5%"><b>����</b></td>
		<td width="20%"><b>�����</b></td>
		<td width="30%"><b>Ÿ��Ʋ</b></td>
		<td width="5%"><b>��뿩��</b></td>
		<td width="15%"><b>������</b></td>
		<td width="15%"><b>�����</b></td>
	</tr>
	
	<% if opart.FResultCount > 0 then %>
	
		<% for i = 0 to opart.FResultCount - 1 %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF'; height="30"> 
			<td style="cursor:hand"  onclick="conwrite('<%= opart.FItemList(i).Fcontentsidx %>');"><%= opart.FItemList(i).Fcontentsidx %></td><!--��ȣ(idx)-->
			
			<td><%= getHitchhikerGubun(opart.FItemList(i).Fgubun) %></td><!--����(PC,MOBILE,MOVIE-->
			
			<td><img src="<%= opart.FItemList(i).Fcon_viewthumbimg %>" width="50" height="50"></td><!--�����-->
	
			<td><%= ReplaceBracket(opart.FItemList(i).Fcon_title) %></td><!--Ÿ��Ʋ-->
	
			<td><%= opart.FItemList(i).FIsusing %></td><!--��뿩��-->
	
			<td align="left">
				<% 
					Response.Write replace(left(opart.FItemList(i).FSdate,10),"-",".") '������
		
					if now() >=  opart.FItemList(i).FSdate then '����,���¿��� ���
						Response.write " <span style=""color:blue"">(����)</span>"
					else
						Response.write " <span style=""color:green"">(���¿���)</span>"
					end if
					Response.Write "<br />"
				%>
			</td>
			<td><% Response.Write left(opart.FItemList(i).FRegdate,22) %></td><!--�����-->
		</tr>
		<% next %>
		<!--����¡ó��------------------------------------------>
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
		<!--����¡ó����------------------------------------------>	
	<% else %>	
		<tr>
			<td colspan=7 align="center">
				�˻������ �����ϴ�.
			</td>
		</tr>
	<% end if %>
</table>
<!--����Ʈ��----------------------------------------------------------------------------------------------->
<%
set opart = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->