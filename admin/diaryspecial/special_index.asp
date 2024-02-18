<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : ����� ���̾ ������ ����
'	History		: 2015.10.05 ���¿� ����
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/diaryspecial/diaryspecialCls.asp"-->

<%
Dim i
Dim Searchevtcode, Searchitemid, SearchUsing
Dim FResultCount, iCurrentpage, iTotCnt

	Searchevtcode = trim(request("evtcode"))
	Searchitemid = trim(request("itemid"))
	SearchUsing = trim(request("isusing"))
	iCurrentpage = NullFillWith(requestCheckVar(Request("IC"),10),1)

if iCurrentpage="" then iCurrentpage=1
	
Dim ospecial
set ospecial = new CDiaryspecial
	ospecial.FCurrPage = iCurrentpage
	ospecial.FPageSize = 15
	ospecial.FrectIsusing	= SearchUsing
	ospecial.Frectevtcode	= Searchevtcode
	ospecial.Frectitemid	= Searchitemid
	ospecial.fnGetDiaryspecial
iTotCnt = ospecial.FTotalCount
%>

<script type="text/javascript">
	function conwrite(idx){
		var conwrite = window.open('/admin/diaryspecial/DiaryspecialReg.asp?idx='+idx,'DiaryspecialReg','width=800,height=768,scrollbars=yes,resizable=yes');
		conwrite.focus();
	}
	function searchFrm(p){
		frm.iC.value = p;
		frm.submit();
	}
</script>

<!--�˻�------------------------------------------------------------------------------------------------->
<form name="frm" action="special_index.asp" method="get">
<input type="hidden" name="iC" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=admincolor("tablebg")%>">
	<tr align="center" bgcolor="#FFFFFF">
		<td lowsapn="2" with="50" bgcolor="<%=admincolor("gray")%>"> <b>�˻�����</b> </td>
		<td align="left">
			<b> �� �� : </b>
			<select name="isusing">
				<option value ="" style="color:blue">�� ü</option>
				<option value="Y" <% If "Y" = cstr(SearchUsing) Then %> selected <% End if %>>Y</option>
				<option value="N" <% If "N" = cstr(SearchUsing) Then %> selected <% End if %>>N</option>
			</select>
			&nbsp;&nbsp;
			<!--
			<b> ��ǰ�ڵ� : </b>
			<input type=text value ="<%= Searchitemid %>" name="itemid" onKeydown="javascript:if(event.keyCode == 13) form.submit();">
			-->
			&nbsp;&nbsp;
			<b> ��ũ�ڵ� : </b>
			<input type=text value ="<%= Searchevtcode %>" name="evtcode" onKeydown="javascript:if(event.keyCode == 13) form.submit();">&nbsp;&nbsp;&nbsp;
			<input type="button" class="button" value="�˻�" onclick="searchFrm('');">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			
			<input type="button" class="button" value="�˻�����Reset" onClick="location.href='special_index.asp?reload=on&menupos=<%=menupos%>';">
		</td>
	</tr>
</table>
</form>
<!--�˻���----------------------------------------------------------------------------------------------->
<br>
<!--�ű��Է¹�ư---------------------------------------------------------------------------------------->
<table width="100%" align="center">
	<tr>
		<td align="left"><input type="button" value="���ΰ�ħ" onclick="document.location.reload();"></td>
		<td align="right"><input type="button" name="newBT" class="button" value="�ű��Է�" onclick="conwrite('');"></td>
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
		<td width="20%"><b>�����</b></td>
		<td width="20%"><b>���θ�ũ����</b></td>
		<td width="20%"><b>���θ�ũ�ڵ�</b></td>
		<td width="5%"><b>��뿩��</b></td>
		<td width="5%"><b>���ļ���</b></td>
		<td width="15%"><b>�����</b></td>
	</tr>
	
	<% if ospecial.FResultCount > 0 then %>
	
		<% for i = 0 to ospecial.FResultCount - 1 %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF'; height="30"> 

			<!--��ȣ(idx)-->
			<td style="cursor:hand"  onclick="conwrite('<%= ospecial.FItemList(i).Fidx %>');"><%= ospecial.FItemList(i).Fidx %></td>

			<!--�����-->
			<td><img src="<%= ospecial.FItemList(i).Fpcmainimage %>" width="50" height="50"></td>

			<!--�����̹��� ��ũ ����(�̺�Ʈ,��ǰ��)-->
			<td>
				<%
				if ospecial.FItemList(i).Flinkgubun = "i" then
					response.write "��ǰ"
				else
					response.write "�̺�Ʈ"
				end if
				%>
			</td>

			<!--�����̹��� ��ũ �ڵ�-->
			<td><%= ospecial.FItemList(i).Flinkcode %></td>

			<!--��뿩��-->
			<td><%= ospecial.FItemList(i).FIsusing %></td>

			<!--���ļ���-->
			<td><%= ospecial.FItemList(i).Fsortnum %></td>

			<!--�����-->
			<td><% Response.Write left(ospecial.FItemList(i).FRegdate,22) %></td>
		</tr>
		<% next %>
		<!--����¡ó��------------------------------------------>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="7" align="center">
		       	<% if ospecial.HasPreScroll then %>
					<span class="list_link"><a href="javascript:searchFrm('<%= ospecial.StartScrollPage-1 %>')">[pre]</a></span> '&menupos=<%=menupos%>
				<% else %>
				[pre]
				<% end if %>
					<% for i = 0 + ospecial.StartScrollPage to ospecial.StartScrollPage + ospecial.FScrollCount - 1 %>
						<% if (i > ospecial.FTotalpage) then Exit for %>
						<% if CStr(i) = CStr(iCurrentpage) then %>
						<span class="page_link"><font color="red"><b><%= i %></b></font></span>
						<% else %>
						<a href="javascript:searchFrm('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
						<% end if %>
					<% next %>
				<% if ospecial.HasNextScroll then %>
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
set ospecial = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->