 <%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : �ν�Ÿ�׷� �̺�Ʈ�� ���� ���������
'	History		: 2016.06.23 ���¿� ����
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/only_sys/instagrameventCls.asp"-->

<%
Dim i
Dim FResultCount, iCurrentpage, iTotCnt , eventid
Dim Searchgubun, SearchTitle, SearchUsing, SearchEvtCode
	SearchUsing = request("SearchUsing")
	eventid = request("eventid")

	Response.write "���� �̺�Ʈ �ڵ� : "& eventid

	iCurrentpage = NullFillWith(requestCheckVar(Request("IC"),10),1)
if iCurrentpage="" then iCurrentpage=1
	
Dim oinsta
set oinsta = new CInstagramevent
	oinsta.FCurrPage = iCurrentpage
	oinsta.FPageSize = 15
	oinsta.FrectIsusing = SearchUsing
	oinsta.Feventid = eventid
	oinsta.fnGetInstagrameventList
iTotCnt = oinsta.FTotalCount
%>

<script type="text/javascript">
function conwrite(contentsidx,md){
	var conwrite = window.open('/admin/etc/only_sys/instagramevent_write.asp?mode='+md+'&contentsidx='+contentsidx,'instagramevent_write','width=800,height=768,scrollbars=yes,resizable=yes');
	conwrite.focus();
}

function searchFrm(p){
	frm.iC.value = p;
	frm.submit();
}

</script>


<form name="frm" action="10x10instagram.asp" method="get">
<input type="hidden" name="iC" value="1">
<input type="hidden" name="menupos" value="<%'= menupos %>" >
<!--�˻�-----------------------------------------------------------------------------------------------
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%'=admincolor("tablebg")%>">
	<tr align="center" bgcolor="#FFFFFF">
		<td lowsapn="2" with="50" bgcolor="<%'=admincolor("gray")%>"> <b>�˻�����</b> </td>
		<td align="left">
			<b> �� �� : </b>
			<select name="SearchUsing">
				<option value ="" style="color:blue">�� ü</option>
				<option value="Y" <%' If "Y" = cstr(SearchUsing) Then%> selected <%'End if%>>Y</option>
				<option value="N" <%' If "N" = cstr(SearchUsing) Then%> selected <%'End if%>>N</option>
			</select>
			<input type="button" class="button" value="�˻�����Reset" onClick="location.href='about_contents_list.asp?reload=on&menupos=<%'=menupos%>';">
		</td>
		<td lowsapn="2" with="50" bgcolor="<%'=admincolor("gray")%>">
			<input type="button" class="button" value="�˻�" onclick="searchFrm('');">&nbsp;
		</td>
	</tr>
</table>
�˻���----------------------------------------------------------------------------------------------->
</form>

<br>
<!--�ű��Է¹�ư---------------------------------------------------------------------------------------->
<table width="100%" align="center">
	<tr>
		<td align="right"><input type="button" name="newBT" class="button" value="�ű��Է�" onclick="conwrite('<%=eventid%>','NEW');"></td>
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
		<td width="5%"><b>��ȣ</b></td>
		<td width="5%"><b>�̺�Ʈ�ڵ�</b></td>
		<td width="20%"><b>�Խ���ID</b></td>
		<td width="30%"><b>�̹���URL</b></td>
		<td width="30%"><b>�Խù���ũ</b></td>
		<td width="5%"><b>��뿩��</b></td>
		<td width="10%"><b>�����</b></td>
	</tr>

	<% if oinsta.FResultCount > 0 then %>
	
		<% for i = 0 to oinsta.FResultCount - 1 %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF'; height="30"> 
			<td onclick="conwrite('<%= oinsta.FItemList(i).Fcontentsidx %>','EDIT');"><%= oinsta.FItemList(i).Fcontentsidx %></td><!--��ȣ(idx)-->
			
			<td><%= oinsta.FItemList(i).Fevt_code %></td><!--�̺�Ʈ�ڵ�-->
	
			<td><%= oinsta.FItemList(i).Fuserid %></td><!--�Խ���ID-->
	
			<td><a href="<%=oinsta.FItemList(i).Fimgurl %>" target="_blank"><img src="<%= oinsta.FItemList(i).Fimgurl %>" width="50" height="50"  border=0></a></td><!--�̹���URL-->
	
			<td><%= oinsta.FItemList(i).Flinkurl %></td><!--�Խù���ũ-->
			
			<td onclick="conwrite('<%= oinsta.FItemList(i).Fcontentsidx %>','EDIT');"><%= oinsta.FItemList(i).Fisusing %></td><!--��뿩��-->
			
			<td><% Response.Write left(oinsta.FItemList(i).FRegdate,22) %></td><!--�����-->
		</tr>
		<% next %>
		<!--����¡ó��------------------------------------------>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="15" align="center">
		       	<% if oinsta.HasPreScroll then %>
					<span class="list_link"><a href="javascript:searchFrm('<%= oinsta.StartScrollPage-1 %>')">[pre]</a></span> '&menupos=<%=menupos%>
				<% else %>
				[pre]
				<% end if %>
					<% for i = 0 + oinsta.StartScrollPage to oinsta.StartScrollPage + oinsta.FScrollCount - 1 %>
						<% if (i > oinsta.FTotalpage) then Exit for %>
						<% if CStr(i) = CStr(iCurrentpage) then %>
						<span class="page_link"><font color="red"><b><%= i %></b></font></span>
						<% else %>
						<a href="javascript:searchFrm('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
						<% end if %>
					<% next %>
				<% if oinsta.HasNextScroll then %>
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
set oinsta = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->