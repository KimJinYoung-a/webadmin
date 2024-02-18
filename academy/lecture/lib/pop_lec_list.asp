<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%

dim yyyy1,mm1,nowdate
dim yyyy2,mm2,dd2
dim lecturer
dim lec_idx, lec_title, lecturdate
dim lecturdateyn
dim page
Dim weclassYN

lec_idx = RequestCheckvar(request("lec_idx"),10)
lecturer = RequestCheckvar(request("lecturer"),32)
lec_title = request("lec_title")
weclassYN = RequestCheckvar(request("weclassYN"),1)

page = RequestCheckvar(request("page"),10)
if page="" then page=1

nowdate = now()

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1   = RequestCheckvar(request("mm1"),2)

if yyyy1="" then
	yyyy1 = Left(Cstr(nowdate),4)
	mm1	  = Mid(Cstr(nowdate),6,2)
end if

lecturdateyn = RequestCheckvar(request("lecturdateyn"),10)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2   = RequestCheckvar(request("mm2"),2)
dd2   = RequestCheckvar(request("dd2"),2)

if yyyy2="" then
	yyyy2 = Left(Cstr(nowdate),4)
	mm2	  = Mid(Cstr(nowdate),6,2)
	dd2	  = Mid(Cstr(nowdate),9,2)
end if
lecturdate = yyyy2 + "-" + mm2 + "-" + dd2

dim olecture
set olecture = new CLecture
olecture.FCurrPage = page
olecture.FPageSize=20
olecture.FRectidx = lec_idx
olecture.FRectSearchYYYYMM = yyyy1 + "-" + mm1
olecture.FRectSearchLecturer = lecturer
olecture.FRectSearchTitle = lec_title
olecture.FweclassYN = weclassYN

if lecturdateyn="on" then
	olecture.FRectSearchYYYYMM = lecturdate
end if

olecture.GetNewLectureList '������

dim i
%>
<script language='javascript'>
function NextPage(page){
	frm.page.value= page;
	frm.submit();
}

function GetOnload(){
	ckEnabled(frm.lecturdateyn);
}

function ckEnabled(comp){
	frm.yyyy2.disabled = (!comp.checked);
	frm.mm2.disabled = (!comp.checked);
	frm.dd2.disabled = (!comp.checked);
}


window.onload = GetOnload;

</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
    <tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="30" >
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
        	�˻��� : <% DrawYMBox yyyy1,mm1 %>&nbsp;
			�����ڵ� : <input type="text" name="lec_idx" size="8" value="<%= lec_idx %>">&nbsp;���¸� :
			<input type="text" name="lec_title" size="20"  value="<%= lec_title %>"><br>
			���� : <% drawSelectBoxLecturer "lecturer",lecturer  %>

			<input type="checkbox" name="lecturdateyn" <% if lecturdateyn = "on" then response.write "checked" %> onclick="ckEnabled(this)">
			������ : <% DrawOneDateBox2 yyyy2,mm2,dd2 %>
        </td>
        <td valign="top" align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22"  border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">

	<tr align="center" bgcolor="#DDDDFF">
		<td></td>
		<td colspan="13" align="right">�˻��Ǽ� : <%= olecture.FTotalCount %> �� Page : <%= page %>/<%= olecture.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40"></td>
		<td align="center" width="50">�����ڵ�</td>
		<td align="center" width="50">�̹���</td>
		<td align="center">���¸�</td>
		<td align="center" width="70">�����</td>
		<td align="center" width="60">����(����)��</td>

		<td align="center" width="70">�����Ⱓ</td>
		<td align="center" width="40">������</td>
		<td align="center" width="30">����</td>
	</tr>
<% for i=0 to olecture.FResultCount - 1 %>
	<% if olecture.FItemList(i).FIsUsing="N" then %>
	<tr align="center" bgcolor="#EEEEEE">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
	<form name="itemfrm_<%= i %>" method="post" target="doframe" action="/academy/lecture/lib/pop_lec_list2.asp">
	<input type="hidden" name="lec_idx" value="<%= olecture.FItemList(i).Fidx %>">
	<input type="hidden" name="" value="">
	<input type="hidden" name="" value="">
	<input type="hidden" name="" value="">
	<input type="hidden" name="" value="">
	<input type="hidden" name="" value="">

		<td><input type="submit" value="����"></td>
		<td><%= olecture.FItemList(i).Fidx %></td>
		<td><img src="<%= olecture.FItemList(i).Fsmallimg %>" width="50" border="0"></td>
		<td><%= olecture.FItemList(i).Flec_title %></td>
		<td><%= olecture.FItemList(i).Flecturer_id %><br>(<%= olecture.FItemList(i).Flecturer_name %>)</td>
		<td width="70"><%= olecture.FItemList(i).Flec_startday1 %></td>
		<td align="center"><%= olecture.FItemList(i).Freg_startday %><br>~<br><%= olecture.FItemList(i).Freg_endday %></td>
		<td align="right"><%= FormatNumber(olecture.FItemList(i).Flec_cost,0) %></td>
		<td><%= olecture.FItemList(i).Flimit_count %></td>
	</tr>
	</form>
<% next %>
	<tr>
		<td colspan="9" align="Center" bgcolor="#FFFFFF">
			<% if olecture.HasPreScroll then %>
				<a href="javascript:NextPage('<%= olecture.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + olecture.StartScrollPage to olecture.FScrollCount + olecture.StartScrollPage - 1 %>
				<% if i>olecture.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if olecture.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>
<iframe name="doframe" src="" width="0" height="0" frameborder="0"></iframe>


<%
set olecture = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->