<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �ΰŽ�
' History : 2010.10.12 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%
dim yyyy2,mm2,dd2 ,lec_idx, lec_title, lecturdate ,lecturdateyn ,page ,yyyy1,mm1,nowdate ,i
	lec_idx = RequestCheckvar(request("lec_idx"),10)
	lec_title = RequestCheckvar(request("lec_title"),64)
	page = RequestCheckvar(request("page"),10)
	if page="" then page=1
	nowdate = now()
	yyyy1 = RequestCheckvar(request("yyyy1"),4)
	mm1   = RequestCheckvar(request("mm1"),2)
  	if lec_title <> "" then
		if checkNotValidHTML(lec_title) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end if
	if yyyy1="" then
		yyyy1 = Left(Cstr(nowdate),4)
		mm1	  = Mid(Cstr(nowdate),6,2)
	end if
	
	lecturdateyn = RequestCheckvar(request("lecturdateyn"),2)
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
	olecture.FRectSearchidx = lec_idx
	olecture.FRectSearchYYYYMM = yyyy1 + "-" + mm1
	olecture.FRectSearchLecturer = Session("ssBctId")
	olecture.FRectSearchTitle = lec_title
	
	if lecturdateyn="on" then
		olecture.FRectSearchLectureDay = lecturdate
	end if
	
	olecture.GetNewLectureList
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

function popwaiting(v){
	popwin = window.open('pop_waituser_list.asp?lec_idx='+ v + '&menupos=<%=menupos%>','popwait','width=720,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

window.onload = GetOnload;

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get>
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" >
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		�˻��� 		: <% DrawYMBox yyyy1,mm1 %>&nbsp;
		�����ڵ�	: <input type="text" name="lec_idx" size="8" value="<%= lec_idx %>">&nbsp;
		���¸� 		:	<input type="text" name="lec_title" size="20"  value="<%= lec_title %>"><br>
		<input type="checkbox" name="lecturdateyn" <% if lecturdateyn = "on" then response.write "checked" %> onclick="ckEnabled(this)">
		������ 		: <% DrawOneDateBox2 yyyy2,mm2,dd2 %>			
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
<% if olecture.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= olecture.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= olecture.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">�̹���</td>
	<td align="center">�����ڵ�<br>�ɼ��ڵ�</td>
	<td align="center">���¸�</td>
	<td align="center">�����</td>
	<td align="center">����(����)��</td>
	<td align="center">�����Ⱓ</td>
	<td align="center">������</td>
	<td align="center">����<br>��û�ο�(����)</td>	
	<td align="center">����ο�<br>��û����</td>
	<td align="center">����<br>����</td>	
</tr>
<% for i=0 to olecture.FResultCount - 1 %>
<% if olecture.FItemList(i).FIsUsing="N" then %>
<tr align="center" bgcolor="#eeeeee" onmouseover=this.style.background="orange"; onmouseout=this.style.background='eeeeee';>
<% else %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% end if %>
	<td><img src="<%= olecture.FItemList(i).Fsmallimg %>" width="50" height=50 border="0"></td>
	<td>
		<%= olecture.FItemList(i).Fidx %>
		<br><%= olecture.FItemList(i).FlecOption %>
	</td>
	<td><%= olecture.FItemList(i).Flec_title %></td>
	<td>
		<%= olecture.FItemList(i).Flecturer_id %>
		<br>(<%= olecture.FItemList(i).Flecturer_name %>)
	</td>
	<td><%= olecture.FItemList(i).Flec_startday1 %></td>
	<td align="center"><%= olecture.FItemList(i).Freg_startday %><br>~<br><%= olecture.FItemList(i).Freg_endday %></td>
	<td align="right">
		<%
		Response.Write FormatNumber(olecture.FItemList(i).Flec_cost,0)
		'������
		if olecture.FItemList(i).FlecturerCouponYn="Y" then
			Select Case olecture.FItemList(i).FlecturerCouponType
				Case "1"
					Response.Write "<br><font color=#5080F0><img src='http://fiximage.10x10.co.kr/web2008/category/icon_coupon.gif' border=0> " & FormatNumber(olecture.FItemList(i).Flec_cost*((100-olecture.FItemList(i).FlecturerCouponValue)/100),0) & ""
					Response.Write "<br>-"&olecture.FItemList(i).FlecturerCouponValue&"%</font>"
				Case "2"
					Response.Write "<br><font color=#5080F0><img src='http://fiximage.10x10.co.kr/web2008/category/icon_coupon.gif' border=0> " & FormatNumber(olecture.FItemList(i).Flec_cost-olecture.FItemList(i).FlecturerCouponValue,0) & ""
					Response.Write "<br>-"&olecture.FItemList(i).FlecturerCouponValue&"</font>"			
			end Select
		end if
		%>		
	</td>
	<td>
		<%= olecture.FItemList(i).Flimit_count %>
		<br><%= olecture.FItemList(i).Flimit_sold %>
	</td>
	<td>
		<a href="javascript:popwaiting('<%= olecture.FItemList(i).Fidx %>')"><%= olecture.FItemList(i).FWaitCount %></a>
		<br><a href="lectureOrderlist.asp?searchfield=itemid&itemid=<%= olecture.FItemList(i).Fidx %>&menupos=<%=menupos%>">
		<img src="http://webadmin.10x10.co.kr/images/icon_search.jpg" width="16" border="0" align="absbottom"></a>		
	</td>
	<td>
		<% if olecture.FItemList(i).IsSoldOut then %>
		<font color="#CC3333">����</font>
		<% end if %>
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

<%
	set olecture = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->