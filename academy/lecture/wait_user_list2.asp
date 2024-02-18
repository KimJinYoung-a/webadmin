<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/lecture_waitingusercls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%
dim lec_idx, lecOption, page
dim ckonlyopen, ckonlynotstart, research
dim yyyy1,mm1,nowdate

lec_idx = RequestCheckvar(request("lec_idx"),10)
lecOption = RequestCheckvar(request("lecOption"),4)
page = RequestCheckvar(request("page"),10)
ckonlyopen = RequestCheckvar(request("ckonlyopen"),10)
ckonlynotstart = RequestCheckvar(request("ckonlynotstart"),10)
research = RequestCheckvar(request("research"),10)

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1   = RequestCheckvar(request("mm1"),2)

nowdate = now()

if yyyy1="" then
	yyyy1 = Left(Cstr(nowdate),4)
	mm1	  = Mid(Cstr(nowdate),6,2)
end if


if (research="") and (ckonlyopen="") then ckonlyopen="on"
''if (research="") and (ckonlynotstart="") then ckonlynotstart="on"

if page="" then page=1

'//���� ����
dim olecture
set olecture = new CLecture
olecture.FRectIdx = lec_idx
olecture.FRectLecOpt = lecOption
olecture.FPageSize=1
olecture.FCurrpage=page
if lec_idx<>"" then
	olecture.GetOneLecture
end if

'// �ɼ�����
dim oLectOption
Set oLectOption = New CLectOption
oLectOption.FRectidx = lec_idx
oLectOption.FRectOptIsUsing = "Y"
if lec_idx<>"" then
	oLectOption.GetLectOptionInfo
end if

'// ���������
dim owaiting
set owaiting = new CLecWaitUser
owaiting.FRectLecIdx = lec_idx
owaiting.FRectLecOpt = lecOption
owaiting.FPageSize=50
owaiting.Fcurrpage=page
owaiting.FRectOnlyusing = ckonlyopen
owaiting.FRectOnlyNotStart = ckonlynotstart
owaiting.FRectYYYYMM = yyyy1 + "-" + mm1
owaiting.getWaitingList

dim i
%>
<script language='javascript'>
function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

function PopWaitUserEdit(iidx,lec_idx){
	var popwin = window.open('/academy/lecture/lib/popwaitpersonreg.asp?idx=' + iidx + '&lec_idx=' + lec_idx,'popwaitpersonreg','width=400,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function saveopen(){

	var ret = confirm('������ ������� ���� ����� ����մϴ�.');

	if (ret){
		subcheck();
		realfrm.mode.value="open";
		realfrm.submit();
	}
}


function subcheck(){

	for (var i=0;i<document.forms.length;i++){
		sfrm = document.forms[i];
		if (sfrm.name.substr(0,5)=="wfrm_") {
			if (sfrm.cksel.checked){
				realfrm.arridx.value = realfrm.arridx.value + sfrm.widx.value + "," ;
			}
		}
	}
}

function deluser(){

	var ret = confirm('������ ����ڸ� ��⸮��Ʈ���� �����մϴ�.');

	if (ret){
		subcheck();
		realfrm.mode.value="del";
		realfrm.submit();
	}
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
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
        	�˻��� 		: <% DrawYMBox yyyy1,mm1 %>&nbsp;
			�����ڵ�	: <input type="text" name="lec_idx" size="8" value="<%= lec_idx %>">&nbsp;
			�����ڵ�	: <input type="text" name="lecOption" size="8" value="<%= lecOption %>">&nbsp;
			<input type="checkbox" name="ckonlyopen" <% if ckonlyopen="on" then response.write "checked" %> >���� ���� �˻�����
			<input type="checkbox" name="ckonlynotstart" <% if ckonlynotstart="on" then response.write "checked" %> >���۵� ���� �˻�����
        </td>
        <td valign="top" align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22"  border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
<tr>
	<td align="right"><input type="button" value="����� �űԵ��" onclick="PopWaitUserEdit('0','<% =lec_idx %>')"></td>
</tr>
</table>
<% if olecture.FResultCount>0 then %>
<!-- ���� ���� -->
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
		<tr bgcolor="#FFFFFF">
			<td width="15%" bgcolor="#DDDDFF">�����ڵ�</td>
			<td width="26%"><%= olecture.FOneItem.Fidx %></td>
			<td width="15%" bgcolor="#DDDDFF">���¿�����</td>
			<td width="26%" ><b><%= olecture.FOneItem.Flec_date %></b></td>
			<td width="18%" colspan="2" rowspan="5" align="center"><img src="<%= olecture.FOneItem.Foblong_img4 %>" width="150"></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#DDDDFF">���¸�</td>
			<td ><%= olecture.FOneItem.Flec_title %></td>
			<td bgcolor="#DDDDFF">�귣��</td>
			<td ><%= olecture.FOneItem.Flecturer_id %> (<%= olecture.FOneItem.Flecturer_name %>)</td>
		</tr>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF" >���ÿ���</td>
			<td >
			<% if olecture.FOneItem.Fdisp_yn="Y" then %>
			����
			<% else %>
			<font color="#CC3333">���þ���</font>
			<% end if %>
			</td>
			<td bgcolor="#DDDDFF" >��뿩��</td>
			<td>
			<% if olecture.FOneItem.Fisusing="Y" then %>
			���
			<% else %>
			<font color="#CC3333">������</font>
			<% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#DDDDFF">��������</td>
			<td >
			<% if olecture.FOneItem.IsSoldOut then %>
			<font color="#CC3333"><b>����</b></font>
			<% else %>
			������
			<% end if %>
			</td>
			<td bgcolor="#DDDDFF">��������</td>
			<td >
			<% if olecture.FOneItem.Freg_yn="Y" then %>
			������
			<% else %>
			<font color="#CC3333">��������</font>
			<% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="#DDDDFF">�����Ⱓ</td>
			<td>
			<%= olecture.FOneItem.Freg_startday %>~<%= olecture.FOneItem.Freg_endday %>
			</td>
			<td bgcolor="#DDDDFF" >�����</td>
			<td>
			<%= olecture.FOneItem.Fregdate %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF"><td colspan="6"></td></tr>
		<tr  bgcolor="#FFFFFF">
			<td bgcolor="#DDDDFF">������/���԰�</td>
			<td >
			<%= FormatNumber(olecture.FOneItem.Flec_cost,0) %> / <%= FormatNumber(olecture.FOneItem.Fbuying_cost,0) %>
			</td>
			<td bgcolor="#DDDDFF">����</td>
			<td bgcolor="#FFFFFF" >
			<% if olecture.FOneItem.Fmatinclude_yn="Y" then %>
			����(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
			<% else %>
			����(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
			<% end if %>
			</td>
			<td bgcolor="#DDDDFF">���ϸ���</td>
			<td >
			<%= olecture.FOneItem.Fmileage %> (point)
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF">����-��û= �����ο�</td>
			<td bgcolor="#FFFFFF" >
			  <%= olecture.FOneItem.Flimit_count %> ��
			-
			  <%= olecture.FOneItem.Flimit_sold %> ��
			=
			  <%= olecture.FOneItem.GetRemainNo %> ��
			</td>
			<td bgcolor="#DDDDFF">����ο�</td>
			<td bgcolor="#FFFFFF" colspan="4">
			<%= olecture.FOneItem.FWaitcount %> ��
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" >
			<td bgcolor="#DDDDFF">����-������û= �����ο�</td>
			<td bgcolor="#FFFFFF" >
			  <%= olecture.FOneItem.Flimit_count %> ��
			-
			  <%= olecture.FOneItem.FRealJupsuCount %> ��
			=
			  <%= olecture.FOneItem.Flimit_count-olecture.FOneItem.FRealJupsuCount %> ��
			</td>
			<td bgcolor="#DDDDFF">���������ο�</td>
			<td bgcolor="#FFFFFF" colspan="4">
			<%= olecture.FOneItem.Flimit_count-olecture.FOneItem.FRealJupsuCount %> ��
			</td>
		</tr>
	</table>
<% end if %>
<% if oLectOption.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#DDDDFF">
	<td>�ɼ��ڵ�</td>
	<td>�ɼǸ�</td>
	<td>�����Ⱓ</td>
	<td>������</td>
	<td>�����ο�</td>
	<td>����ο�</td>
	<td>��������</td>
</tr>
<% for i=0 to oLectOption.FResultCount -1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><a href="?lec_idx=<% =oLectOption.FRectidx %>&lecOption=<%=oLectOption.FItemList(i).FlecOption%>&menupos=<%=menupos%>"><%=oLectOption.FItemList(i).FlecOption%></a></td>
	<td><%=oLectOption.FItemList(i).FlecOptionName%></td>
	<td><%=FormatDateTime(oLectOption.FItemList(i).FRegStartDate,2) & "~" & FormatDateTime(oLectOption.FItemList(i).FRegEndDate,2)%></td>
	<td><%=FormatDateTime(oLectOption.FItemList(i).FlecStartDate,1) & " " & FormatDateTime(oLectOption.FItemList(i).FlecStartDate,4) & "~" & FormatDateTime(oLectOption.FItemList(i).FlecEndDate,4)%></td>
	<td><%=oLectOption.FItemList(i).Flimit_count & "��-" & oLectOption.FItemList(i).Flimit_sold & "��= " & (oLectOption.FItemList(i).Flimit_count-oLectOption.FItemList(i).Flimit_sold) & "��"%></td>
	<td><%=oLectOption.FItemList(i).Fwait_count%>��</td>
	<td><% if oLectOption.FItemList(i).IsOptionSoldOut then Response.Write "����"%></td>
</tr>
<% next %>
</table>
<% end if %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<% if owaiting.FResultCount>0 then %>
	<tr>
		<td bgcolor="#FFFFFF" colspan="13" align="left">
			<input type="button" value="���� ����� ����" onclick="javascript:saveopen();">
			<input type="button" value="���� ����� ����" onclick="javascript:deluser();">
		</td>
	</tr>
	<% end if %>
	<tr align="center" bgcolor="#DDDDFF">
		<td width="20"></td>
		<td width="50" align="center">�̹���</td>
		<td width="80" align="center">����</td>
		<td width="30" align="center">���<br>����</td>
		<td width="80" align="center">Userid</td>
		<td width="40" align="center">��û<br>�μ�</td>
		<td width="50" align="center">�̸�</td>
		<td width="80" align="center">����ó</td>
		<td width="100" colspan="2" align="center">����/�ɼ�</td>
		<td align="center">���¸�</td>
		<td width="70" align="center">��û��</td>
		<td width="120" align="center">����������</td>
	</tr>
	<% for i=0 to owaiting.FResultCount -1 %>
	<form name="wfrm_<%= i %>" method="get" action="">
	<input type="hidden" name="widx" value="<%= owaiting.FItemList(i).FIdx %>">
	<% if owaiting.FItemList(i).FIsusing="N" then %>
	<tr bgcolor="#EEEEEE">
	<% else %>
	<tr bgcolor="#FFFFFF">
	<% end if %>
		<td >
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if (owaiting.FItemList(i).FCurrstate>=3) or (owaiting.FItemList(i).Fisusing="N") then response.write "disabled" %> >
		</td>
		<td align="center"><img src="<% =owaiting.FItemList(i).FSmallimg %>" width="50"></td>
		<td align="center"><font color="<%= owaiting.FItemList(i).GetStateNameColor %>"><%= owaiting.FItemList(i).GetStateName %></font></td>
		<td align="center"><% =owaiting.FItemList(i).FRegrank %></td>
		<td align="left"><a href="javascript:PopWaitUserEdit('<% =owaiting.FItemList(i).FIdx %>','<% =owaiting.FItemList(i).Flec_idx %>');"><% =owaiting.FItemList(i).FUserid %></a></td>
		<td align="center">
		<% if owaiting.FItemList(i).FRegcount>1 then %>
		<b><%= owaiting.FItemList(i).FRegcount %></b>
		<% else %>
		<%= owaiting.FItemList(i).FRegcount %>
		<% end if %>
		</td>
		<td align="center"><a href="javascript:PopWaitUserEdit('<% =owaiting.FItemList(i).FIdx %>','<% =owaiting.FItemList(i).Flec_idx %>');"><% =owaiting.FItemList(i).Fuser_name %></a></td>
		<td align="center"><% =owaiting.FItemList(i).Fuser_phone %></td>
		<td align="center"><a href="?lec_idx=<% =owaiting.FItemList(i).Flec_idx %>&menupos=<%=menupos%>"><% =owaiting.FItemList(i).Flec_idx %></a></td>
		<td align="center"><% =owaiting.FItemList(i).FlecOption %></td>
		<td align="left"><% =owaiting.FItemList(i).Flec_title %></td>
		<td align="center"><% =owaiting.FItemList(i).FRegdate %></td>
		<td align="center"><% =owaiting.FItemList(i).FRegEndDay %></td>
	</tr>
	</form>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="13" align="center">
		<% if owaiting.HasPreScroll then %>
		<a href="javascript:NextPage('<%= owaiting.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + owaiting.StarScrollPage to owaiting.FScrollCount + owaiting.StarScrollPage - 1 %>
			<% if i>owaiting.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if owaiting.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<form name="realfrm" method="post" action="/academy/lecture/lib/doLecwait.asp">
<input type="hidden" name="arridx" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="lec_idx" value="<%= lec_idx %>">

</form>
<%
set olecture = Nothing
set owaiting = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->