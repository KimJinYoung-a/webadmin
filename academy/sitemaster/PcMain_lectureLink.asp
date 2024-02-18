<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ΰŽ� ��ī���� PC���� �۰�&���� ��ũ
' History : 2016-10-24 ���¿� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/academy/PcMainLectureLinkCls.asp" -->
<%
	Dim oPcMainLectureLink, i , page , idx , startdate , titletext , contentstext, isusing, lectureid
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)
	startdate = requestCheckVar(request("startdate"),10)	'������
	titletext = requestCheckVar(request("titletext"),255)	''����
	contentstext = requestCheckVar(request("contentstext"),255)	''����
	isusing = requestCheckVar(request("isusing"),1)	''��뿩��
	lectureid = requestCheckVar(request("lectureid"),32)	''����/�۰� ID
  	if titletext <> "" then
		if checkNotValidHTML(titletext) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if contentstext <> "" then
		if checkNotValidHTML(contentstext) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	if page = "" then page = 1
	if isusing = "" then isusing = "Y"

set oPcMainLectureLink = new CPcMainLectureLinkContents
	oPcMainLectureLink.FPageSize = 20
	oPcMainLectureLink.FCurrPage = page
	oPcMainLectureLink.FRecttitletext = titletext
	oPcMainLectureLink.FRectcontentstext = contentstext
	oPcMainLectureLink.FRectlectureid = lectureid
	oPcMainLectureLink.FRectisusing = isusing
	oPcMainLectureLink.fnGetPcMainLectureLinkList()
%>
<script type="text/javascript">
	function NextPage(page){
		frm.page.value = page;
		frm.submit();
	}

	function AddNewContents(idx){
		var popwin = window.open('/academy/sitemaster/poppcmainlecturelinkEdit.asp?idx=' + idx,'pcmainlecturelinkEdit','width=700,height=800,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	function jsSerach(){
		var frm;
		frm = document.frm;
		frm.target = "_self";
		frm.action ="PcMain_lectureLink.asp";
		frm.submit();
	}

	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//�̹��� Ȯ�� ��â���� �����ֱ�
	function showimage(img){
		var pop = window.open('/lib/showimage.asp?img='+img,'imgview','width=600,height=600,resizable=yes');
	}

</script>

<form name="frm" method="post" style="margin:0px;">	
<input type="hidden" name="page" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	
	<td align="left">
    ��뱸��
	<select name="isusing">
	<option value="">��ü
	<option value="Y" <% if isusing="Y" then response.write "selected" %> >�����
	<option value="N" <% if isusing="N" then response.write "selected" %> >������
	</select>
	&nbsp;&nbsp;&nbsp;
	�۰�/����ID �˻� : <input type="text" name="lectureid" size=20 value="<%=lectureid%>" />
<!--	����˻� : <input type="text" name="titletext" size=20 value="<%'=titletext%>" />-->
	</td>	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onclick="javascript:jsSerach();">
	</td>

</tr>
</table>
</form>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
	<td align="right">
		<input type="button" class="button" value="�űԵ��" onclick="AddNewContents('0');">
	</td>
</tr>
</table>
<!-- �׼� �� -->
<font color="red">���ֱ� ��ϼ����� �������� �����̰ų� ���ú��� �����ɷ� 1�� �ѷ���</font>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr bgcolor="#FFFFFF">
	<td colspan="6">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				�˻���� : <b><%= oPcMainLectureLink.FTotalCount%></b>
				&nbsp;
				������ : <b><%= page %> / <%=  oPcMainLectureLink.FTotalpage %></b>
			</td>
			<td align="right"></td>			
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="3%">idx</td>
	<td width="15%">�۰�/���� ID</td>
	<td width="15%">����</td>
	<td width="15%">����</td>
	<td width="5%">������</td>
	<td width="5%">����</td>
</tr>
<% if oPcMainLectureLink.FresultCount > 0 then %>
<% for i=0 to oPcMainLectureLink.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center"><%= oPcMainLectureLink.FItemList(i).Fidx %></td>
	<td align="center"><%= oPcMainLectureLink.FItemList(i).Flectureid %></td>
	<td align="center"><%= db2html(oPcMainLectureLink.FItemList(i).Ftitletext) %></td>
	<td align="center"><%= db2html(oPcMainLectureLink.FItemList(i).Fcontentstext) %></td>
	<td align="center"><%= left(oPcMainLectureLink.FItemList(i).Fstartdate,10) %></td>
	<td align="center"><input type="button" class="button" value="����" onclick="AddNewContents('<%= oPcMainLectureLink.FItemList(i).Fidx %>');"/></td>
</tr>
<% Next %>
<tr>
	<td colspan="6" align="center" bgcolor="#FFFFFF">
	 	<% if oPcMainLectureLink.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oPcMainLectureLink.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oPcMainLectureLink.StartScrollPage to oPcMainLectureLink.FScrollCount + oPcMainLectureLink.StartScrollPage - 1 %>
			<% if i>oPcMainLectureLink.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oPcMainLectureLink.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="6" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% end if %>
</table>
<% set oPcMainLectureLink = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->