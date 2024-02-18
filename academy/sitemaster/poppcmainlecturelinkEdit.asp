<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ΰŽ� ��ī���� PC���� �۰�&���� ��ũ �Է�,���� �˾�
' History : 2016-10-24 ���¿� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/academy/PcMainLectureLinkCls.asp" -->
<%
	Dim idx, oPcMainLectureLink, sqlStr
	Dim startdate, titletext, contentstext, lectureid, isusing

	idx = RequestCheckvar(request("idx"),10)
	set oPcMainLectureLink = new CPcMainLectureLinkContents
		 oPcMainLectureLink.FRectIdx = idx

		if idx <> "" Then
			oPcMainLectureLink.GetOneRowPcMainLectureLinkContent()
			if oPcMainLectureLink.FResultCount > 0 then
				titletext	= oPcMainLectureLink.FOneItem.Ftitletext
				contentstext	= oPcMainLectureLink.FOneItem.Fcontentstext
				lectureid	= oPcMainLectureLink.FOneItem.Flectureid
				startdate	= oPcMainLectureLink.FOneItem.Fstartdate
				isusing	= oPcMainLectureLink.FOneItem.Fisusing
			end if
		end if
	set oPcMainLectureLink = Nothing

	if isusing = "" then isusing = "Y"
%>
<script type="text/javascript">
	//''jsPopCal : �޷� �˾�
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//����
	function subcheck(){
		var frm=document.inputfrm;
	    
		if (!frm.titletext.value){
			alert('������ ������ּ���');
			frm.titletext.focus();
			return;
		}

		if (!frm.titletext.value){
			alert('������ ������ּ���');
			frm.titletext.focus();
			return;
		}

		if (!frm.startdate.value){
			alert('�������� ������ּ���');
			frm.startdate.focus();
			return;
		}

//		if (!frm.viewtext1.value){
//			alert('�󼼳����� ������ּ���');
//			frm.viewtext1.focus();
//			return;
//		}

		frm.submit();
	}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="PcMainLectureLinkProc.asp">
<input type="hidden" name="idx" value="<%= idx %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>�ΰŽ�PC���� �۰�/���� ��ũ ���/����</b></font>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">idx</td>
	<td bgcolor="#FFFFFF">
		<b><%=idx%></b>
	</td>
</tr>
<% End If %>
<!---------------------------------------------------------------------------------------->
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�۰�/���� ID</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="lectureid" value="<%=lectureid%>" size="50"/>
	</td>
</tr>

<!---------------------------------------------------------------------------------------->

<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="titletext" value="<%=titletext%>" size="50"/>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��  ��</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="contentstext" value="<%=contentstext%>" size="50"/>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td bgcolor="#FFFFFF">
   		<input type="text" name="startdate" size=20 maxlength=10 value="<%= startdate %>" onClick="jsPopCal('startdate');"  style="cursor:pointer;"/>
		<font color="red">��Ŭ���� �޷¿��� ����</font>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center"> ��뿩�� </td>
	<td colspan="2">
		<input type="radio" name="isusing" value="Y" <%=chkIIF(isusing="Y","checked","")%>/>����� &nbsp;&nbsp;&nbsp; 
		<input type="radio" name="isusing" value="N" <%=chkIIF(isusing="N","checked","")%>/>������
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
		<input type="button" value=" ��� " class="button" onclick="window.close();"/>
	</td>
</tr>
<!---------------------------------------------------------------------------------------->
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->