<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<%
	public Sub SelectLecturerId(byval lecturer_id)
		dim sqlStr,i
		sqlStr = "select  c.userid,p.company_name,c.defaultmargine, c.regdate"
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c"
		sqlStr = sqlStr + " left join [db_partner].[dbo].tbl_partner p on c.userid=p.id"
		sqlStr = sqlStr + " where c.userid<>''" + vbCrlf
		sqlStr = sqlStr + " and c.userdiv < 22" + vbcrlf
		sqlStr = sqlStr + " and c.userdiv='14'" + vbcrlf
	
		rsget.open sqlStr,dbget,1
	
		if not rsget.eof then
				response.write "<select name='lecturerID'>"
				response.write "<option value=''>����</option>"
			for i=0 to rsget.recordcount-1
				if lecturer_id=db2html(rsget("userid")) then
				response.write "<option value='" & db2html(rsget("userid")) & "' selected>" & db2html(rsget("userid")) & "(" & db2html(rsget("company_name")) & ")</option>"
				else
				response.write "<option value='" & db2html(rsget("userid")) & "'>" & db2html(rsget("userid")) & "(" & db2html(rsget("company_name")) & ")</option>"
				end if
			rsget.movenext
			next
				response.write "</select>"
		end if
		rsget.close
	
	end sub
%>
<script language='javascript' src="/js/js_minical_min.js"></script>
<script language='javascript' src="/js/js_minical_ko.js"></script>
<script language='javascript' src="/js/js_minical_setup.js"></script>
<link rel="stylesheet" href="/js/js_minical.css" type="text/css"  />
<style>
.display_date { cursor:pointer; width:80px; border:1px solid; border-color:#a6a6a6 #d8d8d8 #d8d8d8 #a6a6a6; height:1em; padding:1px; }
</style>
<script language='javascript'>
<!--
	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.evtDivCd.value)
		{
			alert("�̺�Ʈ ������ �������ֽʽÿ�.");
			frm.evtDivCd.focus();
			return false;
		}

		if(!frm.evtTitle.value)
		{
			alert("������ �Է����ֽʽÿ�.");
			frm.evtTitle.focus();
			return false;
		}

		if(!frm.evtCont.value)
		{
			alert("������ �ۼ����ֽʽÿ�.");
			frm.evtCont.focus();
			return false;
		}

		// �� ����
		return true;
	}

	// �̺�Ʈ ���� ����
	function chgEvtDiv(dv) {
		if(dv=="J020") {
			//���ᰭ�� �̺�Ʈ
			document.all.lyrLecUID.style.display='';
		} else {
			//�Ϲ� �̺�Ʈ
			document.all.lyrLecUID.style.display='none';
		}
	}

	// �̺�Ʈ ���� ����
	function chgEvtType(tp) {
		if(tp=="M") {
			document.all.lyrImage.style.display='';
			document.all.lyrTitle.innerHTML='Image Map';
			document.frm_write.evtCont.value="<map name='evtMainImg'>\n</map>";
		} else {
			document.all.lyrImage.style.display='none';
			document.all.lyrTitle.innerHTML='HTML';
			document.frm_write.evtCont.value="";
		}
	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="<%=imgFingers%>/linkweb/doEvent.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>�̺�Ʈ �ű� ���</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�̺�Ʈ ����</td>
	<td bgcolor="#FFFFFF">
		<select name="evtDivCd" onchange="chgEvtDiv(this.value)">
			<option value="">::����::</option>
			<% call sbOptCommCd("","J000") %>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
	<td bgcolor="#FFFFFF"><input type="text" name="evtTitle" size="40" maxlength="40"></td>
</tr>
<tr align="center" bgcolor="#DDDDFF" id="lyrLecUID" name="lyrLecUID" style="display:none;">
	<td width="120">��� ����</td>
	<td bgcolor="#FFFFFF" align="left">
		<% SelectLecturerId("") %>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" cellpadding="0" cellspacing="2" border="0" class="a">
		<tr>
			<td width="100" bgcolor="#F3F0F8" align="center">�̺�Ʈ����</td>
			<td>
				<input type="radio" name="evtType" value="M" checked onclick="chgEvtType(this.value)">�Ϲ�����
				<input type="radio" name="evtType" value="H" onclick="chgEvtType(this.value)">���۾� ����
			</td>
		</tr>
		<tr id="lyrImage" name="lyrImage">
			<td bgcolor="#F3F0F8" align="center">���� �̹���</td>
			<td><input type="file" name="contImage" size="60"></td>
		</tr>
		<tr>
			<td id="lyrTitle" name="lyrTitle" bgcolor="#F3F0F8" align="center">Image Map</td>
			<td><textarea name="evtCont" rows="14" cols="80"><map name="evtMainImg"><%=vbCrLf%></map></textarea></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#F3F0F8" align="center">���û���</td>
			<td>
				<input type="radio" name="isComment" value="1">�ڸ�Ʈ ���
				<input type="radio" name="isComment" value="0" checked>�ڸ�Ʈ ������
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>�Ⱓ</td>
	<td bgcolor="#FFFFFF">
		<span class="display_date" id="strSDt"><%=date%></span> ~
		<span class="display_date" id="strEDt"><%=date%></span>
		<input type="hidden" name="evtSdate" value="<%=date%>">
		<input type="hidden" name="evtEdate" value="<%=date%>">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>��÷�� ��ǥ��</td>
	<td bgcolor="#FFFFFF">
		<span class="display_date" id="strPDt"><%=date%></span>
		<input type="hidden" name="prizeDate" value="<%=date%>">
		<script language="javascript">
			DyCalendar.setup( { firstDay : 0, inputField : "prizeDate", ifFormat : "%Y-%m-%d", displayArea : "strPDt", daFormat : "%Y-%m-%d"});
			DyCalendar.setup( { firstDay : 0, inputField : "evtSdate", ifFormat : "%Y-%m-%d", displayArea : "strSDt", daFormat : "%Y-%m-%d"});
			DyCalendar.setup( { firstDay : 0, inputField : "evtEdate", ifFormat : "%Y-%m-%d", displayArea : "strEDt", daFormat : "%Y-%m-%d"});
		</script>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">��� �̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="listImage" size="60">
	</td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->