<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCompanyOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/company/recruit_cls.asp"-->
<%
	Dim page, SearchArea, SearchKeyword, rcbsn

	rcbsn = Request("rcbsn")
	page = Request("page")
	SearchArea = Request("SearchArea")
	SearchKeyword = Request("SearchKeyword")
	if page="" then page=1


	'// ���� ����
	dim oRecruit, lp
	Set oRecruit = new CRecruit
	oRecruit.FRectrcbSn = rcbsn
	
	oRecruit.getRecruitCont
%>
<!-- ��ܶ� ���� -->
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
	// ���˻� �� ����
	function submitForm() {
		var form = document.frm_upload;

		if(!form.rcb_startdate.value) {
			alert("ä��������ڸ� �Է����ֽʽÿ�.");
			return;
		}
		if(!form.rcb_enddate.value) {
			alert("ä�븶�����ڸ� �Է����ֽʽÿ�.");
			return;
		}
		if(!dateChk(form.rcb_startdate.value,form.rcb_enddate.value)) {
			alert("�������� �����Ϻ��� ���ų� ���� �� �����ϴ�.\n\nä��Ⱓ�� Ȯ�����ֽʽÿ�.");
			return;
		}

		if(!form.rcb_subject.value) {
			alert("������ �Է����ֽʽÿ�.");
			form.rcb_subject.focus();
			return;
		}

		//2017-02-16 ���¿��߰�(��¿���, ä������)
		form.rcb_career.value=0
		form.rcb_career1.value=0
		form.rcb_career2.value=0
	    var chk1 = $("#rcb_career1").is(":checked");
	    var chk2 = $("#rcb_career2").is(":checked");
	    if(chk1) $("#rcb_career1").val(1);
	    if(chk2) $("#rcb_career2").val(2);

		form.rcb_career.value = Number(form.rcb_career1.value)+Number(form.rcb_career2.value);

		if(form.rcb_career.value==0) {
			alert("��� ���θ� �������ּ���.");
			form.rcb_career.focus();
			return;
		}

		var personalchk = $("#rcb_personalchk").is(":checked");
	    if(personalchk){
	    	$("#rcb_personal").val(1);
	    }else{
	    	$("#rcb_personal").val(0);
	    }

		if(confirm("�Է��� �������� �����Ͻðڽ��ϱ�?")) {
			form.mode.value = "modi";
			form.submit();
		} else {		
			return;
		}
	}

	function dateChk(dt1,dt2) {
		//�����ڷ� ������ �迭�� ��ȯ
		v0=dt1.split("-");
		v1=dt2.split("-");

		//���ڿ� �ش��ϴ� Ÿ�ӽ������� ��ȯ
		v0=new Date(v0[0],v0[1],v0[2]).valueOf();
		v1=new Date(v1[0],v1[1],v1[2]).valueOf();

		//�����̸� ���ѵ� �Ϸ翡 �ش��ϴ� ������ ���Ͽ�, �ʴ����� �ϴ����� ��ȯ
		cha=(v1-v0)/(1000*60*60*24);

		if(cha>0)
			return true;
		else
			return false;
	}

	//���� ����
	function deleteItem() {
		var form = document.frm_upload;
		if(confirm("�� ä����� �����Ͻðڽ��ϱ�?\n\n�س����� ������ �����Ǹ� ���� �� �� ���� �˴ϴ�.")) {
			form.mode.value = "del";
		    form.submit();
		}
	}

	function fnalways(){
		var Now = new Date();
		var Nowyear = Now.getFullYear();
		var inpuyNowyear = Nowyear+1;
		var alwayschk1 = $("#rcb_alwayschkbox").is(":checked");
	    if(alwayschk1){
	    	$("#rcb_always").val(1);
			$("#rcb_enddate").val(inpuyNowyear+'-12-31');
			$("#rcb_enddate").hide;
			$("input[name=rcb_enddate]").attr("readonly",true);
			$("#rcb_enddate_trigger").hide();
	    }else{
	    	$("#rcb_always").val(0);
			$("#rcb_enddate").val("");
			$("#rcb_enddate").show;
			$("#rcb_enddate_trigger").show();	    	
	    }
	}
<% if oRecruit.FitemList(1).Frcb_always=1 then %>
	$(function(){
		$("#rcb_enddate_trigger").hide();
	});
<% end if %>
</script>
<script language="JavaScript" src="/js/file_upload.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<form method="post" name="frm_upload" action="<%=uploadUrl%>/linkweb/company/Recruit_process.asp" onsubmit="return false" enctype="multipart/form-data">
<input type="hidden" name="retURL" value="<%=manageUrl%>/company/intro/recruit_list.asp?menupos=<%= menupos %>&page=<%=page%>&SearchArea=<%=SearchArea%>&SearchKeyword=<%=SearchKeyword%>">
<input type="hidden" name="mode" value="modi">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="rcbsn" value="<%=rcbsn%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="SearchArea" value="<%=SearchArea%>">
<input type="hidden" name="SearchKeyword" value="<%=SearchKeyword%>">
<tr height="10" valign="bottom">
	<td background="/images/tbl_blue_round_02.gif"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><b>ä����� �󼼺���/����</b></td>
	<td align="right">&nbsp;</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ��ܶ� �� -->
<!-- ���� ���� ���� -->
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr>
	<td width="200" bgcolor="#E6E6E6" align="center">��ȣ</td>
	<td width="200" bgcolor="#FFFFFF"><b><%=rcbSn%></b></td>
	<td width="200" bgcolor="#E6E6E6" align="center">�ۼ���</td>
	<td width="200" bgcolor="#FFFFFF">
		<%=oRecruit.FitemList(1).Fuserid%>
		<input type="hidden" name="userid" value="<%=oRecruit.FitemList(1).Fuserid%>">
	</td>
	<td width="200"  bgcolor="#E6E6E6" align="center">��ȸ��</td>
	<td width="200" colspan="6" bgcolor="#FFFFFF"><%=oRecruit.FitemList(1).Frcb_hit%></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">�Ⱓ</td>
	<td bgcolor="#FFFFFF">
		<input id="rcb_startdate" name="rcb_startdate" value="<%=left(oRecruit.FitemList(1).Frcb_startdate,10)%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="rcb_startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="rcb_enddate" name="rcb_enddate" value="<%=left(oRecruit.FitemList(1).Frcb_enddate,10)%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="rcb_enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "rcb_startdate", trigger    : "rcb_startdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "rcb_enddate", trigger    : "rcb_enddate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		<input type="hidden" name="rcb_always" id="rcb_always" value=<%= oRecruit.FitemList(1).Frcb_always %> />
		&nbsp;&nbsp;<input type="checkbox" name="rcb_alwayschkbox" id="rcb_alwayschkbox" <% if oRecruit.FitemList(1).Frcb_always=1 then Response.write "checked" %> onclick="fnalways();" />���
	</td>
	<td bgcolor="#E6E6E6" align="center">����</td>
	<td  colspan="8" bgcolor="#FFFFFF">
		<select name="rcb_state">
			<option value="0" <% if oRecruit.FitemList(1).Frcb_state="0" then Response.write "selected" %>>�Ϲ�</oprion>
			<option value="1" <% if oRecruit.FitemList(1).Frcb_state="1" then Response.write "selected" %>>���⸶��</oprion>
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">ä������</td>
	<td bgcolor="#FFFFFF"><input type="text" name="rcb_jobtype" size="80" value="<%=oRecruit.FitemList(1).Frcb_jobtype%>"><Br>
<p>[��ǥ����] MD, ��������, ����, ������, ���� ��ȹ, ����, ������, ������ ����, �濵, �λ����, CS, ���� </p>

<p>�ΰ��� �̻��� ������ ���� �ø� ���. ex) MD / ������ </p>
</td>
	<td width="200" bgcolor="#E6E6E6" align="center">��¿���</td>
	<td width="200"  colspan="8" bgcolor="#FFFFFF">
		<input type="hidden" name="rcb_career" value="0" >
		����<input type="checkbox" name="rcb_career1" id="rcb_career1" value="0" <% if oRecruit.FitemList(1).Frcb_career="1" or oRecruit.FitemList(1).Frcb_career="3" then Response.write "checked" %>>
		���<input type="checkbox" name="rcb_career2" id="rcb_career2" value="0" <% if oRecruit.FitemList(1).Frcb_career="2" or oRecruit.FitemList(1).Frcb_career="3"then Response.write "checked" %>>
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">����</td>
	<td bgcolor="#FFFFFF" colspan="10"><input type="text" name="rcb_subject" size="80" value="<%=oRecruit.FitemList(1).Frcb_subject%>"></td>
</tr>

<tr>
	<td bgcolor="#E6E6E6" width="200" align="center">��������Ʈ URL</td>
	<td bgcolor="#FFFFFF" colspan="10"><input type="text" name="rcb_recruit_url" size="80" value="<%=oRecruit.FitemList(1).Frcb_recruit_url%>"></td>
</tr>

<tr>
	<td bgcolor="#E6E6E6" width="200" align="center">�����ι� ��<br>�ڰݿ�� (�̹���)</td>
	<td bgcolor="#FFFFFF" colspan="10">
		<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1">
		<%
			oRecruit.getRecruitFile()
			'// ������ ���� ��� ��� ����
			if oRecruit.FResultCount>0 then
				for lp=0 to oRecruit.FResultCount-1
		%>
			<tr>
				<td>
					<input type="file" name="uploadFile" size="50" /><br />
					<img src="<%= "http://imgstatic.10x10.co.kr/company/recruit/" & oRecruit.FfileList(lp).Ffile_name %>" height="50" />
					<%= oRecruit.FfileList(lp).Ffile_name %>
					<label><input type="checkbox" name="DeletedFile" value="<%= oRecruit.FfileList(lp).Ffile_sn %>" /> ����</label>
				</td>
			</tr>
		<%
				next
			end if

			if oRecruit.FResultCount<3 then
				for lp=1 to 3-oRecruit.FResultCount
					Response.Write "<tr><td><input type=""file"" name=""uploadFile"" size=""50"" /></td></tr>"
				next
			end if
		%>
		</table>
	</td>
</tr>

<tr>
	<td bgcolor="#E6E6E6" align="center">����</td>
	<td bgcolor="#FFFFFF" colspan="10">
		<textarea name="rcb_content" cols="110" rows="20"  id="rcb_content"><%=oRecruit.FitemList(1).Frcb_content%></textarea>
	</td>
</tr>

<tr>
	<td bgcolor="#E6E6E6" align="center">�������� ���� �� �̿� ����</td>
	<td width="180" colspan="10" bgcolor="#FFFFFF">
		<input type="hidden" name="rcb_personal" id="rcb_personal" value=<%= oRecruit.FitemList(1).Frcb_personal %> >
		<input type="checkbox" name="rcb_personalchk" id="rcb_personalchk" <% if oRecruit.FitemList(1).Frcb_personal=1 then Response.write "checked" %> >&nbsp;'�������� ���� �� �̿� ����' �ٿ�ε� ���(�̸��Ϸ� ���� ������ ���)
	</td>
</tr>

</table>
<!-- ���� ���� �� -->
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr valign="bottom" height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr valign="bottom">
			<td align="center">
				<a href="javascript:submitForm();"><img src="/images/icon_confirm.gif" width="45" border="0" align="absbottom"></a>&nbsp;
				<a href="javascript:history.back();"><img src="/images/icon_cancel.gif" width="45" border="0" align="absbottom"></a>&nbsp;
				<a href="javascript:deleteItem();"><img src="/images/icon_delete.gif" width="45" border="0" align="absbottom"></a>
			</td>
		</tr>
		</table>
	</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
	<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</form>
</table>
<!-- ������ �� -->
<%
set oRecruit = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbCompanyClose.asp" -->