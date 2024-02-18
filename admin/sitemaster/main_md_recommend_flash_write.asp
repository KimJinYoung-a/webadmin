<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description : PC���ΰ��� MD��
' History : ������ ����
'			2022.07.01 �ѿ�� ����(isms�������ġ)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_event_rotationcls.asp"-->
<%
dim idx,mode
idx = request("idx")
mode = request("mode")
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>
function SubmitForm(){

	if (GetByteLength(document.SubmitFrm.textinfo.value) > 64){
		alert('TEXT ������ 64�� ���Ϸ� �Է� ���ּ���.\n(���� ' + GetByteLength(document.SubmitFrm.textinfo.value) + '�� �Է� ��)');
		document.SubmitFrm.textinfo.focus();
		return;
	}

//	if (document.SubmitFrm.linkinfo.value.length < 1){
//		alert('��ũ ������ �Է� �ϼ���');
//		document.SubmitFrm.linkinfo.focus();
//		return;
//	}

	if (document.SubmitFrm.disporder.value.length < 1){
		alert('���� ������ �Է� �ϼ���');
		document.SubmitFrm.disporder.focus();
		return;
	}

    if (document.SubmitFrm.startdate.value.length!=10){
        alert('�������� �Է�  �ϼ���.');
        return;
    }

    if (document.SubmitFrm.enddate.value.length!=10){
        alert('�������� �Է�  �ϼ���.');
        return;
    }

    var vstartdate = new Date(document.SubmitFrm.startdate.value.substr(0,4), (1*document.SubmitFrm.startdate.value.substr(5,2))-1, document.SubmitFrm.startdate.value.substr(8,2));
    var venddate = new Date(document.SubmitFrm.enddate.value.substr(0,4), (1*document.SubmitFrm.enddate.value.substr(5,2))-1, document.SubmitFrm.enddate.value.substr(8,2));

    if (vstartdate>venddate){
        alert('�������� �����Ϻ��� ������ �ȵ˴ϴ�.');
        return;
    }


	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if (ret) {
		document.SubmitFrm.submit();
	}
}

</script>

<div style="padding:5px 3px; margin:5px 0; font-size:13px;background:#F0F0FF;"><strong>[����] MD ��õ��ǰ ���/����</strong></div>
<form name="SubmitFrm" method="post" action="<%=uploadUrl%>/linkweb/doMainMdChoiceRotate.asp" onsubmit="return false;" enctype="multipart/form-data">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="regID" value="<%=session("ssBctId")%>">
<%
	if mode = "modify" then
		dim mdchoicerotate
	
		set mdchoicerotate = new CMainMdChoiceRotate
		mdchoicerotate.FCurrPage = 1
		mdchoicerotate.FPageSize = 1
		mdchoicerotate.read idx
%>
<input type="hidden" name="idx" value="<% = idx %>">
<input type="hidden" name="updateID" value="<%=session("ssBctId")%>">
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">�̹���</td>
		<td><input type="file" name="photoimg" value="" size="32" maxlength="32" class="text">
			<br>
			<img src="<%= mdchoicerotate.FItemList(0).Fphotoimg %>" style="max-width:550px; max-height:120px;"><br/>
			<font color="red">(119px �� 135px GIF Ȥ�� JPG �̹��� / �� 2013������: ��� �̹��� ������ ��ǰ�̹��� ���)</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">���ü���</td>
		<td><input type="text" name="disporder" value="<% = mdchoicerotate.FItemList(0).Fdisporder  %>" size="2" class="text">
			<font color="red">(2�ڸ� ����)</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">��ǰ�ڵ�</td>
		<td><input type="text" name="linkitemid" value="<%= ReplaceBracket(mdchoicerotate.FItemList(0).Flinkitemid)  %>" size="6" class="text"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">Text����</td>
		<td>
			<textarea name="textinfo" class="textarea" style="width:90%; height:42px;"><%= ReplaceBracket(mdchoicerotate.FItemList(0).Ftextinfo) %></textarea>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">link����</td>
		<td><input type="text" name="linkinfo" value="<% = mdchoicerotate.FItemList(0).Flinkinfo  %>" size="70" class="text">
			<br>
			<font color="red">(����η� �Է��ϼ��� /shopping/category_prd.asp?itemid=72367)</font>
			<br><font color="red">(��ũ���� ���� ������ ��� ��ǰ�ڵ带 ������� �ڵ����� ��ũ���� ��ü�˴ϴ�.)</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">������ ǥ��</td>
		<td>
			<label><input type="radio" name="lowestPrice" value="Y" <% If trim(mdchoicerotate.FItemList(0).FLowestPrice) = "Y" Then %>checked<% End If %> >���</label>
			<label><input type="radio" name="lowestPrice" value="N" <% If trim(mdchoicerotate.FItemList(0).FLowestPrice) = "N" or isnull(mdchoicerotate.FItemList(0).FLowestPrice) Then %>checked<% End If %> >������</label>
		</td>
	</tr>		

	<tr bgcolor="#FFFFFF">
	    <td width="150" bgcolor="#DDDDFF">�ݿ�������</td>
	    <td>
	        <input id="startdate" name="startdate" value="<%= Left(mdchoicerotate.FItemList(0).Fstartdate,10) %>" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	        <input type="text" name="startdatetime" size="2" maxlength="2" value="<%= Format00(2,Hour(mdchoicerotate.FItemList(0).Fstartdate)) %>" />(�� 00~23)
	        <input type="text" name="dummy0" value="00:00" size="6" readonly class="text_ro" />
		    <script type="text/javascript">
			var CAL_Start = new Calendar({
				inputField : "startdate",
				trigger    : "startdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				},
				bottomBar: true,
				dateFormat: "%Y-%m-%d"
			});
			</script>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" bgcolor="#DDDDFF">�ݿ�������</td>
	    <td>
	        <input id="enddate" name="enddate" value="<%= Left(mdchoicerotate.FItemList(0).Fenddate,10) %>" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
	        <input type="text" name="enddatetime" size="2" maxlength="2" value="<%= ChkIIF(mdchoicerotate.FItemList(0).Fenddate="","23",Format00(2,Hour(mdchoicerotate.FItemList(0).Fenddate))) %>">(�� 00~23)
	        <input type="text" name="dummy1" value="59:59" size="6" readonly class="text_ro" />
		    <script type="text/javascript">
			var CAL_End = new Calendar({
				inputField : "enddate",
				trigger    : "enddate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				},
				bottomBar: true,
				dateFormat: "%Y-%m-%d"
			});
			</script>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">��뿩��</td>
		<td>
			<label><input type="radio" name="isusing" value="Y" <%=chkIIF(mdchoicerotate.FItemList(0).FIsUsing="Y" or mdchoicerotate.FItemList(0).FIsUsing="M" ,"checked","")%> >�����</label>
			<!--<label><input type="radio" name="isusing" value="M" <%=chkIIF(mdchoicerotate.FItemList(0).FIsUsing="M","checked","")%> >PC��+����� ���</label>-->
			<label><input type="radio" name="isusing" value="N" <%=chkIIF(mdchoicerotate.FItemList(0).FIsUsing="N","checked","")%> >������</label>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">�����</td>
		<td><% = mdchoicerotate.FItemList(0).FRegdate  %> (<% = mdchoicerotate.FItemList(0).Fregname  %>) </td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">�����۾���</td>
		<td><% = mdchoicerotate.FItemList(0).Fworkername  %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" align="center">
			<input type="button" value="����" onClick="SubmitForm()">
		</td>
	</tr>
	</table>
<%
		set mdchoicerotate = Nothing
	else
%>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">�̹���</td>
		<td>
			<input type="file" name="photoimg" value="" size="32" maxlength="32" class="file"><br />
			<font color="red">(119px �� 135px GIF Ȥ�� JPG �̹��� / �� 2013������: ��� �̹��� ������ ��ǰ�̹��� ���)</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">���ü���</td>
		<td><input type="text" name="disporder" value="99" size="2" class="text">
			<font color="red">(2�ڸ� ����)</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">��ǰ�ڵ�</td>
		<td><input type="text" name="linkitemid" value="" size="6" class="text"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">Text����</td>
		<td><textarea name="textinfo" class="textarea" style="width:90%; height:42px;"></textarea></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">link����</td>
		<td><input type="text" name="linkinfo" size="70"  class="text">
			<br>
			<font color="red">(����η� �Է��ϼ��� /shopping/category_prd.asp?itemid=72367)</font>
			<br><font color="red">(��ũ���� ���� ������ ��� ��ǰ�ڵ带 ������� �ڵ����� ��ũ���� ��ü�˴ϴ�.)</font>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">������ ǥ��</td>
		<td>
			<label><input type="radio" name="lowestPrice" value="Y">���</label>
			<label><input type="radio" name="lowestPrice" value="N" checked >������</label>		
		</td>
	</tr>	

	<tr bgcolor="#FFFFFF">
	    <td width="150" bgcolor="#DDDDFF">�ݿ�������</td>
	    <td>
	        <input id="startdate" name="startdate" value="" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
	        <input type="text" name="startdatetime" size="2" maxlength="2" value="00" />(�� 00~23)
	        <input type="text" name="dummy0" value="00:00" size="6" readonly class="text_ro" />
		    <script type="text/javascript">
			var CAL_Start = new Calendar({
				inputField : "startdate",
				trigger    : "startdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				},
				bottomBar: true,
				dateFormat: "%Y-%m-%d"
			});
			</script>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td width="150" bgcolor="#DDDDFF">�ݿ�������</td>
	    <td>
	        <input id="enddate" name="enddate" value="" class="text" size="10" maxlength="10" />
	        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
	        <input type="text" name="enddatetime" size="2" maxlength="2" value="23">(�� 00~23)
	        <input type="text" name="dummy1" value="59:59" size="6" readonly class="text_ro" />
		    <script type="text/javascript">
			var CAL_End = new Calendar({
				inputField : "enddate",
				trigger    : "enddate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				},
				bottomBar: true,
				dateFormat: "%Y-%m-%d"
			});
			</script>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="100" align="center" bgcolor="<%= adminColor("gray") %>">��뿩��</td>
		<td>
			<label><input type="radio" name="isusing" value="Y" checked>�����</label>
			<!--<label><input type="radio" name="isusing" value="M" checked >PC��+����� ���</label>-->
			<label><input type="radio" name="isusing" value="N">������</label>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2" align="center">
			<input type="button" value="����" onClick="SubmitForm()">
		</td>
	</tr>
	</table>
<%
	end if
%>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->