<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/breakdown/breakdownCls.asp"-->

<%
	Dim cBreakview, vCode, vGubun, vReqDIdx, vReqEquipment, vWorkType, vWorkTarget, vReqComment, vReqCapImage1, userid, vRequserid, vWorkPartSn
	vReqDIdx 	= requestCheckVar(Request("reqdidx"),10)
	vGubun 		= CHKIIF(vReqDIdx<>"","U","I")
	userid = session("ssBctId")

	If vGubun = "U" Then
		Set cBreakview = New CBreakdown
		 	cBreakview.FReqDIdx = vReqDIdx
			cBreakview.fnGetBreakdownView

			vReqEquipment 		= cBreakview.FReqEquipment
			vWorkPartSn 		= cBreakview.FWorkPartSn
			vWorkType			= cBreakview.FWorkType
			vWorkTarget			= Replace(cBreakview.FWorkTarget,"_break","")
			vReqComment			= cBreakview.FReqComment
			vReqCapImage1		= cBreakview.FReqCapImage1
			vRequserid			= cBreakview.FRequserid
		Set cBreakview = Nothing

		if userid <> vRequserid then
			Response.Write "<script>alert('�߸��� ���� �Դϴ�.');location.href='/admin/breakdown/?menupos="&request("menupos")&"';</script>"
			dbget.close()
			Response.End
		end if
	End If
%>

<script language="javascript">
function image_view(src){
	var image_view = window.open('/admin/culturestation/image_view.asp?image='+src,'image_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	image_view.focus();
}

//''�̹��� ����
function jsImgInput(divnm,iptNm,vPath,Fsize,Fwidth,fheight,thumb){

	//window.open('img_input.asp','imginput','width=350,height=300,menubar=no,toolbar=no,scrollbars=no,status=yes,resizable=yes,location=no');
	document.imginputfrm.divName.value=divnm;
	document.imginputfrm.inputname.value=iptNm;
	document.imginputfrm.ImagePath.value = vPath;
	document.imginputfrm.maxFileSize.value = Fsize;
	document.imginputfrm.maxFileWidth.value = Fwidth;
	document.imginputfrm.maxFileheight.value = fheight;
	document.imginputfrm.makeThumbYn.value = thumb;
	document.imginputfrm.orgImgName.value = eval("document.getElementById('"+iptNm+"')").value;
	document.imginputfrm.target='imginput';
	document.imginputfrm.submit();
}

function delimage(gubun)
{
	var aa = eval("document.frm."+gubun+"");
	aa.value = "";
	frm.mode.value = "edit";
	frm.isimgdel.value = "o";
	frm.submit();
}

document.domain = "10x10.co.kr";


function hideFrame() {
	document.iframeDB1.location.href = "about:blank";
}

function changeWorkType()
{
	var frm = document.frm;

	if (frm.work_part_sn.value === "9") {
		if ((frm.req_comment.value === "") && (frm.work_type.value === "1") && (frm.work_target.value === "1")) {
			//frm.req_comment.value = "\n- �ֹ���ȣ :\n\n- �����ȣ :\n\n- ���ó�� ����� : ex > ��ۻ�� / CJ ������� ������ ��..\n\n- ��ǰ����(�ǰ�����) : �������� ��  (������ ���� ���� �� �����ݾ� �̸� ��ۺ� �����Աݾ�)\n\n- ó������ : ex > �̵��� ��ǰ�н�, CJ������ ���� ���ó�� ���� ��..\n\n";
			frm.req_comment.value = "\n- �ֹ���ȣ :\n\n- ���� :\n\n- �����ȣ :\n\n- ����û�� : ex > ��ۻ�� / �ù�� ������ ��..\n\n- ó������ :\n\n";
		} else if ((frm.req_comment.value === "") && (frm.work_type.value === "4") && (frm.work_target.value === "1")) {
			frm.req_comment.value = "\n- �ֹ���ȣ :\n\n- �ֹ��ڵ� :\n\n- ���Ի��� :\n\n- ��� : ��밡�� or ���Ұ�(������) :\n\n"
		} else if ((frm.req_comment.value === "") && (frm.work_type.value === "1") && (frm.work_target.value === "2")) {
			frm.req_comment.value = "\n- �ֹ���ȣ :\n\n- �����ȣ :\n\n- �ֹ��ڵ�(OJ�ڵ�) :\n\n- �󼼻��� :\n\n"
		} else if ((frm.req_comment.value === "") && (frm.work_type.value !== "") && (frm.work_target.value !== "")) {
			frm.req_comment.value = "\n- �ֹ���ȣ :\n\n- ��ǰ�ڵ� :\n\n- ���� (�ڼ��� �����ּ���.) :\n\n"
		}
	}

	if (frm.work_part_sn.value != "30") {
		hideFrame();
		return;
	}

	if (frm.work_type.value === "" || frm.work_target.value === "") {
		hideFrame();
		return;
	}

	if ((frm.work_type.value == "1" || frm.work_type.value == "2") && (frm.work_target.value == "etc")) {
		alert("��Ÿ������ ���ó������ �ش��մϴ�.\n�ٽ� �����ϼ���.");
		frm.work_target.options[0].selected = true;
		frm.work_target.focus();
		return;
	}

	frm.req_equipment.value = "";

	document.getElementById("iframeDB1").height = "100%";
	document.iframeDB1.location.href = "iframe_selectbox.asp?work_type="+frm.work_type.value+"&work_target="+ frm.work_target.value +"";
}

function checkform(f)
{
	if (f.work_part_sn.value === "") {
		alert("�۾� �μ��� �����ϼ���.");
		f.work_part_sn.focus();
		return false;
	}

	if (f.work_part_sn.value !== "10" && f.work_part_sn.value !== "30" && f.work_part_sn.value !== "9") {
		alert("�ý����� ���� - �۾� �μ��� ���ȹ��/CS/���� ���� ���� �����մϴ�.");
		f.work_part_sn.focus();
		return false;
	}

	if(f.work_type.value == "")
	{
		alert("�۾� ��û ������ �����ϼ���.");
		f.work_type.focus();
		return false;
	}

	if(f.work_target.value == "")
	{
		alert("�۾� ��û ���л� �����ϼ���.");
		f.work_target.focus();
		return false;
	}

	if (f.work_part_sn.value === "30") {
		if(f.work_type.value == "3") {
			if(f.req_equipment.value == "") {
				alert("��ָ���Ʈ�� �����ϼ���.");
				return false;
			}

			f.req_comment.value = f.req_equipment_name.value + String.fromCharCode(13) + f.req_comment.value;
		} else {
			if(f.req_equipment.value == "") {
				alert("�ش� ��񸮽�Ʈ�� �����ϼ���.");
				return false;
			}
		}
	}
}
</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<form name="frm" action="breakdown_req_proc.asp" method="post" onSubmit="return checkform(this);" style="margin:0px;">
		<input type="hidden" name="gb" value="<%=vGubun%>">
		<input type="hidden" name="menupos" value="<%=request("menupos")%>">
		<input type="hidden" name="reqdidx" value="<%=vReqDIdx%>">
		<input type="hidden" name="req_equipment" value="<%=vReqEquipment%>">
		<input type="hidden" name="req_equipment_name" value="">
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<COLGROUP>
			<COL width="100" />
			<COL width="*" />
		</COLGROUP>
		<tr height="30">
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�۾� �μ�</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<%= printPartOption("work_part_sn", vWorkPartSn) %>
				* ���ȹ, CS, ���� �μ��� ���ð���
			</td>
		</tr>
		<tr height="30">
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�۾� ����</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<!-- #include virtual="/admin/breakdown/workgubunselectbox.asp"-->
			</td>
		</tr>
		<tr height="30">
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">ĸ���̹���</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="button" value="�̹��� �ֱ�" onclick="jsImgInput('divcapimg1','req_capimage1','cap1','3000','2000','2000','false');"  class="button" style="width:80px; height:23px;" />
				<input type="hidden" name="req_capimage1" id="req_capimage1" value="<%= vReqCapImage1 %>" />
				<div align="right" id="divcapimg1">
					<% if vReqCapImage1 <> "" then %>
					<a href="javascript:image_view('<%=webImgUrl%>/breakdown<%= vReqCapImage1 %>');" onfocus="this.blur()">
						<img src="<%=webImgUrl%>/breakdown<%= vReqCapImage1 %>" width="25" height="25"  border=0></a>
						<% end if %>
				</div>
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�ڸ�Ʈ</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<textarea name="req_comment" class="textarea" style="width:100%; height:350px;"><%=vReqComment%></textarea>
			</td>
		</tr>
		<tr height="30">
			<td colspan="2" bgcolor="#FFFFFF" style="padding: 0 5px 5px 5px;">
				<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
					<tr>
						<td colspan="2">
							<% If vGubun = "I" Then %>
							<iframe src="about:blank" name="iframeDB1" id="iframeDB1" width="100%" height="0" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
							<% ElseIf vGubun = "U" Then %>
							<iframe src="iframe_selectbox.asp?work_type=<%=vWorkType%>&work_target=<%=vWorkTarget%>&req_equipment=<%=vReqEquipment%>" name="iframeDB1" id="iframeDB1" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
							<% End If %>
						</td>
					</tr>
					<tr>
						<td><input type="button" class="button" value="����Ʈ��" onClick="self.location.href='index.asp?menupos=<%=request("menupos")%>';" style="width:80px; height:23px;"></td>
						<td align="right"><input type="submit" class="button" value=" �� �� " style="width:80px; height:23px;"></td>
					</tr>
				</table>
			</td>
		</tr>
		</table>
		</form>
	</td>
</tr>
</table>
<form name="imginputfrm" method="post" action="img_input.asp" style="margin:0px;">
	<input type="hidden" name="YearUse" value="<%=year(now)%>">
	<input type="hidden" name="divName" value="">
	<input type="hidden" name="orgImgName" value="">
	<input type="hidden" name="inputname" value="">
	<input type="hidden" name="ImagePath" value="">
	<input type="hidden" name="maxFileSize" value="">
	<input type="hidden" name="maxFileWidth" value="">
	<input type="hidden" name="maxFileheight" value="">
	<input type="hidden" name="makeThumbYn" value="">
</form>
<script type="text/javascript">

function getOnLoad(){
	var obj = document.frm.work_part_sn;

	// /cscenter/memo/mmgubunselectbox.asp ����
	startRequest('work_type', '<%= vWorkPartSn %>', '<%= vWorkType %>','<%= vWorkTarget %>');
	obj.onchange = function() {
		startRequest('work_type', obj.value, '','');
	};
}

window.onload = getOnLoad;

</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
