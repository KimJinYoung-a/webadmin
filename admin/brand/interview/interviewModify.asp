<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/interviewCls.asp"-->
<%
Dim mode, ointerview
Dim mainidx, makerid, title, mainimg, detailimg, startdate, isusing, regdate, lastupdate, regadminid,lastadminid
dim comment, detailimglink
	mode	= request("mode")
	mainidx		= request("idx")
	makerid	= request("makerid")
	menupos	= request("menupos")
	
If mainidx = "" Then
	mode = "I"
Else
	mode = "U"
End If

SET ointerview = new cinterview
	ointerview.FrectIdx = mainidx
	ointerview.frectmakerid = ""
	
	if mainidx <> "" then
		ointerview.finterview_modify
	end if
	
	if ointerview.ftotalcount > 0 then
		mainidx = ointerview.FOneItem.fmainidx
		makerid = ointerview.FOneItem.Fmakerid
		startdate = ointerview.FOneItem.Fstartdate
		title = ointerview.FOneItem.Ftitle
		mainimg = ointerview.FOneItem.Fmainimg
		detailimg = ointerview.FOneItem.fdetailimg
		isusing = ointerview.FOneItem.Fisusing
		regdate = ointerview.FOneItem.Fregdate
		lastupdate = ointerview.FOneItem.Flastupdate
		regadminid = ointerview.FOneItem.Fregadminid
		lastadminid = ointerview.FOneItem.Flastadminid
		comment = ointerview.FOneItem.fcomment
		detailimglink = ointerview.FOneItem.fdetailimglink
	end if
%>

<script language="javascript">

function form_check(mode){
	var frm = document.frm;

	if(frm.makerid.value==''){
		alert('�귣�带 �����ϼ���.');
		frm.makerid.focus();
		return;
	}
	if(frm.startdate.value==''){
		alert('�������� �Է��ϼ���.');
		frm.startdate.focus();
		return;
	}
	if(frm.title.value==''){
		alert('������ �Է��ϼ���.');
		frm.title.focus();
		return;
	}

	if (GetByteLength(frm.comment.value) > 512){
		alert("�ڸ�Ʈ�� ���ѱ��̸� �ʰ��Ͽ����ϴ�. 256�� ���� �ۼ� �����մϴ�.");
		frm.comment.focus();
		return;
	}
			
	if(frm.isusing.value==''){
		alert('��뿩�θ� �����ϼ���.');
		frm.isusing.focus();
		return;
	}
	if(frm.mainimg.value==""){
		alert('���� �̹����� ����ϼ���');
		frm.mainimg.focus();
		return;
	}
	if(frm.detailimg.value==""){
		alert('�� �̹����� ����ϼ���');
		frm.detailimg.focus();
		return;
	}
	
	if(confirm('�����Ͻðڽ��ϱ�?')){
		frm.mode.value=mode;
		frm.submit();
	}
}

function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/brand/interview/pop_interview_uploadimg.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

</script>

<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b><b>INTERVIEW ���</b></b>

<table border="0" cellpadding="0" cellspacing="0" class="a" width="100%">
<form name="frm" method="post" action="/admin/brand/INTERVIEW/INTERVIEW_process.asp" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="mainimg" value="<%=mainimg%>">
<input type="hidden" name="detailimg" value="<%=detailimg%>">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" width="100%">
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>" width=200>��ȣ</td>
			<td bgcolor="#FFFFFF">
				<%=mainidx%>
				<input type="hidden" name="idx" value="<%=mainidx%>">
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�귣��</td>
			<td bgcolor="#FFFFFF">
				<% if mode = "U" then %>
					<%= makerid %>
					<input type="hidden" name="makerid" value="<%= makerid %>">
				<% else %>
					<% drawSelectBoxDesignerwithName "makerid",makerid %>
				<% end if %>
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">������</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="startdate" size=10 value="<%= left(startdate,10) %>" class="text">			
				<a href="javascript:calendarOpen3(frm.startdate,'������',frm.startdate.value)">
				<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
			</td>
		</tr>	
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF"><input type="text" size="70" maxlength=50 name="title" value="<%= title %>"></td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">
				�ڸ�Ʈ
			</td>
			<td bgcolor="#FFFFFF" >
				<textarea name="comment" rows="5" cols="69"><%= comment %></textarea>
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">���</td>
			<td bgcolor="#FFFFFF" >
				<% drawSelectBoxUsingYN "isusing", isusing %>
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">�����̹���</td>
			<td bgcolor="#FFFFFF">
				<input type="button" name="btnBan" value="�̹������" onClick="jsSetImg('interview','<%= mainimg %>','mainimg','spanban')" class="button">
	   			<div id="spanban" style="padding: 5 5 5 5">
	   				<% IF mainimg <> "" THEN %>
	   					<img src="<%=mainimg%>" border="0" width="259" height="360">
	   					<a href="javascript:jsDelImg('mainimg','spanban');"><img src="/images/icon_delete2.gif" border="0"></a>
	   				<%END IF%>
	   			</div>
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">���̹���</td>
			<td bgcolor="#FFFFFF">
				<input type="button" name="btnBan" value="�̹������" onClick="jsSetImg('interview','<%= detailimg %>','detailimg','spandetailimgban')" class="button">
	   			<div id="spandetailimgban" style="padding: 5 5 5 5">
	   				<% IF detailimg <> "" THEN %>
	   					<img src="<%=detailimg%>" border="0" width="259" height="360">
	   					<a href="javascript:jsDelImg('detailimg','spandetailimgban');"><img src="/images/icon_delete2.gif" border="0"></a>
	   				<%END IF%>
	   			</div>
			</td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF" colspan=2>
				<b>���̹��� �̹����� & ��ũ �ڵ�</b> <font color="red"> map�̸� ���� ��������!</font>
				<textarea rows="15" name="detailimglink" style="width:100%" class="textarea"><%=chkIIF(mode = "U",detailimglink,"<map name=""interviewmap1""></map>") %></textarea>
			</td>
		</tr>		
		<tr align="center">
			<td bgcolor="#FFFFFF" colspan=2>
				<% If mode = "U" Then %>
					<input type="button" value="����" class="button" onclick="form_check('U');">
				<% elseif mode = "I" Then %>
					<input type="button" value="�űԵ��" class="button" onclick="form_check('I');">
				<% End If %>
			</td>
		</tr>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->