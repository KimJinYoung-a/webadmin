<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.19 ������ ����
'			2013.08.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/lookbookCls.asp"-->
<%
Dim mode, olookbook
Dim idx, makerid, title, state, mainimg, isusing, sortNo, regdate, lastupdate, regadminid,lastadminid
dim comment
	mode	= request("mode")
	idx		= request("idx")
	makerid	= request("makerid")
	menupos	= request("menupos")
	
If idx = "" Then
	mode = "I"
Else
	mode = "U"
End If

SET olookbook = new clookbook
	olookbook.FrectIdx = idx
	olookbook.frectmakerid = ""
	
	if idx <> "" then
		olookbook.sblookbookmodify
	end if
	
	if olookbook.ftotalcount > 0 then
		idx = olookbook.FOneItem.Fidx
		makerid = olookbook.FOneItem.Fmakerid
		title = olookbook.FOneItem.Ftitle
		state = olookbook.FOneItem.Fstate
		mainimg = olookbook.FOneItem.Fmainimg
		isusing = olookbook.FOneItem.Fisusing
		sortNo = olookbook.FOneItem.FsortNo
		regdate = olookbook.FOneItem.Fregdate
		lastupdate = olookbook.FOneItem.Flastupdate
		regadminid = olookbook.FOneItem.Fregadminid
		lastadminid = olookbook.FOneItem.Flastadminid
		comment = olookbook.FOneItem.fcomment
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
	
	if(frm.title.value==''){
		alert('������ �Է��ϼ���.');
		frm.title.focus();
		return;
	}
	
	if(frm.isusing.value==''){
		alert('��뿩�θ� �����ϼ���.');
		frm.isusing.focus();
		return;
	}

	if (GetByteLength(frm.comment.value) > 512){
		alert("�ڸ�Ʈ�� ���ѱ��̸� �ʰ��Ͽ����ϴ�. 256�� ���� �ۼ� �����մϴ�.");
		frm.comment.focus();
		return;
	}
	
	if(frm.mainimg.value==""){
		alert('���� �̹����� ����ϼ���');
		frm.mainimg.focus();
		return;
	}
	
	if(confirm('[������]�����Ͻðڽ��ϱ�?')){
		frm.mode.value=mode;
		frm.submit();
	}
}

function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/brand/lookbook/pop_lookbook_uploadimg.asp?sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}

//���º���
function chstate(state){
	if(confirm("���¸� ���� �Ͻðڽ��ϱ�?")){
		frmchstate.mode.value='chstate';
		frmchstate.state.value=state;
		frmchstate.submit();
	}
}

</script>

<!-- #include virtual="/admin/brand/inc_streetHead.asp"-->

<img src="/images/icon_arrow_link.gif"> <b><b>LOOKBOOK ���</b></b>

<table border="0" cellpadding="0" cellspacing="0" class="a" width="100%">
<form name="frmchstate" method="post" action="/admin/brand/lookbook/lookbook_process.asp" style="margin:0px;">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="state">
	<input type="hidden" name="idx" value="<%=idx%>">
	<input type="hidden" name="mode">
</form>
<form name="frm" method="post" action="/admin/brand/lookbook/lookbook_process.asp" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="mainimg" value="<%=mainimg%>">
<input type="hidden" name="statcd" value="">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" width="100%">
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>" width=200>��ȣ</td>
			<td bgcolor="#FFFFFF">
				<%=idx%>
				<input type="hidden" name="idx" value="<%=idx%>">
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
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF"><input type="text" size="70" maxlength=50 name="title" value="<%= title %>"></td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF">
				<% if mode="U" then %>
					<%' drawlookbookstats "state" , state , " onchange='gosubmit("""");'" %>
					<%= lookbookstatsname(state) %>
					<input type="hidden" name="state" value="<%=state%>">
				<% else %>
					�����
				<% end if %>
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">���</td>
			<td bgcolor="#FFFFFF" >
				<% drawSelectBoxUsingYN "isusing", isusing %>
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">������̹���<br>(�ʼ��Է� : 259 x 360)</td>
			<td bgcolor="#FFFFFF">
				<input type="button" name="btnBan" value="�̹������" onClick="jsSetImg('lookbook','<%= mainimg %>','mainimg','spanban')" class="button">
	   			<div id="spanban" style="padding: 5 5 5 5">
	   				<% IF mainimg <> "" THEN %>
	   					<img src="<%=mainimg%>" border="0" width="259" height="360">
	   					<a href="javascript:jsDelImg('mainimg','spanban');"><img src="/images/icon_delete2.gif" border="0"></a>
	   				<%END IF%>
	   			</div>
			</td>
		</tr>
		<tr>
			<td align="center"  bgcolor="<%= adminColor("tabletop") %>">
				�۾��ڸ�Ʈ
				<Br>(�ݷ������� ��ü�� �ϰ������)
			</td>
			<td bgcolor="#FFFFFF" >
				<textarea name="comment" rows="15" cols="69"><%= comment %></textarea>
			</td>
		</tr>
		
		<% If mode = "U" Then %>
			<tr>
				<td align="center"  bgcolor="<%= adminColor("tabletop") %>">LOOKBOOK�̹���</td>
				<td bgcolor="#FFFFFF">
					<iframe id="iframG" frameborder="0" width="100%" src="/admin/brand/lookbook/iframe_lookbook_detail.asp?idx=<%=idx%>" height=300></iframe>
				</td>
			</tr>
		<% else %>
			<tr>
				<td align="center"  bgcolor="<%= adminColor("tabletop") %>">LOOKBOOK�̹���</td>
				<td bgcolor="#FFFFFF">
					�űԵ�� �Ϸ��� LOOKBOOK�̹����� �Է� �ϽǼ� �ֽ��ϴ�.
				</td>
			</tr>
		<% End If %>
		
		<tr align="center">
			<td bgcolor="#FFFFFF" colspan=2>
				<% If mode = "U" Then %>
					<input type="button" value="����" class="button" onclick="form_check('U');">
				<% elseif mode = "I" Then %>
					<input type="button" value="�űԵ��" class="button" onclick="form_check('I');">
				<% End If %>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<% If mode = "U" Then %>
					<%
					'/�ݷ�(������û)�ϰ��
					If state = "1" Then
					%>
						<input type="button" value="���ο�û" class="button" onclick="chstate('3');">
					<% end if %>
	
					<%
					'/������ϰ��
					If state = "2" Then
					%>
						<input type="button" value="���ο�û" class="button" onclick="chstate('3');">
					<% end if %>
					<%
					'/���ο�û�ϰ��
					If state = "3" Then
					%>
						<input type="button" value="����" class="button" onclick="chstate('7');">
						<input type="button" value="�ݷ�(������û)" class="button" onclick="chstate('1');">
						<input type="button" value="�����ݷ�(������)" class="button" onclick="chstate('9');">
					<% end if %>
					<%
					'/�����ϰ��
					If state = "7" Then
					%>
						<input type="button" value="�ݷ�(������û)" class="button" onclick="chstate('1');">
						<input type="button" value="�����ݷ�(������)" class="button" onclick="chstate('9');">
					<% end if %>
						
					<%
					'/�����ݷ�(������)�ϰ��
					If state = "9" Then
					%>
						<input type="button" value="���ο�û" class="button" onclick="chstate('3');">
					<% end if %>
					
				<% End If %>
			</td>
		</tr>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->