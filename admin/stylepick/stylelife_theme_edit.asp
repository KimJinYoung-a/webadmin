<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��Ÿ���� ����
' Hieditor : 2011.04.06 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

<%
dim menupos ,oTheme ,catename
dim idx,title,subcopy,state,banner_img,title_img,startdate,enddate,isusing,regdate,comment, sortno
dim lastadminid,cd1,opendate,closedate,partMDid,partWDid
	idx = request("idx")
	menupos = request("menupos")

'//�̺�Ʈ����
set oTheme = new ClsStyleLife
	oTheme.frectidx = idx
	
	if idx <> "" then
		oTheme.fnGetTheme_item()
		
		if oTheme.ftotalcount > 0 then			
			title = oTheme.foneitem.ftitle
			subcopy = oTheme.foneitem.fsubcopy
			state = oTheme.foneitem.fstate
			banner_img = oTheme.foneitem.fbanner_img
			title_img = oTheme.foneitem.ftitle_img
			startdate = left(oTheme.foneitem.fstartdate,10)
			enddate = left(oTheme.foneitem.fenddate,10)
			regdate = oTheme.foneitem.fregdate
			comment = oTheme.foneitem.fcomment
			lastadminid = oTheme.foneitem.flastadminid
			cd1 = oTheme.foneitem.fcd1
			opendate = oTheme.foneitem.fopendate
			closedate = oTheme.foneitem.fclosedate
			partMDid = oTheme.foneitem.fpartMDid
			partWDid = oTheme.foneitem.fpartWDid
			catename = oTheme.foneitem.fcatename
			sortno = oTheme.foneitem.fsortno
		end if	
	end if
set oTheme = nothing
	
if isusing = "" then isusing = "Y"
%>

<script language="javascript">

	//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	//�̹��� Ȯ��ȭ�� ��â���� �����ֱ�
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	function jsSetImg(sImg, sName, sSpan){	
		document.domain ="10x10.co.kr";
		
		var winImg;
		winImg = window.open('pop_theme_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	//����
	function jsEvtSubmit(){

		if(!frm.cd1.value){
			alert("��Ÿ���� �������ּ���");
			frm.cd1.focus();
			return;
		}

		if(!frm.title.value){
			alert("������ �Է����ּ���");
			frm.title.focus();
			return;
		}
		
		if(!frm.state.value){
			alert("���¸� �������ּ���");
			frm.state.focus();
			return;
		}
	
		if(!frm.startdate.value){
			alert("�������� �Է����ּ���");
			return;
		}

		if(!frm.partmdid.value){
			alert("��� MD�� �����ϼ���.");
			frm.partmdid.focus();
			return;
		}

		if(!frm.partwdid.value){
			alert("��� WD�� �����ϼ���.");
			frm.partwdid.focus();
			return;
		}

		frm.submit();
	}
	
	function TextCD1(g)
	{
		if(g == "0P0")
		{
			document.getElementById("txtcd1").innerHTML = "��Ÿ����<br>����";
		}
		else
		{
			document.getElementById("txtcd1").innerHTML = "Ÿ��Ʋ";
		}
	}
</script>

<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="/admin/stylepick/stylelife_theme_process.asp" method="post">
<input type="hidden" name="mode" value="eventedit">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="banner_img" value="<%=banner_img%>">
<input type="hidden" name="title_img" value="<%=title_img%>">
<input type="hidden" name="opendate" value="<%=opendate%>">
<input type="hidden" name="closedate" value="<%=closedate%>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ȹ����ȣ</td>
	<td bgcolor="#FFFFFF"><%= idx %><input type="hidden" name="idx" value="<%=idx%>"></td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">��Ÿ��</td>
	<td bgcolor="#FFFFFF">
		<select name="cd1" onChange="TextCD1(this.value);">
			<option value="">-��Ÿ��-</option>
			<option value="010" <%=CHKIIF(cd1="010","selected","")%>>Ŭ����</option>
			<option value="020" <%=CHKIIF(cd1="020","selected","")%>>ťƮ</option>
			<option value="040" <%=CHKIIF(cd1="040","selected","")%>>���</option>
			<option value="050" <%=CHKIIF(cd1="050","selected","")%>>���߷�</option>
			<option value="060" <%=CHKIIF(cd1="060","selected","")%>>������Ż</option>
			<option value="070" <%=CHKIIF(cd1="070","selected","")%>>��</option>
			<option value="080" <%=CHKIIF(cd1="080","selected","")%>>�θ�ƽ</option>
			<option value="090" <%=CHKIIF(cd1="090","selected","")%>>��Ƽ��</option>
			<option value="0P0" <%=CHKIIF(cd1="0P0","selected","")%>>��Ÿ����</option>
		</select>
	</td>
</tr>
	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
	<td bgcolor="#FFFFFF"><input type="text" size=64 maxlength=64 name="title" value="<%=title%>"></td>
</tr>
	
<!--
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����ī��</td>
	<td bgcolor="#FFFFFF"><input type="text" size=64 maxlength=64 name="subcopy" value="<%=subcopy%>"></td>
</tr>
//-->
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
	<td bgcolor="#FFFFFF">
		<% Draweventstate2 "state" , state ,"" %>
		&nbsp;&nbsp;&nbsp;�� ���� ���µǴ� ������ ���°� <font color="red"><b>����</b></font>, �Ⱓ�� <font color="red"><b>������ <= ������</b></font> �ν� �ΰ��� ��� ���� �� �͸� ���̰� �˴ϴ�.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�Ⱓ</td>
	<td bgcolor="#FFFFFF">
   		<%IF state = "9" THEN%>
   			������ : <%=startdate%><input type="hidden" name="startdate" size=10 maxlength=10 value="<%=startdate%>">
   		<%ELSE%>
   			������ : <input type="text" name="startdate" size=10 maxlength=10 value="<%=startdate%>" onClick="jsPopCal('startdate');"  style="cursor:hand;">
   		<%END IF%>
   		<%
		if opendate <> "1900-01-01" and opendate <> "" then response.write " ����ó���� : " & opendate
		if closedate <> "1900-01-01" and closedate <> "" then response.write " ����ó���� : " & closedate
		%>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">���MD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partmdid",partmdid,"","11,21" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">���WD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<% If idx <> "" Then %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
	<td bgcolor="#FFFFFF"><input type="text" size=7 maxlength=5 name="sortno" value="<%=sortno%>"></td>
</tr>
<% End If %>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�۾����޻���</td>
	<td bgcolor="#FFFFFF">
		<textarea rows=10 cols=100 name="comment"><%=comment%></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�⺻����̹���</td>
	<td bgcolor="#FFFFFF">		
		<input type="button" name="btnBanImg" value="�̹������" onClick="jsSetImg('<%=banner_img%>','banner_img','banner_imgdiv')" class="button">
		<div id="banner_imgdiv" style="padding: 5 5 5 5">
			<%IF banner_img <> "" THEN %>			
				<img src="<%=banner_img%>" border="0" width=100 height=100 onclick="jsImgView('<%=banner_img%>');" alt="�����ø� Ȯ�� �˴ϴ�">
				<a href="javascript:jsDelImg('banner_img','banner_imgdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center"><span id="txtcd1">Ÿ��Ʋ</span>�̹���</td>
	<td bgcolor="#FFFFFF">		
		<input type="button" name="btnTitImg" value="�̹������" onClick="jsSetImg('<%=title_img%>','title_img','title_imgdiv')" class="button">
		<div id="title_imgdiv" style="padding: 5 5 5 5">
			<%IF title_img <> "" THEN %>			
				<img src="<%=title_img%>" border="0" width=100 height=100 onclick="jsImgView('<%=title_img%>');" alt="�����ø� Ȯ�� �˴ϴ�">
				<a href="javascript:jsDelImg('title_img','title_imgdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="button" onclick="jsEvtSubmit();" class="button" value="����"></td>
</tr>	
</form>
</table>

<script>
<%
If cd1 = "0P0" Then
Response.Write "TextCD1('0P0');"
End If
%>
</script>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
