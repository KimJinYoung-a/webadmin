<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : T-Episode
' Hieditor : ����ȭ ����
'			 2022.07.07 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/play/playCls.asp" -->
<%
Dim idx, viewtitle, subtitle, isusing, PPimg, regdate, playcate, menupos, style_html_m
Dim oPick , oground
	idx = requestCheckVar(getNumeric(request("idx")),10)
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
playcate = 7 't-episode

Set oPick = new CPlayContents
	oPick.FRectIdx = idx

If idx <> "" Then
	oPick.GetPhotoPickOne()
	If oPick.FResultCount > 0 Then
		idx			= oPick.FOneItem.Fidx
		viewtitle	= oPick.FOneItem.FViewtitle
		subtitle	= oPick.FOneItem.FSubtitle
		isusing		= oPick.FOneItem.FIsusing
		PPimg		= oPick.FOneItem.FPPimg
		regdate		= oPick.FOneItem.FRegdate
		style_html_m =  oPick.FOneItem.fstyle_html_m
	End If
End If
set oPick = Nothing

If isusing = "" Then isusing = "Y"

%>
<script type="text/javascript">
<!--
	//�̹��� Ȯ��ȭ�� ��â���� �����ֱ�
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
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
		winImg = window.open('/admin/sitemaster/play/lib/pop_theme_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsTagview(gidx , idx){	
		var poptag;
		poptag = window.open('/admin/sitemaster/play/lib/pop_tagReg.asp?idx='+gidx+'&playcate='+<%=playcate%>,'poptag','width=500,height=400,scrollbars=yes,resizable=yes');
		poptag.focus();
	}

	function subcheck(){
		var frm=document.inputfrm;

		if (!frm.viewtitle.value){
			alert('viewtitle�� ������ּ���');
			frm.viewtitle.focus();
			return;
		}

		if (!frm.subtitle.value){
			alert('subtitle�� ������ּ���');
			frm.subtitle.focus();
			return;
		}

		frm.submit();
	}

	function jsManagePlayImage(){
		var playManageDir = window.open('http://<%=CHKIIF(application("Svr_Info")="Dev","test","")%>upload.10x10.co.kr/linkweb/play/playManageDir.asp?folder=tepisode&idx=<%=idx%>','playManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
		playManageDir.focus();
	}
//-->
</script>

<form name="inputfrm" method="post" action="photopickProc.asp">
<input type="hidden" name="idx" value="<%= idx %>"/>
<input type="hidden" name="photopickimg" value="<%=PPimg%>"/>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>T-episode���/����</b></font><br/><br/>
	</td>
</tr>
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">idx</td>
	<td bgcolor="#FFFFFF">
		<b><%=idx%></b>
	</td>
</tr>
<% End If %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="viewtitle" value="<%= ReplaceBracket(viewtitle) %>" size="50"/>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����ī��</td>
	<td bgcolor="#FFFFFF">
		<textarea name="subtitle" rows="8" cols="50"><%= ReplaceBracket(subtitle) %></textarea>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <%= chkiif(isusing = "Y", "checked", "")%> >Y
		<input type="radio" name="isusing" value="N" <%= chkiif(isusing = "N", "checked", "")%> >N
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnorgmg" value="�̹������" onClick="jsSetImg('<%=PPimg%>','photopickimg','photopickimgdiv')" class="button"/> **�̹��� ���α��� 566px�� �����ּ���
		<div id="photopickimgdiv" style="padding: 5 5 5 5">
			<%If PPimg <> "" THEN %>
				<img src="<%=PPimg%>" border="0" height=100 onclick="jsImgView('<%=PPimg%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('photopickimg','photopickimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				�̹��� �ּ� : <%=PPimg%>
			<%END IF%>
		</div>
	</td>
</tr>
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">Tag</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnviewImg" value="�±� ����" onClick="jsTagview('<%=idx%>', '')" class="button"/><br/><br/>
		���±װ����� �˾����� ���� �մϴ� ���� ��� ���ּ���.��
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">
		����ϼ��۾�����
		<% If idx <> "" Then %>
			<br /><br /><br /><br /><br /><input type="button" value="�̹�������" class="button" onClick="jsManagePlayImage('<%=idx%>');">
		<% End If %>		
	</td>
	<td bgcolor="#FFFFFF">
		<textarea name="style_html_m" style="width:100%; height:240px;"><%= ReplaceBracket(style_html_m) %></textarea>
	</td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
		<input type="button" value=" ��� " class="button" onclick="location.href='/admin/sitemaster/play/tepisode/?menupos=<%=menupos%>';"/>
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
