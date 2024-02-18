<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

<%
dim menupos ,oWeekly ,catename
dim idx,title,subcopy,state,weekly_banner_img,weekly_title_img,startdate,enddate,isusing,regdate,comment
dim lastadminid,cd1,opendate,closedate,partMDid,partWDid
	idx = request("idx")
	menupos = request("menupos")

'//�̺�Ʈ����
set oWeekly = new ClsStyleLife
	oWeekly.frectidx = idx
	
	if idx <> "" then
		oWeekly.fnGetWeekly_item()
		
		if oWeekly.ftotalcount > 0 then			
			title = oWeekly.foneitem.ftitle
			cd1 = oWeekly.foneitem.fcd1
			state = oWeekly.foneitem.fstate
			weekly_banner_img = oWeekly.foneitem.fbanner_img
			weekly_title_img = oWeekly.foneitem.ftitle_img
			startdate = left(oWeekly.foneitem.fstartdate,10)
			regdate = oWeekly.foneitem.fregdate
			comment = oWeekly.foneitem.fcomment
			lastadminid = oWeekly.foneitem.flastadminid
			partMDid = oWeekly.foneitem.fpartMDid
			partWDid = oWeekly.foneitem.fpartWDid
		end if	
	end if
set oWeekly = nothing
	
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
	
	function jsContentImg(idx)
	{
		if(idx == "")
		{
			alert("No.�� ��ȣ�� �־�� �մϴ�.");
			return;
		}
		else
		{
			var cImg = window.open('stylelife_weekly_CImg.asp?idx='+idx+'','cImg','width=370,height=600,scrollbars=yes');
			cImg.focus();
		}
	}

</script>

<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="/admin/stylepick/stylelife_weekly_process.asp" method="post">
<input type="hidden" name="mode" value="eventedit">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="weekly_banner_img" value="<%=weekly_banner_img%>">
<input type="hidden" name="weekly_title_img" value="<%=weekly_title_img%>">
<input type="hidden" name="opendate" value="<%=opendate%>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">No.</td>
	<td bgcolor="#FFFFFF"><%= idx %><input type="hidden" name="idx" value="<%=idx%>"></td>
</tr>	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
	<td bgcolor="#FFFFFF"><input type="text" size=64 maxlength=64 name="title" value="<%=title%>"></td>
</tr>
	
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">��Ÿ��</td>
	<td bgcolor="#FFFFFF">
		<select name="cd1">
			<option value="">-��Ÿ��-</option>
			<option value="010" <%=CHKIIF(cd1="010","selected","")%>>Ŭ����</option>
			<option value="020" <%=CHKIIF(cd1="020","selected","")%>>ťƮ</option>
			<option value="040" <%=CHKIIF(cd1="040","selected","")%>>���</option>
			<option value="050" <%=CHKIIF(cd1="050","selected","")%>>���߷�</option>
			<option value="060" <%=CHKIIF(cd1="060","selected","")%>>������Ż</option>
			<option value="070" <%=CHKIIF(cd1="070","selected","")%>>��</option>
			<option value="080" <%=CHKIIF(cd1="080","selected","")%>>�θ�ƽ</option>
			<option value="090" <%=CHKIIF(cd1="090","selected","")%>>��Ƽ��</option>
		</select>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
	<td bgcolor="#FFFFFF">
		<% Draweventstate2 "state" , state ,"" %> �� ������ �ؼ� �����Ͽ��� ������ =< ���� �̾�߸� ������ �˴ϴ�.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">������</td>
	<td bgcolor="#FFFFFF">
   		<%IF state = "9" THEN%>
   			<%=startdate%><input type="hidden" name="startdate" size=10 maxlength=10 value="<%=startdate%>">
   		<%ELSE%>
   			<input type="text" name="startdate" size=10 maxlength=10 value="<%=startdate%>" onClick="jsPopCal('startdate');"  style="cursor:hand;">
   		<%END IF%>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partmdid",partmdid,"","23" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">���WD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�۾����޻���</td>
	<td bgcolor="#FFFFFF">
		<textarea rows=10 cols=100 name="comment"><%=comment%></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����̹���</td>
	<td bgcolor="#FFFFFF">		
		<input type="button" name="btnBanImg" value="�̹������" onClick="jsSetImg('<%=weekly_banner_img%>','weekly_banner_img','weekly_banner_imgdiv')" class="button">
		<div id="weekly_banner_imgdiv" style="padding: 5 5 5 5">
			<%IF weekly_banner_img <> "" THEN %>			
				<img src="<%=weekly_banner_img%>" border="0" height=100 onclick="jsImgView('<%=weekly_banner_img%>');" alt="�����ø� Ȯ�� �˴ϴ�">
				<a href="javascript:jsDelImg('weekly_banner_img','weekly_banner_imgdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">Ÿ��Ʋ�̹���</td>
	<td bgcolor="#FFFFFF">		
		<input type="button" name="btnTitImg" value="�̹������" onClick="jsSetImg('<%=weekly_title_img%>','weekly_title_img','weekly_title_imgdiv')" class="button">
		<div id="weekly_title_imgdiv" style="padding: 5 5 5 5">
			<%IF weekly_title_img <> "" THEN %>			
				<img src="<%=weekly_title_img%>" border="0" height=50 onclick="jsImgView('<%=weekly_title_img%>');" alt="�����ø� Ȯ�� �˴ϴ�">
				<a href="javascript:jsDelImg('weekly_title_img','weekly_title_imgdiv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����̹���</td>
	<td bgcolor="#FFFFFF">
		<% If idx = "" Then %>
			�����̹����� ����Ϸ��� �ݵ�� No. �� �ʿ��մϴ�. �׷��Ƿ� �� ������ ���� �� ����ϼ���.
		<% Else %>
			<input type="button" name="btnTitImg" value="�̹������" onClick="jsContentImg('<%=idx%>')" class="button">
		<% End If %>
		<br>�� �����̹����� �������� ���� ����ǹǷ� �˾����������� �����մϴ�.
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="button" onclick="jsEvtSubmit();" class="button" value="����">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="button" value="�ݱ�" class="button" onClick="window.close()">
	</td>
</tr>
</form>
</table>


<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
