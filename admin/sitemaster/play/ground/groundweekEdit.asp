<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/play/playCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
	Dim gidx ,idx , listimg , state , reservationdate , viewtitle , viewtext , playcate
	Dim viewno , textimg , worktext , partmkid,partWDid
	dim playmainimg , beforeimg , afterimg , topbgimg , sideltimg , sidertimg  ,myplayimg
	Dim oPlay
	dim subBGColor , viewcontents , mainTopBGColor , mo_contents , mo_exec_check , exec_check , exec_filepath

	idx  = request("idx")
	gidx = request("gidx")
    playcate = 1 'ground

    '//db 1row
	set oPlay = new CPlayContents
		 oPlay.FRectIdx = idx
 		 oPlay.FRectgIdx = gidx

		if idx <> "" Then
			oPlay.GetRowGroundSub()

			if oPlay.FResultCount > 0 then
				listimg				= oPlay.FOneItem.Flistimg
				textimg				= oPlay.FOneItem.Ftextimg
				viewtitle			= oPlay.FOneItem.Fviewtitle
				reservationdate		= oPlay.FOneItem.Freservationdate
				state				= oPlay.FOneItem.Fstate
				viewno				= oPlay.FOneItem.Fviewno
				worktext			= oPlay.FOneItem.Fworktext
				partmkid			= oPlay.FOneItem.Fpartmkid
				partWDid			= oPlay.FOneItem.FpartWDid
				playmainimg			= oPlay.FOneItem.Fplaymainimg
				beforeimg			= oPlay.FOneItem.Fbeforeimg
				afterimg			= oPlay.FOneItem.Fafterimg
				topbgimg			= oPlay.FOneItem.Ftopbgimg
				sideltimg			= oPlay.FOneItem.Fsideltimg
				sidertimg			= oPlay.FOneItem.Fsidertimg
				subBGColor			= oPlay.FOneItem.FsubBGColor
				mainTopBGColor		= oPlay.FOneItem.FmainTopBGColor
				viewcontents		= oPlay.FOneItem.Fviewcontents
				myplayimg			= oPlay.FOneItem.Fmyplayimg
				mo_contents			= oPlay.FOneItem.Fmo_contents
				mo_exec_check		= oPlay.FOneItem.Fmo_exec_check
				exec_check			= oPlay.FOneItem.Fexec_check
				exec_filepath		= oPlay.FOneItem.Fexec_filepath
			end if
		end if
	set oPlay = Nothing

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/evol.colorpicker.min.js"></script>
<script type="text/javascript">
$(function(){
	//�÷���Ŀ
	//$("input[name='subBGColor']").colorpicker();
	//$("input[name='mainTopBGColor']").colorpicker();
});
</script>
<script type="text/javascript">
<!--
//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){
		var winCal;

		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

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

	function jsTagview(gidx,idx){
		var poptag;
		poptag = window.open('/admin/sitemaster/play/lib/pop_tagReg.asp?idx='+gidx+'&subidx='+idx+'&playcate='+<%=playcate%>,'poptag','width=1100,height=400,scrollbars=yes,resizable=yes');
		poptag.focus();
	}

	function jsSetItem(idx , styleno){
		var popitem;
		popitem = window.open('/admin/sitemaster/play/lib/pop_itemReg.asp?idx='+idx+'&number='+styleno,'popitem','width=500,height=400,scrollbars=yes,resizable=yes');
		popitem.focus();
	}


	function subcheck(){
		var frm=document.inputfrm;

		if (!frm.viewno.value){
			alert('No.�� ������ּ���');
			frm.viewno.focus();
			return;
		}

		if (!frm.viewtitle.value){
			alert('�������� ������ּ���');
			frm.viewtitle.focus();
			return;
		}

		if (!frm.worktext.value){
			alert('�۾������� ������ּ���');
			frm.worktext.focus();
			return;
		}

		if (!frm.reservationdate.value){
			alert('���¿������� ������ּ���');
			frm.reservationdate.focus();
			return;
		}

		if(!frm.state.value){
			alert("���¸� �������ּ���");
			frm.state.focus();
			return;
		}

		frm.submit();
	}

	function workerlist() //�����
	{
		var openWorker = null;
		var worker = inputfrm.selMId.value;
		openWorker = window.open('/admin/sitemaster/play/lib/PopWorkerList.asp?worker='+worker+'&team=22','openWorker','width=570,height=570,scrollbars=yes');
		openWorker.focus();
	}

	function jsManagePlayImage(){
		var playManageDir = window.open('http://<%=CHKIIF(application("Svr_Info")="Dev","test","")%>upload.10x10.co.kr/linkweb/play/playManageDir.asp?folder=ground&idx=<%=idx%>','playManageDir','width=1000,height=600,scrollbars=yes,resizable=yes');
		playManageDir.focus();
	}

	function execchk(v,d){
		if (d == "M")
		{
			if (v == "1")
			{
				document.getElementById("moc1").style.display = "none";
				document.getElementById("moc2").style.display = "block";
			}else{
				document.getElementById("moc1").style.display = "block";
				document.getElementById("moc2").style.display = "none";
			}
		}else{
			if (v == "1")
			{
				document.getElementById("wc").style.display = "none";
			}else{
				document.getElementById("wc").style.display = "block";
			}
		}
	}
//-->
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="/admin/sitemaster/play/ground/groundProc.asp">
<input type="hidden" name="gidx" value="<%= gidx %>">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="playmainimg" value="<%=playmainimg%>">
<input type="hidden" name="beforeimg" value="<%=beforeimg%>">
<input type="hidden" name="afterimg" value="<%=afterimg%>">
<input type="hidden" name="topbgimg" value="<%=topbgimg%>">
<input type="hidden" name="sideltimg" value="<%=sideltimg%>">
<input type="hidden" name="sidertimg" value="<%=sidertimg%>">
<input type="hidden" name="myplayimg" value="<%=myplayimg%>">
<input type="hidden" name="position" value="sub"/>
<tr height="30">
	<td colspan="4" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>Play&gt;&gt;ground �� ���/����</b></font><br/><br/>
		�� �̹��� �� �±� ������ �⺻ ������ �����˴ϴ� ��
	</td>
</tr>
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">idx</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<b><%=idx%></b>
	</td>
</tr>
<% End If %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">No.</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" class="text" name="viewno" value="<%=viewno%>" size="10"/>�� ���ڸ� �����ּ��� ��
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" class="text" name="viewtitle" value="<%=viewtitle%>" size="50"/>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% Draweventstate2 "state" , state ,"" %> �� ������ �ؼ� �����Ͽ��� ������ =< ���� �̾�߸� ������ �˴ϴ�.
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td bgcolor="#FFFFFF" colspan="3">
   		<%IF state = "9" THEN%>
   			<%=reservationdate%><input type="hidden" name="reservationdate" size=20 maxlength=10 value="<%=reservationdate%>"/>
   		<%ELSE%>
   			<input type="text" name="reservationdate" size=20 maxlength=10 value="<%=reservationdate%>" onClick="jsPopCal('reservationdate');"  style="cursor:pointer;"/>
   		<%END IF%>
		��) (<%=Left(Now(),10)%>)
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% sbGetwork "selMId",partMKid,"" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">���WD</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�۾� ���� ����</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="worktext" rows="5" cols="50"><%=worktext%></textarea>
	</td>
</tr>
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���� Play���<br/>(my10x10)����play</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="button" name="btnplaymainimg" value="�̹������" onClick="jsSetImg('<%=myplayimg%>','myplayimg','myplayimgdiv')" class="button"/>
		<div id="myplayimgdiv" style="padding: 5 5 5 5">
			<%IF myplayimg <> "" THEN %>
				<img src="<%=myplayimg%>" border="0" height=100 onclick="jsImgView('<%=myplayimg%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('myplayimg','myplayimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���γ��� ���<br/>(Play ����)</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<input type="button" name="btnplaymainimg" value="�̹������" onClick="jsSetImg('<%=playmainimg%>','playmainimg','playmainimgdiv')" class="button"/>
		<div id="playmainimgdiv" style="padding: 5 5 5 5">
			<%IF playmainimg <> "" THEN %>
				<img src="<%=playmainimg%>" border="0" height=100 onclick="jsImgView('<%=playmainimg%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('playmainimg','playmainimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<%END IF%>
		</div>
	</td>
	<td bgcolor="#FFFFFF">������ : <input type="text" name="mainTopBGColor" value="<%=mainTopBGColor%>" class="text" style="width:80px;" /><br/>�� �÷��ڵ� �տ� # �� �ٿ��ּ��� ex)#F9F9F9</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����(�⺻)</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnbeforeimg" value="�̹������" onClick="jsSetImg('<%=beforeimg%>','beforeimg','beforeimgdiv')" class="button"/>
		<div id="beforeimgdiv" style="padding: 5 5 5 5">
			<%IF beforeimg <> "" THEN %>
				<img src="<%=beforeimg%>" border="0" height=100 onclick="jsImgView('<%=beforeimg%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('beforeimg','beforeimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<%END IF%>
		</div>
	</td>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����(���ý�)</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnlistImg" value="�̹������" onClick="jsSetImg('<%=afterimg%>','afterimg','afterimgdiv')" class="button"/>
		<div id="afterimgdiv" style="padding: 5 5 5 5">
			<%IF afterimg <> "" THEN %>
				<img src="<%=afterimg%>" border="0" height=100 onclick="jsImgView('<%=afterimg%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('afterimg','afterimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ܹ���̹���</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="button" name="btntopbgimg" value="�̹������" onClick="jsSetImg('<%=topbgimg%>','topbgimg','topbgimgdiv')" class="button"/>
		<div id="topbgimgdiv" style="padding: 5 5 5 5">
			<%If topbgimg <> "" THEN %>
				<img src="<%=topbgimg%>" border="0" height=100 onclick="jsImgView('<%=topbgimg%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('topbgimg','topbgimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">PCWEB EXECUTE<br/>��뿩�� (����������)</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="radio" name="exec_check" value="1" <%=chkiif(exec_check = "1" Or mo_exec_check ="","checked","")%> onclick="execchk('1','W');">����
		<input type="radio" name="exec_check" value="2" <%=chkiif(exec_check = "2","checked","")%> onclick="execchk('2','W');">���
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">PC web - �̹��� & HTML</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="viewcontents" rows="10" cols="90"><%=viewcontents%></textarea>
	</td>
</tr>
<tr style="display:<%=chkiif(exec_check = "1" Or exec_check = "","none","block")%>" id="wc">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">PCWEB EXECUTE FilePath (����������)</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="exec_filepath" size="50" value="<%=exec_filepath%>"/>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">MOBILE EXECUTE<br/>��뿩�� (����������)</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="radio" name="mo_exec_check" value="1" <%=chkiif(mo_exec_check = "1" Or mo_exec_check ="","checked","")%> onclick="execchk('1','M');">����
		<input type="radio" name="mo_exec_check" value="2" <%=chkiif(mo_exec_check = "2","checked","")%> onclick="execchk('2','M');">���
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">
		Mobile&App - �̹��� & HTML
		<% If idx <> "" Then %>
			<br /><br /><br /><br /><br /><input type="button" value="�̹�������" class="button" onClick="jsManagePlayImage('<%=idx%>');">
		<% End If %>		
	</td>
	<td bgcolor="#FFFFFF" colspan="3" >
		<span id="moc1" style="display:<%=chkiif(mo_exec_check = "2","block","none")%>;">
			������ ����� ��� ���� <br/>
			ex) /play/groundcnt/iframe_60578.asp<br/>
			ex2) form action�� #ID�� ���� �ʼ�
		</span>
		<span id="moc2" style="display:<%=chkiif(mo_exec_check = "1" Or mo_exec_check = "","block","none")%>;">
			���۾� ���� <br/>
			ex) &lt;iframe src="/play/groundcnt/iframe_60578.asp" width="100%" height="1000" frameborder="0" scrolling="no" class="autoheight"&gt;&lt;/iframe&gt;<br/>
		</span>
		<textarea name="mo_contents" rows="10" cols="90"><%=mo_contents%></textarea>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�������� �±�</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="button" name="btnviewImg" value="�±� ����" onClick="jsTagview('<%=gidx%>','<%=idx%>')" class="button"/><br/><br/>
		���±װ����� �˾����� ���� �մϴ� ���� ��� ���ּ���.��
	</td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF" >
	<td colspan="4" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
		<input type="button" value=" ��� " class="button" onclick="history.back();"/>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
