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
	Dim idx , listimg , state , reservationdate , viewtitle , viewtext , playcate , mainimg
	Dim viewno , worktext , partMKid,partWDid , i
	Dim oPlay , oground
	idx = request("idx")
    playcate = 1 'ground

	'//db 1row
	set oPlay = new CPlayContents
		 oPlay.FRectIdx = idx

		if idx <> "" Then
			oPlay.GetRowGroundMain()

			if oPlay.FResultCount > 0 then
				listimg					= oPlay.FOneItem.Flistimg
				viewtitle				= oPlay.FOneItem.Fviewtitle
				mainimg					= oPlay.FOneItem.Fmainimg
				reservationdate			= oPlay.FOneItem.Freservationdate
				state					= oPlay.FOneItem.Fstate
				viewno					= oPlay.FOneItem.Fviewno
				worktext				= oPlay.FOneItem.Fworktext
				partMKid				= oPlay.FOneItem.FpartMKid
				partWDid				= oPlay.FOneItem.FpartWDid
			end if
		end If
		
	set oPlay = Nothing
%>

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

	function AddNewContents(idx,gidx){
		var popwin = window.open('/admin/sitemaster/play/ground/groundweekEdit.asp?idx=' + idx+'&gidx='+gidx,'cateHotPosCodeEdit','width=800,height=500,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	function jsSetImg(sImg, sName, sSpan){
		document.domain ="10x10.co.kr";

		var winImg;
		winImg = window.open('/admin/sitemaster/play/lib/pop_theme_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsTagview(gidx , idx){	
		var poptag;
		poptag = window.open('/admin/sitemaster/play/lib/pop_tagReg.asp?idx='+gidx+'&subidx='+idx+'&playcate='+<%=playcate%>,'poptag','width=500,height=400,scrollbars=yes,resizable=yes');
		poptag.focus();
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

	function jsSetItem(idx){//��ǰ��� Ȯ��
		var popitem;
		popitem = window.open('pop_itemReg.asp?idx='+idx,'popitem','width=500,height=400,scrollbars=yes,resizable=yes');
		popitem.focus();
	}
//-->
</script>
<script type="text/javascript">
<!--
	function copy_url(url) {
		var IE=(document.all)?true:false;
		if (IE) {
			if(confirm("�� ���� URL �ּҸ� Ŭ�����忡 �����Ͻðڽ��ϱ�?"))
				window.clipboardData.setData("Text", url);
		} else {
			temp = prompt("�� ���� Ʈ���� �ּ��Դϴ�. Ctrl+C�� ���� Ŭ������� �����ϼ���", url);
		}
	}
//-->
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="groundProc.asp">
<input type="hidden" name="idx" value="<%= idx %>"/>
<input type="hidden" name="groundtitleimg" value="<%=listimg%>"/>
<input type="hidden" name="playmainimg" value="<%=mainimg%>"/>
<input type="hidden" name="position" value="main"/>
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>Play&gt;&gt;ground ���� ���/����</b></font><br/><br/>
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
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">No.</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="viewno" value="<%=viewno%>" size="10"/>�� ���ڸ� �����ּ��� ��
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="viewtitle" value="<%=viewtitle%>" size="50"/>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">������</td>
	<td bgcolor="#FFFFFF">
   		<%IF state = "9" THEN%>
   			<%=reservationdate%><input type="hidden" name="reservationdate" size=20 maxlength=10 value="<%=reservationdate%>"/>
   		<%ELSE%>
   			<input type="text" name="reservationdate" size=20 maxlength=10 value="<%=reservationdate%>" onClick="jsPopCal('reservationdate');"  style="cursor:pointer;"/>
   		<%END IF%>
		��) (<%=Left(Now(),10)%>)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<% Draweventstate2 "state" , state ,"" %> �� ������ �ؼ� �����Ͽ��� ������ =< ���� �̾�߸� ������ �˴ϴ�.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����</td>
	<td bgcolor="#FFFFFF">
		<% sbGetwork "selMId",partMKid,"" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">���WD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">Ÿ��Ʋ�̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnorgmg" value="�̹������" onClick="jsSetImg('<%=listimg%>','groundtitleimg','groundtitleimgdiv')" class="button"/>
		<div id="groundtitleimgdiv" style="padding: 5 5 5 5">
			<%If listimg <> "" THEN %>
				<img src="<%=listimg%>" border="0" height=100 onclick="jsImgView('<%=listimg%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('groundtitleimg','groundtitleimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				�̹��� �ּ� : <%=listimg%>
			<%END IF%>
		</div>
	</td>
</tr>
<!-- <tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">play���ο�����</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnorgmg" value="�̹������" onClick="jsSetImg('<%=mainimg%>','playmainimg','playmainimgdiv')" class="button"/>
		<div id="playmainimgdiv" style="padding: 5 5 5 5">
			<%If mainimg <> "" THEN %>
				<img src="<%=mainimg%>" border="0" height=100 onclick="jsImgView('<%=mainimg%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('playmainimg','playmainimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				�̹��� �ּ� : <%=mainimg%>
			<%END IF%>
		</div>
	</td>
</tr> -->
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�۾� ���� ����</td>
	<td bgcolor="#FFFFFF">
		<textarea name="worktext" rows="8" cols="50"><%=worktext%></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
		<input type="button" value=" ��� " class="button" onclick="history.back();"/>
	</td>
</tr>
</form>
</table>

<% If idx > "0" Then %>
<%
	set oground = new CPlayContents
		oground.FPageSize = 50
		oground.FCurrPage = 1
		oground.FRPlaycate = playcate
		oground.FRectIdx = idx
		oground.fnGetGroundSubList()

	
%>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 10px 0;">
<tr>
	<td align="left">
		<font color="red"> �� ����Ʈ ���� : ���°� ������ �Ͱ� ������ =< ���� �ΰ͸� ������ �˴ϴ�. ������ No. ��ȣ(��������) ������ ����˴ϴ�.<br/>�� �ϴ� ����Ʈ�� ���� �κ��� �����ø� ���� �������� �����ϴ�.</font>
	</td>
	<td align="right">
		<input type="button" class="button" value="�űԵ��" onclick="AddNewContents('0','<%=idx%>');">
	</td>
</tr>
</table>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="5%">ȸ��</td>
	<td width="10%">����</td>
	<td width="10%">Ÿ��Ʋ�̹���</td>
	<td width="10%">����</td>
	<td width="5%">������</td>
	<td width="10%">�±�</td>
	<td width="5%">�����</td>
	<td width="5%">��ȹWD</td>
	<td width="7%">����</td>
</tr>
<% if oground.FresultCount > 0 then %>
<% for i=0 to oground.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
	<td align="center"><%= oground.FItemList(i).Fviewno %></td>
	<td align="center"><%= geteventstate(oground.FItemList(i).Fstate) %> (<%=oground.FItemList(i).Fstate %>)<br/><br/>
	<a href="http://www.10x10.co.kr/play/playGround_review.asp?gidx=<%=idx%>&gcidx=<%=oground.FItemList(i).Fidxsub%>" target="_blank">PC�̸�����</a>&nbsp;&nbsp;<input type="button" onclick="copy_url('http://www.10x10.co.kr/play/playGround.asp?gidx=<%=idx%>&gcidx=<%=oground.FItemList(i).Fidxsub%>')" value="PC-URL����"/><br/>
	<a href="http://m.10x10.co.kr/play/playGround_review.asp?idx=<%=oground.FItemList(i).Fmo_idx%>&contentsidx=<%=oground.FItemList(i).Fidxsub%>" target="_blank">M �̸�����</a>&nbsp;&nbsp;<input type="button" onclick="copy_url('http://m.10x10.co.kr/play/playGround.asp?idx=<%=oground.FItemList(i).Fmo_idx%>&contentsidx=<%=oground.FItemList(i).Fidxsub%>')" value="M-URL����"/></td>
	<td align="center"><img src="<%= oground.FItemList(i).Fviewthumbimg1 %>" width="80"/>&nbsp;<img src="<%= oground.FItemList(i).Fviewthumbimg2 %>" width="80"/></td>
	<td align="center"><%= oground.FItemList(i).Fviewtitle %></td>
	<td align="center"><%= left(oground.FItemList(i).Freservationdate,10) %></td>
	<td align="center"><a href="#" onclick="jsTagview('<%=idx%>','<%= oground.FItemList(i).Fidxsub %>');" style="cursor:pointer;"><%=chkiif(oground.FItemList(i).Ftagcnt>0,"���","�̵��")%>(<%=oground.FItemList(i).Ftagcnt%>) </a></td>
	<td align="center"><%= oground.FItemList(i).FpartMKname %></td>
	<td align="center"><%= oground.FItemList(i).FpartWDname %></td>
	<td align="center">
		<input type="button" class="button" value="����" onclick="AddNewContents('<%= oground.FItemList(i).Fidxsub %>','<%=idx%>');"/>
		<input type="button" value="�����۵��[<%=oground.FItemList(i).Fitemcnt%>]" onClick="jsSetItem('<%= oground.FItemList(i).Fidxsub %>')" class="button"/>
	</td>
</tr>
<% Next %>
<% end if %>
</table>
<%
	set oground = nothing
%>
<% End If %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
