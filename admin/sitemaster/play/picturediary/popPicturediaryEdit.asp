<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/play/playCls.asp" -->
<%
	Dim idx , listimg , viewimg , state , reservationdate , viewtitle , viewtext , playcate
	Dim viewno , orgimg , worktext
	Dim oPlay
	idx = request("idx")	
    playcate = 5 '�׸��ϱ�
	'//db 1row
	set oPlay = new CPlayContents
		 oPlay.FRectIdx = idx
		
		if idx <> "" Then
			oPlay.GetOneRowContent()

			if oPlay.FResultCount > 0 then			
				listimg = oPlay.FOneItem.Flistimg
				viewimg = oPlay.FOneItem.Fviewimg
				viewtitle = oPlay.FOneItem.Fviewtitle
				viewtext = oPlay.FOneItem.Fviewtext
				viewno = oPlay.FOneItem.Fviewno
				orgimg = oPlay.FOneItem.Forgimg
				worktext = oPlay.FOneItem.Fworktext
				reservationdate = oPlay.FOneItem.Freservationdate
				state = oPlay.FOneItem.Fstate
			end if	
		end if
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

	function jsSetImg(sImg, sName, sSpan){	
		document.domain ="10x10.co.kr";
		
		var winImg;
		winImg = window.open('/admin/sitemaster/play/lib/pop_theme_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	function jsTagview(idx){	
		var poptag;
		poptag = window.open('/admin/sitemaster/play/lib/pop_tagReg.asp?idx='+idx+'&playcate='+<%=playcate%>,'poptag','width=1024,height=768,scrollbars=yes,resizable=yes');
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

		if (!frm.viewtext.value){
			alert('�󼼳����� ������ּ���');
			frm.viewtext.focus();
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
//-->
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="inputfrm" method="post" action="PicturediaryProc.asp">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="pdviewimg" value="<%=viewimg%>">
<input type="hidden" name="pdlistimg" value="<%=listimg%>">
<input type="hidden" name="pdorgimg" value="<%=orgimg%>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>Play&gt;&gt;�׸��ϱ� ���/����</b></font>
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
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnorgmg" value="�̹������" onClick="jsSetImg('<%=orgimg%>','pdorgimg','orgimgdiv')" class="button"/>
		<div id="orgimgdiv" style="padding: 5 5 5 5">
			<%If orgimg <> "" THEN %>			
				<img src="<%=orgimg%>" border="0" height=100 onclick="jsImgView('<%=orgimg%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('pdorgimg','orgimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
				�̹��� �ּ� : <%=orgimg%>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�󼼳���</td>
	<td bgcolor="#FFFFFF">
		<textarea name="viewtext" rows="12" cols="50"><%=viewtext%></textarea>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�۾� ���� ����</td>
	<td bgcolor="#FFFFFF">
		<textarea name="worktext" rows="8" cols="50"><%=worktext%></textarea>
	</td>
</tr>
<% If idx <> "0" Then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">����Ʈ�̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnlistImg" value="�̹������" onClick="jsSetImg('<%=listimg%>','pdlistimg','listimgdiv')" class="button"/>
		<div id="listimgdiv" style="padding: 5 5 5 5">
			<%IF listimg <> "" THEN %>			
				<img src="<%=listimg%>" border="0" height=100 onclick="jsImgView('<%=listimg%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('pdlistimg','listimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnviewImg" value="�̹������" onClick="jsSetImg('<%=viewimg%>','pdviewimg','viewimgdiv')" class="button"/>
		<div id="viewimgdiv" style="padding: 5 5 5 5">
			<%IF viewimg <> "" THEN %>			
				<img src="<%=viewimg%>" border="0" height=100 onclick="jsImgView('<%=viewimg%>');" alt="�����ø� Ȯ�� �˴ϴ�"/>
				<a href="javascript:jsDelImg('pdviewimg','viewimgdiv');"><img src="/images/icon_delete2.gif" border="0"/></a>
			<%END IF%>
		</div>
		(�̹��� Size�� 560x560 �Դϴ�.)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�������� �±�</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnviewImg" value="�±� ����" onClick="jsTagview('<%=idx%>')" class="button"/><br/><br/>
		���±װ����� �˾����� ���� �մϴ� ���� ��� ���ּ���.��
	</td>
</tr>
<% End If %>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="button" value=" ���� " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
		<input type="button" value=" ��� " class="button" onclick="history.back();"/>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
