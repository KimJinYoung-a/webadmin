<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �÷�Ʈ���� ����
' Hieditor : 2012.03.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/color/colortrend_cls.asp"-->

<%
dim menupos ,ctcode ,iColorCd ,isusing ,state ,mainimage ,mainimagelink ,textimage , listimage , Nmainimage , mainimagelinknew
dim startdate ,lastupdate ,regdate ,lastadminid ,i ,ocolor , colortitle
Dim partmdid ,  partwdid , viewno
	ctcode = request("ctcode")
	menupos = request("menupos")
	
'//������
set ocolor = new ccolortrend_list
	ocolor.frectctcode = ctcode
	
	if ctcode <> "" then
		ocolor.getcolortrend_one()
		
		if ocolor.ftotalcount > 0 then
			ctcode = ocolor.foneitem.fctcode
			iColorCd = ocolor.foneitem.fcolorCode
			isusing = ocolor.foneitem.fisusing
			state = ocolor.foneitem.fstate
			mainimage = ocolor.foneitem.fmainimage
			mainimagelink = ocolor.foneitem.fmainimagelink
			mainimagelinknew = ocolor.foneitem.fmainimagelinknew
			textimage = ocolor.foneitem.ftextimage
			startdate = ocolor.foneitem.fstartdate
			lastupdate = ocolor.foneitem.flastupdate
			regdate = ocolor.foneitem.fregdate
			lastadminid = ocolor.foneitem.flastadminid
			viewno = ocolor.foneitem.Fviewno
			partwdid = ocolor.foneitem.Fpartwdid
			partmdid = ocolor.foneitem.Fpartmdid
			listimage = ocolor.foneitem.Flistimg
			Nmainimage = ocolor.foneitem.FNmainimg
			colortitle = ocolor.foneitem.Fcolortitle
		end if	
	end if
set ocolor = nothing
	
if isusing = "" then isusing = "Y"
if mainimagelink = "" then mainimagelink = "<map name='Mapmainimage'></map>"	
if mainimagelinknew = "" then mainimagelinknew = "<map name='Mapmainimagenew'></map>"	
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
	 wImgView = window.open('/admin/itemmaster/colortrend_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
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
		winImg = window.open('/admin/itemmaster/colortrend_imagereg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}
		
	//����
	function jsSubmit(){
		
		if (frm.viewno.value == ''){
			alert('No.�� �Է� ���ּ���');
			return;
		}
		
		if (frm.iCD.value == ''){
			alert('�÷�Ĩ�� ������ �ּ���');
			return;
		}

		if (frm.startdate.value == ''){
			alert('�������� ������ �ּ���');
			frm.startdate.focus();
			return;
		}

		if (frm.state.value == ''){
			alert('���¸� ������ �ּ���');
			frm.state.focus();
			return;
		}

		if (frm.isusing.value == ''){
			alert('��뿩�θ� ������ �ּ���');
			frm.isusing.focus();
			return;
		}

		if (frm.partmdid.value == ''){
			alert('�����MD�� ������ �ּ���');
			frm.partmdid.focus();
			return;
		}

		if (frm.partwdid.value == ''){
			alert('�����WD�� ������ �ּ���');
			frm.partwdid.focus();
			return;
		}
		
		frm.submit();
	}

	//�����ڵ� ����
	function selColorChip(cd) {
		var i;
		document.frm.iCD.value= cd;
		for(i=0;i<=30;i++) {
			document.all("cline"+i).bgColor='#DDDDDD';
		}
		if(!cd) document.all("cline0").bgColor='#DD3300';
		else document.all("cline"+cd).bgColor='#DD3300';
	}	
	
</script>

<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="/admin/itemmaster/colortrend_process.asp">
<input type="hidden" name="mode" value="trendreg">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mainimage" value="<%=mainimage%>">
<input type="hidden" name="textimage" value="<%=textimage%>">
<input type="hidden" name="listimage" value="<%=listimage%>">
<input type="hidden" name="Nmainimage" value="<%=Nmainimage%>">
<input type="hidden" name="ctcode" value="<%=ctcode%>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">No.</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="viewno" value="<%=viewno%>" size="10">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�÷�Ʈ�����ڵ�</td>
	<td bgcolor="#FFFFFF">
		<%= ctcode %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�÷�Ĩ</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="iCD" value="<%= iColorCd %>">
		<%=FnSelectColorBar(iColorCd,32)%>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
	<td bgcolor="#FFFFFF">
   		<input type="text" name="colortitle" size="50" value="<%=colortitle%>"/>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">������</td>
	<td bgcolor="#FFFFFF">
   		<input type="text" name="startdate" size=10 maxlength=10 value="<%=startdate%>" onClick="jsPopCal('startdate');"  style="cursor:pointer;">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
	<td bgcolor="#FFFFFF">
		<% Drawcolortrendstate "state" , state ,"" %>		
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partmdid",partmdid,"","23" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����WD</td>
	<td bgcolor="#FFFFFF">
		<% sbGetpartid "partwdid",partwdid,"","12" %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">��뿩��</td>
	<td bgcolor="#FFFFFF">
		<% drawSelectBoxUsingYN "isusing", isusing %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����� �̹���<br/>(���θ���Ʈ��)</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnmainlistimg" value="�̹������" onClick="jsSetImg('<%=listimage%>','listimage','listimagediv')" class="button">
		<div id="listimagediv" style="padding: 5 5 5 5">
			<%IF listimage <> "" THEN %>			
				<img src="<%=listimage%>" border="0" width=100 height=100 onclick="jsImgView('<%=listimage%>');" alt="�����ø� Ȯ�� �˴ϴ�">
				<a href="javascript:jsDelImg('listimage','listimagediv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">���� �̹���<br/>(2013������ ������)</td>
	<td bgcolor="#FFFFFF">		
		<input type="button" name="btnBan2011" value="�̹������" onClick="jsSetImg('<%=mainimage%>','mainimage','mainimagediv')" class="button">
		<div id="mainimagediv" style="padding: 5 5 5 5">
			<%IF mainimage <> "" THEN %>			
				<img src="<%=mainimage%>" border="0" width=100 height=100 onclick="jsImgView('<%=mainimage%>');" alt="�����ø� Ȯ�� �˴ϴ�">
				<a href="javascript:jsDelImg('mainimage','mainimagediv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">���� �̹��� (2013��)</td>
	<td bgcolor="#FFFFFF">		
		<input type="button" name="btnviewimg" value="�̹������" onClick="jsSetImg('<%=Nmainimage%>','Nmainimage','Nmainimagediv')" class="button">&lt;-- �űԵ�� �����̹����� 2013������ ���ּ���
		<div id="Nmainimagediv" style="padding: 5 5 5 5">
			<%IF Nmainimage <> "" THEN %>			
				<img src="<%=Nmainimage%>" border="0" width=100 height=100 onclick="jsImgView('<%=Nmainimage%>');" alt="�����ø� Ȯ�� �˴ϴ�">
				<a href="javascript:jsDelImg('Nmainimage','Nmainimagediv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
			<br/>�� �̹��� ��� ������� 1140 x 640 �Դϴ�.
		</div>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����̹�����</td>
	<td bgcolor="#FFFFFF">
		�� �� �̸� ���� ���� ������<br>
		<textarea name="mainimagelink" cols="80" rows="6"><%=mainimagelink%></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�����̹����� (2013��)</td>
	<td bgcolor="#FFFFFF">
		�� �� �̸� ���� ���� ������<br>
		<textarea name="mainimagelinknew" cols="80" rows="6"><%=mainimagelinknew%></textarea>
	</td>
</tr>
<!-- <tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">���̹���</td>
	<td bgcolor="#FFFFFF">
		<input type="button" name="btnBan2011" value="�̹������" onClick="jsSetImg('<%=textimage%>','textimage','textimagediv')" class="button">
		<div id="textimagediv" style="padding: 5 5 5 5">
			<%IF textimage <> "" THEN %>			
				<img src="<%=textimage%>" border="0" width=100 height=100 onclick="jsImgView('<%=textimage%>');" alt="�����ø� Ȯ�� �˴ϴ�">
				<a href="javascript:jsDelImg('textimage','textimagediv');"><img src="/images/icon_delete2.gif" border="0"></a>
			<%END IF%>
		</div>
	</td>
</tr> -->
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�ֱټ���</td>
	<td bgcolor="#FFFFFF">
		<%= lastadminid %>
		<Br><%= lastupdate %>	
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="button" onclick="jsSubmit();" class="button" value="����"></td>
</tr>	
</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->