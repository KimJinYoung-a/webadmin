<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �̺�Ʈ SNS ���
' History : 2016-10-27 ����ȭ ����
' History : 2017-04-17 ���¿� ����
'####################################################
%>
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<%
Dim eCode : eCode = Request("evtid")
  	IF (eCode = "" or eCode = "0" or isnull(eCode)) THEN
	  	response.write "<script language='javascript'>alert('�߸��� ���� �Դϴ�.'); window.close();</script>"
	  	dbget.close(): response.End
	END IF 

Dim cESNS
Dim idx, fbtitle, fbdesc, fbimage, twlink, twtag1, twtag2, katitle, kaimage, kalink
Dim arrImg, slen, sImgName
 set cESNS = new ClsEventSNS
 	cESNS.FECode = eCode
  	IF (eCode <> "" and eCode <> "0" and not isnull(eCode)) THEN
	  	cESNS.fnGetEventItemSNSCont	
	  	idx 	= cESNS.Fidx
	  	fbtitle	= cESNS.Ffbtitle
	  	fbdesc	= cESNS.Ffbdesc
	  	fbimage	= cESNS.Ffbimage
	  	twlink  = cESNS.Ftwlink
	  	twtag1	= cESNS.Ftwtag1
	  	twtag2	= cESNS.Ftwtag2
	  	katitle	= cESNS.Fkatitle
	  	kaimage	= cESNS.Fkaimage
	  	kalink	= cESNS.Fkalink
	END IF  	
 set cESNS = nothing
%>
<style>
input:focus::-webkit-input-placeholder {opacity: 0;}
input:focus::-moz-placeholder {opacity: 0;}
input::-webkit-input-placeholder {color:#b2b2b2;}
input::-moz-placeholder {color:#b2b2b2;} /* firefox 19+ */
input:-ms-input-placeholder {color:#b2b2b2;} /* ie */
input:-moz-placeholder {color:#b2b2b2;}
</style>
<script type="text/javascript">
function jsEvtSnsSubmit(frm){
 	if(!frm.fbtitle.value){
	 	alert("Ÿ��Ʋ�� �Է����ּ���");
	 	frm.fbtitle.focus();
	 	return false;
 	}
 	if(!frm.fbdesc.value){
	 	alert("����Ÿ��Ʋ�� �Է��� �ּ���.");
	 	frm.fbdesc.focus();
	 	return false;
 	}
 	if(!frm.fbimage.value){
	 	alert("���̽��� �̹�����ũ�� �Է��� �ּ���.");
	 	frm.fbimage.focus();
	 	return false;
 	}
 	if(!frm.twlink.value){
	 	alert("Ʈ���� ��ũ�� �Է��� �ּ���.");
	 	frm.twlink.focus();
	 	return false;
 	}
 	if(!frm.twtag2.value){
	 	alert("Ʈ���� �±׸� �Է��� �ּ���.");
	 	frm.twtag2.focus();
	 	return false;
 	}
 	if(!frm.katitle.value){
	 	alert("īī���� Ÿ��Ʋ�� �Է��� �ּ���.");
	 	frm.katitle.focus();
	 	return false;
 	}
 	if(!frm.kaimage.value){
	 	alert("īī���� �̹����� �Է��� �ּ���.");
	 	frm.kaimage.focus();
	 	return false;
 	}
 	if(!frm.kalink.value){
	 	alert("īī���� ��ũ�� �Է��� �ּ���.");
	 	frm.kalink.focus();
	 	return false;
 	}
}
</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̺�Ʈ SNS ���</div>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
<form name="frmG" method="get" action="do_eventsns_proc.asp" enctype="MULTIPART/FORM-DATA" onSubmit="return jsEvtSnsSubmit(this);">
<input type="hidden" name="eCode" value="<%=eCode%>">
<tr>
	<td>
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<input type="hidden" name="idx" value="<%= idx %>" >
			<tr>
				<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�ڵ�</td>
				<td bgcolor="#FFFFFF"><input type="text" name="eCodetext" size="10" value="<%= eCode %>" disabled ></td>
			</tr><tr></tr><tr></tr><tr></tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">Ÿ��Ʋ</td>
				<td bgcolor="#FFFFFF"><input type="fbtitle" placeholder="����� ������ ��ũ��"  size="70" name="fbtitle"  value="<%= fbtitle %>"></td>
			</tr><tr></tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">����Ÿ��Ʋ</td>
				<td bgcolor="#FFFFFF"><input type="fbdesc"  placeholder="������ �� �־��ּ���" size="70" name="fbdesc"  value="<%= fbdesc %>"></td>
			</tr><tr></tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">���̽��� �̹���</td>
				<td bgcolor="#FFFFFF"><input type="fbimage" placeholder="http://image.thefingers.co.kr/m/2017/event/174/img_fb.jpg" size="70" name="fbimage"  value="<%= fbimage %>"></td>
			</tr><tr></tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">��ũ</td>
				<td bgcolor="#FFFFFF"><input type="twlink" placeholder="http://www.thefingers.co.kr/event/evt174.asp" size="70" name="twlink" value="<%= twlink %>"></td>
			</tr><tr></tr><tr></tr><tr></tr><tr></tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">Ʈ����TAG2</td>
				<td bgcolor="#FFFFFF"><input type="twtag2" placeholder="#���ΰŽ� #����� ������ ��ũ��" size="50" name="twtag2" value="<%= twtag2 %>"></td>
			</tr><tr></tr>
			<!--
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">Ʈ����TAG1</td>
				<td bgcolor="#FFFFFF"><input type="twtag1" placeholder="" size="50" name="twtag1" value="<%= twtag1 %>"></td>
			</tr><tr></tr>
			-->
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">īī�� Ÿ��Ʋ</td>
				<td bgcolor="#FFFFFF"><input type="katitle" placeholder="[���ΰŽ�] ����� ������ ��ũ��\n������ �� �־��ּ���" size="70" name="katitle" value="<%= katitle %>"></td>
			</tr><tr></tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">īī�� �̹���</td>
				<td bgcolor="#FFFFFF"><input type="kaimage" placeholder="http://image.thefingers.co.kr/m/2017/event/174/img_kakao.jpg [ 200x200 �̹���]" size="70" name="kaimage"   value="<%= kaimage %>"></td>
			</tr><tr></tr>
			<tr>
				<td align="center" bgcolor="<%= adminColor("tabletop") %>">īī�� ��ũ</td>
				<td bgcolor="#FFFFFF"><input type="kalink" placeholder="http://m.thefingers.co.kr/event/evt174.asp" size="70" name="kalink"   value="<%= kalink %>"></td>
			</tr>
		</table>
	</td>
	</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right" height="40">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>	
</form>	
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->