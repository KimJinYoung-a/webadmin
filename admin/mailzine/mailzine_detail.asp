<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� ������
' History : 2018.04.27 �̻� ����(���Ϸ� ���� ���� ���Ϸ��� �߼� ���� ����. ���� �������� ����.)
'			2019.06.24 ������ ����(���ø� ��� �ű� �߰�)
'			2020.05.28 �ѿ�� ����(TMS ���Ϸ� �߰�)
'###########################################################
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mailzinecls.asp"-->
<%
CONST MAXHeightPX = 1400    '''�� ��ġ�� ���ؼ��� Ȯ������ ����.. (2,000px ���� ���� ������ 2���� �־����� �� ������찡 ����)

dim idx, code, omail ,yyyy1, mm1, dd1 , tmp , area, mngUserid, mailergubun
dim title,regdate,img1,img2,img3,img4,imgmap1,imgmap2,imgmap3,imgmap4,isusing,gubun,memgubun,secretGubun,reservationDATE
	idx = requestcheckvar(getNumeric(request("idx")),10)
	mailergubun = requestcheckvar(request("mailergubun"),16)

if mailergubun="" or isnull(mailergubun) then
	response.write "���Ϸ� ������ �����ϴ�."
	dbget.close() : response.end
end if

set omail = new CMailzineList
	omail.frectidx = idx
	omail.frectmailergubun = mailergubun

	'//idx ���� ������쿡�� ����(�������)
	if idx <> "" then
		omail.MailzineDetail()

		if omail.ftotalcount > 0 then
			title = omail.foneitem.ftitle
			regdate = omail.foneitem.fregdate
			img1 = omail.foneitem.fimg1
			img2 = omail.foneitem.fimg2
			img3 = omail.foneitem.fimg3
			img4 = omail.foneitem.fimg4
			imgmap1 = omail.foneitem.fimgmap1
			imgmap2 = omail.foneitem.fimgmap2
			imgmap3 = omail.foneitem.fimgmap3
			imgmap4 = omail.foneitem.fimgmap4
			isusing = omail.foneitem.fisusing
			gubun = omail.foneitem.fgubun
			area = omail.foneitem.farea
			mngUserid = omail.foneitem.fmngUserid
			memgubun = omail.foneitem.fmemgubun
			secretGubun = omail.foneitem.fsecretGubun
			reservationDATE = omail.foneitem.freservationDATE
			tmp = split(omail.foneitem.fregdate,".")
			yyyy1 = tmp(0)
			mm1 = tmp(1)
			dd1 = tmp(2)
			code = mm1 & dd1
		end if
	end if

if area = "" then area = "ten_all"
if isusing = "" then isusing = "N"
if memgubun = "" then memgubun ="member_all"
If secretGubun = "" then secretGubun="N"
%>
<style> 
#mask {  
	position:absolute;  
	z-index:9000;  
	background-color:#000;  
	display:none;  
	left:0;
	top:0;
} 
.window{
	display: none;
	position:absolute;  
	left:100px;
	bottom:10px;
	z-index:10000;
}
</style> 
<script type="text/javascript" src="/js/jquery-latest.js"></script>
<script type="text/javascript">

function wrapWindowByMask(){
	//ȭ���� ���̿� �ʺ� ���Ѵ�.
	var maskHeight = $(document).height();  
	var maskWidth = $(window).width();  

	//����ũ�� ���̿� �ʺ� ȭ�� ������ ����� ��ü ȭ���� ä���.
	$('#mask').css({'width':maskWidth,'height':maskHeight});  

	//�ִϸ��̼� ȿ�� - �ϴ� 1�ʵ��� ��İ� �ƴٰ� 80% �������� ����.
	$('#mask').fadeIn(1000);      
	$('#mask').fadeTo("slow",0.8);    

	//������ ���� �� ����.
	$('.window').show();
}

$(document).ready(function(){
	//���� �� ����
	$('.openMask').click(function(e){
		e.preventDefault();
		wrapWindowByMask();
	});

	//�ݱ� ��ư�� ������ ��
	$('.window .close').click(function (e) {  
	    //��ũ �⺻������ �۵����� �ʵ��� �Ѵ�.
	    e.preventDefault();  
	    $('#mask, .window').hide();
	});
});
</script>
<script language="JavaScript">
	<% if date() >= "2017-11-13" and date() < "2017-11-18" then %>
		alert('[�ſ��߿�]\n\n 11�� 20�� ���� \n���Ϲ߼� ����ڴ�\n������ �븮 �Դϴ�.\n\n 11�� 20������ ���Ϲ߼��� \n�������븮���� ��!! �˷��ּ���.');
	<% end if %>

	alert('[�ſ��߿�]\n�̹��� ���� ���̴� 1,400 px �̸�\n\n�̹����� ��ũ ��Ͻ� ���� ������ ���\n\n�̹����� Ÿ�� target="_top"\n\n�̹����� �̸� �����Ұ�');
	function checkok(frm){
		if (document.modify.gubun.value == "1"){
			if (modify.isusing.value==''){
				alert('���⿩�θ� �������ּ���');
				modify.isusing.focus();
				return;
			}
			if ("<%= hour(now()) %>" >= 18){
				alert('�� 18�� ����\n�� �߼� ����ڿ���\n�� �ϼ� ���θ� �˸��� �ʾ������\n�� ���� �߼��� ���� ������\n�� �̿����� å���� ���������ڰ� ���� �˴ϴ�.');
			}
			frm.submit();
			document.getElementById('goproc').disabled = true;
		}else{
			/* �ð�üũ */
			if ("<%= hour(now()) %>" <= 17){
				if (confirm("�ۼ���Ȳ�� �ϼ��Դϴ�. Ȯ�ι�ư�� �����ø� ������ �Ұ��մϴ�.\n�����Ͻðڽ��ϱ�?") == true) {
					frm.submit();
				}else{
					return false;
				}
			}else{
				$('.openMask2').click(function(e){
					e.preventDefault();
					wrapWindowByMask();
				});
			}
		}
	}

	//���� �߼� �Ϸ� �� 1,2,3,4 �̹��� ����
	function imageedit(frm){
		if(confirm("���� �̹����� ���� �Ͻðڽ��ϱ�?")){
			if(confirm("�߸��� �̹����� ���ε� �� ���\n\n�߼۵� ���Ͽ� ������ ���� �� �ֽ��ϴ�.")){
				frm.mode.value='imageedit';
				frm.submit();
				document.getElementById('goimgproc').disabled = true;
			}
		}
		
	}

	function LastConfirm(yn){
		var frm = document.modify;
		if (yn == 'Y'){
			if (confirm("Ȯ�ι�ư�� �����ø� ������ �Ұ��մϴ�.\n�����Ͻðڽ��ϱ�?") == true) {
				frm.submit();
			}else{
				return false;
			}
		}else{
			 $('#mask, .window').hide();
		}
	}

	function delimg(imgnumber){
		frm_mail.action = '/admin/mailzine/mailzine_process.asp';
		frm_mail.imgnumber.value=imgnumber;
		frm_mail.mode.value='imgdel';
		frm_mail.submit();
	}

</script>

<table width="95%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">

<% IF application("Svr_Info")="Dev" THEN %>
	<form name="modify" method="post" action="<%=mailzine%>/ftp/mailzine_input_ok.asp" enctype="multipart/form-data" style="margin:0px;">
<% else %>
	<form name="modify" method="post" action="https://omailzine.10x10.co.kr/ftp/mailzine_input_ok.asp" enctype="multipart/form-data" style="margin:0px;">
<% end if %>

<input type="hidden" name="idx" value="<% = idx %>">
<input type="hidden" name="mailergubun" value="<%= mailergubun %>">
<input type="hidden" name="mode">
<input type="hidden" name="img1editname" value="<%= img1 %>">
<input type="hidden" name="img2editname" value="<%= img2 %>">
<input type="hidden" name="img3editname" value="<%= img3 %>">
<input type="hidden" name="img4editname" value="<%= img4 %>">
<tr bgcolor="#FFFFFF">
	<td colspan=2>
		<br>
		<font size=4>�� ���ǻ���. �ݵ�� ���� �ּž� �մϴ�. �ſ��߿�!!</font>
		<br>&nbsp;&nbsp;&nbsp;�̹����� Ÿ�� <font color="red">target="_top"</font> ���� �ֽð�, �̹����� <font color="red">�̸�</font>�� ��ġ�� �����ּ���.
		<br>&nbsp;&nbsp;&nbsp;�̹��� ������ ���� ���� <font color="red"><%= FormatNumber(MAXHeightPX,0) %> px</font> �� �������, �ƿ��迡�� ©���ϴ�.
		<br>&nbsp;&nbsp;&nbsp;�̹����� ��ũ ��Ͻ� <font color="red">���� ������</font> ����� �ֽñ� �ٶ��ϴ�. ���� ���� �Ϻ����п��� ©���ϴ�.
		<br>&nbsp;&nbsp;&nbsp;ex)
		<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(map name="ImgMap1")
		<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(area shape="rect" coords="6,5,130,114" href="http://www.10x10.co.kr" target="_top" onfocus="blur()")
		<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(area shape="rect" coords="11,136,694,665" href="http://www.10x10.co.kr" target="_top" onfocus="blur()")
		<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(/map)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">��������</td>
	<td><input type="text" name="title" class="input" size="55" value="<% = title %>"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">������ �����</td>
	<td><% DrawOneDateBox_2012 yyyy1,mm1,dd1 %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">�������</td>
	<td><%sbGetDesignerid "mngUserid",mngUserid, ""%></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">�������ۼ���Ȳ</td>
	<td>
		<select name="gubun" class="select">
			<option value="1" <% if gubun = "1" then response.write "selected"%>>�̿ϼ�</option>
			<option value="5" <% if gubun = "5" then response.write "selected"%>>�ϼ�</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">�ý�����_������</td>
	<td><%= Chkiif(isnull(reservationDATE), "������", reservationDATE) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">1���̹���</td>
	<td>
		<input type="file" name="img1" class="input" size="40"> <strong>(height <%= FormatNumber(MAXHeightPX,0) %> px �̸�)</strong>
		<br>
		<%
		if img1 <> "" then
			response.write mailzine&"/"&yyyy1&"/"&img1
		%>
			<input type="button" onclick="delimg('1');" class="button" value="�̹�������">
		<%
		end if
		%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">�̹����� �ڵ�ֱ�</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
	   <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		   <tr>
				<td>
					<% if imgmap1 = "" then %>
						<textarea name="imagemap1" rows="10" class="textarea" style="width:100%;"><map name="ImgMap1"></map></textarea>
					<% else %>
						<textarea name="imagemap1" rows="10" class="textarea" style="width:100%;"><% = imgmap1 %></textarea>
					<% end if %>
				</td>
		   </tr>
	   </table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">2���̹���</td>
	<td>
		<input type="file" name="img2" class="input" size="40"> <strong>(height <%= FormatNumber(MAXHeightPX,0) %> px �̸�)</strong>
		<br>
		<%
		if img2 <> "" then
			response.write mailzine&"/"&yyyy1&"/"&img2
		%>
			<input type="button" onclick="delimg('2');" class="button" value="�̹�������">
		<%
		end if
		%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">�̹����� �ڵ�ֱ�</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
	   <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	   <tr>
		<td>
			<% if imgmap2 = "" then %>
				<textarea name="imagemap2" rows="10" class="textarea" style="width:100%;"><map name="ImgMap2"></map></textarea>
			<% else %>
				<textarea name="imagemap2" rows="10" class="textarea" style="width:100%;"><%= imgmap2 %></textarea>
			<% end if %>
		</td>
	   </tr>
	   </table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">3���̹���</td>
	<td>
		<input type="file" name="img3" class="input" size="40"> <strong>(height <%= FormatNumber(MAXHeightPX,0) %> px �̸�)</strong>
		<br>
		<%
		if img3 <> "" then
			response.write mailzine&"/"&yyyy1&"/"&img3
		%>
			<input type="button" onclick="delimg('3');" class="button" value="�̹�������">
		<%
		end if
		%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">�̹����� �ڵ�ֱ�</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
	   <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	   <tr>
		<td>
			<% if imgmap3 = "" then %>
				<textarea name="imagemap3" rows="10" class="textarea" style="width:100%;"><map name="ImgMap3"></map></textarea>
			<% else %>
				<textarea name="imagemap3" rows="10" class="textarea" style="width:100%;"><%= imgmap3 %></textarea>
			<% end if %>
		</td>
	   </tr>
	   </table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">4���̹���</td>
	<td>
		<input type="file" name="img4" class="input" size="40"> <strong>(height <%= FormatNumber(MAXHeightPX,0) %> px �̸�)</strong>
		<br>
		<%
		if img4 <> "" then
			response.write mailzine&"/"&yyyy1&"/"&img4
		%>
			<input type="button" onclick="delimg('4');" class="button" value="�̹�������">
		<%
		end if
		%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">�̹����� �ڵ�ֱ�</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
	   <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	   <tr>
		<td>
			<% if imgmap4 = "" then %>
				<textarea name="imagemap4" rows="10" class="textarea" style="width:100%;"><map name="ImgMap4"></map></textarea>
			<% else %>
				<textarea name="imagemap4" rows="10" class="textarea" style="width:100%;"><%= imgmap4 %></textarea>
			<% end if %>
		</td>
	   </tr>
	   </table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">�߼�����</td>
	<td>
		<% Drawareagubun "area" , area , "class='select'" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">�߼�ȸ�����</td>
	<td>
		<% DrawMemberGubun "memgubun" , memgubun , "class='select'" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">����Ʈ����</td>
	<td>
		<% Drawisusing "isusing" , isusing , "class='select'" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">��ũ�� ����</td>
	<td>
		<% DrawsecretGubun "secretGubun" , secretGubun , "class='select'" %> ����Ʈ�����, ��ũ�� ������ Y�� �θ� Ÿ��Ʋ�� ����ǰ� Ŭ���� ���� �ʽ��ϴ�.
	</td>
</tr>

<% If gubun <> "5" or session("ssAdminPsn") = "7" or session("ssAdminPsn") = "11" Then %>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan=2><input type="button" id="goproc" value="������ ����" onclick="checkok(this.form);" class="button"></td>
	</tr>
	<% if reservationDATE <> "" then %>
		<tr bgcolor="#FFFFFF">
			<td align="center" colspan=2><br><br>
				<font color="red"><b>�̹��� ������, ���� �߼� �� 1,2,3,4 �̹����� �����Ҷ� ������ּ���.<br>�̹��� ���� �� �̹����ּҷ� �� ����Ǿ����� Ȯ�����ּ���!</b></font><br>
				<input type="button" id="goimgproc" value="1,2,3,4 �̹��� ����" onclick="imageedit(this.form);" class="openMask2">
			</td>
		</tr>
	<% end if %>
<% Else %>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan=2>
			<b>���� ���Ͽ����� �Ǿ��ִ� �����̹Ƿ� ���� ������ �����Ͻ� �� �����ϴ�.</b><br>
			<b>�� ������ �ʿ��� ��� [��ǰ���]�̽��񿡰� �����ٶ��ϴ�.</b>
		</td>
	</tr>
	<% if reservationDATE <> "" then %>
		<tr bgcolor="#FFFFFF">
			<td align="center" colspan=2><br><br>
				<font color="red"><b>�̹��� ������, ���� �߼� �� 1,2,3,4 �̹����� �����Ҷ� ������ּ���.<br>�̹��� ���� �� �̹����ּҷ� �� ����Ǿ����� Ȯ�����ּ���!</b></font><br>
				<input type="button" id="goimgproc" value="1,2,3,4 �̹����� ����" onclick="imageedit(this.form);" class="openMask2">
			</td>
		</tr>
	<% end if %>
<% End If %>
</form>
<form name="frm_mail" method="post">
	<input type="hidden" name="idx" value="<% = idx %>">
	<input type="hidden" name="imgnumber">
	<input type="hidden" name="mode">
</form>
</table>
<div id="mask"></div> 
<div class="window">
	<table>
	<tr>
		<td>
			<font size="3" color="#FFFFFF">�ٹ��ð��� ������ �������� �ϼ��� �����Դϴ�.<br><br>
				���� �޽����� �ڸ��� <strong>��ǰ���_�̽���</strong>�� ������<br><br>
				���� ������ �� �ֵ��� �޼����� �ֽð�<br><br>
				����� ���¶�� 010-2991-3466��<br><br>
				�ݵ�� ��ȭ ��Ź�帳�ϴ�(���ڴ� �����մϴ�)<br><br>
				���� ��ȭ �� �޾Ҵٸ� ���������� ��� �����ּ���<br><br>
				���� ������ �� �ϸ� ���� �߼��� �� �� �� �ֽ��ϴ�.<br><br>
			</font>
		</td>
	</tr>
	<tr>
		<td align="center">
			<input type="button" class="button" value="��, ����� �������ѵ�Ƚ��ϴ�" onclick="LastConfirm('Y')">
			<input type="button" class="button" value="�ƴϿ�, �����κ��� �ֽ��ϴ�" onclick="LastConfirm('N')">
		</td>
	</tr>
	</table>
</div> 
<% set omail = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
