<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü�����ֹ�������NEW
' History : �̻� ����
'			2020.07.23 �ѿ�� ����(�̸��Ϲ߼�. ���Ϸ��� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<!-- #include virtual="/lib/email/maillib2.asp"-->
<%
dim opartner,i,page, designer, idx, mode, mailfrom, reqhp, mailto, smstext, mailtitle,mailcontent
dim selltotal, buytotal, sqlstr, oMail
	page    = requestCheckVar(request("page"),10)
	designer = requestCheckVar(request("designer"),32)
	idx     = requestCheckVar(request("idx"),10)
	mode        = request("mode")
	mailfrom    = request("mailfrom")
	mailto	    = request("mailto")
	reqhp 	    = request("reqhp")
	smstext     = request("smstext")

if page="" then page=1

set opartner = new CPartnerUser
	opartner.FCurrpage = page
	opartner.FRectDesignerID = designer
	opartner.FPageSize = 1
	opartner.GetPartnerNUserCList

Dim groupid : groupid=opartner.FPartnerList(0).FGroupid

dim ogroup
set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = groupid
	ogroup.GetOneGroupInfo

dim osheet
set osheet = new COrderSheet
	osheet.FRectIdx = idx
	osheet.GetOneOrderSheetMaster

mailtitle = "[�ٹ�����] " + opartner.FRectDesignerID + " �귣���� �ֹ��� (" + osheet.FOneItem.Fbaljucode + ")�� �����Ǿ����ϴ�."

dim osheetdetail
set osheetdetail = new COrderSheet

selltotal =0
buytotal = 0
if mode="sendall" then
	if reqhp<>"" then
		sqlStr = " exec [db_sms].[dbo].[usp_SendSMS] '"+reqhp+"','1644-1851','"+html2db(smstext)+"'"
		dbget.execute sqlStr
	end if

	if mailto<>"" then

		osheetdetail.FRectIdx = idx
		osheetdetail.GetOrderSheetDetail

		'��ǰ����Ʈ �������� ���Ϲ߼� ���� �ۼ�
		ChgCont =""
		ChgCont = ChgCont + "<table width='600' border='0' align='center' cellpadding='0' cellspacing='0' class='a'>"
	    ChgCont = ChgCont + "<tr height='25' valign='top'>"
		ChgCont = ChgCont + "<td>"
		ChgCont = ChgCont + "<font color='red'><strong>�ֹ���</strong>&nbsp;<b>[" + opartner.FRectDesignerID + "]</b>&nbsp;&nbsp;�ֹ��ڵ�(" + osheet.FOneItem.Fbaljucode + ")</font></td>"
	    ChgCont = ChgCont + "</tr>"
	    ChgCont = ChgCont + "<tr valign='top'>"
		ChgCont = ChgCont + "<td>"
		ChgCont = ChgCont + "	<br>�ȳ��ϼ���. �ٹ������Դϴ�."
		ChgCont = ChgCont + "	<br>���� <b>�ֹ���/���⳻������>>�ֹ�������</b>���� �ֹ�Ȯ�� �� �԰� ��Ź�帳�ϴ�."
		ChgCont = ChgCont + "	<br>"
		ChgCont = ChgCont + "	<br>�귣��ID :<b>" + opartner.FRectDesignerID + "</b>"
		ChgCont = ChgCont + "	<br>�ֹ��ڵ� :<b>" + osheet.FOneItem.Fbaljucode + "</b>"
		ChgCont = ChgCont + "	<br><a href='http://scm.10x10.co.kr/'>���� �ٷΰ���>><a>"
		ChgCont = ChgCont + "	<br>"
		ChgCont = ChgCont + "	<br><b><font color='red'>[�ֹ�Ȯ��]</font></b>"
		ChgCont = ChgCont + "	<br>�ֹ����� Ȯ���Ͻ� �Ŀ��� �� �ֹ����������� <b>[�ֹ�Ȯ��]</b>���� ��ȯ�Ͽ��ֽð�,"
		ChgCont = ChgCont + "	<br>������ �����ϰų� ������ ���, �������ͷ� ������ �ֽðų�,"
		ChgCont = ChgCont + "	<br>�ֹ�Ȯ�������� �����Ͽ��ֽñ� �ٶ��ϴ�."
		ChgCont = ChgCont + "	<br>"
		ChgCont = ChgCont + "	<br><b><font color='red'>[���Ϸ�]</font></b>"
		ChgCont = ChgCont + "	<br>����ϽǶ��� ���������� �������, Ȯ�������� �����Ͽ� �ֽð�,"
		ChgCont = ChgCont + "	<br><b>[���Ϸ�]</b>�� ������, ����� ��ȣ�� �Է��Ͽ� �ֽñ� �ٶ��ϴ�."
		ChgCont = ChgCont + "	<br><br><br>���� ���� ��ȭ ��ȣ�� ����Ǿ����ϴ�. (1644-1851)"
		ChgCont = ChgCont + "</td>"
	    ChgCont = ChgCont + "</tr>"
        ChgCont = ChgCont + "</table>"

		'�̸��� ���ø� ����
		'//�Ǽ�,�׼�����
		dim dfPath, fso, ffso, ChgCont
		IF application("Svr_Info")="Dev" THEN
			dfPath = Server.MapPath("\lib\email\mailtemplate")				'// �Ǽ�(scm)
		ELSE
		    dfPath = Server.MapPath("\lib\email\mailtemplate")				'// �Ǽ�(scm)
		END IF

		'/* ���� �ҷ����� */
		Set fso = server.CreateObject("Scripting.FileSystemObject")
			'IF fso.FileExists(dfPath & "\mail_u01.htm") then
				'set ffso = fso.OpenTextFile(dfPath & "\mail_u01.htm",1)
			IF fso.FileExists(dfPath & "\mail_basic.html") then
				set ffso = fso.OpenTextFile(dfPath & "\mail_basic.html",1)
				mailcontent = ffso.ReadAll
				ffso.close
				set ffso = nothing
			ELSE
				mailcontent = ""
			End IF
		Set fso = nothing

		mailcontent = Replace(mailcontent,":HTMLTITLE:",mailtitle)			'���� Ÿ��Ʋ
		mailcontent = Replace(mailcontent,":CONTENTSHTML:",ChgCont)	'���� ����

		set oMail = New MailCls

		IF mailto<>"" THEN

			oMail.MailTitles	= mailtitle
			oMail.SenderNm		= "�ٹ�����"
			'oMail.SenderMail	= "mailzine@10x10.co.kr"
			oMail.SenderMail	= "customer@10x10.co.kr"
			oMail.AddrType		= "string"
			oMail.ReceiverNm	= mailto
			oMail.ReceiverMail	= mailto
			oMail.MailConts 	= mailcontent
			oMail.MailerMailGubun = 13		' ���Ϸ� �ڵ����� ��ȣ

			oMail.Send_TMSMailer()		'TMS���Ϸ�
			'oMail.Send_Mailer()		'EMS���Ϸ�
			''oMail.Send_CDO()	'cdo
		End IF

		SET oMail = nothing
	end if

	sqlstr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlstr = sqlstr + " set sendsms='Y'" + VbCrlf
	sqlstr = sqlstr + " where idx=" + Cstr(idx) + VbCrlf
	dbget.execute sqlstr

	response.write "<script>alert('���۵Ǿ����ϴ�.');</script>"
	response.write "<script>window.close();</script>"
	dbget.close()	:	response.End
end if
%>
<script type='text/javascript'>

function CopyInfo(ihp,iemail){
	document.frm.reqhp.value = ihp;
	document.frm.mailto.value = iemail;
}

function SendSMS(frm){
<% if osheet.FOneItem.IsSendedSMS then %>
	if (!confirm('�̹� ���۵� �ֹ� �Դϴ�. �� ���� �Ͻðڽ��ϱ�?')){ return };
<% end if %>

    if (frm.reqhp.value.length>15){
        alert('�ڵ��� ��ȣ�� 15�� ���Ϸ� �Է��ϼ���.\n�ڵ��� ��ȣ�� ��ü�������� ���� �����մϴ�.');
        frm.reqhp.focus();
		return;
    }

	if ((frm.reqhp.value.length<1)&&(frm.mailto.value.length<1)){
		alert('�ڵ��� ��ȣ�� �̸����ּ� �� �ϳ��� �ԷµǾ�� �մϴ�.');
		return;
	}

	var ret= confirm('���� �Ͻðڽ��ϱ�?');
	if(ret){
		frm.submit();
	}
}

</script>
<table width="500" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td colspan=5><%= opartner.FPartnerList(0).FCompany_name %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=5>[<%= opartner.FPartnerList(0).Fzipcode %>] <%= opartner.FPartnerList(0).Faddress %> <%= opartner.FPartnerList(0).Fmanager_address %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=5>��ǥ��ȭ : <%= opartner.FPartnerList(0).Ftel %> �ѽ� : <%= opartner.FPartnerList(0).Ffax %></td>
</tr>

<tr bgcolor="#DDDDFF">
	<td width=80>����</td>
	<td width=80>����</td>
	<td width=80>��ȭ</td>
	<td width=80>�ڵ���</td>
	<td width=*>�̸���</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td ><a href="#" onClick="CopyInfo('<%= ogroup.FOneItem.Fmanager_hp %>','<%= ogroup.FOneItem.Fmanager_email %>');">�׷�����</a></td>
	<td ><%= ogroup.FOneItem.Fmanager_name %></td>
	<td ><%= ogroup.FOneItem.Fmanager_phone %></td>
	<td ><%= ogroup.FOneItem.Fmanager_hp %></td>
	<td ><%= ogroup.FOneItem.Fmanager_email %></td>
</tr>
<!-- ��۴���ڴ� �귣�庰
<tr bgcolor="#FFFFFF">
	<td ><a href="#" onClick="CopyInfo('<%= ogroup.FOneItem.Fdeliver_hp %>','<%= ogroup.FOneItem.Fdeliver_email %>');">�׷��۴����</a></td>
	<td ><%= ogroup.FOneItem.Fdeliver_name %></td>
	<td ><%= ogroup.FOneItem.Fdeliver_phone %></td>
	<td ><%= ogroup.FOneItem.Fdeliver_hp %></td>
	<td ><%= ogroup.FOneItem.Fdeliver_email %></td>
</tr>
 -->

<tr bgcolor="#FFFFFF">
	<td ><a href="#" onClick="CopyInfo('<%= opartner.FPartnerList(0).Fmanager_hp %>','<%= opartner.FPartnerList(0).Femail %>');">�����</a></td>
	<td ><%= opartner.FPartnerList(0).Fmanager_name %></td>
	<td ><%= opartner.FPartnerList(0).Fmanager_phone %></td>
	<td ><%= opartner.FPartnerList(0).Fmanager_hp %></td>
	<td ><%= opartner.FPartnerList(0).Femail %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td ><a href="#" onClick="CopyInfo('<%= opartner.FPartnerList(0).Fdeliver_hp %>','<%= opartner.FPartnerList(0).Fdeliver_email %>');">�귣���۴����</a></td>
	<td ><%= opartner.FPartnerList(0).Fdeliver_name %></td>
	<td ><%= opartner.FPartnerList(0).Fdeliver_phone %></td>
	<td ><%= opartner.FPartnerList(0).Fdeliver_hp %></td>
	<td ><%= opartner.FPartnerList(0).Fdeliver_email %></td>
</tr>
</table>

<form name="frm" method=post action="" style="margin:0px;">
<input type="hidden" name="mode" value="sendall">
<input type="hidden" name="idx" value="<%= osheet.FOneItem.Fidx %>">
<input type="hidden" name="mailfrom" value="<%= session("ssBctEmail") %>">
<table width="500" cellspacing="1" class="a" bgcolor=#FFFFFF cellpadding="2">
<tr>
    <td colspan="2">
        ** ��� ����� ����ó�� <strong>�귣�庰</strong>�� ����Ǿ����ϴ�.
    </td>
</tr>
<tr>
	<td width=100>�ڵ���</td>
	<td><input type="text" name="reqhp" value="<%= opartner.FPartnerList(0).Fdeliver_hp %>" size=16 maxlength=16></td>
</tr>
<tr>
	<td width=100>�̸���</td>
	<td><input type="text" name="mailto" value="<%= opartner.FPartnerList(0).Fdeliver_email %>" size=30 maxlength=80></td>
</tr>
<tr>
	<td width=100>SMS����</td>
	<td>
	<textarea name="smstext" cols=60 rows=3>[�ٹ�����]<%= opartner.FRectDesignerID %>�ֹ����� �����Ǿ����ϴ�.�ֹ����������� Ȯ���� �԰����ּ���.</textarea>
	</td>
</tr>
<tr>
	<td colspan="2" align=center><input type="button" value=" �� �� " onclick="SendSMS(frm);"></td>
</tr>
</table>
</form>

<% if osheet.FOneItem.IsSendedSMS then %>
	<script type='text/javascript'>
		alert('�̹� SMS���۵� �ֹ��Դϴ�.');
	</script>
<% end if %>

<%
set opartner = Nothing
set ogroup = Nothing
set osheet = Nothing
set osheetdetail = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
