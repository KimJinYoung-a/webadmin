<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� �ֹ� �̸��� & ���� �߼�
' History : 2011.05.16 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<!-- #include virtual="/lib/email/maillib.asp"-->
<%
dim opartner,i,page ,designer ,idx ,ogroup ,osheet ,mode, mailfrom, reqhp, mailto, smstext
dim mailtitle,mailcontent ,selltotal, buytotal ,sqlstr
	page    = requestCheckVar(request("page"),10)
	designer = requestCheckVar(request("designer"),32)
	idx     = requestCheckVar(request("idx"),10)
	mode        = requestCheckVar(request("mode"),32)
	mailfrom    = requestCheckVar(request("mailfrom"),128)
	mailto	    = requestCheckVar(request("mailto"),128)
	reqhp 	    = requestCheckVar(request("reqhp"),16)
	smstext     = request("smstext")

if page="" then page=1
	
set opartner = new CPartnerUser
	opartner.FCurrpage = page
	opartner.FRectDesignerID = designer
	opartner.FPageSize = 1
	opartner.GetPartnerNUserCList

Dim groupid : groupid=opartner.FPartnerList(0).FGroupid

set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = groupid
	ogroup.GetOneGroupInfo

set osheet = new CShopIpChul
	osheet.FRectIdx = idx
	osheet.GetOneIpChulMaster

mailtitle = "[�ٹ�����] " + opartner.FRectDesignerID + " �귣���� �������� �ֹ��� (" + cstr(osheet.FOneItem.fidx) + ")�� �����Ǿ����ϴ�."

selltotal =0
buytotal = 0

if mode="sendall" then
	if reqhp<>"" then
		if smstext <> "" then
			if checkNotValidHTML(smstext) then
			response.write "<script type='text/javascript'>"
			response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
			response.write "</script>"
			dbget.close()	:	response.End
			end if
		end if

		'sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
		'sqlStr = sqlStr + " values('" + reqhp + "',"
		'sqlStr = sqlStr + " '070-7515-5421',"
		'sqlStr = sqlStr + " '1',"
		'sqlStr = sqlStr + " getdate(),"
		'sqlStr = sqlStr + " '" + smstext + "')"

		sqlStr = " exec [db_sms].[dbo].[usp_SendSMS] '"+reqhp+"','070-7515-5421','"+html2db(smstext)+"'"
		dbget.execute sqlStr
	end if

	if mailto<>"" then
		
		'��ǰ����Ʈ �������� ���Ϲ߼� ���� �ۼ�
		ChgCont =""
		ChgCont = ChgCont + "<table width='600' border='0' align='center' cellpadding='0' cellspacing='0' class='a'>"
	    ChgCont = ChgCont + "<tr height='25' valign='top'>"
		ChgCont = ChgCont + "<td>"
		ChgCont = ChgCont + "<font color='red'><strong>�ֹ���</strong>&nbsp;<b>[" + opartner.FRectDesignerID + "]</b>&nbsp;&nbsp;�����ڵ�(" + cstr(osheet.FOneItem.fidx) + ")</font></td>"
	    ChgCont = ChgCont + "</tr>"
	    ChgCont = ChgCont + "<tr valign='top'>"
		ChgCont = ChgCont + "<td>"
		ChgCont = ChgCont + "	<br>�ȳ��ϼ���. �ٹ������Դϴ�."
		ChgCont = ChgCont + "	<br>���� <b>����������>>�������Ʈ </b>���� �ֹ�Ȯ�� ��Ź�帳�ϴ�."
		ChgCont = ChgCont + "	<br>"
		ChgCont = ChgCont + "	<br>�귣��ID :<b>" + opartner.FRectDesignerID + "</b>"
		ChgCont = ChgCont + "	<br>�����ڵ� :<b>" + cstr(osheet.FOneItem.fidx) + "</b>"
		ChgCont = ChgCont + "	<br>"
		ChgCont = ChgCont + "	<br><b><font color='red'>[�ֹ�Ȯ��]</font></b>"
		ChgCont = ChgCont + "	<br>�ֹ����� Ȯ���Ͻ� �Ŀ���"
		ChgCont = ChgCont + "	<br>������ �����ϰų� ������ ���, �������������� ������ �ֽðų�,"
		ChgCont = ChgCont + "	<br>�԰�Ȯ�������� �����Ͽ��ֽñ� �ٶ��ϴ�."
		ChgCont = ChgCont + "	<br>"
		ChgCont = ChgCont + "	<br><b><font color='red'>[���Ϸ�]</font></b>"
		ChgCont = ChgCont + "	<br>����ϽǶ��� ���������� �������, Ȯ�������� �����Ͽ� �ֽðų�,"
		ChgCont = ChgCont + "	<br>�������������� ������ �ּ���."
		ChgCont = ChgCont + "</td>"
	    ChgCont = ChgCont + "</tr>"
        ChgCont = ChgCont + "</table>"

		'�̸��� ���ø� ����
		'//�Ǽ�,�׼�����
		dim dfPath, fso, ffso, ChgCont
		IF application("Svr_Info")="Dev" THEN
			dfPath = Server.MapPath("\lib\email\mailtemplate") 		'// �׼�(scm)
		ELSE
		    dfPath = Server.MapPath("\lib\email\mailtemplate")				'// �Ǽ�(scm)
		END IF

		'/* ���� �ҷ����� */
		Set fso = server.CreateObject("Scripting.FileSystemObject")
			IF fso.FileExists(dfPath & "\mail_u01.htm") then
				set ffso = fso.OpenTextFile(dfPath & "\mail_u01.htm",1)
				mailcontent = ffso.ReadAll
				ffso.close
				set ffso = nothing
			ELSE
				mailcontent = ""
			End IF
		Set fso = nothing

		mailcontent = Replace(mailcontent,":HTMLTITLE:",mailtitle)			'���� Ÿ��Ʋ
		mailcontent = Replace(mailcontent,":CONTENTSHTML:",ChgCont)	'���� ����

		'// ���� �߼�
		call sendmail(mailfrom, mailto, mailtitle, mailcontent)
	end if

	sqlstr = " update [db_shop].dbo.tbl_shop_ipchul_master" + VbCrlf
	sqlstr = sqlstr + " set sendsms='Y'" + VbCrlf
	sqlstr = sqlstr + " where idx=" + Cstr(idx) + VbCrlf
	rsget.Open sqlStr,dbget,1

	response.write "<script>alert('���۵Ǿ����ϴ�.');</script>"
	response.write "<script>window.close();</script>"
	dbget.close()	:	response.End
end if
%>

<script language='javascript'>

function CopyInfo(ihp,iemail){
	document.frm.reqhp.value = ihp;
	document.frm.mailto.value = iemail;
}

function SendSMS(frm){
<% if osheet.FOneItem.fsendsms = "Y" then %>
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

<table width="500" cellspacing="1" class="a" bgcolor=#FFFFFF cellpadding="2">
<form name="frm" method=post action="">
<input type="hidden" name="mode" value="sendall">
<input type="hidden" name="idx" value="<%= osheet.FOneItem.Fidx %>">
<input type="hidden" name="mailfrom" value="<%= session("ssBctEmail") %>">
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
		<textarea name="smstext" cols=60 rows=3>[�ٹ�����]<%= opartner.FRectDesignerID %>�������� �ֹ��� ����. ����������>>�������Ʈ Ȯ�����ּ���.</textarea>
	</td>
</tr>
<tr>
	<td colspan="2" align="center"><input type="button" value="����" onclick="SendSMS(frm);" class="button"></td>
</tr>
</form>
</table>

<% if osheet.FOneItem.fsendsms = "Y" then %>
	<script>alert('�̹� SMS���۵� �ֹ��Դϴ�.');</script>
<% end if %>

<%
set opartner = Nothing
set ogroup = Nothing
set osheet = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->