<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� ��� �ֹ� �̸��� & ���� �߼�
' History : 2012.05.10 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->
<!-- #include virtual="/lib/email/maillib.asp"-->
<!-- #include virtual="/lib/email/smsLib.asp"-->
<%
dim opartner,i,page ,makerid ,ogroup ,osheet ,mode, mailfrom, reqhp, mailto, smstext ,groupid
dim mailtitle,mailcontent ,selltotal, buytotal ,sqlstr ,orderno , masteridx , detailidx
dim shopid ,shopphone
	page    = requestCheckVar(request("page"),10)
	makerid = requestCheckVar(request("makerid"),32)
	mode        = request("mode")
	mailfrom    = request("mailfrom")
	mailto	    = request("mailto")
	reqhp 	    = request("reqhp")
	smstext     = request("smstext")
	orderno     = requestCheckVar(request("orderno"),16)
	masteridx     = requestCheckVar(request("masteridx"),10)
	detailidx     = requestCheckVar(request("detailidx"),10)
	shopphone = requestcheckvar(request("shopphone"),32)
	
if page="" then page=1
	
set opartner = new CPartnerUser
	opartner.FCurrpage = page
	opartner.FRectDesignerID = makerid
	opartner.FPageSize = 1
	opartner.GetPartnerNUserCList

if opartner.FTotalCount > 0 then
	groupid=opartner.FPartnerList(0).FGroupid
else
	response.write "<script language='javascript'>"
	response.write "	alert('�ش� �귣�� ������ �����ϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	response.end	:	dbget.close()
end if

set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = groupid
	
	if groupid <> "" then
		ogroup.GetOneGroupInfo
	end if

set osheet = new cupchebeasong_list
	osheet.FRectorderno = orderno
	osheet.FRectmasteridx = masteridx
	'osheet.FRectdetailidx = detailidx
	osheet.FRectmakerid = makerid
	osheet.FRectIsUpcheBeasong = "Y"
	osheet.fbeasongsmslist

if opartner.FTotalCount < 1 then
	response.write "<script language='javascript'>"
	response.write "	alert('�ش� ��� ������ �����ϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	response.end	:	dbget.close()
else
	shopid = osheet.FItemList(0).fshopid
	
	if shopphone = "" then
		shopphone = osheet.FItemList(0).fshopphone
	end if
end if
shopphone = "1644-6030"

mailtitle = "[�ٹ�����] " + opartner.FRectDesignerID + " �귣���� �������� ����� (" + cstr(osheet.fitemlist(0).forderno) + ")�� �����Ǿ����ϴ�."

selltotal =0
buytotal = 0

if mode="sendall" then
	if reqhp<>"" then
'		sqlStr = "Insert into [db_sms].[ismsuser].em_tran(tran_phone, tran_callback, tran_status, tran_date, tran_msg ) "
'		sqlStr = sqlStr + " values('" + reqhp + "',"
'		sqlStr = sqlStr + " '"&shopphone&"',"
'		sqlStr = sqlStr + " '1',"
'		sqlStr = sqlStr + " getdate(),"
'		sqlStr = sqlStr + " '" + smstext + "')"
'
'		'response.write sqlStr &"<br>"
'		dbget.execute sqlStr
		Call SendNormalSMS_LINK(reqhp, shopphone, smstext)
	end if

	if mailto<>"" then
		
		'��ǰ����Ʈ �������� ���Ϲ߼� ���� �ۼ�
		ChgCont =""
		ChgCont = ChgCont + "<table width='600' border='0' align='center' cellpadding='0' cellspacing='0' class='a'>"
	    ChgCont = ChgCont + "<tr height='25' valign='top'>"
		ChgCont = ChgCont + "<td>"
		ChgCont = ChgCont + "<font color='red'><strong>�ֹ���</strong>&nbsp;<b>[" + opartner.FRectDesignerID + "]</b>&nbsp;&nbsp;�ֹ���ȣ(" + cstr(osheet.Fitemlist(0).forderno) + ")</font></td>"
	    ChgCont = ChgCont + "</tr>"
	    ChgCont = ChgCont + "<tr valign='top'>"
		ChgCont = ChgCont + "<td>"
		ChgCont = ChgCont + "	<br>�ȳ��ϼ���. �ٹ����� �Դϴ�."
		ChgCont = ChgCont + "	<br>�������� ���� <b>�����ް���>>*[������]��ۿ�û����Ʈ </b>���� �ֹ�Ȯ�� ��Ź�帳�ϴ�."
		ChgCont = ChgCont + "	<br>"
		ChgCont = ChgCont + "	<br>�귣��ID :<b>" + opartner.FRectDesignerID + "</b>"
		ChgCont = ChgCont + "	<br>�ֹ���ȣ :<b>" + cstr(osheet.Fitemlist(0).forderno) + "</b>"
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

	sqlstr = " update db_shop.dbo.tbl_shopbeasong_order_detail" + VbCrlf
	sqlstr = sqlstr + " set upchesendsms='Y'" + VbCrlf
	sqlstr = sqlstr + " where isupchebeasong='Y'"
	sqlstr = sqlstr + " and makerid='"&makerid&"'"
	sqlstr = sqlstr + " and masteridx="&masteridx&""
	
	'response.write sqlstr &"<Br>"
	dbget.execute sqlstr

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
<% if osheet.Fitemlist(0).fupchesendsms = "Y" then %>
	if (!confirm('�̹� �̸���&���ڰ� �߼۵� �귣�� �Դϴ�. �� ���� �Ͻðڽ��ϱ�?')){ return };
<% end if %>
    
    if (frm.reqhp.value.length>15){
        alert('�޴��� ��ȣ�� 15�� ���Ϸ� �Է��ϼ���.\n�ڵ��� ��ȣ�� ��ü�������� ���� �����մϴ�.');
        frm.reqhp.focus();
		return;
    }

    if (frm.shopphone.value.length>15){
        alert('ȸ�Ź�ȣ�� 15�� ���Ϸ� �Է��ϼ���.\n���� ��ȭ��ȣ�� ������������ ���� �����մϴ�.');
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
<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
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

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		** ��� ����� ����ó�� <strong>�귣�庰</strong>�� ����Ǿ����ϴ�.
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" cellspacing="1" class="a" bgcolor=#FFFFFF cellpadding="2">
<form name="frm" method=post action="">
<input type="hidden" name="mode" value="sendall">
<input type="hidden" name="orderno" value="<%= osheet.Fitemlist(0).forderno %>">
<input type="hidden" name="masteridx" value="<%= osheet.Fitemlist(0).fmasteridx %>">
<input type="hidden" name="makerid" value="<%= osheet.Fitemlist(0).fmakerid %>">
<input type="hidden" name="mailfrom" value="<%= session("ssBctEmail") %>">
<tr>
	<td width=100>�߼��޴���</td>
	<td><input type="text" name="reqhp" value="<%= opartner.FPartnerList(0).Fdeliver_hp %>" size=16 maxlength=16></td>
</tr>
<tr>
	<td>ȸ����ȭ��ȣ</td>
	<td>
		<input type="text" name="shopphone" readonly value="<%= shopphone %>" size=16 maxlength=16>
	</td>
</tr>
<tr>
	<td>�߼��̸���</td>
	<td><input type="text" name="mailto" value="<%= opartner.FPartnerList(0).Fdeliver_email %>" size=30 maxlength=80></td>
</tr>
<tr>
	<td>SMS����</td>
	<td>
		<textarea name="smstext" cols=60 rows=3>[�ٹ�����] <%= opartner.FRectDesignerID %> �������� �������. ����������>>*[������]��ۿ�û����Ʈ</textarea>
	</td>
</tr>
<tr>
	<td colspan="2" align="center"><input type="button" value="����" onclick="SendSMS(frm);" class="button"></td>
</tr>
</form>
</table>

<% if osheet.Fitemlist(0).fupchesendsms = "Y" then %>
	<script>alert('�̹� �̸���&���ڰ� �߼۵� �귣�� �Դϴ�.');</script>
<% end if %>

<%
set opartner = Nothing
set ogroup = Nothing
set osheet = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->