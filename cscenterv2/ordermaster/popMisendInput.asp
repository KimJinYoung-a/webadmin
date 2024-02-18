<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/oldmisendcls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/order/upchebeasongcls.asp"-->
<style type="text/css" >
.sale11px01 {font-family: dotum; FONT-SIZE: 11px; font-weight:bold ; COLOR: #b70606;}
</style>
<%
'' XXXX �귣��/ ���� ������
'' �ΰŽ�
dim C_ADMIN_USER : C_ADMIN_USER = True

dim idx : idx= requestCheckVar(request("idx"),10)

dim omisend
set omisend = new COldMiSend
omisend.FRectDetailIDx = idx
omisend.getOneOldMisendItem

if (omisend.FResultCount<1) then
    response.write "�˻������ �����ϴ�."
    dbget.close() : response.end
end if

''��ü�ΰ��
if (Not C_ADMIN_USER) then
    if (LCase(omisend.FOneItem.FMakerid) <> LCASE(session("ssBctID"))) then
        response.write "������ �����ϴ�."
        dbget.close() : response.end
    end if
end if

dim PreDispMail
PreDispMail = (omisend.FOneItem.isMisendAlreadyInputed) and (omisend.FOneItem.FMisendReason<>"05")


dim MisendReasonStr : MisendReasonStr = "03,02,08,09,04,10,07"
dim MisendReasonArr : MisendReasonArr = Split(MisendReasonStr, ",")
dim i, tmpStr

%>
<script language='javascript'>

function SetStockOut() {
    var itemSoldOutFlag = document.all.itemSoldOutFlag;
	var itemSoldOutContent = document.all.itemSoldOutContent;
	var itemSoldOutButton = document.all.itemSoldOutButton;

	if (itemSoldOutFlag.style.display == "none") {
        itemSoldOutFlag.style.display = "inline";
		itemSoldOutContent.style.display = "inline";
		if (itemSoldOutButton) {
			itemSoldOutButton.disabled = false;
		}
	} else {
        itemSoldOutFlag.style.display = "none";
		itemSoldOutContent.style.display = "none";
		if (itemSoldOutButton) {
			itemSoldOutButton.disabled = true;
		}
	}
}

function ShowHideObject(comp) {
    var frm = comp.form;
	var doc = document.all;
	var tmpObj;

	// �԰�����
    var divipgodate = doc.divipgodate;

	// ǰ�����Ұ�
	var itemSoldOutFlag = doc.itemSoldOutFlag;
	var itemSoldOutContent = doc.itemSoldOutContent;

	// SMS/MAIL
	var SMSContentAll = doc.SMSContentAll;
	var MailContentAll = doc.MailContentAll;

	SMSContentAll.style.display = "none";
	MailContentAll.style.display = "none";

	<% for i = 0 to UBound(MisendReasonArr) %>
	doc.SMSContent<%= MisendReasonArr(i) %>.style.display = "none";
	doc.MailContent<%= MisendReasonArr(i) %>.style.display = "none";
	<% next %>

	if (divipgodate) {
		if (comp.value == "05") {
			divipgodate.style.display = "none";
		} else {
			divipgodate.style.display = "inline";
		}
	}

	if (comp.value == "05") {
		itemSoldOutFlag.style.display = "inline";
		itemSoldOutContent.style.display = "inline";

		SMSContentAll.style.display = "none";
		MailContentAll.style.display = "none";
	} else {
		itemSoldOutFlag.style.display = "none";
		itemSoldOutContent.style.display = "none";

		if (comp.value != "") {
			tmpObj = eval("doc.SMSContent" + comp.value);
			tmpObj.style.display = "inline";

			tmpObj = eval("doc.MailContent" + comp.value);
			tmpObj.style.display = "inline";

			SMSContentAll.style.display = "inline";
			MailContentAll.style.display = "inline";
		}
	}

	<% if (C_ADMIN_USER) then %>
		if (comp.value == "05") {
			frm.ckSendSMS.disabled = true;
			frm.ckSendEmail.disabled = true;
			frm.ckSendSMS.checked = false;
			frm.ckSendEmail.checked = false;
		} else {
			frm.ckSendSMS.disabled = false;
			frm.ckSendEmail.disabled = false;
			frm.ckSendSMS.checked = true;
			frm.ckSendEmail.checked = true;
		}
	<% end if %>
}

function MisendInput(){
    var frm = document.frmMisend;
    var today= new Date();
    //today = new Date(today.getYear(),today.getMonth(),today.getDate());  //���õ� �����ϵ���
    today = new Date(<%=year(now())%>,<%=month(now())-1%>,<%=Day(now())%>);  //2016/09/08 ����.

    var inputdate;

    if (frm.MisendReason.value.length<1){
        alert('����� ������ �Է��ϼ���.');
        frm.MisendReason.focus();
        return;
    }


    // ǰ�����Ұ�(05)
    if (frm.MisendReason.value != "05") {
        var ipgodate = eval("frm.ipgodate");
        if (ipgodate.value.length!=10){
            alert('��� �������� �Է��ϼ���.(YYYY-MM-DD)');
            ipgodate.focus();
            return;
        }

        inputdate = new Date(ipgodate.value.substr(0,4),ipgodate.value.substr(5,2)*1-1,ipgodate.value.substr(8,2));
        if (today>inputdate){
            alert('��� �������� ���� ���ĳ�¥�� ������ �����մϴ�.');
            ipgodate.focus();
            return;
        }

		/*
        if (frm.ckSendSMS && frm.ckSendEmail) {
        	if ((frm.ckSendSMS.checked != true) && (frm.ckSendEmail.checked != true)) {
				alert("SMS �� ���Ϲ߼� ���� �ϳ��� üũ�ؾ� �մϴ�.");
				return;
        	}
        }
        */
    } else {
		// ǰ����Ͻ� ���氡�� �ɼ� ����
		<% if (omisend.FOneItem.FItemoption <> "0000") then %>
		if (frm.reqaddstr) {
			var regExp = /���氡�� �ɼ� :[ \t]*\r?\n/;

			if(regExp.test(frm.reqaddstr.value) == true) {
				frm.reqaddstr.value = frm.reqaddstr.value.replace("���氡�� �ɼ� :", "���氡�� �ɼ� : ����");

				alert('���氡�� �ɼ��� �Է��ϼ���.\n\n==>>> ���Է½� \"����\" ���� �Էµ˴ϴ�. <<<==');
				frm.reqaddstr.focus();
				return;
			}
		}
		<% end if %>
	}

	if (frm.MisendReason.value != "05") {
		frm.reqaddstr.value = "";
	}

    if (confirm('����� ������ ���� �Ͻðڽ��ϱ�?')){
	    frm.action = "upchebeasong_Process.asp";
	    frm.submit();
	}
}

function MisendInputUpche() {
	var frm = document.frmMisend;

	if (confirm('ǰ����� �Ͻðڽ��ϱ�?')){
	    frm.action = "upchebeasong_Process.asp";
	    frm.submit();
	}
}

function SetupObject() {
	<% if (C_ADMIN_USER) then %>
		ShowHideObject(frmMisend.MisendReason);
	<% end if %>
	popupResize(680);
}
window.onload = SetupObject;

</script>

<% if omisend.FResultCount>0 then %>
<table width="610" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmMisend" method="post" action="upchebeasong_Process.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="misendInputOne">
	<input type="hidden" name="detailidx" value="<%= omisend.FOneItem.Fidx %>">
	<input type="hidden" name="Sitemid" value="<%= omisend.FOneItem.FItemID %>">
	<input type="hidden" name="Sitemoption" value="<%= omisend.FOneItem.FItemOption %>">
	<tr height="30" bgcolor="<%= adminColor("tabletop") %>">
	    <td colspan="2">
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�������� �Է�</b>
	    </td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
    	<td width="130">��ǰ�ڵ�</td>
    	<td width="480"><%= omisend.FOneItem.FItemID %>
    	    <% if (omisend.FOneItem.FCancelyn<>"N") then %>
				<b><font color="#CC3333">[����ֹ�]</font></b>
				<script language="javascript">alert("��ҵ� �ŷ� �Դϴ�.");</script>
			<% else %>
			    <% if (omisend.FOneItem.FDetailCancelYn="Y") then %>
				    <b><font color="#CC3333">[��һ�ǰ]</font></b>
			    <% else %>
				    [�����ֹ�]
			    <% end if%>
			<% end if %>
    	</td>
    </tr>
	<tr bgcolor="#FFFFFF">
	    <td>�̹���</td>
	    <td><img src="<%= omisend.FOneItem.Fsmallimage %>" width="50" height="50"></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
	    <td>��ǰ��</td>
	    <td><%= omisend.FOneItem.FItemName %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
	    <td>�ɼ�</td>
	    <td><%= omisend.FOneItem.FItemoptionName %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
	    <td>�ֹ�����</td>
	    <td>
			<%= omisend.FOneItem.FItemcnt %>��
			<% if (C_ADMIN_USER) then %>
				(�������� <%= omisend.FOneItem.Fitemlackno %>)
			<% end if %>
	    </td>
	</tr>

	<tr height="25" bgcolor="#FFFFFF">
	    <td>��������</td>
	    <td>
	        <% if (Not C_ADMIN_USER) and omisend.FOneItem.isMisendAlreadyInputed then %>
				<%= omisend.FOneItem.getMiSendCodeName %>
				<% if omisend.FOneItem.isMisendAlreadyInputed and (omisend.FOneItem.FMisendReason <> "05") then %>
					<input type="button" class="button" value="ǰ����ȯ" onClick="SetStockOut();">
					<input type="hidden" name="MisendReason" value="05">
				<% end if %>
	        <% else %>
				<select name="MisendReason" id="MisendReason" class="select" onChange="ShowHideObject(this);">
					<option value="">---------</option>
					<option value="03" <%= ChkIIF(omisend.FOneItem.FMisendReason="03","selected"," ") %> >�������</option>
					<option value="02" <%= ChkIIF(omisend.FOneItem.FMisendReason="02","selected"," ") %> >�ֹ�����</option>
					<option value="08" <%= ChkIIF(omisend.FOneItem.FMisendReason="08","selected"," ") %> >����</option>
					<option value="09" <%= ChkIIF(omisend.FOneItem.FMisendReason="09","selected"," ") %> >�������</option>
					<option value="04" <%= ChkIIF(omisend.FOneItem.FMisendReason="04","selected"," ") %> >������</option>
					<option value="10" <%= ChkIIF(omisend.FOneItem.FMisendReason="10","selected"," ") %> >��ü�ް�</option>
					<option value="07" <%= ChkIIF(omisend.FOneItem.FMisendReason="07","selected"," ") %> >���������</option>
					<option value="">---------</option>
					<option value="05" <%= ChkIIF(omisend.FOneItem.FMisendReason="05","selected"," ") %> >ǰ�����Ұ�</option>
					<option value="">---------</option>
				</select>
			<% end if %>
			<span id="itemSoldOutFlag" name="itemSoldOutFlag" style="display=none" align="right" >
			<input type="radio" name="itemSoldOut" value="N" checked >��ǰ ǰ��ó��
			<input type="radio" name="itemSoldOut" value="S">��ǰ �Ͻ�ǰ��ó��
			</span>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
	    <td>ǰ�����Ұ���<br>(���������޻���)</td>
	    <td>
			<span id="itemSoldOutContent" name="itemSoldOutContent" style="display=<% if (omisend.FOneItem.FupcheRequestString = "") then %>none<% else %>inline<% end if %>" align="right" >
			<textarea class="textarea" name="reqaddstr" cols="65" rows="9" <% if (Not C_ADMIN_USER) and omisend.FOneItem.isMisendAlreadyInputed then %>readonly<% end if %> ><% if (omisend.FOneItem.FupcheRequestString = "") then %>�������� : <%= omisend.FOneItem.FItemcnt %>��
����ȭ���� : N
���氡�� �ɼ� :
��Ÿ ���޻��� :
<% else %><%= omisend.FOneItem.FupcheRequestString %><% end if %></textarea>
			</span>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
	    <td>�������</td>
	    <td>
	        <% if (Not C_ADMIN_USER) and omisend.FOneItem.isMisendAlreadyInputed then %>
	        	<%= omisend.FOneItem.FMisendIpgodate %>
	        <% else %>
				<div id="divipgodate" name="divipgodate">
					<input class="text" type="text" name="ipgodate" value="<%= omisend.FOneItem.FMisendIpgodate %>" size="10" maxlength="10">
					<a href="javascript:calendarOpen(frmMisend.ipgodate);"><img src="/images/calicon.gif" border="0" align="top" height=20></a>
				</div>
			<% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td>���ȳ�����</td>
	    <td>
	        <% if (C_ADMIN_USER) then %>
				<input name="ckSendSMS" type="checkbox" checked  >SMS�߼�<% if (omisend.FOneItem.FisSendSms="Y") then %>(Y)<% end if %>
				&nbsp;
				<input name="ckSendEmail" type="checkbox" checked  >MAIL�߼�<% if (omisend.FOneItem.FisSendEmail="Y") then %>(Y)<% end if %>
	        <% else %>
    	        <% if omisend.FOneItem.isMisendAlreadyInputed then %>
    	            <%= CHKIIF(omisend.FOneItem.FisSendSms="Y","SMS�߼ۿϷ� &nbsp; ","") %>
    	            <%= CHKIIF(omisend.FOneItem.FisSendEmail="Y","MAIL�߼ۿϷ� &nbsp; ","") %>
    	            <%= CHKIIF(omisend.FOneItem.FisSendCall="Y","��ȭ�ȳ��Ϸ�","") %>
    	        	<!-- ���ȳ��� �Ϸ�� ���� �������� �� ������� ���� �Ұ� -->
    	        <% else %>
        	        <input name="ckSendSMS" type="checkbox" checked disabled >SMS�߼�
        	        &nbsp;
        	        <input name="ckSendEmail" type="checkbox" checked disabled >MAIL�߼�
    	        <% end if %>
    	    <% end if %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF">
	    <td colspan="2">
	    	<font color="blue">
	    	����� ������ ������� �� �ֹ�����(����)�� ���, �Ʒ��� �������� ���Բ� SMS�� ������ �߼۵˴ϴ�.<br>
	    	���Բ� �ȳ��� ��������� �� �����ֽñ� �ٶ��, ���������� ������, �����ͷ� ���� ��Ź�帳�ϴ�.<br>
	    	</font>
	    	<font color="red">
	       	ǰ�����Ұ��� ���, ���Բ� SMS �� ������ �߼۵��� ������, �ٹ����ٰ����Ϳ���<br>
	    	������ ���Բ� ������ �帱 �����Դϴ�.
	    	</font>
	    </td>
	</tr>
	<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	    <td colspan="2" align="center">
	    <% if (C_ADMIN_USER) then %>
	        <input type="button" class="button" value="����� ���� ����" onclick="MisendInput();">
	    <% else %>
    	    <% if omisend.FOneItem.isMisendAlreadyInputed then %>
				<% if omisend.FOneItem.isMisendAlreadyInputed and (omisend.FOneItem.FMisendReason <> "05") then %>
					<input type="button" class="button" id="itemSoldOutButton" name="itemSoldOutButton" value="ǰ�����" onClick="MisendInputUpche();" disabled><br><br>
					(�̿��� ���������� �����ͷ� �����ϼ���.)
				<% else %>
					���� �Ұ�
				<% end if %>
    	    <% else %>
    	    <input type="button" class="button" value="����� ���� ����" onclick="MisendInput();">
    	    <% end if %>
    	<% end if %>
	    </td>
	</tr>
	</form>
</table>

<p>

<!-- �������/�ֹ����� ���ý� �Ʒ� ���̴� �����Դϴ�. �������ý� �ǽð����� ���̵��� -->

<table width="610" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	    <td>
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>SMS �߼۳���</b>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" id="SMSContentAll" style="display:none" >
	    <td>
        	<table width="100%" align="center" cellspacing="1" cellpadding="0" class="a" >
			<%
			for i = 0 to UBound(MisendReasonArr)
				tmpStr = GetMichulgoSMSString(MisendReasonArr(i))
				tmpStr = Replace(tmpStr, "[�������]", "<span id='MaySendDate" + MisendReasonArr(i) + "' name='MaySendDate" + MisendReasonArr(i) + "'>" + CStr(CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD")) + "</span>")

				tmpStr = Replace(tmpStr, "[��ǰ��]", DdotFormat(omisend.FOneItem.FItemName,16))
				tmpStr = Replace(tmpStr, "[��ǰ�ڵ�]", omisend.FOneItem.FItemID)
			%>
			<tr bgcolor="#FFFFFF" id="SMSContent<%= MisendReasonArr(i) %>">
            	<td>
					<%= tmpStr %>
            	</td>
            </tr>
			<% next %>
            </table>
        </td>
    </tr>
</table>

<p>

<table width="610" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	    <td>
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>MAIL �߼۳���</b>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" id="MailContentAll" style="display:none">
    	<td>
    		<!-- ���� ���� ���� -->
    		<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td>

						<!-- ������ ���� -->
						<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="a">
						<tr>
							<td>
								<a href="http://www.thefingers.co.kr" target="_blank">
									<img src="http://image.thefingers.co.kr/2016/mail/img_logo.png" width="600" height="85" border="0" />
								</a>
							</td>
						</tr>
						<tr>
							<td style="border:7px solid #eeeeee;">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
								<tr>
									<td><img src="http://fiximage.10x10.co.kr/web2008/mail/b01_img.gif" width="586"> </td>
								</tr>
								<tr>
									<td height="30" style="padding:0 15px 0 15px">
										<!-- ���� / �ֹ���ȣ -->
										<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
										<tr>
											<td class="black12px">

											</td>
											<td align="right" class="gray11px02">�ֹ���ȣ : <span class="sale11px01"><%= omisend.FOneItem.FOrderserial %></span></td>
										</tr>
										<tr>
											<td height="3" colspan="2" class="black12px" style="padding:5px;" bgcolor="#99CCCC"></td>
										</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td style="padding:5px 15px 20px 15px">
										<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
										<%
										for i = 0 to UBound(MisendReasonArr)
											tmpStr = GetMichulgoMailString(MisendReasonArr(i))
										%>
										<tr bgcolor="#FFFFFF" id="MailContent<%= MisendReasonArr(i) %>">
											<td>
												<%= Replace(tmpStr, "\n", "<br>") %>
											</td>
										</tr>
										<% next %>

										<tr>
											<td>

												<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
												<tr>
													<td colspan="2" class="sky12pxb" style="padding: 10 0 5 0">*��ǰ����</td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td width="150" height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" align="center" style="padding-top:2px;">��ǰ</td>
													<td width="450"class="gray12px02" style="padding-left:10px;padding-top:2px;"><img src="<%= omisend.FOneItem.Fsmallimage %>" width="50" height="50"></td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" align="center" style="padding-top:2px;">��ǰ�ڵ�</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><%= omisend.FOneItem.FItemID %> </td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" style="padding-top:2px;">��ǰ��</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><%= omisend.FOneItem.FItemName %></td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" style="padding-top:2px;">�ɼǸ�</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><%= omisend.FOneItem.FItemoptionName %></td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" style="padding-top:2px;">�ֹ�����</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><%= omisend.FOneItem.FItemcnt %>��</td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td colspan="2" class="sky12pxb" style="padding: 20 0 5 0">*�߼ۿ����ȳ�</td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" style="padding-top:2px;">�߼�(�Ǹ�)��</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><b><%= omisend.FOneItem.getDlvCompanyName %></b></td>
													<!-- �ٹ����� ����� ��� �ٹ����� ��������, ��ü�ϰ��, ��üȸ���-->
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" style="padding-top:2px;">�߼ۿ�����</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><b><span id="iMisendIpgodate2" name="iMisendIpgodate2"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span></b></td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr id="iEMAILMENTNOTI1" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="07","none","inline") %>">
													<td colspan="2" class="gray12px02" style="padding: 5 0 5 0">
													* �߼ۿ����Ϸκ��� 1~2�� �Ŀ� ��ǰ�� �޾ƺ��� �� �ֽ��ϴ�.<br>
													</td>
												</tr>

												</table>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td><img src="http://image.thefingers.co.kr/academy2009/mail/mail_bottom.gif" width="600" height="30" /></td>
							</tr>
							<tr>
								<td height="51" style="border-bottom:1px solid #eaeaea;">
									<table width="100%" border="0" cellspacing="0" cellpadding="0">
									<tr>
										<td style="padding-left:20px;"><img src="http://image.thefingers.co.kr/academy2009/mail/bottom_text.gif" width="245" height="26" /></td>
										<td width="128"><a href="http://www.thefingers.co.kr/cscenter/csmain.asp" onFocus="blur()" target="_blank"><img src="http://image.thefingers.co.kr/academy2009/mail/btn_cscenter.gif" width="108" height="31" border="0" /></a></td>
									</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td style="padding:10px 0 15px 0;line-height:17px;" class="gray11px02" class="a">
									(03086) ����� ���α� ���з�12�� 31 �������� 5�� (��)�ٹ����� theFingers<br>
									��ǥ�̻� : ������  &nbsp;����ڵ�Ϲ�ȣ : 211-87-00620  &nbsp;����Ǹž� �Ű��ȣ : �� 01-1968ȣ  &nbsp;�������� ��ȣ �� û�ҳ� ��ȣå���� : �̹���<br>
									<span class="black11px">���ູ����:TEL 1644-1557  &nbsp;E-mail:<a href="mailto:customer@thefingers.co.kr" class="link_black11pxb">customer@thefingers.co.kr</a> </span>
								</td>
							</tr>
							</table>
						<!-- ������ �� -->
					</td>
				</tr>
			</table>

    		<!-- ���� ���� �� -->
    	</td>
    </tr>
</table>


<% else %>
<table width="600">
<tr>
    <td align="center">��ҵ� ��ǰ�̰ų� �ش� �ֹ� ������ �����ϴ�.</td>
</tr>
</table>
<% end if %>

<%
set omisend = Nothing
%>
<!-- #include virtual="/cscenterv2/lib/poptail.asp"-->
<!-- #include virtual="/cscenterv2/lib/db/dbclose.asp" -->
