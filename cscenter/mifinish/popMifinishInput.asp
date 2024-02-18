<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : [CS]������>>[CS]��ó��CS����Ʈ
' History : �̻� ����
'			2023.11.15 �ѿ�� ����(�����ϴ� ����ü���� �������� cs������ ���� �̰�)
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_mifinishcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%
dim csdetailidx, ocsmifinishmaster, PreDispMail, isChulgoState, ioneas
    csdetailidx = requestCheckVar(request("csdetailidx"),10)

set ocsmifinishmaster = new CCSMifinishMaster
    ocsmifinishmaster.FRectCSDetailIDx = csdetailidx
    ocsmifinishmaster.getOneMifinishItem

	if ocsmifinishmaster.FtotalCount < 1 then
		ocsmifinishmaster.FRectCSDetailIDx = csdetailidx
		ocsmifinishmaster.FRectorder6MonthBefore = "Y"
		ocsmifinishmaster.getOneMifinishItem
	end if

if (ocsmifinishmaster.FResultCount<1) then
    response.write "�˻������ �����ϴ�."
    dbget.close() : response.end
end if

PreDispMail = (ocsmifinishmaster.FOneItem.isMifinishAlreadyInputed) and (ocsmifinishmaster.FOneItem.FMifinishReason<>"05")
isChulgoState = (ocsmifinishmaster.FOneItem.Fdivcd = "A000") or (ocsmifinishmaster.FOneItem.Fdivcd = "A100")

set ioneas = new CCSASList
    ioneas.FRectCsAsID = ocsmifinishmaster.FOneItem.Fasid
    ioneas.GetOneCSASMaster

%>
<style type="text/css" >
.sale11px01 {font-family: dotum; FONT-SIZE: 11px; font-weight:bold ; COLOR: #b70606;}
</style>
<script language='javascript'>
//function getOnload(){
//    popupResize(700);
//}
//window.onload = getOnload;

function ShowDateBox(comp){
    var frm = comp.form;

    var iid = comp.id;
    var idiv = document.all.divipgodate;
    var isold = document.all.itemSoldOutFlag

	if (comp.value == "05") {
		// ǰ�����Ұ�
		idiv.style.display = "none";
		isold.style.display = "inline";
	} else {
		idiv.style.display = "inline";
		isold.style.display = "none";
	}
}

function ipgodateChange(comp){
    var v = comp.value;
    if (v.length<10) v = "YYYY-MM-DD";

    ShowDateBox(frmMisend.MifinishReason);
}

function MiFinishInput(){
    var frm = document.frmMisend;
    var today= new Date();
    today = new Date(today.getYear(),today.getMonth(),today.getDate());  //���õ� �����ϵ���

    var inputdate;

    if (frm.MifinishReason.value.length<1){
        alert('��ó�� ������ �Է��ϼ���.');
        frm.MifinishReason.focus();
        return;
    }

    if (frm.MifinishReason.value == "05") {

    } else {
        var ipgodate = eval("frm.ipgodate");
        if (ipgodate.value.length!=10){
            alert('ó�� �������� �Է��ϼ���.(YYYY-MM-DD)');
            ipgodate.focus();
            return;
        }

        inputdate = new Date(ipgodate.value.substr(0,4),ipgodate.value.substr(5,2)*1-1,ipgodate.value.substr(8,2));
        if (today>inputdate){
            alert('ó�� �������� ���� ���ĳ�¥�� ������ �����մϴ�.');
            ipgodate.focus();
            return;
        }
    }

    if (confirm('��ó�� ������ ���� �Ͻðڽ��ϱ�?')){
	    frm.action = "/cscenter/mifinish/popMifinishInput_process.asp";
	    frm.submit();
	}
}

function getDiffDay(d1,d2){   // �� ��¥�� ���̱���

  var v1=d1.split("-");
  var v2=d2.split("-");

  var a1=new Date(v1[0],v1[1],v1[2]);
  var a2=new Date(v2[0],v2[1],v2[2]);
  return parseInt((a2-a1)/(1000*3600*24));  //1000*3600*24 �� �������� ���� ���� ���̸� ���ϰ� �ʹٸ� *30���ϸ� �� 12�� ���ϸ� ��

}

</script>

<% if ocsmifinishmaster.FResultCount>0 then %>
<form name="frmMisend" method="post" action="/cscenter/mifinish/popMifinishInput_process.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="MiFinishInputOne">
<input type="hidden" name="csdetailidx" value="<%= ocsmifinishmaster.FOneItem.Fcsdetailidx %>">
<input type="hidden" name="asid" value="<%= ocsmifinishmaster.FOneItem.Fasid %>">
<input type="hidden" name="Sitemid" value="<%= ocsmifinishmaster.FOneItem.FItemID %>">
<input type="hidden" name="Sitemoption" value="<%= ocsmifinishmaster.FOneItem.FItemOption %>">
<input type="hidden" name="ischulgostate" value="<% if isChulgoState then %>Y<% end if %>">
<table width="610" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	    <td colspan="2">
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS��ó������ �Է�</b>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
    	<td width="130">����</td>
    	<td width="480">
    		<font color="<%= ocsmifinishmaster.FOneItem.getDivcdColor %>"><%= ocsmifinishmaster.FOneItem.getDivcdStr %></font>
    	</td>
    </tr>
	<tr bgcolor="#FFFFFF" height="25">
    	<td width="130">��ǰ�ڵ�</td>
    	<td width="480"><%= ocsmifinishmaster.FOneItem.FItemID %>

    	    <% if (ocsmifinishmaster.FOneItem.Fdeleteyn<>"N") then %>
				<b><font color="#CC3333">[���CS]</font></b>
				<script language='javascript'>alert('��ҵ� CS �Դϴ�.');</script>
			<% else %>
			    [����CS]
			<% end if %>
    	</td>
    </tr>
	<tr bgcolor="#FFFFFF">
	    <td>�̹���</td>
	    <td><img src="<%= ocsmifinishmaster.FOneItem.Fsmallimage %>" width="50" height="50"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
	    <td>��ǰ��</td>
	    <td><%= ocsmifinishmaster.FOneItem.FItemName %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
	    <td>�ɼ�</td>
	    <td><%= ocsmifinishmaster.FOneItem.FItemoptionName %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
	    <td>��������</td>
	    <td><%= ocsmifinishmaster.FOneItem.FRegItemNo %>��
	    <% if (isChulgoState = True) then %>
		    <% if ( C_ADMIN_USER) then %>
		    (�������� <%= ocsmifinishmaster.FOneItem.Fitemlackno %>)
		    <% end if %>
		<% end if %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
	    <td>CS����</td>
	    <td>
	    	<%= ioneas.FOneItem.FTitle %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
	    <td>��������</td>
	    <td>
	    	<%= ioneas.FOneItem.Fgubun01Name %>&gt;&gt;<%= ioneas.FOneItem.Fgubun02Name %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
	    <td>��������</td>
	    <td>
	    	<%= replace(ioneas.FOneItem.Fcontents_jupsu,VbCrlf,"<br>") %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
	    <td>��ó������</td>
	    <td>
	        <% if (Not C_ADMIN_USER) and ocsmifinishmaster.FOneItem.isMifinishAlreadyInputed then %>
	        	<%= ocsmifinishmaster.FOneItem.getMiFinishCodeName %>
	        <% else %>
	        <select name="MifinishReason" id="MifinishReason" class="select" onChange="ShowDateBox(this);">
				<option value=""></option>
				<% if (isChulgoState = True) then %>
					<option value="03" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="03","selected"," ") %> >�������</option>
					<option value="05" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="05","selected"," ") %> >ǰ�����Ұ�</option>
					<option value="02" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="02","selected"," ") %> >�ֹ�����(����)</option>
					<option value="04" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="04","selected"," ") %> >������</option>
					<option value="07" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="07","selected"," ") %> >���������</option>
				<% else %>
					<% if (C_ADMIN_USER) then %>
						<option value="25" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="25","selected"," ") %> >�����Է� �ȳ�</option>
						<option value="26" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="26","selected"," ") %> >��ǰ�Ұ� �ȳ�</option>
						<option value="21" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="21","selected"," ") %> >�� ����</option>
						<option value="22" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="22","selected"," ") %> >�� ��ǰ����</option>
						<option value="23" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="23","selected"," ") %> >CS�ù�����</option>
						<option value="12" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="12","selected"," ") %> >��ü����</option>
						<option value="41" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="41","selected"," ") %> >�ù�� ��������</option>
					<% else %>
						<!--
						<option value="11" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="11","selected"," ") %> >��ǰ ȸ������</option>
						<option value="13" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="13","selected"," ") %> >������û(�� ���Է�)</option>
						<option value="14" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="14","selected"," ") %> >��Ÿ</option>
						-->
					<% end if %>
				<% end if %>
			</select>
			<% end if %>
			<span id="itemSoldOutFlag" name="itemSoldOutFlag" style="display=none" align="right" >
			<input type="radio" name="itemSoldOut" value="N" checked >��ǰ ǰ��ó��
			<input type="radio" name="itemSoldOut" value="S">��ǰ �Ͻ�ǰ��ó��
			</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
	    <td>ó��������</td>
	    <td>
	        <% if (Not C_ADMIN_USER) and ocsmifinishmaster.FOneItem.isMifinishAlreadyInputed then %>
	        	<%= ocsmifinishmaster.FOneItem.FMifinishipgodate %>
	        <% else %>
		        <div id="divipgodate" name="divipgodate" <%= ChkIIF((ocsmifinishmaster.FOneItem.FMifinishReason <> "05" and Not IsNull(ocsmifinishmaster.FOneItem.FMifinishReason)),"style='display:inline'","style='display:none'") %> >
				    <input class="text" type="text" name="ipgodate" value="<%= ocsmifinishmaster.FOneItem.FMifinishipgodate %>" size="10" maxlength="10" onKeyup="ipgodateChange(this);">
				    <a href="javascript:calendarOpen(frmMisend.ipgodate);ipgodateChange(frmMisend.ipgodate);"><img src="/images/calicon.gif" border="0" align="top" height=20></a>
				</div>
			<% end if %>
	    </td>
	</tr>

	<% if (not C_ADMIN_USER) then %>
		<tr bgcolor="#FFFFFF" height="25">
		    <td>�󼼻���</td>
		    <td>
		    	<textarea class="textarea" name="finishmemo" cols="60" rows="6" class="a"><%= ioneas.FOneItem.Fcontents_finish %></textarea>
		    </td>
		</tr>
	<% end if %>

	<tr bgcolor="#FFFFFF" height="25">
	    <td>���ȳ�����</td>
	    <td>
	    	<% if isChulgoState then %>
		        <% if (C_ADMIN_USER) then %>
		            <% if (ocsmifinishmaster.FOneItem.FisSendSms="Y") then %>
		                SMS�߼ۿϷ�/
		                <% if (ocsmifinishmaster.FOneItem.FMifinishReason="05") then %>
		                <input name="ckSendSMS" type="checkbox" disabled  >SMS�߼�&nbsp;
		                <% else %>
		                <input name="ckSendSMS" type="checkbox" checked  >SMS�߼�&nbsp;
		                <% end if %>
		            <% else %>
		                <% if (ocsmifinishmaster.FOneItem.FMifinishReason="05") then %>
		                <input name="ckSendSMS" type="checkbox" disabled  >SMS�߼�&nbsp;
		                <% else %>
		                <input name="ckSendSMS" type="checkbox" checked  >SMS�߼�&nbsp;
		                <% end if %>
		            <% end if %>

		            <% if (ocsmifinishmaster.FOneItem.FisSendEmail="Y") then %>
		                MAIL�߼ۿϷ�/
		                <% if (ocsmifinishmaster.FOneItem.FMifinishReason="05") then %>
		                <input name="ckSendEmail" type="checkbox" disabled  >MAIL�߼�&nbsp;
		                <% else %>
		                <input name="ckSendEmail" type="checkbox" checked  >MAIL�߼�&nbsp;
		                <% end if %>
		            <% else %>
		                <% if (ocsmifinishmaster.FOneItem.FMifinishReason="05") then %>
		                <input name="ckSendEmail" type="checkbox" disabled  >MAIL�߼�&nbsp;
		                <% else %>
		                <input name="ckSendEmail" type="checkbox" checked  >MAIL�߼�&nbsp;
		                <% end if %>
		            <% end if %>
		        <% else %>
	    	        <% if ocsmifinishmaster.FOneItem.isMifinishAlreadyInputed then %>
	    	            <!-- ���ȳ��� �Ϸ�� ���� �������� �� ������� ���� �Ұ� -->
	    	            <%= CHKIIF(ocsmifinishmaster.FOneItem.FisSendSms="Y","SMS�߼ۿϷ�/","") %>
	    	            <%= CHKIIF(ocsmifinishmaster.FOneItem.FisSendEmail="Y","MAIL�߼ۿϷ�/","") %>
	    	            <%= CHKIIF(ocsmifinishmaster.FOneItem.FisSendCall="Y","��ȭ�ȳ��Ϸ�","") %>
	    	        <% else %>
	        	        <input name="ckSendSMS" type="checkbox" checked disabled >SMS�߼�
	        	        &nbsp;
	        	        <input name="ckSendEmail" type="checkbox" checked disabled >MAIL�߼�
	    	        <% end if %>
	    	    <% end if %>
			<% else %>
    	        <input name="ckSendSMS" type="checkbox" disabled >SMS�߼�
    	        &nbsp;
    	        <input name="ckSendEmail" type="checkbox" disabled >MAIL�߼�
			<% end if %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
	    <td colspan="2">
	    	<font color="blue">
	    	<% if isChulgoState then %>
		    	����� ������ ������� �� �ֹ�����(����)�� ���, �Ʒ��� �������� ���Բ� SMS�� ������ �߼۵˴ϴ�.<br>
		    	���Բ� �ȳ��� ��������� �� �����ֽñ� �ٶ��, ���������� ������, �����ͷ� ���� ��Ź�帳�ϴ�.<br>
		    	</font>
		    	<font color="red">
		       	ǰ�����Ұ��� ���, ���Բ� SMS �� ������ �߼۵��� ������, �ٹ����ٰ����Ϳ���<br>
		    	������ ���Բ� ������ �帱 �����Դϴ�.
		    	</font>
		    <% else %>

		<% end if %>
	    </td>
	</tr>
	<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	    <td colspan="2" align="center">
	    <% if (C_ADMIN_USER) then %>
	        <% if (ocsmifinishmaster.FOneItem.isMifinishAlreadyInputed) and (ocsmifinishmaster.FOneItem.FisSendSms="Y") and (ocsmifinishmaster.FOneItem.FisSendEmail="Y") then %>
    	    ���� ����� �����Դϴ�.<br>
    	    <input type="button" class="button" value="��ó�� ���� �ٽ� ����" onclick="MiFinishInput();">
    	    <% else %>
	        <input type="button" class="button" value="��ó�� ���� ����" onclick="MiFinishInput();">
	        <% end if %>
	    <% else %>
    	    <% if ocsmifinishmaster.FOneItem.isMifinishAlreadyInputed then %>
    	    ���� �Ұ�
    	    <% else %>
    	    <input type="button" class="button" value="��ó�� ���� ����" onclick="MiFinishInput();">
    	    <% end if %>
    	<% end if %>
	    </td>
	</tr>
</table>
</form>
<br>
<% else %>
<table width="600">
<tr>
    <td align="center">��ҵ� CS�̰ų� �ش� CS ������ �����ϴ�.</td>
</tr>
</table>
<% end if %>

<%
set ocsmifinishmaster = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
