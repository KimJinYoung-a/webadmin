<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/jumun/misendcls.asp"-->

<style type="text/css" >
.sale11px01 {font-family: dotum; FONT-SIZE: 11px; font-weight:bold ; COLOR: #b70606;}
</style>
<%
''�귣��/ ���� ������

dim idx : idx= requestCheckVar(request("idx"),10)

dim omisend
set omisend = new COldMiSend
omisend.FRectDetailIDx = idx
omisend.getOneOldMisendItem

if (omisend.FResultCount<1) then
    response.write "�˻������ �����ϴ�."
    dbget.close() : response.end
end if

if (LCase(omisend.FOneItem.FMakerid) <> LCASE(session("ssBctID"))) then
    response.write "������ �����ϴ�."
    dbget.close() : response.end
end if

dim PreDispMail
PreDispMail = (omisend.FOneItem.isMisendAlreadyInputed) and (omisend.FOneItem.FMisendReason<>"05")

%>
<script language='javascript'>
function getOnload(){
    // popupResize(640);
}
window.onload = getOnload;

function ShowDateBox(comp){
    var frm = comp.form;
    var iid = comp.id;
    var idiv = document.all.divipgodate;
    var isms = document.all.iSMSDISP;
    var iemail = document.all.iEMAILDISP;
    var isDPlusOver = true;
    var isold = document.all.itemSoldOutFlag

    document.all.iSMSDISP02.style.display = "none";
    document.all.iSMSDISP03.style.display = "none";
    document.all.iSMSDISP04.style.display = "none";
    document.all.iSMSDISP02_1.style.display = "none";
    document.all.iSMSDISP03_1.style.display = "none";
    document.all.iSMSDISP04_1.style.display = "none";

    document.all.iEMAILMENT02.style.display = "none";
    document.all.iEMAILMENT03.style.display = "none";
    document.all.iEMAILMENT04.style.display = "none";
    document.all.iEMAILMENT02_1.style.display = "none";
    document.all.iEMAILMENT03_1.style.display = "none";
    document.all.iEMAILMENT04_1.style.display = "none";

    if ((comp.value=="03")||(comp.value=="02")||(comp.value=="04")){
        idiv.style.display = "inline";
        isms.style.display = "inline";
        iemail.style.display = "inline";

        if ((frm.baljudate.value.length>0)&&(frm.ipgodate.value.length>0)){
            if (getDiffDay(frm.baljudate.value,frm.ipgodate.value)<2){
                isDPlusOver=false;
            }
        }

        if (comp.value=="03"){
            if (isDPlusOver){
                document.all.iSMSDISP03.style.display = "inline";
                document.all.iSMSDISP03_1.style.display = "none";

                document.all.iEMAILMENT03.style.display = "inline";
                document.all.iEMAILMENT03_1.style.display = "none";
            }else{
                document.all.iSMSDISP03.style.display = "none";
                document.all.iSMSDISP03_1.style.display = "inline";

                document.all.iEMAILMENT03.style.display = "none";
                document.all.iEMAILMENT03_1.style.display = "inline";
            }
        }else if(comp.value=="02"){
            if (isDPlusOver){
                document.all.iSMSDISP02.style.display = "inline";
                document.all.iSMSDISP02_1.style.display = "none";

                document.all.iEMAILMENT02.style.display = "inline";
                document.all.iEMAILMENT02_1.style.display = "none";
            }else{
                document.all.iSMSDISP02.style.display = "none";
                document.all.iSMSDISP02_1.style.display = "inline";

                document.all.iEMAILMENT02.style.display = "none";
                document.all.iEMAILMENT02_1.style.display = "inline";
            }
        }else if(comp.value=="04"){
            if (isDPlusOver){
                document.all.iSMSDISP04.style.display = "inline";
                document.all.iSMSDISP04_1.style.display = "none";

                document.all.iEMAILMENT04.style.display = "inline";
                document.all.iEMAILMENT04_1.style.display = "none";
            }else{
                document.all.iSMSDISP04.style.display = "none";
                document.all.iSMSDISP04_1.style.display = "inline";

                document.all.iEMAILMENT04.style.display = "none";
                document.all.iEMAILMENT04_1.style.display = "inline";
            }
        }
    }else{
        idiv.style.display = "none";
        isms.style.display = "none";
        iemail.style.display = "none";
    };

   //ǰ�����Ұ�
   if (comp.value=="05") {
        isold.style.display = "inline";
   }else{
        isold.style.display = "none";
   }
}

function ipgodateChange(comp){
    var v = comp.value;
    if (v.length<10) v = "YYYY-MM-DD";

    document.getElementById("iMisendIpgodate02").innerHTML = v;
    document.getElementById("iMisendIpgodate02_1").innerHTML = v;
    document.getElementById("iMisendIpgodate03").innerHTML = v;
    document.getElementById("iMisendIpgodate03_1").innerHTML = v;
    document.getElementById("iMisendIpgodate04").innerHTML = v;
    document.getElementById("iMisendIpgodate04_1").innerHTML = v;

    document.getElementById("iMisendIpgodate2").innerHTML = v;

    ShowDateBox(frmMisend.MisendReason);
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


    //�������(03), �ֹ�����(02), ������(04)
    if ((frm.MisendReason.value=="03")||(frm.MisendReason.value=="02")||(frm.MisendReason.value=="04")){
        var ipgodate = eval("frm.ipgodate");
        if (ipgodate.value.length!=10){
            alert('��� �������� �Է��ϼ���.(YYYY-MM-DD)');
            ipgodate.focus();
            return;
        }

        inputdate = new Date(ipgodate.value.substr(0,4),ipgodate.value.substr(5,2)*1-1,ipgodate.value.substr(8,2));
        if (today>inputdate){
            alert('��� �������� ���� ���ĳ�¥�� ������ �����մϴ�.');
            //ipgodate.focus();
            return;
        }
    }

    if (confirm('����� ������ ���� �Ͻðڽ��ϱ�?')){
	    frm.action = "upchebeasong_Process.asp";
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

<% if omisend.FResultCount>0 then %>
<table width="610" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmMisend" method="post" action="upchebeasong_Process.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="misendInputOne">
	<input type="hidden" name="detailidx" value="<%= omisend.FOneItem.Fidx %>">
	<input type="hidden" name="baljudate" value="<%= CHKIIF(IsNULL(omisend.FOneItem.Fbaljudate),"",Left(omisend.FOneItem.Fbaljudate,10)) %>">
	<input type="hidden" name="upcheconfirmdate" value="<%= CHKIIF(IsNULL(omisend.FOneItem.Fupcheconfirmdate),"",Left(omisend.FOneItem.Fupcheconfirmdate,10)) %>">

	<input type="hidden" name="Sitemid" value="<%= omisend.FOneItem.FItemID %>">
	<input type="hidden" name="Sitemoption" value="<%= omisend.FOneItem.FItemOption %>">

	<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	    <td colspan="2">
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�������� �Է�</b>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
    	<td width="130">��ǰ�ڵ�</td>
    	<td width="480"><%= omisend.FOneItem.FItemID %>

    	    <% if (omisend.FOneItem.FCancelyn<>"N") then %>
			<b><font color="#CC3333">[����ֹ�]</font></b>
			<script language='javascript'>alert('��ҵ� �ŷ� �Դϴ�.');</script>
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
	    <td><img src="<%= omisend.FOneItem.Fsmallimage %>" width="60" ></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td>��ǰ��</td>
	    <td><%= omisend.FOneItem.FItemName %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td>�ɼ�</td>
	    <td><%= omisend.FOneItem.FItemoptionName %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td>�ֹ�����</td>
	    <td><%= omisend.FOneItem.FItemcnt %>��</td>
	</tr>

	<tr bgcolor="#FFFFFF">
	    <td>��������</td>
	    <td>
	        <% if omisend.FOneItem.isMisendAlreadyInputed then %>
	        <%= omisend.FOneItem.getMiSendCodeName %>
	        <% else %>
	        <select name="MisendReason" id="MisendReason" class="select" onChange="ShowDateBox(this);">
				<option value="">---------</option>
				<option value="03" <%= ChkIIF(omisend.FOneItem.FMisendReason="03","selected"," ") %> >�������</option>
				<option value="05" <%= ChkIIF(omisend.FOneItem.FMisendReason="05","selected"," ") %> >ǰ�����Ұ�</option>
				<option value="02" <%= ChkIIF(omisend.FOneItem.FMisendReason="02","selected"," ") %> >�ֹ�����</option>
				<option value="04" <%= ChkIIF(omisend.FOneItem.FMisendReason="04","selected"," ") %> >������</option>
				<!-- �ٹ����ٹ�� ���������� ���������� �մϴ�. -->
			</select>
			<% end if %>
			<span id="itemSoldOutFlag" name="itemSoldOutFlag" style="display=none" align="right" >
			<input type="radio" name="itemSoldOut" value="N" checked >��ǰ ǰ��ó��
			<input type="radio" name="itemSoldOut" value="S">��ǰ �Ͻ�ǰ��ó��
			</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td>�������</td>
	    <td>
	        <% if omisend.FOneItem.isMisendAlreadyInputed then %>
	        <%= omisend.FOneItem.FMisendIpgodate %>
	        <% else %>
	        <div id="divipgodate" name="divipgodate" <%= ChkIIF(omisend.FOneItem.FMisendReason="03" or omisend.FOneItem.FMisendReason="02","style='display:inline'","style='display:none'") %> >
			    <input class="text" type="text" name="ipgodate" value="<%= omisend.FOneItem.FMisendIpgodate %>" size="10" maxlength="10" onChange="ipgodateChange(this);">
			    <a href="javascript:calendarOpen(frmMisend.ipgodate);ipgodateChange(frmMisend.ipgodate);"><img src="/images/calicon.gif" border="0" align="top" height=20></a>
			</div>
			<% end if %>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td>���ȳ�����</td>
	    <td>
	        <% if omisend.FOneItem.isMisendAlreadyInputed then %>
	            <%= CHKIIF(omisend.FOneItem.FisSendSms="Y","SMS�߼ۿϷ�/","") %>
	            <%= CHKIIF(omisend.FOneItem.FisSendEmail="Y","MAIL�߼ۿϷ�/","") %>
	            <%= CHKIIF(omisend.FOneItem.FisSendCall="Y","��ȭ�ȳ��Ϸ�","") %>
	        <!-- ���ȳ��� �Ϸ�� ���� �������� �� ������� ���� �Ұ� -->
	        <% else %>
    	        <input name="ckSendSMS" type="checkbox" checked disabled >SMS�߼�
    	        &nbsp;
    	        <input name="ckSendEmail" type="checkbox" checked disabled >MAIL�߼�
	        <% end if %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF">
	    <td colspan="2">
	    	<font color="blue">
	    	����� ������ ������� �� �ֹ������� ���, �Ʒ��� �������� ���Բ� SMS�� ������ �߼۵˴ϴ�.<br>
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
		    <% if omisend.FOneItem.isMisendAlreadyInputed then %>
		    ���� �Ұ�
		    <% else %>
		    <input type="button" class="button" value="����� ���� ����" onclick="MisendInput();">
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
	<tr bgcolor="#FFFFFF" id="iSMSDISP" style="display:<%= chkIIF(PreDispMail,"inline","none") %>" >
	    <td>
        	<table width="610" align="center" cellspacing="1" cellpadding="0" class="a" >
        	<tr bgcolor="#FFFFFF" id="iSMSDISP02" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="02","inline","none") %>">
            	<td>
            		[���ΰŽ� ��������ȳ�]�ֹ��Ͻ� ��ǰ�� <%= DdotFormat(omisend.FOneItem.FItemName,16) %>(<%= omisend.FOneItem.FItemID %>)��ǰ�� �ֹ����� ��ǰ���� <span id="iMisendIpgodate02" name="iMisendIpgodate02"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span>�� �߼۵� �����Դϴ�. ���ο� ������ ��� �˼��մϴ�.
            	</td>
            </tr>
            <tr bgcolor="#FFFFFF" id="iSMSDISP02_1" style="display:none">
            	<td>
            		[���ΰŽ� ��������ȳ�]�ֹ��Ͻ� ��ǰ�� <%= DdotFormat(omisend.FOneItem.FItemName,16) %>(<%= omisend.FOneItem.FItemID %>)��ǰ�� <span id="iMisendIpgodate02_1" name="iMisendIpgodate02_1"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span>�� �߼۵� �����Դϴ�. �����մϴ�.
            	</td>
            </tr>
        	<tr bgcolor="#FFFFFF" id="iSMSDISP03" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="03","inline","none") %>">
            	<td>
            		[���ΰŽ� ��������ȳ�]�ֹ��Ͻ� ��ǰ�� <%= DdotFormat(omisend.FOneItem.FItemName,16) %>(<%= omisend.FOneItem.FItemID %>)��ǰ�� <span id="iMisendIpgodate03" name="iMisendIpgodate03"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span>�� �߼۵� �����Դϴ�. ���ο� ������ ��� �˼��մϴ�.
            	</td>
            </tr>
            <tr bgcolor="#FFFFFF" id="iSMSDISP03_1" style="display:none">
            	<td>
            		[���ΰŽ� ��������ȳ�]�ֹ��Ͻ� ��ǰ�� <%= DdotFormat(omisend.FOneItem.FItemName,16) %>(<%= omisend.FOneItem.FItemID %>)��ǰ�� <span id="iMisendIpgodate03_1" name="iMisendIpgodate03_1"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span>�� �߼۵� �����Դϴ�. �����մϴ�.
            	</td>
            </tr>
            <tr bgcolor="#FFFFFF" id="iSMSDISP04" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="04","inline","none") %>">
            	<td>
            		[���ΰŽ� ������ȳ�]�ֹ��Ͻ� ��ǰ�� <%= DdotFormat(omisend.FOneItem.FItemName,16) %>(<%= omisend.FOneItem.FItemID %>)��ǰ�� �����ۻ�ǰ���� <span id="iMisendIpgodate04" name="iMisendIpgodate4"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span>�� �߼۵� �����Դϴ�. �����մϴ�.
            	</td>
            </tr>
            <tr bgcolor="#FFFFFF" id="iSMSDISP04_1" style="display:none">
            	<td>
            		[���ΰŽ� ������ȳ�]�ֹ��Ͻ� ��ǰ�� <%= DdotFormat(omisend.FOneItem.FItemName,16) %>(<%= omisend.FOneItem.FItemID %>)��ǰ�� �����ۻ�ǰ���� <span id="iMisendIpgodate04_1" name="iMisendIpgodate04_1"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span>�� �߼۵� �����Դϴ�. �����մϴ�.
            	</td>
            </tr>
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
	<tr bgcolor="#FFFFFF" id="iEMAILDISP" style="display:<%= chkIIF(PreDispMail,"inline","none") %>">
    	<td>
    		<!-- ���� ���� ���� -->

    		<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td>

						<!-- ������ ���� -->
						<table width="600" border="0" align="center" cellspacing="0" cellpadding="0" class="a">
						<tr>
							<td><a href="http://www.thefingers.co.kr" target="_blank" onFocus="blur()">
								<img src="http://image.thefingers.co.kr/2016/mail/img_logo.png" width="600" height="85" border="0" /></a>
							</td>
						</tr>
						<tr>
							<td style="border:7px solid #eeeeee;">
								<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
								<tr>
									<td><img src="http://image.thefingers.co.kr/academy2010/email/sorry.gif" width="586"> </td>
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
										<tr id="iEMAILMENT03" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="03","inline","none") %>">
											<td>
												<!-- ��������� ��� D+2 -->
												�ȳ��ϼ���.   ����<br>
												���Բ��� �ֹ��Ͻ� ��ǰ�� �߼��� ������ �����Դϴ�.<br>
												�Ʒ� �߼ۿ����Ͽ� �߼۵� �����̿���, �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>
												���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>
												���ο� ������ �帰 �� �������� ����帮��, ������� ������ �� �� �ֵ��� �ּ��� ���ϰڽ��ϴ�.<br>
											</td>
										</tr>
										<tr id="iEMAILMENT03_1" style="display:none">
											<td>
												<!-- ��������� ��� D+0/1 -->
												�ȳ��ϼ���.   ����<br>
												���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.<br>
												�Ʒ� �߼ۿ����Ͽ� �߼۵� �����̿���, �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>
												���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>
											</td>
										</tr>
										<tr id="iEMAILMENT02" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="02","inline","none") %>">
										    <td>
												<!-- �ֹ����� ��� D+2 -->
												�ȳ��ϼ���.  ����<br>
												���Բ��� �ֹ��Ͻ� ��ǰ�� �ֹ� �� ���۵Ǵ� ��ǰ����<br>
												�Ϲݻ�ǰ�� �޸� �ֹ����۱Ⱓ�� �ҿ�Ǵ� ��ǰ�Դϴ�.<br>
												�Ʒ��� ���� �߼� �������� �ȳ��ص帮����,<br>
												�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>
											</td>
										</tr>
										<tr id="iEMAILMENT02_1" style="display:none">
										    <td>
												<!-- �ֹ����� ��� D+0/1 -->
												�ȳ��ϼ���.  ����<br>
												���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.<br>
												�Ʒ��� ���� �߼ۿ������� �ȳ��� �帳�ϴ�.<br>
												�Ǹ��ڰ� ��ǰ�� �߼��� ������ ���ݸ� ��ٷ� �ֽø� �����ϰڽ��ϴ�.<br>
											</td>
										</tr>
										<tr id="iEMAILMENT04" style="display:<%= CHKIIF(omisend.FOneItem.FMisendReason="04","inline","none") %>">
										    <td>
												<!-- �����ǰ ��� D+2 -->
												�ȳ��ϼ���.  ����<br>
												���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.<br>
                                                �ֹ��Ͻ� ��ǰ�� �����ۻ�ǰ���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>
                                                �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>
                                                ���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>
											</td>
										</tr>
										<tr id="iEMAILMENT04_1" style="display:none">
										    <td>
												<!-- �����ǰ ��� D+0/1 -->
												�ȳ��ϼ���.  ����<br>
												���Բ��� �ֹ��Ͻ� ��ǰ�� ���ȳ� �����Դϴ�.<br>
                                                �ֹ��Ͻ� ��ǰ�� �����ۻ�ǰ���� �Ʒ� �߼ۿ����Ͽ� �߼۵� �����̸�,<br>
                                                �ε����� �������� ��ǰ��Ҹ� ���Ͻô� ���,<br>
                                                ���ູ���ͷ� ���� ��Ź�帳�ϴ�.<br>

											</td>
										</tr>
										<tr id="iEMAILMENT05" style="display:none">
										    <td>
										        <!-- ǰ�� ���Ұ��� ��� --- �̰� ��ü������ �߼� ���� ���ΰŽ� �����Ϳ����� �߼� ��Ʈ ���߿� �߰�-->
										    </td>
										</tr>
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
													<td width="450"class="gray12px02" style="padding-left:10px;padding-top:2px;"><img src="<%= omisend.FOneItem.Fsmallimage %>" width="60" ></td>
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
													<!-- ���ΰŽ� ����� ��� ���ΰŽ� ��������, ��ü�ϰ��, ��üȸ���-->
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
													<td height="24" align="center" bgcolor="#f7f7f7" class="gray12px02b" style="padding-top:2px;">�߼ۿ�����</td>
													<td class="gray12px02" style="padding-left:10px;padding-top:2px;"><b><span id="iMisendIpgodate2" onClick="ipgodateChange(frmMisend.ipgodate);" name="iMisendIpgodate2"><%= CHKIIF(omisend.FOneItem.FMisendipgodate<>"",omisend.FOneItem.FMisendipgodate,"YYYY-MM-DD") %></span></b></td>
												</tr>
												<tr>
													<td height="1" colspan="2" bgcolor="#cccccc"></td>
												</tr>
												<tr>
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
								(03086) ����� ���α� ���з�12�� 31 �������� 5�� (��)�ٹ����� ���ΰŽ�<br>
								��ǥ�̻� : ������  &nbsp;����ڵ�Ϲ�ȣ:211-87-00620  &nbsp;����Ǹž� �Ű��ȣ : �� 01-1968ȣ  &nbsp;�������� ��ȣ �� û�ҳ� ��ȣå���� : �̹���<br>
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
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->