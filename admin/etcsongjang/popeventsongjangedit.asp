<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ��÷��
' History : 2009.04.17 ���ʻ����� ��
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/event/etcsongjangcls.asp"-->
<%
Dim id
	id = requestCheckvar(request("id"),10)

Dim ibeasong
Set ibeasong = new CEventsBeasong
	ibeasong.FRectId = id
	ibeasong.GetOneWinnerItem

If ibeasong.FResultCount < 1 Then
	response.write "<script>alert('�˻��� ������ �����ϴ�.');</script>"
	response.write "<script>history.back();</script>"
	dbget.close()	:	response.End
End If

Dim i
Dim hpArr,hp1,hp2,hp3
Dim phoneArr,phone1,phone2,phone3

If IsNULL(ibeasong.FOneItem.Freqphone) then ibeasong.FOneItem.Freqphone=""
If IsNULL(ibeasong.FOneItem.Freqhp) then ibeasong.FOneItem.Freqhp=""
If IsNULL(ibeasong.FOneItem.Freqzipcode) then ibeasong.FOneItem.Freqzipcode=""

phoneArr = split(ibeasong.FOneItem.Freqphone,"-")
hpArr = split(ibeasong.FOneItem.Freqhp,"-")

if UBound(hpArr)>=0 then hp1 = hpArr(0)
if UBound(hpArr)>=1 then hp2 = hpArr(1)
if UBound(hpArr)>=2 then hp3 = hpArr(2)

if UBound(phoneArr)>=0 then phone1 = phoneArr(0)
if UBound(phoneArr)>=1 then phone2 = phoneArr(1)
if UBound(phoneArr)>=2 then phone3 = phoneArr(2)

%>
<script type="text/javascript">

function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".zipcode").value = post1 + "-" + post2;
    
    eval(frmname + ".addr1").value = addr;
    eval(frmname + ".addr2").value = dong;
}

function PopSearchZipcode(frmname) {
	var popwin = window.open("/lib/searchzip3.asp?target=" + frmname,"PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function delThis(){
    var frm = document.infoform;

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        if (confirm('������ ���� �Ͻðڽ��ϱ�?')){
            frm.mode.value="del";
    		frm.submit();
		}
	}

}

function gotowrite(){
    var frm = document.infoform;
	if(frm.username.value == ""){
		alert("��÷�ڼ����� �Է����ּ���.");
	    frm.username.focus();
	    return;
	}

    if(frm.reqname.value == ""){
		alert("�����ô� ���� �̸��� �Է����ּ���.");
	    frm.reqname.focus();
	    return;
	}

	if(frm.reqphone1.value == "" || frm.reqphone2.value == "" || frm.reqphone3.value == ""){
		alert("�����ô� ���� ��ȭ��ȣ�� �Է����ּ���.");
	    frm.reqphone1.focus();
	    return;
	}

	if(frm.reqhp1.value == "" || frm.reqhp2.value == "" || frm.reqhp3.value == ""){
		alert("�����ô� ���� �ڵ��� ��ȣ�� �Է����ּ���.");
	    frm.reqphone1.focus();
	    return;
	}

	if(frm.zipcode.value == ""){
		alert("�����ô� ���� �ּҸ� �Է����ּ���.");
	    frm.zipcode.focus();
	    return;
	}

	if(frm.addr2.value == ""){
		alert("�����ô� ���� �������ּҸ� �Է����ּ���.");
	    frm.addr2.focus();
	    return;
	}

	if (frm.reqdeliverdate.value.length!=10){
	    alert('��� ��û���� �Է��ϼ���.');
	    frm.reqdeliverdate.focus();
	    return;
	}

	if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
	    alert('��� ������ ���� �ϼ���.');
	    frm.isupchebeasong[0].focus();
	    return;
	}
	if(frm.isupchebeasong[1].checked&&(frm.jungsan.checked)&&((frm.jungsanValue.value=="")||(frm.jungsanValue.value=="0"))){
	    alert('���� �� �� ��� �����(���԰�)�� �Է��ϼ���');
	    frm.jungsanValue.focus();
	    return;
	}
	if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
	    alert('��ü ����� ��� �귣�� ���̵�  ���� �ϼ���.');
	    frm.makerid.focus();
	    return;
	}


    if (frm.issended.value=="Y"){
        if (frm.songjangdiv.value.length<1){
            alert("�ù�縦 �����ϼ���.");
    	    frm.songjangdiv.focus();
    	    return;
        }

        if (frm.songjangno.value.length<1){
            alert("�����ȣ�� �Է��ϼ���.");
    	    frm.songjangno.focus();
    	    return;
        }
    }

    //�߼ۿϷ�� ���� ���ϴ°�� Check
    if ((frm.isupchebeasong[0].checked)&&(frm.songjangdiv.value.length>0)&&(frm.songjangno.value.length>0)&&(frm.issended.value=="N")){
        alert('�߼� �Ϸ��ΰ�� �߼ۿϷ�� �������ּž� �մϴ�.');
        frm.issended.focus();
        return;
        //if (!confirm("�߼� �Ϸ��ΰ�� �߼ۿϷ�� �������ּž� �մϴ�. \n��� �Ͻðڽ��ϱ�?")){
        //    return;
        //}

    }


	if (confirm('�Է� ������ ��Ȯ�մϱ�?')){
	    frm.mode.value="";
		frm.submit();
	}

}

function disabledBox(comp){
    var frm = comp.form;
    if (comp.value=="Y"){
        frm.makerid.disabled = false;
        frm.jungsan.disabled = false;

		frm.jungsanValue.disabled = false;
        frm.jungsan.checked = true;
    }else{
        //frm.makerid.selectedIndex = 0;
       // frm.makerid.value = '';
		frm.makerid.disabled = true;
		frm.jungsan.disabled = true;

        //frm.jungsanValue.value = '';
        frm.jungsanValue.disabled = true;
        frm.jungsan.checked = false;
    }
}

function jungsanYN(){
	var frm = document.infoform;
	if(frm.jungsan.checked==true){
		frm.jungsanValue.disabled = false;
	}else{
		frm.jungsanValue.value = '';
		frm.jungsanValue.disabled = true;
	}
}

function checkover1(obj) {
	var val = obj.value;
	if (val) {
		if (val.match(/^\d+$/gi) == null) {
			alert("���ڸ� ��������!");
			document.infoform.jungsanValue.value = '';
			obj.select();
			return;
		}
	}
}

</script>
<!--
<table width="600" border="0" cellpadding="0" cellspacing="0" height="50">
  <tr valign="middle">
    <td width="8"><img src="http://fiximage.10x10.co.kr/images/my10x10/myeventmaster_popup_title.gif" width="580" height="50" hspace="10" vspace="10" ></td>
  </tr>
</table>
-->
<table width="100%" border="0" cellpadding="0" cellspacing=0 class="a">
<form name="infoform" method="post" action="/admin/etcsongjang/lib/doeventbeasonginfo.asp">
<input type="hidden" name="id" value="<%= id %>">
<input type="hidden" name="mode" value="">
<tr>
	<td align="center">
		<table width="98%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr height="30">
			<td height="2" colspan="2" >* �̺�Ʈ �� ��Ÿ��� ������� �Է�/ ����</td>
		</tr>
		<tr height="2">
			<td height="2" colspan="2" bgcolor="#AAAAAA"></td>
		</tr>
	<!--
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�̺�Ʈ<br>PrizeCode</td>
			<td style="padding-left:7"></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
	-->

	    <tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����</td>
			<td style="padding-left:7"><%= ibeasong.FOneItem.getEventKind %></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�̺�Ʈ��(���и�)</td>
			<td style="padding-left:7"><%= ibeasong.FOneItem.Fgubunname %></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">���̵�</td>
			<td style="padding-left:7"><%= ibeasong.FOneItem.fuserid %></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">��÷��ǰ</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="prizetitle" size="40" maxlength="64" value="<%= ibeasong.FOneItem.getPrizeTitle %>" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">��÷�ڼ���</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="username" size="20" maxlength="20" value="<%= ibeasong.FOneItem.Fusername %>" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�����μ���</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="reqname" size="20" maxlength="20" value="<%= ibeasong.FOneItem.Freqname %>" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����ó</td>
			<td class="verdana_s" style="padding-left:7">
				<input type="text" class="text" name="reqphone1" size="3" class="verdana_s" maxlength="3" value="<%= phone1 %>">
				-
				<input type="text" class="text" name="reqphone2" size="4" class="verdana_s" maxlength="4" value="<%= phone2 %>">
				-
				<input type="text" class="text" name="reqphone3" size="4" class="verdana_s" maxlength="4" value="<%= phone3 %>">
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�ڵ���</td>
			<td class="verdana_s" style="padding-left:7">
				<input type="text" class="text" name="reqhp1" size="3" class="verdana_s"  maxlength="3" value="<%= hp1 %>">
				-
				<input type="text" class="text" name="reqhp2" size="4" class="verdana_s"  maxlength="4" value="<%= hp2 %>">
				-
				<input type="text" class="text" name="reqhp3" size="4" class="verdana_s"  maxlength="4" value="<%= hp3 %>">
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">������ �ּ�</td>
			<td class="verdana_s" style="padding:5 0 5 7">
				<input type="text" class="text_ro" name="zipcode" size="7" class="verdana_s" readOnly value="<%= ibeasong.FOneItem.Freqzipcode %>">
				<input type="button" class="button" value="�˻�" onClick="FnFindZipNew('infoform','E')">
				<input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('infoform','E')">
				<% '<input type="button" value="�˻�(��)" class="button" onclick="PopSearchZipcode('infoform');" onFocus="this.blur();"> %>
				<br>
				<input type="text" class="text_ro" name="addr1" size="16" maxlength="64"  readOnly value="<%= ibeasong.FOneItem.Freqaddress1 %>" ><br>
				<input type="text" class="text" name="addr2" size="40" maxlength="64" value="<%= ibeasong.FOneItem.Freqaddress2 %>" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">��Ÿ��û����</td>
			<td class="verdana_s" style="padding:5 0 5 7"><textarea class="text" name="reqetc" class="textarea" style="width:350px;height:40px;"><%= ibeasong.FOneItem.Freqetc %></textarea></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		</table>
		<p>
		<table width="98%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
		    <td colspan="4" >* ����ǰ ����</td>
		</tr>
		
		<tr height="1">
			<td height="1" colspan="4" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">��۱���</td>
			<td style="padding-left:7" colspan="3" >
			<% If IsNULL(ibeasong.FOneItem.Fisupchebeasong) or (Not (ibeasong.FOneItem.Fisupchebeasong="Y")) Then %>
				<input type=radio name=isupchebeasong value="N" checked onClick="disabledBox(this);">�ٹ����ٹ��
				<input type=radio name=isupchebeasong value="Y" onClick="disabledBox(this);">��ü�������
			<% Else %>
				<input type=radio name=isupchebeasong value="N" onClick="disabledBox(this);">�ٹ����ٹ��
				<input type=radio name=isupchebeasong value="Y" checked onClick="disabledBox(this);">��ü�������
			<% End If %>
			&nbsp;
			<% drawSelectBoxDesignerwithName "makerid",ibeasong.FOneItem.Fdelivermakerid %>
            
			<% If IsNULL(ibeasong.FOneItem.Fisupchebeasong) or (Not (ibeasong.FOneItem.Fisupchebeasong="Y")) Then %>
			<script language='javascript'>
				document.infoform.makerid.disabled=true;
			</script>
			<% End If %>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="4" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����ǰ�ڵ�</td>
			<td style="padding-left:7" width="30%"><%= ibeasong.FOneItem.Fgift_code %>
			<% if Not isNULL(ibeasong.FOneItem.Fgift_itemid) then %>
			    (��ǰ�ڵ�:<%=ibeasong.FOneItem.Fgift_itemid%>)
			<% end if %>
			</td>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����ǰ��</td>
			<td style="padding-left:7" width="30%"><%= ibeasong.FOneItem.Fgiftkind_name %></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="4" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">���꿩��</td>
			<td style="padding-left:7" width="30%">
				<input type="checkbox" class="checkbox" name="jungsan" id="jungsan" onclick="javascript:jungsanYN();" <%=CHKIIF(ibeasong.FOneItem.FjungsanYN="Y","checked","")%> >������&nbsp;&nbsp;
			</td>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�����(���԰�)</td>
			<td style="padding-left:7" width="30%">
			<input type="text" size="9" style="text-align:right" class="text" id="jungsanValue" name="jungsanValue" value="<%=ibeasong.FOneItem.Fjungsan%>" onkeyup="checkover1(this)" <%=chkiif(IsNULL(ibeasong.FOneItem.Fjungsan) = True,"disabled","")%>>��
			</td>
			
		</tr>
		<tr height="1">
			<td height="1" colspan="4" bgcolor="#DDDDDD"></td>
		</tr>
		</table>
		<p>
		<table width="98%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
		    <td colspan="2" >* �������</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����û��</td>
			<td class="verdana_s" style="padding:5 0 5 7">
			<input type="text" class="text_ro" name="reqdeliverdate" size="10" maxlength="10"  value="<%= ibeasong.FOneItem.FreqDeliverDate %>" >
			<a href="javascript:jsPopCal('reqdeliverdate');"><img src="/images/calicon.gif" border="0" align="absmiddle"></a>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�߼ۻ��� / �����</td>
			<td style="padding-left:7">
				<select name="issended" >
				<% If ibeasong.FOneItem.Fissended="Y" Then %>
					<option value="N">�̹߼�
					<option value="Y" selected >�߼ۿϷ�
				<% Else %>
					<option value="N" selected >�̹߼�
					<option value="Y">�߼ۿϷ�
				<% End If %>
				</select>
				/ <%= ibeasong.FOneItem.Fsenddate %>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����</td>
			<td style="padding-left:7">
				<% drawSelectBoxDeliverCompany "songjangdiv",ibeasong.FOneItem.Fsongjangdiv %>
				<input type="text" class="text" name="songjangno" size="14" maxlength="20" value="<%= ibeasong.FOneItem.Fsongjangno %>">
			</td>
		</tr>
		<tr height="2">
			<td height="2" colspan="2" bgcolor="#AAAAAA"></td>
		</tr>
		<tr height="30">
			<td colspan="2" align="center">
		<% If (ibeasong.FOneItem.IsSended) Then %>
			<input type="button" class="button" value=" �� �� " onClick="if (confirm('�̹� �߼۵� ���� �Դϴ�. ���� �Ͻðڽ��ϱ�?')) { gotowrite(); };" onfocus="this.blur();">
		<% Else %>
			<input type="button" class="button" value=" �� �� " onClick="gotowrite();" onfocus="this.blur();">
			&nbsp;&nbsp;&nbsp;
			<input type="button" class="button" value=" �� �� " onClick="delThis();" onfocus="this.blur();">
		<% End If %>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<% Set ibeasong = Nothing %>
<!-- #include virtual="/admin/lib/poptail.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->