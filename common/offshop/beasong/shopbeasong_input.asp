<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.02.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim i , orderno , itemgubunarr ,itemoptionarr, itemidarr, mode , shopidarr, reqhp,comment, confirmcertno
dim buyname,buyphone, buyhp, buyemail, reqname, reqzipcode, reqzipaddr, reqaddress, reqphone
dim buyphone1, buyphone2, buyphone3 ,buyhp1 ,buyhp2 ,buyhp3 ,reqphone1 ,reqphone2 ,reqphone3
dim reqhp1, reqhp2 ,reqhp3 ,buyemail1 ,buyemail2 , reqaddress1 ,reqaddress2 , ojumun, shopid
dim masteridx_beasong ,oedit , shopname ,shopIpkumDivName ,ipkumdiv, UserHp1, UserHp2, UserHp3, checkblock
dim BeaSongcnt, UserHp, SmsYN, KakaoTalkYN, totrealprice, ExistsItemBeasongYN, ExistsBeasongYN, dbCertNo
	orderno = requestcheckvar(request("orderno"),16)
	masteridx_beasong = requestcheckvar(request("masteridx"),10)
	mode = requestcheckvar(request("mode"),32)

totrealprice=0
ExistsItemBeasongYN="N"
ExistsBeasongYN="N"

if orderno="" or isnull(orderno) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('�ֹ���ȣ�� �����ϴ�');"
	response.write "</script>"
	dbget.close()	:	response.End
end if

set oedit = new cupchebeasong_list
	oedit.frectmasteridx = masteridx_beasong
	oedit.frectorderno = orderno

	if masteridx_beasong <> "" or orderno <> "" then
		oedit.fshopjumun_edit()

		if oedit.ftotalcount > 0 then
			ExistsBeasongYN="Y"

			IpkumDiv = oedit.FOneItem.fIpkumDiv
			shopid = oedit.FOneItem.fshopid
			orderno = oedit.FOneItem.forderno
			buyname = oedit.FOneItem.fbuyname
			buyphone = oedit.FOneItem.fbuyphone
				if buyphone<>"" then
					if instr(buyphone,"-") = 0 then
						buyphone = left(buyphone,3)
						buyphone = mid(buyphone,4,len(buyphone)-3-4)
						buyphone = right(buyphone,4)
					else
						buyphone1 = split(buyphone,"-")(0)
						buyphone2 = split(buyphone,"-")(1)
						buyphone3 = split(buyphone,"-")(2)
					end if
				end if
			buyhp = oedit.FOneItem.fbuyhp
				if buyhp<>"" then
					buyhp1 = split(buyhp,"-")(0)
					buyhp2 = split(buyhp,"-")(1)
					buyhp3 = split(buyhp,"-")(2)
				end if
			buyemail = oedit.FOneItem.fbuyemail
				if buyemail<>"" then
					buyemail1 = split(buyemail,"@")(0)
					buyemail2 = split(buyemail,"@")(1)
				end if
			reqname = oedit.FOneItem.freqname
			reqzipcode = oedit.FOneItem.freqzipcode
			reqzipaddr = oedit.FOneItem.freqzipaddr
			reqaddress = oedit.FOneItem.freqaddress
			reqphone = oedit.FOneItem.freqphone
				if reqphone<>"" then
					reqphone1 = split(reqphone,"-")(0)
					reqphone2 = split(reqphone,"-")(1)
					reqphone3 = split(reqphone,"-")(2)
				end if
			reqhp = oedit.FOneItem.freqhp
				if reqhp<>"" then
					if instr(reqhp,"-") = 0 then
						reqhp1 = left(reqhp,3)
						reqhp2 = mid(reqhp,4,len(reqhp)-3-4)
						reqhp3 = right(reqhp,4)
						'response.write reqhp1 & "/" & reqhp2 & "/" & reqhp3
					else
						reqhp1 = split(reqhp,"-")(0)
						reqhp2 = split(reqhp,"-")(1)
						reqhp3 = split(reqhp,"-")(2)
					end if
				end if
			comment = oedit.FOneItem.fcomment
			shopname = oedit.FOneItem.fshopname
			shopIpkumDivName = oedit.FOneItem.shopIpkumDivName
		
			BeaSongcnt = oedit.FOneItem.fBeaSongcnt
			UserHp = oedit.FOneItem.fUserHp
				if UserHp<>"" then
					if instr(UserHp,"-") = 0 then
						UserHp1 = left(UserHp,3)
						UserHp2 = mid(UserHp,4,len(UserHp)-3-4)
						UserHp3 = right(UserHp,4)
					else
						UserHp1 = split(UserHp,"-")(0)
						UserHp2 = split(UserHp,"-")(1)
						UserHp3 = split(UserHp,"-")(2)
					end if
				end if
			SmsYN = oedit.FOneItem.fSmsYN
			KakaoTalkYN = oedit.FOneItem.fKakaoTalkYN
			dbCertNo = oedit.FOneItem.fCertNo
		end if
	end if

set ojumun = new cupchebeasong_list
	ojumun.frectmasteridx_beasong = masteridx_beasong
	ojumun.frectorderno = orderno

	if orderno <> "" then
		ojumun.fshopbeasong_input()
	end if

function IsUpcheBeasong(odlvType)
	if (CStr(odlvType) = "2") then
		IsUpcheBeasong = "Y"
	else
		IsUpcheBeasong = "N"
	end if
end function
%>

<script language="javascript">

	// �ֹ�������������
	function certedit(orderno,masteridx_beasong,vmode){
		if (vmode==''){
			alert('�����ڰ� �����ϴ�.');
			return;
		}

		frmsmscert.masteridx.value=masteridx_beasong;
		frmsmscert.orderno.value=orderno;
		frmsmscert.action = '/common/offshop/beasong/shopbeasong_process.asp';

		if (vmode=='ReSendKakaotalk'){
			if (confirm('īī������ �߼� �Ͻðڽ��ϱ�?')){
				frmsmscert.mode.value=vmode;
				frmsmscert.submit();
			}
		}else if (vmode=='ReSendSMS'){
			if (confirm('SMS�� �߼� �Ͻðڽ��ϱ�?')){
				frmsmscert.mode.value=vmode;
				frmsmscert.submit();
			}
		}else{
			if (confirm('�ֹ����������� ���� �Ͻðڽ��ϱ�?')){
				frmsmscert.mode.value=vmode;
				frmsmscert.submit();
			}
		}
	}
	
	// �������������. �̹��� �̻�� ��û���� �˾����� ����
	function jumundetail(orderno,masteridx_beasong){
		var popwin = window.open('/common/offshop/beasong/shopjumun_address.asp?mode=addressedit&orderno='+orderno+'&masteridx='+masteridx_beasong+'&menupos=<%=menupos%>','popbeasongedit','width=1280,height=960,scrollbars=yes,resizable=yes');
		popwin.focus();

//		frminfo.masteridx.value=masteridx_beasong;
//		frminfo.orderno.value=orderno;
//		frminfo.mode.value='addressedit';
//		frminfo.action = '/common/offshop/beasong/shopjumun_address.asp';
//		frminfo.submit();
	}

	//������������
	function refer(){
		location.href='/common/offshop/beasong/shopbeasong_list.asp?menupos=<%= menupos %>';
	}

	//��ǰ ����
	function detaildel(detailidx_beasong,masteridx_beasong,odlvType,orderno){
		frminfo.orderno.value=orderno;
		frminfo.odlvType.value=odlvType;
		frminfo.detailidx.value=detailidx_beasong;
		frminfo.masteridx.value=masteridx_beasong;
		frminfo.mode.value='detaildel';
		frminfo.action='/common/offshop/beasong/shopbeasong_process.asp';
		frminfo.submit();
	}

	//��ǰ����
	function jumunedit(upfrm){
		var masteridx_beasong = '<%= masteridx_beasong %>';
		var orderno = '<%= orderno %>';

		upfrm.detailidxarr.value='';
		upfrm.masteridx.value='';
		upfrm.odlvTypearr.value='';

		if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}

		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					if (frm.odlvType.value==''){
						alert('������ɻ��� ��۱����� ���� �ϼ���.');
						frm.odlvType.focus();
						return;
					}
					// comm_cd : B031 ����������� / B012 ��üƯ�� / B013 ���Ư��
					// �������
					if (frm.odlvType.value*1 == '1') {
						if (frm.comm_cd.value == 'B012') {
							alert("�ش� ��ǰ�� ��ü��� or �����۸� ���� �մϴ�.");
							frm.odlvType.focus();
							return;
						}
					}
					// ��ü���
					if (frm.odlvType.value*1 == '2') {
						if (frm.comm_cd.value == 'B031' || frm.comm_cd.value == 'B013') {
							alert("�ش� ��ǰ�� ������� or �����۸� ���� �մϴ�.");
							frm.odlvType.focus();
							return;
						}
					}

/*
					if (frm.defaultbeasongdiv.value*1 == 0) {
						if (frm.odlvType.value*1 != 0) {
							alert("�����Ҽ� ���� ������Դϴ�. �������� �����ϼ���.");
							frm.odlvType.focus();
							return;
						}
					}
*/
					if (frm.currstate.value*1 > 3) {
						alert("�ش� ��ǰ�� �̹� ��� �Ϸ�� ��ǰ �Դϴ�.");
						frm.odlvType.focus();
						return;
					}
					if (frm.currstate.value*1 > 2) {
						alert("[����]�ش� ��ǰ�� �̹� ��ü���� ����� Ȯ���� ���� �Դϴ�.");
					}

					upfrm.odlvTypearr.value = upfrm.odlvTypearr.value + frm.odlvType.value + "," ;
					upfrm.detailidxarr.value = upfrm.detailidxarr.value + frm.detailidx.value + "," ;
					upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "," ;
					upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + "," ;
					upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "," ;
				}
			}
		}

		upfrm.orderno.value= orderno;
		upfrm.masteridx.value= masteridx_beasong;
		upfrm.mode.value='jumunedit';
		upfrm.action='/common/offshop/beasong/shopbeasong_process.asp';
		upfrm.submit();
	}

	function sendSMSEmail(makerid,orderno,masteridx_beasong,detailidx){
		var sendSMSEmail = window.open("/common/offshop/beasong/popupchejumunsms_off.asp?memupos=<%=menupos%>&makerid=" + makerid + "&orderno=" + orderno + "&masteridx=" + masteridx_beasong + "&detailidx=" + detailidx,"sendSMSEmail","width=600 height=500,scrollbars=yes,resizabled=yes");
		sendSMSEmail.focus();
	}

	function CheckThis(frm){
		frm.cksel.checked=true;
		AnCheckClick(frm.cksel);
	}

</script>

<% if ExistsBeasongYN="Y" then %>
	<form name="frmsmscert" method="post">
	<input type="hidden" name="mode">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="masteridx" value="<%= masteridx_beasong %>">
	<input type="hidden" name="orderno" value="<%= orderno %>">
	<input type="hidden" name="loginidshopormaker" value="<%= shopid %>">
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan=8>
			�ֹ���������
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			�޴�����ȣ(�ֹ����Է�)
		</td>
		<td>
			<input type="text" name="UserHp1" value="<%= UserHp1 %>" size=4 maxlength=4>-<input type="text" name="UserHp2" value="<%= UserHp2 %>" size=4 maxlength=4>-<input type="text" name="UserHp3" value="<%= UserHp3 %>" size=4 maxlength=4>
		</td>
		<td>
			īī����߼ۿ���
		</td>
		<td>
			<%= KakaoTalkYN %>
			
			<%
			' ��ü�뺸 ���� ���� ���
			if IpkumDiv < 5 then
			%>
				<input type="button" class="button" value="�߼�" onclick="certedit('<%= orderno %>','','ReSendKakaotalk')">
				<!--<input type="button" class="button" value="�߼�" onclick="alert('�۾���\nSMS�� �߼��ϼ���.'); return false;">-->
			<% elseif  C_ADMIN_AUTH then %>
				<input type="button" class="button" value="�߼�[������]" onclick="certedit('<%= orderno %>','','ReSendKakaotalk')">
			<% end if %>
		</td>
		<td>
			SMS�߼ۿ���
		</td>
		<td>
			<%= SmsYN %>

			<%
			' ��ü�뺸 ���� ���� ���
			if IpkumDiv < 5 then
			%>
				<input type="button" class="button" value="�߼�" onclick="certedit('<%= orderno %>','','ReSendSMS')">
			<% elseif  C_ADMIN_AUTH then %>
				<input type="button" class="button" value="�߼�[������]" onclick="certedit('<%= orderno %>','','ReSendSMS')">
			<% end if %>
		</td>
		<td>
			��ۿ���
		</td>
		<td>
			<% if BeaSongcnt > 0 then %>
				Y
			<% else %>
				N
			<% end if %>
		</td>
	</tr>

	<%
	if C_ADMIN_AUTH then	' or C_OFF_AUTH 
		if KakaoTalkYN="Y" or KakaoTalkYN="N" then
	%>
			<tr align="left" bgcolor="#FFFFFF">
				<td colspan=8>
					<%
					confirmcertno = md5(trim(orderno) & dbCertNo & replace(trim(UserHp1)&trim(UserHp2)&trim(UserHp3),"-",""))
					%>
					�����ڱ��� : <% response.write "https://m.10x10.co.kr/my10x10/order/myshoporder.asp?orderNo="& trim(orderno) &"&certNo="& confirmcertno &"" %>
				</td>
			</tr>
	<%
		end if
	end if
	%>

	<tr align="center" bgcolor="#FFFFFF">
		<td colspan=8>
			<%
			' ��üȮ�� ���� ���� ���
			if IpkumDiv < 6 then
			%>
				<input type="button" onclick="certedit('<%= orderno %>','','certedit')" value="�ֹ�������������" class="button">
			<% elseif  C_ADMIN_AUTH then %>
				<input type="button" onclick="certedit('<%= orderno %>','','certedit')" value="�ֹ�������������[������]" class="button">
			<% end if %>
		</td>
	</tr>
	</table>
	</form>

	<br>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan=8>
			�������
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			�ֹ���ȣ
		</td>
		<td>
			<%= orderno %>
		</td>
		<td>
			�ǸŸ���
		</td>
		<td>
			<%=shopname%>
		</td>
		<td>
			������
		</td>
		<td>
			<font color="red"><%= shopIpkumDivName %></font>
		</td>
		<td>
		</td>
		<td>
		</td>
	</tr>
	<!--<tr align="center" bgcolor="#FFFFFF">
		<td>
			�ֹ��ڼ���
		</td>
		<td>
			<%=buyname%>
		</td>
		<td>
			�ֹ����̸���
		</td>
		<td>
			<%=buyemail1%>@<%=buyemail2%>
		</td>
		<td>
			�ֹ�����ȭ��ȣ
		</td>
		<td>
			<%=buyphone1%> - <%=buyphone2%> - <%=buyphone3%>
		</td>
		<td>
			�ֹ����޴���ȭ
		</td>
		<td>
			<%=buyhp1%> - <%=buyhp2%> - <%=buyhp3%>
		</td>
	</tr>-->
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			�����μ���
		</td>
		<td>
			<%=reqname%>
		</td>
		<td>
			��������ȭ��ȣ
		</td>
		<td>
			<%=reqphone1%> - <%=reqphone2%> - <%=reqphone3%>
		</td>
		<td>
			�������޴���ȭ
		</td>
		<td>
			<%=reqhp1%>-<%=reqhp2%>-<%=reqhp3%>
		</td>
		<td>
			�������̸���
		</td>
		<td>
			<%=buyemail1%>@<%=buyemail2%>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>�ּ�</td>
		<td bgcolor="#FFFFFF" colspan=3>
			(<%= reqzipcode %>) <%= reqzipaddr %> <%= reqaddress %>
		</td>
		<td>�ֹ����ǻ���</td>
		<td bgcolor="#FFFFFF" colspan=3>
			<%= nl2br(comment) %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan=8>
			<%
			' ��üȮ�� ���� ���� ���
			if IpkumDiv < 6 then
			%>
				<input type="button" onclick="jumundetail('<%= orderno %>','')" value="�������������(���������Է�)" class="button">
			<% elseif  C_ADMIN_AUTH then %>
				<input type="button" onclick="jumundetail('<%= orderno %>','')" value="�������������(���������Է�)[������]" class="button">
			<% end if %>
		</td>
	</tr>
	</table>
	<br>
<% end if %>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<form name="frminfo" method="post">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="ipkumdiv" value="<%= ipkumdiv %>">
<input type="hidden" name="itemgubunarr">
<input type="hidden" name="itemidarr">
<input type="hidden" name="itemoptionarr">
<input type="hidden" name="masteridxarr">
<input type="hidden" name="odlvTypearr">
<input type="hidden" name="detailidxarr">
<input type="hidden" name="masteridx">
<input type="hidden" name="detailidx">
<input type="hidden" name="orderno">
<input type="hidden" name="odlvType">
<tr>
	<td align="left">
		<input type="button" value="����Ʈ��������" class="button" onclick="refer();">
		<input type="button" value="���������ΰ�ħ" class="button" onclick="location.reload(); return false;">
	</td>
	<td align="right">
		<%
		' ���Ϸ� ���� ���� ���
		if IpkumDiv < 8 then
		%>
			<input type="button" value="���û�ǰ����" class="button" onclick="jumunedit(frminfo)">
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= ojumun.FTotalCount %></b> &nbsp; �� 500 �� ���� �˻� �˴ϴ�.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>��ǰ�ڵ�</td>
	<td>�귣��ID</td>
	<td>��ǰ��[�ɼǸ�]</td>
	<td>�Ǹűݾ�</td>
	<td>�ǰ�����</td>
	<td>�Ǹż���</td>
	<td>�հ�</td>
	<td>�⺻��۱���</td>
	<td>���������</td>

	<% if ExistsBeasongYN="Y" then %>
		<td>��ۿ�û��</td>
		<td>�����</td>
		<td>��ۻ���</td>
		<td>��������</td>
	<% end if %>

	<td>���</td>
</tr>
<% if ojumun.FTotalCount>0 then %>
<%
for i=0 to ojumun.FTotalCount-1
checkblock = false

if not(ojumun.FItemList(i).fmasteridx_beasong="" or isnull(ojumun.FItemList(i).fmasteridx_beasong)) then
	ExistsItemBeasongYN="Y"
end if

'//���°� �ֹ��뺸 ���� ũ��  disabled
if ojumun.FItemList(i).FCurrState > "2" then
	checkblock = true
end if
if not(trim(ojumun.FItemList(i).fodlvType) = "" or isnull(trim(ojumun.FItemList(i).fodlvType))) then
	checkblock = true
end if
%>
<form action="" name="frmBuyPrc<%=i%>" method="get">
<input type="hidden" name="orderno" value="<%= ojumun.FItemList(i).forderno %>">
<input type="hidden" name="itemgubun" value="<%= ojumun.FItemList(i).fitemgubun %>">
<input type="hidden" name="itemid" value="<%= ojumun.FItemList(i).fitemid %>">
<input type="hidden" name="itemoption" value="<%= ojumun.FItemList(i).fitemoption %>">
<input type="hidden" name="shopid" value="<%= ojumun.FItemList(i).fshopid %>">
<input type="hidden" name="masteridx" value="<%= ojumun.FItemList(i).fmasteridx_beasong %>">
<input type="hidden" name="detailidx" value="<%= ojumun.FItemList(i).fdetailidx_beasong %>">
<input type="hidden" name="defaultbeasongdiv" value="<%= ojumun.FItemList(i).Fdefaultbeasongdiv %>">
<input type="hidden" name="comm_cd" value="<%= ojumun.FItemList(i).fcomm_cd %>">
<input type="hidden" name="currstate" value="<%= ojumun.FItemList(i).FCurrState %>">

<% if ExistsItemBeasongYN="Y" then %>
	<tr align="center" bgcolor="#FFFFFF">
<% else %>
	<tr align="center" bgcolor="#f1f1f1">
<% end if %>
	<td>
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if checkblock then response.write " disabled" %>>
	</td>
	<td>
		<%=ojumun.FItemList(i).fitemgubun%>-<%=CHKIIF(ojumun.FItemList(i).fitemid>=1000000,Format00(8,ojumun.FItemList(i).fitemid),Format00(6,ojumun.FItemList(i).fitemid))%>-<%=ojumun.FItemList(i).fitemoption%>
	</td>
	<td>
		<%=ojumun.FItemList(i).fmakerid%>
	</td>
	<td>
		<%= ojumun.FItemList(i).fitemname %>

		<% if ojumun.FItemList(i).fitemoptionname<>"" then %>
			[<%= ojumun.FItemList(i).fitemoptionname %>]
		<% end if %>
	</td>
	<td><%= FormatNumber(ojumun.FItemList(i).fsellprice,0) %></td>
	<td><%= FormatNumber(ojumun.FItemList(i).frealsellprice,0) %></td>
	<td><%= ojumun.FItemList(i).fitemno %></td>
	<td><%= FormatNumber(ojumun.FItemList(i).frealsellprice*ojumun.FItemList(i).fitemno,0) %></td>
	<td>
		<% if (ojumun.FItemList(i).Fdefaultbeasongdiv <> 0) then %>
			<%= ojumun.FItemList(i).getDefaultBeasongDivName %>
		<% end if %>
	</td>
	<td>
		<% if checkblock then %>
			<% Drawbeasonggubun "odlvType", ojumun.FItemList(i).fodlvType, " onchange='CheckThis(frmBuyPrc"& i &");' disabled" %>
		<% else %>
			<% Drawbeasonggubun "odlvType", ojumun.FItemList(i).fodlvType, " onchange='CheckThis(frmBuyPrc"& i &");'" %>
		<% end if %>
	</td>

	<% if ExistsBeasongYN="Y" then %>
		<td>
			<%= ojumun.FItemList(i).fregdate %>
		</td>
		<td>
			<%= ojumun.FItemList(i).fbeasongdate %>
		</td>
		<td>
			<font color="<%= ojumun.FItemList(i).shopNormalUpcheDeliverStateColor %>">
				<%= ojumun.FItemList(i).shopNormalUpcheDeliverState %>
			</font>
		</td>
		<td>
			<% if (ojumun.FItemList(i).Fsongjangno <> "") then %>
				<a href="<%= fnGetSongjangURL(ojumun.FItemList(i).Fsongjangdiv) %><%= ojumun.FItemList(i).Fsongjangno %>" onfocus="this.blur()" target="_blink">
				<br>[<%= DeliverDivCd2Nm(ojumun.FItemList(i).Fsongjangdiv) %>]<%= ojumun.FItemList(i).Fsongjangno %></a>
			<% end if %>
		</td>
	<% end if %>

	<td>
		<%
		'//������ , ������� �� ���
		if (IsUpcheBeasong(ojumun.FItemList(i).fodlvType) <> "Y") then
			'���Ϸ� ���� ��������
			if ojumun.FItemList(i).FCurrState < "7" then
		%>
				<input type="button" value="����" class="button" onclick="detaildel('<%= ojumun.FItemList(i).fdetailidx_beasong %>','<%=masteridx_beasong%>','<%=ojumun.FItemList(i).fodlvType%>','<%= orderno %>');">
			<% elseif  C_ADMIN_AUTH then %>
				<input type="button" value="����[������]" class="button" onclick="detaildel('<%= ojumun.FItemList(i).fdetailidx_beasong %>','<%=masteridx_beasong%>','<%=ojumun.FItemList(i).fodlvType%>','<%= orderno %>');">
			<% end if %>
		<% else %>
			<%
			'/�ֹ� Ȯ�� ���� ���¸�
			if ojumun.FItemList(i).FCurrState < "3" then
			%>
				<!--<input type="button" class="button" value="SMS" onclick="sendSMSEmail('<%'= ojumun.FItemList(i).fmakerid %>','<%'= ojumun.FItemList(i).forderno %>','<%'= ojumun.FItemList(i).fmasteridx_beasong %>','<%'= ojumun.FItemList(i).fdetailidx_beasong %>')">-->
				<input type="button" value="����" class="button" onclick="detaildel('<%= ojumun.FItemList(i).fdetailidx_beasong %>','<%=masteridx_beasong%>','<%=ojumun.FItemList(i).fodlvType%>','<%= orderno %>');">
			<% elseif  C_ADMIN_AUTH then %>
				<input type="button" value="����[������]" class="button" onclick="detaildel('<%= ojumun.FItemList(i).fdetailidx_beasong %>','<%=masteridx_beasong%>','<%=ojumun.FItemList(i).fodlvType%>','<%= orderno %>');">
			<% end if %>
		<% end if %>
	</td>
</tr>
</form>
<%
totrealprice = totrealprice + (ojumun.FItemList(i).frealsellprice*ojumun.FItemList(i).fitemno)
next
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=5>�հ�</td>
	<td><%= FormatNumber(totrealprice,0) %></td>
	<td colspan=9></td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set oedit = nothing
set ojumun = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->