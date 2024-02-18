<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 매장 고객센터
' Hieditor : 2012.03.20 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/shopcscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/shopcscenter_order_cls.asp"-->
<!-- #include virtual="/admin/offshop/shopcscenter/cscenter_Function_off.asp"-->

<%
dim CsAsID ,asid ,reqname ,reqphone ,reqhp ,reqemail ,reqzipcode ,reqzipaddr ,reqetcaddr ,reqetcstr
	CsAsID = requestCheckVar(request("CsAsID"),10)

dim oCsDeliver
set oCsDeliver = new corder
	oCsDeliver.FRectCsAsID = CsAsID
	oCsDeliver.fGetOneCsDeliveryItem

if oCsDeliver.ftotalcount > 0 then
	reqname = oCsDeliver.FOneItem.Freqname
	reqphone = oCsDeliver.FOneItem.Freqphone
	reqhp = oCsDeliver.FOneItem.Freqhp
	reqemail = oCsDeliver.FOneItem.freqemail
	reqzipcode = oCsDeliver.FOneItem.Freqzipcode
	reqzipaddr = oCsDeliver.FOneItem.Freqzipaddr
	reqetcaddr = oCsDeliver.FOneItem.Freqetcaddr
	reqetcstr = oCsDeliver.FOneItem.Freqetcstr
end if
%>

<script type='text/javascript'>

function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".zipcode").value = post1 + "-" + post2;
    
    eval(frmname + ".addr1").value = addr;
    eval(frmname + ".addr2").value = dong;
}

function PopSearchZipcode(frmname) {
	var popwin = window.open("/lib/searchzip3.asp?target=" + frmname,"PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function gotowrite(){
    if(document.infoform.reqname.value == ""){
		alert("받으시는 분의 이름을 입력해주세요.");
	    document.infoform.reqname.focus();
	}

	else if(document.infoform.reqhp1.value == "" || document.infoform.reqhp2.value == "" || document.infoform.reqhp3.value == ""){
		alert("받으시는 분의 핸드폰 번호를 입력해주세요.");
	    document.infoform.reqphone1.focus();
	}

	else if(document.infoform.zipcode.value == ""){
		alert("받으시는 분의 주소를 입력해주세요.");
	    document.infoform.zipcode.focus();
	}

	else if(document.infoform.addr2.value == ""){
		alert("받으시는 분의 나머지주소를 입력해주세요.");
	    document.infoform.addr2.focus();
	}

    else{
    	if (confirm('입력 내용이 정확합니까?')){
    		document.infoform.submit();
    	}
    }

}

</script>

<table width="100%" height="500" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td valign="top">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		<form name="infoform" method="post" action="pop_CsDeliveryEdit_process.asp">
		<input type="hidden" name="asid" value="<%= CsAsID %>">
		<tr>
			<td height="10"></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td  bgcolor="C02222" height="1" colspan="4">
					</td>
				</tr>
				<tr>
					<td colspan="4" height="1" bgcolor="E5E5E5"></td>
				</tr>
				<tr>
					<td height="28" bgcolor="F8F8F8">&nbsp;</td>
					<td height="28" width=80 bgcolor="F8F8F8">*성함</td>
					<td height="28">&nbsp;</td>
					<td height="28">
						<input type="text" name="reqname" class="input_01" size="10" maxlength="16" value="<%= reqname %>" >
					</td>
				</tr>
				<tr>
					<td colspan="4" height="1" bgcolor="E5E5E5"></td>
				</tr>
				<tr>
					<td bgcolor="F8F8F8">&nbsp;</td>
					<td height="28" bgcolor="F8F8F8">연락처</td>
					<td height="3">&nbsp;</td>
					<td height="3">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td width="35">
								<input type="text" name="reqphone1" size="3"  class="input_01" maxlength="3" value="<%= splitvalue(reqphone,"-",0) %>">
							</td>
							<td width="17" align="center" class="k" style="padding:3 0 0 0">-</td>
							<td width="35">
								<input type="text" name="reqphone2" size="3" class="input_01" maxlength="4" value="<%= splitvalue(reqphone,"-",1) %>">
							</td>
							<td width="17" align="center" class="k" style="padding:3 0 0 0">-</td>
							<td width="476">
							<input type="text" name="reqphone3" size="3" class="input_01" maxlength="4" value="<%= splitvalue(reqphone,"-",2) %>">
							</td>
						</tr>
						</table>
					</td>
				</tr>				
				<tr>
					<td colspan="4" height="1" bgcolor="E5E5E5"></td>
				</tr>
				<tr>
					<td bgcolor="F8F8F8">&nbsp;</td>
					<td height="28" bgcolor="F8F8F8">*핸드폰</td>
					<td height="2">&nbsp;</td>
					<td height="2">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td width="35">
								<input type="text" name="reqhp1" size="3" class="input_01"  maxlength="3" value="<%= splitvalue(reqhp,"-",0) %>">
							</td>
							<td width="17" align="center" class="k" style="padding:3 0 0 0">-</td>
							<td width="35">
								<input type="text" name="reqhp2" size="3" class="input_01"  maxlength="4" value="<%= splitvalue(reqhp,"-",1) %>">
							</td>
							<td width="17" align="center" class="k" style="padding:3 0 0 0">-</td>
							<td width="476">
								<input type="text" name="reqhp3" size="3" class="input_01"  maxlength="4" value="<%= splitvalue(reqhp,"-",2) %>">
							</td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td colspan="4" height="1" bgcolor="E5E5E5"></td>
				</tr>
				<tr>
					<td bgcolor="F8F8F8">&nbsp;</td>
					<td height="28" bgcolor="F8F8F8">이메일</td>
					<td height="2">&nbsp;</td>
					<td height="2">
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td width="35">
								<input type="text" class="input_01" name="reqemail" value="<%= reqemail %>" size=40 maxlength=50>
							</td>
						</tr>
						</table>
					</td>
				</tr>				
				<tr>
					<td colspan="4" height="1" bgcolor="E5E5E5"></td>
				</tr>
				<tr>
					<td bgcolor="F8F8F8">&nbsp;</td>
					<td height="90" bgcolor="F8F8F8"><img src="http://fiximage.10x10.co.kr/web2007/my1010/pop_em07.gif" width="58" height="15"></td>
					<td height="4">&nbsp;</td>
					<td height="4">
						<input name="zipcode" class="input_01"  id="zipcode" style="background-color:#EEEEEE;" value="<%= oCsDeliver.FOneItem.Freqzipcode %>" size="7" readonly>
				        <input type="button" class="button" value="검색" onClick="FnFindZipNew('infoform','E')">
						<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('infoform','E')">
				        <% '<input type="button" value="검색(구)" class="button" onclick="PopSearchZipcode('infoform');" onFocus="this.blur();"> %>
				        <br>
				        <input name="addr1" type="text" class="input_01" id="addr1" style="background-color:#EEEEEE;" value="<%= oCsDeliver.FOneItem.Freqzipaddr %>" size="50" readonly>
				        <input name="addr2" type="text" class="input_01" id="addr2" style="ime-mode:active" value="<%= oCsDeliver.FOneItem.Freqetcaddr %>" size="50" maxlength="80">
					</td>
				</tr>
				<tr>
					<td colspan="4" height="1" bgcolor="E5E5E5"></td>
				</tr>
				<tr>
					<td bgcolor="F8F8F8">&nbsp;</td>
					<td height="80" bgcolor="F8F8F8">기타</td>
					<td height="9">&nbsp;</td>
					<td height="9"><textarea name="reqetc" class="webtextarea" cols="45" rows="3" ><%= reqetcstr %></textarea></td>
				</tr>
				<tr>
					<td colspan="4" height="1" bgcolor="C02222"></td>
				</tr>
				<tr align="center">
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td height="35">&nbsp;</td>
					<td height="35">
						<input type="button" class="button" value=" 저 장 " onclick="gotowrite();">
						<input type="button" class="button" value=" 취 소 " onclick="window.close();">
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<%
set oCsDeliver = Nothing
%>
<!-- #include virtual="/admin/offshop/shopcscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->