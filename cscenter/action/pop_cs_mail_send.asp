<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/board/cs_templatecls.asp"-->

<%
dim email, orderserial, userid, reqhp, makerid, itemid, orderdetailidx
email 			= request("email")
orderserial 	= request("orderserial")
userid 			= request("userid")
reqhp 			= request("reqhp")
makerid 		= request("makerid")
itemid 			= request("itemid")
orderdetailidx 	= request("orderdetailidx")

if (userid = "undefined") then
	userid = ""
end if

%>

<script language='javascript'>
function SendCSMail(mailform){

	if (mailform.mailto.value.length<1){
		alert('메일주소를 입력하세요.');
		return;
	}
	if (mailform.title.value.length<1){
		alert('메일제목를 입력하세요.');
		return;
	}
	if (mailform.contents.value.length<1){
		alert('메일내용를 입력하세요.');
		return;
	}

	var ret= confirm('전송 하시겠습니까?');
	if(ret){
		mailform.submit();
	}
}

function TnCSTemplateGubunChanged(gubun) {
	var frm = document.mailform;

	var orderserial = frm.orderserial.value;
	var userid = frm.userid.value;

	var makerid = frm.makerid.value;
	var itemid = frm.itemid.value;
	var orderdetailidx = frm.orderdetailidx.value;

	CSTemplateFrame.location.href="/cscenter/board/cs_template_select_process.asp?mastergubun=20&gubun=" + gubun + "&orderserial=" + orderserial + "&userid=" + userid + "&makerid=" + makerid + "&itemid=" + itemid + "&orderdetailidx=" + orderdetailidx;
}

 function TnCSTemplateGubunProcess(v, errMSG) {
	var frm = document.mailform;

	if (errMSG != "") {
		alert(errMSG);
		frm.contents.value = "";
		return;
	}

	if(v == ''){
	}
	else{
		var t = v.split("__|__");
		frm.title.value = t[0];
		if (t.length > 1) {
			frm.contents.value = t[1];
		} else {
			alert(t.length);
		}
		// alert(v);
	}
 }

//window.resize(600,450);
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>메일발송</b>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="mailform" method="post" action="pop_cs_mail_send_process.asp">
    <tr>
    	<td width="65" bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="orderserial" value="<%= orderserial %>" size="13" maxlength="16" <% if (orderserial <> "") then %>class="text_ro" readonly<% else %>class="text"<% end if %>>
    	</td>
    	<td width="65" bgcolor="<%= adminColor("tabletop") %>">아이디</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="userid" value="<%= userid %>" size="15" maxlength="32" <% if (userid <> "") then %>class="text_ro" readonly<% else %>class="text"<% end if %>>
    	</td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="reqhp" class="text" value="<%= reqhp %>" size="13">
    	</td>
    	<td bgcolor="<%= adminColor("tabletop") %>">브랜드</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="makerid" value="<%= makerid %>" size="15" maxlength="32" <% if (makerid <> "") then %>class="text_ro" readonly<% else %>class="text"<% end if %>>
    	</td>
    </tr>
    <input type="hidden" name="orderdetailidx" value="<%= orderdetailidx %>">
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="itemid" value="<%= itemid %>" size="15" maxlength="32" <% if (itemid <> "") or True then %>class="text_ro" readonly<% else %>class="text"<% end if %>>
    	</td>
    	<td bgcolor="<%= adminColor("tabletop") %>"></td>
    	<td bgcolor="#FFFFFF">
    	</td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">메일주소</td>
    	<td colspan="3" bgcolor="#FFFFFF"><input type="text" name="mailto" class="text" value="<%= email %>" size="30"></td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">템플릿</td>
    	<td colspan="3" bgcolor="#FFFFFF">
    		<% SelectBoxCSTemplateGubun "20", "" %>
    		<iframe name="CSTemplateFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
    	</td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">메일제목</td>
    	<td colspan="3"bgcolor="#FFFFFF"><input type="text" name="title" class="text" value="" size="70"></td>
    </tr>
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">메일내용</td>
    	<td colspan="3"bgcolor="#FFFFFF"><textarea name="contents" class="textarea" value="" cols="75" rows="11"></textarea></td>
    </tr>

</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            <input type="button" class="button" value="메일발송" onclick="SendCSMail(mailform);">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->