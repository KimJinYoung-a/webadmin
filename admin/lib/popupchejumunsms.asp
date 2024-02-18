<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체개별주문서관리NEW
' History : 이상구 생성
'			2020.07.23 한용민 수정(이메일발송. 메일러로 변경)
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

mailtitle = "[텐바이텐] " + opartner.FRectDesignerID + " 브랜드의 주문서 (" + osheet.FOneItem.Fbaljucode + ")가 접수되었습니다."

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

		'상품리스트 미포함한 메일발송 내용 작성
		ChgCont =""
		ChgCont = ChgCont + "<table width='600' border='0' align='center' cellpadding='0' cellspacing='0' class='a'>"
	    ChgCont = ChgCont + "<tr height='25' valign='top'>"
		ChgCont = ChgCont + "<td>"
		ChgCont = ChgCont + "<font color='red'><strong>주문서</strong>&nbsp;<b>[" + opartner.FRectDesignerID + "]</b>&nbsp;&nbsp;주문코드(" + osheet.FOneItem.Fbaljucode + ")</font></td>"
	    ChgCont = ChgCont + "</tr>"
	    ChgCont = ChgCont + "<tr valign='top'>"
		ChgCont = ChgCont + "<td>"
		ChgCont = ChgCont + "	<br>안녕하세요. 텐바이텐입니다."
		ChgCont = ChgCont + "	<br>어드민 <b>주문서/입출내역관리>>주문서관리</b>에서 주문확인 후 입고 부탁드립니다."
		ChgCont = ChgCont + "	<br>"
		ChgCont = ChgCont + "	<br>브랜드ID :<b>" + opartner.FRectDesignerID + "</b>"
		ChgCont = ChgCont + "	<br>주문코드 :<b>" + osheet.FOneItem.Fbaljucode + "</b>"
		ChgCont = ChgCont + "	<br><a href='http://scm.10x10.co.kr/'>어드민 바로가기>><a>"
		ChgCont = ChgCont + "	<br>"
		ChgCont = ChgCont + "	<br><b><font color='red'>[주문확인]</font></b>"
		ChgCont = ChgCont + "	<br>주문서를 확인하신 후에는 꼭 주문서관리에서 <b>[주문확인]</b>으로 전환하여주시고,"
		ChgCont = ChgCont + "	<br>수량이 부족하거나 단종인 경우, 물류센터로 연락을 주시거나,"
		ChgCont = ChgCont + "	<br>주문확정수량을 조정하여주시기 바랍니다."
		ChgCont = ChgCont + "	<br>"
		ChgCont = ChgCont + "	<br><b><font color='red'>[출고완료]</font></b>"
		ChgCont = ChgCont + "	<br>출고하실때는 수량변경이 있을경우, 확정수량을 조정하여 주시고,"
		ChgCont = ChgCont + "	<br><b>[출고완료]</b>로 설정후, 운송장 번호를 입력하여 주시기 바랍니다."
		ChgCont = ChgCont + "	<br><br><br>물류 센터 전화 번호가 변경되었습니다. (1644-1851)"
		ChgCont = ChgCont + "</td>"
	    ChgCont = ChgCont + "</tr>"
        ChgCont = ChgCont + "</table>"

		'이메일 템플릿 접수
		'//실섭,테섭구분
		dim dfPath, fso, ffso, ChgCont
		IF application("Svr_Info")="Dev" THEN
			dfPath = Server.MapPath("\lib\email\mailtemplate")				'// 실섭(scm)
		ELSE
		    dfPath = Server.MapPath("\lib\email\mailtemplate")				'// 실섭(scm)
		END IF

		'/* 파일 불러오기 */
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

		mailcontent = Replace(mailcontent,":HTMLTITLE:",mailtitle)			'메일 타이틀
		mailcontent = Replace(mailcontent,":CONTENTSHTML:",ChgCont)	'메일 본문

		set oMail = New MailCls

		IF mailto<>"" THEN

			oMail.MailTitles	= mailtitle
			oMail.SenderNm		= "텐바이텐"
			'oMail.SenderMail	= "mailzine@10x10.co.kr"
			oMail.SenderMail	= "customer@10x10.co.kr"
			oMail.AddrType		= "string"
			oMail.ReceiverNm	= mailto
			oMail.ReceiverMail	= mailto
			oMail.MailConts 	= mailcontent
			oMail.MailerMailGubun = 13		' 메일러 자동메일 번호

			oMail.Send_TMSMailer()		'TMS메일러
			'oMail.Send_Mailer()		'EMS메일러
			''oMail.Send_CDO()	'cdo
		End IF

		SET oMail = nothing
	end if

	sqlstr = " update [db_storage].[dbo].tbl_ordersheet_master" + VbCrlf
	sqlstr = sqlstr + " set sendsms='Y'" + VbCrlf
	sqlstr = sqlstr + " where idx=" + Cstr(idx) + VbCrlf
	dbget.execute sqlstr

	response.write "<script>alert('전송되었습니다.');</script>"
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
	if (!confirm('이미 전송된 주문 입니다. 재 전송 하시겠습니까?')){ return };
<% end if %>

    if (frm.reqhp.value.length>15){
        alert('핸드폰 번호를 15자 이하로 입력하세요.\n핸드폰 번호는 업체정보에서 수정 가능합니다.');
        frm.reqhp.focus();
		return;
    }

	if ((frm.reqhp.value.length<1)&&(frm.mailto.value.length<1)){
		alert('핸드폰 번호나 이메일주소 중 하나는 입력되어야 합니다.');
		return;
	}

	var ret= confirm('전송 하시겠습니까?');
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
	<td colspan=5>대표전화 : <%= opartner.FPartnerList(0).Ftel %> 팩스 : <%= opartner.FPartnerList(0).Ffax %></td>
</tr>

<tr bgcolor="#DDDDFF">
	<td width=80>구분</td>
	<td width=80>성명</td>
	<td width=80>전화</td>
	<td width=80>핸드폰</td>
	<td width=*>이메일</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td ><a href="#" onClick="CopyInfo('<%= ogroup.FOneItem.Fmanager_hp %>','<%= ogroup.FOneItem.Fmanager_email %>');">그룹담당자</a></td>
	<td ><%= ogroup.FOneItem.Fmanager_name %></td>
	<td ><%= ogroup.FOneItem.Fmanager_phone %></td>
	<td ><%= ogroup.FOneItem.Fmanager_hp %></td>
	<td ><%= ogroup.FOneItem.Fmanager_email %></td>
</tr>
<!-- 배송담당자는 브랜드별
<tr bgcolor="#FFFFFF">
	<td ><a href="#" onClick="CopyInfo('<%= ogroup.FOneItem.Fdeliver_hp %>','<%= ogroup.FOneItem.Fdeliver_email %>');">그룹배송담당자</a></td>
	<td ><%= ogroup.FOneItem.Fdeliver_name %></td>
	<td ><%= ogroup.FOneItem.Fdeliver_phone %></td>
	<td ><%= ogroup.FOneItem.Fdeliver_hp %></td>
	<td ><%= ogroup.FOneItem.Fdeliver_email %></td>
</tr>
 -->

<tr bgcolor="#FFFFFF">
	<td ><a href="#" onClick="CopyInfo('<%= opartner.FPartnerList(0).Fmanager_hp %>','<%= opartner.FPartnerList(0).Femail %>');">담당자</a></td>
	<td ><%= opartner.FPartnerList(0).Fmanager_name %></td>
	<td ><%= opartner.FPartnerList(0).Fmanager_phone %></td>
	<td ><%= opartner.FPartnerList(0).Fmanager_hp %></td>
	<td ><%= opartner.FPartnerList(0).Femail %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td ><a href="#" onClick="CopyInfo('<%= opartner.FPartnerList(0).Fdeliver_hp %>','<%= opartner.FPartnerList(0).Fdeliver_email %>');">브랜드배송담당자</a></td>
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
        ** 배송 담당자 연락처는 <strong>브랜드별</strong>로 변경되었습니다.
    </td>
</tr>
<tr>
	<td width=100>핸드폰</td>
	<td><input type="text" name="reqhp" value="<%= opartner.FPartnerList(0).Fdeliver_hp %>" size=16 maxlength=16></td>
</tr>
<tr>
	<td width=100>이메일</td>
	<td><input type="text" name="mailto" value="<%= opartner.FPartnerList(0).Fdeliver_email %>" size=30 maxlength=80></td>
</tr>
<tr>
	<td width=100>SMS내용</td>
	<td>
	<textarea name="smstext" cols=60 rows=3>[텐바이텐]<%= opartner.FRectDesignerID %>주문서가 접수되었습니다.주문서관리에서 확인후 입고해주세요.</textarea>
	</td>
</tr>
<tr>
	<td colspan="2" align=center><input type="button" value=" 전 송 " onclick="SendSMS(frm);"></td>
</tr>
</table>
</form>

<% if osheet.FOneItem.IsSendedSMS then %>
	<script type='text/javascript'>
		alert('이미 SMS전송된 주문입니다.');
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
