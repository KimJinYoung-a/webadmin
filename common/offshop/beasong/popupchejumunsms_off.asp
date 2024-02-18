<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  매장 배송 주문 이메일 & 문자 발송
' History : 2012.05.10 한용민 생성
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
	response.write "	alert('해당 브랜드 정보가 없습니다');"
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
	response.write "	alert('해당 배송 정보가 없습니다');"
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

mailtitle = "[텐바이텐] " + opartner.FRectDesignerID + " 브랜드의 오프라인 배송이 (" + cstr(osheet.fitemlist(0).forderno) + ")가 접수되었습니다."

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
		
		'상품리스트 미포함한 메일발송 내용 작성
		ChgCont =""
		ChgCont = ChgCont + "<table width='600' border='0' align='center' cellpadding='0' cellspacing='0' class='a'>"
	    ChgCont = ChgCont + "<tr height='25' valign='top'>"
		ChgCont = ChgCont + "<td>"
		ChgCont = ChgCont + "<font color='red'><strong>주문서</strong>&nbsp;<b>[" + opartner.FRectDesignerID + "]</b>&nbsp;&nbsp;주문번호(" + cstr(osheet.Fitemlist(0).forderno) + ")</font></td>"
	    ChgCont = ChgCont + "</tr>"
	    ChgCont = ChgCont + "<tr valign='top'>"
		ChgCont = ChgCont + "<td>"
		ChgCont = ChgCont + "	<br>안녕하세요. 텐바이텐 입니다."
		ChgCont = ChgCont + "	<br>오프라인 어드민 <b>오프샾관리>>*[매장배송]배송요청리스트 </b>에서 주문확인 부탁드립니다."
		ChgCont = ChgCont + "	<br>"
		ChgCont = ChgCont + "	<br>브랜드ID :<b>" + opartner.FRectDesignerID + "</b>"
		ChgCont = ChgCont + "	<br>주문번호 :<b>" + cstr(osheet.Fitemlist(0).forderno) + "</b>"
		ChgCont = ChgCont + "</td>"
	    ChgCont = ChgCont + "</tr>"
        ChgCont = ChgCont + "</table>"

		'이메일 템플릿 접수
		'//실섭,테섭구분
		dim dfPath, fso, ffso, ChgCont
		IF application("Svr_Info")="Dev" THEN
			dfPath = Server.MapPath("\lib\email\mailtemplate") 		'// 테섭(scm)
		ELSE
		    dfPath = Server.MapPath("\lib\email\mailtemplate")				'// 실섭(scm)
		END IF

		'/* 파일 불러오기 */
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

		mailcontent = Replace(mailcontent,":HTMLTITLE:",mailtitle)			'메일 타이틀
		mailcontent = Replace(mailcontent,":CONTENTSHTML:",ChgCont)	'메일 본문

		'// 메일 발송
		call sendmail(mailfrom, mailto, mailtitle, mailcontent)
	end if

	sqlstr = " update db_shop.dbo.tbl_shopbeasong_order_detail" + VbCrlf
	sqlstr = sqlstr + " set upchesendsms='Y'" + VbCrlf
	sqlstr = sqlstr + " where isupchebeasong='Y'"
	sqlstr = sqlstr + " and makerid='"&makerid&"'"
	sqlstr = sqlstr + " and masteridx="&masteridx&""
	
	'response.write sqlstr &"<Br>"
	dbget.execute sqlstr

	response.write "<script>alert('전송되었습니다.');</script>"
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
	if (!confirm('이미 이메일&문자가 발송된 브랜드 입니다. 재 전송 하시겠습니까?')){ return };
<% end if %>
    
    if (frm.reqhp.value.length>15){
        alert('휴대폰 번호를 15자 이하로 입력하세요.\n핸드폰 번호는 업체정보에서 수정 가능합니다.');
        frm.reqhp.focus();
		return;
    }

    if (frm.shopphone.value.length>15){
        alert('회신번호를 15자 이하로 입력하세요.\n매장 전화번호는 매장정보에서 수정 가능합니다.');
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
<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
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

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		** 배송 담당자 연락처는 <strong>브랜드별</strong>로 변경되었습니다.
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" cellspacing="1" class="a" bgcolor=#FFFFFF cellpadding="2">
<form name="frm" method=post action="">
<input type="hidden" name="mode" value="sendall">
<input type="hidden" name="orderno" value="<%= osheet.Fitemlist(0).forderno %>">
<input type="hidden" name="masteridx" value="<%= osheet.Fitemlist(0).fmasteridx %>">
<input type="hidden" name="makerid" value="<%= osheet.Fitemlist(0).fmakerid %>">
<input type="hidden" name="mailfrom" value="<%= session("ssBctEmail") %>">
<tr>
	<td width=100>발송휴대폰</td>
	<td><input type="text" name="reqhp" value="<%= opartner.FPartnerList(0).Fdeliver_hp %>" size=16 maxlength=16></td>
</tr>
<tr>
	<td>회신전화번호</td>
	<td>
		<input type="text" name="shopphone" readonly value="<%= shopphone %>" size=16 maxlength=16>
	</td>
</tr>
<tr>
	<td>발송이메일</td>
	<td><input type="text" name="mailto" value="<%= opartner.FPartnerList(0).Fdeliver_email %>" size=30 maxlength=80></td>
</tr>
<tr>
	<td>SMS내용</td>
	<td>
		<textarea name="smstext" cols=60 rows=3>[텐바이텐] <%= opartner.FRectDesignerID %> 오프라인 배송접수. 오프샵관리>>*[매장배송]배송요청리스트</textarea>
	</td>
</tr>
<tr>
	<td colspan="2" align="center"><input type="button" value="전송" onclick="SendSMS(frm);" class="button"></td>
</tr>
</form>
</table>

<% if osheet.Fitemlist(0).fupchesendsms = "Y" then %>
	<script>alert('이미 이메일&문자가 발송된 브랜드 입니다.');</script>
<% end if %>

<%
set opartner = Nothing
set ogroup = Nothing
set osheet = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->