<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.24 한용민 생성
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
	response.write "	alert('주문번호가 없습니다');"
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

	// 주문인증정보수정
	function certedit(orderno,masteridx_beasong,vmode){
		if (vmode==''){
			alert('구분자가 없습니다.');
			return;
		}

		frmsmscert.masteridx.value=masteridx_beasong;
		frmsmscert.orderno.value=orderno;
		frmsmscert.action = '/common/offshop/beasong/shopbeasong_process.asp';

		if (vmode=='ReSendKakaotalk'){
			if (confirm('카카오톡을 발송 하시겠습니까?')){
				frmsmscert.mode.value=vmode;
				frmsmscert.submit();
			}
		}else if (vmode=='ReSendSMS'){
			if (confirm('SMS를 발송 하시겠습니까?')){
				frmsmscert.mode.value=vmode;
				frmsmscert.submit();
			}
		}else{
			if (confirm('주문인증정보를 수정 하시겠습니까?')){
				frmsmscert.mode.value=vmode;
				frmsmscert.submit();
			}
		}
	}
	
	// 배송지정보수정. 이문재 이사님 요청으로 팝업으로 변경
	function jumundetail(orderno,masteridx_beasong){
		var popwin = window.open('/common/offshop/beasong/shopjumun_address.asp?mode=addressedit&orderno='+orderno+'&masteridx='+masteridx_beasong+'&menupos=<%=menupos%>','popbeasongedit','width=1280,height=960,scrollbars=yes,resizable=yes');
		popwin.focus();

//		frminfo.masteridx.value=masteridx_beasong;
//		frminfo.orderno.value=orderno;
//		frminfo.mode.value='addressedit';
//		frminfo.action = '/common/offshop/beasong/shopjumun_address.asp';
//		frminfo.submit();
	}

	//이전페이지로
	function refer(){
		location.href='/common/offshop/beasong/shopbeasong_list.asp?menupos=<%= menupos %>';
	}

	//상품 삭제
	function detaildel(detailidx_beasong,masteridx_beasong,odlvType,orderno){
		frminfo.orderno.value=orderno;
		frminfo.odlvType.value=odlvType;
		frminfo.detailidx.value=detailidx_beasong;
		frminfo.masteridx.value=masteridx_beasong;
		frminfo.mode.value='detaildel';
		frminfo.action='/common/offshop/beasong/shopbeasong_process.asp';
		frminfo.submit();
	}

	//상품수정
	function jumunedit(upfrm){
		var masteridx_beasong = '<%= masteridx_beasong %>';
		var orderno = '<%= orderno %>';

		upfrm.detailidxarr.value='';
		upfrm.masteridx.value='';
		upfrm.odlvTypearr.value='';

		if (!CheckSelected()){
			alert('선택아이템이 없습니다.');
			return;
		}

		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					if (frm.odlvType.value==''){
						alert('현장수령빼고 배송구분을 선택 하세요.');
						frm.odlvType.focus();
						return;
					}
					// comm_cd : B031 매입출고정산 / B012 업체특정 / B013 출고특정
					// 물류배송
					if (frm.odlvType.value*1 == '1') {
						if (frm.comm_cd.value == 'B012') {
							alert("해당 상품은 업체배송 or 매장배송만 가능 합니다.");
							frm.odlvType.focus();
							return;
						}
					}
					// 업체배송
					if (frm.odlvType.value*1 == '2') {
						if (frm.comm_cd.value == 'B031' || frm.comm_cd.value == 'B013') {
							alert("해당 상품은 물류배송 or 매장배송만 가능 합니다.");
							frm.odlvType.focus();
							return;
						}
					}

/*
					if (frm.defaultbeasongdiv.value*1 == 0) {
						if (frm.odlvType.value*1 != 0) {
							alert("지정할수 없는 배송자입니다. 매장배송을 선택하세요.");
							frm.odlvType.focus();
							return;
						}
					}
*/
					if (frm.currstate.value*1 > 3) {
						alert("해당 상품은 이미 출고 완료된 상품 입니다.");
						frm.odlvType.focus();
						return;
					}
					if (frm.currstate.value*1 > 2) {
						alert("[참고]해당 상품은 이미 업체에서 배송을 확인한 상태 입니다.");
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
			주문인증정보
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			휴대폰번호(주문시입력)
		</td>
		<td>
			<input type="text" name="UserHp1" value="<%= UserHp1 %>" size=4 maxlength=4>-<input type="text" name="UserHp2" value="<%= UserHp2 %>" size=4 maxlength=4>-<input type="text" name="UserHp3" value="<%= UserHp3 %>" size=4 maxlength=4>
		</td>
		<td>
			카카오톡발송여부
		</td>
		<td>
			<%= KakaoTalkYN %>
			
			<%
			' 업체통보 이전 상태 라면
			if IpkumDiv < 5 then
			%>
				<input type="button" class="button" value="발송" onclick="certedit('<%= orderno %>','','ReSendKakaotalk')">
				<!--<input type="button" class="button" value="발송" onclick="alert('작업중\nSMS로 발송하세요.'); return false;">-->
			<% elseif  C_ADMIN_AUTH then %>
				<input type="button" class="button" value="발송[관리자]" onclick="certedit('<%= orderno %>','','ReSendKakaotalk')">
			<% end if %>
		</td>
		<td>
			SMS발송여부
		</td>
		<td>
			<%= SmsYN %>

			<%
			' 업체통보 이전 상태 라면
			if IpkumDiv < 5 then
			%>
				<input type="button" class="button" value="발송" onclick="certedit('<%= orderno %>','','ReSendSMS')">
			<% elseif  C_ADMIN_AUTH then %>
				<input type="button" class="button" value="발송[관리자]" onclick="certedit('<%= orderno %>','','ReSendSMS')">
			<% end if %>
		</td>
		<td>
			배송여부
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
					관리자권한 : <% response.write "https://m.10x10.co.kr/my10x10/order/myshoporder.asp?orderNo="& trim(orderno) &"&certNo="& confirmcertno &"" %>
				</td>
			</tr>
	<%
		end if
	end if
	%>

	<tr align="center" bgcolor="#FFFFFF">
		<td colspan=8>
			<%
			' 업체확인 이전 상태 라면
			if IpkumDiv < 6 then
			%>
				<input type="button" onclick="certedit('<%= orderno %>','','certedit')" value="주문인증정보수정" class="button">
			<% elseif  C_ADMIN_AUTH then %>
				<input type="button" onclick="certedit('<%= orderno %>','','certedit')" value="주문인증정보수정[관리자]" class="button">
			<% end if %>
		</td>
	</tr>
	</table>
	</form>

	<br>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan=8>
			배송정보
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			주문번호
		</td>
		<td>
			<%= orderno %>
		</td>
		<td>
			판매매장
		</td>
		<td>
			<%=shopname%>
		</td>
		<td>
			출고상태
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
			주문자성함
		</td>
		<td>
			<%=buyname%>
		</td>
		<td>
			주문자이메일
		</td>
		<td>
			<%=buyemail1%>@<%=buyemail2%>
		</td>
		<td>
			주문자전화번호
		</td>
		<td>
			<%=buyphone1%> - <%=buyphone2%> - <%=buyphone3%>
		</td>
		<td>
			주문자휴대전화
		</td>
		<td>
			<%=buyhp1%> - <%=buyhp2%> - <%=buyhp3%>
		</td>
	</tr>-->
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			수령인성함
		</td>
		<td>
			<%=reqname%>
		</td>
		<td>
			수령인전화번호
		</td>
		<td>
			<%=reqphone1%> - <%=reqphone2%> - <%=reqphone3%>
		</td>
		<td>
			수령인휴대전화
		</td>
		<td>
			<%=reqhp1%>-<%=reqhp2%>-<%=reqhp3%>
		</td>
		<td>
			수령인이메일
		</td>
		<td>
			<%=buyemail1%>@<%=buyemail2%>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td>주소</td>
		<td bgcolor="#FFFFFF" colspan=3>
			(<%= reqzipcode %>) <%= reqzipaddr %> <%= reqaddress %>
		</td>
		<td>주문유의사항</td>
		<td bgcolor="#FFFFFF" colspan=3>
			<%= nl2br(comment) %>
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan=8>
			<%
			' 업체확인 이전 상태 라면
			if IpkumDiv < 6 then
			%>
				<input type="button" onclick="jumundetail('<%= orderno %>','')" value="배송지정보수정(직원수기입력)" class="button">
			<% elseif  C_ADMIN_AUTH then %>
				<input type="button" onclick="jumundetail('<%= orderno %>','')" value="배송지정보수정(직원수기입력)[관리자]" class="button">
			<% end if %>
		</td>
	</tr>
	</table>
	<br>
<% end if %>

<!-- 액션 시작 -->
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
		<input type="button" value="리스트페이지로" class="button" onclick="refer();">
		<input type="button" value="페이지새로고침" class="button" onclick="location.reload(); return false;">
	</td>
	<td align="right">
		<%
		' 출고완료 이전 상태 라면
		if IpkumDiv < 8 then
		%>
			<input type="button" value="선택상품수정" class="button" onclick="jumunedit(frminfo)">
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= ojumun.FTotalCount %></b> &nbsp; ※ 500 건 까지 검색 됩니다.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>상품코드</td>
	<td>브랜드ID</td>
	<td>상품명[옵션명]</td>
	<td>판매금액</td>
	<td>실결제액</td>
	<td>판매수량</td>
	<td>합계</td>
	<td>기본배송구분</td>
	<td>배송자지정</td>

	<% if ExistsBeasongYN="Y" then %>
		<td>배송요청일</td>
		<td>배송일</td>
		<td>배송상태</td>
		<td>송장정보</td>
	<% end if %>

	<td>비고</td>
</tr>
<% if ojumun.FTotalCount>0 then %>
<%
for i=0 to ojumun.FTotalCount-1
checkblock = false

if not(ojumun.FItemList(i).fmasteridx_beasong="" or isnull(ojumun.FItemList(i).fmasteridx_beasong)) then
	ExistsItemBeasongYN="Y"
end if

'//상태가 주문통보 보다 크면  disabled
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
		'//매장배송 , 물류배송 의 경우
		if (IsUpcheBeasong(ojumun.FItemList(i).fodlvType) <> "Y") then
			'출고완료 이전 삭제가능
			if ojumun.FItemList(i).FCurrState < "7" then
		%>
				<input type="button" value="삭제" class="button" onclick="detaildel('<%= ojumun.FItemList(i).fdetailidx_beasong %>','<%=masteridx_beasong%>','<%=ojumun.FItemList(i).fodlvType%>','<%= orderno %>');">
			<% elseif  C_ADMIN_AUTH then %>
				<input type="button" value="삭제[관리자]" class="button" onclick="detaildel('<%= ojumun.FItemList(i).fdetailidx_beasong %>','<%=masteridx_beasong%>','<%=ojumun.FItemList(i).fodlvType%>','<%= orderno %>');">
			<% end if %>
		<% else %>
			<%
			'/주문 확인 이전 상태만
			if ojumun.FItemList(i).FCurrState < "3" then
			%>
				<!--<input type="button" class="button" value="SMS" onclick="sendSMSEmail('<%'= ojumun.FItemList(i).fmakerid %>','<%'= ojumun.FItemList(i).forderno %>','<%'= ojumun.FItemList(i).fmasteridx_beasong %>','<%'= ojumun.FItemList(i).fdetailidx_beasong %>')">-->
				<input type="button" value="삭제" class="button" onclick="detaildel('<%= ojumun.FItemList(i).fdetailidx_beasong %>','<%=masteridx_beasong%>','<%=ojumun.FItemList(i).fodlvType%>','<%= orderno %>');">
			<% elseif  C_ADMIN_AUTH then %>
				<input type="button" value="삭제[관리자]" class="button" onclick="detaildel('<%= ojumun.FItemList(i).fdetailidx_beasong %>','<%=masteridx_beasong%>','<%=ojumun.FItemList(i).fodlvType%>','<%= orderno %>');">
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
	<td colspan=5>합계</td>
	<td><%= FormatNumber(totrealprice,0) %></td>
	<td colspan=9></td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set oedit = nothing
set ojumun = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->