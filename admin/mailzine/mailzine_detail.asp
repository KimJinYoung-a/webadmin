<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mailzinecls.asp"-->
<%
CONST MAXHeightPX = 1400    '''이 수치에 대해서는 확실하지 않음.. (2,000px 보다 작은 사이즈 2개를 넣었을때 도 깨진경우가 있음)

dim idx, code, omail ,yyyy1, mm1, dd1 , tmp , area, mngUserid, mailergubun
dim title,regdate,img1,img2,img3,img4,imgmap1,imgmap2,imgmap3,imgmap4,isusing,gubun,memgubun,secretGubun,reservationDATE
	idx = requestcheckvar(getNumeric(request("idx")),10)
	mailergubun = requestcheckvar(request("mailergubun"),16)

if mailergubun="" or isnull(mailergubun) then
	response.write "메일러 구분이 없습니다."
	dbget.close() : response.end
end if

set omail = new CMailzineList
	omail.frectidx = idx
	omail.frectmailergubun = mailergubun

	'//idx 값이 있을경우에만 쿼리(수정모드)
	if idx <> "" then
		omail.MailzineDetail()

		if omail.ftotalcount > 0 then
			title = omail.foneitem.ftitle
			regdate = omail.foneitem.fregdate
			img1 = omail.foneitem.fimg1
			img2 = omail.foneitem.fimg2
			img3 = omail.foneitem.fimg3
			img4 = omail.foneitem.fimg4
			imgmap1 = omail.foneitem.fimgmap1
			imgmap2 = omail.foneitem.fimgmap2
			imgmap3 = omail.foneitem.fimgmap3
			imgmap4 = omail.foneitem.fimgmap4
			isusing = omail.foneitem.fisusing
			gubun = omail.foneitem.fgubun
			area = omail.foneitem.farea
			mngUserid = omail.foneitem.fmngUserid
			memgubun = omail.foneitem.fmemgubun
			secretGubun = omail.foneitem.fsecretGubun
			reservationDATE = omail.foneitem.freservationDATE
			tmp = split(omail.foneitem.fregdate,".")
			yyyy1 = tmp(0)
			mm1 = tmp(1)
			dd1 = tmp(2)
			code = mm1 & dd1
		end if
	end if

if area = "" then area = "ten_all"
if isusing = "" then isusing = "N"
if memgubun = "" then memgubun ="member_all"
If secretGubun = "" then secretGubun="N"
%>
<style> 
#mask {  
	position:absolute;  
	z-index:9000;  
	background-color:#000;  
	display:none;  
	left:0;
	top:0;
} 
.window{
	display: none;
	position:absolute;  
	left:100px;
	bottom:10px;
	z-index:10000;
}
</style> 
<script type="text/javascript" src="/js/jquery-latest.js"></script>
<script type="text/javascript">

function wrapWindowByMask(){
	//화면의 높이와 너비를 구한다.
	var maskHeight = $(document).height();  
	var maskWidth = $(window).width();  

	//마스크의 높이와 너비를 화면 것으로 만들어 전체 화면을 채운다.
	$('#mask').css({'width':maskWidth,'height':maskHeight});  

	//애니메이션 효과 - 일단 1초동안 까맣게 됐다가 80% 불투명도로 간다.
	$('#mask').fadeIn(1000);      
	$('#mask').fadeTo("slow",0.8);    

	//윈도우 같은 거 띄운다.
	$('.window').show();
}

$(document).ready(function(){
	//검은 막 띄우기
	$('.openMask').click(function(e){
		e.preventDefault();
		wrapWindowByMask();
	});

	//닫기 버튼을 눌렀을 때
	$('.window .close').click(function (e) {  
	    //링크 기본동작은 작동하지 않도록 한다.
	    e.preventDefault();  
	    $('#mask, .window').hide();
	});
});
</script>
<script language="JavaScript">
	<% if date() >= "2017-11-13" and date() < "2017-11-18" then %>
		alert('[매우중요]\n\n 11월 20일 이후 \n메일발송 담당자는\n김진영 대리 입니다.\n\n 11월 20일이후 메일발송은 \n김진영대리에게 꼭!! 알려주세요.');
	<% end if %>

	alert('[매우중요]\n이미지 개당 높이는 1,400 px 미만\n\n이미지맵 태크 등록시 한줄 내려서 등록\n\n이미지맵 타켓 target="_top"\n\n이미지맵 이름 수정불가');
	function checkok(frm){
		if (document.modify.gubun.value == "1"){
			if (modify.isusing.value==''){
				alert('노출여부를 선택해주세요');
				modify.isusing.focus();
				return;
			}
			if ("<%= hour(now()) %>" >= 18){
				alert('※ 18시 이후\n※ 발송 담당자에게\n※ 완성 여부를 알리지 않았을경우\n※ 메일 발송이 되지 않으며\n※ 이에대한 책임은 최종수정자가 지게 됩니다.');
			}
			frm.submit();
			document.getElementById('goproc').disabled = true;
		}else{
			/* 시간체크 */
			if ("<%= hour(now()) %>" <= 17){
				if (confirm("작성상황이 완성입니다. 확인버튼을 누르시면 수정이 불가합니다.\n저장하시겠습니까?") == true) {
					frm.submit();
				}else{
					return false;
				}
			}else{
				$('.openMask2').click(function(e){
					e.preventDefault();
					wrapWindowByMask();
				});
			}
		}
	}

	//메일 발송 완료 후 1,2,3,4 이미지 수정
	function imageedit(frm){
		if(confirm("정말 이미지를 수정 하시겠습니까?")){
			if(confirm("잘못된 이미지를 업로드 할 경우\n\n발송된 메일에 문제가 생길 수 있습니다.")){
				frm.mode.value='imageedit';
				frm.submit();
				document.getElementById('goimgproc').disabled = true;
			}
		}
		
	}

	function LastConfirm(yn){
		var frm = document.modify;
		if (yn == 'Y'){
			if (confirm("확인버튼을 누르시면 수정이 불가합니다.\n저장하시겠습니까?") == true) {
				frm.submit();
			}else{
				return false;
			}
		}else{
			 $('#mask, .window').hide();
		}
	}

	function delimg(imgnumber){
		frm_mail.action = '/admin/mailzine/mailzine_process.asp';
		frm_mail.imgnumber.value=imgnumber;
		frm_mail.mode.value='imgdel';
		frm_mail.submit();
	}

</script>

<table width="95%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">

<% IF application("Svr_Info")="Dev" THEN %>
	<form name="modify" method="post" action="<%=mailzine%>/ftp/mailzine_input_ok.asp" enctype="multipart/form-data" style="margin:0px;">
<% else %>
	<form name="modify" method="post" action="https://omailzine.10x10.co.kr/ftp/mailzine_input_ok.asp" enctype="multipart/form-data" style="margin:0px;">
<% end if %>

<input type="hidden" name="idx" value="<% = idx %>">
<input type="hidden" name="mailergubun" value="<%= mailergubun %>">
<input type="hidden" name="mode">
<input type="hidden" name="img1editname" value="<%= img1 %>">
<input type="hidden" name="img2editname" value="<%= img2 %>">
<input type="hidden" name="img3editname" value="<%= img3 %>">
<input type="hidden" name="img4editname" value="<%= img4 %>">
<tr bgcolor="#FFFFFF">
	<td colspan=2>
		<br>
		<font size=4>※ 주의사항. 반드시 지켜 주셔야 합니다. 매우중요!!</font>
		<br>&nbsp;&nbsp;&nbsp;이미지맵 타켓 <font color="red">target="_top"</font> 으로 주시고, 이미지맵 <font color="red">이름</font>은 고치지 말아주세요.
		<br>&nbsp;&nbsp;&nbsp;이미지 사이즈 개당 높이 <font color="red"><%= FormatNumber(MAXHeightPX,0) %> px</font> 가 넘을경우, 아웃룩에서 짤립니다.
		<br>&nbsp;&nbsp;&nbsp;이미지맵 태크 등록시 <font color="red">한줄 내려서</font> 등록해 주시기 바랍니다. 맵이 길경우 일부포털에서 짤립니다.
		<br>&nbsp;&nbsp;&nbsp;ex)
		<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(map name="ImgMap1")
		<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(area shape="rect" coords="6,5,130,114" href="http://www.10x10.co.kr" target="_top" onfocus="blur()")
		<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(area shape="rect" coords="11,136,694,665" href="http://www.10x10.co.kr" target="_top" onfocus="blur()")
		<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(/map)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">메일제목</td>
	<td><input type="text" name="title" class="input" size="55" value="<% = title %>"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">메일진 등록일</td>
	<td><% DrawOneDateBox_2012 yyyy1,mm1,dd1 %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">담당웹디</td>
	<td><%sbGetDesignerid "mngUserid",mngUserid, ""%></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">디자인작성상황</td>
	<td>
		<select name="gubun" class="select">
			<option value="1" <% if gubun = "1" then response.write "selected"%>>미완성</option>
			<option value="5" <% if gubun = "5" then response.write "selected"%>>완성</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">시스템팀_예약일</td>
	<td><%= Chkiif(isnull(reservationDATE), "예약전", reservationDATE) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">1번이미지</td>
	<td>
		<input type="file" name="img1" class="input" size="40"> <strong>(height <%= FormatNumber(MAXHeightPX,0) %> px 미만)</strong>
		<br>
		<%
		if img1 <> "" then
			response.write mailzine&"/"&yyyy1&"/"&img1
		%>
			<input type="button" onclick="delimg('1');" class="button" value="이미지삭제">
		<%
		end if
		%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">이미지맵 코드넣기</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
	   <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		   <tr>
				<td>
					<% if imgmap1 = "" then %>
						<textarea name="imagemap1" rows="10" class="textarea" style="width:100%;"><map name="ImgMap1"></map></textarea>
					<% else %>
						<textarea name="imagemap1" rows="10" class="textarea" style="width:100%;"><% = imgmap1 %></textarea>
					<% end if %>
				</td>
		   </tr>
	   </table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">2번이미지</td>
	<td>
		<input type="file" name="img2" class="input" size="40"> <strong>(height <%= FormatNumber(MAXHeightPX,0) %> px 미만)</strong>
		<br>
		<%
		if img2 <> "" then
			response.write mailzine&"/"&yyyy1&"/"&img2
		%>
			<input type="button" onclick="delimg('2');" class="button" value="이미지삭제">
		<%
		end if
		%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">이미지맵 코드넣기</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
	   <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	   <tr>
		<td>
			<% if imgmap2 = "" then %>
				<textarea name="imagemap2" rows="10" class="textarea" style="width:100%;"><map name="ImgMap2"></map></textarea>
			<% else %>
				<textarea name="imagemap2" rows="10" class="textarea" style="width:100%;"><%= imgmap2 %></textarea>
			<% end if %>
		</td>
	   </tr>
	   </table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">3번이미지</td>
	<td>
		<input type="file" name="img3" class="input" size="40"> <strong>(height <%= FormatNumber(MAXHeightPX,0) %> px 미만)</strong>
		<br>
		<%
		if img3 <> "" then
			response.write mailzine&"/"&yyyy1&"/"&img3
		%>
			<input type="button" onclick="delimg('3');" class="button" value="이미지삭제">
		<%
		end if
		%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">이미지맵 코드넣기</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
	   <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	   <tr>
		<td>
			<% if imgmap3 = "" then %>
				<textarea name="imagemap3" rows="10" class="textarea" style="width:100%;"><map name="ImgMap3"></map></textarea>
			<% else %>
				<textarea name="imagemap3" rows="10" class="textarea" style="width:100%;"><%= imgmap3 %></textarea>
			<% end if %>
		</td>
	   </tr>
	   </table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">4번이미지</td>
	<td>
		<input type="file" name="img4" class="input" size="40"> <strong>(height <%= FormatNumber(MAXHeightPX,0) %> px 미만)</strong>
		<br>
		<%
		if img4 <> "" then
			response.write mailzine&"/"&yyyy1&"/"&img4
		%>
			<input type="button" onclick="delimg('4');" class="button" value="이미지삭제">
		<%
		end if
		%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">이미지맵 코드넣기</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
	   <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
	   <tr>
		<td>
			<% if imgmap4 = "" then %>
				<textarea name="imagemap4" rows="10" class="textarea" style="width:100%;"><map name="ImgMap4"></map></textarea>
			<% else %>
				<textarea name="imagemap4" rows="10" class="textarea" style="width:100%;"><%= imgmap4 %></textarea>
			<% end if %>
		</td>
	   </tr>
	   </table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">발송지역</td>
	<td>
		<% Drawareagubun "area" , area , "class='select'" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">발송회원등급</td>
	<td>
		<% DrawMemberGubun "memgubun" , memgubun , "class='select'" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">사이트노출</td>
	<td>
		<% Drawisusing "isusing" , isusing , "class='select'" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">시크릿 적용</td>
	<td>
		<% DrawsecretGubun "secretGubun" , secretGubun , "class='select'" %> 사이트노출시, 시크릿 적용을 Y로 두면 타이틀만 노출되고 클릭이 되지 않습니다.
	</td>
</tr>

<% If gubun <> "5" or session("ssAdminPsn") = "7" or session("ssAdminPsn") = "11" Then %>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan=2><input type="button" id="goproc" value="메일진 수정" onclick="checkok(this.form);" class="button"></td>
	</tr>
	<% if reservationDATE <> "" then %>
		<tr bgcolor="#FFFFFF">
			<td align="center" colspan=2><br><br>
				<font color="red"><b>이미지 수정은, 메일 발송 후 1,2,3,4 이미지만 수정할때 사용해주세요.<br>이미지 수정 후 이미지주소로 꼭 변경되었는지 확인해주세요!</b></font><br>
				<input type="button" id="goimgproc" value="1,2,3,4 이미지 수정" onclick="imageedit(this.form);" class="openMask2">
			</td>
		</tr>
	<% end if %>
<% Else %>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan=2>
			<b>현재 메일예약이 되어있는 상태이므로 메일 내용은 수정하실 수 없습니다.</b><br>
			<b>꼭 수정이 필요한 경우 [상품운영팀]이슬비에게 연락바랍니다.</b>
		</td>
	</tr>
	<% if reservationDATE <> "" then %>
		<tr bgcolor="#FFFFFF">
			<td align="center" colspan=2><br><br>
				<font color="red"><b>이미지 수정은, 메일 발송 후 1,2,3,4 이미지만 수정할때 사용해주세요.<br>이미지 수정 후 이미지주소로 꼭 변경되었는지 확인해주세요!</b></font><br>
				<input type="button" id="goimgproc" value="1,2,3,4 이미지만 수정" onclick="imageedit(this.form);" class="openMask2">
			</td>
		</tr>
	<% end if %>
<% End If %>
</form>
<form name="frm_mail" method="post">
	<input type="hidden" name="idx" value="<% = idx %>">
	<input type="hidden" name="imgnumber">
	<input type="hidden" name="mode">
</form>
</table>
<div id="mask"></div> 
<div class="window">
	<table>
	<tr>
		<td>
			<font size="3" color="#FFFFFF">근무시간이 지나서 디자인이 완성된 상태입니다.<br><br>
				만약 메신저나 자리에 <strong>상품운영팀_이슬비</strong>가 있으면<br><br>
				제가 인지할 수 있도록 메세지를 주시고<br><br>
				퇴근한 상태라면 010-2991-3466로<br><br>
				반드시 통화 부탁드립니다(문자는 제외합니다)<br><br>
				제가 통화 못 받았다면 새벽까지라도 계속 콜해주세요<br><br>
				제가 인지를 못 하면 메일 발송이 안 될 수 있습니다.<br><br>
			</font>
		</td>
	</tr>
	<tr>
		<td align="center">
			<input type="button" class="button" value="네, 충분히 인지시켜드렸습니다" onclick="LastConfirm('Y')">
			<input type="button" class="button" value="아니요, 수정부분이 있습니다" onclick="LastConfirm('N')">
		</td>
	</tr>
	</table>
</div> 
<% set omail = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
