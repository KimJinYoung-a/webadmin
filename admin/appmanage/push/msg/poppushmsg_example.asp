<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 예약 푸시 메시지 작성
' Hieditor : 서동석 생성
'			 2017.03.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead_utf8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/push/apppush_msg_cls.asp" -->
<%
Dim idx, pushtitle , pushurl, pushimg2, pushimg3, pushimg4, pushimg5 , pushimg , state , testpush, baseIdx, pushcontents, oPush , oPushinfo
dim repeatpushyn, targetKey , i, userid
	idx = requestCheckVar(request("idx"),10)
	userid = requestCheckVar(request("userid"),32)
	repeatpushyn = requestCheckVar(request("repeatpushyn"),1)

'//db 1row
If idx <> "" then
	set oPush = new cpush_msg_list
			oPush.FRectIdx = idx
		
		if idx <> "" Then
			' 반복푸시관리에서 타고 들어온거
			if repeatpushyn="Y" then
				oPush.fpush_RepeatOne_Getrow()
			else
				oPush.pushmsgtest_getrow()
			end if

			if oPush.FResultCount > 0 then			
				pushtitle	= oPush.FOneItem.fpushtitle
				pushurl		= oPush.FOneItem.fpushurl
				pushimg		= oPush.FOneItem.fpushimg
				pushimg2	= oPush.FOneItem.fpushimg2
				pushimg3	= oPush.FOneItem.fpushimg3
				pushimg4	= oPush.FOneItem.fpushimg4
				pushimg5	= oPush.FOneItem.fpushimg5
				state		= oPush.FOneItem.fstate
				testpush	= oPush.FOneItem.ftestpush
				baseIdx     = oPush.FOneItem.fbaseIdx
				targetKey 	= oPush.FOneItem.ftargetKey
				if oPush.FOneItem.fpushcontents<>"" then
					pushcontents     = replace(oPush.FOneItem.fpushcontents,"\n",vbcrlf)
				end if
			end if	
		end if
	set oPush = Nothing
Else
	Response.write "<script type='text/javascript'>alert('잘못된 접근입니다.');</script>"
	session.codePage = 949
	Response.write "<script type='text/javascript'>self.close();</script>"
	Response.End 
End If 

'// user_chk
set oPushinfo = new cpush_msg_list
	oPushinfo.FRectuserid = userid
	oPushinfo.FPageSize = 100
	oPushinfo.FCurrPage = 1

	If userid <> "" Then
		oPushinfo.pushmsg_userinfo()
	End If 

dim testPkeyVal : testPkeyVal=""
dim testikeyVal : testikeyVal=idx+1000000

if Not (IsNULL(baseIdx) or (baseIdx="")) then
	testPkeyVal = baseIdx+1000000
end if

''해당기능사용안함. 사용할경우 아래 주석처리
testikeyVal = ""
testPkeyVal = ""
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

	// 선택기기1개
	function iSubmit(){
		var frm = document.frmAct;

		<% If userid <> "" Then %>
			<% if oPushinfo.FresultCount > 0 then %>
		if(!$('#appkey > option:selected').val()) {
			alert("앱종류를 선택하세요");
			frm.appkey.focus();
			return;
		}

		if (frm.deviceid.value.length<10){
			alert('디바이스 ID를 입력하세요.');
			frm.deviceid.focus();
			return;
		}

		if (frm.deviceid.value.length<10){
			alert('디바이스ID를 입력하세요.');
			frm.deviceid.focus();
			return;
		}

		if (frm.message.value==''){
			alert('제목을 등록해주세요');
			frm.message.focus();
			return;
		}

		if (frm.pushcontents.value==''){
			alert('내용을 등록해주세요');
			frm.pushcontents.focus();
			return;
		}

		if (confirm('발송하시겠습니까?')){
			frm.mode.value="test_insert";
			 frm.submit();
		}
			<% else %>
			alert('등록된 디바이스 정보가 없습니다. 아이디를 확인 해주세요');
			return;
			<% end if %>
		<% end if %>
	}

	// 등록된전체기기
	function allSubmit(){
		if (frmAct.useridarr.value.length<1){
			alert('일괄발송할 아이디를 입력하세요.');
			frmAct.useridarr.focus();
			return;
		}
		if (frmAct.message.value==''){
			alert('제목을 등록해주세요');
			frmAct.message.focus();
			return;
		}

		if (frmAct.pushcontents.value==''){
			alert('내용을 등록해주세요');
			frmAct.pushcontents.focus();
			return;
		}

		if (confirm('발송하시겠습니까?')){
			frmAct.mode.value="test_allinsert";
			frmAct.submit();
		}
	}

	function testchkid(){
		var frm = document.frm;
		var testid = frm.userid.value;
		
		frm.submit();		
	}

	$(function(){
		$('#appkey').change(function () {
			var optionSelected = $(this).find("option:selected");
			var deviceid  = optionSelected.attr("value2");
			var paramurl  = optionSelected.attr("value3");
            var appkey = optionSelected.attr("value4");  //2017/07/03
            
			var frm = document.frmAct;
			frm.deviceid.value = deviceid;
			$("#appverval").text(paramurl);
            
            if (appkey=='6'){
                document.frmAct.paramvalue[0].value='<%=testPkeyVal%>';
            }else{
                document.frmAct.paramvalue[0].value='';
            }
		});
	});

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="repeatpushyn" value="<%= repeatpushyn %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>푸시메시지 테스트 메시지 발송</b></font><br/>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="150">테스트용 푸시 아이디 검색</td>
	<td bgcolor="#FFFFFF">
		<b><input type="text" name="userid" value="<%=userid%>" size="30" bgcolor="#FFFFFF">&nbsp;&nbsp;<input type="button" value="검색" onclick="testchkid();" class="button"></b>
	</td>
</tr>
</table> 
</form>

<form name="frmAct" method="post" action="/admin/appmanage/push/msg/doPushmsg_proc.asp" style="margin:0px;">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="simg1" value="">
<input type="hidden" name="mode" value="test_insert">
<input type="hidden" name="repeatpushyn" value="<%= repeatpushyn %>">

<% If userid <> "" Then %>
	<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oPushinfo.FresultCount>0 then %>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%= adminColor("tabletop") %>" align="center">앱선택</td>
			<td >
				<select name="appkey" id="appkey">
				<option value="" value2="">===선택하세요===</option>
				<% for i=0 to oPushinfo.FresultCount-1 %>
				<option value="<%=oPushinfo.FItemList(i).fappkey%>" value2="<%=oPushinfo.FItemList(i).fdeviceid%>" value3="<%=oPushinfo.FItemList(i).Fappver%>" value4="<%=oPushinfo.FItemList(i).fappkey%>"><%= Selectappname(oPushinfo.FItemList(i).fappkey)%>(<%=oPushinfo.FItemList(i).fappVer%>)</option>
				<!-- <option value="5">wishApp(ios)</option> -->
				<% Next %>
				</select><br>(최근순)
			</td>
			<td> 앱버전 : <span id="appverval" style="text-align:center;"></span></td>
			<td>필수</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%= adminColor("tabletop") %>" align="center">deviceid</td>
			<td colspan="2"><input type="text" name="deviceid" value="" size="60" bgcolor="#FFFFFF"></td>
			<td>필수,디바이스ID</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td align="center" colspan="4">
				<input type="button" value="발송(선택기기1개)" onClick="iSubmit();" class="button">
			</td>
		</tr>
	<% Else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="4" height="30" align="center">등록된 디바이스 정보가 없습니다. 아이디를 확인 해주세요</td>
		</tr>
	<% End If %>
	</table>
<% End If %>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="margin-top: 20px;">
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">일괄발송 아이디</td>
	<td colspan=3>
		<input type="text" name="useridarr" value="" size="180" maxlength="150" bgcolor="#FFFFFF">
		<br>예) tozzinet,kobula
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan="4">
		<input type="button" value="발송(등록된전체기기)" onClick="allSubmit();" class="button">
	</td>
</tr>
</table>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="margin-top: 20px;" >
<tr bgcolor="#FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width=100 >푸시번호</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<b><%=idx%>번 메시지</b>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">푸시제목</td>
	<td>
		<input type="text" name="param0" value="message" readonly>
	</td>
	<td>
		<input type="text" name="message" value="<%= pushtitle %>" size="140" />
	</td>
	<td>필수,alert메세지</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">푸시내용</td>
	<td>
		<input type="text" name="param1" value="pushcontents" readonly>
	</td>
	<td>
		<textarea name="pushcontents" cols=100 rows=8><%= pushcontents %></textarea>
	</td>
	<td>필수,alert메세지</td>
</tr>
<% if pushimg<>"" and not(isnull(pushimg)) then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">이미지1</td>
		<td>
			<input type="text" name="param2" value="array-image-url" >
		</td>
		<td><input type="text" name="array-image-url" value="<%= pushimg %>" size="140"></td>
		<td></td>
	</tr>
<% end if %>
<% if pushimg2<>"" and not(isnull(pushimg2)) then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">이미지2</td>
		<td>
			<input type="text" name="param2" value="array-image-url" >
		</td>
		<td><input type="text" name="array-image-url" value="<%= pushimg2 %>" size="140"></td>
		<td></td>
	</tr>
<% end if %>
<% if pushimg3<>"" and not(isnull(pushimg3)) then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">이미지3</td>
		<td>
			<input type="text" name="param2" value="array-image-url" >
		</td>
		<td><input type="text" name="array-image-url" value="<%= pushimg3 %>" size="140"></td>
		<td></td>
	</tr>
<% end if %>
<% if pushimg4<>"" and not(isnull(pushimg4)) then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">이미지4</td>
		<td>
			<input type="text" name="param2" value="array-image-url" >
		</td>
		<td><input type="text" name="array-image-url" value="<%= pushimg4 %>" size="140"></td>
		<td></td>
	</tr>
<% end if %>
<% if pushimg5<>"" and not(isnull(pushimg5)) then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">이미지5</td>
		<td>
			<input type="text" name="param2" value="array-image-url" >
		</td>
		<td><input type="text" name="array-image-url" value="<%= pushimg5 %>" size="140"></td>
		<td></td>
	</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">중복방지키</td>
	<td>
		<input type="text" name="params" value="pkey" > <% ''안드로이드 중복 방지용 key 상위  %>
	</td>
	<td><input type="text" name="paramvalue" value="" size="140"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">이미지썸네일</td>
	<td>
		<input type="text" name="params" value="imgurl" >
	</td>
	<td><input type="text" name="paramvalue" value="<%= pushimg %>" size="140"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">타켓키</td>
	<td>
		<input type="text" name="params" value="targetkey" >
	</td>
	<td><input type="text" name="paramvalue" value="<%= targetkey %>" size="140"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">푸시알림음</td>
	<td>
		<input type="text" name="params" value="sound" >
	</td>
	<td><input type="text" name="paramvalue" value="default" size="140"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">푸시경로</td>
	<td>
		<input type="text" name="params" value="url" >
	</td>
	<td><input type="text" name="paramvalue" value="<%=pushurl%>" size="140"></td>
	<td>필수,alert메세지</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">푸시타입</td>
	<td>
		<input type="text" name="params" value="type" >
	</td>
	<td><input type="text" name="paramvalue" value="event" size="140"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">배찌</td>
	<td>
		<input type="text" name="params" value="badge" >
	</td>
	<td><input type="text" name="paramvalue" value="1" size="140"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">클릭체크키</td>
	<td>
		<input type="text" name="params" value="ikey" >
	</td>
	<td><input type="text" name="paramvalue" value="<%=testikeyVal%>" size="140"></td>
	<td></td>
</tr>

<% if (pushimg<>"" and not(isnull(pushimg))) or (pushimg2<>"" and not(isnull(pushimg2))) or (pushimg3<>"" and not(isnull(pushimg3))) or (pushimg4<>"" and not(isnull(pushimg4))) or (pushimg5<>"" and not(isnull(pushimg5))) then %>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">다중이미지여부</td>
		<td>
			<input type="text" name="params" value="category" >
		</td>
		<td><input type="text" name="paramvalue" value="image-notification" size="140" ></td>
		<td></td>
	</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">파라메타</td>
	<td>
		<input type="text" name="params" value="" >
	</td>
	<td><input type="text" name="paramvalue" value="" size="140"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">파라메타</td>
	<td>
		<input type="text" name="params" value="" >
	</td>
	<td><input type="text" name="paramvalue" value="" size="140"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">파라메타</td>
	<td>
		<input type="text" name="params" value="" >
	</td>
	<td><input type="text" name="paramvalue" value="" size="140"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">파라메타</td>
	<td>
		<input type="text" name="params" value="" >
	</td>
	<td><input type="text" name="paramvalue" value="" size="140"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">파라메타</td>
	<td>
		<input type="text" name="params" value="" >
	</td>
	<td><input type="text" name="paramvalue" value="" size="140"></td>
	<td></td>
</tr>
</table>
</form>

<%
session.codePage = 949
set oPushinfo = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->