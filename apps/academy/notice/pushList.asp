<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 푸시 알림"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/common/appFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'####################################################
' Description : 소식 리스트
' History : 2015.01.14 서동석 추가
'####################################################
Function fnUTC2LocalTime(iutcTime)
    ''iutcTime like 2015-07-27T09:20:27Z
    Dim buf : buf = Trim(iutcTime)
    If (right(buf,1)="Z") Then
        buf = replace(LEFT(iutcTime,19),"T"," ")
        buf = dateAdd("n",9*60,buf)
        fnUTC2LocalTime=FormatDateTime(buf,2)&" "&FormatDateTime(buf,4)&":"&Right(FormatDateTime(buf,3),2)
    Else
        buf = replace(LEFT(iutcTime,19),"T"," ")
        fnUTC2LocalTime=buf
    End If
    
End Function

Function fnStuffArray(tnArr,nudArr)
    Dim retArr, tRowCnt, nRowCnt
    If Not isArray(nudArr) Then
        'nudArr = array(nudArr)
        fnStuffArray = tnArr
        Exit Function
    End If
    
    If Not isArray(tnArr) Then
        'tnArr = array(tnArr)
        fnStuffArray = nudArr
        Exit Function 
    End If

    Dim total_size
    tRowCnt = UBound(tnArr,2)
    nRowCnt = UBound(nudArr,2)
    total_size = tRowCnt + nRowCnt +1

    Dim merged
    ReDim merged(5,total_size)
    
    Dim counter : counter = 0
    Dim xindex, jindex

    For xindex = 0 To ubound(tnArr,2)''-1
        For jindex = 0 To ubound(tnArr,1)''-1
            merged(jindex,counter) = tnArr(jindex,xindex)
            ''response.write merged(jindex,counter)&"<br>"
        Next 
        merged(4,counter) = ""
        merged(5,counter) = ""
        counter=counter+1
    Next
    For xindex = 0 To ubound(nudArr,2)''-1
        For jindex = 0 To ubound(nudArr,1)''-1
            merged(jindex,counter) = nudArr(jindex,xindex)
        Next 
        counter=counter+1
    Next
    Call QuickSortADO(merged,0,ubound(merged,2),2,"DESC")
    fnStuffArray = merged
End Function

Function fnParseNudgeData(sData, byref iNDataArr)
    Dim i,j,jlen
    Dim oResult
    Dim MaxNdata : MaxNdata     = 4 '넛지 PUSH 최대 표시 갯수
    Dim MaxPreDate : MaxPreDate = 2 '넛지 PUSH 최대 표시 기간(일)
    Dim bufTime
    
    fnParseNudgeData = False
    If (sData="") Then Exit Function
    If (sData="[]") Then Exit Function
        
    Set oResult = JSON.parse(sData)
   
    If Not (oResult is Nothing) Then
        jlen = oResult.length
        If (jlen>MaxNdata) Then jlen = MaxNdata
            
        ReDim iNDataArr(5,jlen)
        j=0
        
        For i=0 To jlen-1
            bufTime = fnUTC2LocalTime(oResult.get(i).display_time)
            If (datediff("d",bufTime,now())<=MaxPreDate) Then
                iNDataArr(0,j) = -1
                iNDataArr(1,j) = oResult.get(i).title
                iNDataArr(2,j) = bufTime
                iNDataArr(3,j) = datediff("n",CDate(bufTime),now())
                iNDataArr(4,j) = replace(oResult.get(i).deep_link,"tfartist://","")
                iNDataArr(5,j) = oResult.get(i).token
                iNDataArr(5,j) = server.urlencode(iNDataArr(5,j))
                j=j+1
            End If
        Next
        redim preserve iNDataArr(5,j-1)
        fnParseNudgeData = true
    End If
    Set oResult = Nothing
End Function

'' 기본값 A
Function GetSearchPushYN(byref iappKey,deviceid)
    dim sqlStr, reFAddr
    dim ret
    sqlStr = "select top 1 isNULL(isAlarm01,'Y') as pushyn, appKey from [db_academy].[dbo].[tbl_app_regInfo] where deviceid='" + deviceid + "' order by lastupdate desc" + vbCrlf
    rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
	If Not rsACADEMYget.EOF Then
		ret     =rsACADEMYget("pushyn")
		iappKey =rsACADEMYget("appKey")
	Else
		ret="N"  ''deviceid 없는경우가 있음..
	End If
	rsACADEMYget.close
	
	if (ret<>"N") then ret="Y"  '' A/C/P/N
	    
	GetSearchPushYN = ret
	
End Function

Function GetSearchPushYNByUserID(userid, byref ideviceid, byref iAppKey)
    dim sqlStr, reFAddr
    dim ret
    sqlStr = "select top 1 isNULL(pushyn,'A') as pushyn, deviceid, appkey from  db_academy.[dbo].[tbl_APP_pushYN_Academy_log] where lastuserid='" + CStr(userid) +"'" + vbCrlf
    if (ideviceid<>"") then
        sqlStr = sqlStr & " and deviceid='"&ideviceid&"'"
    end if
    sqlStr = sqlStr & " order by idx desc"

    rsACADEMYget.CursorLocation = adUseClient
	rsACADEMYget.Open sqlStr,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
	If Not rsACADEMYget.EOF Then
		ret         = rsACADEMYget("pushyn")
		ideviceid   = rsACADEMYget("deviceid")
		iAppKey     = rsACADEMYget("appkey")
	Else
		ret="N"  ''deviceid 없는경우가 있음..
	End If
	rsACADEMYget.close
	
	if (ret<>"N") then ret="Y"  '' A/C/P/N
	    
	GetSearchPushYNByUserID = ret
	
End Function


Dim pid : pid = requestCheckvar(request("pid"),200)
Dim iNDataArr, Appkey, SortDiv, OS, PushYN
Dim sData : sData = Request("njson")
Dim nDataExists : nDataExists = False
Dim userid : userid = getLoginUserID()
Dim SortDivCode 

nDataExists = fnParseNudgeData(sData, iNDataArr)
OS = request("OS")
AppKey = getWishAppKey(OS)

If Appkey="" Then Appkey=8 ''7
SortDiv = request("SortDiv")
If SortDiv="" Then SortDiv=0

''test
If (application("Svr_Info")="Dev") And pid="" Then
    pid = "APA91bETnAM7DIpp81b1b0s6ELS9sEoe2hi7vPlNySc-_as1YYRryVCztx_UXKtYKED-U8cSxQpCRw3Q5pHYGCtzzSJocmsJxjAkC6tLq0zX8kZyGnmjAAl_YhIzjgi3ez_wQxcHydrVW4eFJYNfdlrqkeH5j7rKuw"
End If



if (pid<>"") then
    PushYN = GetSearchPushYN(AppKey,pid)
else
    '' pid 가 안넘어옴.. 일단 임시.. 왠지 pid 가 넘어와야할듯함.
    PushYN = GetSearchPushYNByUserID(userid, pid, AppKey)
end if

'response.write PushYN&"|"&AppKey&"|"&userid&"|"&pid
'response.end


Dim sqlStr, ArrRows, TnArrRows, RowCnt, i 
sqlStr = "exec db_AppNoti.dbo.sp_ACA_getAppHisRecentNotiList_Academy_Artist '"&pid&"','"&userid&"'," & Appkey &","&SortDiv
'Response.write sqlstr
'Response.end
If pid<>"" Then
    
    rsAppNotiget.CursorLocation = adUseClient
    rsAppNotiget.Open sqlStr,dbAppNotiget,adOpenForwardOnly, adLockReadOnly
    If Not rsAppNotiget.Eof Then
        TnArrRows = rsAppNotiget.getRows
    End If
    rsAppNotiget.Close
End If

ArrRows = fnStuffArray(TnArrRows,iNDataArr)

RowCnt = 0
If isArray(ArrRows) Then
    RowCnt = UBound(ArrRows,2)+1
End If

Dim multipskey,sendmsg,resultdate,diffmin,pos1,pos2
Dim notititle,notitime,notiurl,noticolor, isOrderPop
Dim isNudgePush, nudgeToken, notitype

%>
<script>
$(function() {
	//sorting control
	$(".pushSort button").click(function(){
		if($(".pushSort ul").is(":hidden")){
			$(this).parent().children('ul').show();
			$(this).addClass("active");
		}else{
			$(this).parent().children('ul').hide();
			$(this).removeClass("active");
		};
	});
});


function fnPushYNSet(){
	var onoff=$("#onoff").val();
	if(onoff=="on"){
		$(".btnPushSet button").toggleClass('settingOn');
		$("#onoff").val("off");
	}else{
		$(".btnPushSet button").toggleClass('settingOn');
		$("#onoff").val("on");
	}
	
	onoff=$("#onoff").val();
	callNativeFunction('setPushReceiveYN',{"switchstate": onoff});
}

function fnPushOnoffReturnSet(onoff){
	var onoffval;
	if(onoff=="A"){
		onoffval="off"
	}else{
		onoffval="on"
	}
	$("#onoff").val(onoff);
}


// 팝업?
function fnAPPpopupNotice(iURL,ititle,ipageType){
	fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], ititle, [], iURL, ipageType);
	//fnAPPpopupBrowser(OpenType.FROM_RIGHT, [], "상품 정보", [BtnType.CONFIRM], iURL, "itemdetail");
	return false;
}


//function appPushYNSetCallBack(ret){
//    if(ret){
//		$(this).toggleClass('settingOn');
//	}
//	else{
//		alert("알림 설정 처리중 오류가 발생했습니다.");
//	}
//}
function fnSortDiv(objvalue){
	location.href="?OS=<%=OS%>&pid=<%=pid%>&sortdiv="+objvalue;
}
</script>
<style>
.pushHead .pushSort select {height:2.5rem; background-color:transparent; color:#666; font-size:1.3rem;}
.pushHead .pushSort span.schSlt:after {display:block; position:absolute; right:1.6rem; top:50%; width:0.8rem; height:0.5rem; margin-top:-0.25rem; background:url(http://image.thefingers.co.kr/apps/2016/blt_select.png) no-repeat 100% 0; background-size:auto 1.2rem; content:'';}
</style>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content">
			<h1 class="hidden">알림 메시지</h1>
			<div class="pushMsg">
				<div class="pushHead">
					<div class="pushSort">
						<span class="schSlt">
						<select name="SortDiv" onchange="fnSortDiv(this.value)">
							<option value="0"<% If SortDiv=0 Then Response.write " selected"%>>전체알림</option>
							<option value="1"<% If SortDiv=1 Then Response.write " selected"%>>공지사항</option>
							<% If userid="fingertest01" Or userid="fingertest02" Or userid="fingertest03" Or userid="fingertest04" Then %>
							<option value="6"<% If SortDiv=6 Then Response.write " selected"%>>주문접수</option>
							<% End If %>
							<option value="2"<% If SortDiv=2 Then Response.write " selected"%>>등록승인</option>
							<option value="3"<% If SortDiv=3 Then Response.write " selected"%>>등록보류</option>
							<option value="4"<% If SortDiv=4 Then Response.write " selected"%>>등록반려</option>
							<option value="5"<% If SortDiv=5 Then Response.write " selected"%>>가격변경</option>
							<% If userid="fingertest01" Or userid="fingertest02" Or userid="fingertest03" Or userid="fingertest04" Then %>
							<option value="8"<% If SortDiv=8 Then Response.write " selected"%>>Q&A</option>
							<option value="7"<% If SortDiv=7 Then Response.write " selected"%>>관련 CS</option>
							<% End If %>
						</select>
						</span>
					</div>
					<div class="btnPushSet">
						<label>알림</label>
						<button id="onoff" value="<%=CHKIIF(PushYN="Y","on","off")%>" type="button" onclick="fnPushYNSet()"<% If PushYN="Y" Then%> class="settingOn"<% End If %>>알림 설정</button>
					</div>
				</div>
				<% If RowCnt<1 Then %>
					<div class="noData"><p>최근 받은 소식이 없습니다.</p></div>
				<% Else %>				
				<ul class="pushList">

					<% For i=0 To RowCnt-1 %>
					<%
					multipskey  = ArrRows(0,i)
					sendmsg     = ArrRows(1,i)
					resultdate  = ArrRows(2,i)
					diffmin     = ArrRows(3,i)
					notiurl =""
					notitime =""
					noticolor =""
					notitype = ""
					if (multipskey="1") then notitype="notice"
					if (multipskey="2") then notitype="itemreg"
					if (multipskey="3") then notitype="itemrjt"     
					      
					isOrderPop = False
					isNudgePush = false ''(multipskey<0)
					nudgeToken =""

					If (isNudgePush) Then  ''nudge CASE 추가
						notititle = sendmsg
						notiurl   = ArrRows(4,i)
						nudgeToken = ArrRows(5,i)
						If (diffmin>=1440) Then
							notitime = Mid(resultdate,6,2) & "월 " & Mid(resultdate,9,2) & "일"
						ElseIf (diffmin>=60) Then
							notitime = CLng(diffmin/60) & "시간 전"
						ElseIf (diffmin>1) Then
							notitime = diffmin & "분 전"
						ElseIf (diffmin<1) Then
							notitime = "방금 전"
						End If
						noticolor = CStr("alram01")
					Else
						pos1 = InStr(sendmsg,"{""noti"":""")
						If (pos1>0) Then 
							pos1=pos1+LEN("{""noti"":""")
							pos2=InStr(MID(sendmsg,pos1,1024),"""")
							If (pos2>0) Then
								notititle = Mid(sendmsg,pos1,pos2-1)
							End If
						End If
						pos1 = InStr(sendmsg,"""type"":""")
						If (pos1>0) Then 
							pos1=pos1+LEN("""type"":""")
							pos2=InStr(MID(sendmsg,pos1,1024),"""")
							''If (pos2>0) Then
							''	notitype = Mid(sendmsg,pos1,pos2-1)
							''End If
						End If
						pos1 = InStr(sendmsg,"""url"":""")
						If (pos1>0) Then 
							pos1=pos1+LEN("""url"":""")
							pos2=InStr(MID(sendmsg,pos1,1024),"""")
							If (pos2>0) Then
								notiurl = Mid(sendmsg,pos1,pos2-1)
							End If
						End If

						If (diffmin>=1440) Then
							notitime = Mid(resultdate,6,2)&"월 "&Mid(resultdate,9,2)&"일"
						ElseIf (diffmin>=60) Then
							notitime = CLNG(diffmin/60)&"시간 전"
						ElseIf (diffmin>1) Then
							notitime = diffmin&"분 전"
						ElseIf (diffmin<1) Then
							notitime = "방금 전"
						End If

						If multipskey="0" Or multipskey="1" Or multipskey="2" Or multipskey="6" Or multipskey="7"  Or multipskey="8" Then
							noticolor = CStr("flagNoti")
						Else
							noticolor = CStr("flagAprv")
						End If
					End If
					notiurl = Replace(notiurl,"pmode=pms&","")
					if (notiurl<>"") then
					    if (multipskey="1") then
					        notiurl="javascript:fnAPPpopupNotice('"&notiurl&"','공지사항','notice');"
					    elseif (multipskey="2") then
					        notiurl="javascript:fnAPPpopupNotice('"&notiurl&"','상품보기','itemview');"
					    elseif (multipskey="3") then
					        notiurl="javascript:fnAPPpopupNotice('"&notiurl&"','작품등록','itemview');"
						elseif (multipskey="4") then
					        notiurl="javascript:fnAPPpopupNotice('"&notiurl&"','작품등록','itemview');"
						elseif (multipskey="5") then
					        notiurl="javascript:fnAPPpopupNotice('"&notiurl&"','상품정보','itemview');"
						elseif (multipskey="6") then
					        notiurl="javascript:fnAPPpopupNotice('"&notiurl&"','주문관리','orderview');"
						elseif (multipskey="7") then
					        notiurl="javascript:fnAPPpopupNotice('"&notiurl&"','관련 CS','csview');"
						elseif (multipskey="8") then
					        notiurl="javascript:fnAPPpopupNotice('"&notiurl&"','Q&A','qnaview');"
					    else
					        notiurl="javascript:fnAPPpopupNotice('"&notiurl&"','알림','alram');"
					    end if
					end if
					%>
					
					<li class="<%=noticolor%>">
						<a<% If (notiurl<>"") Then %> href="<%=notiurl%>"<% Else %><% End If %>>
							<% If multipskey="1" Then %><dfn>공지사항</dfn><% ElseIF multipskey="2" THEN %><dfn>등록승인</dfn><% ElseIF multipskey="3" THEN %><dfn>등록보류</dfn><% ElseIF multipskey="4" THEN %><dfn>등록반려</dfn><% ElseIF multipskey="5" THEN %><dfn>가격변경</dfn><% ElseIF multipskey="6" THEN %><dfn>주문접수</dfn><% ElseIF multipskey="8" THEN %><dfn>Q&A</dfn><% ElseIF multipskey="7" THEN %><dfn>관련 CS</dfn><% Else %><dfn>공지사항</dfn><% End If %>
							<p><%=notititle%></p>
							<span><%=notitime%></span>
						</a>
					</li>
					<% Next %>
				<% End If %>
				</ul>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<%
''Push message Badge Count Reset
sqlStr = "exec [db_academy].[dbo].[sp_ACA_sendPushMsgBadgeCount_Reset] '" & pid & "'"
dbACADEMYget.Execute sqlStr
%>
<script>
<!--
	fnAPPChangeBadgeCount("noticount",0)
//-->
</script>
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->