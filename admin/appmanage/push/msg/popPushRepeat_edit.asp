<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 푸시 반복 관리
' Hieditor : 2019.05.29 한용민 생성
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
Dim repeatidx , listimg , state , stitle, targetKey, admcomment, targetState, targetStateName, mayTargetCnt
Dim viewno , textimg , worktext, pushimg, oPush , subtitle , mode , isusing, makeridarr, itemidarr, keywordarr, i
dim bonuscouponidxarr, notclickyn, noduppDate, regadminid, lastadminid, regdate, lastupdate, pushcontents, oschedule
dim divcountrepeatgubundisp, yyyymmddttdisp, yyyydisp, iIsTargetActionValid, imgtype, targetName
dim pushimg2, pushimg3, pushimg4, pushimg5, sendranking, privateYN
	repeatidx = requestcheckvar(request("repeatidx"),10)

	If repeatidx = "0" Or repeatidx = "" Then 
		mode = "repeatInsert"
	Else
		mode = "repeatmEdit"
	End If 
	
	noduppDate  =0
	imgtype  =0
	iIsTargetActionValid = false

If repeatidx <> "0" then
	set oPush = new cpush_msg_list
		oPush.FRectidx = repeatidx
	
		if repeatidx <> "" Then
			oPush.fpush_RepeatOne_Getrow()

			if oPush.FResultCount > 0 then			
				stitle			= oPush.FOneItem.fpushtitle
				subtitle		= oPush.FOneItem.fpushurl
				state			= oPush.FOneItem.fstate
				isusing			= oPush.FOneItem.fisusing
				pushimg			= oPush.FOneItem.fpushimg
				pushimg2		= oPush.FOneItem.fpushimg2
				pushimg3		= oPush.FOneItem.fpushimg3
				pushimg4		= oPush.FOneItem.fpushimg4
				pushimg5		= oPush.FOneItem.fpushimg5
				imgtype		= oPush.FOneItem.fimgtype
				noduppDate     = oPush.FOneItem.fnoduppDate
				targetKey      = oPush.FOneItem.ftargetKey
				admcomment      = oPush.FOneItem.fadmcomment
				targetState     = oPush.FOneItem.ftargetState
				targetStateName = oPush.FOneItem.getTargetStateName
				mayTargetCnt    = oPush.FOneItem.fmayTargetCnt
				iIsTargetActionValid = oPush.FOneItem.IsTargetActionValid
				privateYN    = oPush.FOneItem.fprivateYN

				if trim(oPush.FOneItem.fmakeridarr) <> "" then
					makeridarr = replace(oPush.FOneItem.fmakeridarr,"""","")
				end if
				if trim(oPush.FOneItem.fkeywordarr) <> "" then
					keywordarr = replace(oPush.FOneItem.fkeywordarr,"""","")
				end if

				itemidarr = oPush.FOneItem.fitemidarr
				bonuscouponidxarr = oPush.FOneItem.fbonuscouponidxarr
				notclickyn = oPush.FOneItem.fnotclickyn
				regadminid = oPush.FOneItem.fregadminid
				lastadminid = oPush.FOneItem.flastadminid
				regdate = oPush.FOneItem.fregdate
				lastupdate = oPush.FOneItem.flastupdate
				if oPush.FOneItem.fpushcontents<>"" then
					pushcontents = replace(oPush.FOneItem.fpushcontents,"\n",vbcrlf)
				end if
				targetName = oPush.FOneItem.ftargetName
				sendranking = oPush.FOneItem.fsendranking
			end if
		end if
	set oPush = Nothing
else
    stitle="제목입력하세요"
	'pushcontents="(광고) 내용입력하세요"&vbcrlf&"※ 수신거부 : 마이텐바이텐 > 설정"
	pushcontents="내용입력하세요"
End If 
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">

	function jsgolist(){
		opener.location.reload();
		self.close();
	}

	//이미지 확대화면 새창으로 보여주기
	function jsImgView(sImgUrl){
		var wImgView;
		wImgView = window.open('/admin/sitemaster/play/lib/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
		wImgView.focus();
	}

	function jsDelImg(sName, sSpan){
		if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
		   eval("document.all."+sName).value = "";
		   eval("document.all."+sSpan).style.display = "none";
		}
	}

	function chgstate(v){
		var frm = document.frmstate;
		if ( v == "I" ){
			frm.state.value = 1;
		}else{
			frm.state.value = 0;
		}

		frm.target = "FrameCKP";
        frm.submit();
	}

	function chgusing(){
		var frm = document.frmdel

		frm.target = "FrameCKP";
		frm.submit();
	}

	//타켓대상
	function setComp(comp){
	    if (comp.name=="targetKey"){
			callreplacetagcodeajax(comp.value);
    	}
	}

	function jsSetImg(sImg, sName, sSpan){	
		document.domain = '10x10.co.kr';

		var winImg;
		winImg = window.open('/admin/mobile/lib/pop_uploadimg.asp?sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
		winImg.focus();
	}

	//저장
	function subcheck(){
		var frm=document.inputfrm;

		var lenRow = tabledate.rows.length;
		var daterepeatgubun = document.getElementsByName("daterepeatgubun");
		var countrepeatgubun = document.getElementsByName("countrepeatgubun");
		var yyyy = document.getElementsByName("yyyy");
		var mm = document.getElementsByName("mm");
		var dd = document.getElementsByName("dd");
		var time1 = document.getElementsByName("time1");
		var time2 = document.getElementsByName("time2");

		if(lenRow>0)	{
			i=0
			for(l=0;l<daterepeatgubun.length;l++)	{
				if( daterepeatgubun[l].value=="" ) {
					alert("수행구분을 선택해 주세요.");
					daterepeatgubun[l].focus();
					return;
				}
				// 수행구분 : 상시
				if( daterepeatgubun[l].value!="4" ) {
					if( countrepeatgubun[l].value=="" ) {
						alert("발송빈도를 선택해 주세요.");
						countrepeatgubun[l].focus();
						return;
					}
				}
				/*
				if( yyyy[l].value=="" ) {
					alert("수행일(년)을 선택해 주세요.");
					yyyy[l].focus();
					return;
				}
				if( mm[l].value=="" ) {
					alert("수행일(월)을 선택해 주세요.");
					mm[l].focus();
					return;
				}
				if( dd[l].value=="" ) {
					alert("수행일(일)을 선택해 주세요.");
					dd[l].focus();
					return;
				}
				if( time1[l].value=="" ) {
					alert("수행일(시)을 선택해 주세요.");
					time1[l].focus();
					return;
				}
				if( time2[l].value=="" ) {
					alert("수행일(분)을 선택해 주세요.");
					time2[l].focus();
					return;
				}
				*/
			}
		} else {
			//alert('발송일정을 추가해 주세요.')
			//return;
		}

        if ( frm.targetKey.value.length<1 ){
            alert('타겟을 선택해주세요');
			frm.targetKey.focus();
			return;
        }

		if (frm.stitle.value==''){ 
			alert('제목을 등록해주세요.');
			frm.stitle.focus();
			return;
		}

		if (frm.pushcontents.value==''){ 
			alert('내용을 등록해주세요.');
			frm.pushcontents.focus();
			return;
		}

		if (frm.subtitle.value==''){ 
			alert('링크을 등록해주세요');
			frm.subtitle.focus();
			return;
		}

		if(frm.state.value==''){ 
			alert("상태를 선택해주세요");
			frm.state.focus();
			return;
		}

		if(frm.sendranking.value==''){ 
			alert("발송우선순위를 선택해주세요");
			frm.sendranking.focus();
			return;
		}

		//frm.target="_blank";
		frm.submit();
	}

	// 선택삭제
	function delSelectddate(){
		
		if(confirm("선택한 일정을 삭제하시겠습니까?"))
			tabledate.deleteRow(tabledate.clickedRowIndex);
	}

	function chdaterepeatgubun(selectvalue, lenRow){
		// 수행구분 : 일별
		if (selectvalue=='1'){
			document.getElementById("divcountrepeatgubun_"+lenRow).style.display="";
			document.getElementById("yyyymmddtt_"+lenRow).style.display="";
			document.getElementById("yyyy_"+lenRow).style.display="none";
			document.getElementById("mm_"+lenRow).style.display="none";
			document.getElementById("dd_"+lenRow).style.display="none";

		// 수행구분 : 월별
		} else if (selectvalue=='2'){
			document.getElementById("divcountrepeatgubun_"+lenRow).style.display="";
			document.getElementById("yyyymmddtt_"+lenRow).style.display="";
			document.getElementById("yyyy_"+lenRow).style.display="none";
			document.getElementById("mm_"+lenRow).style.display="none";
			document.getElementById("dd_"+lenRow).style.display="";

		// 수행구분 : 년별
		} else if (selectvalue=='3'){
			document.getElementById("divcountrepeatgubun_"+lenRow).style.display="";
			document.getElementById("yyyymmddtt_"+lenRow).style.display="";
			document.getElementById("yyyy_"+lenRow).style.display="none";
			document.getElementById("mm_"+lenRow).style.display="";
			document.getElementById("dd_"+lenRow).style.display="";

		// 수행구분 : 상시
		} else if (selectvalue=='4'){
			document.getElementById("divcountrepeatgubun_"+lenRow).style.display="none";
			document.getElementById("yyyymmddtt_"+lenRow).style.display="none";
		}
	}

	function popdateSelect(){
		var tmpstr;
		var lenRow = tabledate.rows.length;

		// 행추가
		var oRow = tabledate.insertRow(lenRow);
		oRow.onmouseover=function(){tabledate.clickedRowIndex=this.rowIndex};

		// 셀추가
		var oCell1 = oRow.insertCell(0);

		tmpstr = "수행구분 : <select name='daterepeatgubun' id='daterepeatgubun_"+ parseInt(lenRow+1) +"' onchange='chdaterepeatgubun(this.value,"+ parseInt(lenRow+1) +");'><option value=''>선택</option><option value='1'>일별</option><option value='2'>월별</option><option value='3'>년별</option><option value='4'>상시</option></select>"
		tmpstr = tmpstr + " <div id='divcountrepeatgubun_"+ parseInt(lenRow+1) +"'>발행빈도 : <select name='countrepeatgubun' id='countrepeatgubun_"+ parseInt(lenRow+1) +"'><option value=''>선택</option><option value='1'>한번수행</option></select></div>";
		tmpstr = tmpstr + " <div id='yyyymmddtt_"+ parseInt(lenRow+1) +"'>수행일 : " + "<% DrawOneDateBoxdynamic "yyyy", "", "mm", "", "dd", "", "", "yyyy_""+ parseInt(lenRow+1) +""", "mm_""+ parseInt(lenRow+1) +""", "dd_""+ parseInt(lenRow+1) +""" %>" + " <% DrawTimeBoxdynamic "time1", "", "time2", "", "", "time1_""+ parseInt(lenRow+1) +""", "time2_""+ parseInt(lenRow+1) +""", "Y" %>"
		tmpstr = tmpstr + " <img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectddate()' align=absmiddle></div>"
		oCell1.innerHTML = tmpstr;
	}

	// 타켓 치환코드 가져오기. 아작스
	function callreplacetagcodeajax(targetkey){
		$("#replacetagcode").empty().html("");
		str = $.ajax({
			type: "POST",
			url: "/admin/appmanage/push/msg/pushtargetqueryrepeat_act.asp",
			data: "targetkey="+targetkey+"&mode=replacetagcode",
			dataType: "html",
			async: false
		}).responseText;
		if(str!="") {
			$("#replacetagcode").empty().html(str);
		}
	}

</script>

<form name="inputfrm" id="inputfrm" method="post" action="/admin/appmanage/push/msg/doPushRepeat_proc.asp" style="margin:0px;">
<input type="hidden" name="repeatidx" value="<%= repeatidx %>">
<input type="hidden" name="mode" value="<%=mode%>">
<table width="100%" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td colspan="4" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle"/>
		<font color="red"><b>반복푸시 등록/수정</b></font><br/><br/>
	</td>
</tr>

<% If repeatidx <> "0" Then %>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">번호</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<b><%=repeatidx%></b>
		</td>
	</tr>
<% End If %>

<tr>
	<td width="150" align="center" bgcolor="<%= adminColor("tabletop") %>">
		발송일정
		<Br><input type="button" class='button' value="일정추가" onClick="popdateSelect()">
	</td>
	<td colspan="3" bgcolor="#FFFFFF">
	    <table name='tabledate' id='tabledate' width="100%" cellpadding="5" cellspacing="1" class="a" bgcolor="C1C1C1">
		<%
		if mode = "repeatmEdit" then
			set oschedule = new cpush_msg_list
				oschedule.FRectrepeatidx = repeatidx
				oschedule.fpush_Repeat_Schedule_List()
			if oschedule.FResultCount > 0 then
		%>
			<% for i = 0 to oschedule.FResultCount - 1 %>
			<tr>
				<td bgcolor='#FFFFFF'>
					수행구분 : 
					<select name='daterepeatgubun' id='daterepeatgubun_<%= i+1 %>' onchange="chdaterepeatgubun(this.value,'<%= i+1 %>');">
					<option value='' <% if oschedule.FItemList(i).Fdaterepeatgubun="" then response.write " selected" %> >선택</option>
					<option value='1' <% if oschedule.FItemList(i).Fdaterepeatgubun="1" then response.write " selected" %> >일별</option>
					<option value='2' <% if oschedule.FItemList(i).Fdaterepeatgubun="2" then response.write " selected" %> >월별</option>
					<option value='3' <% if oschedule.FItemList(i).Fdaterepeatgubun="3" then response.write " selected" %> >년별</option>
					<option value='4' <% if oschedule.FItemList(i).Fdaterepeatgubun="4" then response.write " selected" %> >상시</option>
					</select>
					<div id='divcountrepeatgubun_<%= i+1 %>'>
					발행빈도 : 
					<select name='countrepeatgubun' id='countrepeatgubun_<%= i+1 %>'>
					<option value='' <% if oschedule.FItemList(i).fcountrepeatgubun="" then response.write " selected" %> >선택</option>
					<option value='1' <% if oschedule.FItemList(i).fcountrepeatgubun="1" then response.write " selected" %> >한번수행</option>
					</select>
					</div>
					<div id='yyyymmddtt_<%= i+1 %>'>
					수행일 : <% DrawOneDateBoxdynamic "yyyy", year(oschedule.FItemList(i).frepeatdate), "mm", Format00(2,month(oschedule.FItemList(i).frepeatdate)), "dd", Format00(2,day(oschedule.FItemList(i).frepeatdate)), "", "yyyy_" & i+1, "mm_" & i+1, "dd_" & i+1 %>
					<% DrawTimeBoxdynamic "time1", Format00(2,Hour(oschedule.FItemList(i).frepeatdate)), "time2", Format00(2,Minute(oschedule.FItemList(i).frepeatdate)), "", "time1_" & i+1, "time2_" & i+1, "Y" %>
					<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectddate()' align=absmiddle>
					</div>
				</td>
			</tr>
			<script type="text/javascript">
				chdaterepeatgubun('<%= oschedule.FItemList(i).Fdaterepeatgubun %>','<%= i+1 %>');
			</script>
			<% next %>
			<% end if %>
		<% end if %>
		</table>
		<table class=a>
		<tr>
			<td>
				<font color="red">어드민에서 일정추가/수정시 즉시 적용 안됨.</font>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">발송제외</td>
	<td colspan="3" bgcolor="#FFFFFF">
   		금일 이전 광고 발송 대상자에 대해 발송에서 제외
		&nbsp;&nbsp;
		<input type="radio" name="noduppDate" value="0" <%=CHKIIF(noduppDate=0,"checked","")%> >선택안함
		&nbsp;<input type="radio" name="noduppDate" value="1" <%=CHKIIF(noduppDate=1,"checked","")%> >(광고 1일 1회)
		&nbsp;<input type="radio" name="noduppDate" value="2" <%=CHKIIF(noduppDate=2,"checked","")%> >(광고 1일 2회)
		&nbsp;<input type="radio" name="noduppDate" value="3" <%=CHKIIF(noduppDate=3,"checked","")%> >(광고 1일 3회)
		<br><br>
		<input type="checkbox" name="notclickyn" <%=CHKIIF(notclickyn="Y","checked","")%>> 금일 이전 광고 발송 대상자에 대해 클릭한 사람 발송에서 제외
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">타게팅여부</td>
	<td colspan="3" bgcolor="#FFFFFF">
	    타겟 대상 : 
		<%' if targetKey<>"" then %>
			<%'= targetName %>
			<!--<input type="hidden" name="targetKey" value="<%'= targetKey %>" />-->
		<% 'else %>
			<% call drawSelectBoxTarget("targetKey",targetKey," onChange='setComp(this);'", "Y", "") %>
		<% 'end if %>
		<br>
		<!--타게팅 코멘트 : -->
		<input type="hidden" name="admcomment" value="<%=admcomment%>" size="140"/>
	    <div id="itargetcmt" >
		    <span id="makeridarr" style='display:none' >
		    	<br><br>브랜드ID:<textarea name="makeridarr" cols=40 rows=3><%= makeridarr %></textarea> EX) ithinkso,7321
		    </span>
		    <span id="itemidarr" style='display:none' >
		    	<br><br>상품코드:<textarea name="itemidarr" cols=40 rows=3><%= itemidarr %></textarea> EX) 12334,432132
		    </span>
		    <span id="keywordarr" style='display:none' >
		    	<br><br>키워드:<textarea name="keywordarr" cols=40 rows=3><%= keywordarr %></textarea> EX) 우산,책상
		    </span>
		    <span id="bonuscouponidxarr" style='display:none' >
		    	<br><br>보너스쿠폰번호:<textarea name="bonuscouponidxarr" cols=40 rows=3><%= bonuscouponidxarr %></textarea> EX) 652,671
		    </span>
	    </div>
    </td>
</tr>
<tr height="60">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">제목</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="stitle" value="<%= stitle %>" size="160"/>
	</td>
</tr>
<tr height="60">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">내용</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<textarea name="pushcontents" cols=100 rows=8><%= pushcontents %></textarea>
		<!--<br><br>맨앞에 <font color="red">(광고)</font> 꼭! 넣어주세요.-->
		<!--<br>맨뒤에 <font color="red">※ 수신거부 : 마이텐바이텐 > 설정</font> 꼭! 넣어주세요.-->
		<span id="replacetagcode"></span>
	</td>
</tr>
<!--<input type="hidden" name="pushcontents" value="<%'= pushcontents %>">-->
<tr height="100">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">링크</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="subtitle" value="<%=subtitle%>" size="160"/><br/>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">상태</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% If repeatidx = "0" Or repeatidx = "" Then %>
			<input type="hidden" name="state" value="0" />작성중
		<% Else %>
			<input type="hidden" name="state" value="<%=state%>" />
			<% If state = 0 Then %>
				<strong>작성중</strong>
				&nbsp;<input type="button" value="발송예약으로 변경" onclick="<%=chkiif(isusing ="Y","chgstate('I')","alert('사용중이 아닙니다.');")%>;" class="button" />
				&nbsp;
				<% If isusing="Y" then%>
					<span style="float:right;clear:both;"><input type="button" value="사용중지" onclick="chgusing();" class="button" /></span>
				<% Else %>
					<strong>&lt;&nbsp;<font color="red">사용중이 아닙니다.</font>&nbsp;&gt;</strong>
				<% End If %>
			<% ElseIf state = 1 then %>
				<strong>발송예약</strong>
				&nbsp;<input type="button" value="작성중으로 변경" onclick="<%=chkiif(isusing ="Y","chgstate('R')","alert('사용중이 아닙니다.');")%>;" class="button" />
				&nbsp;
				<% If isusing="Y" then%>
					<span style="float:right;clear:both;"><input type="button" value="사용중지" onclick="chgusing();" class="button" /></span>
				<% Else %>
					<strong>&lt;&nbsp;<font color="red">사용중이 아닙니다.</font>&nbsp;&gt;</strong>
				<% End If %>
			<% Else %>
				<%= pushmsgstate(state) %>
			<% End If %>
		<% End If %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">동시발송우선순위</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% Drawsendranking "sendranking",sendranking,"" %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">
		이미지사용여부
		<% '(최대 1000x1000) %>
		<br>가로사이즈 : 단말기가로폭/이미지수
		<br>세로사이즈 : 560
	</td>
	<td bgcolor="#FFFFFF" width="40%">
   		<input type="radio" name="imgtype" value="0" <%=CHKIIF(imgtype=0,"checked","")%> >이미지사용안함
		&nbsp;<input type="radio" name="imgtype" value="1" <%=CHKIIF(imgtype=1,"checked","")%> >등록이미지
		&nbsp;<input type="radio" name="imgtype" value="2" <%=CHKIIF(imgtype=2,"checked","")%> >상품이미지(1000)
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">이미지1</td>
	<td bgcolor="#FFFFFF" width="40%">
   		<input type="text" name="pushimg" value="<%=pushimg%>" size=160 maxlength=200 >
		<Br>
		<input type="button" name="btnpushimg" value="이미지등록" onClick="jsSetImg('<%= pushimg %>','pushimg','simgdiv1')" class="button"/>
		<div id="simgdiv1" style="padding: 5 5 5 5">
			<% IF pushimg <> "" THEN %>			
				<img src="<%=pushimg%>" border="0" height=100 onclick="jsImgView('<%=pushimg%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('pushimg','simgdiv1');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<% END IF %>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">이미지2</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="pushimg2" value="<%=pushimg2%>" size=160 maxlength=200 >
		<Br>
		<input type="button" name="btnpushimg2" value="이미지등록" onClick="jsSetImg('<%= pushimg2 %>','pushimg2','simgdiv2')" class="button"/>
		<div id="simgdiv2" style="padding: 5 5 5 5">
			<% IF pushimg2 <> "" THEN %>			
				<img src="<%=pushimg2%>" border="0" height=100 onclick="jsImgView('<%=pushimg2%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('pushimg2','simgdiv2');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<% END IF %>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">이미지3</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="pushimg3" value="<%=pushimg3%>" size=160 maxlength=200 >
		<Br>
		<input type="button" name="btnpushimg3" value="이미지등록" onClick="jsSetImg('<%= pushimg3 %>','pushimg3','simgdiv3')" class="button"/>
		<div id="simgdiv3" style="padding: 5 5 5 5">
			<% IF pushimg3 <> "" THEN %>			
				<img src="<%=pushimg3%>" border="0" height=100 onclick="jsImgView('<%=pushimg3%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('pushimg3','simgdiv3');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<% END IF %>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">이미지4</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="pushimg4" value="<%=pushimg4%>" size=160 maxlength=200 >
		<Br>
		<input type="button" name="btnpushimg4" value="이미지등록" onClick="jsSetImg('<%= pushimg4 %>','pushimg4','simgdiv4')" class="button"/>
		<div id="simgdiv4" style="padding: 5 5 5 5">
			<% IF pushimg4 <> "" THEN %>			
				<img src="<%=pushimg4%>" border="0" height=100 onclick="jsImgView('<%=pushimg4%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('pushimg4','simgdiv4');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<% END IF %>
		</div>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="10%">이미지5</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="pushimg5" value="<%=pushimg5%>" size=160 maxlength=200 >
		<Br>
		<input type="button" name="btnpushimg5" value="이미지등록" onClick="jsSetImg('<%= pushimg5 %>','pushimg5','simgdiv5')" class="button"/>
		<div id="simgdiv5" style="padding: 5 5 5 5">
			<% IF pushimg5 <> "" THEN %>			
				<img src="<%=pushimg5%>" border="0" height=100 onclick="jsImgView('<%=pushimg5%>');" alt="누르시면 확대 됩니다"/>
				<a href="javascript:jsDelImg('pushimg5','simgdiv5');"><img src="/images/icon_delete2.gif" border="0"/></a><br/>
			<% END IF %>
		</div>
	</td>
</tr>

<% if mode = "repeatmEdit" then %>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">개인화푸시여부</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<%= privateYN %>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">최초등록</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<%=regdate%><br><%=regadminid%>
		</td>
	</tr>
	<tr>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>">마지막수정</td>
		<td colspan="3" bgcolor="#FFFFFF">
			<%=lastupdate%><br><%=lastadminid%>
		</td>
	</tr>
<% end if %>

<tr bgcolor="#FFFFFF">
	<td colspan="4" align="center">
	    <% If (state < 7) Then %>
			<input type="button" value=" 저장 " class="button" onclick="subcheck();"/> &nbsp;&nbsp;
	    <% end if %>

		<input type="button" value=" 취소(main) " class="button" onclick="jsgolist();"/>
	</td>
</tr>
</table>
</form>
<form name="frmtarget" method="post" action="/admin/appmanage/push/msg/doPushrepeat_proc.asp" style="margin:0px;">
	<input type="hidden" name="repeatidx" value="<%= repeatidx %>">
	<input type="hidden" name="state" value="">
	<input type="hidden" name="mode" value="target">
</form>
<form name="frmdel" method="get" action="/admin/appmanage/push/msg/doPushrepeat_proc.asp" style="margin:0px;">
	<input type="hidden" name="repeatidx" value="<%= repeatidx %>">
	<input type="hidden" name="state" value="">
	<input type="hidden" name="mode" value="del">
</form>
<form name="frmstate" method="get" action="/admin/appmanage/push/msg/doPushrepeat_proc.asp" style="margin:0px;">
	<input type="hidden" name="repeatidx" value="<%= repeatidx %>">
	<input type="hidden" name="state" value="">
	<input type="hidden" name="mode" value="state">
</form>

<% if (application("Svr_Info")="Dev") then %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="100%" height="500"></iframe>
<% else %>
	<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<script type="text/javascript">
	<% If repeatidx <> "0" Then %>
		callreplacetagcodeajax('<%= targetkey %>');
	<% end if %>
</script>
<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->