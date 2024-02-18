<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_functioninfo.asp
' Discription : I형(통합형) 이벤트 기능 정보 등록 창
' History : 2019.02.15 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Dim cEvtCont
Dim eCode, eFolder, eeday, estate, esday
dim ecomment, ebbs, eisblogurl, eitemps, eregdate, ekind
dim comm_isusing, comm_text, freebie_img, comm_start, comm_end
dim eval_isusing, eval_text, eval_freebie_img, eval_start, eval_end
dim board_isusing, board_text, board_freebie_img, board_start, board_end

eCode = requestCheckVar(Request("eC"),10)

ecomment = False
ebbs = False
eisblogurl = False
eitemps = False
eFolder = eCode
IF eCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'이벤트 코드
	'이벤트 내용 가져오기
	cEvtCont.fnGetEventCont
	esday = cEvtCont.FESDay
	eeday = cEvtCont.FEEDay
	estate = cEvtCont.FEState
	eregdate = cEvtCont.FERegdate
	ekind = cEvtCont.FEKind
	IF datediff("d",now,eeday) <0 THEN estate = 9 '기간 초과시 종료표기
	'이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay
	ecomment = cEvtCont.FECommnet
	ebbs = cEvtCont.FEBbs
	eitemps = cEvtCont.FEItemps
 	eisblogurl = cEvtCont.FSisGetBlogURL

	'엠디 등록 이벤트 테마 정보
	cEvtCont.fnGetEventMDThemeInfo
	comm_isusing = cEvtCont.Fcomm_isusing
	comm_text = cEvtCont.Fcomm_text
	freebie_img = cEvtCont.Ffreebie_img
	comm_start = cEvtCont.Fcomm_start
	comm_end = cEvtCont.Fcomm_end
	eval_isusing = cEvtCont.Feval_isusing
	eval_text = cEvtCont.Feval_text
	eval_freebie_img = cEvtCont.Feval_freebie_img
	eval_start = cEvtCont.Feval_start
	eval_end = cEvtCont.Feval_end
	board_isusing = cEvtCont.Fboard_isusing
	board_text = cEvtCont.Fboard_text
	board_freebie_img = cEvtCont.Fboard_freebie_img
	board_start = cEvtCont.Fboard_start
	board_end = cEvtCont.Fboard_end
	set cEvtCont = nothing
	if comm_start="" or IsNull(comm_start) then comm_start=esday
	if comm_end="" or IsNull(comm_end) then comm_end=eeday
	if eval_start="" or IsNull(eval_start) then eval_start=esday
	if eval_end="" or IsNull(eval_end) then eval_end=eeday
	if board_start="" or IsNull(board_start) then board_start=esday
	if board_end="" or IsNull(board_end) then board_end=eeday
end if 
%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
function jsEvtSubmit(frm){
    frm.action="functioninfo_process.asp";
    frm.submit();
}

function jsNowDate(){
    var mydate=new Date()
    var year=mydate.getYear()
        if (year < 1000)
            year+=1900

    var day=mydate.getDay()
    var month=mydate.getMonth()+1
        if (month<10)
            month="0"+month

    var daym=mydate.getDate()
        if (daym<10)
            daym="0"+daym

    return year+"-"+month+"-"+ daym
}

function fnChangeDivece(div){
    if(div=="C"){
		if (document.frmEvt.chComm.checked){
        	$("#commentdiv").show();
		}
		else{
        	$("#commentdiv").hide();
		}
    }
    else if(div=="E"){
        if (document.frmEvt.chItemps.checked){
        	$("#evaldiv").show();
		}
		else{
        	$("#evaldiv").hide();
		}
    }
	else if(div=="B"){
        if (document.frmEvt.chBbs.checked){
        	$("#boarddiv").show();
		}
		else{
        	$("#boarddiv").hide();
		}
    }
}

function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/event/v5/lib/pop_event_uploadimg.asp?yr=<%=Year(eregdate)%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function popCommentXLS(ecd) {
	var wCmtXls = window.open('/admin/eventmanage/event/v5/popup/pop_event_Comment_xls.asp?eC='+ecd,'pXls','width=400,height=150');
	wCmtXls.focus();
}

function popBBSXLS(ecd) {
	var wBBSXls = window.open('/admin/eventmanage/event/v5/popup/pop_event_board_xls.asp?eC='+ecd,'pXls','width=400,height=150');
	wBBSXls.focus();
}
</script>

<form name="frmEvt" method="post" style="margin:0px;">
<% if eCode<>"0" then %>
<input type="hidden" name="imod" value="BU">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<% else %>
<input type="hidden" name="imod" value="BI">
<% end if %>
<div class="popV19">
	<div class="popHeadV19">
		<h1>기획전 기능 정보</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
					<th>기능</th>
					<td>
						<% if ekind<> "5" then %>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chItemps"<% if eitemps  then %> checked<% end if %> value="1" onclick="fnChangeDivece('E');">
								상품후기
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chComm"<% if ecomment then %> checked<% end if %> value="1" onclick="fnChangeDivece('C');">
								코멘트
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chBbs"<% if ebbs  then %> checked<% end if %> value="1" onclick="fnChangeDivece('B');">
								포토코멘트
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="isblogurl"<% if eisblogurl then %> checked<% end if %> value="1">
								blog URL
								<i class="inputHelper"></i>
							</label>
						</div>
						<% else %>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="checkbox" class="formCheckInput" name="chComm"<% if ecomment then %> checked<% end if %> value="1" onclick="fnChangeDivece('C');">
								코멘트
								<i class="inputHelper"></i>
							</label>
						</div>
						<% End If %>
					</td>
				</tr>
				<tr>
					<th>신청리스트</th>
					<td>
						<img src="/images/icon_excel_reply.gif" alt="코멘트 참여자 Excel다운로드" onClick="popCommentXLS(<%=eCode%>);" style="cursor:pointer" align="absmiddle">
		   				<img src="/images/icon_excel_bbs.gif" alt="게시판 참여자 Excel다운로드" onClick="popBBSXLS(<%=eCode%>);" style="cursor:pointer" align="absmiddle">
						<img src="/images/icon_excel_vote.gif" alt="응모 참여자 Excel다운로드" onClick="window.open('/admin/eventmanage/event/v5/popup/pop_event_votelist_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle" title ="xls 다운로드 회원기반">
						<img src="/images/icon_excel_vote.gif" alt="응모 참여자 Excel다운로드 비회원"  title ="xls 다운로드 비회원" onClick="window.open('/admin/eventmanage/event/v5/popup/pop_event_votelist_guest_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle">
						<img src="/images/icon_excel_vote.gif" alt="응모 참여자 Excel다운로드 Lite버전"  title ="xls 다운로드 Lite버전" onClick="window.open('/admin/eventmanage/event/v5/popup/pop_event_votelist_lite_xls.asp?eC=<%=eCode%>','voteXls','width=400,height=150');" style="cursor:pointer" align="absmiddle">
					</td>
				</tr>
			</tbody>
		</table>
		<div class="tMar15 tPad25 topLineGrey2" id="commentdiv" style="display:<% If ecomment Then %><% Else %>none<% End If %>">
			<h3 class="fs15">코멘트 이벤트 정보</h3>
			<table class="tableV19A tMar10">
				<colgroup>
					<col style="width:150px;">
					<col style="width:auto;">
				</colgroup>
				<tbody>
					<tr>
						<th>이벤트 내용</th>
						<td>
							<input type="hidden" name="comm_isusing" value="Y">
							<textarea name="comm_text" rows="5" cols="50" placeholder="내용입력"><%=comm_text%></textarea>
						</td>
					</tr>
					<tr>
						<th>사은품 이미지</th>
						<td>
							<button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=freebie_img%>','freebie_img','spanfreebie_img');return false;">사은품 이미지 등록</button>
							<div class="formInline lMar10 ">
                            	<span class="cGy1 fs12">*이미지 사이즈 250px * 250px</span>
                        	</div>
							<input type="hidden" name="freebie_img" value="<%=freebie_img%>">
							<%IF freebie_img <> "" THEN %><button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('freebie_img','spanfreebie_img');return false;">삭제</button><%END IF%>
							<div class="previewThumb150W tMar10" id="spanfreebie_img"><%IF freebie_img <> "" THEN %><img src="<%=freebie_img%>" alt=""><%END IF%></div>
						</td>
					</tr>
					<tr>
						<th>기간</th>
						<td>
							<div class="formInline"><span class="datepicker">시작일 <span class="mdi mdi-calendar-month"></span> <input type="text" id="datepicker1" name="comm_start" placeholder="시작일" value="<%=comm_start%>" readonly="true"></span></div>
							<div class="formInline"><span class="datepicker">종료일 <span class="mdi mdi-calendar-month"></span> <input type="text" id="datepicker2" name="comm_end" placeholder="종료일" value="<%=comm_end%>" readonly="true"></span></div>
						</td>
					</tr>
				</tbody>
			</table>
		</div>
		<div class="tMar15 tPad25 topLineGrey2" id="evaldiv" style="display:<% If eitemps Then %><% Else %>none<% End If %>">
			<h3 class="fs15">상품후기 이벤트 정보</h3>
			<table class="tableV19A tMar10">
				<colgroup>
					<col style="width:150px;">
					<col style="width:auto;">
				</colgroup>
				<tbody>
					<tr>
						<th>이벤트 내용</th>
						<td>
							<input type="hidden" name="eval_isusing" value="Y">
							<textarea name="eval_text" rows="5" cols="50" placeholder="내용입력"><%=eval_text%></textarea>
						</td>
					</tr>
					<tr>
						<th>사은품 이미지</th>
						<td>
							<button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=eval_freebie_img%>','eval_freebie_img','spaneval_freebie_img');return false;">사은품 이미지 등록</button>
							<div class="formInline lMar10 ">
                            	<span class="cGy1 fs12">*이미지 사이즈 250px * 250px</span>
                        	</div>
							<input type="hidden" name="eval_freebie_img" value="<%=eval_freebie_img%>">
							<%IF eval_freebie_img <> "" THEN %><button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('eval_freebie_img','spaneval_freebie_img');return false;">삭제</button><%END IF%>
							<div class="previewThumb150W tMar10" id="spaneval_freebie_img"><%IF eval_freebie_img <> "" THEN %><img src="<%=eval_freebie_img%>" alt=""><%END IF%></div>
						</td>
					</tr>
					<tr>
						<th>기간</th>
						<td>
							<div class="formInline"><span class="datepicker">시작일 <span class="mdi mdi-calendar-month"></span> <input type="text" id="datepicker5" name="eval_start" placeholder="시작일" value="<%=eval_start%>" readonly="true"></span></div>
							<div class="formInline"><span class="datepicker">종료일 <span class="mdi mdi-calendar-month"></span> <input type="text" id="datepicker6" name="eval_end" placeholder="종료일" value="<%=eval_end%>" readonly="true"></span></div>
						</td>
					</tr>
				</tbody>
			</table>
		</div>
		<div class="tMar15 tPad25 topLineGrey2" id="boarddiv" style="display:<% If ebbs Then %><% Else %>none<% End If %>">
			<h3 class="fs15">포토게시판 이벤트 정보</h3>
			<table class="tableV19A tMar10">
				<colgroup>
					<col style="width:150px;">
					<col style="width:auto;">
				</colgroup>
				<tbody>
					<tr>
						<th>이벤트 내용</th>
						<td>
							<input type="hidden" name="board_isusing" value="Y">
							<textarea name="board_text" rows="5" cols="50" placeholder="내용입력"><%=board_text%></textarea>
						</td>
					</tr>
					<tr>
						<th>사은품 이미지</th>
						<td>
							<button class="btn4 btnBlue1" onClick="jsSetImg('<%=eFolder%>','<%=board_freebie_img%>','board_freebie_img','spanboard_freebie_img');return false;">사은품 이미지 등록</button>
							<div class="formInline lMar10 ">
                            	<span class="cGy1 fs12">*이미지 사이즈 250px * 250px</span>
                        	</div>
							<input type="hidden" name="board_freebie_img" value="<%=board_freebie_img%>">
							<%IF board_freebie_img <> "" THEN %><button class="btn4 btnGrey1 lMar05" onClick="jsDelImg('eval_freebie_img','spanboard_freebie_img');return false;">삭제</button><%END IF%>
							<div class="previewThumb150W tMar10" id="spanboard_freebie_img"><%IF board_freebie_img <> "" THEN %><img src="<%=board_freebie_img%>" alt=""><%END IF%></div>
						</td>
					</tr>
					<tr>
						<th>기간</th>
						<td>
							<div class="formInline"><span class="datepicker">시작일 <span class="mdi mdi-calendar-month"></span> <input type="text" id="datepicker7" name="board_start" placeholder="시작일" value="<%=board_start%>" readonly="true"></span></div>
							<div class="formInline"><span class="datepicker">종료일 <span class="mdi mdi-calendar-month"></span> <input type="text" id="datepicker8" name="board_end" placeholder="종료일" value="<%=board_end%>" readonly="true"></span></div>
						</td>
					</tr>
				</tbody>
			</table>
		</div>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">저장</button>
	</div>
</div>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->