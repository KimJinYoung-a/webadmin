<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  이벤트 등록 - 화면설정
' History :  
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventPartnerWaitCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->

<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim clsEvt 
dim arrList ,intLoop
dim iTotCnt, iPageSize,iCurrPage,iTotalPage
dim evtstype, evtSD, evtED, evtusing, evtstate, evtdisp1, evtdisp2, evtNm
dim dispcate, maxDepth,evtbrand
dispcate	= requestCheckVar(Request("disp"),10) 		'전시 카테고리
maxDepth = 2 '전시카테고리 2depth까지 보여준다
dim sSort	,ebrand
iPageSize = 30
iCurrPage =  requestCheckVar(Request("iC"),10) 	
if iCurrPage ="" then iCurrPage =1
evtstype= requestCheckVar(Request("evtSType"),1)
if evtstype ="" then evtstype =1
evtNm = requestCheckVar(Request("evtNm"),64)
 evtSD = requestCheckVar(Request("evtSD"),10)
 evtED = requestCheckVar(Request("evtED"),10)
 evtusing =requestCheckVar(Request("evtUsing"),1)
 if evtusing ="" then evtusing = "Y"
 evtstate=requestCheckVar(Request("evtState"),50)
  if evtstate ="" then evtstate ="5"
  	ebrand = requestCheckVar(Request("ebrand"),32)
  

set clsEvt = new CEvent
clsEvt.FRectmakerid =  ebrand
clsEvt.FRectSType 	= evtstype
clsEvt.FRectSDate 	= evtSD
clsEvt.FRectEDate 	= evtED
clsEvt.FRectUsing 	=evtusing
clsEvt.FRectState 	=evtstate
clsEvt.FRectDispcate 	= dispcate 
clsEvt.FRectNm       =   evtNm
clsEvt.FPSize     = iPageSize
clsEvt.FCPage    = iCurrPage
  
arrList = clsEvt.fnGetEventList
iTotCnt = clsEvt.FTotcnt

set clsEvt = nothing
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수	
	dim arrstate , i, strstate(4)
	
	for i =1 to 4
	 strstate(i) = ""
	next
	 if evtstate <> "A" then
	  arrstate = split(evtstate,",")
	  for i=0 to ubound(arrstate)
	   if arrstate(i) = 0 then
	  		strstate(0) = "checked"
		elseif arrstate(i) =5 then
			strstate(1) ="checked"
		elseif arrstate(i) =7 then
			strstate(2) ="checked"
		elseif arrstate(i) =3 then
			strstate(3) ="checked"	
		end if		
	 next
	end if
%>
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.swiper-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/tag-it.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
	function jsevtMod(eC,mps){
		location.href = "/admin/eventmanage/wait/modEvent.asp?ec="+eC+"&menupos="+mps;
	}
	function jsevtDetail(eC,mps){
		location.href = "/admin/eventmanage/wait/contEvent.asp?ec="+eC+"&menupos="+mps; 
	}
	function jsSearch(){
		document.frmSearch.submit();
	}
	function jsStateAll(iVal){
		var i;
			if (document.frmSearch.evtState[0].checked){
					if (iVal==0){
						for(i=1;i<document.frmSearch.evtState.length;i++){
							document.frmSearch.evtState[i].checked = false;
						}
						}else{
							document.frmSearch.evtState[0].checked =false;
					}
			}
	}
	
		function jsPreviewEvt(eC,makerid){
		var pvW = window.open("about:blank");
		pvW.location.href='http://scm.10x10.co.kr/partner/event/plan/pvEventmain.asp?eC='+eC+'&mid='+makerid;
	}
	
	function jsPreviewMEvt(eC,makerid){
		var pvM = window.open("http://scm.10x10.co.kr/partner/event/plan/pvMEventmain.asp?eC="+eC+'&mid='+makerid,"wM","width=400, height=600,scrollbars=yes,resizable=yes");		
	}
</script>
<div class="content scrl" style="top:25px;">
	<form name="frmSearch" method="get"  action="index.asp" >
	<input type="hidden" name="menupos" value="<%=menupos%>"> 
	<input type="hidden" name="isResearch" value="1">   
 	<!-- ========== search ============================== -->
	<div class="searchWrap"  >
		<div class="search">
			<ul>
				<li>
					<label class="formTit">기간 :</label>
					<select class="formSlt" id="evtSType" name="evtSType" title="옵션 선택">
						<option value="1" <%if evtstype="1" then%>selected<%end if%>>시작일</option>
						<option value="2" <%if evtstype="2" then%>selected<%end if%>>종료일</option>
						<option value="3" <%if evtstype="3" then%>selected<%end if%>>작성일</option>
					</select>
					<input type="text" class="formTxt" id="evtSD" name="evtSD" style="width:100px" placeholder="시작일"  value="<%=evtSD%>"/>
					<input type="image" name="evtSD_trigger" id="evtSD_trigger" src="/images/admin_calendar.png" alt="달력으로 검색"  onclick="return false;" />
					~
					<input type="text" class="formTxt" id="evtED"  name="evtED" style="width:100px" placeholder="종료일" value="<%=evtED%>"/>
					<input type="image" name="evtED_trigger"  id="evtED_trigger" src="/images/admin_calendar.png" alt="달력으로 검색"   onclick="return false;"/>
					<script type="text/javascript"> 
						var CAL_Start = new Calendar({
							inputField : "evtSD", trigger    : "evtSD_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
						var CAL_End = new Calendar({
							inputField : "evtED", trigger    : "evtED_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_Start.args.max = date;
								CAL_Start.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<!--<li>
					<p class="formTit">사용여부 :</p>
					<select class="formSlt" id="evtUsing" name="evtUsing" title="옵션 선택">
						<option value="A" <%if evtusing ="A" then%>selected<%end if%>>전체</option>
						<option value="Y" <%if evtusing ="Y" then%>selected<%end if%>>사용</option>
						<option value="N" <%if   evtusing ="N" then%>selected<%end if%>>사용안함</option>
					</select>
				</li>-->
				<li>
					<label class="formTit" for="schWord">상태 :</label>
					<span class="rMar10">
						<input type="checkbox" id="evtState" name="evtState" class="formCheck" value="A" <%if evtState ="A" then%>checked<%end if%> onClick="jsStateAll(0);" />
						<label for="evtType1">전체</label>
					</span> 
					 
					<span>
						<input type="checkbox"  id="evtState" name="evtState" class="formCheck" value="5" <%=strstate(1)%> onClick="jsStateAll(1);"/>
						<label for="evtType3">승인요청</label>
					</span>
					<span>
						<input type="checkbox"  id="evtState" name="evtState"  class="formCheck" value="7" <%=strstate(2)%> onClick="jsStateAll(2);"/>
						<label for="evtType3">승인</label>
					</span> 
					<span>
						<input type="checkbox"  id="evtState" name="evtState" class="formCheck" value="3" <%=strstate(3)%> onClick="jsStateAll(3);"/>
						<label for="evtType3">반려</label>
					</span> 
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<label class="formTit" for="ctgy1">카테고리 :</label>
						<!-- #include virtual="/common/module/dispCateSelectBoxDepth_upche.asp"-->
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn> 
		<div class="search">
			<ul>
				<li>
					<label class="formTit" for="ctgy1">브랜드 :</label>
						<% drawSelectBoxDesignerwithName "ebrand", ebrand %>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>		
		<div class="search">
			<ul>
				<li>
					<label class="formTit" for="schWord">기획전명 :</label> 
					<input type="text" class="formTxt" id="evtNm" name="evtNm" style="width:400px" placeholder="기획전명을 입력해주세요." value="<%=evtNm%>"/>
				</li>
			</ul>
		</div>
		<input type="button" class="schBtn" value="검색"  onClick="jsSearch();"/>
	</div>
	</form>
	<!-- ========== //search ============================== -->
	<div class="cont">
		<div class="pad20">
			<div class="overHidden">
				 
				<div class="ftRt tPad10">
					<span>검색결과 : <strong><%=formatnumber(itotcnt,0)%></strong></span> <span class="lMar10">페이지 : <strong><%=iCurrPage%> / <%=formatnumber(iTotalPage,0)%></strong></span>
				</div>
			</div>
			<table class="tbType1 listTb tMar10">
				<thead>
				<tr>
					<th><div>No.</div></th>
					<th><div>기획전 코드</div></th>
					<th><div>브랜드ID</div></th> 
					<th><div>기획전명</div></th>
					<th><div>할인정보<br/>(상품, 쿠폰)</div></th>
					<th><div>카테고리</div></th>
					<th><div>사용여부</div></th>
					<th><div>테마</div></th>
					<th><div>상태</div></th>
					<th><div>시작일</div></th>
					<th><div>종료일</div></th> 
					<th><div>작성일</div></th>
					<th><div>미리보기</div></th>
					<th><div>관리</div></th>
				</tr>
				</thead>
				<tbody>
				<%
			if isArray(arrList) then 
					For intLoop = 0 To UBound(arrList,2)
					'evt_code,evt_name,evt_startdate,evt_enddate,evt_state,evt_regdate,evt_using,adminid,evt_dispcate,brand,salePer,saleCPer,mdtheme ,dc1nm, dc2nm
				%>
				<tr >
					<td><a href="javascript:jsevtDetail('<%=arrList(0,intLoop)%>','<%=menupos%>');"><%=itotcnt-(intLoop+((iCurrPage-1)*iPageSize))%></a></td>
					<td><a href="javascript:jsevtDetail('<%=arrList(0,intLoop)%>','<%=menupos%>');"><%=arrList(0,intLoop)%></a></td>
					<td><a href="javascript:jsevtDetail('<%=arrList(0,intLoop)%>','<%=menupos%>');"><%=arrList(9,intLoop)%></a></td>
					<td class="lt"><a href="javascript:jsevtDetail('<%=arrList(0,intLoop)%>','<%=menupos%>');"><%=arrList(1,intLoop)%></a></td>
					<td><%if arrList(10,intLoop) >"0" or arrList(10,intLoop)<> "" then%><span class="cRd1"><%=arrList(10,intLoop)%></span><%end if%>
						 <%if (arrList(10,intLoop) > "0" or arrList(10,intLoop)<> "") and (arrList(11,intLoop) >"0" or arrList(11,intLoop)<>"") then %>, <%end if%>
						<%if arrList(11,intLoop) >"0" or arrList(11,intLoop) <> "" then%><span class="cGn1"><%=arrList(11,intLoop)%></span><%end if%></td>
					<td nowrap><%if len(arrList(8,intLoop)) >3 then %><%=arrList(13,intLoop)%> > <%=arrList(14,intLoop)%><%else%><%=arrList(13,intLoop)%> <%end if%></td>
					<td><%if arrList(6,intLoop) ="Y" then%>사용<%else%>사용 안함<%end if%></td>
					<td nowrap><%=fnSetThemeNm(arrList(12,intLoop))%></td>
					<td nowrap><%=fnSetStatusNm(arrList(4,intLoop))%></span></td>
					<td><%=formatdate(arrList(2,intLoop),"0000.00.00")%></td>
					<td><%=formatdate(arrList(3,intLoop),"0000.00.00")%></td> 
					<td><%=formatdate(arrList(5,intLoop),"0000.00.00")%></td>
					<td nowrap><button type="button" class="btnIntb" onclick="jsPreviewEvt('<%=arrList(0,intLoop)%>','<%=arrList(9,intLoop)%>')">PC</button> <button type="button" class="btnIntb"  onclick="jsPreviewMEvt('<%=arrList(0,intLoop)%>','<%=arrList(9,intLoop)%>')">Mob</button></td>
					<td nowrap><button type="button" class="btnIntb" onClick="jsevtMod('<%=arrList(0,intLoop)%>','<%=menupos%>');">수정</button></td>
				</tr>
				<% NEXT
			else
				%>
				<tr><td colspan="14"> 등록된 내용이 없습니다.</td></tr>
				<%
			end if
				%>
				 
				</tbody>
			</table> 
		 </form>
		<!-- 페이징처리 --> 
		<div class="ct tPad20 cBk1">
		<table width="100%" cellpadding="10" >
			<tr>
				<td align="center">  
					<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
				</td>
			</tr>
		</table>
		</div>
	</div>
</div>
</div>
</div>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
