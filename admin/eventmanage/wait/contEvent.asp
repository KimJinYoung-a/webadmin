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

<!-- #include virtual="/admin/lib/adminbodyhead_html5.asp"-->
<%dim evtCode
evtCode =    requestCheckVar(Request("eC"),10)

if evtCode = "" then
		Call Alert_return ("유입경로에 문제가 생겼습니다.   ")
end if

dim makerid, brandNm 
dim ClsEvt
dim evtkind,evtmanager  ,evtname,evtstartdate,evtenddate,evtstate,evtregdate,evtusing ,evtlastupdate,adminid, evtcategory  ,evtcateMid,isgift ,brand  ,evttag
dim titlepc, titlemo,issale, iscoupon, saleper, salecper
dim etcitemimg,evt_mo_listbanner,subcopyK  ,evtsubname,mdtheme   ,themecolor,themecolormo ,textbgcolor       
dim giftisusing ,gifttext1 ,giftimg1  ,gifttext2 ,giftimg2  ,gifttext3 ,giftimg3          
dim giftimg1Nm, giftimg2Nm, giftimg3Nm
dim catenm, cateMnm
dim arrList, intLoop
dim arrimg, arrimgmo
dim arrlog,dispcate
dim iTotCnt, iPageSize,iCurrPage,iTotalPage
dim evttext,filelink
dim arrF,arrFName,sFName

iPageSize = 30
iCurrPage =  requestCheckVar(Request("iC"),10) 	
if iCurrPage ="" then iCurrPage =1
set ClsEvt = new CEvent 
ClsEvt.FevtCode = evtCode
ClsEvt.fnGetEventST4

evtkind       =clsEvt.Fevtkind      
evtmanager   = clsEvt.Fevtmanager   
evtname      = clsEvt.Fevtname      
evtstartdate  =clsEvt.Fevtstartdate 
evtenddate   = clsEvt.Fevtenddate   
evtstate      =clsEvt.Fevtstate     
evtregdate   = clsEvt.Fevtregdate   
evtusing     = clsEvt.Fevtusing     
evtlastupdate= clsEvt.Fevtlastupdate
adminid      = clsEvt.Fadminid     
dispcate =  clsEvt.Fevtdispcate   
catenm 		= clsEvt.FevtCateNm
cateMnm 		= clsEvt.FevtCateMNm
issale       = clsEvt.Fissale       
isgift      =  clsEvt.Fisgift       
iscoupon    =  clsEvt.Fiscoupon     
brand       =  clsEvt.Fbrand        
evttag      =  clsEvt.Fevttag    
brandNm = ClsEvt.FBrandNm
titlepc = ClsEvt.FTitlePC
titlemo = ClsEvt.FTitleMO 
saleper =  ClsEvt.Fsaleper
salecper =  ClsEvt.Fsalecper
etcitemimg        =ClsEvt.Fetcitemimg
evt_mo_listbanner =ClsEvt.Fevt_mo_listbanner 
subcopyK          =ClsEvt.FsubcopyK          
evtsubname        =ClsEvt.Fevtsubname        
mdtheme           =ClsEvt.Fmdtheme           
themecolor        =ClsEvt.Fthemecolor        
themecolormo      =ClsEvt.Fthemecolormo      
textbgcolor       =ClsEvt.Ftextbgcolor       
giftisusing       =ClsEvt.Fgiftisusing       
gifttext1         =ClsEvt.Fgifttext1         
giftimg1          =ClsEvt.Fgiftimg1          
gifttext2         =ClsEvt.Fgifttext2         
giftimg2          =ClsEvt.Fgiftimg2          
gifttext3         =ClsEvt.Fgifttext3         
giftimg3          =ClsEvt.Fgiftimg3          
evttext						=ClsEvt.FevtText
filelink					=ClsEvt.Ffilelink
 
 arrList = clsEvt.fnGetEventGroup
if mdtheme="3" then
 	ClsEvt.Fsdiv ="W"
 	arrimg 		= ClsEvt.fnGetEventItemImg
 	ClsEvt.Fsdiv ="M"
 	arrimgmo 		= ClsEvt.fnGetEventItemImg
elseif mdtheme ="2" then
	 ClsEvt.Fsdiv ="W"
	arrimg = ClsEvt.fnGetEventSlideImg
	 ClsEvt.Fsdiv ="M"
	 arrimgmo = ClsEvt.fnGetEventSlideImg
end if
arrlog = clsEvt.fnGetEventLog
iTotCnt = clsEvt.FTotcnt
set ClsEvt = nothing
iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수	
%>
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />


<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.swiper-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/tag-it.min.js"></script>
<script type="text/javascript">
	
	//그룹 상품등록
function jsSetItem(eGC){
	var winItem = window.open('/admin/eventmanage/wait/popDispItem.asp?eC=<%=evtCode%>&eGC='+eGC,'popItem','width=700,height=750,scrollbars=yes,resizable=yes');
 	winItem.focus();
}
//파일 다운로드
    function jsDownload(sDownURL, sRFN, sFN){
    var winFD = window.open(sDownURL+"/linkweb/partnerAdmin/procDownload.asp?sRFN="+sRFN+"&sFN="+sFN,"popFD","");
    winFD.focus();
 }
</script> 		
 
<div class="content scrl" style="top:25px;">
	<!-- content--->	 
	<div class="tab" style="margin:0 0 0 -1px;">
		<ul>
			<li class="col11 selected"><a href="#exhDetail01">기획전 정보</a></li>
			<li class="col11 "><a href="#exhDetail02">테마 정보</a></li>
			<li class="col11 "><a href="#exhDetail03">이력 정보</a></li>
		</ul>
	</div>
	<div class="cont">
		<div class="pad20">
			<div class="exhibit-detail" id="exhDetail01">
				<div class="overHidden">
					<div class="ftRt tPad10">
					<input type="button" class="btn" value="목록" onClick="location.href='/admin/eventmanage/wait/?menupos=<%=menupos%>'"/>
						<%if evtstate = 5 then %>
						<input type="button" class="btn" value="수정" onclick="location.href='/admin/eventmanage/wait/modEvent.asp?menupos=<%=menupos%>&ec=<%=evtCode%>'"/>
						<%end if%>
						<%if evtstate =5  then%>
						<input type="button" class="btn cBl1" value="승인"  onClick="jsSetState(7);"/>
						<input type="button" class="btn cRd1" value="반려"  onClick="jsSetState(3);"/>
						<%end if%>
					</div>
				</div>

				<div class="basicInfo tMar30">
					<h3 class="bltNo">1. 기본 정보</h3>
					<table class="tbType1 writeTb tMar10">
						<colgroup>
							<col width="14%" /><col width="" />
						</colgroup>
						<tbody>
						<tr>
							<th><div>기획전 명</div></th>
							<td><%=evtname%></td>
						</tr>
						<tr>
							<th><div>기획전 코드</div></th>
							<td><%=evtCode%></td>
						</tr>
						<tr>
							<th><div>상태</div></th>
							<td><%=fnSetStatusNm(evtstate)%></td>
							<!--
							<span class="tag bgYw1">승인요청</span>
							<span class="tag bgGy1">종료</span>
							<span class="tag bgRd1">반려</span>
							<span class="tag bgBl1">오픈</span>
							<span class="tag bgGn1">승인</span>
							-->
						</tr>
						<tr>
							<th><div>기간</div></th>
							<td><%=evtstartdate%> ~<%=evtenddate%></td>
						</tr>
						<tr>
							<th><div>할인정보</div></th>
							<td><%if isSale  then%><span class="cRd1">할인</span><%end if%>
							<%if isCoupon  then%><span  class="cGn1 lMar05">쿠폰</span><%end if%>
							</td> 
						</tr>
						<tr>
							<th><div>기능</div></th>
							<td><%if isGift then%>사은품(GIFT)<%end if%></td>
						</tr>
						<tr>
							<th><div>진열 카테고리</div></th>
							<td><%if len(dispcate)>3 then %><%=catenm%> > <%=cateMnm%><%else%><%=catenm%><%end if%></td>
						</tr>
						<tr>
							<th><div>검색 Tag</div></th>
							<td><%=evtTag%></td>
						</tr>
					<!--	<tr>
							<th><div>작성자</div></th>
							<td>홍길동</td>
						</tr>-->
						<tr>
							<th><div>요청사항</div></th>
							<td><%=evttext%>  
									<%if  filelink <>"" and not isNull(filelink) then 
											arrF= ""
											arrFName =""
											sFName=""
												arrF = split(filelink,"/") 
							 					arrFName = arrF(ubound(arrF))
												sFName = split(arrFName,".")(0)  
									%>
								<a href="javascript:jsDownload('<%=uploadImgUrl%>','<%=arrFName%>','<%=arrF(ubound(arrF)-1)&"/"&arrFName%>');" class="cBl1 tLine lMar10 fs11">파일 다운받기</a>
								 <%end if%>
						</tr>
						<tr>
							<th><div>작성일</div></th>
							<td><%=evtregdate%></td>
						</tr>
						</tbody>
					</table>
				</div>

				<div class="displayInfo tMar50">
					<h3 class="bltNo">2. 상품 진열 정보</h3>
					<div class="tbListWrap tMar10">
						<ul class="thDataList">
							<li>
								<p style="width:90px">순서</p>
								<p class="">그룹명</p>
								<p style="width:150px">상품 진열</p>
							</li>
						</ul>
						<ul class="tbDataList">
							<%if isArray(arrList) then
								for intLoop = 0 To uBound(arrList,2)
							%>
							<li>
								<p style="width:90px"><%=intLoop+1%></p>
								<p class="lt"><%=arrList(1,intLoop)%></p>
								<p style="width:150px"><input type="button" class="btn3 btnIntb" value="상품(<%=arrList(3,intLoop)%>)" onclick="jsSetItem('<%=arrList(0,intLoop)%>')" /></p>
							</li>
							<% next %>
						<%end if%>	
						</ul>
					</div>
				</div>

				<div class="saleInfo tMar50">
					<h3 class="bltNo">3. 기획전 할인 정보</h3>
					<table class="tbType1 writeTb tMar10">
						<colgroup>
							<col width="14%" /><col width="" />
						</colgroup>
						<tbody>
						<tr>
							<th><div>상품 할인 정보</div></th>
							<td><span class="cRd1"><%=saleper%></span></span></td>
						</tr>
						<tr>
							<th><div>쿠폰 할인 정보</div></th>
							<td><span class="cGn1"><%=salecper%></span></td>
						</tr>
						</tbody>
					</table>
				</div>

				<div class="bnrInfo tMar50">
					<h3 class="bltNo">4. 목록 배너 이미지 정보</h3>
					<div class="tbListWrap tMar10">

						<table class="tbType1 writeTb">
							<colgroup>
								<col width="18%" /><col width="" />
							</colgroup>
							<tbody>
							<tr>
								<th><div>기본 배너</div></th>
								<td><img src="<%=etcitemimg%>" alt="" style="width:105px;" /></td>
							</tr>
							<tr>
								<th><div>와이드 배너</div></th>
								<td><img src="<%=evt_mo_listbanner%>" alt="" style="width:194px" /></td>
							</tr>
							</tbody>
						</table>
					</div>
				</div>
			</div>

			<div class="exhibit-detail" id="exhDetail02" style="display:none;">
				<div class="overHidden">
					<div class="ftRt tPad10">
						<input type="button" class="btn" value="목록" onClick="location.href='/admin/eventmanage/wait/?menupos=<%=menupos%>'"/>
						<%if evtstate = 5 then %>
						<input type="button" class="btn" value="수정" onclick="location.href='/admin/eventmanage/wait/modEvent.asp?menupos=<%=menupos%>&ec=<%=evtCode%>'"/>
						<%end if%>
						<%if evtstate =5  then%>
						<input type="button" class="btn cBl1" value="승인"  onClick="jsSetState(7);"/>
						<input type="button" class="btn cRd1" value="반려"  onClick="jsSetState(3);"/>
						<%end if%>
					</div>
				</div>

				<div class="themeInfo tMar50">
					<h3 class="bltNo">1. 기획전 테마 정보</h3>
					<table class="tbType1 writeTb tMar10">
						<colgroup>
							<col width="14%" /><col width="" />
						</colgroup>
						<tbody>
						<tr>
							<th><div>상품테마 적용</div></th><!-- 텍스트 테마 or 이미지 테마 or 상품테마 -->
							<td>
								<div class="themaViewWrap type<%if mdtheme="2" then%>B <% If textbgcolor<>"1" Then %> typeBblack<% End If %><%elseif mdtheme="3" then%>C<%else%>A<%end if%>"><!-- for dev msg : 이벤트 유형에 따라 typeA(텍스트 테마), typeB(이미지 테마-전체롤링), typeC(상품테마-부분롤링) 클래스 넣어주세요. -->
									<div class="chPcWeb tMar30">
										<p><strong>[PC Web]</strong></p>
										<div class="fullTemplatev17" style="background-color:<%=fnEventColorCode(themecolor)%>;">
											<div class="fullContV17">
												<div class="txtCont">
													<div class="inner">
														<a href="" class="brandName arrow"><%=brandNm%><i></i></a>
														<p class="title"><%=fnSetTextForm(titlepc)%></p>
														<p class="subcopy"><%=fnSetTextForm(subcopyK)%></p>
														<%if issale or iscoupon then %>
														<div class="discount">
															<%if issale then%><span class="cRd0V15"><%=saleper%></span><%end if%><!-- for dev msg : 상품할인 cRd0V15, 쿠폰할인 cGr0V15 클래스 넣어주세요 / 상품할인 쿠폰할인 동시에 들어갈 경우 쿠폰할인율 앞에 + 붙여주세요 -->
														<%if iscoupon then%><span class="cGr0V15"><%if issale then%>+<%end if%><%=salecper%></span><%end if%>
														</div>
														<%end if%>
													</div>
												</div>
												<div class="slide">
													<%if isArray(arrimg) then
														for intLoop =0 To uBound(arrimg,2)
														 if mdtheme = "3" then
													%>
												<div><img src="<%=webImgUrl%>/image/basic/<%=GetImageSubFolderByItemid(arrimg(1,intLoop)) %>/<%=arrImg(0,intLoop)%>" alt=""></div>
												<% else%>
												<div><img src="<%=arrImg(0,intLoop)%>" alt=""></div>
												<% end if
														next
													end if
												%>	
												</div>
											</div>
										</div>
										<!-- 기차바 -->
										<!-- for dev msg : 각 기차별 배경컬러 등록 -->
										<div class="pdtGroupBarV17" id="groupBar01" name="groupBar01" style="background-color:<%=fnEventBarColorCode(themecolor)%>;">
											<p>그룹</p>
											<!-- 브랜드 링크는 있을수도, 없을수도 있음--><a href="" class="arrow btnBrand">브랜드 보러가기<i></i></a>
										</div>
										<!--// 기차바 -->
									</div>

									<div class="chMoApp tMar30">
										<p><strong>[Mobile]</strong></p>
										<div class="event-article">
											<section class="section-event hgroup-event" style="background-color:<%=fnEventColorCode(themecolormo)%>;">
												<h2><%=fnSetTextForm(titlemo)%></h2>
												<p class="subcopy"><%=fnSetTextForm(evtsubname)%></p>
												<%if isSale or iscoupon then %>
												<div class="discount tPad05">
													<%if isSale then%><b class="red rMar05"><span><%=saleper%></span></b><%end if%>
												<%if iscoupon then%><b class="green"><small>쿠폰</small><span><%=saleCper%></span></b><%end if%>
												</div>
												<%end if%>
												<div class="btnGroup"><a href="" class="btnV16a"><%=brandNm%></a></div>
											</section>
											<!-- for dev msg : 최대 5개 -->
											<div id="mdRolling" class="swiper">
												<div class="swiper-container">
													<div class="swiper-wrapper">
														<%if isArray(arrimgmo) then
														for intLoop =0 To uBound(arrimgmo,2)
														if mdtheme="3" then
													%>
												<div class="swiper-slide">
													<div class="thumbnail"><img src="<%=webImgUrl%>/image/basic/<%=GetImageSubFolderByItemid(arrimgmo(1,intLoop)) %>/<%=arrimgmo(0,intLoop)%>" alt=""></div>
												</div>
												<%else%>
												<div class="swiper-slide">
													<div class="thumbnail"><img src="<%=arrimgmo(0,intLoop)%>" alt=""></div>
												</div>
											<% 		end if
														next
													end if
												%>	
													</div>
													<div class="pagination-line"></div>
													<button type="button" class="btnNav btnPrev">이전</button>
													<button type="button" class="btnNav btnNext">다음</button>
												</div>
											</div>
										</div>
										<h3 class="groupBar">
											<span style="background-color:<%=fnEventBarColorCode(themecolormo)%>;"></span><b>BAR2</b>
										</h3>
									</div>
								</div>
							</td>
						</tr>
						</tbody>
					</table>
				</div>

				<div class="giftInfo tMar50">
					<h3 class="bltNo">6. GIFT 안내 정보</h3>
					<table class="tbType1 writeTb tMar10">
						<colgroup>
							<col width="14%" /><col width="" />
						</colgroup>
						<tbody>
						<tr>
							<th><div>사은품 종류</div></th>
							<td><%if giftisusing ="1" then%>
							1종 사은품 
							<%elseif giftisusing ="2" then%>
							2종 사은품
							<%elseif giftisusing ="3" then%>
							3종 사은품
							<%else %>
							사용안함
							<%end if%></td>
						</tr>
						<tr>
							<th><div>GIFT1</div></th>
							<td>
								<div class="inTbSet">
									<div><%=gifttext1%></div>
									<div style="width:105px;"><%if giftimg1 <> "" then%>
									<img src="<%=giftimg1%>" alt="" style="width:105px;" />
									<%end if%></div>
								</div>
							</td>
						</tr>
						<tr>
						<th><div>GIFT2</div></th>
						<td>
							<div class="inTbSet">
								<div><%=gifttext2%></div>
								<div style="width:105px;">
									<%if giftimg2 <> "" then%>
									<img src="<%=giftimg2%>" alt="" style="width:105px;" />
									<%end if%>
									</div>
							</div>
						</td>
					</tr> 
					<tr>
						<th><div>GIFT3</div></th>
						<td>
							<div class="inTbSet">
								<div><%=gifttext3%></div>
								<div style="width:105px;">
									<%if giftimg3 <> "" then%>
									<img src="<%=giftimg3%>" alt="" style="width:105px;" />
									<%end if%>
									</div>
							</div>
						</td>
					</tr> 
						</tbody>
					</table>
				</div>
			</div>

			<div class="exhibit-detail" id="exhDetail03" style="display:none;">
				<div class="overHidden">
					<div class="ftRt tPad10">
						<input type="button" class="btn" value="목록" onClick="location.href='/admin/eventmanage/wait/?menupos=<%=menupos%>'"/>
						<%if evtstate = 5 then %>
						<input type="button" class="btn" value="수정" onclick="location.href='/admin/eventmanage/wait/modEvent.asp?menupos=<%=menupos%>&ec=<%=evtCode%>'"/>
						<%end if%>
						<%if evtstate =5  then%>
						<input type="button" class="btn cBl1" value="승인"  onClick="jsSetState(7);"/>
						<input type="button" class="btn cRd1" value="반려"  onClick="jsSetState(3);"/>
						<%end if%>
					</div>
				</div>

				<div class="historyInfo tMar30">
					<h3 class="bltNo">1. 이력 조회</h3>
					<table class="tbType1 listTb tMar10">
						<thead>
						<tr>
							<th><div>상태</div></th>
							<th><div>날짜</div></th>
							<th><div>소속</div></th>
							<th><div>작성자</div></th>
							<th><div>비고</div></th>
						</tr>
						</thead>
						<tbody>
						<%if isArray(arrlog) then
								for intLoop = 0 To ubound(arrlog,2)		
								%>
						<tr>
							<td><%=fnSetStatusNm(arrlog(2,intLoop))%></td>
							<td><%=arrlog(5,intLoop)%></td>
							<td><%if arrlog(6,intLoop) =1 then%>10X10<%else%><%=arrlog(8,intLoop)%><%end if%></td>
							<td><%=arrlog(8,intLoop)%>(<%=arrlog(7,intLoop)%>)</td> 
							<td class="lt"><%=arrlog(3,intLoop)%>
								<%if  arrlog(4,intLoop) <>"" and not isNull(arrlog(4,intLoop)) then 
									arrF= ""
									arrFName =""
									sFName=""
												arrF = split(arrlog(4,intLoop),"/") 
							 					arrFName = arrF(ubound(arrF))
												sFName = split(arrFName,".")(0)  
									%>
								<a href="javascript:jsDownload('<%=uploadImgUrl%>','<%=arrFName%>','<%=arrF(ubound(arrF)-1)&"/"&arrFName%>');" class="cBl1 tLine lMar10 fs11">파일 다운받기</a>
								<%end if%> 
							</td>
						</tr>
						<%
								next
							end if
						%>														 
						</tbody>
					</table> 
						<div class="ct tPad20 cBk1" >
							<center>
						<%sbDisplayPaging "iC",iCurrPage, iTotCnt, iPageSize, 10,menupos %>
						</center>
					</div>
				</div>
			</div>

			<div class="tMar50">
					<!-- #include virtual="/admin/eventmanage/wait/incComment.asp" -->
			</div>
		</div>
	</div>
</div>
</div>
</div>
<div id="boxes"  >
	<div id="mask"></div>
	<div id="dialog" class="window" style="width:500px;">
		<form name="frmS" method="post" action="procEvent.asp">
			<input type="hidden" name="hidM" value="">
			<input type="hidden" name="eC" value="<%=evtCode%>">
			<input type="hidden" name="menupos" value="<%=menupos%>"> 
		<div id="stA">
			<dl class="lyrType">
				<dt class=""> 기획전 승인</dt>
				<dd>
					<%if isSale or isCoupon then%>
					<p>- 기획전 할인정보</p>
					<p>
						<table class="tbType1 writeTb tMar10">
							<colgroup>
								<col width="14%" /><col width="" />
							</colgroup>
							<tbody>
							<tr>
								<th><div>상품 할인 정보</div></th>
								<td><span class="rMar20"><input type="text" class="formTxt" name="eSP" value="<%=salePer%>" placeholder="0%" style="width:50px" /> (예:~10%)</span><span><input type="button" class="button" value="최대값 가져오기" onclick="fnGetMaxSalevalue('S')" /></span></td>
							</tr>
							<tr>
								<th><div>쿠폰 할인 정보</div></th>
								<td><span class="rMar20"><input type="text" class="formTxt" name="eCP" value="<%=saleCper%>" placeholder="0%" style="width:50px" /> (예:~10%)</span><span><input type="button" class="button" value="최대값 가져오기" onclick="fnGetMaxSalevalue('C')" /></span></td>
							</tr>
							</tbody>
						</table>
					</p>
					<%end if%>
				</dd>
				<dd> 
					<p>-승인 후 기획전 진행상태</p>
					<p>
						<table class="tbType1 writeTb tMar10">
							<colgroup>
								<col width="14%" /><col width="" />
							</colgroup>
							<tbody>
							<tr>
								<th><div>상태</div></th>
								<td><%sbGetOptStatusCodeSort "eventstate",0,false,"" %></td>
							</tr>
							</tbody>
						</table>
					</p> 
				</dd>
			</dl>
			<div class="ct tPad15 tMar20">
				<input type="button" class="btn3 btnDkGy" value="취소" onclick="jsCancelState();">
				<input type="button" class="btn3 btnRd" value="승인" onclick="jsProcState('7');">
			</div>
		</div>
		<!--------------------------------------------------------->
		<div id="stR">
			<dl class="lyrType">
				<dt class="">- 기획전 반려</dt>
				<dd>
					<p>반려사유를  입력해주세요(최대 100자)</p>
					<p class="tPad10"><textarea name="etext" class="formTxtA" style="width:100%; height:60px;"></textarea></p>
				</dd>				 
			</dl>
			<div class="ct tPad15 tMar20">
				<input type="button" class="btn3 btnDkGy" value="취소" onclick="jsCancelState();">
				<input type="button" class="btn3 btnRd" value="반려" onclick="jsProcState('3');">
			</div>
		</div> 	
		</form>
	</div> 
</div>
 <script type="text/javascript" src="/js/jquery.slides.min2.js"></script>

<script>
	
		
	function jsSetDisp(sType){ 
			if (sType=="B"){
				//var textW = $(".typeB .fullTemplatev17 .title").outerWidth();
				var textH = $(".typeB .fullTemplatev17 .inner").outerHeight()/2;
				var pgnW = $(".fullTemplatev17 .slide .slidesjs-pagination").outerWidth()/2;
				//$(".fullTemplatev17.typeB .inner").css("width",textW +160);
				$(".typeB .fullTemplatev17 .inner").css("margin-top",-textH);
				$(".typeB .fullTemplatev17 .slide .slidesjs-pagination").css("margin-left",-pgnW);
				$(".typeB .fullTemplatev17 .slidesjs-previous").css("margin-left",-pgnW);
				$(".typeB .fullTemplatev17 .slidesjs-next").css("margin-left",pgnW - 20);
		}else if (sType=="A"){
			var textH = 0;
			$(".typeA .fullTemplatev17 .inner").css("margin-top",-textH);
		}
	}
	
	// 이벤트 상품 최대 할인율 접수
	function fnGetMaxSalevalue(saildiv) {
		var evtcd = document.frmS.eC.value;
		$.ajax({
			type: "POST",
			url: "ajaxGetEvtMaxItemSalePer.asp",
			data: "eC="+evtcd+"&saildiv="+saildiv,
			cache: false,
			success: function(message) {
				if(message) {
					if(saildiv=="S"){
						document.frmS.eSP.value=message;
					}else{
						document.frmS.eCP.value=message;
					}
					
				} else {
					alert("이벤트의 상품이 없거나 할인중인 상품이 없습니다.");
				}
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
	}
	
$(function() {
	if ($(".fullTemplatev17 .slide div").length > 1) {
		$('.fullTemplatev17 .slide').slidesjs({
			pagination:{effect:'fade'},
			navigation:{effect:'fade'},
			play:{interval:3000, effect:'fade', auto:true},
			effect:{fade: {speed:800, crossfade:true}},
			callback: {
				complete: function(number) {
					var pluginInstance = $('.fullTemplatev17 .slide').data('plugin_slidesjs');
					setTimeout(function() {
						pluginInstance.play(true);
					}, pluginInstance.options.play.interval);
				}
			}
		});
	}

	jsSetDisp('B');
	//var textW = $(".typeB .fullTemplatev17 .title").outerWidth();
	var textH = $(".typeB .fullTemplatev17 .inner").outerHeight()/2;
	var pgnW = $(".fullTemplatev17 .slide .slidesjs-pagination").outerWidth()/2;
	//$(".fullTemplatev17.typeB .inner").css("width",textW +160);
	$(".typeB .fullTemplatev17 .inner").css("margin-top",-textH);
	$(".typeB .fullTemplatev17 .slide .slidesjs-pagination").css("margin-left",-pgnW);
	$(".typeB .fullTemplatev17 .slidesjs-previous").css("margin-left",-pgnW);
	$(".typeB .fullTemplatev17 .slidesjs-next").css("margin-left",pgnW - 20);

	/* rolling for md event */
	if ($("#mdRolling .swiper-container .swiper-slide").length > 1) {
		var mdRolling = new Swiper("#mdRolling .swiper-container", {
			pagination:"#mdRolling .pagination-line",
			paginationClickable:true,
			loop:true,
			speed:800,
			nextButton:"#mdRolling .btnNext",
			prevButton:"#mdRolling .btnPrev",
			observer:true,
			observeParents:true,
			autoplay:1700
		});
	} else {
		var mdRolling = new Swiper("#mdRolling .swiper-container", {
			pagination:false,
			noSwipingClass:".noswiping",
			noSwiping:true
		});
	}

	$("#mdRolling .pagination-line").each(function(){
		var checkItem = $(this).children("span").length;
		if (checkItem == 2) {
			$(this).addClass("grid2");
		}
		if (checkItem == 3) {
			$(this).addClass("grid3");
		}
		if (checkItem == 4) {
			$(this).addClass("grid4");
		}
		if (checkItem == 5) {
			$(this).addClass("grid5");
		}
	});

	$(".tab li").click(function() {
		$(".tab li").removeClass('selected');
		$(this).addClass('selected');
		$('.exhibit-detail').hide();
		var activeTab = $(this).find("a").attr("href");
		$(activeTab).show();
//		if (activeTab=='#exhDetail02') {
			mdRolling.init();
			jsSetDisp('B');
		//}
		return;
	});
});

 

function jsSetState(sES) { 
	$('#mask').show();
	$('#boxes').show();
	if (sES==7){
		$('#stA').show();
		$('#stR').hide();
	}else{
		$('#stA').hide();
		$('#stR').show();
	}
	var maskHeight = $(document).height();
	var maskWidth = $(window).width();
	$('#mask').css({'width':maskWidth,'height':maskHeight});

	var contH = $('#dialog').outerHeight();
	var contW = $('#dialog').outerWidth();
	$('#dialog').css('margin-left', -contW/2+'px');
	$('#dialog').css('margin-top', -contH/2+'px');

	$('#mask').click(function () {
		$('#boxes').hide();
	//	$('.window').hide();
	});

	$(window).resize(function () {
		var maskHeight = $(document).height();
		var maskWidth = $(window).width();
		$('#mask').css({'width':maskWidth,'height':maskHeight});

		var contH = $('#dialog').outerHeight();
		var contW = $('#dialog').outerWidth();
		$('#dialog').css('margin-left', -contW/2+'px');
		$('#dialog').css('margin-top', -contH/2+'px');
	});
}

function jsCancelState(){
		$('#boxes').hide();
	//	$('.window').hide();
}
function jsProcState(sES){
	var strmsg ="";
	if (sES=="7"){
	<%if isSale  then%>
	strmsg = strmsg + "상품할인이 등록되어야 합니다. 확인해주세요\n\n"
	<%end if%>
	<%if isCoupon  then%>
	strmsg = strmsg + "상품쿠폰이 등록되어야 합니다. 확인해주세요\n\n"
	<%end if%>
	<%if isGift then%>
	strmsg = strmsg + "사은품이 등록되야 합니다. 사은품 종류와 기획전 등록상품이 같은 배송타입인지 확인해주세요\n\n"
	<%end if%>
	 if(confirm(strmsg+"해당 기획전을 승인완료 하시겠습니까?")){
	 	document.frmS.hidM.value = "C";
		document.frmS.submit();
	}
	}else{
	if(document.frmS.etext.value.length >100){
		alert("반려사유를 100자 이내로 작성해주세요");
		return;
	}
	 if(confirm("해당 기획전을 반려 하시겠습니까?")){
	 	document.frmS.hidM.value = "R";
		document.frmS.submit();
	}
}
}

function jsDelFile(){
	 $("#eFileNm").empty();   
	 $("#eFile").val("");
}
</script> 
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
