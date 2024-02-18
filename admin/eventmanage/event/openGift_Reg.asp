<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/openGift.asp
' Description :  전체증정이벤트 관리 369등.
' History : 2010.04 서동석 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<!-- #include virtual="/lib/classes/event/openGiftCls.asp"-->
<%
dim eCode : eCode=request("eC")
Dim cEvtCont
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate, etag
Dim echkdisp, ecategory,esale,egift,ecoupon,ecomment,ebbs,eitemps,eapply,ebimg,etemp,emimg,ehtml,eisort,eiaddtype,edid,emid,efwd,selPartner
Dim eusing, tmp_cdl, tmp_cdm, elktype, elkurl, ebimg2010, gimg, ebrand, eicon, ecommenttitle, elinkcode
dim dopendate, dclosedate, blnFull, blnIteminfo
dim ehtml5

dim eFolder : eFolder = eCode

Dim oOpenGift
Dim imod : imod="I"
Dim frontopen,  OGtitle, OopenHtml, OopenHtmlWeb, opengiftType, opengiftScope
set oOpenGift=new CopenGift
oOpenGift.FRectEventCode = eCode
oOpenGift.getOneOpenGift

if (oOpenGift.FResultCount>0) then
    frontopen = oOpenGift.FOneItem.FfrontOpen
    OGtitle   = oOpenGift.FOneItem.FopenImage1
    imod      = "E"
    eCode     = oOpenGift.FOneItem.Fevent_code
    OopenHtml = db2Html(oOpenGift.FOneItem.FopenHtml)
    OopenHtmlWeb = db2Html(oOpenGift.FOneItem.FopenHtmlWeb)
    opengiftType = oOpenGift.FOneItem.FopengiftType
    opengiftScope = oOpenGift.FOneItem.FopengiftScope
end if


IF eCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'이벤트 코드
	'이벤트 내용 가져오기
	cEvtCont.fnGetEventCont
	ekind 		=	cEvtCont.FEKind
	eman 		=	cEvtCont.FEManager
	escope 		=	cEvtCont.FEScope
	selPartner	=	cEvtCont.FEPartnerID
	ename 		=	db2html(cEvtCont.FEName)
	esday 		=	cEvtCont.FESDay
	eeday 		=	cEvtCont.FEEDay
	epday 		=	cEvtCont.FEPDay
	elevel 		=	cEvtCont.FELevel
	estate 		=	cEvtCont.FEState
	IF datediff("d",now,eeday) <0 THEN estate = 9 '기간 초과시 종료표기
	eregdate	=	cEvtCont.FERegdate
	eusing		= 	cEvtCont.FEUsing

	'이벤트 화면설정 내용 가져오기
	cEvtCont.fnGetEventDisplay
	echkdisp 		= cEvtCont.FChkDisp
	tmp_cdl 		= cEvtCont.FECategory

	tmp_cdm		= cEvtCont.FECateMid
	esale 			= cEvtCont.FESale
	egift 			= cEvtCont.FEGift
	ecoupon 		= cEvtCont.FECoupon
	ecomment 		= cEvtCont.FECommnet
	ebbs 			= cEvtCont.FEBbs
	eitemps 		= cEvtCont.FEItemps
	eapply 			= cEvtCont.FEApply
	elktype			= cEvtCont.FELinkType
	IF elktype="" Then elktype="E" '//링크타입 기본값 설정
	elkurl			= cEvtCont.FELinkURL
	ebimg 			= cEvtCont.FEBImg
	ebimg2010		= cEvtCont.FEBImg2010
	gimg			= cEvtCont.FEGImg
	etemp			= cEvtCont.FETemp
	if etemp = 5 or etemp = 6  THEN	'수작업 이벤트 일 경우 처리
		ehtml5 		= db2html(cEvtCont.FEHtml)
	else
		emimg 		= cEvtCont.FEMImg
		ehtml 		= db2html(cEvtCont.FEHtml)
	end if
	eisort 			= cEvtCont.FEISort
	edid 			= cEvtCont.FEDId
	emid 			= cEvtCont.FEMId
	efwd 			= db2html(cEvtCont.FEFwd)
	ebrand			= cEvtCont.FEBrand
	eicon   		= cEvtCont.FEIcon
	ecommenttitle   = db2html(cEvtCont.FECommentTitle)
	elinkcode   	= cEvtCont.FELinkCode
	dopendate		= cEvtCont.FEOpenDate
	dclosedate		= cEvtCont.FECloseDate
 	blnFull			= cEvtCont.FEFullYN
 	blnIteminfo		= cEvtCont.FEIteminfoYN
 	etag			= db2html(cEvtCont.FETag)

	set cEvtCont = nothing
END IF

Dim arreventstate
arreventstate= fnSetCommonCodeArr("eventstate",False)


%>

<script language="javascript">
function saveOpenGift(){
    if (confirm('저장 하시겠습니까?')){
        frmOpenGift.submit();
    }
}

function jsLastEvent(){
  var winLast,eKind;
  eKind = 1;
  winLast = window.open('pop_event_lastlist.asp?menupos=<%=menupos%>&eventkind='+eKind,'pLast','width=550,height=600, scrollbars=yes')
  winLast.focus();
}

function jsSetImg(sFolder, sImg, sName, sSpan){
	document.domain ="10x10.co.kr";
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimg.asp?yr=<%=Year(eregdate)%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("이미지를 삭제하시겠습니까?\n\n삭제 후 저장버튼을 눌러야 처리완료됩니다.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >

<tr>
	<td> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 오픈 사은 이벤트 지정 </td>
</tr>

<tr>
	<td>
		<table width="1100" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트코드</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%= eCode %>
		   		</td>
		   	</tr>
		   	<tr id="eNameTr_A">
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트명</B></td>
		   		<td bgcolor="#FFFFFF">
		   			<%=ename%>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트기간</B></td>
		   		<td bgcolor="#FFFFFF">
		   			시작일 : <%=esday%>
		   			~ 종료일 : <%=eeday%>
		   		</td>
		   	</tr>

		   	<tr>
		   		<td align="center" bgcolor="<%= adminColor("tabletop") %>"><B>이벤트상태</B></td>
		   		<td bgcolor="#FFFFFF">
		   		    <%= Replace(fnGetCommCodeArrDesc(arreventstate,estate),"오픈예정","오픈") %>

		   			<%IF not isnull(dopendate) THEN%><span style="padding-left:10px;">  오픈처리일 : <%=dopendate%>  </span><%END IF%>
		   			<%IF not isnull(dclosedate) THEN%>/ <span style="padding-left:10px;">  종료처리일 : <%=dclosedate%>  </span><%END IF%>
		   		</td>
		   	</tr>
		</table>
	</td>
</tr>
<tr>
    <td>
    <table width="1100" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frmOpenGift" method="post"  action="openGift_process.asp" onSubmit="return jsEvtSubmit(this);">
    <input type="hidden" name="imod" value="<%= imod %>">
    <input type="hidden" name="menupos" value="<%=menupos%>">
    <input type="hidden" name="OGtitle" value="<%=OGtitle%>">
    <input type="hidden" name="eCode" value="<%=eCode%>">

		   <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>오픈여부</B></td>
		   		<td bgcolor="#FFFFFF">
		   		    <% if (imod="I") then %>
		   		    <input type="radio" name="frontopen" value="Y" disabled >Open
		   			<input type="radio" name="frontopen" value="N" checked >Close
		   			(신규등록시에는 Close로 지정됩니다.)
		   		    <% else %>
		   		    <input type="radio" name="frontopen" value="Y" <%= chkIIF(frontopen="Y","checked","") %> >Open
		   			<input type="radio" name="frontopen" value="N" <%= chkIIF(frontopen="N","checked","") %> >Close
		   			<% end if %>
		   		</td>
		   	</tr>
		   	<tr>
		   	    <td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>전체 사은 구분</B></td>
		   		<td bgcolor="#FFFFFF">
		   		    <% if (imod="I") then %>
		   		    <input type="radio" name="opengiftType" value="1" checked >전체사은 이벤트
		   		    <input type="radio" name="opengiftType" value="9"  >다이어리 이벤트
		   		    <% else %>
		   		    <input type="radio" name="opengiftType" value="1" <%= chkIIF(opengiftType=1,"checked","") %> >전체사은 이벤트
		   		    <input type="radio" name="opengiftType" value="9" <%= chkIIF(opengiftType=9,"checked","") %> >다이어리 이벤트
		   		    <% end if %>
		   		</td>
		   	</tr>
		   	<tr>
		   	    <td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>적용 범위</B></td>
		   		<td bgcolor="#FFFFFF">
		   		    <% if (imod="I") then %>
		   		    <label><input type="radio" name="opengiftScope" value="1" checked >전체</label>
		   		    <label><input type="radio" name="opengiftScope" value="3"  >모바일</label>
		   		    <label><input type="radio" name="opengiftScope" value="5"  >APP</label>
		   		    <% else %>
		   		    <label><input type="radio" name="opengiftScope" value="1" <%= chkIIF(opengiftScope="1","checked","") %> >전체</label>
		   		    <label><input type="radio" name="opengiftScope" value="3" <%= chkIIF(opengiftScope="3","checked","") %> >모바일</label>
		   		    <label><input type="radio" name="opengiftScope" value="5" <%= chkIIF(opengiftScope="5","checked","") %> >APP</label>
		   		    <% end if %>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>타이틀이미지</B></td>
		   		<td bgcolor="#FFFFFF">
		   		    <input type="button" class="button" value="이미지등록" onClick="jsSetImg('<%=eFolder%>','<%=OGtitle%>','OGtitle','spantitle')">
		   		    (장바구니에 표시되는 이미지)
		   		    <div id="spantitle" style="padding: 5 5 5 5">
		   				<%IF OGtitle <> "" THEN %>
		   				<img  src="<%=OGtitle%>" width="100%" />
		   				<a href="javascript:jsDelImg('OGtitle','spantitle');"><img src="/images/icon_delete2.gif" border="0"></a>
		   				<%END IF%>
		   			</div>

		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>웹용 간략설명<br>및<br>유의사항</B></td>
		   		<td bgcolor="#FFFFFF">
		   		<textarea name="openHtmlWeb" cols="120" rows="10"><%=OopenHtmlWeb%></textarea>
<p><font color="blue">샘플코드</font></p>
<textarea name="openHtmlWeb_Sample" cols="120" rows="8" style="border:0">
<li>3/7/15만원 이상 구매시 마일리지, 쿠폰, 할인카드 등의 사용 후 실제 결제금액 기준</li>
<li class="tMar07">텐바이텐 배송상품을 포함하여 3/7/15만원으로 구매시 상품 또는 쿠폰 중에 선택</li>
<li class="tMar07">텐바이텐 배송상품 없이 업체상품만으로 7/15만원으로 구매시 사은품은 쿠폰만 선택가능합니다.<br />(3만원이상 구매시 조건은 쿠폰사은품은 없습니다.)</li>
<li class="tMar07">사은품 쿠폰은 10월 26일 일괄 발급 해드립니다.</li>
<li class="tMar07">사은품 상품만 다른 배송지로 받는것은 불가합니다.</li>
<li class="tMar07">컬러는 랜덤으로 발송되며 교환은 불가합니다.</li>
</textarea>
		   		</td>
		   	</tr>
		   	<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>"><B>모바일용 간략설명<br>및<br>유의사항</B></td>
		   		<td bgcolor="#FFFFFF">
		   		<textarea name="openHtml" cols="140" rows="25"><%=OopenHtml%></textarea>
<p><font color="blue">샘플코드</font></p>
<textarea name="openHtml_Sample" cols="140" rows="25"  style="border:0">
<div id="lyGiftNoti" style="display:none;">
	<div class="layerPopup lyGiftNoti">
		<dl>
			<dt>유의사항</dt>
			<dd>
				<ul class="cartInfoV16a">
					<li>마일리지, 쿠폰 등의 사용 후 구매확정금액 기준</li>
					<li>텐바이텐 배송상품포함 2만원이상 구매시 사은품 증정</li>
					<li>사은품은 텐바이텐 배송상품과 함께 배송</li>
					<li>환불 및 교환으로 인한 구매조건 미달 시, 사은품 및 마일리지 반품 필수</li>
					<li>사은품은 랜덤으로 발송되며 교환 불가</li>
					<li>사은품 소진시 이벤트 종료</li>
				</ul>
			</dd>
		</dl>
		<button class="lyClose" onclick="fnClosePartLayer();">닫기</button>
	</div>
</div>
<div class="bxLGy2V16a grpTitV16a">
	<h2>[오늘은 기분이 좋아 완전 좋아 최고 좋아] 사은품 선택</h2>
	<i class="icoQuestV16a" onClick="fnOpenPartLayer();return false;">사은품 선택 유의사항</i>
</div>
<div class="bxWt1V16a freebieSltV16a">
	<div class="bxWt1V16a">2016.03.12~ 2016.03.28 (조기 소진시 종료)</div></textarea>
		   		</td>
		   	</tr>
    </table>
    </form>
    </td>
</tr>
<tr>
	<td width="800" height="40" align="right">
		<img src="/images/icon_save.gif" onClick="saveOpenGift()"  style="cursor:pointer">
		<a href="/admin/eventmanage/event/openGift.asp?menupos=1184"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>

</table>
<%
set oOpenGift=Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->