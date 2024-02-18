<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/items/ticketItemCls.asp"-->
<%

dim itemid, oitem

itemid  = requestCheckvar(request("itemid"),10)
if (itemid = "") then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if

'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetOneItem

Dim isTravelItem : isTravelItem=(oitem.FoneItem.Fitemdiv="18")

Dim oticketItem
set oticketItem = new CTicketItem
oticketItem.FRectItemID = itemid
oticketItem.GetOneTicketItem

Dim oticketSchdule
set oticketSchdule = new CTicketSchedule
oticketSchdule.FRectItemID = itemid
oticketSchdule.getTicketSchduleList

Dim i
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>

function popTicketPlcaeInfo(compPlaceName, compPlaceIdx){
    var popwin = window.open('pop_TicketPlaceList.asp?itemid=<%= itemid %>&compPlaceName=' + compPlaceName + '&compPlaceIdx='+compPlaceIdx,'popTicketPlcaeList','width=900,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popTicketPlcaeList(compPlaceName, compPlaceIdx){
    var popwin = window.open('pop_TicketPlaceList.asp?compPlaceName=' + compPlaceName + '&compPlaceIdx='+compPlaceIdx,'popTicketPlcaeList','width=900,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function saveFrm(frm){
    
    if (frm.itemid.value.length<1){
        alert('상품 코드 없음.');
        return;
    }
    
    <% if (isTravelItem) then %>
        if (frm.itemdiv.value!="18"){
            alert('여행 상품만 저장 가능합니다. 먼저 상품 수정에서 여행상품으로 저장후 사용요망');
            return;
        }
        
        /*
        if (frm.ticketPlaceIdx.value.length<1){
            alert('여행장소를 선택하세요.');
            return;
        }
        
        
        if (frm.txGenre.value.length<1){
            alert('공연장르를 입력하세요.');
            frm.txGenre.focus();
            return;
        }
        
        if (frm.txGrade.value.length<1){
            alert('관람등급을 입력하세요.');
            frm.txGrade.focus();
            return;
        }
        
        if (frm.txRunTime.value.length<1){
            alert('관람시간을 입력하세요.');
            frm.txRunTime.focus();
            return;
        }
        */
        
        if (getFieldValue(frm.ticketDlvType)==''){
            alert('티켓 수령 방식을 선택하세요.');
            frm.ticketDlvType[3].focus();
            return;
        }
        
        
        if (frm.bookingStDt.value.length<1){
            alert('예매 시작일을  입력하세요.');
            //frm.bookingStDt.focus();
            return;
        }
        
        if (frm.bookingEdDt.value.length<1){
            alert('예매 종료일을  입력하세요.');
            //frm.bookingEdDt.focus();
            return;
        }
        
        if (frm.stDt.value.length<1){
            alert('여행 시작일을  입력하세요.');
            //frm.stDt.focus();
            return;
        }
        
        if (frm.edDt.value.length<1){
            alert('여행 종료일을  입력하세요.');
            //frm.edDt.focus();
            return;
        }
        
        //실 Schedule
        if (frm.Tk_StSchedule.length){
            for(var i=0;i<frm.Tk_StSchedule.length;i++){
                if (frm.Tk_StSchedule[i].value.length<1){
                    alert('여행일정 - 시작일을  입력하세요.');
                    return;
                }
                
                if (frm.Tk_StScheduleTime[i].value.length!=8){
                    alert('여행일정 - 시작시간  입력하세요.');
                    frm.Tk_StScheduleTime[i].focus();
                    return;
                }
                
                if (frm.Tk_EdSchedule[i].value.length<1){
                    alert('여행일정 - 종료일을  입력하세요.');
                    return;
                }
                
                if (frm.Tk_EdScheduleTime[i].value.length!=8){
                    alert('여행일정 - 종료시간  입력하세요.');
                    frm.Tk_EdScheduleTime[i].focus();
                    return;
                }
            }
        }else{
            if (frm.Tk_StSchedule.value.length<1){
                alert('공연일정 - 시작일을  입력하세요.');
                return;
            }
            
            if (frm.Tk_StScheduleTime.value.length!=8){
                alert('공연일정 - 시작시간  입력하세요.');
                frm.Tk_StScheduleTime.focus();
                return;
            }
            
            if (frm.Tk_EdSchedule.value.length<1){
                alert('공연일정 - 종료일을  입력하세요.');
                return;
            }
            
            if (frm.Tk_EdScheduleTime.value.length!=8){
                alert('공연일정 - 종료시간  입력하세요.');
                frm.Tk_EdScheduleTime.focus();
                return;
            }
        }
    <% else %>
        if (frm.itemdiv.value!="08"){
            alert('티켓 상품만 저장 가능합니다. 먼저 상품 수정에서 티켓상품으로 저장후 사용요망');
            return;
        }
        
        if (frm.ticketPlaceIdx.value.length<1){
            alert('공연장소를 선택하세요.');
            return;
        }
        
        if (frm.txGenre.value.length<1){
            alert('공연장르를 입력하세요.');
            frm.txGenre.focus();
            return;
        }
        
        if (frm.txGrade.value.length<1){
            alert('관람등급을 입력하세요.');
            frm.txGrade.focus();
            return;
        }
        
        
        if (getFieldValue(frm.ticketDlvType)==''){
            alert('티켓 수령 방식을 선택하세요.');
            frm.ticketDlvType[3].focus();
            return;
        }
        
        
        if (frm.txRunTime.value.length<1){
            alert('관람시간을 입력하세요.');
            frm.txRunTime.focus();
            return;
        }
        
        if (frm.bookingStDt.value.length<1){
            alert('예매 시작일을  입력하세요.');
            //frm.bookingStDt.focus();
            return;
        }
        
        if (frm.bookingEdDt.value.length<1){
            alert('예매 종료일을  입력하세요.');
            //frm.bookingEdDt.focus();
            return;
        }
        
        if (frm.stDt.value.length<1){
            alert('공연 시작일을  입력하세요.');
            //frm.stDt.focus();
            return;
        }
        
        if (frm.edDt.value.length<1){
            alert('공연 종료일을  입력하세요.');
            //frm.edDt.focus();
            return;
        }
        
        //실 Schedule
        if (frm.Tk_StSchedule.length){
            for(var i=0;i<frm.Tk_StSchedule.length;i++){
                if (frm.Tk_StSchedule[i].value.length<1){
                    alert('공연일정 - 시작일을  입력하세요.');
                    return;
                }
                
                if (frm.Tk_StScheduleTime[i].value.length!=8){
                    alert('공연일정 - 시작시간  입력하세요.');
                    frm.Tk_StScheduleTime[i].focus();
                    return;
                }
                
                if (frm.Tk_EdSchedule[i].value.length<1){
                    alert('공연일정 - 종료일을  입력하세요.');
                    return;
                }
                
                if (frm.Tk_EdScheduleTime[i].value.length!=8){
                    alert('공연일정 - 종료시간  입력하세요.');
                    frm.Tk_EdScheduleTime[i].focus();
                    return;
                }
            }
        }else{
            if (frm.Tk_StSchedule.value.length<1){
                alert('공연일정 - 시작일을  입력하세요.');
                return;
            }
            
            if (frm.Tk_StScheduleTime.value.length!=8){
                alert('공연일정 - 시작시간  입력하세요.');
                frm.Tk_StScheduleTime.focus();
                return;
            }
            
            if (frm.Tk_EdSchedule.value.length<1){
                alert('공연일정 - 종료일을  입력하세요.');
                return;
            }
            
            if (frm.Tk_EdScheduleTime.value.length!=8){
                alert('공연일정 - 종료시간  입력하세요.');
                frm.Tk_EdScheduleTime.focus();
                return;
            }
        }
    <% end if %>
    
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}

function setDefault(comp,defaultVal){
    if (comp.value==''){comp.value=defaultVal};
}

</script>


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <form name="ticketreg" method="post" action="ticketItem_Process.asp" onsubmit="return false;">
  <input type="hidden" name="mode" value="ticketInfo">
  <input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
  <input type="hidden" name="itemdiv" value="<%= oitem.FOneItem.Fitemdiv %>">
  <input type="hidden" name="makerid" value="<%= oitem.FOneItem.Fmakerid %>">
  
  <tr align="left" bgcolor="#FFFFFF">
    <td height="30" colspan="4">
    <strong>** 차후 상품 수정 페이지와 통합..</strong>
    </td>
  </tr>
  <tr align="left" bgcolor="F4F4F4">
    <td height="30" colspan="4">
    <%= CHKIIF(isTravelItem,"여행","공연") %> 기본정보
    </td>
  </tr>
  
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 상품코드 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <%= oitem.FOneItem.Fitemid %>
  	  &nbsp;&nbsp;&nbsp;&nbsp;
  	  <input type="button" value="미리보기" onclick="window.open('http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FOneItem.Fitemid %>');">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 업체ID :</td>
  	<td bgcolor="#FFFFFF" colspan="3"><%=oitem.FOneItem.FMakerid %></td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* <%= CHKIIF(isTravelItem,"여행","공연") %>명 :</td>
  	<td bgcolor="#FFFFFF" colspan="3"><%= oitem.FOneItem.Fitemname %></td>
  </tr>
  
  <!-- 공연 기본정보 -->
  <% if (isTravelItem) then %>
  <input type="hidden" name="ticketPlaceName" value="">
  <input type="hidden" name="ticketPlaceIdx" value="0">
  <input type="hidden" name="txGenre" value="">
  <input type="hidden" name="txGrade" value="">
  <input type="hidden" name="txRunTime" value="">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 취소수수료 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="bookingCharge" value="<%= oticketItem.FOneItem.FbookingCharge %>" size="10" class="text" maxlength="10"  >
  	
  	</td>
  </tr>
  
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 티켓 수령 방식 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	
  	
  	<input type="radio" name="ticketDlvType" value="1" disabled <%= CHKIIF(oticketItem.FOneItem.FticketDlvType="1","checked","") %> > 현장수령
  	<input type="radio" name="ticketDlvType" value="2" <%= CHKIIF(oticketItem.FOneItem.FticketDlvType="2","checked","") %> > 일반배송
  	<input type="radio" name="ticketDlvType" value="3" disabled <%= CHKIIF(oticketItem.FOneItem.FticketDlvType="3","checked","") %> > 현장수령 or 일반배송 택일
  	<input type="radio" name="ticketDlvType" value="9" disabled <%= CHKIIF(oticketItem.FOneItem.FticketDlvType="9","checked","") %> > 티켓 현장수령 및 사은품 상품은 배송
  	
  	<br>(현재 [일반배송] 방식만 유효 )
  	</td>
  </tr>
  
  <% else %>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* <%= CHKIIF(isTravelItem,"여행","공연") %>장소 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="ticketPlaceName" value="<%= oticketItem.FOneItem.FticketPlaceName %>" size="50" class="text_ro" readonly >
  	<input type="hidden" name="ticketPlaceIdx" value="<%= oticketItem.FOneItem.FticketPlaceIdx %>" >
  	<input type="button" style="button" value="관람장소 선택" onclick="popTicketPlcaeList('ticketPlaceName','ticketPlaceIdx');">
  	</td>
  </tr>
  
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 장르 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	
  	<input type="text" name="txGenre" value="<%= oticketItem.FOneItem.FtxGenre %>" size="20" class="text" maxlength="30" >
  	(ex 뮤지컬, 콘서트, 연극, 클래식, 무용 등)
  	</td>
  </tr>
  
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 관람등급 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="txGrade" value="<%= oticketItem.FOneItem.FtxGrade %>" size="30" class="text" maxlength="64" >
  	(ex 전체관람, 만 18세 이상 등)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 관람시간 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="txRunTime" value="<%= oticketItem.FOneItem.FtxRunTime %>" size="30" class="text" maxlength="32" >
  	(ex 120분, 100분 등)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 예매수수료 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" name="bookingCharge" value="<%= oticketItem.FOneItem.FbookingCharge %>" size="10" class="text" maxlength="10" readOnly >
  	(당분간 수수료 없음 0원)
  	</td>
  </tr>
  
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 티켓 수령 방식 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	
  	
  	<input type="radio" name="ticketDlvType" value="1" <%= CHKIIF(oticketItem.FOneItem.FticketDlvType="1","checked","") %> > 현장수령
  	<input type="radio" name="ticketDlvType" value="2" disabled <%= CHKIIF(oticketItem.FOneItem.FticketDlvType="2","checked","") %> > 일반배송
  	<input type="radio" name="ticketDlvType" value="3" disabled <%= CHKIIF(oticketItem.FOneItem.FticketDlvType="3","checked","") %> > 현장수령 or 일반배송 택일
  	<input type="radio" name="ticketDlvType" value="9" <%= CHKIIF(oticketItem.FOneItem.FticketDlvType="9","checked","") %> > 티켓 현장수령 및 사은품 상품은 배송
  	
  	<br>(현재 [티켓 현장수령] 및 [사은품 상품은 배송] 방식만 유효 )
  	</td>
  </tr>
  
  <% end if %>
  
  
  
  
  
  <tr align="left" bgcolor="F4F4F4">
    <td height="30" colspan="4">
    <%= CHKIIF(isTravelItem,"여행","공연") %> 일시
    </td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 예매가능일시 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	예매시작일 <input id="bookingStDt" name="bookingStDt" value="<%= Left(oticketItem.FOneItem.FbookingStDt,10) %>" class="text" size="10" maxlength="10" onChange="setDefault(ticketreg.bookingStDtTime,'00:00:00')" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="bookingStDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
  	           &nbsp;<input type="text" name="bookingStDtTime" value="<%= Right(oticketItem.FOneItem.FbookingStDt,8) %>" size="8" maxlength="8" class="text"> 
  	~
  	예매종료일 <input id="bookingEdDt" name="bookingEdDt" value="<%= Left(oticketItem.FOneItem.FbookingEdDt,10) %>" class="text" size="10" maxlength="10" onChange="setDefault(ticketreg.bookingEdDtTime,'23:59:59')" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="bookingEdDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
  	           &nbsp;<input type="text" name="bookingEdDtTime" value="<%= Right(oticketItem.FOneItem.FbookingEdDt,8) %>" size="8" maxlength="8" class="text">
  	<script language="javascript">
		var BKG_Start = new Calendar({
			inputField : "bookingStDt", trigger    : "bookingStDt_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				BKG_End.args.min = date;
				BKG_End.redraw();
				setDefault(ticketreg.bookingStDtTime,'00:00:00');
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		var BKG_End = new Calendar({
			inputField : "bookingEdDt", trigger    : "bookingEdDt_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				BKG_Start.args.max = date;
				BKG_Start.redraw();
				setDefault(ticketreg.bookingEdDtTime,'23:59:59');
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
	</script>
  	(ex <%= Left(now(),10) %>&nbsp;23:59:59)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">* 전체<%= CHKIIF(isTravelItem,"여행","공연") %>일정 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<%= CHKIIF(isTravelItem,"여행","공연") %>시작일 <input id="stDt" name="stDt" value="<%= Left(oticketItem.FOneItem.FstDt,10) %>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="stDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
  	~
  	<%= CHKIIF(isTravelItem,"여행","공연") %>종료일 <input id="edDt" name="edDt" value="<%= Left(oticketItem.FOneItem.FedDt,10) %>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="edDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
  	<script language="javascript">
		var CAL_Start = new Calendar({
			inputField : "stDt", trigger    : "stDt_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_End.args.min = date;
				CAL_End.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		var CAL_End = new Calendar({
			inputField : "edDt", trigger    : "edDt_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
	</script>
				
  	</td>
  </tr>
  
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF"><%= CHKIIF(isTravelItem,"여행","공연") %>시간 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	<input type="text" class="text" name="txplayTimInfo" value="<%= oticketItem.FOneItem.FtxplayTimInfo %>" size="80" maxlength=250>
  	<br>(ex 수-금 8시 / 토 4시, 8시 / 일 7시 (월,화 공연없음))
  	</td>
  </tr>
  <tr >
  	<td height="30" width="15%" bgcolor="#DDDDFF" align="right"><%= CHKIIF(isTravelItem,"여행","관람") %>일정 </td>
  	<td bgcolor="#FFFFFF" colspan="3" align="left">
      	<table width="96%" border=0 cellspacing=2 cellpadding=3 bgcolor="#CCCCCC">
      	<tr bgcolor="#FFFFFF">
      	    <td width="60">일정코드</td>
      	    <td width="160">일정명</td>
      	    <td>일정(YYYY-MM-DD HH:NN:SS)</td>
      	    <td width="120">취소마감일</td>
      	</tr>
      	<% for i=0 to oticketSchdule.FResultCount-1%>
      	<tr bgcolor="#FFFFFF">
      	    <td><%= oticketSchdule.FItemList(i).FTk_itemoption %></td>
      	    <td><%= oticketSchdule.FItemList(i).FTk_optName %></td>
      	    <td>
      	        <input type="hidden" name="Tk_itemoption" value="<%= oticketSchdule.FItemList(i).FTk_itemoption %>">
      	        
      	        <input type="text" class="text" name="Tk_StSchedule" value="<%= Left(oticketSchdule.FItemList(i).FTk_StSchedule,10) %>" size="10" maxlength="10">
      	        <input type="text" class="text" name="Tk_StScheduleTime" value="<%= Right(oticketSchdule.FItemList(i).FTk_StSchedule,8) %>" size="8" maxlength="8">
      	        ~
      	        <input type="text" class="text" name="Tk_EdSchedule" value="<%= Left(oticketSchdule.FItemList(i).FTk_EdSchedule,10) %>" size="10" maxlength="10">
      	        <input type="text" class="text" name="Tk_EdScheduleTime" value="<%= Right(oticketSchdule.FItemList(i).FTk_EdSchedule,8) %>" size="8" maxlength="8">
      	        
      	        
      	    </td>
      	    <td>
      	        <% IF IsNULL(oticketSchdule.FItemList(i).FreturnExpireDate) THEN %>
      	        자동계산
      	        <% else %>
      	        <%= oticketSchdule.FItemList(i).FreturnExpireDate %>
      	        <% end if %>
      	        </td>
      	</tr>
      	<% next %>
      	</table>
  	</td>
  </tr>
  
  <!--
  <tr align="left" bgcolor="F4F4F4">
    <td height="30" colspan="4">
    티켓 좌석 종류
    </td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">좌석 종류 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	차후 개발 (일반석, R석, S석 등 생길경우)
  	</td>
  </tr>
  -->
  <tr >
  	<td height="30" bgcolor="#FFFFFF" colspan="4" align="center">
  	<input type="button" value=" 저 장 " onClick="saveFrm(ticketreg);">
  	</td>
  </tr>
  </form>
</table>

<%
set oitem = Nothing
set oticketItem = Nothing
set oticketSchdule = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->