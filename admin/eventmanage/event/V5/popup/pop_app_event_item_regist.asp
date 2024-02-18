<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/v5/popup/pop_app_event_item_regist.asp
' Description :  앱전용 이벤트 상품 등록 메인
' History : 2023.01.09 정태훈 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/appDedicatedEventCls.asp"-->
<%
'변수선언
dim evt_code : evt_code = requestCheckvar(request("evt_code"),10)
Dim oAppDedicated, arrList, intLoop

set oAppDedicated = new AppEventCls
oAppDedicated.FRectEventCode = evt_code
arrList = oAppDedicated.fnGetAppDedicatedItemList
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<link rel="stylesheet" href="http://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">

<style type="text/css">
div.btmLine {background:url(/images/partner/admin_grade.png) left bottom repeat-x; padding-bottom:5px;}    
.tab {position:relative; z-index:50;}
.tab ul {_zoom:1; border-left:1px solid #ccc; border-bottom:1px solid #ccc; list-style:none; margin:0; padding:0;}
.tab ul:after {content:""; display:block; height:0; clear:both; visibility:hidden;}
.tab ul li {float:left; text-align:center;height:23px;padding-top:7px; border:1px solid #ccc; margin:0 0 -1px -1px; cursor:pointer;  background-color:#fff; }
.tab ul li.selected {background-color:#FAECC5; position:relative; font-weight:bold;}
.col11 {width:15% !important;}

select {font-size:12px; vertical-align:top;}
input[type=button], input[type=text] {vertical-align:top;}
</style>
<script>
$(function(){
    $("#datepicker").datepicker({
        changeMonth: true, 
        changeYear: true,
        minDate: '-50y',
        nextText: '다음 달',
        prevText: '이전 달',
        yearRange: 'c-50:c+20',
        showButtonPanel: true, 
        currentText: '오늘 날짜',
        closeText: '닫기',
        dateFormat: "yy-mm-dd",
        showAnim: "slide",
        showMonthAfterYear: true, 
        dayNamesMin: ['월', '화', '수', '목', '금', '토', '일'],
        monthNamesShort: ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월']
    });
    $("#datepicker2").datepicker({
        changeMonth: true, 
        changeYear: true,
        minDate: '-50y',
        nextText: '다음 달',
        prevText: '이전 달',
        yearRange: 'c-50:c+20',
        showButtonPanel: true, 
        currentText: '오늘 날짜',
        closeText: '닫기',
        dateFormat: "yy-mm-dd",
        showAnim: "slide",
        showMonthAfterYear: true, 
        dayNamesMin: ['월', '화', '수', '목', '금', '토', '일'],
        monthNamesShort: ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월']
    });
    $("#datepicker3").datepicker({
        changeMonth: true, 
        changeYear: true,
        minDate: '-50y',
        nextText: '다음 달',
        prevText: '이전 달',
        yearRange: 'c-50:c+20',
        showButtonPanel: true, 
        currentText: '오늘 날짜',
        closeText: '닫기',
        dateFormat: "yy-mm-dd",
        showAnim: "slide",
        showMonthAfterYear: true, 
        dayNamesMin: ['월', '화', '수', '목', '금', '토', '일'],
        monthNamesShort: ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월']
    });
});
function fnAddItem(frm){
    if(frm.itemid.value==""){
        alert("상품코드를 입력하세요.");
    }else if(frm.startdate.value==""){
        alert("시작일을 입력하세요.");
    }else if(frm.startdate.value==""){
        alert("종료일을 입력하세요.");
    }else{
        frm.submit();
    }
}
//삭제
function jsDel(sType, iValue){	
	var frm;		
	var sValue;		
	frm = document.fitem;
	sValue = "";
	
	if (sType ==0) {
		if(!frm.chkI) return;
		
		if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked){
				if (sValue==""){
					sValue = frm.chkI[i].value;		
				}else{
					sValue =sValue+","+frm.chkI[i].value;		
				}	
			}
		}	
		}else{
			if(frm.chkI.checked){
				sValue = frm.chkI.value;
			}	
		}
	
		if (sValue == "") {
			alert('선택 상품이 없습니다.');
			return;
		}
		document.frmDel.itemidarr.value = sValue;
	}else{
		document.frmDel.itemidarr.value = iValue;
	}	
		
	if(confirm("선택하신 상품을 삭제하시겠습니까?")){
		document.frmDel.submit();
	}
}
//전체선택
var ichk;
ichk = 1;
	
function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA		
		if ((e.type=="checkbox")) {				
			e.checked = blnChk ;
		}
	}
}
function fnEventPrizeAdd(episode,itemid){
	var winpop = window.open('/admin/eventmanage/event/v5/popup/pop_appDedicatedEvent_PrizeSet.asp?evt_code=<%=evt_code%>&episode='+episode+'&itemid='+itemid,'winPrize','width=600,height=400,scrollbars=yes,resizable=yes');
}
</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a"  style="padding-top:10px">
	<tr><!-- 검색--->
		<td>     
		    <table cellspacing="5"  bgcolor="FAECC5" width="100%" class="a" cellpadding="0">
		        <tr>
		            <td bgcolor="#FFFFFF"> 
            			<form name="frmitemAdd" method="post" action="appDedicatedItem_process.asp"> 
                        <input type="hidden" name="evt_code" value="<%=evt_code%>">
                        <input type="hidden" name="mode" value="add">
            			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
            				<tr align="center" >
            					<td  width="100" bgcolor="<%= adminColor("tabletop") %>">회차(상품) 추가</td>
            					<td align="left"  bgcolor="#ffffff">  	 
            						<table border="0" cellpadding="1" cellspacing="1" class="a">
            						<tr>
            							<td style="white-space:nowrap;padding-left:10px;">회차: <select name="episode"><option value="1">1</option><option value="2">2</option><option value="3">3</option><option value="4">4</option><option value="5">5</option><option value="6">6</option><option value="7">7</option><option value="8">8</option><option value="9">9</option></select></td>  
            							<td style="white-space:nowrap;padding-left:10px;">상품코드: <input type="text" class="text" name="itemid" size="15" maxlength="10" /></td>
										<td style="white-space:nowrap;padding-left:10px;">당첨수: <input type="text" class="text" name="prize_count" size="4" maxlength="2" /></td>
										<td style="white-space:nowrap;padding-left:10px;">당첨수컬러: <input type="text" class="text" name="prize_count_color" size="15" maxlength="10" /></td>
                                        <td style="white-space:nowrap;padding-left:10px;">시작일: <input type="text" class="text" name="startdate" id="datepicker" size="15" maxlength="10" /></td>
                                        <td style="white-space:nowrap;padding-left:10px;">종료일: <input type="text" class="text" name="enddate" id="datepicker2" size="15" maxlength="10" /></td>
                                        <td style="white-space:nowrap;padding-left:10px;">당첨발표일: <input type="text" class="text" name="prizedate" id="datepicker3" size="15" maxlength="10" />
											<select name="prizetime">
												<option value="0">0시</option>
												<option value="1">1시</option>
												<option value="2">2시</option>
												<option value="3">3시</option>
												<option value="4">4시</option>
												<option value="5">5시</option>
												<option value="6">6시</option>
												<option value="7">7시</option>
												<option value="8">8시</option>
												<option value="9">9시</option>
												<option value="10">10시</option>
												<option value="11">11시</option>
												<option value="12">12시</option>
												<option value="13">13시</option>
												<option value="14">14시</option>
												<option value="15">15시</option>
												<option value="16">16시</option>
												<option value="17">17시</option>
												<option value="18">18시</option>
												<option value="19">19시</option>
												<option value="20">20시</option>
												<option value="21">21시</option>
												<option value="22">22시</option>
												<option value="23">23시</option>
												<option value="24">24시</option>
											</select>
										</td>
                                        <td style="white-space:nowrap;padding-left:10px;"><input type="button" class="button_s" value="등록" onClick="javascript:fnAddItem(this.form);"></td>
            						</tr>
            			        	</table>
            			        </td>
            			    </tr> 
            			</table>
            			</form>
            		</td>
            	</tR> <!-- 검색--->
            	
            	<tr>
            		<td style="padding-top:10px;" valign="top" >  
            		  <div id="divPC">
            		 <form name="fitem" method="post">
                     <input type="hidden" name="mode" value="">
                     <input type="hidden" name="evt_code" value="<%=evt_code%>">
            		  <table width="100%" border="0" align="center" cellpadding="0"  class="a" cellspacing="0"  >	  
            		     <tr>
            		       <td>
                                <input type="button" value="선택삭제" onclick="jsDel(0,'');" class="button">
                      	    </td>
                          </tr>  
                          <tr>
                      	    <td> 
                      			<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>"> 
                      			    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
                      			    	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>    				    	
                       			    	<td>회차</td>
										<td>상품ID</td>
                      					<td>이미지</td>
										<td>상품명</td>
                      					<td>판매가</td>
										<td>당첨수</td>
                      					<td>시작일</td>
                      					<td>종료일</td>
                                        <td>당첨발표일</td>
                                        <td>당첨발표</td>
                      				</tr>
									<tbody>
                      			    <%
									IF isArray(arrList) THEN 
                      			    	For intLoop = 0 To UBound(arrList,2)
                      			    %>
                      			    <tr align="center" bgcolor="#FFFFFF">    
                      			    	<td><input type="checkbox" name="chkI" value="<%=arrList(0,intLoop)%>"></td>    				    	
                      			    	<td><%=arrList(1,intLoop)%></td>
										<td><%=arrList(2,intLoop)%></td>
                                        <td>
                                            <% if (Not IsNull(arrList(7,intLoop))) and (arrList(7,intLoop)<>"") then %>
                                                <img src="http://webimage.10x10.co.kr/image/list/<%=GetImageSubFolderByItemid(arrList(2,intLoop))%>/<%=arrList(7,intLoop)%>" width="50%">
                                            <%end if%>
                                        </td>
										<td><%=db2html(arrList(5,intLoop))%></td>
                                        <td><%=formatnumber(arrList(6,intLoop),0)%>원</td>
                      			    	<td><%=arrList(10,intLoop)%><br><%=arrList(11,intLoop)%></td>
										<td><%=arrList(3,intLoop)%></td>
                                        <td><%=arrList(4,intLoop)%></td>
                                        <td><%=arrList(9,intLoop)%></td>
                                        <td>
                                            <% if arrList(8,intLoop)="N" then %>
                                                <input type="button" value="당첨발표" onclick="fnEventPrizeAdd(<%=arrList(1,intLoop)%>,<%=arrList(2,intLoop)%>);" class="button">
                                            <% else %>
                                                <input type="button" value="당첨자보기" onclick="fnEventPrizeAdd(<%=arrList(1,intLoop)%>,<%=arrList(2,intLoop)%>);" class="button">
                                            <% end if %>
                                        </td>
                      			    </tr>   
                      				<%	
										Next
                      				ELSE
                      				%>
                      			   	<tr align="center" bgcolor="#FFFFFF">
                      			   		<td colspan="19">등록된 내용이 없습니다.</td>
                      			   	</tr>	
                      			   	<%END IF%>
									</tbody>
                      			</table>
                              </td>
                             </tr>
                          </table>
                          </form>    
                      </div> 
            		</td>
            	</tR> 
            </table>
        </tD>
    </tr> 
</table>
<%
	set oAppDedicated = nothing
%>	
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<!-- 선택 삭제--->
<form name="frmDel" method="post" action="appDedicatedItem_process.asp">
<input type="hidden" name="mode" value="del">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
<input type="hidden" name="itemidarr">
</form>
<!-- 표 하단바 끝-->		
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->