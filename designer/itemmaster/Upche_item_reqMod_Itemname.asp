<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체배송상품상품수정요청
' Hieditor : 2014.03.17 정윤정 생성 
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_upche_2014.asp"-->  
<!-- #include virtual="/lib/function_item.asp"-->
<!-- #include virtual="/designer/lib/incPageFunction.asp"-->
<%
dim itemid, itemname, sellyn, limityn, dispCate
dim clsItem 
dim arrList,intLoop
dim iTotCnt, iCurrpage, iPageSize
dim sSort

itemid  = RequestCheckVar(request("itemid"),500) 
itemname = RequestCheckVar(request("itemname"),32) 
sellyn  = RequestCheckVar(request("sellyn"),10) 
limityn = RequestCheckVar(request("limityn"),10)
dispCate = requestCheckvar(request("disp"),16)
iCurrpage= requestCheckvar(request("iCP"),10)
sSort =  requestCheckVar(request("sS"),2)

iPageSize = 30
IF iCurrpage = "" THEN iCurrpage = 1
IF sSort = "" THEN sSort = "ID"	
'상품코드 유효성검사	
if itemid<>"" then
	dim iA ,arrTemp,arrItemid 
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
 
	iA = 0
	do while iA <= ubound(arrTemp) 	
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then 
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if 
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if	

'업체배송 상품리스트 가져오기(텐배제외)	
set clsItem = new CItem
	clsItem.FRectMakerId = session("ssBctID")
	clsItem.FRectItemId = itemid
	clsItem.FRectItemname = itemname
	clsItem.FRectSellYN	= sellYN
	clsItem.FRectLimitYN = limityn
	clsItem.FRectDispCate	= dispCate
	clsItem.FRectSort = sSort
	clsItem.FCurrPage		= iCurrpage
	clsItem.FPageSize		= iPageSize
	arrList = clsItem.fnGetItemUpcheBaesongList
	iTotCnt	= clsItem.FTotCnt
set clsItem = nothing
%>
<style> 
	#dialog {display:none; position:absolute;left:100px;top:100px; z-index:9100;background:#efefef; padding:10px;width:650;}
	#mask {display:none; position:absolute; left:0; top:0; z-index:9000; background:url(http://webadmin.10x10.co.kr/images/mask_bg.png) left top repeat;}
	#boxes .window {position:fixed; left:0; top:0; display:none; z-index:99999;} 
</style> 
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
	//검색
	function jsSearch(){
			//상품코드 숫자&엔터만 입력가능하도록 체크-----------------------------
	var itemid = document.frmSearch.itemid.value;  
	 itemid =  itemid.replace(",","\r");    //콤마는 줄바꿈처리 
		 for(i=0;i<itemid.length;i++){ 
			if ( itemid.charCodeAt(i) != "13" && itemid.charCodeAt(i) != "10" && "0123456789".indexOf(itemid.charAt(i)) < 0){ 
					alert("상품코드는 숫자만 입력가능합니다.");
					return;
			}
		}  
	//---------------------------------------------------------------------
	
	document.frmSearch.submit();
	}
	
//선택상품 수정요청 
function jsChkSubmit(){
	var chkV = false;
	var itemcount = 0;
	var frm = document.frm;
	var len = 64;
	frm.itemidarr.value="";
 	frm.olditemname.value= "";
  frm.itemname.value = "";
  frm.etcStr.value=""
  frm.itemcount.value="";
  		
	if(frm.chkI){ 
		if (!frm.chkI.length){
	        if (frm.chkI.checked){
	        	if(!frm.sCIN.value){
	        		alert("수정할 상품명을 입력해주세요");
	        		frm.sCIN.focus();
	        		return;
	        	}
	        	
	        	if (returnByteCount(frm.sCIN.value) > len)
				{
					alert("상품명은 한글기준 "+parseInt(len/2)+"자, 영문기준 "+len+"자 이내로 해주세요.");
					frm.sCIN.focus();
					return;
				}
				
	          	frm.itemidarr.value=frm.chkI.value;
		 		frm.olditemname.value= frm.sOIN.value;
		  	    frm.itemname.value = frm.sCIN.value;  
	        	chkV = true;
	        	itemcount = 1;
	        }
	   }else{ 
	   	  for (var i=0;i<frm.chkI.length;i++){
	            if (frm.chkI[i].checked){
	            	if(!frm.sCIN[i].value){
				    	alert("수정할 상품명을 입력해주세요");
				        frm.sCIN[i].focus();
				        return;
				     } 
				     
					if (returnByteCount(frm.sCIN[i].value) > len){
					
						alert("상품명은 한글기준 "+parseInt(len/2)+"자, 영문기준 "+len+"자 이내로 해주세요.");
						frm.sCIN[i].focus();
						return;
					}
				   	
				    if( frm.itemidarr.value==""){ 
					      frm.itemidarr.value=frm.chkI[i].value;
						  frm.olditemname.value= frm.sOIN[i].value;
						  frm.itemname.value = frm.sCIN[i].value;  	
					}else{
						  frm.itemidarr.value=frm.itemidarr.value+"|,|"+frm.chkI[i].value;
						  frm.olditemname.value= frm.olditemname.value+"|,|"+frm.sOIN[i].value;
						  frm.itemname.value =frm.itemname.value+"|,|"+ frm.sCIN[i].value;  
					} 
				    itemcount = itemcount + 1;	
	            	chkV = true;
	            }
	      }      
	  }   
	 
	} 
 
	 if (!chkV){
	  	alert("선택된 상품이 없습니다.");
			return;
	  }
	  frm.itemcount.value = itemcount;
	  var maskHeight = $(document).height();
		var maskWidth = $(document).width(); 
		$('#mask').css({'width':maskWidth,'height':maskHeight}); 
		$('#boxes').show();
		$('#mask').show(); 
		$("#dialog").show(); 

}

	$('#mask').click(function () {
		$('#boxes').hide();
		$('.window').hide();
		$('#dialog').hide(); 
	});
	
 
	function jsModSubmit(){
		if(!document.frmMS.sES.value){
			alert("상품수정변경사유를 입력해 주세요.");
			document.frmMS.sES.focus();
			return;
		}
	  	document.frm.etcStr.value = document.frmMS.sES.value; 
		document.frm.submit();
	 
	}
	
	function jsCancel(){
			document.frm.itemidarr.value="";
 			document.frm.olditemname.value= "";
  		document.frm.itemname.value = "";
  		document.frm.etcStr.value=""
  		document.frm.itemcount.value="";
  	
  	 $( "#dialog" ).hide();
  	 $('#mask').hide();
  	 $('#boxes').hide();
	}
	
	//리스트 정렬
function jsSort(sValue,i){ 
	 	document.frmSearch.sS.value= sValue;
	 	 
		   if (-1 < eval("document.frm.img"+i).src.indexOf("_alpha")){
	        document.frmSearch.sS.value= sValue+"D";  
	    }else if (-1 < eval("document.frm.img"+i).src.indexOf("_bot")){
	     		document.frmSearch.sS.value= sValue+"A";  
	    }else{
	       document.frmSearch.sS.value= sValue+"D";  
	    } 
		 document.frmSearch.submit();
	} 
	
	function jsChkClick(idx){
		var len=64; 
		if (!frm.chkI.length){
			document.frm.chkI.checked = true;
			if (returnByteCount(frm.sCIN.value) > len){
				alert("상품명은  한글기준 "+parseInt(len/2)+"자, 영문기준 "+len+"자 이내로 해주세요.");
				frm.sCIN.focus();
				return;
			}
		}else{
			document.frm.chkI[idx].checked = true;
			if (returnByteCount(frm.sCIN[idx].value) > len){ 
				alert("상품명은 한글기준 "+parseInt(len/2)+"자, 영문기준 "+len+"자 이내로 해주세요.");
				frm.sCIN[idx].focus();
				return;
			}
		}	
		
	} 
	 			
</script>
<!-- 표 상단바 시작-->    
<form name="frmSearch" method="get" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sS" value="<%=ssort %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left"> 
			 <table border="0" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>판매:<% drawSelectBoxSellYN "sellyn", sellyn %></td>
				<td> 한정:<% drawSelectBoxLimitYN "limityn", limityn %> </td>
				<td> 상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="30"></td>
				<td> 상품코드: </td>
				<Td rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea></td>
			</tr>
			<tr>
				<td colspan="4">	
					 전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox_upche.asp"-->
				</td> 
			</tr>
			</table>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="jsSearch();">
		</td>
	</tr> 
</table> 
</form> 

<table width="100%" border="0" class="a" >
<tr>
	<td align="left" style="padding-top:5px;">
		<input type="button" class="button" style="width:240px;background-color:#F8DFF0;" value="선택상품명수정요청"   onClick="jsChkSubmit();"/> 
	</td> 
</tr>
</table>
<form name="frm" method="post" action="/designer/itemmaster/upche_item_reqMod_proc.asp">
<input type="hidden" name="hidM" value="N">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="olditemname" value="">
<input type="hidden" name="itemname" value="">
<input type="hidden" name="etcStr" value="">
<input type="hidden" name="itemcount" value=""> 
<input type="hidden" name="menupos" value="<%=menupos%>">
 
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
	 <td colspan="13" align="left">검색조건: <%= formatnumber(iTotCnt,0)%> </td>
 </tr>
 <tr  align="center" bgcolor="<%= adminColor("tabletop") %>">
 	<td><input type="checkbox" name="chkAI" onClick="fnCheckAll(this.checked,frm.chkI);"></td>
 	<td  onClick="javascript:jsSort('I','1');" style="cursor:hand;">상품코드 <img src="/images/list_lineup<%IF sSort="ID" THEN%>_bot<%ELSEIF sSort="IA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
 	<td>이미지</td>
 	<td onClick="javascript:jsSort('N','2');" style="cursor:hand;">상품명 <img src="/images/list_lineup<%IF sSort="ND" THEN%>_bot<%ELSEIF sSort="NA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
 	<td>판매</td>
 	<td>한정</td>
 	<td>판매가</td>
 	<td>공급가</td>
 	<td>공급마진</td>
</tr>
<%IF isArray(arrList) THEN
	For intLoop = 0 To UBound(arrList,2)
	%>
<tr bgcolor="#FFFFFF" align="center">
	<td><input type="checkbox" name="chkI" value="<%=arrList(0,intLoop)%>" onClick="AnCheckClick(this);"></td>
	<td><%=arrList(0,intLoop)%></td>
	<td><img src="<%=webImgUrl%>/image/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(11,intLoop)%>"></td> 
	<td align="left"><input type="hidden" name="sOIN" value="<%=arrList(2,intLoop)%>"><%=arrList(2,intLoop)%> &nbsp;&nbsp;<a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank"><font color="blue">확인하기</font></a>
		<br>->수정: <input type="text" name="sCIN" size="32" maxlength="64" onKeyUp="jsChkClick(<%=intLoop%>);">
		</td>
	<td><%=fnColor(arrList(5,intLoop),"yn")%></td>
	<td><%IF arrList(7,intLoop) ="Y" THEN%> 
      <%= fnColor(arrList(7,intLoop),"yn") %>
       <br>(<%= (arrList(8,intLoop) - arrList(9,intLoop)) %>)
      <% else %>	
       <%= fnColor(arrList(7,intLoop),"yn") %>
      <% end if %>
		
		</td>
<td align="right"><input type="hidden" name="mOSeP" value="<%=arrList(3,intLoop)%>">
		<%=formatnumber(arrList(15,intLoop),0)%> 
		<%'할인가
			if arrList(14,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((arrList(15,intLoop)-arrList(17,intLoop))/arrList(15,intLoop)*100) & "%할)" & FormatNumber(arrList(17,intLoop),0) & "</font>"
			end if
			'쿠폰가
			if arrList(19,intLoop)="Y" then
				IF arrList(20,intLoop) = "1" or arrList(20,intLoop) ="2" THEN 
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(GetCouponAssignPrice(arrList(19,intLoop),arrList(20,intLoop),arrList(22,intLoop),arrList(3,intLoop)),0) & "</font>"
				END IF 
			end if
		%>
		</td>
	<td align="right"><input type="hidden" name="mOSuP" value="<%=arrList(4,intLoop)%>">
		<%=formatnumber(arrList(16,intLoop),0)%> 
	 <%	'할인
	 		if arrList(14,intLoop)="Y" then 
				Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(18,intLoop),0) & "</font>"
			end if
			'쿠폰가
		if arrList(19,intLoop)="Y" then
			IF arrList(20,intLoop) = "1" or arrList(20,intLoop) ="2" THEN 
					if  arrList(21,intLoop)=0 or isNull(arrList(21,intLoop)) then
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(16,intLoop),0) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(21,intLoop),0) & "</font>"
					end if
			END IF 
		END IF
		%>
	</td>
	<td>
		<%=fnPercent(arrList(16,intLoop),arrList(15,intLoop),1)%>
		<% '할인가
			if arrList(14,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>" & fnPercent(arrList(18,intLoop),arrList(17,intLoop),1) & "</font>"
			end if
			'쿠폰가
			if arrList(19,intLoop)="Y" then
					IF arrList(20,intLoop) = "1" or arrList(20,intLoop) ="2" THEN 
			 			if  arrList(21,intLoop)=0 or isNull(arrList(21,intLoop)) then 
							Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(16,intLoop),GetCouponAssignPrice(arrList(19,intLoop),arrList(20,intLoop),arrList(22,intLoop),arrList(3,intLoop)),1) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(21,intLoop),GetCouponAssignPrice(arrList(19,intLoop),arrList(20,intLoop),arrList(22,intLoop),arrList(3,intLoop)),1) & "</font>"
						end if
					END IF
			END IF			 
		%>
	</td>
</tr> 
<%Next
ELSE
%>
<tr  bgcolor="#FFFFFF">
	<td colspan="9" align="center">등록된 내용이 없습니다.</td>
</tr>
<%END IF%>
</table>
</form>
<table width="100%" cellpadding="10" cellspacing="0">
	<tr>
		<td align="center"><%Call sbDisplayPaging("iCP",iCurrpage, iTotCnt,iPageSize, 10,menupos )%></td>
	</tr>
</table>
<div id="boxes">  
<div id="mask"></div>
<div id="dialog">   
<form name="frmMS" method="post"  onsubmit="return false;">
<div style="padding:10px;background-color:#FFFFFF"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 상품명 수정요청<hr></div> 
<table width="100%" border="0" class="a" cellpadding="5" cellspacing="5"  bgcolor="#FFFFFF">
	<tr >
		<td>상품명 수정은 <font color="red">담당MD의 승인 완료 후 사이트에 반영</font>됨을 참고 부탁드립니다.</td>
	</tr>
	<tr>
		<td>
		상품명 수정사유: <input type="text" name="sES" size="45" maxlength="64" value="">	
		</td>
	</tr>
	<tr>
		<td align="center">
			<input type="button" class="button" value="취소" onClick="jsCancel();">
			<input type="button" class="button"  style="color:blue;" value="수정요청" onClick="jsModSubmit();">
		</td>
	</tr>
</table> 
</div> 
</form>


<!-- #include virtual="/lib/db/dbclose.asp" -->