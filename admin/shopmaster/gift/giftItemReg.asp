<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  사은품 상품등록
' History : 2008.04.04 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemgiftcls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim clsGItem 
Dim gCode, acURL
Dim iTotCnt, arrList,intLoop
Dim iPageSize, iCurrpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
Dim strparm
Dim clsGift
Dim sTitle, dSDay, dEDay, igStatus,igType,igR1,igR2,igkCode,igkType, igkCnt, igkLimit, igkName,sgDelivery
	
gCode = requestCheckVar(Request("gC"),10)
acURL = Server.HTMLEncode("/admin/shopmaster/gift/giftitemProc.asp?gC="&gCode)
strParm = request("strParm")
'=== 코드값이 없을 경우 back
	IF gCode = "" THEN	
%>
		<script language="javascript">
		<!--
			alert("전달값에 문제가 발생하였습니다. 관리자에게 문의해주십시오");
			history.back();
		//-->
		</script>
	<%	dbget.close()	:	response.End
	END IF	

'=== 사은품 정보 
 set clsGift = new CGift 
 	clsGift.FGCode = gCode
 	clsGift.fnGetGiftConts
 	sTitle		= clsGift.FGName
 	igType		= clsGift.FGType      
	igR1		= clsGift.FGRange1     
	igR2 		= clsGift.FGRange2    	
	igkCode		= clsGift.FGKindCode  
	igkType		= clsGift.FGKindType  
	igkCnt		= clsGift.FGKindCnt   
	igkLimit	= clsGift.FGKindlimit
 	dSDay		= clsGift.FSDate   	
	dEDay		= clsGift.FEDate    
	igStatus	= clsGift.FGStatus	
	igkName 	= clsGift.FGKindName
	sgDelivery = clsGift.FGDelivery
  set clsGift = nothing	
 	IF igkLimit = 0 THEN igkLimit = ""	
'=== 파라미터값 받기 & 기본 변수 값 세팅  
	iCurrpage = Request("iC")	'현재 페이지 번호


	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF	  
		
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격

'=== 사은품 대상 상품 리스트 
 set clsGItem = new CGiftItem
 	clsGItem.FGCode = gCode
 	clsGItem.FCPage = iCurrpage
 	clsGItem.FPSize = iPageSize
 	arrList = clsGItem.fnGetItemConts
 	iTotCnt = clsGItem.FTotCnt
 
 
 iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
 
 '공통코드 값 배열로 한꺼번에 가져온 후 값 보여주기
	Dim  arrgifttype,arrgiftstatus		
	arrgifttype 	= fnSetCommonCodeArr("gifttype",False)
	arrgiftstatus 	= fnSetCommonCodeArr("giftstatus",False)
%>
<script language="javascript">
<!--
// 새상품 추가 팝업
function addnewItem(){
		var popwin;
		popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?acURL=<%=acURL%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		popwin.focus();
}

// 페이지 이동
function jsGoPage(iP){
		document.frmitem.iC.value = iP;		
		document.frmitem.submit();	
}

//전체선택
var ichk;
ichk = 1;
	
function jsChkAll(){			
	    var frm, blnChk;
		frm = document.frmitem;
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
//삭제
function jsDel(sType, iValue){	
		var frm;		
		var sValue;		
		frm = document.frmitem;
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
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" class="a">
<form name="frmitem" method="get" action="giftItemReg.asp">
<input type="hidden" name="iC">
<input type="hidden" name="gC" value="<%=gCode%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr>
		<td colspan="2">
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">사은품코드</td>
			<td bgcolor="#FFFFFF"><%=gCode%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">사은품명</td>
			<td bgcolor="#FFFFFF"><%=sTitle%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">기간</td>
			<td bgcolor="#FFFFFF" colspan="3"><%=dSDay%> ~ <%=dEDay%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">한정</td>
			<td bgcolor="#FFFFFF"><%=igkLimit%></td>		
		</tr>
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">상태</td>
			<td bgcolor="#FFFFFF"><%=fnGetCommCodeArrDesc(arrgiftstatus,igStatus)%></td>				
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">증정조건</td>
			<td bgcolor="#FFFFFF"><%=fnGetCommCodeArrDesc(arrgifttype,igType)%>&nbsp; <%=igR1%>이상~ <%=igR2%>미만</td>	
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">수량 </td>
			<td bgcolor="#FFFFFF"><%=igkCnt%> <%IF igkType =2 THEN%>(1+1)<%END IF%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">종류 </td>
			<td bgcolor="#FFFFFF"><%=igkName%>(<%=igkCode%>)
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">배송 </td>
			<td bgcolor="#FFFFFF"><%=fnSetDelivery(sgDelivery)%></td>
			
			
		</tr>
		</table>
		</td>
	</tr>
	<tr height="40" valign="bottom">
		<td align="left">
			<input type="button" value="선택삭제" onClick="jsDel(0,'');" class="button">
		</td>
		<td align="right">	
			<input type="button" value="새상품 추가" onclick="addnewItem();" class="button">
		</td>
	</tr>
	<tr>
		<td colspan="2"> 
			<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			    <tr bgcolor="#FFFFFF">
			   		<td colspan="15" align="left">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
			   	</tr>
			    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			    	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>    				    				    	
			    	<td align="center">상품ID</td>
					<td align="center">이미지</td>
					<td align="center">브랜드</td>
					<td align="center">상품명</td>
					<td align="center">판매가</td>
					<td align="center">매입가</td>
					<td align="center">배송</td>	
					<td align="center">판매여부</td>	
					<td align="center">사용여부</td>	
					<td align="center">한정여부</td>				    
			    	<td>처리</td>
			    </tr>			  
			   <%IF isArray(arrList) THEN 
			    	For intLoop = 0 To UBound(arrList,2)
			   %>
			    <tr align="center" bgcolor="#FFFFFF">    
			    	<td><input type="checkbox" name="chkI" value="<%=arrList(0,intLoop)%>"></td>    				    				    
			    	<td>
			    		<!-- 2007/05/05 김정인 수정 -- 품절 표시 -->			    		
			    		<A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a>
			    		<% if clsGItem.IsSoldOut(arrList(12,intLoop),arrList(14,intLoop),arrList(18,intLoop),arrList(19,intLoop)) then %>
			    			<br><img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
			    		<% end if %>
			    	</td>
			    	<td><% if (Not IsNull(arrList(10,intLoop)) ) and (arrList(10,intLoop)<>"") then %>
					 <img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(10,intLoop)%>">
					<%end if%>
			    	</td>    	
			    	<td><%=db2html(arrList(1,intLoop))%></td>
			    	<td align="left">&nbsp;<%=db2html(arrList(2,intLoop))%></td>
			    	<td><%
						Response.Write FormatNumber(arrList(5,intLoop),0)
						'할인가
						if arrList(16,intLoop)="Y" then
							Response.Write "<br><font color=#F08050>(할)" & FormatNumber(arrList(7,intLoop),0) & "</font>"
						end if
						'쿠폰가
						if arrList(20,intLoop)="Y" then
							Select Case arrList(21,intLoop)
								Case "1"
									Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(5,intLoop)*((100-arrList(22,intLoop))/100),0) & "</font>"
								Case "2"
									Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(5,intLoop)-arrList(22,intLoop),0) & "</font>"
							end Select
						end if
					%></td>
			    	<td><%
			Response.Write FormatNumber(arrList(6,intLoop),0)
			'할인가
			if arrList(16,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(8,intLoop),0) & "</font>"
			end if
			'쿠폰가
			if arrList(20,intLoop)="Y" then
				if arrList(21,intLoop)="1" or arrList(21,intLoop)="2" then
					if arrList(23,intLoop)=0 or isNull(arrList(23,intLoop)) then
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(6,intLoop),0) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(23,intLoop),0) & "</font>"
					end if
				end if
			end if
				%></td>
			    	<td><%= fnColor(clsGItem.IsUpcheBeasong(arrList(13,intLoop)),"delivery")%></td>    	
			    	<td><%= fnColor(arrList(12,intLoop),"yn") %></td>
			    	<td><%= fnColor(arrList(17,intLoop),"yn") %></td>
			    	<td><%= fnColor(arrList(14,intLoop),"yn") %></td>    				    				    
			    	<td><input type="button" value="삭제" onClick="jsDel(1,<%=arrList(0,intLoop)%>);" class="button"></td>	
			    </tr>   
			   <%	Next
			   	ELSE
			   %>
			   	<tr  align="center" bgcolor="#FFFFFF">
			   		<td colspan="12">등록된 내용이 없습니다.</td>
			   	</tr>	
			   <%END IF%>
			</table>
			<!-- 페이징처리 -->
			<%		
			iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	
			
			If (iCurrpage mod iPerCnt) = 0 Then																
				iEndPage = iCurrpage
			Else								
				iEndPage = iStartPage + (iPerCnt-1)
			End If	
			%>
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			    <tr valign="bottom" height="25">			      
			        <td valign="bottom" align="center">
			         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
					<% else %>[pre]<% end if %>
			        <%
						for ix = iStartPage  to iEndPage
							if (ix > iTotalPage) then Exit for
							if Cint(ix) = Cint(iCurrpage) then
					%>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
					<%		else %>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
					<%
							end if
						next
					%>
			    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
					<% else %>[next]<% end if %>
			        </td>
			        <td  width="50" align="right"><a href="giftList.asp?menupos=<%=menupos%>&<%=strparm%>"><img src="/images/icon_list.gif" border="0"></a></td>			        
			    </tr>				  
			 </table>
		</td>	    
	</tr>
		
</form>
</table> 
<%
set clsGItem = nothing
%>
<!-- 선택삭제--->
<form name="frmDel" method="post" action="giftItemProc.asp">
<input type="hidden" name="mode" value="D">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="gC" value="<%=gCode%>">
<input type="hidden" name="itemidarr" value="">
</form>
<!-- 표 하단바 끝-->		
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->