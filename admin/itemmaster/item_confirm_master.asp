<%@ language=vbscript %>
<% option explicit %>
<%
'########################################################### 
' Description :  승인대기상품리스트
' History : 2014.01.06 정윤정 수정
'						currstate: 0-승인반려,1-승인대기,2-승인보류,5-승인대기(재요청),7-승인완료,9-업체취소
'						승인반려는 최근 3개월 내역만 보여준다.
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/waititemcls_2014.asp"-->
<%
Dim sListType, sCurrstate,sSort,makerid
Dim clsWait, arrList, intLoop

 sListType =  requestCheckVar(request("sLT"),1)
 sCurrstate =  requestCheckVar(request("selCS"),1)
 sSort =  requestCheckVar(request("sS"),2)
 makerid =  requestCheckVar(request("makerid"),32)
 
 IF sListType = "" THEN sListType = "C"
 IF sCurrstate = "" THEN sCurrstate ="1"
 IF sSort = "" THEN
 	 IF sListType ="C" THEN 	 
 	 		sSort = "CA"	
 	ELSE 
 			sSort = "LD"	
 	END IF
END IF
 set clsWait = new CWaitItemlist2014
	clsWait.FListType 	= sListType
	clsWait.Fcurrstate	= sCurrstate
	clsWait.FSort				= sSort
	clsWait.Fmakerid		= makerid
	arrList = clsWait.fnGetSummaryList
 set clsWait = nothing
%>
<style type="text/css">
.tableList:hover {background-color:lightblue;}
</style>
<script type="text/javascript">
	//검색
	function jsSearch(){
		document.frm.submit();
	}
	
	//리스트 종류선택
	function jsSetList(sListType){
		location.href = "/admin/itemmaster/item_confirm_master.asp?sLT="+sListType+"&menupos=<%= request("menupos") %>";
	}
	
	//상품상세 리스트
	function PopItemconfirm(sListType,makerid,disp){ 
		if (disp ==""){disp="n"};
		var popwin=window.open('item_confirm.asp?sLT=' + sListType + '&makerid='+makerid+'&disp='+disp+'&sCS=<%=sCurrstate%>','_blank','');
		popwin.focus();
	}
	
	//브랜드정보
	function PopUpcheBrandInfoEdit(v){
		window.open("/admin/member/popupchebrandinfo.asp?designer=" + v,"PopUpcheBrandInfoEdit","width=640,height=580,scrollbars=yes,resizabled=yes");
	}
	 
	 //리스트 정렬
	 function jsSort(sValue,i){ 
	 	document.frm.sS.value= sValue;
	 	 
		   if (-1 < eval("img"+i).src.indexOf("_alpha")){
	        document.frm.sS.value= sValue+"D";  
	    }else if (-1 < eval("img"+i).src.indexOf("_bot")){
	     		document.frm.sS.value= sValue+"A";  
	    }else{
	       document.frm.sS.value= sValue+"D";  
	    } 
		 document.frm.submit();
	}
	
	
</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a"> 
	<tr>
		<td><a href="javascript:jsSetList('B');"><%IF sListType="B" THEN%><B>브랜드별</B><%ELSE%>브랜드별<%END IF%> </a> | <a href="javascript:jsSetList('C');"><%IF sListType="C" THEN%><B>카테고리별</B><%ELSE%>카테고리별<%END IF%></a> </td>
	</tr> 
	<tr>
		<td>
			<form name="frm" method="get" action="">
			<input type="hidden" name="page" value="1">
			<input type="hidden" name="menupos" value="<%= request("menupos") %>"> 
			<input type="hidden" name="sS" value=""><!--정렬-->
			<input type="hidden" name="sLT" value="<%=sListType%>"><!--리스트타입(b:브랜드, c:카테고리)-->
				<%IF sListType ="B" THEN%> 
				<table width="100%" border="0" cellpadding="10" cellspacing="1" bgcolor="#CCCCCC" class="a"> 
					<tr align="center" bgcolor="#FFFFFF">
						<td width="50" bgcolor="#EEEEEE">검색조건</td>
						<td align="left">
							브랜드: <%	drawSelectBoxDesignerWithName "makerid", makerid %>&nbsp;&nbsp;
							진행상태:  
							<select name="selCS" class="select">
								<%sbOptItemWaitStatus sCurrState%>
							</select> 
							</td> 
							<td  width="50" bgcolor="#EEEEEE">
								<input type="button" class="button_s" value="검색" onClick="jsSearch();">
							</td> 
					</tr>
				</table>
				<%END IF%>
			</form>
		</td>
	</tr>  
	<tr>
		<td> 
		<%IF sListType ="B" THEN%> 
			<div id="dvBrand">
			<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="<%=adminColor("tablebg")%>" class=a>
				<tr bgcolor="#EEEEEE" align="center">
					<td onClick="javascript:jsSort('B','1');" style="cursor:hand;">브랜드ID <img src="/images/list_lineup<%IF sSort="BD" THEN%>_bot<%ELSEIF sSort="BA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
					<td onClick="javascript:jsSort('N','2');" style="cursor:hand;">브랜드명 <img src="/images/list_lineup<%IF sSort="ND" THEN%>_bot<%ELSEIF sSort="NA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
					<td onClick="javascript:jsSort('W','3');" style="cursor:hand;">승인대기 <img src="/images/list_lineup<%IF sSort="WD" THEN%>_bot<%ELSEIF sSort="WA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img3"></td>
					<td onClick="javascript:jsSort('P','8');" style="cursor:hand;">승인대기(재등록) <img src="/images/list_lineup<%IF sSort="PD" THEN%>_bot<%ELSEIF sSort="PA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img8"></td>
					<td onClick="javascript:jsSort('R','4');" style="cursor:hand;">승인보류 <img src="/images/list_lineup<%IF sSort="RD" THEN%>_bot<%ELSEIF sSort="RA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img4"></td>
					<td onClick="javascript:jsSort('F','5');" style="cursor:hand;">승인반려 <img src="/images/list_lineup<%IF sSort="FD" THEN%>_bot<%ELSEIF sSort="FA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img5"></td>
					<td onClick="javascript:jsSort('C','6');" style="cursor:hand;">대표카테고리 <img src="/images/list_lineup<%IF sSort="CD" THEN%>_bot<%ELSEIF sSort="CA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img6"></td>
					<td onClick="javascript:jsSort('L','7');" style="cursor:hand;">최종등록일 <img src="/images/list_lineup<%IF sSort="LD" THEN%>_bot<%ELSEIF sSort="LA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img7"></td>
					<td>비고</td> 
				</tr> 
				<%IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2)
					%>
				<tr bgcolor="#FFFFFF" align="center" class="tableList">
					<td><a href="javascript:PopUpcheBrandInfoEdit('<%=arrList(0,intLoop)%>');"><%=arrList(0,intLoop)%></a></td>
					<td><%=arrList(1,intLoop)%></td>
					<td><%=arrList(5,intLoop)%></td>
					<td><%=arrList(6,intLoop)%></td>
					<td><%=arrList(7,intLoop)%></td>
					<td><%=arrList(8,intLoop)%></td>
					<td><%=arrList(3,intLoop)%></td>
					<td><%=arrList(9,intLoop)%></td>
					<td><a href="javascript:PopItemconfirm('B','<%=arrList(0,intLoop)%>','<%=arrList(2,intLoop)%>');">승인대기리스트>></a></td> 
				</tr> 
				<%	Next 
				ELSE
				%>
				<tr>
					<td colspan="9" align="center" bgcolor="#FFFFFF">등록된 내용이 없습니다.</td>
				</tr>
				<%
				END IF%>
			</table>
		</div>
		<%ELSE%>
		<div id="dvCategory">
			<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="<%=adminColor("tablebg")%>" class=a>
				<tr bgcolor="#EEEEEE" align="center">
					<td>카테고리코드</td>
					<td onClick="javascript:jsSort('C','1');" style="cursor:hand;">카테고리 <img src="/images/list_lineup<%IF sSort="CD" THEN%>_bot<%ELSEIF sSort="CA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
					<td onClick="javascript:jsSort('W','2');" style="cursor:hand;">승인대기 <img src="/images/list_lineup<%IF sSort="WD" THEN%>_bot<%ELSEIF sSort="WA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
					<td onClick="javascript:jsSort('P','6');" style="cursor:hand;">승인대기(재등록) <img src="/images/list_lineup<%IF sSort="PD" THEN%>_bot<%ELSEIF sSort="PA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img6"></td>
					<td onClick="javascript:jsSort('R','3');" style="cursor:hand;">승인보류 <img src="/images/list_lineup<%IF sSort="RD" THEN%>_bot<%ELSEIF sSort="RA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img3"></td> 
					<td onClick="javascript:jsSort('F','4');" style="cursor:hand;">승인반려 <img src="/images/list_lineup<%IF sSort="FD" THEN%>_bot<%ELSEIF sSort="FA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img4"></td> 
					<td onClick="javascript:jsSort('L','5');" style="cursor:hand;">최종등록일 <img src="/images/list_lineup<%IF sSort="LD" THEN%>_bot<%ELSEIF sSort="LA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img5"></td>
					<td>비고</td> 
				</tr> 
					<%IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2)
					%>
				<tr bgcolor="#FFFFFF" align="center" class="tableList">
					<td><%=arrList(0,intLoop)%></td>
					<td><%=arrList(1,intLoop)%></td>
					<td><%=arrList(3,intLoop)%></td> 
					<td><%=arrList(4,intLoop)%></td>
					<td><%=arrList(5,intLoop)%></td>
					<td><%=arrList(6,intLoop)%></td>
					<td><%IF arrList(7,intLoop) <> "1900-01-01" THEN %><%=arrList(7,intLoop)%><%END IF%></td> 
					<td><a href="javascript:PopItemconfirm('C','','<%=arrList(0,intLoop)%>');">승인대기리스트>></a></td> 
				</tr> 
				<%	Next 
				ELSE
				%>
				<tr>
					<td colspan="8" align="center"  bgcolor="#FFFFFF">등록된 내용이 없습니다.</td>
				</tr>
				<%
				END IF%>
			</table>
		</div>
		<%END IF%>
		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->