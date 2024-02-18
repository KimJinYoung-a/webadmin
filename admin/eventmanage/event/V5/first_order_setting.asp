<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################################
' Page : /admin/eventmanage/event/v5/first_order_setting.asp
' Description :  첫 구매 상품 등록
' History : 2023.05.09 정태훈 생성
'#######################################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/items/first_order_itemcls.asp"-->
<%
'변수선언
dim idx : idx = requestCheckvar(request("idx"),10)
dim itemsort : itemsort = requestCheckvar(request("itemsort"),32)
dim strG : strG = requestCheckvar(Request("selG"),10)
dim makerid : makerid = requestCheckvar(request("makerid"),32)
dim itemid : itemid = request("itemid")
dim itemname : itemname = requestCheckvar(request("itemname"),64)
dim sellyn : sellyn = requestCheckvar(request("sellyn"),2)
dim dispCate : dispCate = requestCheckvar(request("disp"),16)
dim iCurrpage : iCurrpage = Request("iC")	'현재 페이지 번호
dim sorting : sorting = requestCheckvar(Request("sorting"),1)
Dim iTotCnt, arrList, intLoop
Dim iPageSize, iDelCnt, oDealItem
Dim iStartPage, iEndPage, iTotalPage, ix, iPerCnt

if itemsort = "" then itemsort = 3
if iCurrpage = "" then iCurrpage = 1
iPageSize = 105		'한 페이지의 보여지는 열의 수
iPerCnt = 10		'보여지는 페이지 간격

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp) 
		if trim(arrTemp(iA))<>"" then
			'상품코드 유효성 검사(2008.08.05;허진원)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	itemid = left(arrItemid,len(arrItemid)-1)
end if

set oDealItem = new CDealItem
oDealItem.FPSize = iPageSize
oDealItem.FRectMasterIDX = idx
oDealItem.FRectMakerid = makerid
oDealItem.FRectItemid = itemid
oDealItem.FRectItemName = itemname
oDealItem.FRectDispCate = dispCate
oDealItem.FCPage = iCurrpage
oDealItem.FESGroup = strG
oDealItem.FESSort = itemsort
oDealItem.FRectSellYN = sellyn
arrList = oDealItem.fnGetDealEventItemNew
iTotCnt = oDealItem.FTotCnt	'전체 데이터  수
iTotalPage = int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>

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
$(function() {
	$("#sortable").sortable({
		placeholder: "ui-state-highlight",
		cancel : ".sortablearrow",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).find("input[name^='sSort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='sSort']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});

//삭제
function jsDel(){	
	var frm;		
	var sValue;		
	frm = document.fitem;
	sValue = "";
	
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
		
	if(confirm("선택하신 상품을 삭제하시겠습니까?")){
        document.frmDel.isusing.value="N";
		document.frmDel.target="ifrmProc";
		document.frmDel.submit();
	}
}
//순서 정렬
function jsSort(){	
	var frm;
	var sValue, sSort, sDisp ;
	frm = document.fitem;
	sValue = "";
	sSort = "";
	sDisp = ""; 

	var itemid;
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){ 
			if (frm.chkI[i].checked){
					if(!IsDigit(frm.sSort[i].value)){
						alert("순서지정은 숫자만 가능합니다.");
						frm.sSort[i].focus();
						return;
					}
				itemid = frm.chkI[i].value;
				if (sValue==""){
					sValue = frm.chkI[i].value;
				}else{
					sValue =sValue+","+frm.chkI[i].value;
				}
				// 정렬순서
				if (sSort==""){
					sSort = frm.sSort[i].value;
				}else{
					sSort =sSort+","+frm.sSort[i].value;
				}
			}
    	}
	}else{
		sValue = frm.chkI.value;
		if(!IsDigit(frm.sSort.value)){
			alert("순서지정은 숫자만 가능합니다.");
			frm.sSort.focus();
			return;
		} 
		sSort   = frm.sSort.value ;
	}
		document.frmSort.itemidarr.value = sValue;
		document.frmSort.sortarr.value = sSort;
        document.frmSort.target="ifrmProc";
	 	document.frmSort.submit();
}
//수정
function jsEdit(itemid, iValue){
    document.frmEdit.itemid.value=itemid;
    document.frmEdit.isusing.value=iValue;
    document.frmEdit.target="ifrmProc";
    document.frmEdit.submit();
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
function jsAddNewItem(){
	var winpop = window.open('/admin/eventmanage/event/v5/popup/pop_first_order_additemlist.asp','winItempop','width=1024,height=768,scrollbars=yes,resizable=yes');
}
function jsGoPage(iCurrpage){ 
    document.fitem.iC.value = iCurrpage;		
    document.fitem.submit();	
}
</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a"  style="padding-top:10px">
	<tr><!-- 검색--->
		<td>     
		    <table cellspacing="5"  bgcolor="FAECC5" width="100%" class="a" cellpadding="0">
		        <tr>
		            <td bgcolor="#FFFFFF"> 
            			<form name="fsearch" method="post" action="dealitem_regist.asp"> 
            				<input type="hidden" name="idx" value="<%=idx%>">
            				<input type="hidden" name="mode" value="">
            				<input type="hidden" name="selGroup" value="">
            				<input type="hidden" name="itemsort" value="<%=itemsort%>">
            			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
            				<tr align="center" >
            					<td  width="100" bgcolor="<%= adminColor("tabletop") %>">검색 조건</td>
            					<td align="left"  bgcolor="#ffffff">  	 
            						<table border="0" cellpadding="1" cellspacing="1" class="a">
            						<tr>
            							<td style="white-space:nowrap;padding-left:10px;">브랜드: <% drawSelectBoxDesignerWithName "makerid", makerid %></td>  
            							<td style="white-space:nowrap;padding-left:10px;">상품코드:</td>
            							<td style="white-space:nowrap;" rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea> </td>
            						</tr> 
            						<tr>
            						    <td colspan="6">
											<label>상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20" /></label>
										</td>
            						</tr>
            			        	</table>        			        	
            			        </td>
            			        <td   width="100" bgcolor="<%= adminColor("gray") %>">
            						<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
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
                     <input type="hidden" name="iC" value="">
                     <input type="hidden" name="idx" value="<%=idx%>">
                     <input type="hidden" name="selGroup" value="">
                     <input type="hidden" name="selG" value="<%=strG%>">
                     <input type="hidden" name="makerid" value="<%=makerid%>">
                     <input type="hidden" name="itemname" value="<%=itemname%>">
                     <input type="hidden" name="itemid" value="<%=itemid%>">
					 <input type="hidden" name="sorting">
            		  <table width="100%" border="0" align="center" cellpadding="0"  class="a" cellspacing="0"  >	  
            		     <tr>
            		       <td>
                      		
                      			<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a"  >		 
                      			    <tr height="35">			      
                      			        <td align="left">       	
                          			       	<input type="button" value="선택삭제" onclick="jsDel();" class="button">&nbsp;&nbsp;&nbsp;
                                            <input type="button" value="선택정렬" onclick="jsSort();" class="button">
                      			    	</td>
                      			    	<td align="right">
                      			       	    <input type="button" value="새상품 추가" onclick="jsAddNewItem();" class="button"> 
                      			        </td>			      
                      			    </tr>
                      			</table>
                      	    </td>
                          </tr>  
                          <tr>
                      	    <td> 
                      			<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>"> 
                      			    <tr>
                              	        <td colspan="20" bgcolor="#FFFFFF">
                                  	        <table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="0" > 
                                  	            <tr>
                                      		        <td align="left">[검색결과] <b>총: <%=iTotCnt%></b>&nbsp;&nbsp;&nbsp;&nbsp;페이지: <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
                                      		    </tr>
                                      		</table>
                                      	</td>
                                  	</tr>
                      			    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
                      			    	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>    				    	
										<td>상품ID</td>
                      					<td>이미지</td>
										<td>브랜드</td>
                      					<td>상품명</td>
                      					<td>판매가</td>
                      					<td>매입가</td>
                      					<td>할인율</td>
                      					<td>순서</td>
                      					<td>사용유무</td>
                      				</tr>
									<tbody id="sortable">
                      			    <%
									IF isArray(arrList) THEN 
                      			    	For intLoop = 0 To UBound(arrList,2)
                      			    %>
                      			    <tr align="center" bgcolor="<% if arrList(23,intLoop)="Y" then %>#FFFFFF<% else %>#f5f5f5<% end if %>">    
                      			    	<td class="sortablearrow"><input type="checkbox" name="chkI" value="<%=arrList(0,intLoop)%>"></td>    				    	
										<td class="sortablearrow">
											<%=arrList(0,intLoop)%>
                      			    		<% if oDealItem.IsSoldOut(arrList(10,intLoop),arrList(11,intLoop),arrList(12,intLoop),arrList(13,intLoop)) then %>
                      			    			<br><img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
                      			    		<% end if %>
                      			    	</td>
                      			    	<td class="sortablearrow">
											<% if (Not IsNull(arrList(14,intLoop)) ) and (arrList(14,intLoop)<>"") then %>
											<img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(14,intLoop)%>">
											<%end if%>
                      			    	</td>
										<td class="sortablearrow"><%=db2html(arrList(22,intLoop))%></td>
                      			    	<td align="left" class="sortablearrow">&nbsp;<%=db2html(arrList(1,intLoop))%></td>
                      			    	<td class="sortablearrow">
											<%
                      						Response.Write FormatNumber(arrList(4,intLoop),0)
                      						'할인가
                      						if arrList(8,intLoop)="Y" then
                      							Response.Write "<br><font color=#F08050>(할)" & FormatNumber(arrList(6,intLoop),0) & "</font>"
                      						end if
                      						'쿠폰가
                      						if arrList(9,intLoop)="Y" then
                      							Select Case arrList(15,intLoop)
                      								Case "1"
                      									Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(3,intLoop)*((100-arrList(16,intLoop))/100),0) & "</font>"
                      								Case "2"
                      									Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(3,intLoop)-arrList(16,intLoop),0) & "</font>"
                      							end Select
                      						end if
                      						%>
										</td>
                      			    	<td class="sortablearrow">
											<%
											Response.Write FormatNumber(arrList(5,intLoop),0)
											'할인가
											if arrList(8,intLoop)="Y" then
												Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(7,intLoop),0) & "</font>"
											end if
											'쿠폰가
											if arrList(9,intLoop)="Y" then
												if arrList(15,intLoop)="1" or arrList(15,intLoop)="2" then
													if arrList(19,intLoop)=0 or isNull(arrList(19,intLoop)) then
														Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(5,intLoop),0) & "</font>"
													else
														Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(19,intLoop),0) & "</font>"
													end if
												end if
											end if
											%>
										</td>
                      					<td class="sortablearrow">
											<%if arrList(8,intLoop)="Y" then%>
                      						<font color=#F08050><%=CLng(((arrList(4,intLoop)-arrList(6,intLoop))/arrList(4,intLoop))*100)%>%</font>		
                      						<%end if%>
                      						<%
											if arrList(9,intLoop)="Y" then 
												if arrList(15,intLoop)="1" or arrList(15,intLoop)="2" then
													if arrList(19,intLoop)=0 or isNull(arrList(19,intLoop)) then
														Response.Write "<br><font color=#5080F0>" & FormatNumber( arrList(5,intLoop),0) & "</font>"
													else
														Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(16,intLoop),0) 
														if arrList(15,intLoop)="1" then 
														Response.Write "%"
														else
														Response.Write "원"
														end if
														Response.Write "</font>"
													end if
												end if
                      						end if
											%>
                      					</td>
										<% if sorting="Y" then %>
										<td class="sortablearrow"><input type="text" name="sSort" value="<%=intLoop+1%>" size="4" style="text-align:right;"></td>
										<% else %>
                      			    	<td class="sortablearrow"><input type="text" name="sSort" value="<%=arrList(18,intLoop)%>" size="4" style="text-align:right;"></td>
										<% end if %>
                      			    	<td class="sortablearrow">
                                        <%=arrList(23,intLoop)%>
                                        <% if arrList(23,intLoop)="Y" then %>
                                            <input type="button" value="삭제" onClick="jsEdit(<%=arrList(0,intLoop)%>,'N');" class="button">
                                        <% else %>
                                            <input type="button" value="사용" onClick="jsEdit(<%=arrList(0,intLoop)%>,'Y');" class="button">
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
                             <tr>
                          	    <td> <!-- 페이징처리 -->
                                      <%		
                          			iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	
                          			
                          			If (iCurrpage mod iPerCnt) = 0 Then																
                          				iEndPage = iCurrpage
                          			Else								
                          				iEndPage = iStartPage + (iPerCnt-1)
                          			End If	
                          			%>
                          			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" >
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
                          			    </tr>	 
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
	set oDealItem = nothing
%>	
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<!-- 선택삭제--->
<form name="frmDel" method="post" action="/admin/eventmanage/event/v5/popup/first_order_item_process.asp">
<input type="hidden" name="mode" value="delarr">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="isusing">
<input type="hidden" name="itemidarr" value="">
</form>
<!-- 사용 삭제 변경--->
<form name="frmEdit" method="post" action="/admin/eventmanage/event/v5/popup/first_order_item_process.asp">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="itemid" value="<%=itemid%>">
<input type="hidden" name="isusing">
</form> 
<!-- 순서 및 이미지크기 변경--->
<form name="frmSort" method="post" action="/admin/eventmanage/event/v5/popup/first_order_item_process.asp">
<input type="hidden" name="mode" value="sort">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="sortarr" value="">
</form> 
<!-- 표 하단바 끝-->		
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->