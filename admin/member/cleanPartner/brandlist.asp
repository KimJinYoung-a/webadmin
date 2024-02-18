<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/cleanPartnerBrand.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
dim yyyy , mm,d,mend,mdispend, mstart, monthDiff,favCnt
dim makerid,   cdl, cdm, cds
dim OnlySellyn, OnlyIsUsing,danjongyn,mwdiv,mode, dispCate
dim itemid, itemoption
dim arrList, intLoop
dim cCItem
dim iCurrpage, iPageSize,iTotCnt, iTotalPage,iPerCnt
dim chkImg
dim sSort
dim groupid,socname_kr,companyno, crect

iCurrpage   = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
yyyy        = requestCheckVar(Request("yyyy1"),4)
mm          = requestCheckVar(Request("mm1"),2)
favCnt      = requestCheckVar(Request("favCnt"),10)
mwdiv       = requestCheckVar(Request("mwdiv"),2)
OnlySellyn  = requestCheckVar(Request("OnlySellyn"),2)
dispCate    = requestCheckvar(request("disp"),16)
makerid 	= requestCheckvar(Request("makerid"),32)

chkImg      = requestCheckvar(Request("chkImg"),1)
sSort				= requestCheckvar(Request("sSort"),2)

socname_kr  = requestCheckVar(request("socname_kr"),60) 
crect       = RequestCheckVar(request("crect"),32) 
companyno   = RequestCheckVar(request("companyno"),32) 
groupid     = RequestCheckVar(request("groupid"),32)
	
if (yyyy = "") then
	d = CStr(dateadd("m" ,-3, now()))
	yyyy = Left(d,4)
	mm = Mid(d,6,2)
end if

if (monthDiff = "") then
	monthDiff = "3"
end if 
    mend    = dateadd("m",1,yyyy&"-"&mm&"-01") 
    mdispend = dateadd("d",-1,dateadd("m",1,yyyy&"-"&mm&"-01"))
    mstart  =  dateadd("m", monthdiff*-1,mend)
    response.write mend
if favCnt ="" then
    favCnt = 10
end if
   
if OnlySellyn ="" then OnlySellyn ="YS"
if mwdiv ="" then mwdiv ="U"
        
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 50		'한 페이지의 보여지는 열의 수
    iPerCnt = 10		'보여지는 페이지 간격
   
 
	
set cCItem = new CCleanItem
	cCItem.FCurrPage = iCurrpage		'현재페이지
	cCItem.FPageSize = iPageSize	    '페이지사이즈 
	cCItem.FRectStdate   = mstart
  cCItem.FRectEddate   = mend
  cCItem.FRectWishCount=favcnt 
	cCItem.FRectMakerid	 = makerid
	cCItem.FRectDispCate = dispCate
  cCItem.FRectSellYN   = OnlySellyn
  cCItem.FRectMwdiv    = mwdiv
  cCItem.FRectSort     = sSort
  CCItem.FRectsocname_kr= socname_kr  
  CCItem.FRectcrect     = crect       
  CCItem.FRectcompanyno = companyno   
  CCItem.FRectgroupid   = groupid     
  
	arrList = cCItem.fnGetCleanBrandList
	iTotCnt = cCItem.FTotCnt
set cCItem = nothing
iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript"> 
function jsChkAll(){
var frm;
frm = document.frmList;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
	   	   	if(frm.chkitem.disabled==false){
		   	 	frm.chkitem.checked = true;
		   	}
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					 	if(frm.chkitem[i].disabled==false){
					frm.chkitem[i].checked = true;
				}
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkitem) !="undefined"){
	  	if(!frm.chkitem.length){
	   	 	frm.chkitem.checked = false;
	   	}else{
			for(i=0;i<frm.chkitem.length;i++){
				frm.chkitem[i].checked = false;
			}
		}
	  }

	}

}

function jsSetUseYN() {
	var frm = document.frmList;
	 
	if(typeof(frm.chkitem) =="undefined"){
	   return;
     }    
	 	
	if(!frm.chkitem.length){
	    if(!frm.chkitem.checked){
	 		alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
	 		return;
	 	}
	 	
	 	frm.itemidarr.value = frm.chkitem.value;
	 	 
	 }else{
    	for(i=0;i<frm.chkitem.length;i++){
    		if(frm.chkitem[i].checked) {
    	  			if (frm.itemidarr.value==""){
    	      			 frm.itemidarr.value =  frm.chkitem[i].value;
    	  			}else{
    	  	    		 frm.itemidarr.value = frm.itemidarr.value + "," +frm.chkitem[i].value;
    	  			} 
    	  	} 
	  	 }

    	 if (frm.itemidarr.value == ""){
    	 	alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
	 		return;
	  	 }
	 }
	  

	if (confirm("선택 상품을 사용안함 처리하시겠습니까?") == true) {
		frm.submit();
	}
}
 

function jsWinOpen(sUrl){ 
    var winOpen = window.open(sUrl,"popWin", "width=700 height=700 scrollbars=yes resizable=yes");
    winOpen.focus();
}


	 //리스트 정렬
	 function jsSort(sValue,i){  
	  
	 	document.frm.sSort.value= sValue; 
	 	 
		   if (-1 < eval("document.all.img"+i).src.indexOf("_alpha")){
	        document.frm.sSort.value= sValue+"D";  
	    }else if (-1 < eval("document.all.img"+i).src.indexOf("_bot")){
	     		document.frm.sSort.value= sValue+"A";  
	    }else{
	       document.frm.sSort.value= sValue+"D";  
	    } 
	    
	   
		 document.frm.submit();
	}
 </script>
<form name="frm" method="get" action=""> 
<input type="hidden" name="menupos" value="<%= menupos %>"> 
<input type="hidden" name="sSort" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			정리대상조건: 검색기간 <% DrawYMBox yyyy, mm %>  부터 <input type="text" name="monthDiff" value="<%=monthDiff%>" size="4" style="text-align:right" class="input"> 개월 이전까지 <span style="color:gray">[<%=mstart%>~<%=mdispend%>]</span> 판매수량 0 ,
			 
		     &nbsp; &nbsp;
		     위시수 <input type="text" name="favCnt" value="<%=favCnt%>" size="4"  style="text-align:right" class="input"> 개 미만 
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	<tr    bgcolor="#FFFFFF" >	
		<td>
		    <table class="a" >
		        <tr>
		            <td> 
			            브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
            			&nbsp;
            			전시 카테고리:  <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		            </td>		            
	            </tr> 
               <tr>
               		<td>그룹코드 <input type="text" name="groupid" value="<%= groupid %>" Maxlength="32" size="7">
										&nbsp;
										스트리트명(한글) : <input type="text" class="text" name="socname_kr" value="<%= socname_kr %>" Maxlength="32" size="16">
										&nbsp; 
										회사명 <input type="text" name="crect" value="<%= crect %>" Maxlength="32" size="12">
										&nbsp;
										사업자번호 <input type="text" name="companyno" value="<%=companyno %>" Maxlength="32" size="12">
		            </td>		            
              </tr> 	
			    </table>
		</td> 
	</tr> 
</table>
</form> 
<p> 
    <div align="right">
		<input type="button" class="button" value="선택브랜드 [On]종료처리" onClick="jsSetUseYN()">
	</div>
    </p>
<!-- 리스트 시작 -->
<form name="frmList" method="post" action="procClean.asp">
    <input type="hidden" name="hidM" value="I">
	<input type="hidden" name="itemidarr" value="">   
	<input type="hidden" name="sRU" value="itemlist.asp?menupos=<%=menupos%>&makerid=<%=makerid%>&disp=<%=dispCate%>&OnlySellyn=<%=OnlySellyn%>&mwdiv=<%=mwdiv%>&itemid=<%=itemid%>&yyyy=<%=yyyy%>&mm=<%=mm%>&monthdiff=<%=monthdiff%>&favcnt=<%=favcnt%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b> <%=formatnumber(iTotCnt,0)%></b>
			&nbsp;
			페이지 : <b><%=formatnumber(iCurrpage,0)%>/ <%=formatnumber(iTotalPage,0)%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	  <td width="20"><input type="checkbox" name="chkAll"  onClick="jsChkAll()"></td><!--list_lineup_bot_on--> 
		<td>브랜드ID</td> 
		<td>브랜드명(한글)</br>브랜드명(영문)</td> 
		<td>그룹코드<br>사업자번호</td> 
		<td>회사명</td>
		<td>위시수</td> 
		<td>[On]판매중상품수</td>  
		<td>Off사용여부</td>   
		<td>제휴몰사용여부</td>   
		<td>등록일</td>
		
	</tr>
	<%if isArray(arrList) then
	    For intLoop = 0 To UBound(arrList,2) 
	    %>
	<tr bgcolor="#FFFFFF" align="center">
	    <td><input type="checkbox" id="chkitem" name="chkitem" value="<%=arrList(0,intLoop)%>"></td>
	    <td><a href="javascript:jsWinOpen('/admin/member/popbrandadminusing.asp?designer=<%=arrList(0,intLoop)%>');"><%=arrList(0,intLoop)%></a></td>  
	    <td><%=arrList(1,intLoop)%><br><%=arrList(2,intLoop)%></td>
	    <td><%=arrList(3,intLoop)%><br><%=arrList(4,intLoop)%></td>
	    <td><%=arrList(5,intLoop)%></td>
	    <td><%=arrList(7,intLoop)%></td>
	    <td><%=arrList(8,intLoop)%></td>
	    <td><%=arrList(9,intLoop)%></td>
	    <td><%=arrList(10,intLoop)%></td>
	    <td><%=arrList(6,intLoop)%></td>
	</tr>
    <% Next 
    else
    %>
    <tr bgcolor="#FFFFFF" >
        <td colspan="10" align="center">검색조건에 해당하는  내용이 존재하지 않습니다.</td>
    </tr>
    <%
	end if%>
</table>
</form>
<!-- 페이징처리 --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table> 
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->