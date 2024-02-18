<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2010.10.11 한용민 생성
' Description : 상품 추가 - 할인, 사은품 상품등록에 사용
'				input - actionURL(db 처리에 필요한 파라미터까지 포함) ex.acURL = "/admin/eventmanage/event/eventitem_process.asp?eC=1234"
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%
dim target, actionURL ,page ,cdl, cdm, cds ,i
dim lectureridx, lecturer_name, makerid, disp_yn, usingyn, deliverytype, limityn, vatyn, sailyn, couponyn, mwdiv,defaultmargin
	actionURL	= request("acURL")
	lectureridx      = request("lectureridx")
	lecturer_name    = RequestCheckvar(request("lecturer_name"),32)
	makerid     = RequestCheckvar(request("makerid"),32)
	disp_yn      = RequestCheckvar(request("disp_yn"),1)
	usingyn     = RequestCheckvar(request("usingyn"),1)
	mwdiv       = RequestCheckvar(request("mwdiv"),1)
	limityn     = RequestCheckvar(request("limityn"),1)
	sailyn      = RequestCheckvar(request("sailyn"),1)
	couponyn	= RequestCheckvar(request("couponyn"),1)
	defaultmargin = RequestCheckvar(request("defaultmargin"),10)
	deliverytype       = RequestCheckvar(request("deliverytype"),1)
	cdl = RequestCheckvar(request("cdl"),3)
	cdm = RequestCheckvar(request("cdm"),3)
	cds = RequestCheckvar(request("cds"),3)
	page = RequestCheckvar(request("page"),10)
  	if actionURL <> "" then
		if checkNotValidHTML(actionURL) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
	if lectureridx <> "" then
		if checkNotValidHTML(lectureridx) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	if (page="") then page=1
	if sailyn="" and instr(actionURL,"saleitem")>0 then sailyn="N"			'할인페이지에서 검색된거라면 기본값: 할인안함(쿠폰도 동일)
	if couponyn="" and instr(actionURL,"saleitem")>0 then couponyn="N"
	'if disp_yn = "" then disp_yn ="Y"

if lectureridx<>"" then
	dim iA ,arrTemp,arrlectureridx

	arrTemp = Split(lectureridx,",")

	iA = 0
	do while iA <= ubound(arrTemp)

		if trim(arrTemp(iA))<>"" then
			'상품코드 유효성 검사(2008.08.04;허진원)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrlectureridx = arrlectureridx & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	lectureridx = left(arrlectureridx,len(arrlectureridx)-1)
end if

dim oitem
set oitem = new CLecture
	oitem.FPageSize         = 30
	oitem.FCurrPage         = page
	oitem.frectlecturer_id      = makerid
	oitem.FRectlectureridx       = lectureridx
	oitem.FRectlecturer_name     = lecturer_name
	oitem.FRectdisp_yn       = disp_yn
	oitem.FRectIsUsing      = usingyn
	oitem.FRectCouponYn		= couponyn
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.GetlecturerList()		
%>

<script language="javascript">

function jsSerach(){
	var frm;
	frm = document.frm;
	frm.target = "_self";
	frm.action ="pop_lecturerAddInfo.asp";
	frm.submit();
}

function SelectItems(sType){	
var frm;
var tmp = 0 ;
var itemcount = 0;
frm = document.frm;
frm.sType.value = sType;   //전체선택 or 선택상품 여부 구분

	if (sType == "sel"){
		 if(typeof(frm.chkitem) !="undefined"){
	   	   	if(!frm.chkitem.length){
	   	   		if(!frm.chkitem.checked){
	   	   			alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
	   	   			return;
	   	   		}
	   	   		 frm.lectureridxarr.value = frm.chkitem.value;
	   	   		 itemcount = 1;
	   	    }else{
	   	    
	   	    	for(i=0;i<frm.chkitem.length;i++){
	   	    		if(frm.chkitem[i].checked) {	   	    			
	   	    			if (frm.lectureridxarr.value==""){
							frm.lectureridxarr.value =  frm.chkitem[i].value;
	   	    			}else{
							frm.lectureridxarr.value = frm.lectureridxarr.value + "," +frm.chkitem[i].value;
	   	    			}
	   	    		tmp = tmp + 1	
	   	    		}
	   	    		
	   	    		itemcount = frm.chkitem.length;
	   	    	}
	   	    	if (tmp > 1){
	   	    		alert("강좌쿠폰은 하나의 쿠폰에 하나의 강좌만 선택 하실수 있습니다");
	   	   			return;
	   	    	}
	   	    		   	    	
	   	    	if (frm.lectureridxarr.value == ""){
	   	    		alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
	   	   			return;
	   	    	}
	   	    }
	   	  }else{
	   	  	alert("추가할 상품이 없습니다.");
	   	  	return;
	   	  } 
	}else{
		if(typeof(frm.chkitem) !="undefined"){
			itemcount = "<%= oitem.FTotalCount%>";
		  if(confirm(itemcount +"건의 검색된 모든 상품을 추가하시겠습니까?")){
		  	if(itemcount > 1000) {
		  		alert("상품은 최대 1000건까지 가능합니다. 조건을 다시 설정해주세요");
		  		return;
		  	}
			frm.lectureridxarr.value = frm.lectureridx.value;
			
		  }else{
		  	return;
		  }
		}else{
		 	alert("추가할 상품이 없습니다.");
	   	  	return;
		}	
	}
	
	//frm.target = opener.name;
	frm.target = "FrameCKP";
	frm.action = "<%=actionURL%>";
	frm.itemcount.value = itemcount;
	frm.submit();
	frm.lectureridxarr.value = "";
	frm.itemcount.value = 0;	
	opener.history.go(0);	
	//window.close();
}

//전체 선택
function jsChkAll(){	
var frm;
frm = document.frm;
	if (frm.chkAll.checked){			      
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
		   	 	frm.chkitem.checked = true;	   	 
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					frm.chkitem[i].checked = true;
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

// 페이지 이동
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.target = "_self";
	document.frm.action ="pop_lecturerAddInfo.asp";
	document.frm.submit();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">	
<input type="hidden" name="page" >
<input type="hidden" name="sType" >
<input type="hidden" name="lectureridxarr" >
<input type="hidden" name="itemcount" value="0">
<input type="hidden" name="mode" value="I">
<input type="hidden" name="acURL" value="<%=actionURL%>">
<input type="hidden" name="defaultmargin" value="<%=defaultmargin%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<!-- include virtual="/common/module/categoryselectbox.asp"-->
		<br>
		브랜드 :<% drawSelectBoxLecturer "makerid", makerid %>		
		강좌코드 :
		<input type="text" class="text" name="lectureridx" value="<%= lectureridx %>" size="40" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">			
		<br>
		강좌명 :
		<input type="text" class="text" name="lecturer_name" value="<%= lecturer_name %>" size="32" maxlength="20">			
		<div style="font-size:11px; color:gray;padding-left:60px;">(쉼표로 복수입력가능)</div>
	</td>
	
	<td rowspan="2" width="30" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSerach();">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td align="left">
		전시:<% drawSelectBoxUsingYN "disp_yn", disp_yn %>
     	 
     	사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
         	     	     	      	    	     
     	쿠폰: <% drawSelectBoxCouponYN "couponyn", couponyn %>
     	     	
	</td>
</tr>    
</table>

<table width="100%" height="40" align="center" cellpadding="3" cellspacing="1" class="a" border="0">	
	<tr>
		<td  valign="bottom">				
				<input type="button" value="선택상품 추가" onClick="SelectItems('sel')" class="button">
				<!-- 전체선택 사용안함 하나의 강좌쿠폰에 하나의 강좌만 등록 가능-->
				<input type="button" value="전체선택 추가" onClick="SelectItems('all')" class="button" disabled >
		</td>				
	</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr  bgcolor="#FFFFFF">
	<td colspan="13">
	검색결과 : <b><%= oitem.FTotalCount%></b>
	&nbsp;
	페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>		
</tr>		
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td align="center">상품ID</td>
	<td align="center">이미지</td>
	<td align="center">강사</td>
	<td align="center">상품명</td>
	<td align="center">판매가</td>
	<td align="center">매입가</td>
	<td align="center" nowrap>전시<br>여부</td>	
	<td align="center" nowrap>사용<br>여부</td>		
</tr>
<% if oitem.FresultCount<1 then %>
<tr bgcolor="#FFFFFF" >
	<td colspan="13" align="center">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<tr class="a" height="25" bgcolor="#FFFFFF">
	<td  align="center"><input type="checkbox" name="chkitem" value="<%= oitem.FItemList(i).Flectureridx %>"></td>
	<td align="center"><A href="<%=wwwFingers%>/lecture/lecturedetail.asp?lectureridx=<%= oitem.FItemList(i).Flectureridx %>" target="_blank"><%= oitem.FItemList(i).Flectureridx %></a></td>
	<td align="center"><%IF oitem.FItemList(i).FSmallImage <> "" THEN%><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border=0 alt=""><%END IF%></td>
	<td align="center"><% =oitem.FItemList(i).flecturer_id %><br>(<% =oitem.FItemList(i).Flecturer_name %>)</td>
	<td>&nbsp;<% =oitem.FItemList(i).flec_title %></td>
	<td align="center">
		<%
		Response.Write FormatNumber(oitem.FItemList(i).Fsellcash,0)
		'쿠폰가
		if oitem.FItemList(i).FlecturerCouponYn="Y" then
			Select Case oitem.FItemList(i).FlecturerCouponType
				Case "1"
					Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Fsellcash*((100-oitem.FItemList(i).FlecturerCouponValue)/100),0) & "</font>"
				Case "2"
					Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).Fsellcash-oitem.FItemList(i).FlecturerCouponValue,0) & "</font>"
			end Select
		end if
		%>
	</td>
	<td align="center">
		<%
		Response.Write FormatNumber(oitem.FItemList(i).Fbuycash,0)
		'쿠폰가
		if oitem.FItemList(i).FlecturerCouponYn="Y" then
			if oitem.FItemList(i).FlecturerCouponType="1" or oitem.FItemList(i).FlecturerCouponType="2" then
				if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
					Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Forgsuplycash,0) & "</font>"
				else
					Response.Write "<br><font color=#5080F0>" & FormatNumber(oitem.FItemList(i).Fcouponbuyprice,0) & "</font>"
				end if
			end if
		end if
		%>
	</td>
	<td align="center">
		<%= oitem.FItemList(i).Fdisp_yn %>
	</td>
	<td align="center">
		<%= oitem.FItemList(i).Fisusing %>
	</td>	
</tr>
<% next %>
<tr>
	<td colspan="13" align="center" bgcolor="#FFFFFF">
		<!-- 페이징처리 -->
	 	<% if oitem.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oitem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</form>
<% end if %>
</table>

<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="300"></iframe>

<%
 set oitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
