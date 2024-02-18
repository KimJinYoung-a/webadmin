<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체 계약 관리
' Hieditor : 2013.11.20 서동석 생성
'						2016.08.26 정윤정 수정 - 계약종료기능 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<%
dim jMonth : jMonth=3
if (application("Svr_Info")	= "Dev") then
    jMonth=60
end if

dim page, makerid, catecode ,arect, crect, mrect, contractNo, ContractState , regScope, dispCate
dim reqCtrSearch, reqCtr, grpType, notboru, subtype, ctrType, onusing, offusing, ctryyyy,ctrnum,ctrdtype,ctrmm
dim i
dim ctrenddate, chkend
dim strParm 
	page    = requestCheckVar(request("page"),10)
	makerid = requestCheckVar(request("makerid"),32)
	catecode= requestCheckVar(request("catecode"),10)
	arect   = requestCheckVar(request("arect"),32)
	crect   = requestCheckVar(request("crect"),32)
	mrect   = requestCheckVar(request("mrect"),32)
	contractNo  = requestCheckVar(request("contractNo"),20)
	ContractState = requestCheckVar(request("ContractState"),10)
	regScope = requestCheckVar(request("regScope"),10)
	dispCate = requestCheckvar(request("dispCate"),10)
	catecode = requestCheckvar(request("catecode"),10)
    reqCtrSearch = requestCheckvar(request("reqCtrSearch"),10)
    reqCtr = requestCheckvar(request("reqCtr"),10)
    grpType = requestCheckvar(request("grpType"),10)
    notboru = requestCheckvar(request("notboru"),10)
    subtype = requestCheckvar(request("subtype"),10)
	ctrType = requestCheckvar(request("ctrType"),10)
  onusing = requestCheckvar(request("selonyn"),1)
  offusing = requestCheckvar(request("seloffyn"),1)
 	chkend		=requestCheckvar(request("chkend"),1)
  
  ctryyyy = year(date())
  ctrmm = month(date())  
  
  if ctrmm >=1 and ctrmm<=3 then
  	ctrenddate = ctryyyy&"-04-01"
  elseif 	ctrmm >=4 and ctrmm<=6 then
  	ctrenddate = ctryyyy&"-07-01"
  elseif 	ctrmm >=7 and ctrmm<=9 then
  	ctrenddate = ctryyyy&"-10-01"	
 	elseif 	ctrmm >=10 and ctrmm<=12 then
  	ctrenddate = LEFT(dateadd("yyyy",1,now()),4)&"-01-01"   ''2016/10/04 dateadd(y,ctryyyy,1)&"-01-01" => LEFT(dateadd("y",1,now()),4)&"-01-01"  ??
  end if 
  
	if (page="") then page=1
  if (reqCtrSearch="") then reqCtrSearch="P"
  if chkend ="" then chkend = "0" 
  
   
dim ocontract
set ocontract = new CPartnerContract
	ocontract.FPageSize=100
	ocontract.FCurrPage = page
	ocontract.FRectCateCode = catecode
	ocontract.FRectMakerid = makerid
	ocontract.FRectCompanyName = arect
	ocontract.FRectGroupID = crect
	ocontract.FRectContractno  = contractNo
	ocontract.FRectContractState = ContractState
	ocontract.FRectRegScope = regScope
	ocontract.FRectDispCateCode = dispCate
	ocontract.FRectCateCode = catecode
  ocontract.FRectGrpType  = grpType
  ocontract.FRectOnUsing  = onusing
  ocontract.FRectoffusing = offusing
  ocontract.FRectctrenddate = ctrenddate 
  ocontract.FRectchkend		= chkend
  
	if (reqCtrSearch="N") and (reqCtr<>"") then
	    ocontract.FPageSize=500
	    ocontract.FRectNotIncboru = notboru
	    ocontract.GetNewContractListReq reqCtr,jMonth
	else
	    ocontract.FsubType = subtype
		ocontract.FRectContractType = ctrType
	    ocontract.GetNewContractList
    end if

dim uniqGroupID

if (ocontract.FResultCount>0) then
    if (makerid<>"") or (crect<>"") then
        uniqGroupID = ocontract.FItemList(0).Fgroupid
    end if
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="contract.js"></script>
<script language='javascript'>
$(document).on('change','input[name="reqCtrSearch"]',function() {
    if (this.value=="P"){
        document.frm.reqCtr.disabled=true;
        document.frm.ContractState.disabled=false;
        document.frm.arect.disabled=false;
        document.frm.contractNo.disabled=false;
        $("#dvBoru").hide();
    }else{
        document.frm.reqCtr.disabled=false;
        document.frm.ContractState.disabled=true;
        document.frm.arect.disabled=true;
        document.frm.contractNo.disabled=true;

        $("#dvBoru").show();
    }

});

function RegContractProtoType(){
    var popwin = window.open('ctrPrototypeReg.asp','contractPrototypeReg','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function regContract(makerid,groupid){
    var popwin = window.open('ctrReg.asp?makerid=' + makerid + '&groupid=' + groupid,'contractReg','width=1124,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function NextPage(page){
	frm.page.value = page;
	frm.submit();
}

function modiContract(ctrkey){
    var popwin = window.open('editContract.asp?ctrkey=' + ctrkey,'editContract','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}



//전체 선택
function jsChkAll(){
var frm;
frm = document.frmList;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkid) !="undefined"){
	   	   if(!frm.chkid.length){
	   	   	if(frm.chkid.disabled==false){
		   	 	frm.chkid.checked = true;
		   	}
		   }else{
				for(i=0;i<frm.chkid.length;i++){
					 	if(frm.chkid[i].disabled==false){
					frm.chkid[i].checked = true;
				}
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkid) !="undefined"){
	  	if(!frm.chkid.length){
	   	 	frm.chkid.checked = false;
	   	}else{
			for(i=0;i<frm.chkid.length;i++){
				frm.chkid[i].checked = false;
			}
		}
	  }

	}

} 

//선택계약종료
function endContract(){
	   if (confirm("선택하신 계약서를 종료하시겠습니까?")){
     document.frmList.mode.value ="ctrend";
    	document.frmList.submit();
    }
}


function jsSetEcState(){
	$("#btnSubmit").prop("disabled", true);
	document.frmEcState.submit();
}

function ChangeContractType(frm) {
	return;
}
</script>
<form name="frmEcState" method="post"  action="ctrEcStateUpdate.asp">
<input type="hidden" name="mode" value="ecstate">	
</form>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드포함 :
    		<% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
        &nbsp;
		&nbsp;

		계약 등록자:
		<select name="regScope">
		<option value="">전체
		<option value="R" <%=CHKIIF(regScope="R","selected","")%> >내가등록한계약
		<option value="R" <%=CHKIIF(regScope="S","selected","")%> >내가발송한계약
		<option value="R" <%=CHKIIF(regScope="F","selected","")%> >내가완료한계약
		</select>

        &nbsp;
		&nbsp;
		<span style="white-space:nowrap;">관리카테고리 : <% CALL SelectBoxBrandCategory("catecode", catecode) %></span>
        &nbsp;
		<span style="white-space:nowrap;">전시카테고리 : <% CALL DrawSelectBoxDispCateLarge("dispCate",dispCate,"")%></span>
		<br>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	    회사명/사업자번호 : <input type="text" name="arect" value="<%= arect %>" Maxlength="32" size="16" <%=CHKIIF(reqCtrSearch="N","disabled","") %>>
		&nbsp;&nbsp;
	    그룹코드 : <input type="text" name="crect" value="<%= crect %>" Maxlength="32" size="16">
	    &nbsp;&nbsp;
		계약서번호 : <input type="text" name="contractNo" value="<%= contractNo %>" Maxlength="32" size="16" <%=CHKIIF(reqCtrSearch="N","disabled","") %>>

		<% if (uniqGroupID<>"") and (reqCtrSearch<>"N") then %>
		&nbsp;&nbsp;
		<input type="radio" name="grpType" value="" <%=CHKIIF(grpType="","checked","")%> >계약서별 보기 &nbsp;<input type="radio" name="grpType" value="M" <%=CHKIIF(grpType="M","checked","")%> >판매처/계약형태별 펼쳐보기
		<% end if %>
		&nbsp;&nbsp;
		계약서 구분 :
    	<% drawSubTypeGubun "subtype" , subtype %>
		&nbsp;&nbsp;
		계약서명 :
    	<%  drawSelectBoxContractTypeWithChangeEvent "ctrType" , ctrType %>
	</td>
</tr>

<tr align="center" bgcolor="#FFFFFF" >
    <td align="left" height="30">
    계약 진행상태 :
    <input type="radio" name="reqCtrSearch" id="reqCtrSearch1" value="P" <%=CHKIIF(reqCtrSearch="P","checked","") %> ><label for="reqCtrSearch1">계약진행중</label>
	<select name="ContractState" <%=CHKIIF(reqCtrSearch<>"P","disabled","") %>>
	<option value="">전체</option>
	<option value="M" <% if ContractState="M" then response.write "selected" %> >미완료전체</option>
	<option value="0" <% if ContractState="0" then response.write "selected" %> >수정중(미전송)</option>
	<option value="1" <% if ContractState="1" then response.write "selected" %> >계약오픈(검토대기)</option>
	<option value="2" <% if ContractState="2" then response.write "selected" %> >계약반려(검토반려)</option>
	<option value="3" <% if ContractState="3" then response.write "selected" %> >계약확인(결재완료)</option>
	<option value="6" <% if ContractState="6" then response.write "selected" %> >서명진행</option>
	<option value="7" <% if ContractState="7" then response.write "selected" %> >계약완료</option>
	<option value="8" <% if ContractState="8" then response.write "selected" %> >계약파기요청</option>
	<option value="9" <% if ContractState="9" then response.write "selected" %> >계약종료</option>
	<option value="-1" <% if ContractState="-1" then response.write "selected" %> >삭제</option>
	<option value="-2" <% if ContractState="-2" then response.write "selected" %> >등록오류</option>
	</select>
	&nbsp;

	&nbsp;

    <input type="radio" name="reqCtrSearch" id="reqCtrSearch2" value="N" <%=CHKIIF(reqCtrSearch="N","checked","") %> ><label for="reqCtrSearch2">미계약</label>
    <select name="reqCtr" <%=CHKIIF(reqCtrSearch<>"N","disabled","") %> >
    <option value="OJ" <% if reqCtr="OJ" then response.write "selected" %> >온라인 <%=jMonth%>개월 정산기준,판매상품>0
    <option value="OT" <% if reqCtr="OT" then response.write "selected" %> >온라인 <%=jMonth%>개월 정산기준,오프라인정산없음,판매상품>0
    <option value="OJN" <% if reqCtr="OJN" then response.write "selected" %> >온라인 <%=jMonth%>개월 정산기준,정산액0, 판매상품>0
    <option value="OJNN" <% if reqCtr="OJNN" then response.write "selected" %> >온라인 <%=jMonth%>개월 정산기준,정산액0, 판매상품=0

    <option value="FJ" <% if reqCtr="FJ" then response.write "selected" %> >오프라인 <%=jMonth%>개월 정산기준,판매상품>0
    <option value="FN" <% if reqCtr="FN" then response.write "selected" %> >오프라인 <%=jMonth%>개월 정산기준,온라인정산없음,판매상품>0
    </select>

    <span id="dvBoru" style="display:<%=CHKIIF(reqCtrSearch="N","","none") %>"><input type="checkbox" name="notboru" <%=CHKIIF(notboru="on","checked","")%> >보류브랜드 표시안함</span>
    <% if reqCtrSearch="N" then %>
    (최대 <%=ocontract.FPageSize%>건 검색됨)
    <% end if %>
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left" height="30">    	
    	<input type="checkbox" name="chkEnd" value="1" <% if chkEnd ="1" then%>checked<%end if%>> <%=ctrenddate%>  자동갱신 리스트
    <!--	계약일:
    	<input type="radio" name="rdoCD" value="0"  <%=CHKIIF(ctrdtype="0","checked","") %> onclick="jssetdate(this.value);">전체
    	<input type="radio" name="rdoCD" value="1" <%=CHKIIF(ctrdtype="1","checked","") %>  onclick="jssetdate(this.value);">
    	<select name="selyyyy" class="select" <%if ctrdtype="0" then%> disabled<%end if%>> 
    		<%for i=year(date()) to "2016"  step -1%>
    		<option value="<%=i%>" <%=CHKIIF(ctryyyy=i,"selected","") %>><%=i%></option>
    		<%next%>
    	</select> 년 
    	<select name="selnum" class="select"  <%if ctrdtype="0" then%> disabled<%end if%>>  
    		<option value="1" <%=CHKIIF(ctrnum="1","selected","") %>>[1회차] 01월~03월</option> 
    		<option value="2" <%=CHKIIF(ctrnum="2","selected","") %>>[2회차] 04월~06월</option> 
    		<option value="3" <%=CHKIIF(ctrnum="3","selected","") %>>[3회차] 07월~09월</option> 
    		<option value="4" <%=CHKIIF(ctrnum="4","selected","") %>>[4회차] 10월~12월</option> 
    	</select> -->
    		&nbsp;&nbsp;
    	브랜드 사용여부
    	[ON]: <select name="selonyn" class="select">
    		<option value="" <%=CHKIIF(onusing="","selected","") %>>전체</option>
    		<option value="Y" <%=CHKIIF(onusing="Y","selected","") %>>Y</option>
    		<option value="N"<%=CHKIIF(onusing="N","selected","") %> >N</option>
    	</select>
    	[OFF]: <select name="seloffyn" class="select">
    		<option value="" <%=CHKIIF(offusing="","selected","") %>>전체</option>
    		<option value="Y" <%=CHKIIF(offusing="Y","selected","") %>>Y</option>
    		<option value="N"<%=CHKIIF(offusing="N","selected","") %> >N</option>
    	</select>
    </td>
  </tr>	
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:5;padding-bottom:5;">
	<tr>
		<td align="left">
		    <input type="button" value="신규 계약 등록" onClick="regContract('<%=makerid%>','<%=uniqGroupID%>')" class="button">

		    <% if (uniqGroupID<>"") then  %>
		    &nbsp;&nbsp;
		    <%=uniqGroupID%>
		    <% end if %>
		    &nbsp;&nbsp;
		     <!--input type="button" value="선택 계약 종료" onClick="endContract()" class="button"-->
		</td>
		<td align="right">
			<!--<span style="left-margin:10px;"><input type="button" id="btnSubmit" class="button" value="전자계약서 상태Update" onClick="jsSetEcState();"></span>-->
        	<% if (C_ADMIN_AUTH) then %>
        	<input type="button" value="계약서 원본 등록" onClick="RegContractProtoType()" class="button">
        	<% end if %>
		</td>
	</tr>
</table>
<!-- 액션 끝 -->
<form name="frmList" method="post" action="/admin/member/contract/ctrReg_process.asp?<%=strParm%>">
<input type="hidden" name="mode" value=""> 
<input type="hidden" name="ctred" value="<%=ctrenddate%>"> 
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="18" align="right">총 <%= FormatNumber(ocontract.FTotalCount,0) %> 건 <%=page%>/<%=ocontract.FTotalPage%> page</td>
</tr>
<% if (reqCtr="") then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan="2"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
    <td width="100" rowspan="2">계약서 명</td>
    <td width="100" rowspan="2">계약서번호</td>
    <td width="70"  rowspan="2">그룹코드</td>
    <td width="80"  rowspan="2">업체명</td>
    <td width="120" rowspan="2" >브랜드ID</td>
    <td width="100" rowspan="2">판매처</td>
    <td width="100" rowspan="2">마진</td>
    <td width="100" rowspan="2">계약일</td>
    <td width="100" rowspan="2">계약종료일</td>
    <td width="100" rowspan="2" >상태</td>
    <td width="80" rowspan="2">등록자</td>
    <td width="80" rowspan="2" >발송자</td>
    <td width="80" rowspan="2" >완료자</td>
    <td width="80" colspan="2">브랜드사용여부</td>
    <td rowspan="2">계약구분</td>
	<td rowspan="2">비고</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>ON</td>
	<td>OFF</td>
</tr>
<% for i=0 to ocontract.FResultCount - 1 %>
<tr bgcolor="#FFFFFF">
	  <td align="center"><input type="checkbox" name="chkid" value="<%= ocontract.FITemList(i).FctrKey %>"></td>
    <td><a href="javascript:regContract('<%= ocontract.FITemList(i).FMakerid %>','<%= ocontract.FITemList(i).Fgroupid %>');"><%= ocontract.FITemList(i).FContractName %></a></td>
    <td align="center"><a href="javascript:modiContract('<%= ocontract.FITemList(i).FctrKey %>');"><%= CHKIIF(isNULL(ocontract.FITemList(i).FctrNo) or ocontract.FITemList(i).FctrNo="","-",ocontract.FITemList(i).FctrNo) %></a></td>
    <td align="center"><%= ocontract.FITemList(i).FGroupid %>
    <% if (ocontract.FITemList(i).FGroupid<>ocontract.FITemList(i).FcurrGroupid) then %>
    <br><font color=red><%=ocontract.FITemList(i).FcurrGroupid%></font>
    <% end if %>
    </td>
    <td><%= ocontract.FITemList(i).FcompanyName %></td>
    <td><%= ocontract.FITemList(i).FMakerid %></td>
    <td align="center"><%= ocontract.FITemList(i).getMajorSellplaceName %></td>
    <td align="center"><%= ocontract.FITemList(i).getMajorMarginStr %></td>
    <td align="center"><%= ocontract.FITemList(i).FcontractDate %></td>
    <td align="center"><%= ocontract.FITemList(i).FendDate %></td>
    <td align="center"><font color="<%= ocontract.FITemList(i).GetContractStateColor %>" title="<%= ocontract.FITemList(i).GetStateActiondate %>"><%= ocontract.FITemList(i).GetContractStateName %></font></td>
    <td align="center"><span title="<%= ocontract.FITemList(i).FRegDate%>"><%= ocontract.FITemList(i).FRegUserName %></span></td>
    <td align="center"><span title="<%= ocontract.FITemList(i).FSendDate%>"><%= ocontract.FITemList(i).FSendUserName %></span></td>
    <td align="center"><span title="<%= ocontract.FITemList(i).FfinishDate%>"><%= ocontract.FITemList(i).FfinUserName %></span></td>
    <td align="center"><%= fncolor(ocontract.FITemList(i).Fonbrandusing,"yn") %></td>
    <td align="center"><%= fncolor(ocontract.FITemList(i).Foffbrandusing ,"yn")%></td>
	<td align="center"><%=ocontract.FITemList(i).GetSignType%></td>
    <td align="center"><img src="/images/pdficon.gif" style="cursor:pointer" onClick="dnPdfAdm('<%=ocontract.FITemList(i).getPdfDownLinkUrlAdm %>');"></td>
</tr>
<% next %>
</form>
<tr bgcolor="#FFFFFF">
    <td colspan="18" align="center">
        <% if ocontract.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ocontract.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ocontract.StartScrollPage to ocontract.FScrollCount + ocontract.StartScrollPage - 1 %>
			<% if i>ocontract.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ocontract.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
</tr>
<% else %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="70" rowspan="2">그룹코드</td>
    <td width="80" rowspan="2">업체명</td>
    <td width="120" rowspan="2">브랜드ID</td>
    <td width="120" rowspan="2">브랜드등록일</td>
    <td colspan="4" >판매상품수</td>
    <td colspan="4" ><%=JMonth%>개월 정산</td>

    <td rowspan="2">비고</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>매입</td>
    <td>위탁</td>
    <td>업체</td>
    <td>합계</td>

    <td>매입</td>
    <td>위탁</td>
    <td>업체</td>
    <td>합계</td>
</tr>


<% for i=0 to ocontract.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" align="right">
    <td align="center"><a href="javascript:regContract('<%= ocontract.FITemList(i).FMakerid %>','<%= ocontract.FITemList(i).Fgroupid %>');"><%= ocontract.FITemList(i).FGroupid %></a></td>
    <td align="center"><a href="javascript:regContract('<%= ocontract.FITemList(i).FMakerid %>','<%= ocontract.FITemList(i).Fgroupid %>');"><%= ocontract.FITemList(i).Fcompany_name %></a></td>
    <td align="center"><a href="javascript:regContract('<%= ocontract.FITemList(i).FMakerid %>','<%= ocontract.FITemList(i).Fgroupid %>');"><%= ocontract.FITemList(i).Fmakerid %></a></td>

    <td align="center"><%=ocontract.FITemList(i).getBrandActiveMonth%> 개월</td>
    <td><%= FormatNumber(ocontract.FITemList(i).FMsellcnt,0) %></td>
    <td><%= FormatNumber(ocontract.FITemList(i).FWsellcnt,0) %></td>
    <td><%= FormatNumber(ocontract.FITemList(i).FUsellcnt,0) %></td>
    <td><%= FormatNumber(ocontract.FITemList(i).FTTLsellcnt,0) %></td>

    <td><%= FormatNumber(ocontract.FITemList(i).FMjungsanSum,0) %></td>
    <td><%= FormatNumber(ocontract.FITemList(i).FWjungsanSum,0) %></td>
    <td><%= FormatNumber(ocontract.FITemList(i).FUjungsanSum,0) %></td>
    <td><%= FormatNumber(ocontract.FITemList(i).FTTLjungsanSum,0) %></td>
    <td align="center">
        <% if Not isNULL(ocontract.FITemList(i).FHolddate) then %>
        <span title="보류자ID <%=ocontract.FITemList(i).FHoldregID%>&#13;<%=ocontract.FITemList(i).FHolddate%>">보류</span>
        <% end if %>
    </td>
</tr>
<% next %>
<% end if %>
</table>

<%
	set ocontract = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->