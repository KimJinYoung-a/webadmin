<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 핑거스 계약 관리
' Hieditor : 2016.08.10 서동석 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/classes/partners/fingersUpcheAgreeCls.asp"-->
<%
dim page, makerid,groupid, ctrtype, chkdelinc
dim arect, reqCtrSearch, agreeState
dim i

page    = requestCheckVar(request("page"),10)
makerid = requestCheckVar(request("makerid"),32)
groupid = requestCheckVar(request("groupid"),10)
ctrtype  = requestCheckVar(request("ctrtype"),10)
arect   = requestCheckVar(request("arect"),32)
agreeState = requestCheckVar(request("agreeState"),10)
reqCtrSearch = requestCheckVar(request("reqCtrSearch"),10)
chkdelinc = requestCheckVar(request("chkdelinc"),10)  

if (page="") then page=1
dim ocontract
set ocontract = new CFingersUpcheAgree
ocontract.FPageSize = 30
ocontract.FCurrPage =page
ocontract.FRectContractType = ctrtype 
ocontract.FRectagreeState = agreeState
ocontract.FRectMakerid  = makerid
ocontract.FRectGroupID  = groupid
ocontract.FRectDelInclude = chkdelinc
ocontract.GetFingersUpcheAgreeHistList


%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function dnPdfFingers(iUri,ctrKey){
    var popwin = window.open(iUri,'dnPdf'+ctrKey,'width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popViewFingerUpcheAgree(agreeIdx){
    var popwin = window.open('fingersAgreeView.asp?agreeIdx='+agreeIdx,'fingersAgreeView','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function RegFingersAgreeProtoType(){
    var popwin = window.open('/admin/member/contract/ctrPrototypeReg.asp','contractPrototypeReg','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function regFingersContract(makerid,groupid){
    var popwin = window.open('ctrReg_fingers.asp?makerid=' + makerid + '&groupid=' + groupid,'contractFingersReg','width=1124,height=768,scrollbars=yes,resizable=yes');
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

function jssetdate(ctrdtype){
	if (ctrdtype==1){
		document.frm.selyyyy.disabled=false;
		document.frm.selnum.disabled=false;
	}else{
		document.frm.selyyyy.disabled=true;
		document.frm.selnum.disabled=true;
	}
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
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드ID :
    		<% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
        &nbsp;
		&nbsp;
        그룹코드 : <input type="text" name="groupid" value="<%= groupid %>" Maxlength="32" size="16">
	    &nbsp;&nbsp;
		<br>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	    
		
		계약서 구분
    	<% call DrawfingerAgreeMasterCombo("ctrtype",ctrtype) %>
    	&nbsp;&nbsp;
    	계약 진행상태 :
        <% Call DrawAgreeStateCombo("agreestate",agreestate) %>
        &nbsp;&nbsp;
        <input type="checkbox" name="chkdelinc" <%=CHKIIF(chkdelinc="on","checked","")%> >삭제내역포함 검색
	</td>
</tr>

</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:5;padding-bottom:5;">
	<tr>
		<td align="left">
<% if (FALSE) then %>
		    <input type="button" value="신규 계약 등록" onClick="regFingersContract('<%=makerid%>','')" class="button">
<% end if %>
		   
		</td>
		<td align="right">
        	<% if (C_ADMIN_AUTH) then %>
        	<input type="button" value="계약서 원본 등록" onClick="RegFingersAgreeProtoType()" class="button">
        	<% end if %>
		</td>
	</tr>
</table>
<!-- 액션 끝 -->
<form name="frmList" >
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="12" align="right">총 <%= FormatNumber(ocontract.FTotalCount,0) %> 건 <%=page%>/<%=ocontract.FTotalPage%> page</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="100" >번호</td>
    <td width="70"  >계약서명</td>
    <td width="100" >계약서번호</td>
    <td width="70"  >그룹코드</td>
    <td width="80"  >업체명</td>
    <td width="120"  >브랜드ID</td>
    <td width="80" >등록일</td>
    <td width="80" >조회일</td>
    <td width="80" >동의일</td>
    <td width="100" >상태</td>
    <td width="100"  >계약서보기</td>
    <td  >다운로드</td>
</tr>
<% for i=0 to ocontract.FResultCount - 1 %>
<tr bgcolor="<%=CHKIIF(isNULL(ocontract.FITemList(i).Fdeldate),"#FFFFFF","#CCCCCC")%>">
    <td><%= ocontract.FITemList(i).FagreeIdx %></td>
    <td><%= ocontract.FITemList(i).FcontractName %></td>
    <td align="center"><%= ocontract.FITemList(i).FContractNo %></td>
    <td align="center"><%= ocontract.FITemList(i).Fgroupid %></td>
    <td><%= ocontract.FITemList(i).Fcompanyname %></td>
    <td><%= ocontract.FITemList(i).Fmakerid %></td>
    <td><%= ocontract.FITemList(i).Fregdate %></td>
    <td><%= ocontract.FITemList(i).Fviewdate %></td>
    <td><%= ocontract.FITemList(i).Fagreedate %></td>
    <td align="center"><%= ocontract.FITemList(i).getAgreeStateName %></td>
    
    <td align="center"><img src="/images/iexplorer.gif" style="cursor:pointer" onClick="popViewFingerUpcheAgree('<%=ocontract.FITemList(i).FagreeIdx %>');"></td>
    <td align="center">
        <% if ocontract.FITemList(i).IsAgreeFinished then %>
        <img src="/images/pdficon.gif" style="cursor:pointer" onClick="dnPdfFingers('<%=ocontract.FITemList(i).getPdfDownLinkUrlAdm %>','<%=ocontract.FITemList(i).FagreeIdx %>');">
        <% end if %>
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="12" align="center">
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

</table>
</form>
<%
	set ocontract = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->