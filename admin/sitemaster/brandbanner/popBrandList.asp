<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/brand_banner_manageCls.asp"-->
<%
dim menupos, imenuposStr, imenuposnotice, imenuposhelp
menupos = request("menupos")
if menupos ="" then menupos=1

imenuposStr = fnGetMenuPos(menupos, imenuposnotice, imenuposhelp)

'// 즐겨찾기
dim IsMenuFavoriteAdded

IsMenuFavoriteAdded = fnGetMenuFavoriteAdded(session("ssBctID"), menupos)


Dim i, cQuick, vPage, vDateGubun, vSDate, vEDate, vEndType, vUseYN, vSearchGubun, vSearchTxt, vidx
vPage = NullFillWith(requestCheckVar(Request("page"),10),1)
vidx = requestCheckVar(Request("idx"),10)
vSearchGubun = requestCheckVar(Request("SearchGubun"),10)
vSearchTxt = requestCheckVar(Request("searchtxt"),50)

Set cQuick = New CSearchMng
cQuick.FCurrPage = vPage
cQuick.FPageSize = 15
cQuick.FRectMasterIDX = vidx
cQuick.FRectSearchGubun = vSearchGubun
cQuick.FRectSearchTxt = vSearchTxt
cQuick.fnQuickLinkBrandList

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<script language='javascript'>
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function searchFrm(p){
	frm1.page.value = p;
	frm1.submit();
}

//브랜드 ID 검색 팝업창
function jsBrandBannerSearchBrandID(){
    var popwin2 = window.open("popBrandSearch.asp?idx=<%=vidx%>","popBrandSearch2","width=800 height=400 scrollbars=yes resizable=yes");
	popwin2.focus();
}

//브랜드 ID 등록 팝업창
function fnBrandBulkADD(){
    var popwin3 = window.open("popBrandBulkAdd.asp?idx=<%=vidx%>","popBrandSearch3","width=800 height=400 scrollbars=yes resizable=yes");
	popwin3.focus();
}

function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function fnDelBrand(){
	var chk=0;
	$("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("브랜드를 선택해주세요.");
		return;
	}
	if(confirm("선택한 브랜드를 삭제 하시겠습니까?")) {
		document.frmList.mode.value="del";
		document.frmList.action="addbrandproc.asp";
		document.frmList.submit();
	}
}

function fnreload(){
	window.location.reload();
}
</script>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<div class="contSectFix scrl">
	<!-- 상단 검색폼 시작 -->
	<form name="frm1" method="get" action="">
	<input type="hidden" name="page" value="<%=vPage%>">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
    <input type="hidden" name="idx" value="<%=vidx%>">
	<!-- search -->
	<div class="searchWrap">
		<div class="search">
			<ul>
				<li>
					<p class="formTit">검색 :</p>
					<select class="formSlt" id="SearchGubun" name="SearchGubun" title="옵션 선택">
                        <option value="" <%=CHKIIF(vSearchGubun="","selected","")%>>전체</option>
						<option value="brandid" <%=CHKIIF(vSearchGubun="brandid","selected","")%>>브랜드ID</option>
						<option value="socname_kr" <%=CHKIIF(vSearchGubun="socname_kr","selected","")%>>스트리트명(한글)</option>
						<option value="company_name" <%=CHKIIF(vSearchGubun="company_name","selected","")%>>회사명</option>
					</select>
                    <input type="text" class="formTxt" id="searchtxt" name="searchtxt" value="<%=vSearchTxt%>" style="width:200px" placeholder="" />
				</li>
			</ul>
		</div>
		<input type="button" class="schBtn" value="검색" onClick="searchFrm(1);" />
	</div>
	<!-- //search -->
	</form>

	<div class="cont">
		<div class="pad20">
			<div class="overHidden">
				<div class="ftLt">
					<input type="button" class="btn" value="브랜드 일괄 등록" onClick="fnBrandBulkADD();" />
					<input type="button" class="btn" value="브랜드 검색/추가" onClick="jsBrandBannerSearchBrandID();" />
					<input type="button" class="btn" value="선택 브랜드 삭제" onClick="fnDelBrand();" />
				</div>
			</div>
			<div>
				<div class="rt pad10">
					<span>검색결과 : <strong><%=FormatNumber(cQuick.FTotalCount,0)%></strong></span> <span class="lMar10">페이지 : <strong><%=cQuick.FtotalPage%> / <%=FormatNumber(vPage,0)%></strong></span>
				</div>
				<form name="frmList" method="post">
				<input type="hidden" name="mode">
				<input type="hidden" name="idx" value="<%=vidx%>">
				<table class="tbType1 listTb">
					<thead>
					<tr>
						<th><div><input type="checkbox" onclick="chkAllItem();"/></div></th>
						<th><div>브랜드ID</div></th>
						<th><div>스트리트명(한글) / 스트리트명(영문)</div></th>
						<th><div>회사명</div></th>
						<th><div>작성일</div></th>
					</tr>
					</thead>
					<tbody>
					<%
						If cQuick.FResultCount > 0 Then
							For i=0 To cQuick.FResultCount-1
					%>
							<tr>
								<td><input type="checkbox" name="chkIdx" id="chkIdx" value="<%=cQuick.FItemList(i).Fidx%>" /></td>
								<td class="lt"><%=cQuick.FItemList(i).Fbrandid%></td>
                                <td><%=cQuick.FItemList(i).Fsocname_kor%> / <%=cQuick.FItemList(i).Fsocname%></td>
                                <td><%=cQuick.FItemList(i).Fcompany_name%></td>
								<td><%=Left(cQuick.FItemList(i).Fregdate, 10)%></td>
							</tr>
					<%
							Next
						End If
					%>
					</tfoot>
				</table>
				</form>
				<div class="ct tPad20 cBk1">
					<% if cQuick.HasPreScroll then %>
					<a href="javascript:searchFrm('<%= cQuick.StartScrollPage-1 %>')">[pre]</a>
					<% else %>
		    			[pre]
		    		<% end if %>
		    		
		    		<% for i=0 + cQuick.StartScrollPage to cQuick.FScrollCount + cQuick.StartScrollPage - 1 %>
		    			<% if i>cQuick.FTotalpage then Exit for %>
		    			<% if CStr(vPage)=CStr(i) then %>
		    			<span class="cRd1">[<%= i %>]</span>
		    			<% else %>
		    			<a href="javascript:searchFrm('<%= i %>')">[<%= i %>]</a>
		    			<% end if %>
		    		<% next %>
					
					<% if cQuick.HasNextScroll then %>
		    			<a href="javascript:searchFrm('<%= i %>')">[next]</a>
		    		<% else %>
		    			[next]
		    		<% end if %>
				</div>
			</div>
		</div>
	</div>
</div>
<% Set cQuick = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->