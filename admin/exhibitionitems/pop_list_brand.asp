<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/exhibitionitems/lib/classes/exhibitionCls.asp"-->
<%
'###########################################################
' Description :  브랜드 리스트 목록
' History : 2022-12-22 허진원 생성
'###########################################################
dim isusing , prevDate , i
dim page , mastercode , detailcode
dim oExhibition

mastercode = request("mastercode")
detailcode = request("detailcode")
prevDate = request("prevDate")
isusing = request("isusing")
page = request("page")

if page = "" then page = 1

set oExhibition = new ExhibitionCls
    oExhibition.FPageSize = 10
	oExhibition.FCurrPage = page
	oExhibition.FrectMasterCode = mastercode
	oExhibition.FrectDetailCode = detailcode
    oExhibition.FRectSelDate = prevDate
    oExhibition.FRectIsusing = isusing
	oExhibition.getBrandList
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<script type="text/javascript">
    function AddNewMainContents(idx,mastercode){
        var popwin = window.open('/admin/exhibitionitems/pop_reg_brand.asp?idx='+idx+'&mastercode='+mastercode,'popaddbrand','width=800,height=575,scrollbars=yes,resizable=yes');
        popwin.focus();
    }

    // 페이지 이동
	function NextPage(ipage)
	{
		document.frm.page.value= ipage;
		document.frm.submit();
	}
</script>
</head>
<body>
<div class="contSectFix scrl">
    <div class="contHead">
		<div class="locate"><h2>기획전상품관리 &gt; <strong>브랜드 관리</strong></h2></div>
	</div>
	<div class="pad10">
        <div class="tPad15">
            <form name="frm" method="get" action="">
            <input type="hidden" name="page" value="">
            <input type="hidden" name="research" value="on">
            <input type="hidden" name="mastercode" value="<%=masterCode%>">
            <input type="hidden" name="menupos" value="<%= request("menupos") %>">
                <table class="tbType1 listTb">
                    <tr>
                        <td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
                        <td style="text-align:left;">
                            사용구분 :
                            <select name="isusing" class="select">
                                <option value="" <% if isusing="" then response.write "selected" %>>전체
                                <option value="1" <% if isusing="1" then response.write "selected" %>>사용함
                                <option value="0" <% if isusing="0" then response.write "selected" %>>사용안함
                            </select>
                            &nbsp;&nbsp;
                            지정일자 : <input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" style="vertical-align:middle;"/>
                            <script language="javascript">
                                var CAL_Start = new Calendar({
                                    inputField : "prevDate", trigger    : "prevDate_trigger",
                                    onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
                                });
                            </script>
                        </td>
                        <td width="50" bgcolor="<%= adminColor("gray") %>">
                            <input type="submit" class="button_s" value="검색">
                        </td>
                    </tr>
                </table>                
            </form>
        </div>
		<div class="tPad15">
            <table class="tbType1 listTb">
                <tr height="25" bgcolor="FFFFFF">
                    <td style="text-align:left;" colspan="8">
                        <div style="float:left;">
                            검색결과 : <b><%=oExhibition.FtotalCount%></b>&nbsp;페이지 : <b><%= page %> / <%=oExhibition.FtotalPage%></b>
                        </div>
                        <div style="float:right;vertial-align:bottom;">
                            <a href="javascript:AddNewMainContents('0','<%=mastercode%>');"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
                        <div>
                    </td>
                </tr>
                <tr bgcolor="<%= adminColor("tabletop") %>" height="25" >
                    <th width="50">idx</th>
                    <th>브랜드ID</th>
                    <th>브랜드명</th>
                    <th>대표상품</th>
                    <th>배너 노출 시작일</th>
                    <th>배너 노출 종료일</th>
                    <th>사용여부</th>
                    <th>우선순위</th>
                </tr>
                <%
                    for i=0 to oExhibition.FResultCount - 1
                %>
                <tr bgcolor="<%=chkiif((oExhibition.FItemList(i).IsEndDateExpired) or (oExhibition.FItemList(i).FIsusing="0"),"#DDDDDD","#FFFFFF")%>">
                    <td><%= "<a href=""javascript:AddNewMainContents('" & oExhibition.FItemList(i).Fidx & "','"& mastercode &"');"">" & oExhibition.FItemList(i).Fidx & "</a>" %></td>
                    <td><%= oExhibition.FItemList(i).Fmakerid %></td>
                    <td>
                        [<a href="https://www.10x10.co.kr/street/street_brand_sub06.asp?makerid=<%= oExhibition.FItemList(i).Fmakerid %>" target="_blank"><%= oExhibition.FItemList(i).Fsocname %></a>] <%= oExhibition.FItemList(i).FsocnameKor %><br/><br/>
                    </td>
                    <td>
                        <% if oExhibition.FItemList(i).FbannerImage <> "" then %>
                            <img src="<%= oExhibition.FItemList(i).FbannerImage %>" border="0" width="75">
                        <% else %>
                            <a href="<%=wwwUrl & "/" & oExhibition.FItemList(i).FmodelItem%>" title="<%= oExhibition.FItemList(i).FmodelItem %>"><img src="<%= oExhibition.FItemList(i).FmodelImg %>" border="0" width="75" /></a>
                        <% end if %>
                    </td>
                    <td>
                        <%= oExhibition.FItemList(i).FStartDate %><br/><br/>
                    </td>
                    <td>
                        <%= oExhibition.FItemList(i).FEndDate %><br/><br/>
                    </td>
                    <td><%= chkiif(oExhibition.FItemList(i).FIsusing = 1 , "Y" , "N") %></td>
                    <td><%= oExhibition.FItemList(i).FsortNo %></td>
                </tr>
                <% next %>
                <tr bgcolor="#FFFFFF">
                    <td colspan="12" align="center" height="30">
                    <% if oExhibition.HasPreScroll then %>
                        <a href="javascript:NextPage('<%= oExhibition.StartScrollPage-1 %>');">[pre]</a>
                    <% else %>
                        [pre]
                    <% end if %>

                    <% for i=0 + oExhibition.StartScrollPage to oExhibition.FScrollCount + oExhibition.StartScrollPage - 1 %>
                        <% if i>oExhibition.FTotalpage then Exit for %>
                        <% if CStr(page)=CStr(i) then %>
                        <font color="red">[<%= i %>]</font>
                        <% else %>
                        <a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
                        <% end if %>
                    <% next %>

                    <% if oExhibition.HasNextScroll then %>
                        <a href="javascript:NextPage('<%= i %>');">[next]</a>
                    <% else %>
                        [next]
                    <% end if %>
                    </td>
                </tr>
            </table>
		</div>
	</div>
</div>
<%
    SET oExhibition = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->