<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 상품 히스토리
' History : 2020.08.17 이상구 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemHistoryCls.asp"-->
<!-- #include virtual="/admin/etc/only_sys/only_sys_cls.asp"-->
<%
dim itemid, i, oitemHostory, clist, iTotCnt, vS_PKIdx_txt
	itemid = getNumeric(requestCheckVar(request("itemid"),16))

set oitemHostory = new CItemHistory
	oitemHostory.FRectItemID = itemid
    oitemHostory.FPageSize = 30
	if itemid<>"" then
		oitemHostory.getItemHistoryList
	end if

vS_PKIdx_txt = itemid
Set clist = New cOnlySys
	clist.FCurrPage = 1
	clist.FPageSize = 5
	clist.FGubun = "item"
	clist.FEvtSDate = Left(DateAdd("yyyy", -1, Now()), 10)
	clist.FPK_Idx = "itemid"
	clist.FPK_Idx_txt = itemid
    if itemid<>"" then
	    clist.fnSCMChangeList
	    iTotCnt = clist.ftotalcount
    end if

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>

<script type="text/javascript">

function jsViewContents(i){
	if($("#span_contents"+i+"").is(":hidden")){
		$("#span_contents"+i+"").show();
	}else{
		$("#span_contents"+i+"").hide();
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a" >
		<tr>
			<td>판매정보 변경 히스토리</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>등록일</td>
	<td>판매상태</td>
	<td>판매가</td>
	<td>매입가</td>
	<td>매입구분</td>
	<td>배송구분</td>
	<td>상품쿠폰</td>
	<td>브랜드명</td>
	<td>한정판매여부</td>
</tr>
<% If oitemHostory.FResultCount > 0 Then %>
	<% for i=0 to oitemHostory.FResultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= FormatDate(oitemHostory.FItemList(i).Fregdate,"0000-00-00 00:00:00") %></td>
		<td><%= getSellYnName(oitemHostory.FItemList(i).Fsellyn) %></td>
		<td><%= FormatNumber(oitemHostory.FItemList(i).Fsellcash,0) %></td>
		<td><%= FormatNumber(oitemHostory.FItemList(i).Fbuycash,0) %></td>
		<td><%= mwdivName(oitemHostory.FItemList(i).Fmwdiv) %></td>
		<td><%= GetdeliverytypeName(oitemHostory.FItemList(i).Fdeliverytype) %></td>
		<td><%= oitemHostory.FItemList(i).Fcitemcouponidx %></td>
		<td><%= oitemHostory.FItemList(i).Fbrandname %></td>
		<td>
			<%= oitemHostory.FItemList(i).Flimityn %>
			<% if oitemHostory.FItemList(i).Flimityn="Y" then %>
				&nbsp;(<%= oitemHostory.FItemList(i).FLimitNo %>-<%= oitemHostory.FItemList(i).FLimitSold %>=<%= oitemHostory.FItemList(i).FLimitNo-oitemHostory.FItemList(i).FLimitSold %>)
			<% end if %>
		</td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="center">
			[검색결과가 없습니다.]
		</td>
	</tr>
<% End If %>
</table>

<br />

<table border="0" width="100%" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a" >
		<tr>
			<td>상품정보 변경 히스토리</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50" nowrap>idx</td>
  	<td width="180" nowrap>등록일</td>
  	<td width="150" nowrap>작업자</td>
	<td width="110" nowrap>구분</td>
	<td width="450" nowrap>작업메뉴</td>
  	<td width="110" nowrap>접근IP</td>
</tr>
<%
If clist.FResultCount > 0 Then
	For i=0 To clist.FResultCount-1
%>
	<tr height="25" bgcolor="#FFFFFF" onClick="jsViewContents('<%=clist.FItemList(i).Fidx%>');" style="cursor:pointer;" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F0F0F0'">
		<td width="50" align="center"><%=clist.FItemList(i).Fidx%></td>
		<td width="180" align="center"><%=clist.FItemList(i).Fregdate%></td>
		<td width="150" align="center"><%=clist.FItemList(i).Fusername%>(<%=clist.FItemList(i).Fuserid%>)</td>
		<td width="110" align="center"><%=clist.FItemList(i).Fgubun%></td>
		<td width="450" style="padding-left:5px;">
			<%=clist.FItemList(i).Fmenuname%>(menupos:<%=clist.FItemList(i).Fmenupos%>)
			&nbsp;
			<% If clist.FItemList(i).Fmenupos <> "0" Then %>
			[<a href="<%=clist.FItemList(i).Fmenulink%>" target="_blank">링 크</a>]
			<% End If %>
		</td>
		<td width="110" align="center"><%=clist.FItemList(i).Frefip%></td>
	</tr>
	<tr>
		<td colspan="6" width="1150" bgcolor="#FFFFFF" id="td_contents" style="word-break:break-all;">
			<span class="dummy1" id="span_contents<%=clist.FItemList(i).Fidx%>" style="display:none;">
			<% If clist.FItemList(i).Fgubun = "item" OR clist.FItemList(i).Fgubun = "dispcate" Then %>
				<% If Len(clist.FItemList(i).Fpk_idx) < 4 Then %>
				<% Else %>
				- 상품코드 : <%=clist.FItemList(i).Fpk_idx%> [<a href="http://www.10x10.co.kr/<%=clist.FItemList(i).Fpk_idx%>" target="_blank">링 크</a>]<br />
				<% End If %>
			<% End If %>
			<%=Replace(clist.FItemList(i).Fcontents,vbCrLf,"<br />")%>
			</span>
		</td>
	</tr>
<% Next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="8" align="center">
			[검색결과가 없습니다.]
		</td>
	</tr>
<% End If %>
</table>

<%
set oitemHostory = nothing
set clist = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
