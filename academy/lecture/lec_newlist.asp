<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스
' History : 2009.04.07 서동석 생성
'			2010.05.12 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%

dim yyyy1,mm1,nowdate , yyyy2,mm2,dd2 , lecturer , lec_idx, lec_title, lecturdate , lecturdateyn
dim page , waitlec ,CateCD1, CateCD2, CateCD3 ,i ,weclass , classlevel, lec_gubun
dim limitsoldnotZero

Dim code_large , code_mid

	code_large = RequestCheckvar(request("code_large"),3)
	code_mid = RequestCheckvar(request("code_mid"),3)

	lec_idx = RequestCheckvar(request("lec_idx"),10)
	lecturer = RequestCheckvar(request("lecturer"),32)
	lec_title = request("lec_title")
	waitlec = RequestCheckvar(request("waitlec"),10)
	CateCD1 = RequestCheckvar(request("CateCD1"),3)
	'CateCD2 = request("CateCD2")
	CateCD3 = RequestCheckvar(request("CateCD3"),3)
	lecturdateyn = RequestCheckvar(request("lecturdateyn"),10)
	yyyy2 = RequestCheckvar(request("yyyy2"),4)
	mm2   = RequestCheckvar(request("mm2"),2)
	dd2   = RequestCheckvar(request("dd2"),2)
	yyyy1 = RequestCheckvar(request("yyyy1"),4)
	mm1   = RequestCheckvar(request("mm1"),2)
	page = RequestCheckvar(request("page"),10)
	weclass = RequestCheckvar(request("weclass"),1)
	classlevel = RequestCheckvar(request("classlevel"),1)
	lec_gubun = RequestCheckvar(request("lec_gubun"),10)
    limitsoldnotZero = RequestCheckvar(request("limitsoldnotZero"),1)
    
	if page="" then page=1

	nowdate = now()

if yyyy1="" then
	yyyy1 = Left(Cstr(nowdate),4)
	mm1	  = Mid(Cstr(nowdate),6,2)
end if

if yyyy2="" then
	yyyy2 = Left(Cstr(nowdate),4)
	mm2	  = Mid(Cstr(nowdate),6,2)
	dd2	  = Mid(Cstr(nowdate),9,2)
end if

lecturdate = yyyy2 + "-" + mm2 + "-" + dd2

dim olecture
set olecture = new CLecture
	olecture.FCurrPage = page
	if (limitsoldnotZero="Y") then
	    olecture.FPageSize=200
	else
	    olecture.FPageSize=20
    end if
    
	if lec_idx<>"" then
		olecture.FRectSearchidx = lec_idx
		olecture.FweclassYN = weclass
	else
		olecture.FRectSearchYYYYMM = yyyy1 + "-" + mm1
		olecture.FRectSearchLecturer = lecturer
		olecture.FRectSearchTitle = lec_title
		olecture.FRectCateCD1 = CateCD1
		'olecture.FRectCateCD2 = CateCD2
		olecture.FRectCateCD3 = CateCD3
		olecture.Fcode_Large = code_large
		olecture.Fcode_Mid = code_mid
		olecture.FweclassYN = weclass
		olecture.Fclasslevel = classlevel
		olecture.Flec_gubun = lec_gubun
        olecture.FRectlimitsoldnotZero= limitsoldnotZero

		if lecturdateyn="on" then
			olecture.FRectSearchLectureDay = lecturdate
		end if
	end if

	if waitlec="on" then
		olecture.GetWaitManageLectureList
	else
		olecture.GetNewLectureList
	end If
%>
<script type='text/javascript'>

function NextPage(page){
	frm.page.value= page;
	frm.submit();
}

function GetOnload(){
	ckEnabled(frm.lecturdateyn);
}

function ckEnabled(comp){
	frm.yyyy2.disabled = (!comp.checked);
	frm.mm2.disabled = (!comp.checked);
	frm.dd2.disabled = (!comp.checked);
}

function popLecSimpleEdit(lec_idx){
	popwin = window.open('/academy/lecture/poplecsimpleedit.asp?lec_idx=' + lec_idx,'popLecSimpleEdit','width=600,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popLecReg(lec_idx,lecOption){
	popwin = window.open('/academy/lecture/poplecreg.asp?lec_idx=' + lec_idx + '&lecOption='+lecOption,'popLecReg','width=800,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function Popimage(lec_idx){
	popwin = window.open('/academy/lecture/lec_image_edit.asp?lec_idx=' + lec_idx,'Popimage','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function Popview(lec_idx){
	popwin = window.open('<%=wwwFingers%>/lecture/lecturedetail.asp?lec_idx=' + lec_idx,'Popview','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopEditWin(lec_idx, hidweclass){
	popwin = window.open('/academy/lecture/lec_edit.asp?lec_idx=' + lec_idx + '&hidweclass=' + hidweclass,'PopEditWin','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopRegWin(lec_idx){
	popwin = window.open('/academy/lecture/lec_reg.asp','popLecEdit','width=720,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function popmap(){
	popwin = window.open('/academy/lecture/lib/pop_lec_mapimg.asp','popMap','width=1024,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popexcel(y1, m1, leci, lect, lectu, classle, lecgu, catecd1, catecd2, catecd3, codelarge, limitsoldnotZero){
	popwin = window.open('/academy/lecture/lec_newlist_excel.asp?yyyy1='+y1+'&mm1='+m1+'&lec_idx='+leci+'&lec_title='+lect+'&lecturer='+lectu+'&classlevel='+classle+'&lec_gubun='+lecgu+'&CateCD1='+catecd1+'&CateCD3='+catecd3+'&code_large='+codelarge+'&limitsoldnotZero='+limitsoldnotZero,'popMap','width=1024,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popwaiting(v,lecOption){
	popwin = window.open('/academy/lecture/wait_user_list2.asp?menupos=835&lec_idx='+ v + '&lecOption='+lecOption,'popwait','width=840,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
function onlyNumberInput() 
{ 
	var code = window.event.keyCode; 
	if ((code > 34 && code < 41) || (code > 47 && code < 58) || (code > 95 && code < 106) || code == 8 || code == 9 || code == 13 || code == 46) { 
		window.event.returnValue = true; 
		return; 
	} 
	window.event.returnValue = false; 
}
window.onload = GetOnload;

</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" >
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<% If weclass <> "Y" Then %>
		검색월 		: <% DrawYMBox yyyy1,mm1 %>&nbsp;
		<% End If %>
		강좌코드	: <input type="text" name="lec_idx" size="8" value="<%= lec_idx %>"  style="IME-MODE: disabled" />&nbsp;
		강좌명 		:	<input type="text" name="lec_title" size="20"  value="<%= lec_title %>">
		강사 			: <% drawSelectBoxLecturer "lecturer",lecturer  %>
		등급			: <select name="classlevel">
								<option value="" <%=chkiif(classlevel="","selected","")%>>선택</option>
								<option value="1" <%=chkiif(classlevel="1","selected","")%>>초급</option>
								<option value="2" <%=chkiif(classlevel="2","selected","")%>>중급</option>
								<option value="3" <%=chkiif(classlevel="3","selected","")%>>고급</option>
						  </select>

		강좌구분			: <select name="lec_gubun">
								<option value="" <%=chkiif(lec_gubun="","selected","")%>>선택</option>
								<option value="0" <%=chkiif(lec_gubun="0","selected","")%>>일반</option>
								<option value="1" <%=chkiif(lec_gubun="1","selected","")%>>단체</option>
							  </select>

		<br><input type="checkbox" name="lecturdateyn" <% if lecturdateyn = "on" then response.write "checked" %> onclick="ckEnabled(this)">
		강좌일 		: <% DrawOneDateBox2 yyyy2,mm2,dd2 %>
		분류 : <%=makeCateSelectBox("CateCD1",CateCD1) & " " & makeCateSelectBox("CateCD3",CateCD3)%>
		대카테고리 :  <% DrawSelectBoxLecCategoryLarge "code_large" ,  code_large  , "N" %>
		<% 
        		if code_large <> "" Then
   					Response.Write "&nbsp;"
					Response.Write "중카테고리 : "
        			call DrawSelectBoxLecCategoryMid("code_mid",code_large, code_mid  , "N")
				End If 
		%>
		<input type="checkbox" name="waitlec" <% if waitlec = "on" then response.write "checked" %> >대기자 관리 필요 강좌
		<input type="checkbox" name="weclass" value="Y" <%=chkiif(weclass="Y","checked","")%> >단체강좌(weclass)
		
		<input type="checkbox" name="limitsoldnotZero" value="Y" <%=chkiif(limitsoldnotZero="Y","checked","")%> >신청인원<>0
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" onclick="popmap();" value="약도이미지" class="button">
			<input type="button" onclick="popexcel('<%=yyyy1%>', '<%=mm1%>', '<%=lec_idx%>', '<%=lec_title%>', '<%=lecturer%>', '<%=classlevel%>', '<%=lec_gubun%>', '<%=CateCD1%>', '', '<%=CateCD3%>', '<%=code_large%>','<%=limitsoldnotZero%>');" value="엑셀출력" class="button">
		</td>
		<td align="right">
			<a href="/academy/lecture/lec_reg.asp<%=chkiif(weclass="Y","?weclass="&weclass&"","")%>" target="_blank"><img src="/images/icon_new_registration.gif" border="0"></a>
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if olecture.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= olecture.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= olecture.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <% if (limitsoldnotZero<>"Y") then %>
	<td align="center"></td>
	<td align="center">이미지</td>
    <% end if %>
	<td align="center">강좌코드<br>옵션코드</td>
	<td align="center">강좌명</td>
	<td align="center">강사명</td>
	<% If weclass <> "Y" Then %>
	<td align="center">강좌(시작)일</td>
	<td align="center">접수기간</td>
	<% End If %>
	<td align="center">수강료<br>재료비</td>
	<td align="center">매입가</td>
	<td align="center">마진</td>
	<td align="center">재료비<br>포함결제</td>
	<td align="center">정원<br>신청인원(웹상)</td>
	<td align="center">대기인원<br>신청내역</td>
	<td align="center">마감<br>여부</td>
	<td align="center">수정</td>
	<td align="center">수강<br>입력</td>
</tr>
<%
Dim couponsellcash, couponbuycash
for i=0 to olecture.FResultCount - 1
%>
<% if olecture.FItemList(i).FIsUsing="N" then %>
<tr align="center" bgcolor="#eeeeee" >
<% else %>
<tr align="center" bgcolor="#FFFFFF" >
<% end if %>
    <% if (limitsoldnotZero<>"Y") then %>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td><a href="#" onclick="Popimage('<%= olecture.FItemList(i).Fidx %>'); return false;"><img src="<%= olecture.FItemList(i).Fsmallimg %>" width="50" height=50 border="0"></a></td>
	<% end if %>
	<td>
		<a href="#" onclick="Popview('<%= olecture.FItemList(i).Fidx %>'); return false;" title="페이지 보기"><%= olecture.FItemList(i).Fidx %></a>
		<br><%= olecture.FItemList(i).FlecOption %>
	</td>
	<td><a href="#" onclick="PopEditWin('<%= olecture.FItemList(i).Fidx %>','<%=weclass%>'); return false;"><%= olecture.FItemList(i).Flec_title %></a>
	<% if (olecture.FItemList(i).FlecOptionName<>"") then %>
	<br><font color="#888888">(<%= olecture.FItemList(i).FlecOptionName %>)</font>
	<% end if%>
	</td>
	<td><%= olecture.FItemList(i).Flecturer_id %><br>(<%= olecture.FItemList(i).Flecturer_name %>)</td>
	<% If weclass <> "Y" Then %>
	<td><%= olecture.FItemList(i).Flec_startday1 %></td>
	<td align="center"><%= olecture.FItemList(i).Freg_startday %><br>~<br><%= olecture.FItemList(i).Freg_endday %></td>
	<% End If %>
	<td align="right">
		<%
		Response.Write FormatNumber(olecture.FItemList(i).Flec_cost,0)
		'쿠폰가
		if olecture.FItemList(i).FlecturerCouponYn="Y" then
			Select Case olecture.FItemList(i).FlecturerCouponType
				Case "1"
				    couponsellcash = olecture.FItemList(i).Flec_cost*((100-olecture.FItemList(i).FlecturerCouponValue)/100)
					Response.Write "<br><font color=#5080F0><img src='http://fiximage.10x10.co.kr/web2008/category/icon_coupon.gif' border=0> " & FormatNumber(couponsellcash,0) & ""
					Response.Write "<br>-"&olecture.FItemList(i).FlecturerCouponValue&"%</font>"
				Case "2"
				    couponsellcash = olecture.FItemList(i).Flec_cost-olecture.FItemList(i).FlecturerCouponValue
					Response.Write "<br><font color=#5080F0><img src='http://fiximage.10x10.co.kr/web2008/category/icon_coupon.gif' border=0> " & FormatNumber(couponsellcash,0) & ""
					Response.Write "<br>-"&olecture.FItemList(i).FlecturerCouponValue&"</font>"
			end Select
		end if
		%>
		<% if (olecture.FItemList(i).Fmatinclude_yn	="C") then %>
		<br><%= FormatNumber(olecture.FItemList(i).Fmat_cost,0) %>
		<% else %>
		<br><font color="#AAAAAA"><%= FormatNumber(olecture.FItemList(i).Fmat_cost,0) %></font>
		<% end if %>
	</td>
	<td align="center">
		<%
		Response.Write FormatNumber(olecture.FItemList(i).fbuying_cost,0)
		'쿠폰가
		if olecture.FItemList(i).FlecturerCouponYn="Y" then
			if olecture.FItemList(i).FlecturerCouponType="1" or olecture.FItemList(i).FlecturerCouponType="2" then
				if olecture.FItemList(i).Fcouponbuyprice=0 or isNull(olecture.FItemList(i).Fcouponbuyprice) then
				   '' couponbuycash = olecture.FItemList(i).Forgsuplycash  ''주석처리2015/09/10
					Response.Write "<br><font color=#5080F0>" & FormatNumber(couponbuycash,0) & "</font>"
				else
				    couponbuycash = olecture.FItemList(i).Fcouponbuyprice
					Response.Write "<br><font color=#5080F0>" & FormatNumber(couponbuycash,0) & "</font>"
				end if
			end if
		end if
		%>
		<% if (olecture.FItemList(i).Fmatinclude_yn	="C") then %>
		<br><%= FormatNumber(olecture.FItemList(i).Fmat_buying_cost,0) %>
		<% else %>
		<br><font color="#AAAAAA"><%= FormatNumber(olecture.FItemList(i).Fmat_buying_cost,0) %></font>
		<% end if %>
	</td>
	<td>
	    <% if (olecture.FItemList(i).Flec_cost<>0) then %>
	    <%= 100-CLng(olecture.FItemList(i).fbuying_cost/olecture.FItemList(i).Flec_cost*100*100)/100 %>%
	    <% end if %>

	    <% if olecture.FItemList(i).FlecturerCouponYn="Y" then %>
	        <%
	        if olecture.FItemList(i).FlecturerCouponType="1" or olecture.FItemList(i).FlecturerCouponType="2" then
				if couponsellcash<>0 then
					response.write "<br><font color=#5080F0>" & 100-CLng(couponbuycash/couponsellcash*100*100)/100 & "%"
				end if
			end if
			%>
	    <% end if %>

	    <% if olecture.FItemList(i).Fmat_cost<>0 then %>
	    <% if (olecture.FItemList(i).Fmatinclude_yn	="C") then %>
		<br><%= 100-CLng(olecture.FItemList(i).Fmat_buying_cost/olecture.FItemList(i).Fmat_cost*100*100)/100 %>%
		<% else %>
		<br><font color="#AAAAAA">-</font>
		<% end if %>
	    <% end if %>

	</td>
	<td>
	<% if (olecture.FItemList(i).Fmatinclude_yn	="C") then %>
	    <strong><%= FormatNumber(olecture.FItemList(i).Flec_cost+olecture.FItemList(i).Fmat_cost,0) %></strong>
	    <br>함께결제
	<% elseif  (olecture.FItemList(i).Fmatinclude_yn="N") and (olecture.FItemList(i).Fmat_cost>0) then %>
	    <strong><%= FormatNumber(olecture.FItemList(i).Flec_cost,0) %></strong>
	    <br><font color="#AAAAAA">현장</font>
	<% elseif (olecture.FItemList(i).Fmatinclude_yn="X") then %>
	    <strong><%= FormatNumber(olecture.FItemList(i).Flec_cost,0) %></strong>
	    <br><font color="#AAAAAA">없음</font>
	<% end if %>
	</td>
	<td>
		<%= olecture.FItemList(i).Flimit_count %>
		<br><%= olecture.FItemList(i).Flimit_sold %>
	</td>
	<td>
		<% if olecture.FItemList(i).WaitOpenRequire then %>
			<b><a href="javascript:popwaiting('<%= olecture.FItemList(i).Fidx %>','<%= olecture.FItemList(i).FlecOption %>')"><font color="red"><%= olecture.FItemList(i).FWaitCount %></font></a></b>
		<% else %>
			<a href="javascript:popwaiting('<%= olecture.FItemList(i).Fidx %>','<%= olecture.FItemList(i).FlecOption %>')"><%= olecture.FItemList(i).FWaitCount %></a>
		<% end if %>
		<br><%= olecture.FItemList(i).FRealJupsuCount %>
		<a href="/academy/lecture/lec_orderlist.asp?searchfield=itemid&itemid=<%= olecture.FItemList(i).Fidx %>&lecOption=<%= olecture.FItemList(i).FlecOption %>" target="_blank">
		<img src="http://webadmin.10x10.co.kr/images/icon_search.jpg" width="16" border="0" align="absbottom"></a>
	</td>
	<td>
		<% if olecture.FItemList(i).Fdisp_yn="N" then %>
		<font color="#3333CC">비전시</font>
		<% end if %>
		<% if olecture.FItemList(i).IsSoldOut then %>
		<font color="#CC3333">마감</font>
		<% end if %>
	</td>
	<td><a href="javascript:popLecSimpleEdit('<%= olecture.FItemList(i).Fidx %>')"><img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a></td>
	<td><a href="javascript:popLecReg('<%= olecture.FItemList(i).Fidx %>','<%= olecture.FItemList(i).FlecOption %>')" ><img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a></td>
</tr>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" height="30" align="center">
	<% if olecture.HasPreScroll then %>
		<a href="javascript:NextPage('<%= olecture.StartScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + olecture.StartScrollPage to olecture.FScrollCount + olecture.StartScrollPage - 1 %>
		<% if i>olecture.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if olecture.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
	set olecture = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->