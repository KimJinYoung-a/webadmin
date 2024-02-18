<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스
' History : 2017.06.05 원승현 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->

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
	    olecture.FPageSize=10000
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


	Response.Expires=0
	response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
	Response.CacheControl = "public"
%>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" border="1">
<% if olecture.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="12">
		검색결과 : <b><%= olecture.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= olecture.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">강좌코드</td>
	<td align="center">옵션코드</td>
	<td align="center">강좌명</td>
	<td align="center">강사ID</td>
	<td align="center">강사명</td>
	<% If weclass <> "Y" Then %>
	<td align="center">강좌(시작)일</td>
	<% End If %>
	<td align="center">수강료<br>재료비</td>
	<td align="center">매입가</td>
	<td align="center">마진</td>
	<td align="center">재료비<br>포함결제</td>
	<td align="center">정원</td>
	<td align="center">신청인원(웹상)</td>
</tr>
<%
Dim couponsellcash, couponbuycash
for i=0 to olecture.FResultCount - 1
%>
<% if olecture.FItemList(i).FIsUsing="N" then %>
<tr align="center">
<% else %>
<tr align="center">
<% end if %>
	<td><%= olecture.FItemList(i).Fidx %></td>
	<td><%= olecture.FItemList(i).FlecOption %></td>
	<td><%= olecture.FItemList(i).Flec_title %>
	</td>
	<td><%= olecture.FItemList(i).Flecturer_id %></td>
	<td><%= olecture.FItemList(i).Flecturer_name %></td>
	<% If weclass <> "Y" Then %>
	<td><%= olecture.FItemList(i).Flec_startday1 %></td>
	<% End If %>
	<td align="right">
		<%
		Response.Write FormatNumber(olecture.FItemList(i).Flec_cost,0)
		'쿠폰가
		if olecture.FItemList(i).FlecturerCouponYn="Y" then
			Select Case olecture.FItemList(i).FlecturerCouponType
				Case "1"
				    couponsellcash = olecture.FItemList(i).Flec_cost*((100-olecture.FItemList(i).FlecturerCouponValue)/100)
					Response.Write "<br><font color=#5080F0> " & FormatNumber(couponsellcash,0) & ""
					Response.Write "<br>-"&olecture.FItemList(i).FlecturerCouponValue&"%</font>"
				Case "2"
				    couponsellcash = olecture.FItemList(i).Flec_cost-olecture.FItemList(i).FlecturerCouponValue
					Response.Write "<br><font color=#5080F0> " & FormatNumber(couponsellcash,0) & ""
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
	<% elseif  (olecture.FItemList(i).Fmatinclude_yn="N") and (olecture.FItemList(i).Fmat_cost>0) then %>
	    <strong><%= FormatNumber(olecture.FItemList(i).Flec_cost,0) %></strong>
	<% elseif (olecture.FItemList(i).Fmatinclude_yn="X") then %>
	    <strong><%= FormatNumber(olecture.FItemList(i).Flec_cost,0) %></strong>
	<% end if %>
	</td>
	<td><%= olecture.FItemList(i).Flimit_count %></td>
	<td><%= olecture.FItemList(i).Flimit_sold %></td>
</tr>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
<%
	set olecture = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->