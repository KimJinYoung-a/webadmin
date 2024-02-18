<%@Language="VBScript" CODEPAGE="65001" %>
<% option explicit %>
<%
Response.CharSet="utf-8" 
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"--> 
<!-- #include virtual="/lib/classes/contribution/contributionCls.asp"--> 
 
<%
dim sdatetype, dstdate,deddate , sdispcate,catekind
Dim  syear,smonth, sday, eday   
dim itotCnt
dim arrcate, intc 
dim clsCMeachulLog 

    sdatetype     = requestCheckvar(request("dategbn"),32)
    sdispcate = requestCheckvar(request("blnDp"),1) 
    catekind =  requestCheckvar(request("rdoCK"),1) 
   	syear     = requestcheckvar(request("sy"),4)
	smonth     = requestcheckvar(request("sM"),2)
    sday     = requestcheckvar(request("sd"),2)
    eday     = requestCheckvar(request("ed"),2)  

 if sdatetype="" then sdatetype="ipkumdate" 
 if  sdispcate ="" then  sdispcate = "N" 
if syear ="" then  syear = Cstr(Year( dateadd("m",-1,date()) ))
if smonth ="" then smonth = Cstr(Month( dateadd("m",-1,date()) ))
if sday ="" then sday =  "01" 
  dstdate = DateSerial(syear, smonth, sday) 
if eday ="" then  eday =  Cstr(Day( dateadd("d",-1,DateSerial(Year(Date()), Month(Date()), 1)) )) 
 deddate = DateSerial(syear, smonth, eday)    
 
 if dateadd("m",1,dstdate) <= deddate then deddate = dateadd("d",dateadd("m",1,dstdate),-1)   
 
 
dim i, tmpJSON, j,m
dim gubunNM
dim buyPer, BCPer, MPPer, MPerPer
dim hidx_buy1,hidx_buy2 ,hidx_buy3
dim hidx_bc1, hidx_bc2
dim hidx_MP1, hidx_MP2
dim hidx_MPPer1, hidx_MPPer2  
dim mper
dim imax
dim catecnt
dim buyVar, buyVarCate, itemVar, odsVar, itemVarcate
 
	set clsCMeachulLog = new CMeachulLog
	clsCMeachulLog.FdateType =sdatetype
	clsCMeachulLog.FstDate =dstdate
	clsCMeachulLog.FedDate =deddate
	clsCMeachulLog.FDispCate =sdispcate
    clsCMeachulLog.Fcatekind =catekind
    
	 clsCMeachulLog.fnGetOrerLogData

    if sdispcate ="Y" then
        iTotCnt = clsCMeachulLog.FCateRow
        arrcate =clsCMeachulLog.fnGetCateList
        if isarray(arrcate) then
         catecnt =ubound(arrcate,2) +1 
        else
        catecnt =0
        end if 
    else
	    iTotCnt = clsCMeachulLog.FTotCnt
    end if

    Response.Expires=0
    response.ContentType = "application/vnd.ms-excel"
     Response.AddHeader "Content-Disposition", "attachment; filename=공헌이익" & Left(dstdate,7)  & ".xls"
     Response.CacheControl = "public"
     Response.Buffer = true    '버퍼사용여부

     
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td colspan="4" rowspan="2" style="text-align:center">구분</td>
    <td colspan="4" style="text-align:center">전체</td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td colspan="3"><%=arrcate(1,intC)%></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td style="text-align:center"> 주문건수</td>
    <td style="text-align:center"> 상품수량</td>
    <td style="text-align:center"> 금액</td>
    <td style="text-align:center"> 율</td> 
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> 상품수량</td>
    <td style="text-align:center"> 금액</td>
    <td style="text-align:center"> 율</td>
    <%   next
    end if%>
</tr>   

<tr bgcolor="#ffffff">
    <td rowspan="13" style="text-align:center">구매총액</td>
    <td colspan="3" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_sum%></td>
    <td><%=clsCMeachulLog.FItem_sum%></td>
    <td><%=clsCMeachulLog.Fbuy_sum%></td>
    <td></td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_Sum(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.Fcatebuy_Sum(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="6" style="text-align:center">10x10</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10%></td>
    <td><%=clsCMeachulLog.FItem_10%></td>
    <td><%=clsCMeachulLog.Fbuy_10%></td>
    <td> </td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.Fcatebuy_10(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="2" style="text-align:center">상품</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10I%></td>
    <td><%=clsCMeachulLog.FItem_10I%></td>
    <td><%=clsCMeachulLog.Fbuy_10I%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.Fcatebuy_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td style="text-align:center">매입</td> 
    <td><%=clsCMeachulLog.FOds_10I%></td>
    <td><%=clsCMeachulLog.FItem_10I%></td>
    <td><%=clsCMeachulLog.Fbuy_10I%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.Fcatebuy_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">수수료</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10C%></td>
    <td><%=clsCMeachulLog.FItem_10C%></td>
    <td><%=clsCMeachulLog.Fbuy_10C%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10C(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.Fcatebuy_10C(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<%for i = 0 to iTotCnt
   if  clsCMeachulLog.Fsitename(i) = "10x10" and ( clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U")  then
%>
<tr bgcolor="#ffffff">
    <td style="text-align:center"><%=getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i))%></td> 
    <td><%=clsCMeachulLog.FOds(i)%></td>
    <td><%=clsCMeachulLog.FItem(i)%></td>
    <td><%=clsCMeachulLog.Fbuy(i)%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%= clsCMeachulLog.FcateItem(i,arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.Fcatebuy(i,arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<% end if
next%> 
<tr bgcolor="#ffffff">
    <td rowspan="6" style="text-align:center">제휴</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_P%></td>
    <td><%=clsCMeachulLog.FItem_P%></td>
    <td><%=clsCMeachulLog.Fbuy_P%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_P(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.Fcatebuy_P(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="2" style="text-align:center">상품</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_PI%></td>
    <td><%=clsCMeachulLog.FItem_PI%></td>
    <td><%=clsCMeachulLog.Fbuy_PI%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.Fcatebuy_PI(arrcate(3,intc)) %></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td style="text-align:center">매입</td> 
    <td><%=clsCMeachulLog.FOds_PI%></td>
    <td><%=clsCMeachulLog.FItem_PI%></td>
    <td><%=clsCMeachulLog.Fbuy_PI%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.Fcatebuy_PI(arrcate(3,intc)) %></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">수수료</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_PC%></td>
    <td><%=clsCMeachulLog.FItem_PC%></td>
    <td><%=clsCMeachulLog.Fbuy_PC%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PC(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.Fcatebuy_PC(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<%for i = 0 to iTotCnt
    if  clsCMeachulLog.Fsitename(i) <> "10x10" and ( clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U")  then
%>
<tr bgcolor="#ffffff">
    <td style="text-align:center"><%=getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i)) %></td> 
    <td><%=clsCMeachulLog.FOds(i)%></td>
    <td><%=clsCMeachulLog.FItem(i)%></td>
    <td><%=clsCMeachulLog.Fbuy(i)%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem(i,arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.Fcatebuy(i,arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
 <% end if
next%> 


<tr bgcolor="#ffffff">
    <td rowspan="13" style="text-align:center">구매총액수익</td>
    <td colspan="3" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_sum%></td>
    <td><%=clsCMeachulLog.FItem_sum%></td>
    <td><%=clsCMeachulLog.FbuyPF_sum%></td>
    <td></td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_Sum(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcatebuyPF_Sum(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="6" style="text-align:center">10x10</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10%></td>
    <td><%=clsCMeachulLog.FItem_10%></td>
    <td><%=clsCMeachulLog.FbuyPF_10%></td>
    <td>
    <% buyPer=0
    if clsCMeachulLog.Fbuy_10 > 0 then 
		buyPer = round((clsCMeachulLog.FbuyPF_10/clsCMeachulLog.Fbuy_10)*100)
	   end if
    %>
    <%=buyPer/100%>
    </td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
             if clsCMeachulLog.Fcatebuy_10(arrcate(3,intc)) > 0 then
               mper = round((clsCMeachulLog.FcatebuyPF_10(arrcate(3,intc))/clsCMeachulLog.Fcatebuy_10(arrcate(3,intc)))*100)
             end if
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcatebuyPF_10(arrcate(3,intc))%></td>
    <td style="text-align:center"><%= mper/100%></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="2" style="text-align:center">상품</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10I%></td>
    <td><%=clsCMeachulLog.FItem_10I%></td>
    <td><%=clsCMeachulLog.FbuyPF_10I%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcatebuyPF_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td style="text-align:center">매입</td> 
    <td><%=clsCMeachulLog.FOds_10I%></td>
    <td><%=clsCMeachulLog.FItem_10I%></td>
    <td><%=clsCMeachulLog.FbuyPF_10I%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcatebuyPF_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">수수료</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10C%></td>
    <td><%=clsCMeachulLog.FItem_10C%></td>
    <td><%=clsCMeachulLog.FbuyPF_10C%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10C(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcatebuyPF_10C(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<%for i = 0 to iTotCnt
   if  clsCMeachulLog.Fsitename(i) = "10x10" and ( clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U")  then
%>
<tr bgcolor="#ffffff">
    <td style="text-align:center"><%=getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i))%></td> 
    <td><%=clsCMeachulLog.FOds(i)%></td>
    <td><%=clsCMeachulLog.FItem(i)%></td>
    <td><%=clsCMeachulLog.FbuyPF(i)%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%= clsCMeachulLog.FcateItem(i,arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcatebuyPF(i,arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<% end if
next%> 
<tr bgcolor="#ffffff">
    <td rowspan="6" style="text-align:center">제휴</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_P%></td>
    <td><%=clsCMeachulLog.FItem_P%></td>
    <td><%=clsCMeachulLog.FbuyPF_P%></td>
    <td><% buyPer= 0
      if clsCMeachulLog.Fbuy_P > 0 then 
             buyPer = round((clsCMeachulLog.FbuyPF_P/clsCMeachulLog.Fbuy_P)*100)
           end if
        %>
     <%=buyPer/100%>
    </td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
            if clsCMeachulLog.Fcatebuy_P(arrcate(3,intc)) > 0 then
             mper = round((clsCMeachulLog.FcatebuyPF_P(arrcate(3,intc))/clsCMeachulLog.Fcatebuy_P(arrcate(3,intc)))*100)
            end if
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_P(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcatebuyPF_P(arrcate(3,intc))%></td>
    <td style="text-align:center"><%=mper/100 %></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="2" style="text-align:center">상품</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_PI%></td>
    <td><%=clsCMeachulLog.FItem_PI%></td>
    <td><%=clsCMeachulLog.FbuyPF_PI%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcatebuyPF_PI(arrcate(3,intc)) %></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td style="text-align:center">매입</td> 
    <td><%=clsCMeachulLog.FOds_PI%></td>
    <td><%=clsCMeachulLog.FItem_PI%></td>
    <td><%=clsCMeachulLog.FbuyPF_PI%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcatebuyPF_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">수수료</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_PC%></td>
    <td><%=clsCMeachulLog.FItem_PC%></td>
    <td><%=clsCMeachulLog.FbuyPF_PC%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PC(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcatebuyPF_PC(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<%for i = 0 to iTotCnt
    if  clsCMeachulLog.Fsitename(i) <> "10x10" and ( clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U")  then
%>
<tr bgcolor="#ffffff">
    <td style="text-align:center"><%=getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i)) %></td> 
    <td><%=clsCMeachulLog.FOds(i)%></td>
    <td><%=clsCMeachulLog.FItem(i)%></td>
    <td><%=clsCMeachulLog.FbuyPF(i)%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem(i,arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcatebuyPF(i,arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
 <% end if
next%> 


<tr bgcolor="#ffffff">
    <td rowspan="13" style="text-align:center">보너스쿠폰</td>
    <td colspan="3" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_sum%></td>
    <td><%=clsCMeachulLog.FItem_sum%></td>
    <td><%=clsCMeachulLog.FBC_sum%></td>
    <td><% buyPer= 0 
			if clsCMeachulLog.Fbuy_sum > 0 then 
			buyPer = round((clsCMeachulLog.FBC_sum/clsCMeachulLog.Fbuy_sum)*100)
			end if
    %><%=buyPer/100%></td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
          if clsCMeachulLog.Fcatebuy_Sum(arrcate(3,intc)) > 0 then 
           mper = round((clsCMeachulLog.FcateBC_Sum(arrcate(3,intc))/clsCMeachulLog.Fcatebuy_Sum(arrcate(3,intc)))*100)
          end if 
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_Sum(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateBC_Sum(arrcate(3,intc))%></td>
    <td style="text-align:center"><%=mper/100 %></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="6" style="text-align:center">10x10</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10%></td>
    <td><%=clsCMeachulLog.FItem_10%></td>
    <td><%=clsCMeachulLog.FBC_10%></td>
    <td>
    <%  buyPer= 0 
			if clsCMeachulLog.Fbuy_10 > 0 then 
			buyPer = round((clsCMeachulLog.FBC_10/clsCMeachulLog.Fbuy_10)*100)
			end if
    %>
    <%=buyPer/100%>
    </td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
           if clsCMeachulLog.Fcatebuy_10(arrcate(3,intc)) > 0 then
              mper = round((clsCMeachulLog.FcateBC_10(arrcate(3,intc))/clsCMeachulLog.Fcatebuy_10(arrcate(3,intc)))*100)
           end if
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateBC_10(arrcate(3,intc))%></td>
    <td style="text-align:center"><%= mper/100%></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="2" style="text-align:center">상품</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10I%></td>
    <td><%=clsCMeachulLog.FItem_10I%></td>
    <td><%=clsCMeachulLog.FBC_10I%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateBC_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> </td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td style="text-align:center">매입</td> 
    <td><%=clsCMeachulLog.FOds_10I%></td>
    <td><%=clsCMeachulLog.FItem_10I%></td>
    <td><%=clsCMeachulLog.FBC_10I%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateBC_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">수수료</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10C%></td>
    <td><%=clsCMeachulLog.FItem_10C%></td>
    <td><%=clsCMeachulLog.FBC_10C%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10C(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateBC_10C(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<%for i = 0 to iTotCnt
   if  clsCMeachulLog.Fsitename(i) = "10x10" and ( clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U")  then
%>
<tr bgcolor="#ffffff">
    <td style="text-align:center"><%=getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i))%></td> 
    <td><%=clsCMeachulLog.FOds(i)%></td>
    <td><%=clsCMeachulLog.FItem(i)%></td>
    <td><%=clsCMeachulLog.FBC(i)%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%= clsCMeachulLog.FcateItem(i,arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateBC(i,arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<% end if
next%> 
<tr bgcolor="#ffffff">
    <td rowspan="6" style="text-align:center">제휴</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_P%></td>
    <td><%=clsCMeachulLog.FItem_P%></td>
    <td><%=clsCMeachulLog.FBC_P%></td>
    <td><% buyPer= 0 
           if clsCMeachulLog.Fbuy_P > 0 then 
            buyPer = round((clsCMeachulLog.FBC_P/clsCMeachulLog.Fbuy_P)*100)
          end if
        %>
     <%=buyPer/100%>
    </td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
            if clsCMeachulLog.Fcatebuy_P(arrcate(3,intc)) > 0 then
             mper = round((clsCMeachulLog.FcateBC_P(arrcate(3,intc))/clsCMeachulLog.Fcatebuy_P(arrcate(3,intc)))*100)
            end if
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_P(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateBC_P(arrcate(3,intc))%></td>
    <td style="text-align:center"><%=mper/100 %></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="2" style="text-align:center">상품</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_PI%></td>
    <td><%=clsCMeachulLog.FItem_PI%></td>
    <td><%=clsCMeachulLog.FBC_PI%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateBC_PI(arrcate(3,intc)) %></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td style="text-align:center">매입</td> 
    <td><%=clsCMeachulLog.FOds_PI%></td>
    <td><%=clsCMeachulLog.FItem_PI%></td>
    <td><%=clsCMeachulLog.FBC_PI%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateBC_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">수수료</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_PC%></td>
    <td><%=clsCMeachulLog.FItem_PC%></td>
    <td><%=clsCMeachulLog.FBC_PC%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PC(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateBC_PC(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<%for i = 0 to iTotCnt
    if  clsCMeachulLog.Fsitename(i) <> "10x10" and ( clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U")  then
%>
<tr bgcolor="#ffffff">
    <td style="text-align:center"><%=getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i)) %></td> 
    <td><%=clsCMeachulLog.FOds(i)%></td>
    <td><%=clsCMeachulLog.FItem(i)%></td>
    <td><%=clsCMeachulLog.FBC(i)%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem(i,arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateBC(i,arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
 <% end if
next%> 


<tr bgcolor="#ffffff">
    <td rowspan="13" style="text-align:center">취급액</td>
    <td colspan="3" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_sum%></td>
    <td><%=clsCMeachulLog.FItem_sum%></td>
    <td><%=clsCMeachulLog.FMP_sum%></td>
    <td> </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2) 
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_Sum(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMP_Sum(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="6" style="text-align:center">10x10</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10%></td>
    <td><%=clsCMeachulLog.FItem_10%></td>
    <td><%=clsCMeachulLog.FMP_10%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2) 
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMP_10(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="2" style="text-align:center">상품</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10I%></td>
    <td><%=clsCMeachulLog.FItem_10I%></td>
    <td><%=clsCMeachulLog.FMP_10I%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMP_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> </td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td style="text-align:center">매입</td> 
    <td><%=clsCMeachulLog.FOds_10I%></td>
    <td><%=clsCMeachulLog.FItem_10I%></td>
    <td><%=clsCMeachulLog.FMP_10I%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMP_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">수수료</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10C%></td>
    <td><%=clsCMeachulLog.FItem_10C%></td>
    <td><%=clsCMeachulLog.FMP_10C%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10C(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMP_10C(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<%for i = 0 to iTotCnt
   if  clsCMeachulLog.Fsitename(i) = "10x10" and ( clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U")  then
%>
<tr bgcolor="#ffffff">
    <td style="text-align:center"><%=getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i))%></td> 
    <td><%=clsCMeachulLog.FOds(i)%></td>
    <td><%=clsCMeachulLog.FItem(i)%></td>
    <td><%=clsCMeachulLog.FMP(i)%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%= clsCMeachulLog.FcateItem(i,arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMP(i,arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<% end if
next%> 
<tr bgcolor="#ffffff">
    <td rowspan="6" style="text-align:center">제휴</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_P%></td>
    <td><%=clsCMeachulLog.FItem_P%></td>
    <td><%=clsCMeachulLog.FMP_P%></td>
    <td> </td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2) 
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_P(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMP_P(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="2" style="text-align:center">상품</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_PI%></td>
    <td><%=clsCMeachulLog.FItem_PI%></td>
    <td><%=clsCMeachulLog.FMP_PI%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMP_PI(arrcate(3,intc)) %></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td style="text-align:center">매입</td> 
    <td><%=clsCMeachulLog.FOds_PI%></td>
    <td><%=clsCMeachulLog.FItem_PI%></td>
    <td><%=clsCMeachulLog.FMP_PI%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMP_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">수수료</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_PC%></td>
    <td><%=clsCMeachulLog.FItem_PC%></td>
    <td><%=clsCMeachulLog.FMP_PC%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PC(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMP_PC(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<%for i = 0 to iTotCnt
    if  clsCMeachulLog.Fsitename(i) <> "10x10" and ( clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U")  then
%>
<tr bgcolor="#ffffff">
    <td style="text-align:center"><%=getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i)) %></td> 
    <td><%=clsCMeachulLog.FOds(i)%></td>
    <td><%=clsCMeachulLog.FItem(i)%></td>
    <td><%=clsCMeachulLog.FMP(i)%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem(i,arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMP(i,arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
 <% end if
next%> 


<tr bgcolor="#ffffff">
    <td rowspan="13" style="text-align:center">취급액 수익</td>
    <td colspan="3" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_sum%></td>
    <td><%=clsCMeachulLog.FItem_sum%></td>
    <td><%=clsCMeachulLog.FMPPF_sum%></td>
    <td><% buyPer= 0 
			if clsCMeachulLog.FMP_sum > 0 then 
			buyPer = round((clsCMeachulLog.FMPPF_sum/clsCMeachulLog.FMP_sum)*100)
			end if
        %>
        <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
          if clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)) > 0 then 
            mper = round((clsCMeachulLog.FcateMPPF_Sum(arrcate(3,intc))/clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)))*100)
          end if 
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_Sum(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_Sum(arrcate(3,intc))%></td>
    <td style="text-align:center"><%=mper/100 %></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="6" style="text-align:center">10x10</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10%></td>
    <td><%=clsCMeachulLog.FItem_10%></td>
    <td><%=clsCMeachulLog.FMPPF_10%></td>
    <td><% buyPer= 0 
			if clsCMeachulLog.FMP_10 > 0 then 
			buyPer = round((clsCMeachulLog.FMPPF_10/clsCMeachulLog.FMP_10)*100)
			end if
    %>
    <%=buyPer/100%>
    </td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
           if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then
               mper = round((clsCMeachulLog.FcateMPPF_10(arrcate(3,intc))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100)
           end if
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_10(arrcate(3,intc))%></td>
    <td style="text-align:center"><%= mper/100%></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="2" style="text-align:center">상품</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10I%></td>
    <td><%=clsCMeachulLog.FItem_10I%></td>
    <td><%=clsCMeachulLog.FMPPF_10I%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> </td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td style="text-align:center">매입</td> 
    <td><%=clsCMeachulLog.FOds_10I%></td>
    <td><%=clsCMeachulLog.FItem_10I%></td>
    <td><%=clsCMeachulLog.FMPPF_10I%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_10I(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">수수료</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_10C%></td>
    <td><%=clsCMeachulLog.FItem_10C%></td>
    <td><%=clsCMeachulLog.FMPPF_10C%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_10C(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_10C(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<%for i = 0 to iTotCnt
   if  clsCMeachulLog.Fsitename(i) = "10x10" and ( clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U")  then
%>
<tr bgcolor="#ffffff">
    <td style="text-align:center"><%=getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i))%></td> 
    <td><%=clsCMeachulLog.FOds(i)%></td>
    <td><%=clsCMeachulLog.FItem(i)%></td>
    <td><%=clsCMeachulLog.FMPPF(i)%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%= clsCMeachulLog.FcateItem(i,arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF(i,arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<% end if
next%> 
<tr bgcolor="#ffffff">
    <td rowspan="6" style="text-align:center">제휴</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_P%></td>
    <td><%=clsCMeachulLog.FItem_P%></td>
    <td><%=clsCMeachulLog.FMPPF_P%></td>
    <td><%  buyPer= 0 
           if clsCMeachulLog.FMP_P > 0 then 
           buyPer = round((clsCMeachulLog.FMPPF_P/clsCMeachulLog.FMP_P)*100)
          end if
        %>
     <%=buyPer/100%>
    </td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
             if  clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then
                 mper = round((clsCMeachulLog.FcateMPPF_P(arrcate(3,intc))/ clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100)
             end if
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_P(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_P(arrcate(3,intc))%></td>
    <td style="text-align:center"><%=mper/100 %></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="2" style="text-align:center">상품</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_PI%></td>
    <td><%=clsCMeachulLog.FItem_PI%></td>
    <td><%=clsCMeachulLog.FMPPF_PI%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_PI(arrcate(3,intc)) %></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td style="text-align:center">매입</td> 
    <td><%=clsCMeachulLog.FOds_PI%></td>
    <td><%=clsCMeachulLog.FItem_PI%></td>
    <td><%=clsCMeachulLog.FMPPF_PI%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_PI(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">수수료</td>
    <td style="text-align:center">합계</td>
    <td><%=clsCMeachulLog.FOds_PC%></td>
    <td><%=clsCMeachulLog.FItem_PC%></td>
    <td><%=clsCMeachulLog.FMPPF_PC%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem_PC(arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_PC(arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
<%for i = 0 to iTotCnt
    if  clsCMeachulLog.Fsitename(i) <> "10x10" and ( clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U")  then
%>
<tr bgcolor="#ffffff">
    <td style="text-align:center"><%=getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i)) %></td> 
    <td><%=clsCMeachulLog.FOds(i)%></td>
    <td><%=clsCMeachulLog.FItem(i)%></td>
    <td><%=clsCMeachulLog.FMPPF(i)%></td>
    <td></td>
     <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
    %>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateItem(i,arrcate(3,intc))%></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF(i,arrcate(3,intc))%></td>
    <td style="text-align:center"></td>
    <%   next
    end if%>
</tr>
 <% end if
next%> 

<%
 dim  totvc,totvc_10, totvc_p, totsc_10, totsc_P , totwh_10, totwh_p
    dim arrPF, intp
    dim hidx_vc1, hidx_vc1_10, hidx_vc1_P, hidx_sc_10 ,hidx_wh_10, hidx_sc_p ,hidx_wh_p ,cper
        totvc = 0 : totvc_10=0:totvc_p=0: totsc_10=0: totsc_P=0 : totwh_10=0:totwh_p=0

	   clsCMeachulLog.fnGetprofitlossdata   
       totsc_10 = clsCMeachulLog.FTotSC_10
       totsc_P = clsCMeachulLog.FTotSC_P
      
       totwh_10= clsCMeachulLog.FTotWH_10
       totwh_p= clsCMeachulLog.FTotWH_P

       totvc_10 = totsc_10 + totwh_10
       totvc_P = totsc_P + totwh_p

       totvc = totvc_10 + totvc_P


%>
<tr bgcolor="#ffffff">
    <td rowspan="19" style="text-align:center">변동비1</td>
    <td colspan="3" style="text-align:center">합계</td>
    <td></td>
    <td></td>
    <td><%=totvc%></td>
    <td><% buyPer= 0 
        if clsCMeachulLog.FMP_sum > 0 then 
            buyPer = round((totvc/clsCMeachulLog.FMP_sum)*100)
        end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
            if  clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)) > 0 then  
               cper = round(( clsCMeachulLog.FTotVC1_Cate(arrcate(3,intC))/ clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)))*100)
            end if 
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FTotVC1_Cate(arrcate(3,intC))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr> 
<tr bgcolor="#ffffff">
    <td rowspan="10" style="text-align:center">10x10</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td></td>
    <td></td>
    <td><%=totvc_10%></td>
    <td><% buyPer= 0 
        if clsCMeachulLog.FMP_10  > 0 then 
            buyPer = round((totvc_10/clsCMeachulLog.FMP_10)*100)
        end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
            if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then  
                  cper = round(( clsCMeachulLog.FTotVC1_Cate_10(arrcate(3,intC))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100)
            end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FTotVC1_Cate_10(arrcate(3,intC))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr> 
<tr bgcolor="#ffffff">
    <td rowspan="5" style="text-align:center">판매수수료</td>
    <td  style="text-align:center">합계</td>
    <td></td>
    <td></td>
    <td><%=totsc_10%></td>
    <td><% buyPer= 0 
        if clsCMeachulLog.FMP_10 > 0 then 
            buyPer = round((totsc_10/clsCMeachulLog.FMP_10)*100)
        end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
            if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then  
                cper = round(( clsCMeachulLog.FTotSC_Cate_10(arrcate(3,intC))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100)
            end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FTotSC_Cate_10(arrcate(3,intC))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr>  
<%for intp = 0 to clsCMeachulLog.Fscrow-1   
  if clsCMeachulLog.FAccCIdx(intp) = 5 and clsCMeachulLog.FsiteNM(intp) ="10x10" then 
%> 
<tr bgcolor="#ffffff">
    <td  style="text-align:center"><%=clsCMeachulLog.FAccNM(intp)%></td> 
    <td></td>
    <td></td>
    <td><%= clsCMeachulLog.FPFPrice(intp)%></td>
    <td> </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2) 
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FCatePrice(intp, arrcate(3,intC))%></td>
    <td style="text-align:center"> </td>
    <%   next
    end if%>
</tr>  
<% end if
next%>
<tr bgcolor="#ffffff">
    <td rowspan="4" style="text-align:center">물류비</td>
    <td  style="text-align:center">합계</td>
    <td></td>
    <td></td>
    <td><%=totwh_10%></td>
    <td><% buyPer= 0 
        if clsCMeachulLog.FMP_10 > 0 then 
            buyPer = round((totwh_10/clsCMeachulLog.FMP_10)*100)
        end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
            if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then  
                cper = round(( clsCMeachulLog.FTotWH_Cate_10(arrcate(3,intC))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100)
            end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FTotWH_Cate_10(arrcate(3,intC))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr>
<%for intp = 0 to clsCMeachulLog.Fscrow-1   
  if clsCMeachulLog.FAccCIdx(intp) = 6 and clsCMeachulLog.FsiteNM(intp) ="10x10" then 
%> 
<tr bgcolor="#ffffff">
    <td  style="text-align:center"><%=clsCMeachulLog.FAccNM(intp)%></td> 
    <td></td>
    <td></td>
    <td><%= clsCMeachulLog.FPFPrice(intp)%></td>
    <td> </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2) 
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FCatePrice(intp, arrcate(3,intC))%></td>
    <td style="text-align:center"> </td>
    <%   next
    end if%>
</tr>  
<% end if
next%>
<tr bgcolor="#ffffff">
    <td rowspan="8" style="text-align:center">제휴</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td></td>
    <td></td>
    <td><%=totvc_P%></td>
    <td><% buyPer= 0 
        if clsCMeachulLog.FMP_P > 0 then 
            buyPer = round((totvc_P/clsCMeachulLog.FMP_P)*100)
        end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
            if clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then  
                  cper = round(( clsCMeachulLog.FTotVC1_Cate_P(arrcate(3,intC))/clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100)
            end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FTotVC1_Cate_P(arrcate(3,intC))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr> 
<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">판매수수료</td>
    <td  style="text-align:center">합계</td>
    <td></td>
    <td></td>
    <td><%=totsc_P%></td>
    <td><% buyPer= 0 
        if clsCMeachulLog.FMP_P > 0 then 
           buyPer = round((totsc_P/clsCMeachulLog.FMP_P)*100)
        end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
           if clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then  
              cper = round(( clsCMeachulLog.FTotSC_Cate_P(arrcate(3,intC))/clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100)
           end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FTotSC_Cate_P(arrcate(3,intC))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr>   
<%for intp = 0 to clsCMeachulLog.Fscrow-1   
  if clsCMeachulLog.FAccCIdx(intp) = 5 and clsCMeachulLog.FsiteNM(intp) ="제휴" then 
%> 
<tr bgcolor="#ffffff">
    <td  style="text-align:center"><%=clsCMeachulLog.FAccNM(intp)%></td> 
    <td></td>
    <td></td>
    <td><%= clsCMeachulLog.FPFPrice(intp)%></td>
    <td> </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2) 
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FCatePrice(intp, arrcate(3,intC))%></td>
    <td style="text-align:center"> </td>
    <%   next
    end if%>
</tr>  
<% end if
next%>
<tr bgcolor="#ffffff">
    <td rowspan="4" style="text-align:center">물류비</td>
    <td  style="text-align:center">합계</td>
    <td></td>
    <td></td>
    <td><%=totwh_P%></td>
    <td><% buyPer= 0 
       if clsCMeachulLog.FMP_P > 0 then 
           buyPer = round((totwh_P/clsCMeachulLog.FMP_P)*100)
       end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
            if clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then  
                cper = round(( clsCMeachulLog.FTotWH_Cate_P(arrcate(3,intC))/clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100)
            end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FTotWH_Cate_P(arrcate(3,intC))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr>
<%for intp = 0 to clsCMeachulLog.Fscrow-1   
  if clsCMeachulLog.FAccCIdx(intp) = 6 and clsCMeachulLog.FsiteNM(intp) ="제휴" then 
%> 
<tr bgcolor="#ffffff">
    <td  style="text-align:center"><%=clsCMeachulLog.FAccNM(intp)%></td> 
    <td></td>
    <td></td>
    <td><%= clsCMeachulLog.FPFPrice(intp)%></td>
    <td> </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2) 
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FCatePrice(intp, arrcate(3,intC))%></td>
    <td style="text-align:center"> </td>
    <%   next
    end if%>
</tr>  
<% end if
next%>
<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">공헌이익 1</td>
    <td colspan="3" style="text-align:center">합계</td>
    <td></td>
    <td></td>
    <td><%=clsCMeachulLog.FMPPF_sum- totvc%></td>
    <td><% buyPer= 0 
       if (clsCMeachulLog.FMP_sum) > 0 then 
			buyPer = round(((clsCMeachulLog.FMPPF_sum- totvc)/clsCMeachulLog.FMP_sum)*100)
	   end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
             if clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)) > 0 then  
                    cper = round(((clsCMeachulLog.FcateMPPF_Sum(arrcate(3,intc))- clsCMeachulLog.FTotVC1_Cate(arrcate(3,intC)))/clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)))*100)
             end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=(clsCMeachulLog.FcateMPPF_Sum(arrcate(3,intc))- clsCMeachulLog.FTotVC1_Cate(arrcate(3,intC)))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td colspan="3" style="text-align:center">공헌이익 1(10x10)</td> 
    <td></td>
    <td></td>
    <td><%=clsCMeachulLog.FMPPF_10-totvc_10%></td>
    <td><% buyPer= 0 
       if (clsCMeachulLog.FMP_10) > 0 then 
			buyPer = round(((clsCMeachulLog.FMPPF_10-totvc_10)/clsCMeachulLog.FMP_10)*100)
	   end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
            if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then
                cper = round(((clsCMeachulLog.FcateMPPF_10(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_10(arrcate(3,intC)))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100)
            end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_10(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_10(arrcate(3,intC))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td colspan="3" style="text-align:center">공헌이익 1(제휴)</td> 
    <td></td>
    <td></td>
    <td><%=clsCMeachulLog.FMPPF_P-totvc_P%></td>
    <td><% buyPer= 0 
        if clsCMeachulLog.FMP_P > 0 then 
                buyPer = round(((clsCMeachulLog.FMPPF_P-totvc_P)/clsCMeachulLog.FMP_P)*100)
            end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
             if clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then  
                cper = round(((clsCMeachulLog.FcateMPPF_P(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_P(arrcate(3,intC)))/clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100) 
             end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_P(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_P(arrcate(3,intC))  %></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td rowspan="5" style="text-align:center">변동비2</td>
    <td colspan="3" style="text-align:center">합계</td>
    <td></td>
    <td></td>
    <td><%=clsCMeachulLog.FTotMF_10 + clsCMeachulLog.FTotMF_P%></td>
    <td><% buyPer= 0 
        if (clsCMeachulLog.FMP_sum) > 0 then  
			buyPer = round(((clsCMeachulLog.FTotMF_10 + clsCMeachulLog.FTotMF_P)/clsCMeachulLog.FMP_sum)*100)
		end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
            if clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)) > 0 then  
                cper = round(( clsCMeachulLog.FTotVC2_Cate(arrcate(3,intC))/clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)))*100)
            end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FTotVC2_Cate(arrcate(3,intC))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr> 
<tr bgcolor="#ffffff">
    <td rowspan="2" style="text-align:center">10x10</td>
    <td colspan="2" style="text-align:center">합계</td>
    <td></td>
    <td></td>
    <td><%=clsCMeachulLog.FTotMF_10%></td>
    <td><% buyPer= 0 
       if clsCMeachulLog.FMP_10 > 0 then 
            buyPer = round((clsCMeachulLog.FTotMF_10/clsCMeachulLog.FMP_10)*100)
       end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
           if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then  
                cper = round(( clsCMeachulLog.FTotVC2_Cate_10(arrcate(3,intC))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100)
           end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FTotVC2_Cate_10(arrcate(3,intC))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr>
 <%for intp = 0 to clsCMeachulLog.Fscrow-1   
  if clsCMeachulLog.FAccCIdx(intp) = 9 and clsCMeachulLog.FsiteNM(intp) ="10x10" then 
%> 
<tr bgcolor="#ffffff">
    <td colspan="2" style="text-align:center"><%=clsCMeachulLog.FAccNM(intp)%></td> 
    <td></td>
    <td></td>
    <td><%= clsCMeachulLog.FPFPrice(intp)%></td>
    <td> </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2) 
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FCatePrice(intp, arrcate(3,intC))%></td>
    <td style="text-align:center"> </td>
    <%   next
    end if%>
</tr>  
<% end if
next%>
<tr bgcolor="#ffffff">
    <td rowspan="2" style="text-align:center">제휴</td>
    <td  colspan="2" style="text-align:center">합계</td>
    <td></td>
    <td></td>
    <td><%=clsCMeachulLog.FTotMF_P%></td>
    <td><% buyPer= 0 
       if  clsCMeachulLog.FMP_P > 0 then 
            buyPer = round((clsCMeachulLog.FTotMF_P/ clsCMeachulLog.FMP_P)*100)
      end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
            if clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then  
                  cper = round(( clsCMeachulLog.FTotVC2_Cate_P(arrcate(3,intC))/clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100)
            end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FTotVC2_Cate_P(arrcate(3,intC))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr> 
<%for intp = 0 to clsCMeachulLog.Fscrow-1   
  if clsCMeachulLog.FAccCIdx(intp) = 9 and clsCMeachulLog.FsiteNM(intp) ="제휴" then 
%> 
<tr bgcolor="#ffffff">
    <td colspan="2" style="text-align:center"><%=clsCMeachulLog.FAccNM(intp)%></td> 
    <td></td>
    <td></td>
    <td><%= clsCMeachulLog.FPFPrice(intp)%></td>
    <td> </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2) 
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FCatePrice(intp, arrcate(3,intC))%></td>
    <td style="text-align:center"> </td>
    <%   next
    end if%>
</tr>  
<% end if
next%>

<tr bgcolor="#ffffff">
    <td rowspan="3" style="text-align:center">공헌이익 2</td>
    <td colspan="3" style="text-align:center">합계</td>
    <td></td>
    <td></td>
    <td><%= clsCMeachulLog.FMPPF_sum- totvc-(clsCMeachulLog.FTotMF_10 + clsCMeachulLog.FTotMF_P) %></td>
    <td><% buyPer= 0 
      if (clsCMeachulLog.FMP_sum) > 0 then 
			buyPer = round(((clsCMeachulLog.FMPPF_sum- totvc- clsCMeachulLog.FTotMF_10 + clsCMeachulLog.FTotMF_P )/clsCMeachulLog.FMP_sum)*100)
			end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
             if clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)) > 0 then  
                    cper = round(((clsCMeachulLog.FcateMPPF_Sum(arrcate(3,intc))- clsCMeachulLog.FTotVC1_Cate(arrcate(3,intC))- clsCMeachulLog.FTotVC2_Cate(arrcate(3,intC)))/clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)))*100)
                    end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%= (clsCMeachulLog.FcateMPPF_Sum(arrcate(3,intc))- clsCMeachulLog.FTotVC1_Cate(arrcate(3,intC))- clsCMeachulLog.FTotVC2_Cate(arrcate(3,intC)))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td  colspan="3" style="text-align:center">공헌이익 2(10x10)</td> 
    <td></td>
    <td></td>
    <td><%=clsCMeachulLog.FMPPF_10-totvc_10-clsCMeachulLog.FTotMF_10%></td>
    <td><% buyPer= 0 
      if clsCMeachulLog.FMP_10 > 0 then 
			buyPer = round(((clsCMeachulLog.FMPPF_10-totvc_10-clsCMeachulLog.FTotMF_10)/clsCMeachulLog.FMP_10)*100)
			end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
            if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then
                        cper = round(((clsCMeachulLog.FcateMPPF_10(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_10(arrcate(3,intC))-clsCMeachulLog.FTotVC2_Cate_10(arrcate(3,intC)))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100)
                    end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_10(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_10(arrcate(3,intC))-clsCMeachulLog.FTotVC2_Cate_10(arrcate(3,intC))%></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr>
<tr bgcolor="#ffffff">
    <td colspan="3"  style="text-align:center">공헌이익 2(제휴)</td> 
    <td></td>
    <td></td>
    <td><%=clsCMeachulLog.FMPPF_P-totvc_P-clsCMeachulLog.FTotMF_P%></td>
    <td><% buyPer= 0 
        if  clsCMeachulLog.FMP_P > 0 then 
                buyPer = round(((clsCMeachulLog.FMPPF_P-totvc_P-clsCMeachulLog.FTotMF_P )/ clsCMeachulLog.FMP_P)*100)
            end if
    %>
    <%=buyPer/100%>
    </td>
    <%if isArray(arrcate) then
         for intc = 0 to ubound(arrcate,2)
              if clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then  
                    cper = round(((clsCMeachulLog.FcateMPPF_P(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_P(arrcate(3,intC))-clsCMeachulLog.FTotVC2_Cate_P(arrcate(3,intC)))/clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100)
                    end if
    %>
    <td style="text-align:center"></td>
    <td style="text-align:center"> <%=clsCMeachulLog.FcateMPPF_P(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_P(arrcate(3,intC))-clsCMeachulLog.FTotVC2_Cate_P(arrcate(3,intC)) %></td>
    <td style="text-align:center"><%=cper/100%></td>
    <%   next
    end if%>
</tr>
</table>
<%
   set clsCMeachulLog  = nothing 
    
 %> 
 <!-- #include virtual="/lib/db/dbSTSclose.asp" -->
 <!-- #include virtual="/lib/db/db3close.asp" --> 
 <!-- #include virtual="/lib/db/dbclose.asp" -->