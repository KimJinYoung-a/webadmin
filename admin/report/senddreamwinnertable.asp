<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<% 
Response.AddHeader "Cache-Control","no-cache" 
Response.AddHeader "Expires","-1" 
Response.AddHeader "Pragma","no-cache" 

Response.ContentType = "application/vnd.ms-excel"

Response.AddHeader "Content-Disposition", "inline; filename=senddreamwinner.xls"
Response.AddHeader "Content-Description", "ASP Generated Data"
%>

<!-- main logic -->    
<%
dim fromDate,toDate
dim fY,fM,fD
dim tY,tM,tD

fY = request("fY")
fM = request("fM")
fD = request("fD")

tY = request("tY")
tM = request("tM")
tD = request("tD")

if (fY="") then fY = cstr(year(now()))
if (fM="") then fM = cstr(month(now()))
if (fD="") then fD = cstr(day(now()))
if (tY="") then tY = cstr(year(now()))
if (tM="") then tM = cstr(month(now()))
if (tD="") then tD = cstr(day(now()))

fromDate = DateSerial(fY, fM, fD) 
toDate = DateSerial(tY, tM, tD+1)
	
    '결과물에 대한 통계
    dim sql,tmp,cnt,total,wintotal
		
	cnt =0
	total = 0
	
    
    sql = "select count(d.id) as cnt"
	sql = sql & " from tbl_item i,"
	sql = sql & " tbl_etc_user e,"
	sql = sql & " tbl_new_senddream d"
	
	sql = sql & " where iswinner='Y'"
	sql = sql & " and d.itemid=i.itemid"
	sql = sql & " and e.sitename='cara'"
	sql = sql & " and e.userid=d.userid"

	sql = sql & " and d.regdate > '" + CStr(fromDate) + "'"
	sql = sql & " and d.regdate < '" + CStr(toDate) + "'"

    rsget.Open sql,dbget,1
    if  not rsget.EOF  then        
        wintotal = rsget("cnt")
    end if
    rsget.Close
%> 

<html>
	<body>
      <table border="1" cellspacing="1" cellpadding="3" >
        <tr align="center"> 
          <td width="120" class="a" ><font >구분</font></td>
          <td width="300" class="a" ><font >내용</font></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="120" class="a">검색기간</td>
         
          	<td class="a" width="300"><%=fromDate%> ~ <%= DateSerial(tY, tM, tD) %></td>
          
        </tr>
        
        
        <tr bgcolor="#FFFFFF"> 
          <td width="120" class="a">총당첨자수</td>
          <td class="a" width="300">총 <%= FormatNumber(FormatCurrency(wintotal),0)%> 명</td>
        </tr>

      </table>
      <br>
<%
	dim maxcnt,sql2,n
			
	sql = "select d.id,d.userid,e.username,i.itemname,mailcontents,d.regdate,T.*"
	sql = sql & " from tbl_item i,"
	sql = sql & " tbl_etc_user e,"
	sql = sql & " tbl_new_senddream d"
	sql = sql & " left join "
	sql = sql & " tbl_new_senddream_winner as T on"
	sql = sql & " d.id=masterid"
	sql = sql & " where iswinner='Y'"
	sql = sql & " and d.itemid=i.itemid"
	sql = sql & " and e.sitename='cara'"
	sql = sql & " and e.userid=d.userid"

	sql = sql & " and d.regdate > '" + CStr(fromDate) + "'"
	sql = sql & " and d.regdate < '" + CStr(toDate) + "'"
	sql = sql & " order by d.id"
	rsget.Open sql,dbget,1 
	
	maxcnt = 0
	    
%>
     
      <table border="1" cellspacing="1" cellpadding="3" bgcolor="#6CA0B2" class="verdana_9pt">
        <tr align="center"> 
          <td width="50" align="center" ><font color="#FFFFFF">ID</font></td> 
          <td width="50" align="center" ><font color="#FFFFFF">상품</font></td> 
          <td width="60" align="center" ><font color="#FFFFFF">아이디</font></td>
          <td width="60" align="center" ><font color="#FFFFFF">이름</font></td>
          <td width="60" align="center" ><font color="#FFFFFF">등록일</font></td>
          <td width="500"><font color="#FFFFFF">내용</font></td>
          
        </tr>
<%
    if  not rsget.EOF  then
        dim pathname,fso,ofile,inx,ipkumdiv_name,subtotalprice_F,fileName,tmpFileName
 
        do until (rsget.EOF)
%>        
        <tr bgcolor="#FFFFFF">
          <td width="50" align="center" ><%= rsget("id") %></td> 
          <td width="50" align="center" ><%= rsget("itemname") %></td> 
          <td width="60" align="center" valign="middle"><%= rsget("userid") %></td>
          <td width="60" align="center" valign="middle"><%= rsget("username") %></td>
          <td width="60" align="center" valign="middle"><%= rsget("regdate") %></td>
          <td width="500" align="left" ><%= db2html(rsget("mailcontents")) %></td>
          
          
        </tr>
<%
            rsget.MoveNext
        loop
%>      
      <tr>
      </form>
      </table>
      
<%
rsget.close
%>	          
      
      <br>
      <div align="center"><br>
      </div>
<%  else %>

  <tr bgcolor="#FFFFFF"> 
    <td colspan="8">
      <center>해당기간에 거래내역이 없습니다.</center>        
    </td>
  </tr>

<%  end if %>   
	</body>
</html>