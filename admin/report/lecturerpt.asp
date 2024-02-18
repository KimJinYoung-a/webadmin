<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
Sub SelectBoxDesignerItem1(selectedId)
   dim query1,tmp_str
   %><select name="tempid">
     <option value=''>-- 업체선택 --</option>
     <%
	query1 = "select  distinct userid, socname_kor from [db_user].[dbo].tbl_user_c" + vbcrlf
	query1 = query1 + " where userdiv=14" + vbcrlf
	query1 = query1 + " group by socname_kor, userid" + vbcrlf
	
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write ("<option value='" & rsget("userid")&"' "&tmp_str&">" & rsget("socname_kor") & "</option>")
			      tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub

dim oreport
dim oldlist
Dim tempid

oldlist = request("oldlist")
tempid = request("tempid")

set oreport = new CJumunMaster
oreport.FRectOldJumun = oldlist
oreport.FRectDesignerID = tempid
oreport.getLectureMeaChul

Dim i,p1,p2
Dim premonth_sellsum,premonth_sellcnt
%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a">
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역&nbsp;&nbsp;
		<% SelectBoxDesignerItem1 tempid %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">기간</font></td>
          <td class="a" width="600"><font color="#FFFFFF"></font></td>
          <td class="a" width="100"><font color="#FFFFFF">액수</font></td>
          <td class="a" width="50"><font color="#FFFFFF">건수</font></td>
          <td class="a" width="80"><font color="#FFFFFF">객단가</font></td>
        </tr>

		<% for i=0 to oreport.FResultCount-1 %>
		<%

			if i = 0 then
				premonth_sellsum = oreport.FMasterItemList(1).Fselltotal
				premonth_sellcnt = oreport.FMasterItemList(1).Fsellcnt 
			end if

			if oreport.maxt<>0 then
				p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*100)
			end if

			if oreport.maxc<>0 then
				p2 = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100)
			end if
		%>
        <tr bgcolor="#FFFFFF" height="10"  class="a">
				<td width="120" height="10">
					<%= oreport.FMasterItemList(i).Fsitename %>(강좌<%= oreport.FMasterItemList(i).FSocName %>건)
				</td>
				<td  height="10" width="600">
					 <div align="left"> <img src="/images/dot1.gif" height="4" width="<%= p1 %>%"></div><br>
					 <div align="left"> <img src="/images/dot2.gif" height="4" width="<%= p2 %>%"></div>
				</td>
				<td class="a" width="100" align="right">
					 <% if i = 0 then %>
					 <font color="#AAAAAA"><% = round((oreport.FMasterItemList(i).Fselltotal/premonth_sellsum)*100)-100 %>%</font><br>
					 <% end if %>
					 <%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %>원
				</td>
				<td class="a" width="80" align="right">
					 <% if i = 0 then %>
					 <font color="#AAAAAA"><% = round((oreport.FMasterItemList(i).Fsellcnt/premonth_sellcnt)*100)-100 %>%</font><br>
					 <% end if %>
					 <%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %>건
				</td>
				<td class="a" width="80" align="right">
					 <% if i = 0 then %>
					 <font color="#AAAAAA"><% = round((Clng(oreport.FMasterItemList(i).Fselltotal/oreport.FMasterItemList(i).Fsellcnt)/Clng(premonth_sellsum/premonth_sellcnt))*100)-100 %>%</font><br>
					 <% end if %>
					 <%= FormatNumber(Clng(oreport.FMasterItemList(i).Fselltotal/oreport.FMasterItemList(i).Fsellcnt),0) %>원
				</td>
        </tr>
        <% next %>
</table>
<%
set oreport = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->