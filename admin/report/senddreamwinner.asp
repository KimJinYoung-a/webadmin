<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/bct_admin_header.asp"-->
<%
const MenuPos1 = "Admin"
const MenuPos2 = "센드드림"
%>
<!-- #include virtual="/admin/bct_admin_menupos.asp"-->

<%
const Maxlines = 10
dim totalpage, totalnum, q

dim gotopage,fY,fM,fD,tY,tM,tD
dim fromDate,toDate,jnx,tmpStr,siteId,settle
dim showtype, IsAdmin


siteId = session("ssBctId")
if (siteID="10x10") then IsAdmin = true

dim searchId,mxlen
searchId = request("searchId")
showtype = request("showtype")
gotopage = request("gotopage")
settle = request("settle")
mxlen = request("mxlen")

fY = request("fY")
fM = request("fM")
fD = request("fD")

tY = request("tY")
tM = request("tM")
tD = request("tD")

''서동팔 수정..
''기본값적용..
If gotopage <> "" then
   session("gotopage") = CInt(gotopage)
else
   Session("gotopage") = 1
   gotopage = session("gotopage") 
end if

gotopage= Cint(gotopage)

if (settle="") then settle ="D"

if (Not IsNumeric(mxlen)) or mxlen="" then mxlen =20


if showtype="2" then
	if (fY="") then fY = cstr(year(now()))
	if (fM="") then fM = cstr(month(now()))
	
	fromDate = DateSerial(fY, fM, 1)
	toDate = DateSerial(fY, fM+1, 1) 	
else
	if (fY="") then fY = cstr(year(now()))
	if (fM="") then fM = cstr(month(now()))
	if (fD="") then fD = cstr(day(now()))
	if (tY="") then tY = cstr(year(now()))
	if (tM="") then tM = cstr(month(now()))
	if (tD="") then tD = cstr(day(now()))
	
	fromDate = DateSerial(fY, fM, fD) 
	toDate = DateSerial(tY, tM, tD+1)
end if

%>    
      <table width="100%" border="0" cellpadding="0" cellspacing="3" bgcolor="#CCCCCC">
      <form name="bari" method="get" action="senddreamwinnertable.asp" target="_blank">
    	<input type="hidden" name="goTopage">
    	<input type="hidden" name="showtype" value="<%= showtype %>">
    	<input type="hidden" name="Xler" value="">
    
        <tr> 
          <td width="5%"></td>
          <td width="600" class="a"> 
                  <select name="fY">
                    <!-- <option value="" <%if fY="" then response.write " selected"%>>년</option> -->
               <% for jnx=1 to 6 %>
               <%   tmpStr = "200"+CStr(jnx) %>
                    <option value="<%=tmpStr%>" <%if fY=tmpStr then response.write " selected"%>><%=tmpStr%></option>
               <% next %>
                  </select>
                  <select name="fM">
                    <!-- <option value="" <%if fM="" then response.write " selected"%>>월</option> -->
               <% for jnx=1 to 12 %>
               <%   tmpStr = CStr(jnx) %>
                    <option value="<%=tmpStr%>" <%if fM=tmpStr then response.write " selected"%>><%=tmpStr%></option>
               <% next %>
                  </select>
                  <select name="fD">
                    <!-- <option value="" <%if fD="" then response.write " selected"%>>일</option> -->
               <% for jnx=1 to 31 %>
               <%   tmpStr = CStr(jnx) %>
                    <option value="<%=tmpStr%>" <%if fD=tmpStr then response.write " selected"%>><%=tmpStr%></option>
               <% next %>
                  </select>
                  <span class="a">~</span> 
                  <select name="tY">
                    <!-- <option value="" <%if tY="" then response.write " selected"%>>년</option> -->
               <% for jnx=1 to 6 %>
               <%   tmpStr = "200"+CStr(jnx) %>
                    <option value="<%=tmpStr%>" <%if tY=tmpStr then response.write " selected"%>><%=tmpStr%></option>
               <% next %>
                  </select>
                  <select name="tM">
                    <!--  <option value="" <%if tM="" then response.write " selected"%>>월</option> -->
               <% for jnx=1 to 12 %>
               <%   tmpStr = CStr(jnx) %>
                    <option value="<%=tmpStr%>" <%if tM=tmpStr then response.write " selected"%>><%=tmpStr%></option>
               <% next %>
                  </select>
                  <select name="tD">
                    <!--  <option value="" <%if tD="" then response.write " selected"%>>일</option> -->
               <% for jnx=1 to 31 %>
               <%   tmpStr = CStr(jnx) %>
                    <option value="<%=tmpStr%>" <%if tD=tmpStr then response.write " selected"%>><%=tmpStr%></option>
               <% next %>
                  </select>
               
          </td>
          <td align="left" class="a">
          	
          </td>
          <td align="right" class="a">
          	<input type="image" src="/images/search2.gif" width="74" height="22" align="middle">
          </td>
          <td width="5%"></td>
        </tr>
        <tr> 
          <td width="5%"></td>
          <td width="600" class="a" valign="center">
          
          </td>
          <td >	
          </td>
          <td class="a">	
          </td>
          <td width="5%"></td>
        </tr>
      </table>
    </form>

</td>
</tr>
</table>


<!-- #include virtual="/admin/bct_admin_tail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
