<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%
dim userid, page
userid  = requestCheckvar(request("userid"),32)
page    = requestCheckvar(request("page"),10)

if (page="") then page=1

dim frmname, rebankaccount, rebankownername, rebankname
frmname         = request("frmname")
rebankaccount   = request("rebankaccount")
rebankownername = request("rebankownername")
rebankname      = request("rebankname")

''환불정보
dim orefundInfo
set orefundInfo = New CCSASList
orefundInfo.FCurrpage = page
orefundInfo.FPageSize = 10
orefundInfo.FRectUserID = userid

if (userid<>"") then
    orefundInfo.GetHisOldRefundInfo
end if

dim i

%>
<body style="margin:10 10 10 10" bgcolor="#FFFFFF">
<script language='javascript'>
var frmname         = "<%= frmname %>";
var rebankname      = "<%= rebankname %>";
var rebankaccount   = "<%= rebankaccount %>";
var rebankownername = "<%= rebankownername %>";


function selectPreAcct(i_rebankname, i_rebankaccount, i_rebankownername){
    eval("opener." + frmname + "." + rebankname).value = i_rebankname;
    eval("opener." + frmname + "." + rebankaccount).value = i_rebankaccount;
    eval("opener." + frmname + "." + rebankownername).value = i_rebankownername;

    window.close();
}

function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function setDisplayNo(iasid){
    if (confirm("계좌정보를 이전 환불계좌목록에서 제외시킵니다.\n\n진행하시겠습니까?") == true) {
    	frmAction.asid.value = iasid;

    	frmAction.submit();
    }
}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" class="a" bgcolor="#FFFFFF">

<form name="frmAction" method="post" action="pop_cs_PreRefundAccount_process.asp">
<input type="hidden" name="mode" value="setdisplayno">
<input type="hidden" name="page" value="1">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="frmname" value="<%= frmname %>">
<input type="hidden" name="rebankname" value="<%= rebankname %>">
<input type="hidden" name="rebankaccount" value="<%= rebankaccount %>">
<input type="hidden" name="rebankownername" value="<%= rebankownername %>">
<input type="hidden" name="asid" value="">
</form>

<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="frmname" value="<%= frmname %>">
<input type="hidden" name="rebankname" value="<%= rebankname %>">
<input type="hidden" name="rebankaccount" value="<%= rebankaccount %>">
<input type="hidden" name="rebankownername" value="<%= rebankownername %>">
<tr>
    <td>
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#FFFFFF">
            <td ><img src="/images/icon_star.gif" align="absbottom">&nbsp; <b>이전 환불 계좌 목록 : <%= userid %></b></td>
        </tr>
        </table>
    </td>
</tr>
</form>
</tr>
    <td>
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="<%= adminColor("topbar") %>">
            <td height=25 align=center>은행</td>
            <td align=center>계좌</td>
            <td align=center>예금주</td>
            <td width="100"align=center>비고</td>
        </tr>
        <% if orefundInfo.FResultCount>0 then %>
        <% for i=0 to orefundInfo.FResultCount-1 %>
			<%
			if (orefundInfo.FItemList(i).Fencmethod = "TBT") then
			    ''사용안함.
				orefundInfo.FItemList(i).Frebankaccount = TBTDecrypt(orefundInfo.FItemList(i).FencAccount)
			elseif (orefundInfo.FItemList(i).Fencmethod = "PH1") then
	            orefundInfo.FItemList(i).Frebankaccount = orefundInfo.FItemList(i).Fdecaccount
	        elseif (orefundInfo.FItemList(i).Fencmethod = "AE2") then
	            orefundInfo.FItemList(i).Frebankaccount = orefundInfo.FItemList(i).Fdecaccount
			end if
			%>
        <tr bgcolor="#FFFFFF">
            <td height=25 align=center><%= orefundInfo.FItemList(i).Frebankname %></td>
            <td align=center><%= orefundInfo.FItemList(i).Frebankaccount %></td>
            <td align=center><%= orefundInfo.FItemList(i).Frebankownername %></td>
            <td align=right>
            	<input class="button" type="button" value="선택" onClick="selectPreAcct('<%= orefundInfo.FItemList(i).Frebankname %>','<%= orefundInfo.FItemList(i).Frebankaccount %>','<%= orefundInfo.FItemList(i).Frebankownername %>');">
            	&nbsp; &nbsp; &nbsp; &nbsp;
            	<input class="button" type="button" value="X" onClick="setDisplayNo(<%= orefundInfo.FItemList(i).Fasid %>);">
            </td>
        </tr>
        <% next %>
        <tr bgcolor="#FFFFFF">
            <td colspan="4" align="center">
            <%
				if orefundInfo.HasPreScroll then
					Response.Write "<a href='javascript:NextPage(" & orefundInfo.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for i=0 + orefundInfo.StarScrollPage to orefundInfo.FScrollCount + orefundInfo.StarScrollPage - 1

					if i>orefundInfo.FTotalpage then Exit for

					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:NextPage(" & i & ")'>[" & i & "]</a> "
					end if

				next

				if orefundInfo.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:NextPage(" & i & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
            </td>
        </tr>
        <% else %>
        <tr bgcolor="#FFFFFF">
            <td colspan="4" align="center">
            <% if (userid<>"") then %>
            [검색 결과가 없습니다.]
            <% else %>
            [<strong>비회원</strong>은 이전 내역을 검색 하실 수 없습니다.]
            <% end if %>
            </td>
        </tr>
        <% end if %>
        </table>
    </td>
</tr>
</table>
</body>
<%
set orefundInfo = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->