<!-- #include virtual="/lib/classes/street/managerCls.asp"-->
<%
'###########################################################
' Description :  브랜드스트리트
' History : 2013.08.19 김진영 생성
'			2013.08.29 한용민 수정
'###########################################################

'########################## 브랜드 매뉴 권한 셋팅 ############################
dim hello_yn, interview_yn, tenbytenand_yn, artistwork_yn, shop_collection_yn, shop_event_yn, lookbook_yn
dim brandgubun, brandgubunname, makerid_confirm
	makerid_confirm = session("ssBctID")

dim omenu
set omenu = new cmanager
	omenu.frectmakerid = makerid_confirm
	
	if makerid_confirm<>"" then
		omenu.sbbrandgubunlist_confirm
	end if
	
	if omenu.Ftotalcount > 0 then
		brandgubun = omenu.FOneItem.fbrandgubun
		brandgubunname = omenu.FOneItem.fbrandgubunname
		
		hello_yn = omenu.FOneItem.fhello_yn
		interview_yn = omenu.FOneItem.finterview_yn
		tenbytenand_yn = omenu.FOneItem.ftenbytenand_yn
		artistwork_yn = omenu.FOneItem.fartistwork_yn
		shop_collection_yn = omenu.FOneItem.fshop_collection_yn
		shop_event_yn = omenu.FOneItem.fshop_event_yn
		lookbook_yn = omenu.FOneItem.flookbook_yn
	end if
set omenu = nothing	
'########################## 브랜드 매뉴 권한 셋팅 ############################
%>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#FFFFFF" >
    <td><a href="/designer/brand/hello/index.asp?menupos=<%=menupos%>">HELLO</a></td>
    
    <% if artistwork_yn="Y" then %>
    	<td><a href="/designer/brand/Artist/index.asp?menupos=<%=menupos%>">Artist Work</a></td>
    <% end if %>

    <% if shop_collection_yn="Y" then %>
    	<td><a href="/designer/brand/shop/collection/index.asp?menupos=<%=menupos%>">SHOP_collection</a></td>
    <% end if %>
    
    <% if lookbook_yn="Y" then %>
		<td><a href="/designer/brand/lookbook/index.asp?menupos=<%=menupos%>">LOOKBOOK</a></td>
    <% end if %>		
</tr>
</table>
<br>
