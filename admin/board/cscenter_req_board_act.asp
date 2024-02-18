<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<%

dim companyrequest
dim boarditem
dim id, mode,categubun
dim user,comment,commmode
dim ipjumYN
dim page
dim gubun,sellgubun,onlymifinish,research,searchkey,catevalue
page=request("pg")
id = request("id")
mode = request("mode")
categubun=request("categubun")
user=html2db(request("user"))
comment=html2db(request("comment"))
sellgubun=request("sellgubun")
ipjumYN=request("ipjumYN")
gubun = request("gubun")
onlymifinish = request("onlymifinish")
research = request("research")
searchkey = request("searchkey")
catevalue=request("catevalue")
commmode=request("commmode")

dim opt
opt="pg="+ page + "&id=" + id + "&mode=" + mode + "&categubun=" + categubun + "&user=" + user + "&sellgubun=" + sellgubun + "&ipjumYN=" + ipjumYN + "&gubun=" + gubun + "&onlymifinish=" + onlymifinish + "&research=" + research +"&searchkey=" + searchkey + "&catevalue="+ catevalue + "&commmode=" + commmode

set companyrequest = New CCompanyRequest
'response.write opt
'dbget.close()	:	response.End
if (mode = "finish") then

        companyrequest.finish(id)
		response.write "<script>location.replace('cscenter_req_board_list.asp?" + opt + "');</script>"
		
elseif (mode="del") then
		
		companyrequest.delitem id
		response.write "<script>location.replace('cscenter_req_board_list.asp?" + opt + "');</script>"

elseif (mode="chcate") then

		companyrequest.catechange id,categubun
		
		response.write "<script>location.replace('cscenter_req_board_view.asp?" + opt + "');</script>"
		
elseif (mode="chsell") then
		
		companyrequest.sellchange id,sellgubun
						
		response.write "<script>location.replace('cscenter_req_board_view.asp?" + opt + "');</script>"
		
elseif (mode="ipjum") then
		
		companyrequest.ipjumchange id,ipjumYN
						
		response.write "<script>location.replace('cscenter_req_board_view.asp?" + opt + "');</script>"
		
elseif (mode="comm") then
		
		companyrequest.writecomm id,user,comment
						
		response.write "<script>location.replace('cscenter_req_board_view.asp?" + opt + "');</script>"

end if
set companyrequest=nothing

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->