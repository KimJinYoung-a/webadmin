<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_cls.asp"-->
<%

dim itemid, catecode, pagesize
itemid  = requestCheckvar(request("itemid"),10)
catecode  = requestCheckvar(request("catecode"),16)

pagesize = 500
if (itemid <> "") then
	pagesize = 20
end if

dim cdata
SET cdata = New cdataanalysis
	cdata.FCurrPage = 1
	cdata.FPageSize = pagesize
	cdata.FRectItemID = itemid
	cdata.FRectCateCode = catecode

	if (itemid = "") then
		cdata.Getdataanalysis_wish_list()
	else
		cdata.Getdataanalysis_wish_detail()
	end if

dim i, colnum

colnum = 10

%>
<table width="80%">
	<tr>
		<td>
			<a href="wishCollection.asp"><img src="http://webimage.10x10.co.kr/image/list/172/L001724513.jpg" width="60" border="0"></a><br />
			전체
		</td>
		<td>
			<a href="wishCollection.asp?catecode=102"><img src="http://webimage.10x10.co.kr/image/list/176/L001761791.jpg" width="60" border="0"></a><br />
			디지털
		</td>
		<td>
			<a href="wishCollection.asp?catecode=117"><img src="http://webimage.10x10.co.kr/image/list/176/L001763019.jpg" width="60" border="0"></a><br />
			패션의류
		</td>
		<td>
			<a href="wishCollection.asp?catecode=121"><img src="http://webimage.10x10.co.kr/image/list/175/L001757671.jpg" width="60" border="0"></a><br />
			가구
		</td>
		<td>
			<a href="wishCollection.asp?catecode=122114"><img src="http://webimage.10x10.co.kr/image/list/176/L001762781.jpg" width="60" border="0"></a><br />
			조명
		</td>
	</tr>
</table>
<table width="100%">
	<tr>
		<% for i = 0 to cdata.FResultCount - 1 %>
		<% if (i > 0) and (i mod colnum) = 0 then %></tr><tr><% end if %>
		<td>
			<% if True or (itemid = "") then %>
			<a href="wishCollection.asp?itemid=<%= cdata.FItemList(i).FitemID %>"><img src="<%= cdata.FItemList(i).FlistImage %>"></a>
			<% else %>
			<img src="<%= cdata.FItemList(i).FlistImage %>">
			<% end if %>
		</td>
	<% next %>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
