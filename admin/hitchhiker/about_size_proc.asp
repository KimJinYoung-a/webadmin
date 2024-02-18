<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : HITCHHIKER ADMIN
'	History		: 2014.07.09 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/about_hitchhiker_contentsCls.asp"-->

<%
dim deviceidx, device_name, contents_size,  mode, sortnum, isusing, gubun
	gubun = requestcheckvar(request("gubun"),1)
	mode = requestcheckvar(request("mode"),16)
	isusing = requestcheckvar(request("isusing"),1)
	sortnum = requestcheckvar(request("sortnum"),5)
	deviceidx = requestcheckvar(request("deviceidx"),10)
	device_name = requestcheckvar(request("device_name"),32)
	contents_size = requestcheckvar(request("contents_size"),32)
	'response.write mode
	
	dim sqlstr, getdate
	if mode = "sizeedit" then 
		if deviceidx<>"" then
			sqlstr = " update db_sitemaster.dbo.tbl_hitchhiker_device_size set " '수정모드일때 db업데이트
			sqlstr = sqlstr & " isusing = '"& isusing &"' "
			sqlstr = sqlstr & " ,sortnum = '"& sortnum &"' "
			sqlstr = sqlstr & " ,device_name = '"& html2db(device_name) &"' "
			sqlstr = sqlstr & " ,contents_size = '"& html2db(contents_size) &"' "
			sqlstr = sqlstr & " where deviceidx = "& deviceidx &" "
			'response.write sqlstr
			dbget.execute sqlstr
		else
			sqlstr = "insert into db_sitemaster.dbo.tbl_hitchhiker_device_size (device_name, contents_size, sortnum, isusing, gubun, regdate)" '신규입력모드
			sqlstr = sqlstr & " values ('" & html2db(device_name) & "','" & html2db(contents_size) & "','" & sortnum & "' , '" & isusing & "', '" & gubun & "',getdate())"
			'response.write sqlstr
			dbget.execute sqlstr
		end if
	else
		response.write "<script language = 'javascript'>alert('정상적인 경로가 아닙니다');</script>"
		dbget.close() : response.end
	end if
%>

<script language = "javascript">
	alert("저장되었습니다."); //저장되었습니다 라는 메시지띄움
	//opener.location.reload(); //이창을 띄운 부모창을 리로드함
	//self.close();			  //이창을 닫음
	//parent.location.reload();
	parent.location.href = "about_size_write.asp?menupos=<%=menupos%>";
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->

