<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�彺Ʈ��Ʈ
' History : 2013.08.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/street/shopcls.asp"-->
<%
Dim itemidarr, sortnoarr, tmpSort, masteridx, cnt, i, sqlStr, mode, adminid
	sortnoarr 	= Request("sortnoarr")
	itemidarr 	= Request("itemidarr")
	masteridx 		= requestCheckVar(Request("masteridx"),20)
	mode 		= requestCheckVar(Request("mode"),20)
	menupos	= requestCheckVar(request("menupos"),10)
	
adminid = session("ssBctId")

if mode="sortedit" then
	If sortnoarr="" THEN
		Response.Write "<script language='javascript'>alert('������ �������� �ʾҽ��ϴ�.'); history.back(-1);</script>"
		dbget.close()	:	response.End
	end if
	
	'���û�ǰ �ľ�
	itemidarr = split(itemidarr,",")
	cnt = ubound(itemidarr)
	
	'// ���ļ��� ����
	If sortnoarr<>"" THEN
		sortnoarr =  split(sortnoarr,",")
		
		For i = 0 to cnt
			IF sortnoarr(i) = "" THEN
				 tmpSort = "0"				
			ELSE	
				 tmpSort = sortnoarr(i)	
			END IF
			
			sqlStr = "UPDATE db_brand.dbo.tbl_street_shop_collection SET" + vbcrlf
			sqlStr = sqlStr & " sortNo = "&tmpSort&"" + vbcrlf
			sqlStr = sqlStr & " ,lastupdate=getdate()" + vbcrlf
			sqlStr = sqlStr & " ,lastadminid = '"&adminid&"'" + vbcrlf
			sqlStr = sqlStr & " WHERE idx =" + itemidarr(i)
			
			'response.write sqlStr & "<Br>"
			dbget.execute sqlStr
		Next
	END IF

	response.write "<script language='javascript'>"
	response.write "	alert('����Ǿ����ϴ�');"
	response.write "	location.replace('/designer/brand/shop/collection/index.asp?menupos="&menupos&"');"
	response.write "</script>"
else
	Response.Write "<script language='javascript'>alert('�����ڰ� �����ϴ�.'); history.back(-1);</script>"
	dbget.close()	:	response.End
END IF	
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->