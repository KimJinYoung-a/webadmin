<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/partner/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/adminpartner/partnerCls.asp" -->
<%
'###########################################################
' Description : ��ü���� header : ���� ��ü������ �귣�� ����
' History : 2014.05.07 ������  ���� 
'########################################################### 
dim userid
dim shiftid
dim groupid
dim preuserdiv
dim ref 
dim ClsPartner,sid,scompany_name,semail,sbigo,suserdiv,sgroupid,curruserdiv
userid  = session("ssBctId")
groupid = session("ssGroupid")
preuserdiv = session("ssBctDiv")
shiftid = requestCheckVar(request("shiftid"),32) 
ref = request.ServerVariables("HTTP_REFERER") 

''���� �귣�� ���� ��������
set ClsPartner = new CPartner
	ClsPartner.FRectShiftID = shiftid
	ClsPartner.FRectGroupID = groupid
	ClsPartner.fnGetBrandChangeLogin  
	sid	 			= ClsPartner.Fid          
	scompany_name 	= ClsPartner.Fcompany_name 
	semail  		= ClsPartner.Femail       
	sbigo 			= ClsPartner.Fbigo        
	suserdiv 	 	= ClsPartner.Fuserdiv     
	sgroupid     	= ClsPartner.Fgroupid     
	curruserdiv 	= ClsPartner.Fcuserdiv     
set ClsPartner = nothing 

 ''�����Ϸ��� �귣���� ������ ������..	
	if  (isNull(sid) or sid="") then
		Call Alert_move("��ȿ�� �귣�尡 �ƴմϴ�.sid="+sid,"/partner/index.asp")  
		response.end
	end if

    session("ssBctId") 		= sid
    session("ssBctDiv") 	= suserdiv
    session("ssBctBigo") 	= sbigo
    session("ssBctCname") 	= db2html(scompany_name)
	session("ssBctEmail") 	= db2html(semail)
	session("ssGroupid") 	= sgroupid
session("chkOffShop") = 0			'-������ ���� �귣�� ���� ��üũó��(incadmingnb.asp ���� Ȯ��)
	response.Cookies("partner").domain = "10x10.co.kr"
    response.Cookies("partner")("userid") = session("ssBctId")
    response.Cookies("partner")("userdiv") = session("ssBctDiv") 

 ''���翡�� <=> �귣�� �� ����Ī �Ѱ��  
if (curruserdiv<>preuserdiv) then 
    ref = "/partner/index.asp"
end if 
if (curruserdiv="14") then 
    ref = "/lectureadmin/index.asp"
end if
if (curruserdiv="15") then ''2016/06/27
    ref = "/diyadmin/index.asp"
end if

 response.redirect ref
%> 
<!-- #include virtual="/lib/db/dbclose.asp" -->