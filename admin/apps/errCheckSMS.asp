<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/lib/email/smslib.asp" -->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("61.252.133.2","61.252.133.9","61.252.133.70","110.93.128.121","110.93.128.92","172.16.0.206")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    response.write "ERR"
    dbget.Close()
    response.end
end if


''SMS
dim imsg, osms, ibody

imsg = requestCheckvar(request("imsg"),70)
ibody = requestCheckvar(request("ibody"),1500)

'dim isAgirl
'if (inStr(imsg,"Agirl")>0) or (inStr(imsg,"youareagirl")>0) or (inStr(imsg,"[29cm]")>0) then  ''or (inStr(imsg,"[29cm]")>0) �߰� 2015/10/20
'    isAgirl = true
'end if

dim isPushDemon, isSearchDemon : isPushDemon = False :  isSearchDemon = false

if (inStr(imsg,"PUSH_DEMON")>0) then
    isPushDemon = true
end if

if (inStr(imsg,"ksearch")>0) then
    isSearchDemon = true
end if

''�߰� �˻� ���л��� ������ �˻�����ڿ��Ը� ��Ƽ�� SKIP
if (InStr(imsg,"ksearch_INC")>0) and (hour(now())<9) then
    set osms = new CSMSClass
    call SendNormalSMS_LINK("010-6324-9110","",imsg)    ''������
    call SendNormalSMS_LINK("010-4782-3272","",imsg)    ''�����
    call SendNormalSMS_LINK("010-3350-6271","",imsg)     ''�̽���
    set osms = Nothing

    dbget.Close(): response.end
end if

IF (imsg<>"") then
    set osms = new CSMSClass
    if (isSearchDemon) then ''�˻� �� ���Ÿ� ������ ���
        call SendNormalSMS_LINK("010-6324-9110","",imsg)     ''������
        call SendNormalSMS_LINK("010-4782-3272","",imsg)     ''�����
        call SendNormalSMS_LINK("010-3350-6271","",imsg)     ''�̽���

        set osms = Nothing

        dbget.Close(): response.end
    end if
    
    call SendNormalSMS_LINK("010-9972-8517","",imsg)     ''������
    call SendNormalSMS_LINK("010-2303-1873","",imsg)     ''������
    call SendNormalSMS_LINK("010-6324-9110","",imsg)     ''������
    call SendNormalSMS_LINK("010-4782-3272","",imsg)     ''�����
    call SendNormalSMS_LINK("010-9177-8708","",imsg)     ''�ѿ��
    call SendNormalSMS_LINK("010-9459-1552","",imsg)     ''������
    call SendNormalSMS_LINK("010-3350-6271","",imsg)     ''�̽���
    call SendNormalSMS_LINK("010-4210-3402","",imsg)     ''���缮
    call SendNormalSMS_LINK("010-5046-1412","",imsg)     ''����ȣ
   
    set osms = Nothing
    

    response.write "OK"
else

ENd IF


%>

<!-- #include virtual="/lib/db/dbclose.asp" -->