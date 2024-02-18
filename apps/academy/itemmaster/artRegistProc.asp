<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : /apps/protoV1/loginProc.asp
' Discription : 로그인 처리
' Request : json > type, pushid, OS, versioncode, versionname, verserion
' Response : response > 결과
' History : 2016.10.18 허진원 : 신규 생성
'###############################################

'//헤더 출력
Response.ContentType = "text/html"

Dim addimgname, addimgtext
Dim imgbasic,imgadd1,imgadd2,itemvideo,opttype,opttypename1,opttypename2,opttypename3,optname1,optname2,optname3,optaddprice1,optaddprice2,optaddprice3,optaddbuyprice1
Dim optaddbuyprice2,optaddbuyprice3,designerid,defultmargine,defaultmaeipdiv,defaultFreeBeasongLimit,defaultDeliverPay,defaultDeliveryType,cd1,cd2,cd3,catecode,catedepth
Dim itemdiv,cstodr,reqMsg,reqImg,vatYn,limityn,useoptionyn,optlevel,optwintitle,keywords,safetyYn,safetyDiv,safetyNum,infoCd,infoChk,infoCont,infoDiv, requirecontents
Dim sellvat,buyvat,sellcash,buycash,mwdiv,sellyn,isusing,mileage,ordercomment,limitno,makername,sourcearea,itemsource,itemsize,itemWeight,makeremail,requireMakeDay,itemname

imgbasic = request.form("imgbasic")
imgadd1 = request.form("imgadd1")
imgadd2 = request.form("imgadd2")
itemvideo = request.form("itemvideo")
opttype = request.form("opttype")
opttypename1 = request.form("opttypename1")
opttypename2 = request.form("opttypename2")
opttypename3 = request.form("opttypename3")
optname1 = request.form("optname1")
optname2 = request.form("optname2")
optname3 = request.form("optname3")
optaddprice1 = request.form("optaddprice1")
optaddprice2 = request.form("optaddprice2")
optaddprice3 = request.form("optaddprice3")
optaddbuyprice1 = request.form("optaddbuyprice1")
optaddbuyprice2 = request.form("optaddbuyprice2")
optaddbuyprice3 = request.form("optaddbuyprice3")
designerid = request.form("designerid")
defultmargine = request.form("defultmargine")
defaultmaeipdiv = request.form("defaultmaeipdiv")
defaultFreeBeasongLimit = request.form("defaultFreeBeasongLimit")
defaultDeliverPay = request.form("defaultDeliverPay")
defaultDeliveryType = request.form("defaultDeliveryType")
cd1 = request.form("cd1")
cd2 = request.form("cd2")
cd3 = request.form("cd3")
catecode = request.form("catecode")
catedepth = request.form("catedepth")
itemdiv = request.form("itemdiv")
cstodr = request.form("cstodr")
reqMsg = request.form("reqMsg")
reqImg = request.form("reqImg")
vatYn = request.form("vatYn")
limityn = request.form("limityn")
useoptionyn = request.form("useoptionyn")
optlevel = request.form("optlevel")
optwintitle = request.form("optwintitle")
keywords = request.form("keywords")
safetyYn = request.form("safetyYn")
safetyDiv = request.form("safetyDiv")
safetyNum = request.form("safetyNum")
infoCd = request.form("infoCd")
infoChk = request.form("infoChk")
infoCont = request.form("infoCont")
infoDiv = request.form("infoDiv")
addimgname= request.form("addimgname")
addimgtext= request.form("addimgtext")
requirecontents= request.form("requirecontents")
sellvat= request.form("sellvat")
sellcash= request.form("sellcash")
buyvat= request.form("buyvat")
buycash= request.form("buycash")
mwdiv= request.form("mwdiv")
sellyn= request.form("sellyn")
isusing= request.form("isusing")
mileage= request.form("mileage")
ordercomment= request.form("ordercomment")
limitno= request.form("limitno")
makername = request.form("makername")
sourcearea = request.form("sourcearea")
itemsource = request.form("itemsource")
itemsize = request.form("itemsize")
itemWeight = request.form("itemWeight")
makeremail = request.form("makeremail")
requireMakeDay = request.form("requireMakeDay")
itemname = request.form("itemname")


Response.write "imgbasic : " + imgbasic + "<br>"
Response.write "imgadd1 : " + imgadd1 + "<br>"
Response.write "imgadd2 : " + imgadd2 + "<br>"
Response.write "itemvideo : " + itemvideo + "<br>"
Response.write "opttype : " + opttype + "<br>"
Response.write "opttypename1 : " + opttypename1 + "<br>"
Response.write "opttypename2 : " + opttypename2 + "<br>"
Response.write "opttypename3 : " + opttypename3 + "<br>"
Response.write "optname1 : " + optname1 + "<br>"
Response.write "optname2 : " + optname2 + "<br>"
Response.write "optname3 : " + optname3 + "<br>"
Response.write "optaddprice1 : " + optaddprice1 + "<br>"
Response.write "optaddprice2 : " + optaddprice2 + "<br>"
Response.write "optaddprice3 : " + optaddprice3 + "<br>"
Response.write "optaddbuyprice1 : " + optaddbuyprice1 + "<br>"
Response.write "optaddbuyprice2 : " + optaddbuyprice2 + "<br>"
Response.write "optaddbuyprice3 : " + optaddbuyprice3 + "<br>"
Response.write "designerid : " +  designerid + "<br>"
Response.write "defultmargine : " +  defultmargine + "<br>"
Response.write "defaultmaeipdiv : " +  defaultmaeipdiv + "<br>"
Response.write "defaultFreeBeasongLimit : " +  defaultFreeBeasongLimit + "<br>"
Response.write "defaultDeliverPay : " +  defaultDeliverPay + "<br>"
Response.write "defaultDeliveryType : " +  defaultDeliveryType + "<br>"
Response.write "cd1 : " + cd1 + "<br>"
Response.write "cd2 : " + cd2 + "<br>"
Response.write "cd3 : " + cd3 + "<br>"
Response.write "catecode : " + catecode + "<br>"
Response.write "catedepth : " + catedepth + "<br>"
Response.write "itemdiv : " + itemdiv + "<br>"
Response.write "cstodr : " + cstodr + "<br>"
Response.write "reqMsg : " + reqMsg + "<br>"
Response.write "reqImg : " + reqImg + "<br>"
Response.write "vatYn : " + vatYn + "<br>"
Response.write "limityn : " + limityn + "<br>"
Response.write "useoptionyn : " + useoptionyn + "<br>"
Response.write "optlevel : " + optlevel + "<br>"
Response.write "optwintitle : " + optwintitle + "<br>"
Response.write "keywords : " + keywords + "<br>"
Response.write "safetyYn : " + safetyYn + "<br>"
Response.write "safetyDiv : " + safetyDiv + "<br>"
Response.write "safetyNum : " + safetyNum + "<br>"
Response.write "infoCd : " +  infoCd + "<br>"
Response.write "infoChk : " +  infoChk + "<br>"
Response.write "infoCont : " +  infoCont + "<br>"
Response.write "infoDiv : " +  infoDiv + "<br>"
Response.write "addimgname : " +  addimgname + "<br>"
Response.write "addimgtext : " +  addimgtext + "<br>"
Response.write "requirecontents : " +  requirecontents + "<br>"
Response.write "sellvat : " +  sellvat + "<br>"
Response.write "sellcash : " +  sellcash + "<br>"
Response.write "buyvat : " +  buyvat + "<br>"
Response.write "buycash : " +  buycash + "<br>"
Response.write "buycash : " +  mwdiv + "<br>"
Response.write "buycash : " +  sellyn + "<br>"
Response.write "buycash : " +  isusing + "<br>"
Response.write "buycash : " +  mileage + "<br>"
Response.write "ordercomment : " +  ordercomment + "<br>"
Response.write "limitno : " +  limitno + "<br>"

Response.write "makername : " +  makername + "<br>"
Response.write "sourcearea : " +  sourcearea + "<br>"
Response.write "itemsource : " +  itemsource + "<br>"
Response.write "itemsize : " +  itemsize + "<br>"
Response.write "itemWeight : " +  itemWeight + "<br>"
Response.write "email : " +  makeremail + "<br>"
Response.write "requireMakeDay : " +  requireMakeDay + "<br>"
Response.write "limitno : " +  itemname + "<br>"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->