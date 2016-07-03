<%
Function iPosProcess(serviceUrl, requestParams)
	Set PostConnection =server.CreateObject("MSXML2.ServerXMLHTTP")
	serviceUrl = serviceUrl
	PostConnection.Open "POST", serviceUrl, False
	PostConnection.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	PostConnection.Send requestParams
	iPosProcess = PostConnection.responseText
	Set PostConnection = Nothing
End Function
' ---------------------------------------
Function xmlParser(xmlString)
	Set objDoc = server.createObject("Microsoft.XMLDOM")
	objDoc.loadXML(xmlString)
	statusNode = objDoc.getElementsByTagName("Status").item(0).text
	If statusNode="Y" Then
	ACSUrlNode = objDoc.getElementsByTagName("ACSUrl").item(0).text
	TermUrlNode = objDoc.getElementsByTagName("TermUrl").item(0).text
	MDNode = objDoc.getElementsByTagName("MD").item(0).text
	PAReqNode = objDoc.getElementsByTagName("PaReq").item(0).text
	MessageErrorCode = objDoc.getElementsByTagName("MessageErrorCode").item(0).text
	resultDic = Array(statusNode,PAReqNode,ACSUrlNode,TermUrlNode,MDNode,MessageErrorCode)
	Else
	MessageErrorCode = objDoc.getElementsByTagName("MessageErrorCode").item(0).text
	resultDic = Array(statusNode,MessageErrorCode)
	End If
	xmlParser = resultDic
End Function
' ---------------------------------------
Function GetGuid() 
	Set TypeLib = CreateObject("Scriptlet.TypeLib") 
	GetGuid = "SIP_" + Replace(Replace(Replace(Left(CStr(TypeLib.Guid), 38),"{",""),"}",""),"-","") 
	Set TypeLib = Nothing 
End Function
' ---------------------------------------

toplam="1.00"
taksit=""
kkno="5549604812345678"
kkay="01"
kkyil="2020"
kkcvv="123"
uyeno="934566782"
uyesifre="911455678"
donusURL="http://www.siteadi.com/sonuc.asp"

	mpiServiceUrl=	"https://3dsecure.vakifbank.com.tr:4443/MPIAPI/MPI_Enrollment.aspx"
	krediKartiNumarasi = trim(kkno)
	sonKullanmaTarihi = right("00"&kkyil,2)&right("00"&kkay,2)
	if left(kkno,1)="4" then kartTipi="100" else kartTipi="200" : end if
	tutar = toplam
	paraKodu = "949"
	taksitSayisi=taksit
	islemNumarasi = GetGuid()
	uyeIsyeriNumarasi = right("000000000000000"&uyeno,15)
	uyeIsYeriSifresi = uyesifre
	SuccessURL = donusURL
	FailureURL = donusURL
	ekVeri = "1"
	params = "Pan="+krediKartiNumarasi+"&ExpiryDate="+sonKullanmaTarihi+"&PurchaseAmount="+tutar+"&Currency="+paraKodu+"&BrandName="+kartTipi+"&VerifyEnrollmentRequestId="+islemNumarasi+"&SessionInfo="+ekVeri+"&MerchantId="+uyeIsyeriNumarasi+"&MerchantPassword="+uyeIsYeriSifresi+"&SuccessURL="+SuccessURL+"&FailureURL="+FailureURL+"&InstallmentCount="+taksitSayisi
	resultXml = iPosProcess(mpiServiceUrl,params)
	result = xmlParser(resultXml)
		If result(0)="Y" Then
			Response.Clear
			Response.Write("<html>")
			Response.Write("<head>")
			Response.Write("<META HTTP-EQUIV='Content-Type' content='text/html; charset=Windows-1254'>")
			Response.Write("<script language='JavaScript'>")
			Response.Write("function submitForm(form) {")
			Response.Write("form.submit();")
			Response.Write("}")
			Response.Write("</script>")
			Response.Write("<title></title></head>")
			Response.Write("<body OnLoad='submitForm(document.downloadForm);' >")
			Response.Write("<FORM id='downloadForm' name='downloadForm' method='post' action='"&result(2)&"'>")
			Response.Write("<INPUT name='PaReq' type='hidden' value='"&result(1)&"'>")
			Response.Write("<INPUT name='TermUrl' type='hidden' value='"&result(3)&"'>")
			Response.Write("<INPUT name='MD' type='hidden' value='"&result(4)&"'>")
			Response.Write("</form>")
			Response.Write("</body>")
			Response.Write("</html>")
			Response.End()
		else
			response.write "Ödeme Tamamlanamadı. Hata Kodu : "&result(0)&" - Hata Mesajı : "&result(1)
			response.end()
		end if
		%>
