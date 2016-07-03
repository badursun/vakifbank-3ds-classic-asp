<%
Function iPosProcess(serviceUrl, requestXML)
		Set PostConnection = CreateObject("MSXML2.ServerXMLHTTP")
		serviceUrl = serviceUrl
		PostConnection.Open "POST", serviceUrl, False
		PostConnection.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		PostConnection.Send requestXML
		iPosProcess = PostConnection.responseText
		Set PostConnection = Nothing
End Function
' -------------------------------------------------------------
	taksit = session("karttk")
	uyeno = "934566782"
	uyesifre = "911455678"
	
	PosXML = "<?xml version=""1.0"" encoding=""utf-8""?>"
	PosXML = PosXML & "<VposRequest>"
		PosXML = PosXML & "<MerchantId>" & Request.Form("MerchantId") &"</MerchantId>"
		PosXML = PosXML & "<Password>"&uyeno&"</Password>"
		PosXML = PosXML & "<TerminalNo>"&uyesifre&"</TerminalNo>"
		PosXML = PosXML & "<TransactionType>Sale</TransactionType>"
		PosXML = PosXML & "<CurrencyAmount>"&session("karttp")&"</CurrencyAmount>"
		PosXML = PosXML & "<CurrencyCode>" & Request.Form("PurchCurrency") &"</CurrencyCode>"
		
		if taksit<>"" then
			if isnumeric(taksit) then
				if int(taksit)>1 then
					PosXML = PosXML & "<NumberOfInstallments>"&taksit&"</NumberOfInstallments>"
				end if
			end if
		end if
		
		PosXML = PosXML & "<Pan>"&session("kartno")&"</Pan>"
		PosXML = PosXML & "<Cvv>"&session("kartcv")&"</Cvv>"
		PosXML = PosXML & "<Expiry>20" & Request.Form("Expiry") &"</Expiry>"
		PosXML = PosXML & "<ECI>" & Request.Form("Eci") &"</ECI>"
		PosXML = PosXML & "<CAVV>" & Request.Form("Cavv") &"</CAVV>"
		PosXML = PosXML & "<MpiTransactionId>" & Request.Form("VerifyEnrollmentRequestId") &"</MpiTransactionId>"
		PosXML = PosXML & "<ClientIp>"&IP&"</ClientIp>"
		PosXML = PosXML & "<TransactionDeviceSource>0</TransactionDeviceSource>"
	PosXML = PosXML & "</VposRequest>"
	
	session("karttp")="abandon"
	session("kartno")="abandon"
	session("kartcv")="abandon"
	session("karttk")="abandon"
	
	  PostUrl="https://onlineodeme.vakifbank.com.tr:4443/VposService/v3/Vposreq.aspx"
		result = iPosProcess(PostUrl,  "prmstr=" + PosXML)
		Set objXML = Server.CreateObject("Microsoft.XMLDOM") 
		objXML.async = False 
		objXML.loadXML(result)
		Set sec = objXML.selectNodes("//VposResponse") 
		For z = 0 to (sec.Length - 1)
		ResultCode = sec(z).selectSingleNode ("ResultCode").ChildNodes(0).Text
		
			if ResultCode="0000" then
        response.write "Tebrikler ! Ödeme başarıyla tamamlandı."
        response.end()
			else
			ResultDetail = sec(z).selectSingleNode ("ResultDetail").ChildNodes(0).Text
  			response.write "Ödeme Tamamlanamadı. Hata Kodu : "&ResultCode&" - Hata Mesajı : "&ResultDetail
  			response.end()
			end if
			
		next
%>
