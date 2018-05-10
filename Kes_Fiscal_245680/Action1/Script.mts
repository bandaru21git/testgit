'Script Name    	: Kes_Fiscal_245680 
'Purpose    		: Make a Reconnect quote payment
'Developed by  		: Mounika
'Developed Date 	: 29/06/2017
'TestDataSheet 		: Kes_Fiscal_245680
'TestCasePath		: Solution.Delivery/Integration testing Africa/Kenya/Fiscal/245680
''**************************************************************

On Error Resume Next

intStepCounter = 0

'Load environment File
LoadEnvironment

'Initializing the Test Script
Call gfOnInitialize(Environment("TestName"))

'Load test data specific to the test script
Call gfGetExcelRecordSet(Environment("TestDataFileLocation"),"CreateCustomer_Residential",gCountryClause,ObjResCustRecordSet,True)

'Load test data specific to the test script
Call gfGetExcelRecordSet(Environment("TestDataFileLocation"),Environment("TestName"),gCountryClause,ObjADORecordSet,True)

'Iterate through all the rows in the test data sheet
Do While Not ObjADORecordSet.EOF

	For intCounter = 1 To 1 Step 1
		'Login into Clarity application...
		If Login_Clarity(gBrowserType,gTestEnvironment,gLoginUser,gLoginPassword,gLoginCountry) Then
			
			Call gfReportExecutionStatus(micPass,"Login to Clarity Application for " & gTestEnvironment,"Successfully logged into " & strURLType & " Portal")
			
			'Loading test data.....
			strpackageName						= ObjADORecordSet.Fields.Item("packageName").Value
			strPaymentMethod					= ObjADORecordSet.Fields.Item("PaymentMethod").Value
			strProductStatus					= ObjADORecordSet.Fields.Item("ProductStatus").Value
			strHigherProduct 					= ObjADORecordSet.Fields.Item("HigherProduct").Value
			strHigherProductCode 				= ObjADORecordSet.Fields.Item("HigherProductCode").Value
			strInteractionMessage 				= ObjADORecordSet.Fields.Item("interactionMessage").value
			strHistoryEventDescription 			= ObjADORecordSet.Fields.Item("HistoryEventDescription").value
			sPackageStatus 						= ObjADORecordSet.Fields.Item("PackageStatus").value
			sTransactionType 					= ObjADORecordSet.Fields.Item("TransactionType").value
			sLedgerAccountDescription 			= ObjADORecordSet.Fields.Item("LedgerAccountDescription").value
			sCustomerAccountType 				= ObjADORecordSet.Fields.Item("CustomerAccountType").value
			strUpgradedpackageName				= ObjADORecordSet.Fields.Item("UpgradedpackageName").value
			strItemsList 						= ObjADORecordSet.Fields.Item("strItemsList").value 
				
			'Assigning Fiscal Related Printer
			 If AssignPrinter("Kenya_Test_Production") = False Then
				Call gfReportExecutionStatus(micFail,"Assign Printer","Failed to assign printer")		
			 End If
					
	 		'Log into the application  		
	   		If  Login_Clarity(gBrowserType,gTestEnvironment,gLoginUser,gLoginPassword,gLoginCountry) Then
	   		
	   			Call gfReportExecutionStatus(micDone,"Loginto Clarity Application for " & gTestEnvironment,"Successfully Logged into " & strURLType & " Portal")
				
				'Capture the device info from database.
				Call QueryHandler(gDecoder_SQLFileName,gDecoderModel_HD,ObjAppDBRecordSet,20,True,True)
				
				'Get device info from Database. 
				If Not ObjAppDBRecordSet.EOF Then
					strDevice = ObjAppDBRecordSet.Fields("FROM_DEVICE_SERIAL_NUMBER").Value
					strSmartCard=ObjAppDBRecordSet.Fields("TO_DEVICE_SERIAL_NUMBER").Value	
					Call gfReportExecutionStatus(micPass, "Getting Device From Database", "Successfully retrived device " & strDevice)
				Else
					Call gfReportExecutionStatus(micFail, "Getting Device From Database", "Failed to get Decoder information from Database")
				End If	
				
				'Create a Residential customer
				If CreateCustomer(ObjResCustRecordSet, "Residential", strCustomerNumber, strAccountNumber) <> True Then
					Call gfReportExecutionStatus(micFail, "Create a Gotv Residential customer", "Gotv Residential customer creation failed ")
				End If
				
				'Search specialist customer through global search.
				bln_GlobalSearch  =  GlobalSearch("Customer", strCustomerNumber)
				If  bln_GlobalSearch <> True Then
					Call gfReportExecutionStatus(micFail, "Load the Customer using Global Search", "Failed to load customer by global search")
				End If
			
				'Add devices and packages
				bln_AddDevicePackages  = AddDevicePackages(strDevice, strpackageName, "", False, False, False, "", "", "", "", "", "")
				If bln_AddDevicePackages <> True Then
					Call gfReportExecutionStatus(micFail, "Add Device and Packages to the Residential GoTv customer", "Failed to add the Device to the Gotv Residential Customer "&strCustomerNumber)
				End If
				
				'Capture paymnet for customer by calling function.
				blnMakePayment  =  MakePayment(strPaymentMethod,gCurrencyValue,"",strCustomerNumber)
				If blnMakePayment <> True Then
					Call gfReportExecutionStatus(micFail, "Verify Payment transaction", "Payment was not successfull for the added decoder "&strDevice)
				End If
								
				'Verifying Device Status after payment...
		        If VerifyProductInformation("", strDevice, strpackageName, "", "", strProductStatus, "", "", "") <> True Then
					Call gfReportExecutionStatus(micFail, "Get Active device for the customer", "Error in retriveing Active device for customer: "&strCustomerNumber)
				End If
				
				'Disconnecting the Packages
				gDictionaryObj_QueryHandler.Add "CustomerId",strCustomerNumber
				Call QueryHandler("sql_Update_AllCP_Status_DIS", "", ObjADORecordSet, 20, True, False)
				
				'Search specialist customer through global search.
				bln_GlobalSearch  =  GlobalSearch("Customer", strCustomerNumber)
				If  bln_GlobalSearch <> True Then
					Call gfReportExecutionStatus(micFail, "Load the Customer using Global Search", "Failed to load customer by global search")
				End If
				
				'Verifying Device Status after payment...
		        If VerifyProductInformation("", strDevice, strpackageName, "", "", "Disconnected", "", "", "") <> True Then
					Call gfReportExecutionStatus(micFail, "Get Active device for the customer", "Error in retriveing Active device for customer: "&strCustomerNumber)
				End If
					
				'Reconnecting Services					
				If ReconnectProduct(strSmartCard, "COMPLE36", "", "", "", True, "", True, "", "", ObjRecordsetAddHocData, ObjQuoteInformation)	<> True Then			
					Call gfReportExecutionStatus(micFail, "Load the Customer using Global Search", "Failed to load customer by global search")
				End If
				
				'Get Quote amount for package...
		        If IsObject(ObjQuoteInformation) Then
		        	intQuoteAmount = ObjQuoteInformation.Item("QuoteAmount")
		        End If
				
				'Validating Fiscal Payment Report
				bln_VerifyFiscalPaymentReceipt = VerifyFiscalPaymentReceipt(strItemsList)
				If bln_VerifyFiscalPaymentReceipt <> True Then
					Call gfReportExecutionStatus(micPass,"Verify Fiscal Payment Report","Failed to Validate Fiscal Payment report")	
				End If
								
				'Verifying Upgraded Product Status
				bln_VerifyProductInformation  =  VerifyProductInformation("", strSmartCardNumber, strUpgradedpackageName, "", "", strProductStatus, "", "", "")
				If bln_VerifyProductInformation <> True Then
					Call gfReportExecutionStatus(micFail, " Verify Changes for Upgrade product", "Failed to verify the product upgrade for smartcard: "&strSmartCardNumber)
				End If
				
				If intQuoteAmount > 0 Then
					If IsNumeric(intQuoteAmount) Then  intQuoteAmount = Abs(intQuoteAmount)
					'Verify Financial transactions for reconnection action.
					bln_VerifyFTs  =  VerifyFTs(strFinancialTransactionType, strLedgerAccountDescription, "", strAccountType)
					If bln_VerifyFTs <>  True Then
						Call gfReportExecutionStatus(micFail, "Verify Financial Transactions for Reconnection action", "Financial Transaction for amount: "&intQuoteAmount&" not generated for reconnection action")
					End If
				End If
								
			Else
        		Call gfReportExecutionStatus(micFail,"Launching Clarity","Failed to Launch " & strURLType & " Portal")
   			End If
			
		Else
			Call gfReportExecutionStatus(micFail, "Lgoin into clarity application", "Failed to Launch " & strURLType & " Portal")		
		End If
	Next
	
	'Close chrome browser
	Call CloseChrome
	
	' Close Clarity Application" 	
	Call gfReportExecutionStatus(micPass, "Close Clarity Application", "Successfully closed Clarity Application")
	
	'Move to next row in test data
	ObjADORecordSet.MoveNext	
Loop

'End of execution, generate the HTML Report
Call gfOnTerminate()




'*****************************************************************************************************************************
'!Function Name		: VerifyFiscalPaymentReceipt
'!Purpose			: To verify Fiscal Payment report
'!Input 			: strItemsList - Items that need to be validated in Payment Report need to be provided from Test Data seperated by ','
'!Output 	  		: Verifies the Existence of Item and Corresponding Value related to that Item
'!Developed By 		: Mounika
'!Date 		  		: 29 June'17
'*****************************************************************************************************************************
Public Function VerifyFiscalPaymentReceipt(ByVal strItemsList)
	VerifyFiscalPaymentReceipt = False
	If Browser("brw_Clarity").Page("pg_MakePayment").WebElement("wele_PrintPaymentReceipt").Exist Then
		Set objReport =	Browser("brw_Clarity").Page("pg_MakePayment").WebTable("column names:=YOUR TAX INVOICE","visible:=True")
		strItemsDisplayed = objReport.GetROProperty("innertext")
		PaymentReportItem = Split (strItemsList,",")
			For Count = 0 To Ubound(PaymentReportItem) step 1
				If Instr(1,strItemsDisplayed,PaymentReportItem(Count),1) >0  AND (PaymentReportItem(Count) <> "Fiscal Signature") Then
					RowNum = objReport.GetRowWithCellText(PaymentReportItem(Count))
					ColNum = objReport.ColumnCount(RowNum)
					'Getting Value displayed for Corresponding item
					ItemValue = objReport.GetCellData(RowNum,ColNum)
					Else If Instr(1,strItemsDisplayed,PaymentReportItem(Count),1) >0  AND (PaymentReportItem(Count) = "Fiscal Signature") Then
						RowNum = objReport.GetRowWithCellText(PaymentReportItem(Count))
						ColNum = objReport.ColumnCount(RowNum)
						'Getting Value displayed for Corresponding item
						ItemValue = objReport.GetCellData(RowNum+1,ColNum)
					End If
					Call gfReportExecutionStatus(micPass,"Validate Payment Report","Item "&ItemValue&" Validated for "&PaymentReportItem(Count)&"in Payment Report")
					If Count = Ubound(PaymentReportItem) Then
						Call gfReportExecutionStatus(micPass,"Verify Payment Report Items","All Items in payment Report are Verified")	
						VerifyFiscalPaymentReceipt = True
					End If
				End If
			Next
	End If
End Function


