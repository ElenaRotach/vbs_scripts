Dim WshShell, test
Set WshShell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\swift\"
wshshell.RegWrite path, ""
wshshell.RegWrite path & "sound", "1"
wshshell.RegWrite path & "dbName", "swift.mdb"
wshshell.RegWrite path & "dbPath", "C:\SWIFT\"
wshshell.RegWrite path & "inMess", "C:\SWIFT\IN\"
wshshell.RegWrite path & "outMess", "C:\SWIFT\OUT\"
wshshell.RegWrite path & "inArhiv", "C:\SWIFT\IN\"
wshshell.RegWrite path & "outArhiv", "C:\SWIFT\OUT\"
path = "HKEY_CURRENT_USER\Software\swift\columnWidth\"
wshshell.RegWrite path & "transactionReferenceNumber_20", "100"
wshshell.RegWrite path & "valueDate_30V", "100"
wshshell.RegWrite path & "date_32", "100"
wshshell.RegWrite path & "currency_32", "100"
wshshell.RegWrite path & "amount_32", "100"
wshshell.RegWrite path & "currency_33B", "100"
wshshell.RegWrite path & "amount_33B", "100"
wshshell.RegWrite path & "orderingCustomer_50", "100"
wshshell.RegWrite path & "orderingInstitution_52", "100"
wshshell.RegWrite path & "senderCorrespondent_53", "100"
wshshell.RegWrite path & "receiverCorrespondent_54", "100"
wshshell.RegWrite path & "intermediaryInstitution_56", "100"
wshshell.RegWrite path & "accountWithInstitution_57", "100"
wshshell.RegWrite path & "beneficiaryInstitution_58", "100"
wshshell.RegWrite path & "beneficiaryCustomer_59", "100"
wshshell.RegWrite path & "processingCharacteristic", "100"
wshshell.RegWrite path & "mess_direction", "100"
wshshell.RegWrite path & "comment", "100"
wshshell.RegWrite path & "dateTime_mess", "100"
wshshell.RegWrite path & "referenceMess", "100"
wshshell.RegWrite path & "fin", "100"
wshshell.RegWrite path & "swiftNumberBankKontragent", "100"
wshshell.RegWrite path & "naimBankKontragent", "100"
wshshell.RegWrite path & "thread", "100"
wshshell.RegWrite path & "fileName", "100"
wshshell.RegWrite path & "direction", "100"
wshshell.RegWrite path & "id", "100"
path = "HKEY_CURRENT_USER\Software\swift\columnIndex\"
wshshell.RegWrite path & "transactionReferenceNumber_20", "1"
wshshell.RegWrite path & "valueDate_30V", "2"
wshshell.RegWrite path & "date_32", "3"
wshshell.RegWrite path & "currency_32", "4"
wshshell.RegWrite path & "amount_32", "5"
wshshell.RegWrite path & "currency_33B", "6"
wshshell.RegWrite path & "amount_33B", "7"
wshshell.RegWrite path & "orderingCustomer_50", "8"
wshshell.RegWrite path & "orderingInstitution_52", "9"
wshshell.RegWrite path & "senderCorrespondent_53", "10"
wshshell.RegWrite path & "receiverCorrespondent_54", "11"
wshshell.RegWrite path & "intermediaryInstitution_56", "12"
wshshell.RegWrite path & "accountWithInstitution_57", "13"
wshshell.RegWrite path & "beneficiaryInstitution_58", "14"
wshshell.RegWrite path & "beneficiaryCustomer_59", "15"
wshshell.RegWrite path & "processingCharacteristic", "16"
wshshell.RegWrite path & "mess_direction", "17"
wshshell.RegWrite path & "comment", "18"
wshshell.RegWrite path & "dateTime_mess", "19"
wshshell.RegWrite path & "referenceMess", "20"
wshshell.RegWrite path & "fin", "21"
wshshell.RegWrite path & "swiftNumberBankKontragent", "22"
wshshell.RegWrite path & "naimBankKontragent", "23"
wshshell.RegWrite path & "thread", "24"
wshshell.RegWrite path & "fileName", "25"
wshshell.RegWrite path & "direction", "26"
wshshell.RegWrite path & "id", "0"
path = "HKEY_CURRENT_USER\Software\swift\columnName\"

wshshell.RegWrite path & "transactionReferenceNumber_20", "20 transactionReferenceNumber_20"
wshshell.RegWrite path & "valueDate_30V", "30 valueDate_30V"
wshshell.RegWrite path & "date_32", "32 date_32"
wshshell.RegWrite path & "currency_32", "32 currency_32"
wshshell.RegWrite path & "amount_32", "32 amount_32"
wshshell.RegWrite path & "currency_33B", "33 currency_33B"
wshshell.RegWrite path & "amount_33B", "33 amount_33B"
wshshell.RegWrite path & "orderingCustomer_50", "50 orderingCustomer_50"
wshshell.RegWrite path & "orderingInstitution_52", "52 orderingInstitution_52"
wshshell.RegWrite path & "senderCorrespondent_53", "53 senderCorrespondent_53"
wshshell.RegWrite path & "receiverCorrespondent_54", "54 receiverCorrespondent_54"
wshshell.RegWrite path & "intermediaryInstitution_56", "56 intermediaryInstitution_56"
wshshell.RegWrite path & "accountWithInstitution_57", "57 accountWithInstitution_57"
wshshell.RegWrite path & "beneficiaryInstitution_58", "58 beneficiaryInstitution_58"
wshshell.RegWrite path & "beneficiaryCustomer_59", "59 beneficiaryCustomer_59"
wshshell.RegWrite path & "processingCharacteristic", "00 processingCharacteristic"
wshshell.RegWrite path & "mess_direction", "00 mess_direction"
wshshell.RegWrite path & "comment", "00 comment"
wshshell.RegWrite path & "dateTime_mess", "00 dateTime_mess"
wshshell.RegWrite path & "referenceMess", "00 referenceMess"
wshshell.RegWrite path & "fin", "00 fin"
wshshell.RegWrite path & "swiftNumberBankKontragent", "00 swiftNumberBankKontragent"
wshshell.RegWrite path & "naimBankKontragent", "00 naimBankKontragent"
wshshell.RegWrite path & "thread", "00 thread"
wshshell.RegWrite path & "fileName", "00 fileName"
wshshell.RegWrite path & "direction", "00 direction"
wshshell.RegWrite path & "id", "00 id"

path = "HKEY_CURRENT_USER\Software\swift\columnAlignment\"

wshshell.RegWrite path & "transactionReferenceNumber_20", "l"
wshshell.RegWrite path & "valueDate_30V", "l"
wshshell.RegWrite path & "date_32", "l"
wshshell.RegWrite path & "currency_32", "l"
wshshell.RegWrite path & "amount_32", "l"
wshshell.RegWrite path & "currency_33B", "l"
wshshell.RegWrite path & "amount_33B", "l"
wshshell.RegWrite path & "orderingCustomer_50", "l"
wshshell.RegWrite path & "orderingInstitution_52", "l"
wshshell.RegWrite path & "senderCorrespondent_53", "l"
wshshell.RegWrite path & "receiverCorrespondent_54", "l"
wshshell.RegWrite path & "intermediaryInstitution_56", "l"
wshshell.RegWrite path & "accountWithInstitution_57", "l"
wshshell.RegWrite path & "beneficiaryInstitution_58", "l"
wshshell.RegWrite path & "beneficiaryCustomer_59", "l"
wshshell.RegWrite path & "processingCharacteristic", "l"
wshshell.RegWrite path & "mess_direction", "l"
wshshell.RegWrite path & "comment", "l"
wshshell.RegWrite path & "dateTime_mess", "l"
wshshell.RegWrite path & "referenceMess", "l"
wshshell.RegWrite path & "fin", "l"
wshshell.RegWrite path & "swiftNumberBankKontragent", "l"
wshshell.RegWrite path & "naimBankKontragent", "l"
wshshell.RegWrite path & "thread", "l"
wshshell.RegWrite path & "fileName", "l"
wshshell.RegWrite path & "direction", "l"
wshshell.RegWrite path & "id", "l"
msgBox("End")