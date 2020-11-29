class DeliveryInfo:

    #General attributes
    sendingDate = "DEFAULT-SENDING-DATE"
    deliveryCode = "DEFAULT-DELIVERY-CODE"
    expectedCOD = 0
    actualCOD = 0
    status = "DEFAULT-STATUS"

    #Dinamic attribute
    customerInfoDict = {}
    
    def setCustomerInfoDict(self, valueDict):
        self.customerInfoDict = valueDict

    def checkStatus(self):

        if float(self.expectedCOD) == float(self.actualCOD) and self.expectedCOD != 0:
            self.status = "success-payment"

        elif float(self.expectedCOD) == 0:
            self.status = "non-COD"

        else:
            self.status = "invalid-COD"
    
    def defaultSetting(self):
        
        self.sendingDate = "DEFAULT-SENDING-DATE"
        self.deliveryCode = "DEFAULT-DELIVERY-CODE"
        self.expectedCOD = 0
        self.actualCOD = 0
        self.status = "DEFAULT-STATUS"
