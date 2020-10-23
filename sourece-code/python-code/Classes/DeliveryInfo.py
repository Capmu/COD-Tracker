class DeliveryInfo:

    #General attributes
    sendingDate = "<DEFAULT-SENDING-DATE>"
    deliveryCode = "DEFAULT-DELIVERY-CODE"
    expectedCOD = 0
    actualCOD = 0

    #Dinamic attribute
    customerInfoColumnNumberDict = {}
    customerInfoValueDict = {}

    def setCustomerInfoColumnNumberDict(self, columnNumberDict):
        self.customerInfoColumnNumberDict = columnNumberDict
    
    def setCustomerInfoValueDict(self, valueDict):
        self.customerInfoValueDict = valueDict

    def checkStatus(self):
        if self.expectedCOD == self.actualCOD:
            self.status = "success-payment"
        elif self.expectedCOD == 0 or self.actualCOD == 0:
            self.status = "invalid-COD"
        else:
            self.status = "invalid-COD"