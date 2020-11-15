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

        if float(self.expectedCOD) == float(self.actualCOD):
            self.status = "success-payment"

        else:
            self.status = "invalid-COD"