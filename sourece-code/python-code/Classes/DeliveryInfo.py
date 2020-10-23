class DeliveryInfo:

    #General attributes
    deliveryCode = "DEFAULT-DELIVERY-CODE"
    expectedCOD = 0
    actualCOD = 0

    #Dinamic attribute
    customerInfoDict = {}

    def setCustomerInfoDict(self, customerInfoDict):
        self.customerInfoDict = customerInfoDict

    def checkStatus(self):
        if self.expectedCOD == self.actualCOD:
            self.status = "success-payment"
        elif self.expectedCOD == 0 or self.actualCOD == 0:
            self.status = "invalid-COD"
        else:
            self.status = "invalid-COD"