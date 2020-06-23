Attribute VB_Name = "Mianyabi"
Public Function mianyabi(size1 As Double, size2 As Double, price1 As Double, price2 As Double, size3 As Double)
    mianyabi = (((size3 - size1) * (price2 - price1)) / (size2 - size1)) + price1
End Function
