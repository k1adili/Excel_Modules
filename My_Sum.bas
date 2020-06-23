Attribute VB_Name = "My_Sum"
Public Function mysum(Number As Double, current_sum As Double, My_sum As Double)
    mysum = Number - ((Number * 100 / current_sum) / 100) * (current_sum - My_sum)
End Function
