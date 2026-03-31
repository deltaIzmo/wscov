' calculator.test.vbs (compatibility path)

Wscov_AddTest "calculator:Add", GetRef("Test_Calculator_Add")
Wscov_AddTest "calculator:Sub", GetRef("Test_Calculator_Sub")

Sub Test_Calculator_Add()
  AssertEqual 3, sut.Add(1, 2)
End Sub

Sub Test_Calculator_Sub()
  AssertEqual 1, sut.Sub(3, 2)
End Sub
