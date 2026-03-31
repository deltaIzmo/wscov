' calculator.test.vbs

Class CalculatorTest
  Public Sub test_add()
    assert_equal 3, sut.Add(1, 2)
  End Sub

  Public Sub test_sub()
    assert_equal 1, sut.Sub(3, 2)
  End Sub
End Class

Wscov_RegisterTestCase "CalculatorTest"
