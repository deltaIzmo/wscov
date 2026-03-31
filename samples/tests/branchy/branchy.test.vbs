' branchy.test.vbs

Class BranchyTest
  Private m_a
  Private m_b

  Public Sub setup()
    m_a = 1
    m_b = 2
  End Sub

  Public Sub test_positive_branch()
    assert_equal 24, sut.Score(1, m_a, m_b)
  End Sub

  Public Sub test_zero_branch()
    assert_equal 12, sut.Score(0, m_a, m_b)
  End Sub
End Class

Wscov_RegisterTestCase "BranchyTest"
