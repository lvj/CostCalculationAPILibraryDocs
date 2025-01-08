Option Explicit

' -------------------------------------------
' Test for CreateCalculation
' -------------------------------------------
Sub Test_CreateCalculation()
    Dim calcId As String
    calcId = CreateCalculation("Test Calculation Name", "string")
    Debug.Print "Created Calculation ID: " & calcId
End Sub

' -------------------------------------------
' Test for CreatePartition
' -------------------------------------------
Sub Test_CreatePartition()
    Dim calculationId As String
    calculationId = CreateCalculation("Partition Test Calc", "string")
    Debug.Print "Created Calculation for Partition Test: " & calculationId
    
    Dim partitionId As String
    partitionId = CreatePartition(calculationId, _
                                  "string", _
                                   "Test Partition", _
                                  1000#, 1200#, 100#, 50#, _
                                  20#, 80#, 30#, 10#, _
                                  15#, 1800#)
    Debug.Print "Created Partition ID: " & partitionId
End Sub

' -------------------------------------------
' Test for CreateSection
' Now includes the required fields: subcontractorName, quantityMadeUnit, status
' -------------------------------------------
Sub Test_CreateSection()
    Dim calcId As String, partitionId As String, sectionId As String
    
    ' Create a Calculation
    calcId = CreateCalculation("Section Test Calc", "string")
    
    ' Create a Partition
    partitionId = CreatePartition(calcId, "string", "Partition for Section", _
                                  1000#, 1200#, 100#, 50#, _
                                  20#, 80#, 30#, 10#, _
                                  15#, 1800#)
                                  
    ' Create a Section with all required fields
    sectionId = CreateSection(partitionId, "string", 100, _
                              "Section Description", "m3", _
                              50#, 10#, 1#, _
                              1#, 50#, _
                              500#, "Subcontractor Inc.", 10#, 500#, _
                              600#, 12#, 12#, _
                              20#, 10#, 5#, _
                              15#, 5#, 2#, _
                              13#, 650#, _
                              "m3", "Active")

    Debug.Print "Created Section ID: " & sectionId
End Sub
Sub Test_CreateSubSection()
    Dim calcId As String, partitionId As String, sectionId As String, subSectionId As String
    
    calcId = CreateCalculation("SubSection Test Calc", "string")
    partitionId = CreatePartition(calcId, "string", "Partition for SubSection", _
                                  1000#, 1200#, 100#, 50#, _
                                  20#, 80#, 30#, 10#, _
                                  15#, 1800#)
    
    ' Updated CreateSection with required fields
    sectionId = CreateSection(partitionId, "string", 100, _
                              "Section Description", "m3", _
                              50#, 10#, 1#, _
                              1#, 50#, _
                              500#, "Subcontractor Inc.", 10#, 500#, _
                              600#, 12#, 12#, _
                              20#, 10#, 5#, _
                              15#, 5#, 2#, _
                              13#, 650#, _
                              "m3", "Active")
        subSectionId = CreateSubSection(sectionId, "string", _
                                    1, 0.9, _
                                    300#, 350#, _
                                    7#, "ss code", "ssname", "efunit")
    Debug.Print "Created SubSection ID: " & subSectionId
    ' Now create a SubSection

End Sub

Sub Test_CreateSubSubSection()
    Dim calcId As String, partitionId As String, sectionId As String, subSectionId As String, subSubSectionId As String
    
    calcId = CreateCalculation("SubSubSection Test Calc", "string")
    partitionId = CreatePartition(calcId, "string", "Partition for SubSubSection", _
                                  1000#, 1200#, 100#, 50#, _
                                  20#, 80#, 30#, 10#, _
                                  15#, 1800#)
    
    sectionId = CreateSection(partitionId, "string", 100, _
                              "Section Description", "m3", _
                              50#, 10#, 1#, _
                              1#, 50#, _
                              500#, "Subcontractor Inc.", 10#, 500#, _
                              600#, 12#, 12#, _
                              20#, 10#, 5#, _
                              15#, 5#, 2#, _
                              13#, 650#, _
                              "m3", "Active")
                              
        subSectionId = CreateSubSection(sectionId, "string", _
                                    1, 0.9, _
                                    300#, 350#, _
                                    7#, "ss code2", "ssname2", "efunit2")
                                    
    subSubSectionId = CreateSubSubSection(subSectionId, _
                                          "string", _
                                          1, 0, _
                                          100#, 120#, _
                                          1.2, 1.2)
    Debug.Print "Created SubSubSection ID: " & subSubSectionId
End Sub

Sub Test_CreateSubSubSectionItem()
    Dim calcId As String, partitionId As String, sectionId As String
    Dim subSectionId As String, subSubSectionId As String, subSubSectionItemId As String
    
    calcId = CreateCalculation("SubSubSectionItem Test Calc", "string")
    partitionId = CreatePartition(calcId, "string", "Partition for SubSubSectionItem", _
                                  1000#, 1200#, 100#, 50#, _
                                  20#, 80#, 30#, 10#, _
                                  15#, 1800#)
    
    sectionId = CreateSection(partitionId, "string", 100, _
                              "Section DescriptionItem", "m3", _
                              50#, 10#, 1#, _
                              1#, 50#, _
                              500#, "Subcontractor Inc.", 10#, 500#, _
                              600#, 12#, 12#, _
                              20#, 10#, 5#, _
                              15#, 5#, 2#, _
                              13#, 650#, _
                              "m3", "Active")
                              
        subSectionId = CreateSubSection(sectionId, "string", _
                                    1, 0.9, _
                                    300#, 350#, _
                                    7#, "ss code2", "ssname2", "efunit2")
                                    
    subSubSectionId = CreateSubSubSection(subSectionId, _
                                          "string", _
                                          1, 0, _
                                          100#, 120#, _
                                          1.2, 1.2)
                                          
    subSubSectionItemId = CreateSubSubSectionItem(subSubSectionId, _
                                                  "string", _
                                                  10#, 10#, _
                                                  5#, 50#, "c", "jm", "ssss")
    Debug.Print "Created SubSubSectionItem ID: " & subSubSectionItemId
End Sub


Sub Test_All()
    Dim calcId As String, partitionId As String, sectionId As String
    Dim subSectionId As String, subSubSectionId As String, subSubSectionItemId As String
    
    calcId = CreateCalculation("SubSubSectionItem Test Calc", "string")
    
    partitionId = CreatePartition(calcId, "string", "Partition for SubSubSectionItem", 1000#, 1200#, 100#, 50#, 0#, 80#, 30#, 10#, 15#, 1800#)
    
    sectionId = CreateSection(partitionId, "string", 100, "Section DescriptionItem", "m3", 50#, 10#, 1#, 1#, 50#, 500#, "Subcontractor Inc.", 10#, 500#, 600#, 12#, 12#, 20#, 10#, 5#, 15#, 5#, 2#, 13#, 650#, "m3", "Active")
                              
    subSectionId = CreateSubSection(sectionId, "string", 1, 0.9, 300#, 350#, 7#, "ss code2", "ssname2", "efunit2")
    subSubSectionId = CreateSubSubSection(subSectionId, "string", 1, 0, 100#, 120#, 1.2, 1.2)
  subSubSectionItemId = CreateSubSubSectionItem(subSubSectionId, "string", 10#, 10#, 5#, 50#, "c", "jm", "ssss")
 
End Sub
