Option Explicit On

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
    partitionId = CreatePartition(calculationId,
                                  "string",
                                   "Test Partition",
                                  1000.0#, 1200.0#, 100.0#, 50.0#,
                                  20.0#, 80.0#, 30.0#, 10.0#,
                                  15.0#, 1800.0#)
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
    partitionId = CreatePartition(calcId, "string", "Partition for Section",
                                  1000.0#, 1200.0#, 100.0#, 50.0#,
                                  20.0#, 80.0#, 30.0#, 10.0#,
                                  15.0#, 1800.0#)

    ' Create a Section with all required fields
    sectionId = CreateSection(partitionId, "string", 100,
                              "Section Description", "m3",
                              50.0#, 10.0#, 1.0#,
                              1.0#, 50.0#,
                              500.0#, "Subcontractor Inc.", 10.0#, 500.0#,
                              600.0#, 12.0#, 12.0#,
                              20.0#, 10.0#, 5.0#,
                              15.0#, 5.0#, 2.0#,
                              13.0#, 650.0#,
                              "m3", "Active")

    Debug.Print "Created Section ID: " & sectionId
End Sub
Sub Test_CreateSubSection()
    Dim calcId As String, partitionId As String, sectionId As String, subSectionId As String

    calcId = CreateCalculation("SubSection Test Calc", "string")
    partitionId = CreatePartition(calcId, "string", "Partition for SubSection",
                                  1000.0#, 1200.0#, 100.0#, 50.0#,
                                  20.0#, 80.0#, 30.0#, 10.0#,
                                  15.0#, 1800.0#)

    ' Updated CreateSection with required fields
    sectionId = CreateSection(partitionId, "string", 100,
                              "Section Description", "m3",
                              50.0#, 10.0#, 1.0#,
                              1.0#, 50.0#,
                              500.0#, "Subcontractor Inc.", 10.0#, 500.0#,
                              600.0#, 12.0#, 12.0#,
                              20.0#, 10.0#, 5.0#,
                              15.0#, 5.0#, 2.0#,
                              13.0#, 650.0#,
                              "m3", "Active")
    subSectionId = CreateSubSection(sectionId, "string",
                                1, 0.9,
                                300.0#, 350.0#,
                                7.0#, "ss code", "ssname", "efunit")
    Debug.Print "Created SubSection ID: " & subSectionId
    ' Now create a SubSection

End Sub

Sub Test_CreateSubSubSection()
    Dim calcId As String, partitionId As String, sectionId As String, subSectionId As String, subSubSectionId As String

    calcId = CreateCalculation("SubSubSection Test Calc", "string")
    partitionId = CreatePartition(calcId, "string", "Partition for SubSubSection",
                                  1000.0#, 1200.0#, 100.0#, 50.0#,
                                  20.0#, 80.0#, 30.0#, 10.0#,
                                  15.0#, 1800.0#)

    sectionId = CreateSection(partitionId, "string", 100,
                              "Section Description", "m3",
                              50.0#, 10.0#, 1.0#,
                              1.0#, 50.0#,
                              500.0#, "Subcontractor Inc.", 10.0#, 500.0#,
                              600.0#, 12.0#, 12.0#,
                              20.0#, 10.0#, 5.0#,
                              15.0#, 5.0#, 2.0#,
                              13.0#, 650.0#,
                              "m3", "Active")

    subSectionId = CreateSubSection(sectionId, "string",
                                1, 0.9,
                                300.0#, 350.0#,
                                7.0#, "ss code2", "ssname2", "efunit2")

    subSubSectionId = CreateSubSubSection(subSectionId,
                                          "string",
                                          1, 0,
                                          100.0#, 120.0#,
                                          1.2, 1.2)
    Debug.Print "Created SubSubSection ID: " & subSubSectionId
End Sub

Sub Test_CreateSubSubSectionItem()
    Dim calcId As String, partitionId As String, sectionId As String
    Dim subSectionId As String, subSubSectionId As String, subSubSectionItemId As String

    calcId = CreateCalculation("SubSubSectionItem Test Calc", "string")
    partitionId = CreatePartition(calcId, "string", "Partition for SubSubSectionItem",
                                  1000.0#, 1200.0#, 100.0#, 50.0#,
                                  20.0#, 80.0#, 30.0#, 10.0#,
                                  15.0#, 1800.0#)

    sectionId = CreateSection(partitionId, "string", 100,
                              "Section DescriptionItem", "m3",
                              50.0#, 10.0#, 1.0#,
                              1.0#, 50.0#,
                              500.0#, "Subcontractor Inc.", 10.0#, 500.0#,
                              600.0#, 12.0#, 12.0#,
                              20.0#, 10.0#, 5.0#,
                              15.0#, 5.0#, 2.0#,
                              13.0#, 650.0#,
                              "m3", "Active")

    subSectionId = CreateSubSection(sectionId, "string",
                                1, 0.9,
                                300.0#, 350.0#,
                                7.0#, "ss code2", "ssname2", "efunit2")

    subSubSectionId = CreateSubSubSection(subSectionId,
                                          "string",
                                          1, 0,
                                          100.0#, 120.0#,
                                          1.2, 1.2)

    subSubSectionItemId = CreateSubSubSectionItem(subSubSectionId,
                                                  "string",
                                                  10.0#, 10.0#,
                                                  5.0#, 50.0#, "c", "jm", "ssss")
    Debug.Print "Created SubSubSectionItem ID: " & subSubSectionItemId
End Sub


Sub Test_All()
    Dim calcId As String, partitionId As String, sectionId As String
    Dim subSectionId As String, subSubSectionId As String, subSubSectionItemId As String

    calcId = CreateCalculation("SubSubSectionItem Test Calc", "string")

    partitionId = CreatePartition(calcId, "string", "Partition for SubSubSectionItem", 1000.0#, 1200.0#, 100.0#, 50.0#, 0#, 80.0#, 30.0#, 10.0#, 15.0#, 1800.0#)

    sectionId = CreateSection(partitionId, "string", 100, "Section DescriptionItem", "m3", 50.0#, 10.0#, 1.0#, 1.0#, 50.0#, 500.0#, "Subcontractor Inc.", 10.0#, 500.0#, 600.0#, 12.0#, 12.0#, 20.0#, 10.0#, 5.0#, 15.0#, 5.0#, 2.0#, 13.0#, 650.0#, "m3", "Active")

    subSectionId = CreateSubSection(sectionId, "string", 1, 0.9, 300.0#, 350.0#, 7.0#, "ss code2", "ssname2", "efunit2")

    subSubSectionId = CreateSubSubSection(subSectionId, "string", 1, 0, 100.0#, 120.0#, 1.2, 1.2)

    subSubSectionItemId = CreateSubSubSectionItem(subSubSectionId, "string", 10.0#, 10.0#, 5.0#, 50.0#, "c", "jm", "ssss")

End Sub
