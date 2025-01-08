Option Explicit On

Private Const BASE_URL As String = "http://localhost:5204" ' <-- Change to your actual base URL

Private Function HttpPostJson(url As String, jsonBody As String) As String
    'Debug.Print (jsonBody)
    'HttpPostJson = "debugHttpJson"
    Debug.Print("------------------------------JSON--------------------------------")
    Debug.Print(jsonBody)
    Debug.Print("-----------------------------/JSON--------------------------------")
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    ' Add other headers if necessary, e.g.:
    ' http.SetRequestHeader "Authorization", "Bearer your_token_here"

    http.Send jsonBody

    If http.status = 200 Or http.status = 201 Then
        HttpPostJson = http.responseText
    Else

        Debug.Print "HTTP Error: " & http.status & " - " & http.StatusText & " Response: " & http.responseText
        Err.Raise vbObjectError + 1000, "HttpPostJson", "HTTP Error: " & http.status & " - " & http.StatusText & " Response: " & http.responseText
    End If
End Function

' This assumes the response is always a single quoted GUID string, e.g.:
' "4300146c-de54-4c5c-b2f8-d83742e39eb8"
Private Function ExtractIdFromJsonResponse(jsonResponse As String) As String
    'Debug.Print ("debugExtractId")
    'ExtractIdFromJsonResponse = "testid11111111111"

    If Len(jsonResponse) > 2 And Left(jsonResponse, 1) = """" And Right(jsonResponse, 1) = """" Then
        ExtractIdFromJsonResponse = Mid(jsonResponse, 2, Len(jsonResponse) - 2)
    Else
        ' If unexpected format, return entire response
        ExtractIdFromJsonResponse = jsonResponse
    End If
End Function
Private Function ToInvariantString(value As Double) As String
    ' Converts a double to a string with '.' as decimal separator regardless of locale.
    ' We convert to a string using CStr and then replace ',' with '.' if needed.
    Dim s As String
    s = CStr(value)
    s = Replace(s, ",", ".")
    ToInvariantString = s
End Function



Public Function CreateCalculation(name As String, username As String) As String
    Dim url As String, jsonBody As String, response As String
    url = BASE_URL & "/api/Calculations"
    jsonBody = "{""name"": """ & name & """, ""username"": """ & username & """}"

    response = HttpPostJson(url, jsonBody)
    CreateCalculation = ExtractIdFromJsonResponse(response)
End Function
Public Function CreatePartition(calculationId As String, username As String, name As String, totalRMSTCost As Double, totalCost As Double,
                                generalConstructionCost As Double, laboratoryCost As Double,
                                geodesyCost As Double, generalCosts As Double,
                                profitOrRisk As Double, financingCost As Double,
                                finalUnitPriceInRequesterUnits As Double, totalValueInRequesterUnits As Double) As String

    Dim url As String, jsonBody As String, response As String
    url = BASE_URL & "/api/Partitions/" & calculationId

    jsonBody = "{" &
        """username"": """ & username & """," &
        """name"": """ & name & """," &
        """totalRMSTCost"": " & ToInvariantString(totalRMSTCost) & "," &
        """totalCost"": " & ToInvariantString(totalCost) & "," &
        """generalConstructionCost"": " & ToInvariantString(generalConstructionCost) & "," &
        """laboratoryCost"": " & ToInvariantString(laboratoryCost) & "," &
        """geodesyCost"": " & ToInvariantString(geodesyCost) & "," &
        """generalCosts"": " & ToInvariantString(generalCosts) & "," &
        """profitOrRisk"": " & ToInvariantString(profitOrRisk) & "," &
        """financingCost"": " & ToInvariantString(financingCost) & "," &
        """finalUnitPriceInRequesterUnits"": " & ToInvariantString(finalUnitPriceInRequesterUnits) & "," &
        """totalValueInRequesterUnits"": " & ToInvariantString(totalValueInRequesterUnits) &
        "}"

    response = HttpPostJson(url, jsonBody)
    CreatePartition = ExtractIdFromJsonResponse(response)
End Function

Public Function CreateSection(partitionId As String,
                              username As String, itemNo As Long,
                              description As String, unit As String,
                              quantity As Double, quantityMade As Double, quantityCorrectionFactor As Double,
                              unitConversionFactor As Double, convertedQuantity As Double,
                              totalRMSTCost As Double, subcontractorName As String, subcontractorUnitPrice As Double, subcontractorValue As Double,
                              totalCost As Double, unitCostInRequesterUnits As Double, unitCostInCalculatedUnits As Double,
                              generalConstructionCost As Double, laboratoryCost As Double, geodesyCost As Double,
                              generalCosts As Double, profitOrRisk As Double, financingCost As Double,
                              finalUnitPriceInRequesterUnits As Double, totalValueInRequesterUnits As Double,
                              quantityMadeUnit As String, status As String, Optional groupOfWork As String = "warnNoGrpAssigned") As String
    Dim url As String, jsonBody As String, response As String
    url = BASE_URL & "/api/Sections/" & partitionId

    jsonBody = "{" &
        """username"": """ & username & """," & """itemNo"": " & itemNo & "," & """description"": """ & description & """," &
        """unit"": """ & unit & """," & """quantity"": " & ToInvariantString(quantity) & "," & """quantityMade"": " & ToInvariantString(quantityMade) & "," & """quantityCorrectionFactor"": " & ToInvariantString(quantityCorrectionFactor) & "," & """unitConversionFactor"": " & ToInvariantString(unitConversionFactor) & "," &
        """convertedQuantity"": " & ToInvariantString(convertedQuantity) & "," & """totalRMSTCost"": " & ToInvariantString(totalRMSTCost) & "," & """subcontractorName"": """ & subcontractorName & """," & """subcontractorUnitPrice"": " & ToInvariantString(subcontractorUnitPrice) & "," &
        """subcontractorValue"": " & ToInvariantString(subcontractorValue) & "," & """totalCost"": " & ToInvariantString(totalCost) & "," & """unitCostInRequesterUnits"": " & ToInvariantString(unitCostInRequesterUnits) & "," & """unitCostInCalculatedUnits"": " & ToInvariantString(unitCostInCalculatedUnits) & "," &
        """generalConstructionCost"": " & ToInvariantString(generalConstructionCost) & "," & """laboratoryCost"": " & ToInvariantString(laboratoryCost) & "," & """geodesyCost"": " & ToInvariantString(geodesyCost) & "," & """generalCosts"": " & ToInvariantString(generalCosts) & "," &
        """profitOrRisk"": " & ToInvariantString(profitOrRisk) & "," & """financingCost"": " & ToInvariantString(financingCost) & "," & """finalUnitPriceInRequesterUnits"": " & ToInvariantString(finalUnitPriceInRequesterUnits) & "," & """totalValueInRequesterUnits"": " & ToInvariantString(totalValueInRequesterUnits) & "," &
        """quantityMadeUnit"": """ & quantityMadeUnit & """," & """status"": """ & status & "" & """," & """groupOfWork"": """ & groupOfWork & """" &
        "}"

    response = HttpPostJson(url, jsonBody)
    CreateSection = ExtractIdFromJsonResponse(response)
End Function

Public Function CreateSubSection(sectionId As String,
                                 username As String, efficiency As Double,
                                 totalRMSTCost As Double, totalCost As Double,
                                 unitCostInRequesterUnits As Double, unitCostInCalculatedUnits As Double,
                                 Optional subsectionCode As String = "",
                                 Optional subsectionName As String = "",
                                 Optional efficiencyUnit As String = "", Optional groupOfWork As String = "warnNoGrpAssigned") As String
    Dim url As String, jsonBody As String, response As String
    url = BASE_URL & "/api/SubSections/" & sectionId

    Dim codeJson As String
    Dim nameJson As String
    Dim effUnitJson As String

    If subsectionCode = "" Then
        codeJson = """subsectionCode"": null,"
    Else
        codeJson = """subsectionCode"": """ & subsectionCode & ""","
    End If

    If subsectionName = "" Then
        nameJson = """subsectionName"": null,"
    Else
        nameJson = """subsectionName"": """ & subsectionName & ""","
    End If

    If efficiencyUnit = "" Then
        effUnitJson = """efficiencyUnit"": null,"
    Else
        effUnitJson = """efficiencyUnit"": """ & efficiencyUnit & ""","
    End If

    jsonBody = "{" &
        """username"": """ & username & """," &
         codeJson &
         nameJson &
        """efficiency"": " & ToInvariantString(efficiency) & "," &
         effUnitJson &
        """totalRMSTCost"": " & ToInvariantString(totalRMSTCost) & "," &
        """totalCost"": " & ToInvariantString(totalCost) & "," &
        """unitCostInRequesterUnits"": " & ToInvariantString(unitCostInRequesterUnits) & "," &
        """unitCostInCalculatedUnits"": " & ToInvariantString(unitCostInCalculatedUnits) & "," & """groupOfWork"": """ & groupOfWork & """" &
        "}"

    response = HttpPostJson(url, jsonBody)
    CreateSubSection = ExtractIdFromJsonResponse(response)
End Function



Public Function CreateSubSubSection(subSectionId As String,
                                    username As String, typeValue As Long,
                                    totalRMSTCost As Double, totalCost As Double,
                                    unitCostInRequesterUnits As Double, unitCostInCalculatedUnits As Double,
                                    Optional subSubSectionName As String = "", Optional groupOfWork As String = "warnNoGrpAssigned") As String

    Dim url As String, jsonBody As String, response As String
    url = BASE_URL & "/api/SubSubSections/" & subSectionId

    Dim nameJson As String
    If subSubSectionName = "" Then
        nameJson = """subSubSectionName"": null,"
    Else
        nameJson = """subSubSectionName"": """ & subSubSectionName & ""","
    End If

    jsonBody = "{" &
        """username"": """ & username & """," &
         nameJson &
        """type"": " & typeValue & "," &
        """totalRMSTCost"": " & ToInvariantString(totalRMSTCost) & "," &
        """totalCost"": " & ToInvariantString(totalCost) & "," &
        """unitCostInRequesterUnits"": " & ToInvariantString(unitCostInRequesterUnits) & "," &
        """unitCostInCalculatedUnits"": " & ToInvariantString(unitCostInCalculatedUnits) & "," & """groupOfWork"": """ & groupOfWork & """" &
        "}"
    'Debug.Print jsonBody
    response = HttpPostJson(url, jsonBody)
    CreateSubSubSection = ExtractIdFromJsonResponse(response)
End Function


Public Function CreateSubSubSectionItem(subSubSectionId As String,
                                        username As String,
                                        resourceQuantityForGivenEfficiency As Double, resourceQuantityTORMSTQuanity As Double,
                                        resourceUnitPrice As Double, totalResourceCost As Double,
                                        Optional resourceCode As String = "", Optional unit As String = "", Optional resourceName As String = "", Optional groupOfWork As String = "warnNoGrpAssigned") As String

    Dim url As String, jsonBody As String, response As String
    url = BASE_URL & "/api/SubSubSectionItems/" & subSubSectionId

    Dim codeJson As String, unitJson As String, nameJson As String

    If resourceCode = "" Then
        codeJson = """resourceCode"": null,"
    Else
        codeJson = """resourceCode"": """ & resourceCode & ""","
    End If

    If unit = "" Then
        unitJson = """unit"": null,"
    Else
        unitJson = """unit"": """ & unit & ""","
    End If

    If resourceName = "" Then
        nameJson = """resourceName"": null,"
    Else
        nameJson = """resourceName"": """ & resourceName & ""","
    End If

    jsonBody = "{" &
        """username"": """ & username & """," &
         codeJson &
         nameJson &
         unitJson &
        """resourceQuantityForGivenEfficiency"": " & ToInvariantString(resourceQuantityForGivenEfficiency) & "," &
        """resourceQuantityTORMSTQuanity"": " & ToInvariantString(resourceQuantityTORMSTQuanity) & "," &
        """resourceUnitPrice"": " & ToInvariantString(resourceUnitPrice) & "," &
        """totalResourceCost"": " & ToInvariantString(totalResourceCost) & "," & """groupOfWork"": """ & groupOfWork & """" &
        "}"

    response = HttpPostJson(url, jsonBody)
    CreateSubSubSectionItem = ExtractIdFromJsonResponse(response)
End Function





Public Function CreateNewVersion(calculationId As String,
                                 previousVersionId As String, description As String, userId As String) As String
    Dim url As String, jsonBody As String, response As String
    url = BASE_URL & "/api/Versions/" & calculationId & "/new-version"

    jsonBody = "{" &
        """previousVersionId"": """ & previousVersionId & """," &
        """description"": """ & description & """," &
        """userId"": """ & userId & """" &
        "}"

    response = HttpPostJson(url, jsonBody)
    CreateNewVersion = ExtractIdFromJsonResponse(response)
End Function





'Option Explicit
'
'Private Const BASE_URL As String = "http://localhost:5204" ' <-- Change to your actual base URL
'
'Private Function HttpPostJson(url As String, jsonBody As String) As String
'    Dim http As Object
'    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
'
'    http.Open "POST", url, False
'    http.SetRequestHeader "Content-Type", "application/json"
'    ' Add other headers if necessary, e.g.:
'    ' http.SetRequestHeader "Authorization", "Bearer your_token_here"
'
'    http.Send jsonBody
'
'    If http.status = 200 Or http.status = 201 Then
'        HttpPostJson = http.responseText
'    Else
'        Debug.Print "HTTP Error: " & http.status & " - " & http.StatusText & " Response: " & http.responseText
'        Err.Raise vbObjectError + 1000, "HttpPostJson", "HTTP Error: " & http.status & " - " & http.StatusText & " Response: " & http.responseText
'    End If
'End Function
'
'' This assumes the response is always a single quoted GUID string, e.g.:
'' "4300146c-de54-4c5c-b2f8-d83742e39eb8"
'Private Function ExtractIdFromJsonResponse(jsonResponse As String) As String
'    If Len(jsonResponse) > 2 And Left(jsonResponse, 1) = """" And Right(jsonResponse, 1) = """" Then
'        ExtractIdFromJsonResponse = Mid(jsonResponse, 2, Len(jsonResponse) - 2)
'    Else
'        ' If unexpected format, return entire response
'        ExtractIdFromJsonResponse = jsonResponse
'    End If
'End Function
'Private Function ToInvariantString(value As Double) As String
'    ' Converts a double to a string with '.' as decimal separator regardless of locale.
'    ' We convert to a string using CStr and then replace ',' with '.' if needed.
'    Dim s As String
'    s = CStr(value)
'    s = Replace(s, ",", ".")
'    ToInvariantString = s
'End Function
'
'
'
'Public Function CreateCalculation(name As String, username As String) As String
'    Dim url As String, jsonBody As String, response As String
'    url = BASE_URL & "/api/Calculations"
'     jsonBody = "{""name"": """ & name & """, ""username"": """ & username & """}"
'
'    response = HttpPostJson(url, jsonBody)
'    CreateCalculation = ExtractIdFromJsonResponse(response)
'End Function
'Public Function CreatePartition(calculationId As String, username As String, name As String, totalRMSTCost As Double, totalCost As Double, _
'                                generalConstructionCost As Double, laboratoryCost As Double, _
'                                geodesyCost As Double, generalCosts As Double, _
'                                profitOrRisk As Double, financingCost As Double, _
'                                finalUnitPriceInRequesterUnits As Double, totalValueInRequesterUnits As Double) As String
'
'    Dim url As String, jsonBody As String, response As String
'    url = BASE_URL & "/api/Partitions/" & calculationId
'
'    jsonBody = "{" & _
'        """username"": """ & username & """," & _
'        """name"": """ & name & """," & _
'        """totalRMSTCost"": " & ToInvariantString(totalRMSTCost) & "," & _
'        """totalCost"": " & ToInvariantString(totalCost) & "," & _
'        """generalConstructionCost"": " & ToInvariantString(generalConstructionCost) & "," & _
'        """laboratoryCost"": " & ToInvariantString(laboratoryCost) & "," & _
'        """geodesyCost"": " & ToInvariantString(geodesyCost) & "," & _
'        """generalCosts"": " & ToInvariantString(generalCosts) & "," & _
'        """profitOrRisk"": " & ToInvariantString(profitOrRisk) & "," & _
'        """financingCost"": " & ToInvariantString(financingCost) & "," & _
'        """finalUnitPriceInRequesterUnits"": " & ToInvariantString(finalUnitPriceInRequesterUnits) & "," & _
'        """totalValueInRequesterUnits"": " & ToInvariantString(totalValueInRequesterUnits) & _
'        "}"
'
'    response = HttpPostJson(url, jsonBody)
'    CreatePartition = ExtractIdFromJsonResponse(response)
'End Function
'
'Public Function CreateSection(partitionId As String, _
'                              username As String, itemNo As Long, _
'                              description As String, unit As String, _
'                              quantity As Double, quantityMade As Double, quantityCorrectionFactor As Double, _
'                              unitConversionFactor As Double, convertedQuantity As Double, _
'                              totalRMSTCost As Double, subcontractorName As String, subcontractorUnitPrice As Double, subcontractorValue As Double, _
'                              totalCost As Double, unitCostInRequesterUnits As Double, unitCostInCalculatedUnits As Double, _
'                              generalConstructionCost As Double, laboratoryCost As Double, geodesyCost As Double, _
'                              generalCosts As Double, profitOrRisk As Double, financingCost As Double, _
'                              finalUnitPriceInRequesterUnits As Double, totalValueInRequesterUnits As Double, _
'                              quantityMadeUnit As String, status As String, Optional groupOfWork As String = "BRAK") As String
'    Dim url As String, jsonBody As String, response As String
'    url = BASE_URL & "/api/Sections/" & partitionId
'
'    jsonBody = "{" & _
'        """username"": """ & username & """," & """itemNo"": " & itemNo & "," & """description"": """ & description & """," & _
'        """unit"": """ & unit & """," & """quantity"": " & ToInvariantString(quantity) & "," & """quantityMade"": " & ToInvariantString(quantityMade) & "," & """quantityCorrectionFactor"": " & ToInvariantString(quantityCorrectionFactor) & "," & """unitConversionFactor"": " & ToInvariantString(unitConversionFactor) & "," & _
'        """convertedQuantity"": " & ToInvariantString(convertedQuantity) & "," & """totalRMSTCost"": " & ToInvariantString(totalRMSTCost) & "," & """subcontractorName"": """ & subcontractorName & """," & """subcontractorUnitPrice"": " & ToInvariantString(subcontractorUnitPrice) & "," & _
'        """subcontractorValue"": " & ToInvariantString(subcontractorValue) & "," & """totalCost"": " & ToInvariantString(totalCost) & "," & """unitCostInRequesterUnits"": " & ToInvariantString(unitCostInRequesterUnits) & "," & """unitCostInCalculatedUnits"": " & ToInvariantString(unitCostInCalculatedUnits) & "," & _
'        """generalConstructionCost"": " & ToInvariantString(generalConstructionCost) & "," & """laboratoryCost"": " & ToInvariantString(laboratoryCost) & "," & """geodesyCost"": " & ToInvariantString(geodesyCost) & "," & """generalCosts"": " & ToInvariantString(generalCosts) & "," & _
'        """profitOrRisk"": " & ToInvariantString(profitOrRisk) & "," & """financingCost"": " & ToInvariantString(financingCost) & "," & """finalUnitPriceInRequesterUnits"": " & ToInvariantString(finalUnitPriceInRequesterUnits) & "," & """totalValueInRequesterUnits"": " & ToInvariantString(totalValueInRequesterUnits) & "," & _
'        """quantityMadeUnit"": """ & quantityMadeUnit & """," & """status"": """ & status & """" & """," & """groupOfWork"": """ & groupOfWork & """" & _
'        "}"
'
'    response = HttpPostJson(url, jsonBody)
'    CreateSection = ExtractIdFromJsonResponse(response)
'End Function
'
'Public Function CreateSubSection(sectionId As String, _
'                                 username As String, efficiency As Double, _
'                                 totalRMSTCost As Double, totalCost As Double, _
'                                 unitCostInRequesterUnits As Double, unitCostInCalculatedUnits As Double, _
'                                 Optional subsectionCode As String = "", _
'                                 Optional subsectionName As String = "", _
'                                 Optional efficiencyUnit As String = "", Optional groupOfWork As String = "BRAK") As String
'    Dim url As String, jsonBody As String, response As String
'    url = BASE_URL & "/api/SubSections/" & sectionId
'
'    Dim codeJson As String
'    Dim nameJson As String
'    Dim effUnitJson As String
'
'    If subsectionCode = "" Then
'        codeJson = """subsectionCode"": null,"
'    Else
'        codeJson = """subsectionCode"": """ & subsectionCode & ""","
'    End If
'
'    If subsectionName = "" Then
'        nameJson = """subsectionName"": null,"
'    Else
'        nameJson = """subsectionName"": """ & subsectionName & ""","
'    End If
'
'    If efficiencyUnit = "" Then
'        effUnitJson = """efficiencyUnit"": null,"
'    Else
'        effUnitJson = """efficiencyUnit"": """ & efficiencyUnit & ""","
'    End If
'
'    jsonBody = "{" & _
'        """username"": """ & username & """," & _
'         codeJson & _
'         nameJson & _
'        """efficiency"": " & ToInvariantString(efficiency) & "," & _
'         effUnitJson & _
'        """totalRMSTCost"": " & ToInvariantString(totalRMSTCost) & "," & _
'        """totalCost"": " & ToInvariantString(totalCost) & "," & _
'        """unitCostInRequesterUnits"": " & ToInvariantString(unitCostInRequesterUnits) & "," & _
'        """unitCostInCalculatedUnits"": " & ToInvariantString(unitCostInCalculatedUnits) & _
'        "}"
'
'    response = HttpPostJson(url, jsonBody)
'    CreateSubSection = ExtractIdFromJsonResponse(response)
'End Function
'
'
'
'Public Function CreateSubSubSection(subSectionId As String, _
'                                    username As String, typeValue As Long, _
'                                    totalRMSTCost As Double, totalCost As Double, _
'                                    unitCostInRequesterUnits As Double, unitCostInCalculatedUnits As Double, _
'                                    Optional subSubSectionName As String = "", Optional groupOfWork As String = "BRAK") As String
'
'    Dim url As String, jsonBody As String, response As String
'    url = BASE_URL & "/api/SubSubSections/" & subSectionId
'
'    Dim nameJson As String
'    If subSubSectionName = "" Then
'        nameJson = """subSubSectionName"": null,"
'    Else
'        nameJson = """subSubSectionName"": """ & subSubSectionName & ""","
'    End If
'
'    jsonBody = "{" & _
'        """username"": """ & username & """," & _
'         nameJson & _
'        """type"": " & typeValue & "," & _
'        """totalRMSTCost"": " & ToInvariantString(totalRMSTCost) & "," & _
'        """totalCost"": " & ToInvariantString(totalCost) & "," & _
'        """unitCostInRequesterUnits"": " & ToInvariantString(unitCostInRequesterUnits) & "," & _
'        """unitCostInCalculatedUnits"": " & ToInvariantString(unitCostInCalculatedUnits) & _
'        "}"
'    Debug.Print jsonBody
'    response = HttpPostJson(url, jsonBody)
'    CreateSubSubSection = ExtractIdFromJsonResponse(response)
'End Function
'
'
'Public Function CreateSubSubSectionItem(subSubSectionId As String, _
'                                        username As String, _
'                                        resourceQuantityForGivenEfficiency As Double, resourceQuantityTORMSTQuanity As Double, _
'                                        resourceUnitPrice As Double, totalResourceCost As Double, _
'                                        Optional resourceCode As String = "", Optional unit As String = "", Optional resourceName As String = "", Optional groupOfWork As String = "BRAK") As String
'
'    Dim url As String, jsonBody As String, response As String
'    url = BASE_URL & "/api/SubSubSectionItems/" & subSubSectionId
'
'    Dim codeJson As String, unitJson As String, nameJson As String
'
'    If resourceCode = "" Then
'        codeJson = """resourceCode"": null,"
'    Else
'        codeJson = """resourceCode"": """ & resourceCode & ""","
'    End If
'
'    If unit = "" Then
'        unitJson = """unit"": null,"
'    Else
'        unitJson = """unit"": """ & unit & ""","
'    End If
'
'    If resourceName = "" Then
'        nameJson = """resourceName"": null,"
'    Else
'        nameJson = """resourceName"": """ & resourceName & ""","
'    End If
'
'    jsonBody = "{" & _
'        """username"": """ & username & """," & _
'         codeJson & _
'         nameJson & _
'         unitJson & _
'        """resourceQuantityForGivenEfficiency"": " & ToInvariantString(resourceQuantityForGivenEfficiency) & "," & _
'        """resourceQuantityTORMSTQuanity"": " & ToInvariantString(resourceQuantityTORMSTQuanity) & "," & _
'        """resourceUnitPrice"": " & ToInvariantString(resourceUnitPrice) & "," & _
'        """totalResourceCost"": " & ToInvariantString(totalResourceCost) & _
'        "}"
'
'    response = HttpPostJson(url, jsonBody)
'    CreateSubSubSectionItem = ExtractIdFromJsonResponse(response)
'End Function
'
'
'
'
'
'Public Function CreateNewVersion(calculationId As String, _
'                                 previousVersionId As String, description As String, userId As String) As String
'    Dim url As String, jsonBody As String, response As String
'    url = BASE_URL & "/api/Versions/" & calculationId & "/new-version"
'
'    jsonBody = "{" & _
'        """previousVersionId"": """ & previousVersionId & """," & _
'        """description"": """ & description & """," & _
'        """userId"": """ & userId & """" & _
'        "}"
'
'    response = HttpPostJson(url, jsonBody)
'    CreateNewVersion = ExtractIdFromJsonResponse(response)
'End Function
'
'
'
