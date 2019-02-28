Attribute VB_Name = "Palladyne"
' All of the functions and sub-routines in this module are specific to Palladyne
Option Base 1
Option Explicit

Public Function PalladyneRegions() As Variant
    Let PalladyneRegions = Array("USA", "EUR", "RW")
End Function
    
Public Function PalladyneSectors() As Variant
    Let PalladyneSectors = Array("ENR", "UTL", "MAT", "TCH", "CAP", "DUR", "STP", "HTH", "SER", "TEL", "FIN", "BNK")
End Function
    
Public Function PalladyneStyles() As Variant
    Let PalladyneStyles = Array("G", "V")
End Function
    
Public Function PalladyneSizes() As Variant
    Let PalladyneSizes = Array("L", "S")
End Function

Public Function MsciSubIndustryNames() As Variant
    Let MsciSubIndustryNames = Array("Energy", "Materials", "Capital Goods", "Commercial & Professional Services", "Transportation", _
                                     "Automobile & Components", "Consumer Durables & Apparel", "Consumer Services", "Media", _
                                     "Retailing", "Food & Staples Retailing", "Food, Beverage & Tobacco", "Household & Personal Products", _
                                     "Health Care Equipment & Services", "Pharmaceuticals & Biotechnology & Life Sciences", "Banks", "Diversified Financials", _
                                     "Insurance", "Real Estate", "Software & Services", "Technology Hardware & Equipment", "Semiconductors & Semiconductor Equipment", _
                                     "Telecommunication Services", "Utilities")
End Function

' This function returns an array with all possible cellular coordinates
Public Function GetCellularCoordinateArray() As Variant
    Dim Sectors As Variant
    Dim Regions As Variant
    Dim Styles As Variant
    Dim Sizes As Variant
    Dim Sector As Variant
    Dim Region As Variant
    Dim Style As Variant
    Dim Size As Variant
    Dim TheCoordinateArray(144) As String
    Dim i As Integer

    ' Set regions, sectors, styles, and sizes arrays
    Let Regions = PalladyneRegions()
    Let Sectors = PalladyneSectors()
    Let Styles = PalladyneStyles()
    Let Sizes = PalladyneSizes()

    ' Create array of all coordinate tuples and store the weight of each cell
    Let i = 1
    For Each Sector In Sectors
        For Each Region In Regions
            For Each Size In Sizes
                For Each Style In Styles
                    ' Create the string representing the current coordinate tuple
                    Let TheCoordinateArray(i) = Sector & "-" & Region & "-" & Size & "-" & Style
                    Let i = i + 1
                Next Style
            Next Size
        Next Region
    Next Sector
    
    ' Return the coordinate array
    Let GetCellularCoordinateArray = TheCoordinateArray
End Function

' Returns an alpha-to-numeric mapping of Palladyne's regional codes
Public Function RegionalNumericalMap() As Dictionary
    Dim aDictionary As Dictionary
    Dim TheRegions As Variant
    Dim i As Integer
    
    Set aDictionary = New Dictionary
    Let TheRegions = PalladyneRegions()
    
    For i = 1 To UBound(TheRegions)
        aDictionary.Add TheRegions(i), i
    Next i

    Set RegionalNumericalMap = aDictionary
End Function

' Returns an alpha-to-numeric mapping of Palladyne's size codes
Public Function SizeNumericalMap() As Dictionary
    Dim aDictionary As Dictionary
    Dim TheSizes As Variant
    Dim i As Integer
    
    Set aDictionary = New Dictionary
    Let TheSizes = PalladyneSizes()
    
    For i = 1 To UBound(TheSizes)
        aDictionary.Add TheSizes(i), i
    Next i

    Set SizeNumericalMap = aDictionary
End Function

' Returns an alpha-to-numeric mapping of Palladyne's style codes
Public Function StyleNumericalMap() As Dictionary
    Dim aDictionary As Dictionary
    Dim TheStyles As Variant
    Dim i As Integer
    
    Set aDictionary = New Dictionary
    Let TheStyles = PalladyneStyles()
    
    For i = 1 To UBound(TheStyles)
        aDictionary.Add TheStyles(i), i
    Next i

    Set StyleNumericalMap = aDictionary
End Function


' Returns an alpha-to-numeric mapping of Palladyne's sectoral codes
Public Function SectoralNumericalMap() As Dictionary
    Dim aDictionary As Dictionary
    Dim TheSectors As Variant
    Dim i As Integer
    
    Set aDictionary = New Dictionary
    Let TheSectors = PalladyneSectors()
        
    For i = 1 To UBound(TheSectors)
        aDictionary.Add TheSectors(i), i
    Next i

    Set SectoralNumericalMap = aDictionary
End Function

' Returns a numerical-to-alpha mapping of MSCI sub-industry code-to-MSCI sub-industry name
Public Function MsciSubIndustryMap() As Dictionary
    Dim aDictionary As Dictionary
    Dim TheSubIndustries As Variant
    Dim i As Integer
    
    Set aDictionary = New Dictionary
    Let TheSubIndustries = MsciSubIndustryNames()
        
    For i = 1 To UBound(TheSubIndustries)
        aDictionary.Add i, TheSubIndustries(i)
    Next i

    Set MsciSubIndustryMap = aDictionary
End Function

