Imports System.Collections.Generic

Public Class TailoringCatalogManager

    Public Class TailoringService
        Public Property CatalogID As Integer
        Public Property ServiceName As String
        Public Property Description As String
        Public Property BasePrice As Decimal
        Public Property EstimatedTime As String
    End Class

    Public Class ServiceRequirement
        Public Property CatalogID As Integer
        Public Property MaterialName As String
        Public Property DefaultQuantity As Integer
        Public Property UnitMeasure As String
        Public Property MaterialID As Integer
    End Class

    Public ReadOnly Property Catalog As List(Of TailoringService)
    Public ReadOnly Property ServiceRequirements As Dictionary(Of String, Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement))))

    Public Sub New()
        Catalog = InitializeTailoringCatalog()
        ServiceRequirements = InitializeServiceRequirements()
    End Sub

    Private Function InitializeTailoringCatalog() As List(Of TailoringService)
        Return New List(Of TailoringService) From {
            New TailoringService With {.CatalogID = 1, .ServiceName = "Hemming Pants", .Description = "Shorten pants legs", .BasePrice = 150D, .EstimatedTime = "1 day"},
            New TailoringService With {.CatalogID = 2, .ServiceName = "Zipper Replacement", .Description = "Replace broken zipper", .BasePrice = 250D, .EstimatedTime = "2 days"},
            New TailoringService With {.CatalogID = 3, .ServiceName = "Button Replacement", .Description = "Replace missing buttons", .BasePrice = 80D, .EstimatedTime = "1 hour"},
            New TailoringService With {.CatalogID = 4, .ServiceName = "Taking In Waist", .Description = "Reduce waist size", .BasePrice = 200D, .EstimatedTime = "2 days"},
            New TailoringService With {.CatalogID = 5, .ServiceName = "Tapering Legs", .Description = "Slim fit leg taper", .BasePrice = 300D, .EstimatedTime = "3 days"},
            New TailoringService With {.CatalogID = 6, .ServiceName = "Shorten Sleeves", .Description = "Shorten shirt/jacket sleeves", .BasePrice = 180D, .EstimatedTime = "1 day"},
            New TailoringService With {.CatalogID = 7, .ServiceName = "Lengthen Hem", .Description = "Add fabric to hem", .BasePrice = 220D, .EstimatedTime = "2 days"}
        }
    End Function

    Private Function InitializeServiceRequirements() As Dictionary(Of String, Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement))))
        Return New Dictionary(Of String, Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement)))) From {
            {"Pants", New Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement))) From {
                {"Cotton", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                    {1, New List(Of ServiceRequirement) From {  ' Hemming Pants
                        New ServiceRequirement With {.CatalogID = 1, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 2, .UnitMeasure = "rolls"},
                        New ServiceRequirement With {.CatalogID = 1, .MaterialID = 7, .MaterialName = "Cotton Fabric Patch", .DefaultQuantity = 1, .UnitMeasure = "pieces"}
                    }},
                    {2, New List(Of ServiceRequirement) From {  ' Zipper Replacement (Pants only)
                        New ServiceRequirement With {.CatalogID = 2, .MaterialID = 5, .MaterialName = "Zipper 7 inch Metal", .DefaultQuantity = 1, .UnitMeasure = "pieces"},
                        New ServiceRequirement With {.CatalogID = 2, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 1, .UnitMeasure = "rolls"}
                    }}
                }},
                {"Denim", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                    {1, New List(Of ServiceRequirement) From {  ' Hemming Pants
                        New ServiceRequirement With {.CatalogID = 1, .MaterialID = 3, .MaterialName = "Heavy Duty Thread", .DefaultQuantity = 3, .UnitMeasure = "rolls"},
                        New ServiceRequirement With {.CatalogID = 1, .MaterialID = 8, .MaterialName = "Denim Patch Material", .DefaultQuantity = 1, .UnitMeasure = "pieces"}
                    }},
                    {2, New List(Of ServiceRequirement) From {  ' Zipper Replacement
                        New ServiceRequirement With {.CatalogID = 2, .MaterialID = 5, .MaterialName = "Zipper 7 inch Metal", .DefaultQuantity = 1, .UnitMeasure = "pieces"},
                        New ServiceRequirement With {.CatalogID = 2, .MaterialID = 3, .MaterialName = "Heavy Duty Thread", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                    }}
                }},
                {"Gabardine", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                    {1, New List(Of ServiceRequirement) From {
                        New ServiceRequirement With {.CatalogID = 1, .MaterialID = 3, .MaterialName = "Heavy Duty Thread", .DefaultQuantity = 3, .UnitMeasure = "rolls"}
                    }},
                    {2, New List(Of ServiceRequirement) From {
                        New ServiceRequirement With {.CatalogID = 2, .MaterialID = 5, .MaterialName = "Zipper 7 inch Metal", .DefaultQuantity = 1, .UnitMeasure = "pieces"},
                        New ServiceRequirement With {.CatalogID = 2, .MaterialID = 3, .MaterialName = "Heavy Duty Thread", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                    }}
                }}
            }},
            {"Top", New Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement))) From {
                {"Cotton", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                    {3, New List(Of ServiceRequirement) From {  ' Button Replacement (Tops only)
                        New ServiceRequirement With {.CatalogID = 3, .MaterialID = 6, .MaterialName = "Buttons Plastic 15mm", .DefaultQuantity = 4, .UnitMeasure = "pieces"},
                        New ServiceRequirement With {.CatalogID = 3, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 1, .UnitMeasure = "rolls"}
                    }},
                    {6, New List(Of ServiceRequirement) From {  ' Shorten Sleeves (Tops only)
                        New ServiceRequirement With {.CatalogID = 6, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                    }}
                }},
                {"Silk", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                    {3, New List(Of ServiceRequirement) From {
                        New ServiceRequirement With {.CatalogID = 3, .MaterialID = 6, .MaterialName = "Buttons Plastic 15mm", .DefaultQuantity = 4, .UnitMeasure = "pieces"},
                        New ServiceRequirement With {.CatalogID = 3, .MaterialID = 4, .MaterialName = "Silk Thread Fine", .DefaultQuantity = 1, .UnitMeasure = "rolls"}
                    }},
                    {6, New List(Of ServiceRequirement) From {
                        New ServiceRequirement With {.CatalogID = 6, .MaterialID = 4, .MaterialName = "Silk Thread Fine", .DefaultQuantity = 2, .UnitMeasure = "rolls"},
                        New ServiceRequirement With {.CatalogID = 6, .MaterialID = 10, .MaterialName = "Silk Fabric Patch", .DefaultQuantity = 1, .UnitMeasure = "pieces"}
                    }}
                }},
                {"Polyester", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                    {3, New List(Of ServiceRequirement) From {
                        New ServiceRequirement With {.CatalogID = 3, .MaterialID = 6, .MaterialName = "Buttons Plastic 15mm", .DefaultQuantity = 4, .UnitMeasure = "pieces"},
                        New ServiceRequirement With {.CatalogID = 3, .MaterialID = 2, .MaterialName = "Polyester Thread Black", .DefaultQuantity = 1, .UnitMeasure = "rolls"}
                    }},
                    {6, New List(Of ServiceRequirement) From {
                        New ServiceRequirement With {.CatalogID = 6, .MaterialID = 2, .MaterialName = "Polyester Thread Black", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                    }}
                }}
            }},
            {"Dress", New Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement))) From {
                {"Silk", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                    {4, New List(Of ServiceRequirement) From {  ' Taking In Waist (Dress)
                        New ServiceRequirement With {.CatalogID = 4, .MaterialID = 4, .MaterialName = "Silk Thread Fine", .DefaultQuantity = 3, .UnitMeasure = "rolls"}
                    }},
                    {7, New List(Of ServiceRequirement) From {  ' Lengthen Hem (Dress)
                        New ServiceRequirement With {.CatalogID = 7, .MaterialID = 15, .MaterialName = "Silk Fabric Roll", .DefaultQuantity = 1, .UnitMeasure = "yards"},
                        New ServiceRequirement With {.CatalogID = 7, .MaterialID = 4, .MaterialName = "Silk Thread Fine", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                    }}
                }},
                {"Linen", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                    {4, New List(Of ServiceRequirement) From {
                        New ServiceRequirement With {.CatalogID = 4, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 3, .UnitMeasure = "rolls"}
                    }},
                    {7, New List(Of ServiceRequirement) From {
                        New ServiceRequirement With {.CatalogID = 7, .MaterialID = 17, .MaterialName = "Linen Fabric Roll", .DefaultQuantity = 2, .UnitMeasure = "yards"},
                        New ServiceRequirement With {.CatalogID = 7, .MaterialID = 1, .MaterialName = "Cotton Thread White", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                    }}
                }}
            }},
            {"Suit", New Dictionary(Of String, Dictionary(Of Integer, List(Of ServiceRequirement))) From {
                {"Twill", New Dictionary(Of Integer, List(Of ServiceRequirement)) From {
                    {4, New List(Of ServiceRequirement) From {
                        New ServiceRequirement With {.CatalogID = 4, .MaterialID = 3, .MaterialName = "Heavy Duty Thread", .DefaultQuantity = 3, .UnitMeasure = "rolls"},
                        New ServiceRequirement With {.CatalogID = 4, .MaterialID = 11, .MaterialName = "Interfacing Cloth", .DefaultQuantity = 1, .UnitMeasure = "meters"}
                    }},
                    {7, New List(Of ServiceRequirement) From {
                        New ServiceRequirement With {.CatalogID = 7, .MaterialID = 16, .MaterialName = "Twill Fabric Roll", .DefaultQuantity = 2, .UnitMeasure = "yards"},
                        New ServiceRequirement With {.CatalogID = 7, .MaterialID = 3, .MaterialName = "Heavy Duty Thread", .DefaultQuantity = 2, .UnitMeasure = "rolls"}
                    }}
                }}
            }}
        }
    End Function

    ' Public helper methods for Form1
    Public Function GetServiceByName(serviceName As String) As TailoringService
        Return Catalog.FirstOrDefault(Function(s) s.ServiceName = serviceName)
    End Function

    Public Function GetServiceRequirements(clothesCategory As String, fabricType As String, catalogId As Integer) As List(Of ServiceRequirement)
        If ServiceRequirements.ContainsKey(clothesCategory) AndAlso
           ServiceRequirements(clothesCategory).ContainsKey(fabricType) AndAlso
           ServiceRequirements(clothesCategory)(fabricType).ContainsKey(catalogId) Then
            Return ServiceRequirements(clothesCategory)(fabricType)(catalogId)
        End If

        ' Fallback: Category match, any fabric
        If ServiceRequirements.ContainsKey(clothesCategory) Then
            Dim categoryDict = ServiceRequirements(clothesCategory)
            For Each fabricDict In categoryDict.Values
                If fabricDict.ContainsKey(catalogId) Then
                    Return fabricDict(catalogId)
                End If
            Next
        End If

        ' Ultimate fallback: Any category + fabric
        For Each categoryDict In ServiceRequirements.Values
            If categoryDict.ContainsKey(fabricType) AndAlso categoryDict(fabricType).ContainsKey(catalogId) Then
                Return categoryDict(fabricType)(catalogId)
            End If
        Next

        Return New List(Of ServiceRequirement)()
    End Function
End Class