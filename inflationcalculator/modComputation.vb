Module modComputation

    Public Function loadInflationYear() As Dictionary(Of Integer, Decimal)
        Dim years As New Dictionary(Of Integer, Decimal)

        With years
            .Add(1960, 4.15)
            .Add(1961, 1.6)
            .Add(1962, 5.8)
            .Add(1963, 5.63)
            .Add(1964, 8.18)
            .Add(1965, 2.56)
            .Add(1966, 5.4)
            .Add(1967, 6.25)
            .Add(1968, 2.36)
            .Add(1969, 1.96)
            .Add(1970, 14.38)
            .Add(1971, 21.4)
            .Add(1972, 8.2)
            .Add(1973, 16.58)
            .Add(1974, 34.16)
            .Add(1975, 6.76)
            .Add(1976, 9.2)
            .Add(1977, 9.9)
            .Add(1978, 7.33)
            .Add(1979, 17.53)
            .Add(1980, 18.2)
            .Add(1981, 13.08)
            .Add(1982, 10.22)
            .Add(1983, 10.03)
            .Add(1984, 50.34)
            .Add(1985, 23.1)
            .Add(1986, 1.15)
            .Add(1987, 4.07)
            .Add(1988, 13.86)
            .Add(1989, 12.24)
            .Add(1990, 12.18)
            .Add(1991, 19.26)
            .Add(1992, 8.65)
            .Add(1993, 6.72)
            .Add(1994, 10.39)
            .Add(1995, 6.83)
            .Add(1996, 7.48)
            .Add(1997, 5.59)
            .Add(1998, 9.23)
            .Add(1999, 5.94)
            .Add(2000, 3.98)
            .Add(2001, 5.35)
            .Add(2002, 2.72)
            .Add(2003, 2.29)
            .Add(2004, 4.83)
            .Add(2005, 6.52)
            .Add(2006, 5.49)
            .Add(2007, 2.9)
            .Add(2008, 8.26)
            .Add(2009, 4.22)
            .Add(2010, 3.79)
            .Add(2011, 4.65)
            .Add(2012, 3.17)
            .Add(2013, 2.6)
            .Add(2014, 3.6)
            .Add(2015, 0.7)
            .Add(2016, 1.3)
            .Add(2017, 2.9)
            .Add(2018, 5.2)
            .Add(2019, 4.4)
        End With

        Return years

    End Function


    Public Class clsKeyValuePair
        Private _Value As Integer
        Private _Key As Decimal

        Public Sub New(Key As Integer, Value As Decimal)
            _Value = Value
            _Key = Key
        End Sub

        ReadOnly Property Value As Object
            Get
                Return _Value
            End Get
        End Property

        ReadOnly Property Tag As Object
            Get
                Return _Key
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return _Key
        End Function
    End Class


End Module
