Attribute VB_Name = "modVBAFunctions"
Option Explicit

' modVBAFunctions.bas
' --------------------
' 4 ham VBA don gian de test IntelliSense Tooltip cua XLL: Cong, Tru, Nhan, Chia
' Mo ta ham/tham so tuong ung nam trong VBAFunctions.xml (cung thu muc voi file .xll)
'
' QUAN TRONG - vi sao cac ham duoi day khai bao Private (khong phai Public):
' Excel TU DONG lo mot Public Function trong standard module dang mo ra
' thanh UDF rieng cua no. Neu ham thuc ten la "Cong" VA la Public, no se
' vua ton tai nhu 1 UDF native cua VBA, VUA bi trung ten voi "Cong" ma XLL
' dang ky qua DangKyHamXLL - Excel co 2 dang ky doc lap cung ten "Cong",
' go =cong se thay 2 dong trong autocomplete du ten ham co doi khac di
' (da thu doi ten thanh Cong_Impl truoc do, van ra 2 dong, chi la khac ten
' - chung to nguyen nhan la o tu khoa Public, khong phai o cai ten).
' Doi Public thanh Private thi Excel KHONG con tu liet ke ham nay ra UDF
' rieng nua (Private khong duoc Excel tu dong expose), nhung xlcRun goi tu
' XLL van goi duoc binh thuong (da kiem chung thuc te) - vi vay giu dung
' ten "Cong" (khong can doi ten _Impl, khong can thuoc tinh proc rieng
' trong XML nua, dung lai TenHam = TenHamThucThi mac dinh).
'
' Cach test: import module nay vao 1 workbook (Alt+F11 > File > Import File...),
' dat VBAFunctions.xml cung thu muc voi .xll, mo lai Excel,
' go =Cong( hoac =Tru( hoac =Nhan( hoac =Chia( trong 1 cell - XLL forward
' qua xlcRun sang dung ham Private nay - tooltip va ket qua phai hien ra
' dung, va chi con DUY NHAT 1 dong "Cong" trong autocomplete.

' ---------------------------------------------------------------------
' Tinh tong cua A va B
' ---------------------------------------------------------------------
Private Function Cong(A As Double, B As Double) As Double
    Cong = A + B
End Function

' ---------------------------------------------------------------------
' Tinh hieu cua A va B
' ---------------------------------------------------------------------
Private Function Tru(A As Double, B As Double) As Double
    Tru = A - B
End Function

' ---------------------------------------------------------------------
' Tinh tich cua A va B
' ---------------------------------------------------------------------
Private Function Nhan(A As Double, B As Double) As Double
    Nhan = A * B
End Function

' ---------------------------------------------------------------------
' Tinh thuong cua A chia B - tra ve loi #DIV/0! chuan cua Excel neu B = 0
' thay vi de VBA tu nem runtime error 11 (khong bat duoc trong cell)
' ---------------------------------------------------------------------
Private Function Chia(A As Double, B As Double) As Variant
    If B = 0 Then
        Chia = CVErr(xlErrDiv0)
    Else
        Chia = A / B
    End If
End Function

' ---------------------------------------------------------------------
' Ham test tham so kieu Range THAT (khong phai gia tri tho) qua cau noi
' uXLLVBABridge.pas (type 'R') - Rng phai la 1 Range that de vong For Each
' Cell In Rng chay duoc, khong the la 1 mang gia tri thuong.
'
' Tinh tong cac gia tri so trong Rng, co the loc theo dieu kien don gian
' dang ">100", "<50", ">=10", "<=20", "=5", "<>0" (bo trong Condition =
' tinh tong tat ca o so trong Rng).
' ---------------------------------------------------------------------
Private Function TinhTongCoDieuKien(Rng As Range, Optional Condition As String = "") As Variant
    Dim Cell As Range
    Dim Tong As Double
    Dim ToanTu As String
    Dim GiaTriSoSanh As Double
    Dim CoDieuKien As Boolean

    CoDieuKien = False
    Tong = 0

    If Len(Trim(Condition)) > 0 Then
        On Error Resume Next
        Call TachDieuKien(Condition, ToanTu, GiaTriSoSanh)
        If Err.Number = 0 Then
            CoDieuKien = True
        End If
        On Error GoTo 0
    End If

    For Each Cell In Rng
        If IsNumeric(Cell.Value) Then
            If Not CoDieuKien Then
                Tong = Tong + Cell.Value
            ElseIf KiemTraDieuKien(Cell.Value, ToanTu, GiaTriSoSanh) Then
                Tong = Tong + Cell.Value
            End If
        End If
    Next Cell

    TinhTongCoDieuKien = Tong
End Function

' Tach chuoi dieu kien (vd ">=10") thanh toan tu (">=") va gia tri so (10)
Private Sub TachDieuKien(ByVal DieuKien As String, ByRef ToanTu As String, ByRef GiaTri As Double)
    Dim S As String
    S = Trim(DieuKien)

    If Left(S, 2) = ">=" Or Left(S, 2) = "<=" Or Left(S, 2) = "<>" Then
        ToanTu = Left(S, 2)
        GiaTri = CDbl(Mid(S, 3))
    ElseIf Left(S, 1) = ">" Or Left(S, 1) = "<" Or Left(S, 1) = "=" Then
        ToanTu = Left(S, 1)
        GiaTri = CDbl(Mid(S, 2))
    Else
        ToanTu = "="
        GiaTri = CDbl(S)
    End If
End Sub

' Kiem tra 1 gia tri co thoa dieu kien (toan tu + gia tri so sanh) hay khong
Private Function KiemTraDieuKien(ByVal GiaTriO As Double, ByVal ToanTu As String, ByVal GiaTriSoSanh As Double) As Boolean
    Select Case ToanTu
        Case ">"
            KiemTraDieuKien = (GiaTriO > GiaTriSoSanh)
        Case "<"
            KiemTraDieuKien = (GiaTriO < GiaTriSoSanh)
        Case ">="
            KiemTraDieuKien = (GiaTriO >= GiaTriSoSanh)
        Case "<="
            KiemTraDieuKien = (GiaTriO <= GiaTriSoSanh)
        Case "<>"
            KiemTraDieuKien = (GiaTriO <> GiaTriSoSanh)
        Case Else
            KiemTraDieuKien = (GiaTriO = GiaTriSoSanh)
    End Select
End Function
