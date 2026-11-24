Attribute VB_Name = "formTanggal"
Attribute VB_Base = "0{7277BBC4-838C-4530-9271-89B15597794E}{43BF3E1D-BD70-4CF7-976E-24C0273D7A27}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()

    Dim i As Long
    Dim dt As Date

    ComboBoxTanggal.Clear
    ComboBoxBulan.Clear
    ComboBoxTahun.Clear

    ' Bulan Indonesia
    ComboBoxBulan.AddItem "Januari"
    ComboBoxBulan.AddItem "Februari"
    ComboBoxBulan.AddItem "Maret"
    ComboBoxBulan.AddItem "April"
    ComboBoxBulan.AddItem "Mei"
    ComboBoxBulan.AddItem "Juni"
    ComboBoxBulan.AddItem "Juli"
    ComboBoxBulan.AddItem "Agustus"
    ComboBoxBulan.AddItem "September"
    ComboBoxBulan.AddItem "Oktober"
    ComboBoxBulan.AddItem "November"
    ComboBoxBulan.AddItem "Desember"

    ' Tahun
    For i = Year(Date) - 10 To Year(Date) + 10
        ComboBoxTahun.AddItem CStr(i)
    Next i

    If Not managerForm.CallerCell Is Nothing Then

        If IsDate(managerForm.CallerCell.value) Then

            dt = CDate(managerForm.CallerCell.value)

            ComboBoxBulan.ListIndex = Month(dt) - 1
            ComboBoxTahun.value = Year(dt)

            RefreshTanggal

            ComboBoxTanggal.value = Format(Day(dt), "00")

            Exit Sub

        End If

    End If

    ' Default hari ini
    ComboBoxBulan.ListIndex = Month(Date) - 1
    ComboBoxTahun.value = Year(Date)

    RefreshTanggal

    ComboBoxTanggal.value = Format(Day(Date), "00")

End Sub

Private Sub ComboBoxBulan_Change()

    RefreshTanggal

End Sub

Private Sub ComboBoxTahun_Change()

    RefreshTanggal

End Sub

Private Sub RefreshTanggal()

    Dim jumlahHari As Long
    Dim selectedDay As Long
    Dim i As Long

    If ComboBoxBulan.ListIndex = -1 Then Exit Sub
    If Trim$(ComboBoxTahun.value) = "" Then Exit Sub

    If IsNumeric(ComboBoxTanggal.value) Then
        selectedDay = CLng(ComboBoxTanggal.value)
    End If

    jumlahHari = Day(DateSerial( _
        CLng(ComboBoxTahun.value), _
        ComboBoxBulan.ListIndex + 2, _
        0))

    ComboBoxTanggal.Clear

    For i = 1 To jumlahHari
        ComboBoxTanggal.AddItem Format(i, "00")
    Next i

    If selectedDay > 0 Then

        If selectedDay > jumlahHari Then
            selectedDay = jumlahHari
        End If

        ComboBoxTanggal.value = Format(selectedDay, "00")

    End If

End Sub

Private Sub buttonInput_Click()

    Dim dt As Date
    Dim bulan As Long

    If ComboBoxTanggal.ListIndex = -1 Then
        managerForm.ShowWarning "Pilih tanggal."
        Exit Sub
    End If

    If ComboBoxBulan.ListIndex = -1 Then
        managerForm.ShowWarning "Pilih bulan."
        Exit Sub
    End If

    If ComboBoxTahun.ListIndex = -1 Then
        managerForm.ShowWarning "Pilih tahun."
        Exit Sub
    End If

    bulan = ComboBoxBulan.ListIndex + 1

    dt = DateSerial( _
        CLng(ComboBoxTahun.value), _
        bulan, _
        CLng(ComboBoxTanggal.value))

    If Not managerForm.CallerCell Is Nothing Then

        managerForm.CallerCell.value = dt

        ' Excel tetap menyimpan sebagai Date
        managerForm.CallerCell.NumberFormat = "dd-mmm-yyyy"

    End If

    Unload Me

End Sub

Private Sub buttonBatal_Click()

    Unload Me

End Sub
