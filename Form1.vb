Imports System.Drawing
Imports System.Globalization
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Data.OleDb
Imports System.IO
Imports System.IO.Ports
Public Class Form1
    Dim baris As Integer
    Dim iterasi As Integer
    ' Dim baris As Integer
    Dim Nam As String
    Dim tipeA As Integer = 3
    Dim l As ListViewItem
    Dim P_Stack As String
    Dim N As Integer
    Dim x As String
    Dim Tc As String
    Dim P_H2O As String
    Dim pp_H2 As String
    Dim pp_O2 As String
    ' Dim e As Exception
    Dim q As Integer
    Dim V_out As String
    Dim z As Integer
    Dim InstRead(100) As Double 'tentukan jumlah pengambilan data (tipe A)
    Dim abort As Integer
    Dim reading1 As String
    '  Dim iterasi As Integer
    ' Dim baris As Integer
    Dim channel As String
    Dim md As String
    Dim arm As String
    Dim sz As String
    Dim disp As String
    Dim clk As String
    Dim fsrc As String
    Dim src As String
    'Dim idn As String
    Dim chn As String
    Private Property SaveFileDialog1 As Object
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TabControl1.SelectedIndex = 2
        wait(1)
        Button29.PerformClick()
        wait(1)
        TabControl1.SelectedIndex = 3
        Button25.PerformClick()

    End Sub
    Public Sub ControlBmpToFile(ByVal control As Control, ByVal file As String)
        Dim bmp As New Bitmap(control.Width, control.Height)
        control.DrawToBitmap(bmp, control.DisplayRectangle)

        bmp.Save(file, System.Drawing.Imaging.ImageFormat.Jpeg)
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

   
    Public Sub wait(ByVal Dt As Double)
        Dim IDay As Double = Date.Now.DayOfYear
        Dim CDay As Double
        Dim ITime As Double = Date.Now.TimeOfDay.TotalSeconds
        Dim CTime As Double
        Dim DiffDay As Double
        Try
            Do
                Application.DoEvents()
                CDay = Date.Now.DayOfYear
                CTime = Date.Now.TimeOfDay.TotalSeconds
                DiffDay = CDay - IDay
                CTime = CTime + 86400 * DiffDay
                If CTime >= ITime + Dt Then Exit Do

            Loop
        Catch e As Exception
        End Try
    End Sub
    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        Dim tk As String
        Dim P_H2 = TextBox4.Text

        Dim i As String
        Dim P_air As String = TextBox5.Text
        Dim b As String
        Dim alpha As String = TextBox9.Text
        Dim F As String = TextBox2.Text
        Dim R As String = TextBox1.Text
        Dim V_act As String
        Dim io As String = TextBox11.Text ^ TextBox12.Text
        Dim V_ohmic As String
        Dim rin As String = TextBox8.Text
        Dim term As String
        Dim Bt As String = TextBox13.Text
        Dim V_conc As String
        Dim Alpha1 As String = TextBox10.Text
        Dim k As String = TextBox17.Text
        Dim Gf_liq As String = TextBox14.Text
        Dim E_nernst As String
        Dim a As Integer
        Chart1.Series(0).Points.Clear()
        ListView2.Items.Clear()
        Chart1.ChartAreas(0).AxisX.Maximum = Double.NaN
        Chart1.ChartAreas(0).AxisY.Maximum = Double.NaN
        Chart1.ChartAreas(0).AxisX.Minimum = Double.NaN
        Chart1.ChartAreas(0).AxisX.Minimum = Double.NaN

        tk = TextBox3.Text + 273.15
        Tc = TextBox3.Text
        ' Create loop for current 
        'loop=1;
        'i=0; 

        N = TextBox24.Text / TextBox26.Text
        z = TextBox25.Text

        For Me.baris = z + 1 To N 'Step 0.01

            l = Me.ListView2.Items.Add("")
            For j As Integer = 1 To Me.ListView2.Columns.Count
                l.SubItems.Add("")
            Next
            Try
                For Me.iterasi = 2 To tipeA
                    ListView2.Items(baris - 1).SubItems(1).Text = CDec((baris - 1) * TextBox26.Text)
                    ListView2.Items(baris - 1).SubItems(0).Text = baris

                    i = CStr((baris - 1) * TextBox26.Text) '* Math.Sqrt(2)))
                    'Calculation of Partial Pressures 
                    'Calculation of saturation pressure of water 

                    x = -2.1794 + (0.02953 * CDec(Tc)) - CDec(9.1837 * (10 ^ -5) * (Tc ^ 2)) + (1.4454 * (10 ^ -7) * (CDec(Tc) ^ 3))
                    P_H2O = (10 ^ x)
                    'Calculation of partial pressure of hydrogen 
                    pp_H2 = (0.5 * CDec((P_H2) / (Math.Exp(1.653 * i / (tk ^ 1.334))) - P_H2O))
                    'Calculation of partial pressure of oxygen 
                    pp_O2 = (P_air / Math.Exp(4.192 * i / (tk ^ 1.334))) - P_H2O
                    'Activation Losses 
                    b = (R * tk / (2 * alpha * F))
                    V_act = (-b * Math.Log(i / io))
                    'Tafel equation 

                    'Ohmic Losses
                    V_ohmic = -(i * rin)

                    'Mass Transport Losses 
                    term = (1.5 - (Bt * i))
                    If term > 0 Then
                        V_conc = Alpha1 * (i ^ k) * Math.Log(term)
                    Else
                        V_conc = 0
                    End If


                    'Calculation of Nernst voltage 
                    E_nernst = CDec(-Gf_liq / (2 * F)) + (((R * tk) / (2 * F)) * Math.Log(P_H2O / (pp_H2 * (pp_O2 ^ 0.5))))
                    'Calculation of output voltage 


                    If i = 0 Then
                        '
                        E_nernst = CDec((-Gf_liq / (2 * F)) - ((R * tk) * (1 + Math.Log(pp_H2 * (pp_O2 ^ 0.5))) / (2 * F)))
                        'i = 0.001
                        V_out = CDec(E_nernst) - i * rin + Val(V_act) + Val(V_conc)
                        ListView2.Items(baris - 1).SubItems(2).Text = V_out
                    Else
                        V_out = CDec(E_nernst) + CDec(V_ohmic) + CDec(V_act) + CDec(V_conc)
                        ListView2.Items(baris - 1).SubItems(2).Text = CDec((V_out)) '+ Math.Abs(Val(ListView2.Items(baris - 1).SubItems(2).Text) - Val(ListView2.Items(baris - 2).SubItems(2).Text)) / Math.Sqrt(2)
                    End If




                    If ListView2.Items(baris - 1).SubItems(2).Text < 0 Then
                        N = baris
                        'Exit For
                        ListView2.Items(baris - 1).Remove()
                        ' ListView2.Items.Clear()

                        ' Exit Sub
                    Else
                        If baris > 1 Then
                            ListView2.Items(baris - 1).SubItems(3).Text = CDec(ListView2.Items(baris - 2).SubItems(2).Text - ListView2.Items(baris - 1).SubItems(2).Text)
                        Else
                            ListView2.Items(baris - 1).SubItems(3).Text = CDec(0)

                        End If
                        ' If Val(ListView2.Items(baris - 1).SubItems(2).Text) < 0 Then
                        'V_out = Val(0)
                        ' ListView2.Items(baris - 1).SubItems(2).Text = V_out
                        'ListView2.Items(baris - 2).SubItems(2).Text = V_out
                        'N = baris - 1
                        ' End If



                        If ComboBox12.Text = "Real Time" Then

                            Chart1.Series(0).Points.AddXY(CDec(ListView2.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView2.Items(baris - 1).SubItems(2).Text))
                            If ComboBox17.Text = "Point" Then
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                            ElseIf ComboBox17.Text = "Bar" Then
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                            ElseIf ComboBox17.Text = "Area" Then
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                            ElseIf ComboBox17.Text = "Fast Line" Then
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
                            Else
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                            End If

                            If ComboBox13.Text = "Red" Then
                                Chart1.Series(0).Color = Color.Red
                            ElseIf ComboBox13.Text = "Green" Then
                                Chart1.Series(0).Color = Color.Green
                            ElseIf ComboBox13.Text = "Blue" Then
                                Chart1.Series(0).Color = Color.Blue
                            Else
                                Chart1.Series(0).Color = Color.Brown
                            End If
                            If ComboBox14.Text = "Dash" Then
                                With Chart1.ChartAreas(0)
                                    .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                    .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                    '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                                End With
                            Else
                                With Chart1.ChartAreas(0)
                                    .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                    .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                    '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                                End With
                            End If

                            wait(0.001)
                        End If
                        Chart1.ChartAreas("ChartArea1").AxisX.Title = "Current density (A/cm^2)"
                        Chart1.ChartAreas("ChartArea1").AxisY.Title = "Output Voltage (V)"
                        'If ListView2.Items(baris - 1).SubItems(3).Text < 0 Then
                        'End

                        'End If

                    End If


                Next
                TextBox27.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum), "0.00E0")
                TextBox28.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum), "0.00E0")
                TextBox93.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum), "0.00E0")
                TextBox94.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum), "0.00E0")
            Catch t As Exception
            End Try
        Next
        For a = 1 To ListView2.Items.Count
            If ListView2.Items(CInt(ListView2.Items.Count) - 1).SubItems(0).Text = "" Then
                ListView2.Items(CInt(ListView2.Items.Count) - 1).Remove()
            End If
            'Next
            'ListView2.Items(baris - 1).Remove()

            'Exit Sub
        Next

        ' Button14.PerformClick()


        'MsgBox(Chart1.ChartAreas(0).AxisX.ScaleView.Size)

        MsgBox("finish")


        'MsgBox(Chart1.ChartAreas(0).AxisX.ScaleView.Size)

    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ListView2.Items.Clear()
    End Sub
    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Chart1.ChartAreas(0).AxisX.ScaleView.Size = TextBox24.Text / 2
        Chart1.ChartAreas(0).AxisY.ScaleView.Size = TextBox33.Text / 2
    End Sub
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Chart1.ChartAreas(0).AxisX.ScaleView.Size = TextBox24.Text * 2
        Chart1.ChartAreas(0).AxisY.ScaleView.Size = TextBox33.Text * 2
    End Sub
    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
        End With
    End Sub
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
        End With
    End Sub
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
        End With
    End Sub
    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Chart1.Series(0).Points.Clear()
        '  Button39.PerformClick()

        Dim baris As Integer
        If TextBox32.Text = "Mode : Fuel Cell Polarization" Then

            With ListView2
                For baris = 1 To .Items.Count
                    Chart1.Series(0).Points.AddXY(CDec(.Items(baris - 1).SubItems(1).Text.ToString), CDec(.Items(baris - 1).SubItems(2).Text))
                    If ComboBox17.Text = "Point" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                    ElseIf ComboBox17.Text = "Bar" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                    ElseIf ComboBox17.Text = "Area" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                    ElseIf ComboBox17.Text = "Fast Line" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
                    Else
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                    End If

                    If ComboBox13.Text = "Red" Then
                        Chart1.Series(0).Color = Color.Red
                    ElseIf ComboBox13.Text = "Green" Then
                        Chart1.Series(0).Color = Color.Green
                    ElseIf ComboBox13.Text = "Blue" Then
                        Chart1.Series(0).Color = Color.Blue
                    Else
                        Chart1.Series(0).Color = Color.Brown
                    End If
                    If ComboBox14.Text = "Dash" Then
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        End With
                    Else
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        End With
                    End If
                Next
                Chart1.ChartAreas(0).AxisX.Title = .Columns(1).Text
                Chart1.ChartAreas(0).AxisY.Title = .Columns(2).Text
                Chart1.ChartAreas("ChartArea1").AxisX.IsLabelAutoFit = True
                Me.Chart1.ChartAreas("ChartArea1").AxisX.IsStartedFromZero = False
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Format(CDbl(TextBox27.Text + (CDec(TextBox27.Text) - CDec(TextBox94.Text)) / 5), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Format(CDbl(TextBox28.Text + (TextBox35.Text - TextBox34.Text) / 7), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Format(CDbl(TextBox94.Text), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Format(CDbl(TextBox93.Text - (TextBox35.Text - TextBox34.Text) / 7), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.00E0"
                Me.Chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "0.00E0"

            End With
        ElseIf TextBox32.Text = "Mode : Fuel Cell Power" Then
            With ListView3
                For baris = 1 To .Items.Count
                    Chart1.Series(0).Points.AddXY(CDec(.Items(baris - 1).SubItems(1).Text.ToString), CDec(.Items(baris - 1).SubItems(2).Text))
                    If ComboBox17.Text = "Point" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                    ElseIf ComboBox17.Text = "Bar" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                    ElseIf ComboBox17.Text = "Area" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                    ElseIf ComboBox17.Text = "Fast Line" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
                    Else
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                    End If

                    If ComboBox13.Text = "Red" Then
                        Chart1.Series(0).Color = Color.Red
                    ElseIf ComboBox13.Text = "Green" Then
                        Chart1.Series(0).Color = Color.Green
                    ElseIf ComboBox13.Text = "Blue" Then
                        Chart1.Series(0).Color = Color.Blue
                    Else
                        Chart1.Series(0).Color = Color.Brown
                    End If
                    If ComboBox14.Text = "Dash" Then
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        End With
                    Else
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        End With
                    End If
                Next
                Chart1.ChartAreas(0).AxisX.Title = .Columns(1).Text
                Chart1.ChartAreas(0).AxisY.Title = .Columns(2).Text
                Chart1.ChartAreas("ChartArea1").AxisX.IsLabelAutoFit = True
                Me.Chart1.ChartAreas("ChartArea1").AxisX.IsStartedFromZero = False
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Format(CDbl(TextBox27.Text + (CDec(TextBox27.Text) - CDec(TextBox94.Text)) / 5), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Format(CDbl(TextBox28.Text + (TextBox35.Text - TextBox34.Text) / 7), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Format(CDbl(TextBox94.Text - (CDec(TextBox27.Text) - CDec(TextBox94.Text)) / 8), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Format(CDbl(TextBox93.Text - (TextBox35.Text - TextBox34.Text) / 7), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.00E0"
                Me.Chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "0.00E0"
            End With

        ElseIf TextBox32.Text = "Mode : Hydrogen Usage" Then
            With ListView4
                For baris = 1 To .Items.Count
                    Chart1.Series(0).Points.AddXY(CDec(.Items(baris - 1).SubItems(1).Text.ToString), CDec(.Items(baris - 1).SubItems(2).Text))
                    If ComboBox17.Text = "Point" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                    ElseIf ComboBox17.Text = "Bar" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                    ElseIf ComboBox17.Text = "Area" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                    ElseIf ComboBox17.Text = "Fast Line" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
                    Else
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                    End If

                    If ComboBox13.Text = "Red" Then
                        Chart1.Series(0).Color = Color.Red
                    ElseIf ComboBox13.Text = "Green" Then
                        Chart1.Series(0).Color = Color.Green
                    ElseIf ComboBox13.Text = "Blue" Then
                        Chart1.Series(0).Color = Color.Blue
                    Else
                        Chart1.Series(0).Color = Color.Brown
                    End If
                    If ComboBox14.Text = "Dash" Then
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        End With
                    Else
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        End With
                    End If
                Next
                Chart1.ChartAreas(0).AxisX.Title = .Columns(1).Text
                Chart1.ChartAreas(0).AxisY.Title = .Columns(2).Text
                Chart1.ChartAreas("ChartArea1").AxisX.IsLabelAutoFit = True
                Me.Chart1.ChartAreas("ChartArea1").AxisX.IsStartedFromZero = False
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Format(CDbl(TextBox27.Text + (CDec(TextBox27.Text) - CDec(TextBox94.Text)) / 8), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Format(CDbl(TextBox28.Text + (TextBox35.Text - TextBox34.Text) / 7), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Format(CDbl(TextBox94.Text - (CDec(TextBox27.Text) - CDec(TextBox94.Text)) / 8), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Format(CDbl(TextBox93.Text - (TextBox35.Text - TextBox34.Text) / 7), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.00E0"
                Me.Chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "0.00E0"
            End With
        ElseIf TextBox32.Text = "Mode : Power Vs Hydrogen" Then
            With ListView8
                For baris = 1 To .Items.Count
                    Chart1.Series(0).Points.AddXY(CDec(.Items(baris - 1).SubItems(1).Text.ToString), CDec(.Items(baris - 1).SubItems(2).Text))
                    If ComboBox17.Text = "Point" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                    ElseIf ComboBox17.Text = "Bar" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                    ElseIf ComboBox17.Text = "Area" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                    ElseIf ComboBox17.Text = "Fast Line" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
                    Else
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                    End If

                    If ComboBox13.Text = "Red" Then
                        Chart1.Series(0).Color = Color.Red
                    ElseIf ComboBox13.Text = "Green" Then
                        Chart1.Series(0).Color = Color.Green
                    ElseIf ComboBox13.Text = "Blue" Then
                        Chart1.Series(0).Color = Color.Blue
                    Else
                        Chart1.Series(0).Color = Color.Brown
                    End If
                    If ComboBox14.Text = "Dash" Then
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        End With
                    Else
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        End With
                    End If
                Next
                Chart1.ChartAreas(0).AxisX.Title = .Columns(1).Text
                Chart1.ChartAreas(0).AxisY.Title = .Columns(2).Text
                Chart1.ChartAreas("ChartArea1").AxisX.IsLabelAutoFit = True
                Me.Chart1.ChartAreas("ChartArea1").AxisX.IsStartedFromZero = False
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Format(CDbl(TextBox27.Text + (CDec(TextBox27.Text) - CDec(TextBox94.Text)) / 5), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Format(CDbl(TextBox28.Text + (TextBox35.Text - TextBox34.Text) / 7), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Format(CDbl(TextBox94.Text - (CDec(TextBox27.Text) - CDec(TextBox94.Text)) / 8), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Format(CDbl(TextBox93.Text - (TextBox35.Text - TextBox34.Text) / 7), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.00E0"
                Me.Chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "0.00E0"
            End With
        ElseIf TextBox32.Text = "Mode : Power VS Partial Pressure" Then
            With ListView9
                For baris = 1 To .Items.Count
                    Chart1.Series(0).Points.AddXY(CDec(.Items(baris - 1).SubItems(1).Text.ToString), CDec(.Items(baris - 1).SubItems(2).Text))
                    If ComboBox17.Text = "Point" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                    ElseIf ComboBox17.Text = "Bar" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                    ElseIf ComboBox17.Text = "Area" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                    ElseIf ComboBox17.Text = "Fast Line" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
                    Else
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                    End If

                    If ComboBox13.Text = "Red" Then
                        Chart1.Series(0).Color = Color.Red
                    ElseIf ComboBox13.Text = "Green" Then
                        Chart1.Series(0).Color = Color.Green
                    ElseIf ComboBox13.Text = "Blue" Then
                        Chart1.Series(0).Color = Color.Blue
                    Else
                        Chart1.Series(0).Color = Color.Brown
                    End If
                    If ComboBox14.Text = "Dash" Then
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        End With
                    Else
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        End With
                    End If
                Next
                Chart1.ChartAreas(0).AxisX.Title = .Columns(1).Text
                Chart1.ChartAreas(0).AxisY.Title = .Columns(2).Text
                Chart1.ChartAreas("ChartArea1").AxisX.IsLabelAutoFit = True
                Me.Chart1.ChartAreas("ChartArea1").AxisX.IsStartedFromZero = False
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum = (CDec(.Items(.Items.Count - 1).SubItems(1).Text) - Math.Abs(CDec(.Items(.Items.Count - 1).SubItems(1).Text) - CDec(.Items(0).SubItems(1).Text)) / 8)
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = (CDec(TextBox28.Text) + CDec(TextBox33.Text) / 5)
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = (CDec(.Items(0).SubItems(1).Text) + Math.Abs(CDec(.Items(.Items.Count - 1).SubItems(1).Text) - CDec(.Items(0).SubItems(1).Text)) / 5)
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = (CDbl(TextBox93.Text) - CDec(TextBox33.Text) / 5)
                Me.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.00E0"
                Me.Chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "0.00E0"
            End With
        ElseIf TextBox32.Text = "Mode : Water Produced" Then
            With ListView5
                For baris = 1 To .Items.Count
                    Chart1.Series(0).Points.AddXY(CDec(.Items(baris - 1).SubItems(1).Text.ToString), CDec(.Items(baris - 1).SubItems(2).Text))
                    If ComboBox17.Text = "Point" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                    ElseIf ComboBox17.Text = "Bar" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                    ElseIf ComboBox17.Text = "Area" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                    ElseIf ComboBox17.Text = "Fast Line" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
                    Else
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                    End If

                    If ComboBox13.Text = "Red" Then
                        Chart1.Series(0).Color = Color.Red
                    ElseIf ComboBox13.Text = "Green" Then
                        Chart1.Series(0).Color = Color.Green
                    ElseIf ComboBox13.Text = "Blue" Then
                        Chart1.Series(0).Color = Color.Blue
                    Else
                        Chart1.Series(0).Color = Color.Brown
                    End If
                    If ComboBox14.Text = "Dash" Then
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        End With
                    Else
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        End With
                    End If
                Next
                Chart1.ChartAreas(0).AxisX.Title = .Columns(1).Text
                Chart1.ChartAreas(0).AxisY.Title = .Columns(2).Text
                Chart1.ChartAreas("ChartArea1").AxisX.IsLabelAutoFit = True
                Me.Chart1.ChartAreas("ChartArea1").AxisX.IsStartedFromZero = False

                Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum = (CDec(.Items(.Items.Count - 1).SubItems(1).Text) - Math.Abs(CDec(.Items(.Items.Count - 1).SubItems(1).Text) - CDec(.Items(0).SubItems(1).Text)) / 5)
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = (CDec(TextBox28.Text) + CDec(TextBox33.Text) / 5)
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = (CDec(.Items(0).SubItems(1).Text) + Math.Abs(CDec(.Items(.Items.Count - 1).SubItems(1).Text) - CDec(.Items(0).SubItems(1).Text)) / 5)
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = (CDbl(TextBox93.Text) - CDec(TextBox33.Text) / 5)
                Me.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.00E0"
                Me.Chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "0.00E0"
            End With

        ElseIf TextBox32.Text = "Mode : Heat Generated" Then
            With ListView6
                For baris = 1 To .Items.Count
                    Chart1.Series(0).Points.AddXY(CDec(.Items(baris - 1).SubItems(1).Text.ToString), CDec(.Items(baris - 1).SubItems(2).Text))
                    If ComboBox17.Text = "Point" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                    ElseIf ComboBox17.Text = "Bar" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                    ElseIf ComboBox17.Text = "Area" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                    ElseIf ComboBox17.Text = "Fast Line" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
                    Else
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                    End If

                    If ComboBox13.Text = "Red" Then
                        Chart1.Series(0).Color = Color.Red
                    ElseIf ComboBox13.Text = "Green" Then
                        Chart1.Series(0).Color = Color.Green
                    ElseIf ComboBox13.Text = "Blue" Then
                        Chart1.Series(0).Color = Color.Blue
                    Else
                        Chart1.Series(0).Color = Color.Brown
                    End If
                    If ComboBox14.Text = "Dash" Then
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        End With
                    Else
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        End With
                    End If
                Next
                Chart1.ChartAreas(0).AxisX.Title = .Columns(1).Text
                Chart1.ChartAreas(0).AxisY.Title = .Columns(2).Text
                Chart1.ChartAreas("ChartArea1").AxisX.IsLabelAutoFit = True
                Me.Chart1.ChartAreas("ChartArea1").AxisX.IsStartedFromZero = False
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Format(CDbl(TextBox27.Text + (CDec(TextBox27.Text) - CDec(TextBox94.Text)) / 8), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Format(CDbl(TextBox28.Text + (TextBox35.Text - TextBox34.Text) / 5), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Format(CDbl(TextBox94.Text - (CDec(TextBox27.Text) - CDec(TextBox94.Text)) / 8), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Format(CDbl(TextBox93.Text - (TextBox35.Text - TextBox34.Text) / 5), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.00E0"
                Me.Chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "0.00E0"
            End With

        ElseIf TextBox32.Text = "Mode : Efficiency" Then
            With ListView7
                For baris = 1 To .Items.Count
                    Chart1.Series(0).Points.AddXY(CDec(.Items(baris - 1).SubItems(1).Text.ToString), CDec(.Items(baris - 1).SubItems(2).Text))
                    If ComboBox17.Text = "Point" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                    ElseIf ComboBox17.Text = "Bar" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                    ElseIf ComboBox17.Text = "Area" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                    ElseIf ComboBox17.Text = "Fast Line" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
                    Else
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                    End If

                    If ComboBox13.Text = "Red" Then
                        Chart1.Series(0).Color = Color.Red
                    ElseIf ComboBox13.Text = "Green" Then
                        Chart1.Series(0).Color = Color.Green
                    ElseIf ComboBox13.Text = "Blue" Then
                        Chart1.Series(0).Color = Color.Blue
                    Else
                        Chart1.Series(0).Color = Color.Brown
                    End If
                    If ComboBox14.Text = "Dash" Then
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        End With
                    Else
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        End With
                    End If
                Next
                Chart1.ChartAreas(0).AxisX.Title = .Columns(1).Text
                Chart1.ChartAreas(0).AxisY.Title = .Columns(2).Text
                Chart1.ChartAreas("ChartArea1").AxisX.IsLabelAutoFit = True
                Me.Chart1.ChartAreas("ChartArea1").AxisX.IsStartedFromZero = False
                'Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Format(CDbl(TextBox27.Text + (CDec(TextBox27.Text) - CDec(TextBox94.Text)) / 5), "0.00E0")
                'Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Format(CDbl(TextBox28.Text + (TextBox35.Text - TextBox34.Text) / 7), "0.00E0")
                ' Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Format(CDbl(TextBox94.Text - (CDec(TextBox27.Text) - CDec(TextBox94.Text)) / 5), "0.00E0")
                ' Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Format(CDbl(TextBox93.Text - (TextBox35.Text - TextBox34.Text) / 7), "0.00E0")
                Me.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.00"
                Me.Chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Format = "0.00"
            End With
        Else

            Timer1.Enabled = True
        End If


    End Sub


    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim saveFileDialog1 As New SaveFileDialog()

        ' Sets the current file name filter string, which determines 
        ' the choices that appear in the "Save as file type" or 
        ' "Files of type" box in the dialog box.
        saveFileDialog1.Filter = "Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg|EMF (*.emf)|*.emf|PNG (*.png)|*.png|SVG (*.svg)|*.svg|GIF (*.gif)|*.gif|TIFF (*.tif)|*.tif"
        saveFileDialog1.FilterIndex = 2
        saveFileDialog1.RestoreDirectory = True

        ' Set image file format
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim format As ChartImageFormat = ChartImageFormat.Bmp

            If saveFileDialog1.FileName.EndsWith("bmp") Then
                format = ChartImageFormat.Bmp
            Else
                If saveFileDialog1.FileName.EndsWith("jpg") Then
                    format = ChartImageFormat.Jpeg
                Else
                    If saveFileDialog1.FileName.EndsWith("emf") Then
                        format = ChartImageFormat.Emf
                    Else
                        If saveFileDialog1.FileName.EndsWith("gif") Then
                            format = ChartImageFormat.Gif
                        Else
                            If saveFileDialog1.FileName.EndsWith("png") Then
                                format = ChartImageFormat.Png
                            Else
                                If saveFileDialog1.FileName.EndsWith("tif") Then
                                    format = ChartImageFormat.Tiff

                                End If
                            End If ' Save image
                        End If
                    End If
                End If
            End If
            Chart1.SaveImage(saveFileDialog1.FileName, format)
        End If
    End Sub

    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        For g = 1 To ListView2.Items.Count
            If ListView2.Items(g - 1).BackColor = Color.LightGreen Then
                'ListView2.Items(g - 1).BackColor = Color.LightGreen
                'TextBox35.Text = Format(Val(TextBox35.Text), "0.00E0")
                Chart1.Series(0).Points(g - 1).Label = "max = " + Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam

                'ChartArea1.AxisX.islabelautofit = True
                Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart1.Series(0).Points(g - 1).MarkerSize = 10
                Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                '  
            ElseIf ListView2.Items(g - 1).BackColor = Color.Yellow Then
                ' ListView2.Items(g - 1).BackColor = Color.Yellow


                Chart1.Series(0).Points(g - 1).Label = "Stable = " + Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart1.Series(0).Points(g - 1).MarkerSize = 10
                Chart1.Series(0).Points(g - 1).MarkerColor = Color.Yellow


            ElseIf Val(ListView2.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then
                ' ListView2.Items(g - 1).BackColor = Color.LightBlue
                ' TextBox31.BackColor = Color.LightBlue

                Chart1.Series(0).Points(g - 1).Label = "Peaky = " & Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart1.Series(0).Points(g - 1).MarkerSize = 10
                Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightBlue
                ' Chart1.Series(0).YValueMembers = "Quan"
                'Chart1.Series(0).Points(g).LabelForeColor = Color.LightBlue
            ElseIf ListView2.Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                'MsgBox(g)
                'MsgBox(ListView2.Items(g - 1).SubItems(2).Text)

                'ListView2.Items(g - 1).BackColor = Color.Orange
                ' TextBox34.BackColor = Color.Orange
                Chart1.Series(0).Points(g - 1).Label = "min = " + Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                ' MsgBox(Chart1.Series(0).Points(g - 1).Label)
                Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart1.Series(0).Points(g - 1).MarkerSize = 10
                Chart1.Series(0).Points(g - 1).MarkerColor = Color.Orange
                'Chart1.Series(0).Points(g).LabelForeColor = Color.Orange
            ElseIf ListView2.Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                ' ListView2.Items(g - 1).BackColor = Color.LightCyan
                ' TextBox30.BackColor = Color.LightCyan
                Chart1.Series(0).Points(g - 1).Label = "mean = " + ListView2.Items(g - 1).SubItems(2).Text & Nam
                Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                Chart1.Series(0).Points(g - 1).MarkerSize = 10
                Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightCyan
                'Chart1.Series(0).Points(g).LabelForeColor = Color.LightCyan

            End If

        Next g

        Chart1.Series(0).SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No
        Chart1.Series(0).SmartLabelStyle.IsMarkerOverlappingAllowed = True
        Chart1.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.Right
        'Chart1.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopRight
        Button12.PerformClick()
    End Sub

    Private Sub Chart1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To ListView2.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = ListView2.Columns(j).Text.ToString
            col = col + 1
        Next


        For i = 0 To ListView2.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = CDec(ListView2.Items.Item(i).Text.ToString)
            xlWorkSheet.Cells(i + 2, 2) = Format(CDec(ListView2.Items.Item(i).SubItems(1).Text), "0.0000")
            xlWorkSheet.Cells(i + 2, 3) = CDec(ListView2.Items.Item(i).SubItems(2).Text)
            xlWorkSheet.Cells(i + 2, 4) = CDec(ListView2.Items.Item(i).SubItems(3).Text)

        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = " "
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()


    End Sub
    ' End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Chart1.Series(0).Points.Clear()
    End Sub

    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.DashDot
            '.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With

    End Sub

    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        With Chart1.ChartAreas(0)
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.DashDot
            '.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        For g = 1 To ListView2.Items.Count
            If CDec(ListView2.Items(g - 1).SubItems(2).Text) = CDec(TextBox35.Text) Then
                'ListView2.Items(g - 1).BackColor = Color.LightGreen
                TextBox35.Text = Format(Val(TextBox35.Text), "0.00E0")
                Chart1.Series(0).Points(g - 1).Label = ""
                'Chart1.Series(0).Points(g - 1).LabelFormat.ToLower()

                Chart1.Series(0).Points(g - 1).MarkerSize = 0

            ElseIf ListView2.Items(g - 1).BackColor = Color.Yellow Then

                TextBox43.BackColor = Color.Yellow
                TextBox43.Text = Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.00E0")
                Chart1.Series(0).Points(g - 1).Label = ""
                Chart1.Series(0).Points(g - 1).MarkerSize = 0



            ElseIf Val(ListView2.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then
                ' ListView2.Items(g - 1).BackColor = Color.LightBlue
                ' TextBox31.BackColor = Color.LightBlue
                TextBox44.BackColor = Color.LightBlue
                TextBox44.Text = Format(Val(ListView2.Items(g - 1).SubItems(2).Text), "0.00E0")
                Chart1.Series(0).Points(g - 1).Label = ""

                Chart1.Series(0).Points(g - 1).MarkerSize = 0

            ElseIf ListView2.Items(g - 1).SubItems(2).Text = TextBox34.Text Then

                Chart1.Series(0).Points(g - 1).Label = ""

                Chart1.Series(0).Points(g - 1).MarkerSize = 0

            ElseIf ListView2.Items(g - 1).SubItems(2).Text = TextBox30.Text Then

                Chart1.Series(0).Points(g - 1).Label = ""

                Chart1.Series(0).Points(g - 1).MarkerSize = 0
            End If

        Next g
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim openFileDialog1 As New OpenFileDialog()
        'set the root to the z drive
        openFileDialog1.InitialDirectory = "D:\"
        'make sure the root goes back to where the user started
        ''   openFileDialog1.RestoreDirectory = True
        'show the dialog
        ''  openFileDialog1.ShowDialog()
        ''
        '' If (openFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK) Then
        ''Path = openFileDialog1.FileName
        ''  End If

        ' Call ShowDialog.
        Dim result As DialogResult = openFileDialog1.ShowDialog()

        ' Test result.
        Dim path As String = openFileDialog1.FileName
        If result = Windows.Forms.DialogResult.OK Then
            Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Extended Properties=Excel 12.0;"
            ' Get the file name.

            '  Try
            Dim table As DataTable = New DataTable()
            Dim excelName As String = "Sheet1"
            Dim strConnection As String = String.Format(connStr)
            Dim conn As OleDbConnection = New OleDbConnection(strConnection)
            conn.Open()
            Dim oada As OleDbDataAdapter = New OleDbDataAdapter("select * from [" & excelName & "$]", strConnection)
            table.TableName = "TableInfo"
            oada.Fill(table)
            conn.Close()
            ListView2.Items.Clear()

            For i As Integer = 0 To table.Rows.Count - 1
                Dim drow As DataRow = table.Rows(i)

                If drow.RowState <> DataRowState.Deleted Then
                    Dim lvi As ListViewItem = New ListViewItem(drow("No").ToString())
                    lvi.SubItems.Add(drow("Curr density (A/cm^2)").ToString())
                    lvi.SubItems.Add(drow("Output voltage (Volts)").ToString())
                    lvi.SubItems.Add(drow("Sequence Discreapancy").ToString())
                    ListView2.Items.Add(lvi)
                End If
            Next
            ' Read in text.
            '' Dim text As String = File.ReadAllText(path)

            ' For debugging.
            ''  Me.Text = text.Length.ToString

            '  Catch ex As Exception

            ' Report an error.
            ''   Me.Text = "Error"

            ' End Try
        End If


        ' For Me.baris = 1 To Me.ListView2.Items.Count
        ' l = Me.ListView2.Items.Add("")
        ' For j As Integer = 1 To Me.ListView2.Columns.Count
        ' l.SubItems.Add("")
        'Next
        'ListView2.Items(baris - 1).SubItems(0).Text = baris
        ' Next


    End Sub

    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            '.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        With Chart1.ChartAreas(0)
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
            '.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TextBox29.Text = ""
        TextBox29.BackColor = Color.White
        TextBox30.Text = ""
        TextBox30.BackColor = Color.White
        TextBox31.Text = ""
        TextBox31.BackColor = Color.White
        TextBox33.Text = ""
        TextBox33.BackColor = Color.White
        TextBox34.Text = ""
        TextBox34.BackColor = Color.White
        TextBox35.Text = ""
        TextBox35.BackColor = Color.White
        TextBox44.BackColor = Color.White
        TextBox44.Text = ""
        TextBox43.Text = ""
        TextBox43.BackColor = Color.White

    End Sub

    Private Sub Button40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Button39.PerformClick()
        Button6.PerformClick()
        Button34.PerformClick()
        MsgBox("All data have cleared")
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        Try
            Chart1.Series(0).Points.Clear()
            Nam = " Watts"
            ListView3.Items.Clear()
            Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
            Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
            Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
            Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
            Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
            Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
            Chart1.Series(0).Name = "P Out"
            If Me.ListView2.Items(0).SubItems(0).Text = "" Then
                MsgBox("Run FuelCell Polarization First")
            Else
                q = 0
                For i = 1 To ListView2.Items.Count
                    If Not ListView2.Items(i - 1).SubItems(0).Text = "" Then
                        q = q + 1
                    Else
                        q = q
                    End If
                Next
                wait(0.001)
                'N = ListView2.Items.Count
                z = TextBox25.Text
                For Me.baris = z + 1 To q 'Step 0.01
                    l = Me.ListView3.Items.Add("")
                    For j As Integer = 1 To Me.ListView3.Columns.Count
                        l.SubItems.Add("")
                    Next
                    For Me.iterasi = 2 To tipeA

                        P_Stack = CDec(ListView2.Items(baris - 1).SubItems(2).Text) * CDec(ListView2.Items(baris - 1).SubItems(1).Text) * Val(TextBox6.Text) * Val(TextBox7.Text)

                        ListView3.Items(baris - 1).SubItems(1).Text = CDec((baris - 1) / 100)
                        ListView3.Items(baris - 1).SubItems(0).Text = baris


                        ListView3.Items(baris - 1).SubItems(2).Text = CDec(P_Stack)
                        If baris > 1 Then
                            ListView3.Items(baris - 1).SubItems(3).Text = ListView3.Items(baris - 2).SubItems(2).Text - ListView3.Items(baris - 1).SubItems(2).Text
                        Else
                            ListView3.Items(baris - 1).SubItems(3).Text = 0

                        End If

                        If ComboBox12.Text = "Real Time" Then

                            Chart1.Series(0).Points.AddXY(CDec(ListView3.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView3.Items(baris - 1).SubItems(2).Text))
                            If ComboBox17.Text = "Point" Then
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                            ElseIf ComboBox17.Text = "Bar" Then
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                            ElseIf ComboBox17.Text = "Area" Then
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                            ElseIf ComboBox17.Text = "Fast Line" Then
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                            Else
                                Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                            End If

                            If ComboBox13.Text = "Red" Then
                                Chart1.Series(0).Color = Color.Red
                            ElseIf ComboBox13.Text = "Green" Then
                                Chart1.Series(0).Color = Color.Green
                            ElseIf ComboBox13.Text = "Blue" Then
                                Chart1.Series(0).Color = Color.Blue
                            Else
                                Chart1.Series(0).Color = Color.Brown
                            End If
                            If ComboBox14.Text = "Dash" Then
                                With Chart1.ChartAreas(0)
                                    .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                    .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                    '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                                End With
                            Else
                                With Chart1.ChartAreas(0)
                                    .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                    .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                    '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                                End With
                            End If

                            wait(0.001)
                        End If

                        Chart1.ChartAreas("ChartArea1").AxisX.Title = ListView3.Columns(1).Text
                        Chart1.ChartAreas("ChartArea1").AxisY.Title = ListView3.Columns(2).Text

                    Next
                    TextBox27.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum), "0.00E0")
                    TextBox28.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum), "0.00E0")
                    TextBox93.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum), "0.00E0")
                    TextBox94.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum), "0.00E0")

                Next

                Button14.PerformClick()
                Button19.PerformClick()
                Button32.PerformClick()
                Button12.PerformClick()
                TextBox34.Text = Format(CDbl(TextBox34.Text), "0.00E0")
                TextBox35.Text = Format(CDbl(TextBox35.Text), "0.00E0")
                TextBox31.Text = Format(CDbl(TextBox31.Text), "0.00E0")
                TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
                TextBox45.Text = Format(CDbl(TextBox45.Text), "0.00E0")
                TextBox44.Text = Format(CDbl(TextBox44.Text), "0.00E0")
                TextBox43.Text = Format(CDbl(TextBox43.Text), "0.00E0")
                TextBox29.Text = Format(CDbl(TextBox29.Text), "0.00E0")
                TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
                TextBox33.Text = Format(CDbl(TextBox33.Text), "0.00E0")
                'msgbox("finish")

            End If
        Catch t As Exception
        End Try
    End Sub


    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim openFileDialog1 As New OpenFileDialog()
        'set the root to the z drive
        openFileDialog1.InitialDirectory = "D:\"
        'make sure the root goes back to where the user started
        ''   openFileDialog1.RestoreDirectory = True
        'show the dialog
        ''  openFileDialog1.ShowDialog()
        ''
        '' If (openFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK) Then
        ''Path = openFileDialog1.FileName
        ''  End If

        ' Call ShowDialog.
        Dim result As DialogResult = openFileDialog1.ShowDialog()

        ' Test result.
        Dim path As String = openFileDialog1.FileName
        If result = Windows.Forms.DialogResult.OK Then
            Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Extended Properties=Excel 12.0;"
            ' Get the file name.

            '  Try
            Dim table As DataTable = New DataTable()
            Dim excelName As String = "Sheet1"
            Dim strConnection As String = String.Format(connStr)
            Dim conn As OleDbConnection = New OleDbConnection(strConnection)
            conn.Open()
            Dim oada As OleDbDataAdapter = New OleDbDataAdapter("select * from [" & excelName & "$]", strConnection)
            table.TableName = "TableInfo"
            oada.Fill(table)
            conn.Close()
            ListView3.Items.Clear()

            For i As Integer = 0 To table.Rows.Count - 1
                Dim drow As DataRow = table.Rows(i)

                If drow.RowState <> DataRowState.Deleted Then
                    Dim lvi As ListViewItem = New ListViewItem(drow("No").ToString())
                    lvi.SubItems.Add(drow("Curr density (A/cm^2)").ToString())
                    lvi.SubItems.Add(drow("Power (Watts)").ToString())
                    lvi.SubItems.Add(drow("Sequence Discreapancy").ToString())
                    ListView3.Items.Add(lvi)
                End If
            Next
            ' Read in text.
            '' Dim text As String = File.ReadAllText(path)

            ' For debugging.
            ''  Me.Text = text.Length.ToString

            '  Catch ex As Exception

            ' Report an error.
            ''   Me.Text = "Error"

            ' End Try
        End If


        ' For Me.baris = 1 To Me.listview3.Items.Count
        ' l = Me.listview3.Items.Add("")
        ' For j As Integer = 1 To Me.listview3.Columns.Count
        ' l.SubItems.Add("")
        'Next
        'listview3.Items(baris - 1).SubItems(0).Text = baris
        ' Next


    End Sub
    Private Sub Button48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To ListView3.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = ListView3.Columns(j).Text.ToString
            col = col + 1
        Next


        For i = 0 To ListView3.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = CDec(ListView3.Items.Item(i).Text.ToString)
            xlWorkSheet.Cells(i + 2, 2) = Format(CDec(ListView3.Items.Item(i).SubItems(1).Text), "0.0000")
            xlWorkSheet.Cells(i + 2, 3) = CDec(ListView3.Items.Item(i).SubItems(2).Text)
            xlWorkSheet.Cells(i + 2, 4) = CDec(ListView3.Items.Item(i).SubItems(3).Text)

        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = " "
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()


    End Sub

    Private Sub Button46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ListView3.Items.Clear()
    End Sub
    Private Sub Button57_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button57.Click
        Dim N_cells As String = TextBox7.Text
        Dim A_cell As String = TextBox6.Text
        Dim Hyd_usage As String
        Dim F As String = TextBox2.Text
        'N = ListView2.Items.Count
        Chart1.Series(0).Points.Clear()
        Nam = " Watts"
        ListView4.Items.Clear()
        Chart1.Series(0).Name = "H2 Usg"
        Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN

        Try
            z = TextBox25.Text

            For Me.baris = z + 1 To ListView3.Items.Count 'Step 0.01
                l = Me.ListView4.Items.Add("")
                For j As Integer = 1 To Me.ListView3.Columns.Count
                    l.SubItems.Add("")
                Next

                P_Stack = N_cells * (CDec(ListView2.Items(baris - 1).SubItems(2).Text) * (CDec(ListView3.Items(baris - 1).SubItems(1).Text) * A_cell))
                If CDec(ListView2.Items(baris - 1).SubItems(2).Text) < Val(0) Then
                    Hyd_usage = 0
                Else
                    Hyd_usage = (P_Stack) / (2 * CDec(ListView2.Items(baris - 1).SubItems(2).Text) * F)
                End If

                ListView4.Items(baris - 1).SubItems(0).Text = baris
                ListView4.Items(baris - 1).SubItems(1).Text = CDec(P_Stack)
                ListView4.Items(baris - 1).SubItems(2).Text = CDec(Hyd_usage)
                If Not baris < 2 Then
                    ListView4.Items(baris - 1).SubItems(3).Text = CDec(ListView4.Items(baris - 1).SubItems(2).Text) - CDec(ListView4.Items(baris - 2).SubItems(2).Text)
                Else
                    ListView4.Items(baris - 1).SubItems(3).Text = 0
                End If
                If ComboBox12.Text = "Real Time" Then

                    Chart1.Series(0).Points.AddXY(CDec(ListView4.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView4.Items(baris - 1).SubItems(2).Text))
                    If ComboBox17.Text = "Point" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                    ElseIf ComboBox17.Text = "Bar" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                    ElseIf ComboBox17.Text = "Area" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                    ElseIf ComboBox17.Text = "Fast Line" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                    Else
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                    End If

                    If ComboBox13.Text = "Red" Then
                        Chart1.Series(0).Color = Color.Red
                    ElseIf ComboBox13.Text = "Green" Then
                        Chart1.Series(0).Color = Color.Green
                    ElseIf ComboBox13.Text = "Blue" Then
                        Chart1.Series(0).Color = Color.Blue
                    Else
                        Chart1.Series(0).Color = Color.Brown
                    End If
                    If ComboBox14.Text = "Dash" Then
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                        End With
                    Else
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                        End With
                    End If

                    wait(0.001)
                End If

                Chart1.ChartAreas("ChartArea1").AxisY.Title = ListView4.Columns(2).Text
                Chart1.ChartAreas("ChartArea1").AxisX.Title = ListView4.Columns(1).Text
            Next
            TextBox27.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum), "0.00E0")
            TextBox28.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum), "0.00E0")
            TextBox93.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum), "0.00E0")
            TextBox94.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum), "0.00E0")




            Button14.PerformClick()
            Button19.PerformClick()
            Button32.PerformClick()
            Button12.PerformClick()
            TextBox34.Text = Format(CDbl(TextBox34.Text), "0.00E0")
            TextBox35.Text = Format(CDbl(TextBox35.Text), "0.00E0")
            TextBox31.Text = Format(CDbl(TextBox31.Text), "0.00E0")
            TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
            TextBox45.Text = Format(CDbl(TextBox45.Text), "0.00E0")
            TextBox44.Text = Format(CDbl(TextBox44.Text), "0.00E0")
            TextBox43.Text = Format(CDbl(TextBox43.Text), "0.00E0")
            TextBox29.Text = Format(CDbl(TextBox29.Text), "0.00E0")
            TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
            TextBox33.Text = Format(CDbl(TextBox33.Text), "0.00E0")

        Catch t As Exception
        End Try

        'msgbox("finish")
    End Sub

    Private Sub Button78_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button78.Click
        'P_Stack = N_cells * (V_out * (i * A_cell))
        Dim Water_prod As String
        Dim mm_H2o As String = TextBox19.Text
        Dim F As String = TextBox2.Text
        Chart1.Series(0).Points.Clear()
        Nam = " litter/hour"
        ListView5.Items.Clear()
        Chart1.Series(0).Name = "H2O Out"
        Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN

        Try
            For Me.baris = 1 To ListView3.Items.Count 'Step 0.01
                l = Me.ListView5.Items.Add("")
                For j As Integer = 1 To Me.ListView5.Columns.Count
                    l.SubItems.Add("")
                Next
                If ListView2.Items(baris - 1).SubItems(2).Text < 0 Then
                    Water_prod = 0
                Else
                    Water_prod = (mm_H2o * ListView3.Items(baris - 1).SubItems(2).Text) / (3600 * ListView2.Items(baris - 1).SubItems(2).Text * F)
                End If

                ListView5.Items(baris - 1).SubItems(0).Text = baris
                ListView5.Items(baris - 1).SubItems(1).Text = CDec(ListView2.Items(baris - 1).SubItems(2).Text)
                ListView5.Items(baris - 1).SubItems(2).Text = CDec(Water_prod)
                If Not baris < 2 Then
                    ListView5.Items(baris - 1).SubItems(3).Text = Val(ListView5.Items(baris - 1).SubItems(2).Text) - Val(ListView5.Items(baris - 2).SubItems(2).Text)
                Else
                    ListView5.Items(baris - 1).SubItems(3).Text = 0
                End If
                If ComboBox12.Text = "Real Time" Then

                    Chart1.Series(0).Points.AddXY(CDec(ListView5.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView5.Items(baris - 1).SubItems(2).Text))
                    If ComboBox17.Text = "Point" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                    ElseIf ComboBox17.Text = "Bar" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                    ElseIf ComboBox17.Text = "Area" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                    ElseIf ComboBox17.Text = "Fast Line" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                    Else
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                    End If

                    If ComboBox13.Text = "Red" Then
                        Chart1.Series(0).Color = Color.Red
                    ElseIf ComboBox13.Text = "Green" Then
                        Chart1.Series(0).Color = Color.Green
                    ElseIf ComboBox13.Text = "Blue" Then
                        Chart1.Series(0).Color = Color.Blue
                    Else
                        Chart1.Series(0).Color = Color.Brown
                    End If
                    If ComboBox14.Text = "Dash" Then
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                        End With
                    Else
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                        End With
                    End If

                    wait(0.001)
                End If
                Me.Chart1.ChartAreas(0).AxisY.LabelStyle.Format = "0.00E0"
                Me.Chart1.ChartAreas(0).AxisX.LabelStyle.Format = "0.00E0"
                Chart1.ChartAreas("ChartArea1").AxisX.Title = ListView5.Columns(1).Text
                Chart1.ChartAreas("ChartArea1").AxisY.Title = ListView5.Columns(2).Text
                TextBox27.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum), "0.00E0")
                TextBox28.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum), "0.00E0")
                TextBox93.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum), "0.00E0")
                TextBox94.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum), "0.00E0")
            Next

            Button14.PerformClick()
            Button19.PerformClick()
            Button32.PerformClick()
            Button12.PerformClick()
            TextBox34.Text = Format(CDbl(TextBox34.Text), "0.00E0")
            TextBox35.Text = Format(CDbl(TextBox35.Text), "0.00E0")
            TextBox31.Text = Format(CDbl(TextBox31.Text), "0.00E0")
            TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
            TextBox45.Text = Format(CDbl(TextBox45.Text), "0.00E0")
            TextBox44.Text = Format(CDbl(TextBox44.Text), "0.00E0")
            TextBox43.Text = Format(CDbl(TextBox43.Text), "0.00E0")
            TextBox29.Text = Format(CDbl(TextBox29.Text), "0.00E0")
            TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
            TextBox33.Text = Format(CDbl(TextBox33.Text), "0.00E0")

        Catch t As Exception
        End Try

        'msgbox("finish")

    End Sub



    Private Sub Button153_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To ListView8.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = ListView8.Columns(j).Text.ToString
            col = col + 1
        Next


        For i = 0 To ListView3.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = CDec(ListView8.Items.Item(i).Text.ToString)
            xlWorkSheet.Cells(i + 2, 2) = Format(CDec(ListView8.Items.Item(i).SubItems(1).Text), "0.0000")
            xlWorkSheet.Cells(i + 2, 3) = CDec(ListView8.Items.Item(i).SubItems(2).Text)
            xlWorkSheet.Cells(i + 2, 4) = CDec(ListView8.Items.Item(i).SubItems(3).Text)

        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = " "
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()


    End Sub

    Private Sub Button69_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To ListView4.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = ListView4.Columns(j).Text.ToString
            col = col + 1
        Next


        For i = 0 To ListView3.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = CDec(ListView4.Items.Item(i).Text.ToString)
            xlWorkSheet.Cells(i + 2, 2) = Format(CDec(ListView4.Items.Item(i).SubItems(1).Text), "0.0000")
            xlWorkSheet.Cells(i + 2, 3) = CDec(ListView4.Items.Item(i).SubItems(2).Text)
            xlWorkSheet.Cells(i + 2, 4) = CDec(ListView4.Items.Item(i).SubItems(3).Text)

        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = " "
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()


    End Sub

    Private Sub Button162_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button162.Click
        Dim tk As String
        Dim P_H2 = TextBox4.Text * 101325
        Chart1.Series(0).Name = "P Out"
        Dim i As String
        Dim P_air As String = TextBox5.Text * 101325

        Dim alpha As String = TextBox9.Text
        Dim F As String = TextBox2.Text
        Dim R As String = TextBox1.Text

        Dim io As String = TextBox11.Text ^ TextBox12.Text

        Dim rin As String = TextBox8.Text

        Dim Bt As String = TextBox13.Text

        Dim Alpha1 As String = TextBox10.Text
        Dim k As String = TextBox17.Text
        Dim Gf_liq As String = TextBox14.Text
        Nam = " Watts"

        Chart1.Series(0).Points.Clear()
        ListView9.Items.Clear()
        tk = TextBox3.Text + 273.15
        Tc = TextBox3.Text
        ' Create loop for current 
        'loop=1;
        'i=0; 
        Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN

        Try
            z = TextBox25.Text

            For Me.baris = z + 1 To ListView3.Items.Count 'Step 0.01

                l = Me.ListView9.Items.Add("")
                For j As Integer = 1 To Me.ListView9.Columns.Count
                    l.SubItems.Add("")
                Next

                For Me.iterasi = 2 To tipeA

                    ListView9.Items(baris - 1).SubItems(0).Text = baris

                    i = CStr((baris - 1) * TextBox26.Text) '* Math.Sqrt(2)))
                    'Calculation of Partial Pressures 
                    'Calculation of saturation pressure of water 

                    x = -2.1794 + (0.02953 * CDec(Tc)) - CDec(9.1837 * (10 ^ -5) * (Tc ^ 2)) + (1.4454 * (10 ^ -7) * (CDec(Tc) ^ 3))
                    P_H2O = (10 ^ x) * 101325
                    'Calculation of partial pressure of hydrogen 
                    pp_H2 = (0.5 * CDec((P_H2) / (Math.Exp(1.653 * i / (tk ^ 1.334))) - P_H2O))

                    ListView9.Items(baris - 1).SubItems(1).Text = CDec(pp_H2)
                    ListView9.Items(baris - 1).SubItems(2).Text = ListView3.Items(baris - 1).SubItems(2).Text
                    If Not baris < 2 Then
                        ListView9.Items(baris - 1).SubItems(3).Text = Val(ListView9.Items(baris - 1).SubItems(2).Text) - Val(ListView9.Items(baris - 2).SubItems(2).Text)
                    Else
                        ListView9.Items(baris - 1).SubItems(3).Text = 0
                    End If
                    If ComboBox12.Text = "Real Time" Then

                        Chart1.Series(0).Points.AddXY(CDec(ListView9.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView9.Items(baris - 1).SubItems(2).Text))
                        If ComboBox17.Text = "Point" Then
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                        ElseIf ComboBox17.Text = "Bar" Then
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                        ElseIf ComboBox17.Text = "Area" Then
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                        ElseIf ComboBox17.Text = "Fast Line" Then
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                        Else
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                        End If

                        If ComboBox13.Text = "Red" Then
                            Chart1.Series(0).Color = Color.Red
                        ElseIf ComboBox13.Text = "Green" Then
                            Chart1.Series(0).Color = Color.Green
                        ElseIf ComboBox13.Text = "Blue" Then
                            Chart1.Series(0).Color = Color.Blue
                        Else
                            Chart1.Series(0).Color = Color.Brown
                        End If
                        If ComboBox14.Text = "Dash" Then
                            With Chart1.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                        Else
                            With Chart1.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                        End If

                        wait(0.001)
                    End If

                    Chart1.ChartAreas("ChartArea1").AxisX.Title = ListView9.Columns(1).Text
                    Chart1.ChartAreas("ChartArea1").AxisY.Title = ListView9.Columns(2).Text

                Next
                TextBox27.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum), "0.00E0")
                TextBox28.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum), "0.00E0")
                TextBox93.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum), "0.00E0")
                TextBox94.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum), "0.00E0")
            Next


            Button14.PerformClick()
            Button19.PerformClick()
            Button32.PerformClick()
            Button12.PerformClick()
            TextBox34.Text = Format(CDbl(TextBox34.Text), "0.00E0")
            TextBox35.Text = Format(CDbl(TextBox35.Text), "0.00E0")
            TextBox31.Text = Format(CDbl(TextBox31.Text), "0.00E0")
            TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
            TextBox45.Text = Format(CDbl(TextBox45.Text), "0.00E0")
            TextBox44.Text = Format(CDbl(TextBox44.Text), "0.00E0")
            TextBox43.Text = Format(CDbl(TextBox43.Text), "0.00E0")
            TextBox29.Text = Format(CDbl(TextBox29.Text), "0.00E0")
            TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
            TextBox33.Text = Format(CDbl(TextBox33.Text), "0.00E0")

        Catch t As Exception
        End Try

        'msgbox("finish")




    End Sub

    Private Sub Button174_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        Dim col As Integer = 1
        For j As Integer = 0 To ListView9.Columns.Count - 1
            xlWorkSheet.Cells(1, col) = ListView9.Columns(j).Text.ToString
            col = col + 1
        Next


        For i = 0 To ListView3.Items.Count - 1
            xlWorkSheet.Cells(i + 2, 1) = CDec(ListView9.Items.Item(i).Text.ToString)
            xlWorkSheet.Cells(i + 2, 2) = Format(CDec(ListView9.Items.Item(i).SubItems(1).Text), "0.0000")
            xlWorkSheet.Cells(i + 2, 3) = CDec(ListView9.Items.Item(i).SubItems(2).Text)
            xlWorkSheet.Cells(i + 2, 4) = CDec(ListView9.Items.Item(i).SubItems(3).Text)

        Next
        Dim dlg As New SaveFileDialog
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FilterIndex = 1
        dlg.InitialDirectory = My.Application.Info.DirectoryPath & "\EXCEL\\EICHER\BILLS\"
        dlg.FileName = " "
        Dim ExcelFile As String = ""
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            ExcelFile = dlg.FileName
            xlWorkSheet.SaveAs(ExcelFile)
        End If
        xlWorkBook.Close()

        xlApp.Quit()


    End Sub

    Private Sub Button99_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button99.Click
        Dim tk As String
        Dim P_H2 = TextBox4.Text * 101325
        Nam = " Watts"
        ' Dim i As String
        Dim P_air As String = TextBox5.Text * 101325
        Chart1.Series(0).Name = "P Out"
        Dim alpha As String = TextBox9.Text
        Dim F As String = TextBox2.Text
        Dim R As String = TextBox1.Text
        Dim H_G As String
        Dim io As String = TextBox11.Text ^ TextBox12.Text

        Dim rin As String = TextBox8.Text

        Dim Bt As String = TextBox13.Text
        Dim E_HHV As String = TextBox22.Text
        Dim Alpha1 As String = TextBox10.Text
        Dim k As String = TextBox17.Text
        Dim Gf_liq As String = TextBox14.Text


        Chart1.Series(0).Points.Clear()
        ListView6.Items.Clear()
        tk = TextBox3.Text + 273.15
        Tc = TextBox3.Text
        ' Create loop for current 
        'loop=1;
        'i=0; 
        Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN

        Try
            z = TextBox25.Text

            For Me.baris = z + 1 To ListView3.Items.Count 'Step 0.01

                l = Me.ListView6.Items.Add("")
                For j As Integer = 1 To Me.ListView6.Columns.Count
                    l.SubItems.Add("")
                Next

                For Me.iterasi = 2 To tipeA

                    ListView6.Items(baris - 1).SubItems(0).Text = baris
                    If ListView2.Items(baris - 1).SubItems(2).Text < 0 Then
                        H_G = 0
                    Else
                        H_G = CDec(ListView2.Items(baris - 1).SubItems(1).Text * TextBox7.Text * (E_HHV - ListView2.Items(baris - 1).SubItems(2).Text))
                    End If

                    ListView6.Items(baris - 1).SubItems(1).Text = CDec(ListView3.Items(baris - 1).SubItems(2).Text)
                    ListView6.Items(baris - 1).SubItems(2).Text = CDec(H_G)
                    If Not baris < 2 Then
                        ListView6.Items(baris - 1).SubItems(3).Text = CDec(ListView6.Items(baris - 1).SubItems(2).Text) - CDec(ListView6.Items(baris - 2).SubItems(2).Text)
                    Else
                        ListView6.Items(baris - 1).SubItems(3).Text = 0
                    End If
                    If ComboBox12.Text = "Real Time" Then

                        Chart1.Series(0).Points.AddXY(CDec(ListView6.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView6.Items(baris - 1).SubItems(2).Text))
                        If ComboBox17.Text = "Point" Then
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                        ElseIf ComboBox17.Text = "Bar" Then
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                        ElseIf ComboBox17.Text = "Area" Then
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                        ElseIf ComboBox17.Text = "Fast Line" Then
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                        Else
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                        End If

                        If ComboBox13.Text = "Red" Then
                            Chart1.Series(0).Color = Color.Red
                        ElseIf ComboBox13.Text = "Green" Then
                            Chart1.Series(0).Color = Color.Green
                        ElseIf ComboBox13.Text = "Blue" Then
                            Chart1.Series(0).Color = Color.Blue
                        Else
                            Chart1.Series(0).Color = Color.Brown
                        End If
                        If ComboBox14.Text = "Dash" Then
                            With Chart1.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                        Else
                            With Chart1.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                        End If

                        wait(0.001)
                    End If

                    Chart1.ChartAreas("ChartArea1").AxisX.Title = ListView6.Columns(1).Text
                    Chart1.ChartAreas("ChartArea1").AxisY.Title = ListView6.Columns(2).Text

                Next
                TextBox27.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum), "0.00E0")
                TextBox28.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum), "0.00E0")
                TextBox93.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum), "0.00E0")
                TextBox94.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum), "0.00E0")
            Next


            Button14.PerformClick()
            Button19.PerformClick()
            Button32.PerformClick()
            Button12.PerformClick()
            TextBox34.Text = Format(CDbl(TextBox34.Text), "0.00E0")
            TextBox35.Text = Format(CDbl(TextBox35.Text), "0.00E0")
            TextBox31.Text = Format(CDbl(TextBox31.Text), "0.00E0")
            TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
            TextBox45.Text = Format(CDbl(TextBox45.Text), "0.00E0")
            TextBox44.Text = Format(CDbl(TextBox44.Text), "0.00E0")
            TextBox43.Text = Format(CDbl(TextBox43.Text), "0.00E0")
            TextBox29.Text = Format(CDbl(TextBox29.Text), "0.00E0")
            TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
            TextBox33.Text = Format(CDbl(TextBox33.Text), "0.00E0")

        Catch t As Exception
        End Try

        'msgbox("finish")
    End Sub

    Private Sub Button120_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button120.Click

        Nam = "Volts"
        Dim tk As String
        Dim P_H2 = TextBox4.Text * 101325
        Dim Eff_HHV As String
        Dim UF As String = TextBox16.Text
        Dim E_HHV As String = TextBox22.Text
        'Dim i As String
        Dim P_air As String = TextBox5.Text * 101325

        Dim alpha As String = TextBox9.Text
        Dim F As String = TextBox2.Text
        Dim R As String = TextBox1.Text

        Dim io As String = TextBox11.Text ^ TextBox12.Text

        Dim rin As String = TextBox8.Text

        Dim Bt As String = TextBox13.Text

        Dim Alpha1 As String = TextBox10.Text
        Dim k As String = TextBox17.Text
        Dim Gf_liq As String = TextBox14.Text

        Chart1.Series(0).Points.Clear()
        ListView7.Items.Clear()
        tk = TextBox3.Text + 273.15
        Tc = TextBox3.Text
        ' Create loop for current 
        'loop=1;
        'i=0; 
        Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN

        Try
            z = TextBox25.Text

            For Me.baris = z + 1 To ListView3.Items.Count 'Step 0.01

                l = Me.ListView7.Items.Add("")
                For j As Integer = 1 To Me.ListView7.Columns.Count
                    l.SubItems.Add("")
                Next

                For Me.iterasi = 2 To tipeA

                    ListView7.Items(baris - 1).SubItems(0).Text = baris



                    If ListView2.Items(baris - 1).SubItems(2).Text < 0 Then
                        Eff_HHV = 0
                    Else
                        Eff_HHV = (UF * ListView2.Items(baris - 1).SubItems(2).Text * 100) / (E_HHV)
                    End If
                    ListView7.Items(baris - 1).SubItems(1).Text = CDec(ListView2.Items(baris - 1).SubItems(1).Text)
                    ListView7.Items(baris - 1).SubItems(2).Text = (CDec(Eff_HHV))
                    If Not baris < 2 Then
                        ListView7.Items(baris - 1).SubItems(3).Text = CDec(ListView7.Items(baris - 1).SubItems(2).Text) - CDec(ListView7.Items(baris - 2).SubItems(2).Text)
                    Else
                        ListView7.Items(baris - 1).SubItems(3).Text = 0
                    End If
                    If ComboBox12.Text = "Real Time" Then

                        Chart1.Series(0).Points.AddXY(CDec(ListView7.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView7.Items(baris - 1).SubItems(2).Text))
                        If ComboBox17.Text = "Point" Then
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                        ElseIf ComboBox17.Text = "Bar" Then
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                        ElseIf ComboBox17.Text = "Area" Then
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                        ElseIf ComboBox17.Text = "Fast Line" Then
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                        Else
                            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                        End If

                        If ComboBox13.Text = "Red" Then
                            Chart1.Series(0).Color = Color.Red
                        ElseIf ComboBox13.Text = "Green" Then
                            Chart1.Series(0).Color = Color.Green
                        ElseIf ComboBox13.Text = "Blue" Then
                            Chart1.Series(0).Color = Color.Blue
                        Else
                            Chart1.Series(0).Color = Color.Brown
                        End If
                        If ComboBox14.Text = "Dash" Then
                            With Chart1.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                        Else
                            With Chart1.ChartAreas(0)
                                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                                '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                            End With
                        End If

                        wait(0.001)
                    End If

                    Chart1.ChartAreas("ChartArea1").AxisX.Title = ListView7.Columns(1).Text
                    Chart1.ChartAreas("ChartArea1").AxisY.Title = ListView7.Columns(2).Text

                Next
                TextBox27.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum), "0.00E0")
                TextBox28.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum), "0.00E0")
                TextBox93.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum), "0.00E0")
                TextBox94.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum), "0.00E0")
            Next

            Button14.PerformClick()
            Button19.PerformClick()
            Button32.PerformClick()
            Button12.PerformClick()
            TextBox34.Text = Format(CDbl(TextBox34.Text), "0.00E0")
            TextBox35.Text = Format(CDbl(TextBox35.Text), "0.00E0")
            TextBox31.Text = Format(CDbl(TextBox31.Text), "0.00E0")
            TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
            TextBox45.Text = Format(CDbl(TextBox45.Text), "0.00E0")
            TextBox44.Text = Format(CDbl(TextBox44.Text), "0.00E0")
            TextBox43.Text = Format(CDbl(TextBox43.Text), "0.00E0")
            TextBox29.Text = Format(CDbl(TextBox29.Text), "0.00E0")
            TextBox33.Text = Format(CDbl(TextBox33.Text), "0.00E0")

        Catch t As Exception
        End Try

        'msgbox("finish")
    End Sub
    Private Sub Button14_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim i As Integer
        Dim g As Integer
        Try
            Button39.PerformClick()
            If TextBox32.Text = "Mode : Fuel Cell Polarization" Then

                With ListView2
                    TextBox35.Text = .Items(0).SubItems(2).Text
                    For g = 2 To .Items.Count
                        If (.Items(g - 2).SubItems(2).Text) < (.Items(g - 1).SubItems(2).Text) And TextBox35.Text < (.Items(g - 1).SubItems(2).Text) Then
                            TextBox35.Text = (.Items(g - 1).SubItems(2).Text)
                        Else
                            TextBox35.Text = TextBox35.Text
                        End If
                    Next
                    TextBox34.Text = .Items(0).SubItems(2).Text
                    For i = 2 To .Items.Count
                        If CDec(.Items(i - 1).SubItems(2).Text) < CDec(.Items(i - 2).SubItems(2).Text) And TextBox34.Text > CDec(.Items(i - 1).SubItems(2).Text) Then
                            TextBox34.Text = CDec(.Items(i - 1).SubItems(2).Text)
                        Else
                            TextBox34.Text = Val(TextBox34.Text)
                        End If
                    Next i
                    TextBox33.Text = TextBox35.Text - TextBox34.Text
                    TextBox45.Text = .Items(.Items.Count - 1).SubItems(1).Text
                    TextBox31.Text = Math.Abs(CDec(.Items(0).SubItems(3).Text))
                    For g = 2 To .Items.Count
                        If Math.Abs(CDec(.Items(g - 2).SubItems(3).Text)) < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) And TextBox31.Text < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) Then
                            TextBox31.Text = Math.Abs(CDec(.Items(g - 1).SubItems(3).Text))
                            TextBox44.BackColor = Color.LightBlue
                            TextBox44.Text = Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0")
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = Val("-" + TextBox31.Text)
                            End If
                        Else
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = Val("-" + TextBox31.Text)
                            Else
                                TextBox31.Text = Val(TextBox31.Text)
                            End If
                        End If
                    Next g
                    TextBox29.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For i = 3 To .Items.Count
                        If Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) < Math.Abs(CDec(.Items(i - 2).SubItems(3).Text)) And TextBox29.Text > Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) Then
                            TextBox29.Text = Math.Abs(CDec(.Items(i - 2).SubItems(3).Text))
                            TextBox43.BackColor = Color.Yellow
                            TextBox43.Text = Format(Val(.Items(i - 2).SubItems(2).Text), "0.00E0")
                        Else
                            TextBox29.Text = Val(TextBox29.Text)
                        End If
                    Next i

                    wait(0.001)
                    TextBox30.Text = "0"
                    For i = 1 To .Items.Count
                        TextBox30.Text = TextBox30.Text + CDec(.Items(i - 1).SubItems(2).Text)
                    Next i
                    TextBox30.Text = TextBox30.Text / .Items.Count
                    For g = 1 To .Items.Count
                        If CDec(.Items(g - 1).SubItems(2).Text) = CDec(TextBox35.Text) Then
                            .Items(g - 1).BackColor = Color.LightGreen
                            TextBox35.BackColor = Color.LightGreen
                        ElseIf .Items(g - 1).SubItems(3).Text = Val(TextBox29.Text) Then
                            .Items(g - 1).BackColor = Color.Yellow
                            TextBox29.BackColor = Color.Yellow
                        ElseIf Val(.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then
                            .Items(g - 1).BackColor = Color.LightBlue
                            TextBox31.BackColor = Color.LightBlue
                        ElseIf .Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                            .Items(g - 1).BackColor = Color.Orange
                            TextBox34.BackColor = Color.Orange
                        ElseIf .Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                            .Items(g - 1).BackColor = Color.LightCyan
                            TextBox30.BackColor = Color.LightCyan
                        End If
                    Next g
                    Chart1.ChartAreas(0).AxisX.LineWidth = 1
                End With
            ElseIf TextBox32.Text = "Mode : Fuel Cell Power" Then
                With ListView3
                    TextBox35.Text = .Items(0).SubItems(2).Text
                    For g = 2 To .Items.Count
                        If CDec(.Items(g - 2).SubItems(2).Text) < CDec(.Items(g - 1).SubItems(2).Text) And TextBox35.Text < CDec(.Items(g - 1).SubItems(2).Text) Then
                            TextBox35.Text = (.Items(g - 1).SubItems(2).Text)
                        Else
                            TextBox35.Text = TextBox35.Text
                        End If
                    Next
                    TextBox34.Text = .Items(0).SubItems(2).Text
                    For i = 2 To .Items.Count
                        If CDec(.Items(i - 1).SubItems(2).Text) < CDec(.Items(i - 2).SubItems(2).Text) And TextBox34.Text > CDec(.Items(i - 1).SubItems(2).Text) Then
                            TextBox34.Text = CDec(.Items(i - 1).SubItems(2).Text)
                        Else
                            TextBox34.Text = Val(TextBox34.Text)
                        End If
                    Next i
                    TextBox33.Text = TextBox35.Text - TextBox34.Text
                    TextBox45.Text = .Items(.Items.Count - 1).SubItems(1).Text
                    TextBox31.Text = Math.Abs(CDec(.Items(0).SubItems(3).Text))
                    For g = 2 To .Items.Count
                        If Math.Abs(CDec(.Items(g - 2).SubItems(3).Text)) < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) And TextBox31.Text < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) Then
                            TextBox31.Text = Math.Abs(CDec(.Items(g - 1).SubItems(3).Text))
                            TextBox44.BackColor = Color.LightBlue
                            TextBox44.Text = Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0")
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = Val("-" + TextBox31.Text)
                            End If
                        Else
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = Val("-" + TextBox31.Text)
                            Else
                                TextBox31.Text = Val(TextBox31.Text)
                            End If
                        End If
                    Next g
                    TextBox29.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For i = 3 To .Items.Count
                        If Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) < Math.Abs(CDec(.Items(i - 2).SubItems(3).Text)) And TextBox29.Text > Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) Then
                            TextBox29.Text = Math.Abs(CDec(.Items(i - 2).SubItems(3).Text))
                            TextBox43.BackColor = Color.Yellow
                            TextBox43.Text = Format(Val(.Items(i - 2).SubItems(2).Text), "0.00E0")
                        Else
                            TextBox29.Text = Val(TextBox29.Text)
                        End If
                    Next i

                    wait(0.001)
                    TextBox30.Text = "0"
                    For i = 1 To .Items.Count
                        TextBox30.Text = TextBox30.Text + CDec(.Items(i - 1).SubItems(2).Text)
                    Next i
                    TextBox30.Text = TextBox30.Text / .Items.Count
                    For g = 1 To .Items.Count
                        If CDec(.Items(g - 1).SubItems(2).Text) = CDec(TextBox35.Text) Then
                            .Items(g - 1).BackColor = Color.LightGreen
                            TextBox35.BackColor = Color.LightGreen
                        ElseIf .Items(g - 1).SubItems(3).Text = Val(TextBox29.Text) Then
                            .Items(g - 1).BackColor = Color.Yellow
                            TextBox29.BackColor = Color.Yellow
                        ElseIf Val(.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then
                            .Items(g - 1).BackColor = Color.LightBlue
                            TextBox31.BackColor = Color.LightBlue
                        ElseIf .Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                            .Items(g - 1).BackColor = Color.Orange
                            TextBox34.BackColor = Color.Orange
                        ElseIf .Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                            .Items(g - 1).BackColor = Color.LightCyan
                            TextBox30.BackColor = Color.LightCyan
                        End If
                    Next g
                    Chart1.ChartAreas(0).AxisX.LineWidth = 1
                End With
            ElseIf TextBox32.Text = "Mode : Hydrogen Usage" Then

                With ListView4
                    TextBox35.Text = .Items(0).SubItems(2).Text
                    For g = 2 To .Items.Count
                        If CDec(.Items(g - 2).SubItems(2).Text) < CDec(.Items(g - 1).SubItems(2).Text) And TextBox35.Text < CDec(.Items(g - 1).SubItems(2).Text) Then
                            TextBox35.Text = (.Items(g - 1).SubItems(2).Text)
                        Else
                            TextBox35.Text = TextBox35.Text
                        End If
                    Next
                    TextBox34.Text = .Items(0).SubItems(2).Text
                    For i = 2 To .Items.Count
                        If CDec(.Items(i - 1).SubItems(2).Text) < CDec(.Items(i - 2).SubItems(2).Text) And TextBox34.Text > CDec(.Items(i - 1).SubItems(2).Text) Then
                            TextBox34.Text = CDec(.Items(i - 1).SubItems(2).Text)
                        Else
                            TextBox34.Text = Val(TextBox34.Text)
                        End If
                    Next i
                    TextBox33.Text = TextBox35.Text - TextBox34.Text
                    TextBox45.Text = .Items(.Items.Count - 1).SubItems(1).Text
                    TextBox31.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For g = 2 To .Items.Count
                        If Math.Abs(CDec(.Items(g - 2).SubItems(3).Text)) < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) And TextBox31.Text < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) Then
                            TextBox31.Text = Math.Abs(CDec(.Items(g - 1).SubItems(3).Text))
                            TextBox44.BackColor = Color.LightBlue
                            TextBox44.Text = (CDec(.Items(g - 1).SubItems(2).Text))
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = Val("-" + TextBox31.Text)
                            End If
                        Else
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = Val("-" + TextBox31.Text)
                            Else
                                TextBox31.Text = Val(TextBox31.Text)
                            End If
                        End If
                    Next g
                    TextBox29.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For i = 3 To .Items.Count
                        If Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) < Math.Abs(CDec(.Items(i - 2).SubItems(3).Text)) And TextBox29.Text > Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) Then
                            TextBox29.Text = Math.Abs(CDec(.Items(i - 2).SubItems(3).Text))
                            TextBox43.BackColor = Color.Yellow
                            TextBox43.Text = (CDec(.Items(i - 2).SubItems(2).Text))
                        Else
                            TextBox29.Text = CDec(TextBox29.Text)
                        End If
                    Next i

                    wait(0.001)
                    TextBox30.Text = "0"
                    For i = 1 To .Items.Count
                        TextBox30.Text = TextBox30.Text + CDec(.Items(i - 1).SubItems(2).Text)
                    Next i
                    TextBox30.Text = TextBox30.Text / .Items.Count
                    For g = 1 To .Items.Count
                        If CDec(.Items(g - 1).SubItems(2).Text) = CDec(TextBox35.Text) Then
                            .Items(g - 1).BackColor = Color.LightGreen
                            TextBox35.BackColor = Color.LightGreen
                        ElseIf CDec(.Items(g - 1).SubItems(3).Text) = CDec(TextBox29.Text) Then
                            .Items(g - 1).BackColor = Color.Yellow
                            TextBox29.BackColor = Color.Yellow
                        ElseIf Val(.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then
                            .Items(g - 1).BackColor = Color.LightBlue
                            TextBox31.BackColor = Color.LightBlue
                        ElseIf .Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                            .Items(g - 1).BackColor = Color.Orange
                            TextBox34.BackColor = Color.Orange
                        ElseIf .Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                            .Items(g - 1).BackColor = Color.LightCyan
                            TextBox30.BackColor = Color.LightCyan
                        End If
                    Next g
                    Chart1.ChartAreas(0).AxisX.LineWidth = 1
                End With

            ElseIf TextBox32.Text = "Mode : Power Vs Hydrogen" Then

                With ListView8
                    TextBox35.Text = CDec(.Items(0).SubItems(2).Text)
                    For g = 2 To .Items.Count
                        If CDec(.Items(g - 2).SubItems(2).Text) < CDec(.Items(g - 1).SubItems(2).Text) And CDec(TextBox35.Text) < CDec(.Items(g - 1).SubItems(2).Text) Then
                            TextBox35.Text = CDec(.Items(g - 1).SubItems(2).Text)
                        Else
                            TextBox35.Text = CDec(TextBox35.Text)
                        End If
                    Next
                    TextBox34.Text = .Items(0).SubItems(2).Text
                    For i = 2 To .Items.Count
                        If CDec(.Items(i - 1).SubItems(2).Text) < CDec(.Items(i - 2).SubItems(2).Text) And TextBox34.Text > CDec(.Items(i - 1).SubItems(2).Text) Then
                            TextBox34.Text = CDec(.Items(i - 1).SubItems(2).Text)
                        Else
                            TextBox34.Text = CDec(TextBox34.Text)
                        End If
                    Next i
                    TextBox33.Text = CDec(TextBox35.Text - TextBox34.Text)
                    TextBox45.Text = CDec(.Items(.Items.Count - 1).SubItems(1).Text)
                    TextBox31.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For g = 2 To .Items.Count
                        If Math.Abs(CDec(.Items(g - 2).SubItems(3).Text)) < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) And TextBox31.Text < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) Then
                            TextBox31.Text = Math.Abs(CDec(.Items(g - 1).SubItems(3).Text))
                            TextBox44.BackColor = Color.LightBlue
                            TextBox44.Text = (CDec(.Items(g - 1).SubItems(2).Text))
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = CDec("-" + TextBox31.Text)
                            End If
                        Else
                            If CDec(.Items(g - 1).SubItems(3).Text) < 0 Then
                                TextBox31.Text = CDec("-" + TextBox31.Text)
                            Else
                                TextBox31.Text = CDec(TextBox31.Text)
                            End If
                        End If
                    Next g
                    TextBox29.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For i = 3 To .Items.Count
                        If Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) < Math.Abs(CDec(.Items(i - 2).SubItems(3).Text)) And TextBox29.Text > Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) Then
                            TextBox29.Text = Math.Abs(CDec(.Items(i - 2).SubItems(3).Text))
                            TextBox43.BackColor = Color.Yellow
                            TextBox43.Text = (CDec(.Items(i - 2).SubItems(2).Text))
                        Else
                            TextBox29.Text = CDec(TextBox29.Text)
                        End If
                    Next i

                    wait(0.001)
                    TextBox30.Text = "0"
                    For i = 1 To .Items.Count
                        TextBox30.Text = TextBox30.Text + CDec(.Items(i - 1).SubItems(2).Text)
                    Next i
                    TextBox30.Text = CDec(TextBox30.Text / .Items.Count)
                    For g = 1 To .Items.Count
                        If CDec(.Items(g - 1).SubItems(2).Text) = CDec(TextBox35.Text) Then
                            .Items(g - 1).BackColor = Color.LightGreen
                            TextBox35.BackColor = Color.LightGreen
                        ElseIf .Items(g - 1).SubItems(3).Text = CDec(TextBox29.Text) Then
                            .Items(g - 1).BackColor = Color.Yellow
                            TextBox29.BackColor = Color.Yellow
                        ElseIf Val(.Items(g - 1).SubItems(3).Text) = CDec(TextBox31.Text) Then
                            .Items(g - 1).BackColor = Color.LightBlue
                            TextBox31.BackColor = Color.LightBlue
                        ElseIf .Items(g - 1).SubItems(2).Text = CDec(TextBox34.Text) Then
                            .Items(g - 1).BackColor = Color.Orange
                            TextBox34.BackColor = Color.Orange
                        ElseIf .Items(g - 1).SubItems(2).Text = CDec(TextBox30.Text) Then
                            .Items(g - 1).BackColor = Color.LightCyan
                            TextBox30.BackColor = Color.LightCyan
                        End If
                    Next g
                    Chart1.ChartAreas(0).AxisX.LineWidth = 1

                End With

            ElseIf TextBox32.Text = "Mode : Power VS Partial Pressure" Then
                With ListView9
                    TextBox35.Text = CDec(.Items(0).SubItems(2).Text)
                    For g = 2 To .Items.Count
                        If CDec(.Items(g - 2).SubItems(2).Text) < CDec(.Items(g - 1).SubItems(2).Text) And CDec(TextBox35.Text) < CDec(.Items(g - 1).SubItems(2).Text) Then
                            TextBox35.Text = CDec(.Items(g - 1).SubItems(2).Text)
                        Else
                            TextBox35.Text = CDec(TextBox35.Text)
                        End If
                    Next
                    TextBox34.Text = .Items(0).SubItems(2).Text
                    For i = 2 To .Items.Count
                        If CDec(.Items(i - 1).SubItems(2).Text) < CDec(.Items(i - 2).SubItems(2).Text) And TextBox34.Text > CDec(.Items(i - 1).SubItems(2).Text) Then
                            TextBox34.Text = CDec(.Items(i - 1).SubItems(2).Text)
                        Else
                            TextBox34.Text = CDec(TextBox34.Text)
                        End If
                    Next i
                    TextBox33.Text = CDec(TextBox35.Text - TextBox34.Text)
                    TextBox45.Text = CDec(.Items(.Items.Count - 1).SubItems(1).Text)
                    TextBox31.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For g = 2 To .Items.Count
                        If Math.Abs(CDec(.Items(g - 2).SubItems(3).Text)) < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) And TextBox31.Text < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) Then
                            TextBox31.Text = Math.Abs(CDec(.Items(g - 1).SubItems(3).Text))
                            TextBox44.BackColor = Color.LightBlue
                            TextBox44.Text = (CDec(.Items(g - 1).SubItems(2).Text))
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = CDec("-" + TextBox31.Text)
                            End If
                        Else
                            If CDec(.Items(g - 1).SubItems(3).Text) < 0 Then
                                TextBox31.Text = CDec("-" + TextBox31.Text)
                            Else
                                TextBox31.Text = CDec(TextBox31.Text)
                            End If
                        End If
                    Next g
                    TextBox29.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For i = 3 To .Items.Count
                        If Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) < Math.Abs(CDec(.Items(i - 2).SubItems(3).Text)) And TextBox29.Text > Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) Then
                            TextBox29.Text = Math.Abs(CDec(.Items(i - 2).SubItems(3).Text))
                            TextBox43.BackColor = Color.Yellow
                            TextBox43.Text = (CDec(.Items(i - 2).SubItems(2).Text))
                        Else
                            TextBox29.Text = CDec(TextBox29.Text)
                        End If
                    Next i

                    wait(0.001)
                    TextBox30.Text = "0"
                    For i = 1 To .Items.Count
                        TextBox30.Text = TextBox30.Text + CDec(.Items(i - 1).SubItems(2).Text)
                    Next i
                    TextBox30.Text = CDec(TextBox30.Text / .Items.Count)
                    For g = 1 To .Items.Count
                        If CDec(.Items(g - 1).SubItems(2).Text) = CDec(TextBox35.Text) Then
                            .Items(g - 1).BackColor = Color.LightGreen
                            TextBox35.BackColor = Color.LightGreen
                        ElseIf .Items(g - 1).SubItems(3).Text = CDec(TextBox29.Text) Then
                            .Items(g - 1).BackColor = Color.Yellow
                            TextBox29.BackColor = Color.Yellow
                        ElseIf Val(.Items(g - 1).SubItems(3).Text) = CDec(TextBox31.Text) Then
                            .Items(g - 1).BackColor = Color.LightBlue
                            TextBox31.BackColor = Color.LightBlue
                        ElseIf .Items(g - 1).SubItems(2).Text = CDec(TextBox34.Text) Then
                            .Items(g - 1).BackColor = Color.Orange
                            TextBox34.BackColor = Color.Orange
                        ElseIf .Items(g - 1).SubItems(2).Text = CDec(TextBox30.Text) Then
                            .Items(g - 1).BackColor = Color.LightCyan
                            TextBox30.BackColor = Color.LightCyan
                        End If
                    Next g
                    Chart1.ChartAreas(0).AxisX.LineWidth = 1
                End With
            ElseIf TextBox32.Text = "Mode : Water Produced" Then
                With ListView5
                    TextBox35.Text = .Items(0).SubItems(2).Text
                    For g = 2 To .Items.Count
                        If (.Items(g - 2).SubItems(2).Text) < (.Items(g - 1).SubItems(2).Text) And TextBox35.Text < (.Items(g - 1).SubItems(2).Text) Then
                            TextBox35.Text = (.Items(g - 1).SubItems(2).Text)
                        Else
                            TextBox35.Text = TextBox35.Text
                        End If
                    Next
                    TextBox34.Text = .Items(0).SubItems(2).Text
                    For i = 2 To .Items.Count
                        If CDec(.Items(i - 1).SubItems(2).Text) < CDec(.Items(i - 2).SubItems(2).Text) And TextBox34.Text > CDec(.Items(i - 1).SubItems(2).Text) Then
                            TextBox34.Text = CDec(.Items(i - 1).SubItems(2).Text)
                        Else
                            TextBox34.Text = Val(TextBox34.Text)
                        End If
                    Next i
                    TextBox33.Text = TextBox35.Text - TextBox34.Text
                    TextBox45.Text = .Items(.Items.Count - 1).SubItems(1).Text
                    TextBox31.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For g = 2 To .Items.Count
                        If Math.Abs(CDec(.Items(g - 2).SubItems(3).Text)) < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) And TextBox31.Text < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) Then
                            TextBox31.Text = Math.Abs(CDec(.Items(g - 1).SubItems(3).Text))
                            TextBox44.BackColor = Color.LightBlue
                            TextBox44.Text = Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0")
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = Val("-" + TextBox31.Text)
                            End If
                        Else
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = Val("-" + TextBox31.Text)
                            Else
                                TextBox31.Text = Val(TextBox31.Text)
                            End If
                        End If
                    Next g
                    TextBox29.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For i = 3 To .Items.Count
                        If Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) < Math.Abs(CDec(.Items(i - 2).SubItems(3).Text)) And TextBox29.Text > Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) Then
                            TextBox29.Text = Math.Abs(CDec(.Items(i - 2).SubItems(3).Text))
                            TextBox43.BackColor = Color.Yellow
                            TextBox43.Text = Format(Val(.Items(i - 2).SubItems(2).Text), "0.00E0")
                        Else
                            TextBox29.Text = Val(TextBox29.Text)
                        End If
                    Next i

                    wait(0.001)
                    TextBox30.Text = "0"
                    For i = 1 To .Items.Count
                        TextBox30.Text = TextBox30.Text + CDec(.Items(i - 1).SubItems(2).Text)
                    Next i
                    TextBox30.Text = TextBox30.Text / .Items.Count
                    For g = 1 To .Items.Count
                        If CDec(.Items(g - 1).SubItems(2).Text) = CDec(TextBox35.Text) Then
                            .Items(g - 1).BackColor = Color.LightGreen
                            TextBox35.BackColor = Color.LightGreen
                        ElseIf .Items(g - 1).SubItems(3).Text = Val(TextBox29.Text) Then
                            .Items(g - 1).BackColor = Color.Yellow
                            TextBox29.BackColor = Color.Yellow
                        ElseIf Val(.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then
                            .Items(g - 1).BackColor = Color.LightBlue
                            TextBox31.BackColor = Color.LightBlue
                        ElseIf .Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                            .Items(g - 1).BackColor = Color.Orange
                            TextBox34.BackColor = Color.Orange
                        ElseIf .Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                            .Items(g - 1).BackColor = Color.LightCyan
                            TextBox30.BackColor = Color.LightCyan
                        End If
                    Next g
                    Chart1.ChartAreas(0).AxisX.LineWidth = 1
                End With

            ElseIf TextBox32.Text = "Mode : Heat Generated" Then
                With ListView6
                    TextBox35.Text = CDec(.Items(0).SubItems(2).Text)
                    For g = 2 To .Items.Count
                        If CDec(.Items(g - 2).SubItems(2).Text) < CDec(.Items(g - 1).SubItems(2).Text) And CDec(TextBox35.Text) < CDec(.Items(g - 1).SubItems(2).Text) Then
                            TextBox35.Text = CDec(.Items(g - 1).SubItems(2).Text)
                        Else
                            TextBox35.Text = CDec(TextBox35.Text)
                        End If
                    Next
                    TextBox34.Text = .Items(0).SubItems(2).Text
                    For i = 2 To .Items.Count
                        If CDec(.Items(i - 1).SubItems(2).Text) < CDec(.Items(i - 2).SubItems(2).Text) And TextBox34.Text > CDec(.Items(i - 1).SubItems(2).Text) Then
                            TextBox34.Text = CDec(.Items(i - 1).SubItems(2).Text)
                        Else
                            TextBox34.Text = CDec(TextBox34.Text)
                        End If
                    Next i
                    TextBox33.Text = CDec(TextBox35.Text - TextBox34.Text)
                    TextBox45.Text = CDec(.Items(.Items.Count - 1).SubItems(1).Text)
                    TextBox31.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For g = 2 To .Items.Count
                        If Math.Abs(CDec(.Items(g - 2).SubItems(3).Text)) < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) And TextBox31.Text < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) Then
                            TextBox31.Text = Math.Abs(CDec(.Items(g - 1).SubItems(3).Text))
                            TextBox44.BackColor = Color.LightBlue
                            TextBox44.Text = (CDec(.Items(g - 1).SubItems(2).Text))
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = CDec("-" + TextBox31.Text)
                            End If
                        Else
                            If CDec(.Items(g - 1).SubItems(3).Text) < 0 Then
                                TextBox31.Text = CDec("-" + TextBox31.Text)
                            Else
                                TextBox31.Text = CDec(TextBox31.Text)
                            End If
                        End If
                    Next g
                    TextBox29.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For i = 3 To .Items.Count
                        If Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) < Math.Abs(CDec(.Items(i - 2).SubItems(3).Text)) And TextBox29.Text > Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) Then
                            TextBox29.Text = Math.Abs(CDec(.Items(i - 2).SubItems(3).Text))
                            TextBox43.BackColor = Color.Yellow
                            TextBox43.Text = (CDec(.Items(i - 2).SubItems(2).Text))
                        Else
                            TextBox29.Text = CDec(TextBox29.Text)
                        End If
                    Next i

                    wait(0.001)
                    TextBox30.Text = "0"
                    For i = 1 To .Items.Count
                        TextBox30.Text = TextBox30.Text + CDec(.Items(i - 1).SubItems(2).Text)
                    Next i
                    TextBox30.Text = CDec(TextBox30.Text / .Items.Count)
                    For g = 1 To .Items.Count
                        If CDec(.Items(g - 1).SubItems(2).Text) = CDec(TextBox35.Text) Then
                            .Items(g - 1).BackColor = Color.LightGreen
                            TextBox35.BackColor = Color.LightGreen
                        ElseIf CDec(.Items(g - 1).SubItems(3).Text) = CDec(TextBox29.Text) Then
                            .Items(g - 1).BackColor = Color.Yellow
                            TextBox29.BackColor = Color.Yellow
                        ElseIf CDec(.Items(g - 1).SubItems(3).Text) = CDec(TextBox31.Text) Then
                            .Items(g - 1).BackColor = Color.LightBlue
                            TextBox31.BackColor = Color.LightBlue
                        ElseIf CDec(.Items(g - 1).SubItems(2).Text) = CDec(TextBox34.Text) Then
                            .Items(g - 1).BackColor = Color.Orange
                            TextBox34.BackColor = Color.Orange
                        ElseIf CDec(.Items(g - 1).SubItems(2).Text) = CDec(TextBox30.Text) Then
                            .Items(g - 1).BackColor = Color.LightCyan
                            TextBox30.BackColor = Color.LightCyan
                        End If
                    Next g
                    Chart1.ChartAreas(0).AxisX.LineWidth = 1

                End With

            ElseIf TextBox32.Text = "Mode : Efficiency" Then
                With ListView7
                    TextBox35.Text = .Items(0).SubItems(2).Text
                    For g = 2 To .Items.Count
                        If CDec(.Items(g - 2).SubItems(2).Text) < CDec(.Items(g - 1).SubItems(2).Text) And TextBox35.Text < (.Items(g - 1).SubItems(2).Text) Then
                            TextBox35.Text = CDec(.Items(g - 1).SubItems(2).Text)
                        Else
                            TextBox35.Text = TextBox35.Text
                        End If
                    Next
                    TextBox34.Text = .Items(0).SubItems(2).Text
                    For i = 2 To .Items.Count
                        If (CDec(.Items(i - 1).SubItems(2).Text)) < (CDec(.Items(i - 2).SubItems(2).Text)) And Format(CDbl(TextBox34.Text), "0.00E0") > (CDec(.Items(i - 1).SubItems(2).Text)) Then
                            TextBox34.Text = CDec(.Items(i - 1).SubItems(2).Text)
                        ElseIf (CDec(.Items(i - 1).SubItems(2).Text)) > (CDec(.Items(i - 2).SubItems(2).Text)) And Format(CDbl(TextBox34.Text), "0.00E00") > (CDec(.Items(i - 2).SubItems(2).Text)) Then
                            TextBox34.Text = CDec(.Items(i - 2).SubItems(2).Text)
                        Else
                            TextBox34.Text = CDec(TextBox34.Text)
                        End If
                    Next i
                    TextBox33.Text = TextBox35.Text - TextBox34.Text
                    TextBox45.Text = CDec(.Items(.Items.Count - 1).SubItems(1).Text)
                    TextBox31.Text = Math.Abs(CDec(.Items(0).SubItems(3).Text))
                    For g = 2 To .Items.Count
                        If Math.Abs(CDec(.Items(g - 2).SubItems(3).Text)) < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) And TextBox31.Text < Math.Abs(CDec(.Items(g - 1).SubItems(3).Text)) Then
                            TextBox31.Text = Math.Abs(CDec(.Items(g - 1).SubItems(3).Text))
                            TextBox44.BackColor = Color.LightBlue
                            TextBox44.Text = (CDec(.Items(g - 1).SubItems(2).Text))
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = Val("-" + TextBox31.Text)
                            End If
                        Else
                            If .Items(g - 1).SubItems(3).Text < 0 Then
                                TextBox31.Text = Val("-" + TextBox31.Text)
                            Else
                                TextBox31.Text = CDec(TextBox31.Text)
                            End If
                        End If
                    Next g
                    TextBox29.Text = Math.Abs(CDec(.Items(1).SubItems(3).Text))
                    For i = 3 To .Items.Count
                        If Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) < Math.Abs(CDec(.Items(i - 2).SubItems(3).Text)) And TextBox29.Text > Math.Abs(CDec(.Items(i - 1).SubItems(3).Text)) Then
                            TextBox29.Text = Math.Abs(CDec(.Items(i - 2).SubItems(3).Text))
                            TextBox43.BackColor = Color.Yellow
                            TextBox43.Text = (CDec(.Items(i - 2).SubItems(2).Text))
                            .Items(i - 2).SubItems(2).BackColor = Color.Yellow
                        Else
                            TextBox29.Text = CDec(TextBox29.Text)
                        End If
                    Next i

                    wait(0.001)
                    TextBox30.Text = "0"
                    For i = 1 To .Items.Count
                        TextBox30.Text = TextBox30.Text + CDec(.Items(i - 1).SubItems(2).Text)
                    Next i
                    TextBox30.Text = TextBox30.Text / .Items.Count
                    For g = 1 To .Items.Count
                        If CDec(.Items(g - 1).SubItems(2).Text) = CDec(TextBox35.Text) Then
                            .Items(g - 1).BackColor = Color.LightGreen
                            TextBox35.BackColor = Color.LightGreen
                        ElseIf CDec(.Items(g - 1).SubItems(3).Text) = CDec(TextBox29.Text) Then
                            .Items(g - 1).BackColor = Color.Yellow
                            TextBox29.BackColor = Color.Yellow
                        ElseIf CDec(.Items(g - 1).SubItems(3).Text) = CDec(TextBox31.Text) Then
                            .Items(g - 1).BackColor = Color.LightBlue
                            TextBox31.BackColor = Color.LightBlue
                        ElseIf CDec(.Items(g - 1).SubItems(2).Text) = CDec(TextBox34.Text) Then
                            ListView7.Items(g - 1).BackColor = Color.Orange
                            TextBox34.BackColor = Color.Orange
                        ElseIf CDec(.Items(g - 1).SubItems(2).Text) = CDec(TextBox30.Text) Then
                            .Items(g - 1).BackColor = Color.LightCyan
                            TextBox30.BackColor = Color.LightCyan
                        End If
                    Next g
                    Chart1.ChartAreas(0).AxisX.LineWidth = 1

                End With
            Else

                Timer1.Enabled = True
            End If
        Catch v As Exception
        End Try

    End Sub

    Private Sub Button32_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        Dim g As Integer
        If TextBox32.Text = "Mode : Fuel Cell Polarization" Then

            With ListView2
                For g = 1 To .Items.Count
                    If .Items(g - 1).BackColor = Color.LightGreen Then
                        Chart1.Series(0).Points(g - 1).Label = "max = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                    ElseIf .Items(g - 1).BackColor = Color.Yellow Then

                        Chart1.Series(0).Points(g - 1).Label = "Stable = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Yellow
                    ElseIf Val(.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then

                        Chart1.Series(0).Points(g - 1).Label = "Peaky = " & Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightBlue
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "min = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Orange
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "mean = " + .Items(g - 1).SubItems(2).Text & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightCyan
                    End If
                Next g
                Chart1.Series(0).SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No
                Chart1.Series(0).SmartLabelStyle.IsMarkerOverlappingAllowed = True
                Chart1.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.Right
            End With
        ElseIf TextBox32.Text = "Mode : Fuel Cell Power" Then
            With ListView3
                For g = 1 To .Items.Count
                    If .Items(g - 1).BackColor = Color.LightGreen Then
                        Chart1.Series(0).Points(g - 1).Label = "max = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                    ElseIf .Items(g - 1).BackColor = Color.Yellow Then

                        Chart1.Series(0).Points(g - 1).Label = "Stable = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Yellow
                    ElseIf Val(.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then

                        Chart1.Series(0).Points(g - 1).Label = "Peaky = " & Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightBlue
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "min = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Orange
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "mean = " + .Items(g - 1).SubItems(2).Text & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightCyan
                    End If
                Next g
                Chart1.Series(0).SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No
                Chart1.Series(0).SmartLabelStyle.IsMarkerOverlappingAllowed = True
                Chart1.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.Right
            End With
        ElseIf TextBox32.Text = "Mode : Hydrogen Usage" Then
            With ListView4
                For g = 1 To .Items.Count
                    If .Items(g - 1).BackColor = Color.LightGreen Then
                        Chart1.Series(0).Points(g - 1).Label = "max = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                    ElseIf .Items(g - 1).BackColor = Color.Yellow Then

                        Chart1.Series(0).Points(g - 1).Label = "Stable = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Yellow
                    ElseIf Val(.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then

                        Chart1.Series(0).Points(g - 1).Label = "Peaky = " & Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightBlue
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "min = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Orange
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "mean = " + .Items(g - 1).SubItems(2).Text & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightCyan
                    End If
                Next g
                Chart1.Series(0).SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No
                Chart1.Series(0).SmartLabelStyle.IsMarkerOverlappingAllowed = True
                Chart1.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.Right
            End With
        ElseIf TextBox32.Text = "Mode : Power Vs Hydrogen" Then
            With ListView8
                For g = 1 To .Items.Count
                    If .Items(g - 1).BackColor = Color.LightGreen Then
                        Chart1.Series(0).Points(g - 1).Label = "max = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                    ElseIf .Items(g - 1).BackColor = Color.Yellow Then

                        Chart1.Series(0).Points(g - 1).Label = "Stable = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Yellow
                    ElseIf Val(.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then

                        Chart1.Series(0).Points(g - 1).Label = "Peaky = " & Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightBlue
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "min = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Orange
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "mean = " + .Items(g - 1).SubItems(2).Text & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightCyan
                    End If
                Next g
                Chart1.Series(0).SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No
                Chart1.Series(0).SmartLabelStyle.IsMarkerOverlappingAllowed = True
                Chart1.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.Right
            End With
        ElseIf TextBox32.Text = "Mode : Power VS Partial Pressure" Then
            With ListView9
                For g = 1 To .Items.Count
                    If .Items(g - 1).BackColor = Color.LightGreen Then
                        Chart1.Series(0).Points(g - 1).Label = "max = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                    ElseIf .Items(g - 1).BackColor = Color.Yellow Then

                        Chart1.Series(0).Points(g - 1).Label = "Stable = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Yellow
                    ElseIf Val(.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then

                        Chart1.Series(0).Points(g - 1).Label = "Peaky = " & Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightBlue
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "min = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Orange
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "mean = " + .Items(g - 1).SubItems(2).Text & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightCyan
                    End If
                Next g
                Chart1.Series(0).SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No
                Chart1.Series(0).SmartLabelStyle.IsMarkerOverlappingAllowed = True
                Chart1.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.Right
            End With
        ElseIf TextBox32.Text = "Mode : Water Produced" Then
            With ListView5
                For g = 1 To .Items.Count
                    If .Items(g - 1).BackColor = Color.LightGreen Then
                        Chart1.Series(0).Points(g - 1).Label = "max = " + Format(CDbl(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                    ElseIf .Items(g - 1).BackColor = Color.Yellow Then

                        Chart1.Series(0).Points(g - 1).Label = "Stable = " + Format(CDbl(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Yellow
                        ' ElseIf Val(.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then

                        ' Chart1.Series(0).Points(g - 1).Label = "Peaky = " & Format(CDbl(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        ' Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        ' Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        ' Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightBlue
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "min = " + Format(CDbl(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Orange
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "mean = " + Format(CDbl(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightCyan
                    End If
                Next g
                Chart1.Series(0).SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Yes
                Chart1.Series(0).SmartLabelStyle.IsMarkerOverlappingAllowed = True
                Chart1.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.Right

            End With

        ElseIf TextBox32.Text = "Mode : Heat Generated" Then
            With ListView6
                For g = 1 To .Items.Count
                    If .Items(g - 1).BackColor = Color.LightGreen Then
                        Chart1.Series(0).Points(g - 1).Label = "max = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                    ElseIf .Items(g - 1).BackColor = Color.Yellow Then

                        Chart1.Series(0).Points(g - 1).Label = "Stable = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Yellow
                    ElseIf Val(.Items(g - 1).SubItems(3).Text) = Val(TextBox31.Text) Then

                        Chart1.Series(0).Points(g - 1).Label = "Peaky = " & Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightBlue
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox34.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "min = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Orange
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "mean = " + .Items(g - 1).SubItems(2).Text & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightCyan
                    End If
                Next g
                Chart1.Series(0).SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No
                Chart1.Series(0).SmartLabelStyle.IsMarkerOverlappingAllowed = True
                Chart1.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.Right
            End With

        ElseIf TextBox32.Text = "Mode : Efficiency" Then
            With ListView7
                For g = 1 To .Items.Count
                    If .Items(g - 1).BackColor = Color.LightGreen Then
                        Chart1.Series(0).Points(g - 1).Label = "max = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightGreen
                    ElseIf .Items(g - 1).BackColor = Color.Yellow Then

                        Chart1.Series(0).Points(g - 1).Label = "Stable = " + Format(Val(.Items(g - 1).SubItems(2).Text), "0.00E0") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Yellow
                    ElseIf Format(CDbl(.Items(g - 1).SubItems(3).Text.ToString), "0.00") = Format(CDbl(TextBox31.Text.ToString), "0.00") Then
                        Chart1.Series(0).Points(g - 1).Label = "Peaky = " & Format(Val(.Items(g - 1).SubItems(2).Text), "0.00") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightBlue
                    ElseIf CDbl(.Items(g - 1).SubItems(2).Text) = CDbl(TextBox34.Text) Then
                        Chart1.Series(0).Points(g - 1).Label = "min = " + Format(CDbl(.Items(g - 1).SubItems(2).Text), "0.00") & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.Orange
                    ElseIf .Items(g - 1).SubItems(2).Text = TextBox30.Text Then
                        Chart1.Series(0).Points(g - 1).Label = "mean = " + .Items(g - 1).SubItems(2).Text & Nam
                        Chart1.Series(0).Points(g - 1).MarkerStyle = MarkerStyle.Circle
                        Chart1.Series(0).Points(g - 1).MarkerSize = 10
                        Chart1.Series(0).Points(g - 1).MarkerColor = Color.LightCyan
                    End If
                Next g
                Chart1.Series(0).Points(.Items.Count - 1).Label = "min = " + Format(CDbl(.Items(.Items.Count - 1).SubItems(2).Text), "0.00") & Nam
                Chart1.Series(0).Points(.Items.Count - 1).MarkerStyle = MarkerStyle.Circle
                Chart1.Series(0).Points(.Items.Count - 1).MarkerSize = 8
                Chart1.Series(0).Points(.Items.Count - 1).MarkerColor = Color.Orange
                Chart1.Series(0).SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No
                Chart1.Series(0).SmartLabelStyle.IsMarkerOverlappingAllowed = True
                Chart1.Series(0).SmartLabelStyle.MovingDirection = LabelAlignmentStyles.Right
            End With
        Else

            Timer1.Enabled = True
        End If



    End Sub

    Private Sub Button34_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        Chart1.Series(0).Points.Clear()
        TextBox28.Text = ""
        TextBox27.Text = ""
        TextBox93.Text = ""
        TextBox94.Text = ""
    End Sub

    Private Sub Button39_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button39.Click
        TextBox29.Text = ""
        TextBox29.BackColor = Color.White
        TextBox30.Text = ""
        TextBox30.BackColor = Color.White
        TextBox31.Text = ""
        TextBox31.BackColor = Color.White
        TextBox33.Text = ""
        TextBox33.BackColor = Color.White
        TextBox34.Text = ""
        TextBox34.BackColor = Color.White
        TextBox35.Text = ""
        TextBox35.BackColor = Color.White
        TextBox44.BackColor = Color.White
        TextBox44.Text = ""
        TextBox43.Text = ""
        TextBox43.BackColor = Color.White
        TextBox45.Text = ""
        TextBox45.BackColor = Color.White
    End Sub

    Private Sub Button141_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button141.Click
        Dim N_cells As String = TextBox7.Text
        Dim A_cell As String = TextBox6.Text
        Dim Hyd_usage As String
        Dim F As String = TextBox2.Text
        'N = ListView2.Items.Count
        Chart1.Series(0).Points.Clear()
        Nam = " Watts"
        ListView8.Items.Clear()
        z = TextBox25.Text
        Chart1.Series(0).Name = "P Out"
        Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN

        Try
            For Me.baris = z + 1 To ListView3.Items.Count 'Step 0.01
                l = Me.ListView8.Items.Add("")
                For j As Integer = 1 To Me.ListView3.Columns.Count
                    l.SubItems.Add("")
                Next

                ListView8.Items(baris - 1).SubItems(0).Text = baris
                ListView8.Items(baris - 1).SubItems(2).Text = ListView3.Items(baris - 1).SubItems(2).Text ' CDec(P_Stack)
                ListView8.Items(baris - 1).SubItems(1).Text = ListView4.Items(baris - 1).SubItems(2).Text * 22.4 / 1000 'CDec(Hyd_usage)
                If Not baris < 2 Then
                    ListView8.Items(baris - 1).SubItems(3).Text = CDec(ListView8.Items(baris - 1).SubItems(2).Text) - CDec(ListView8.Items(baris - 2).SubItems(2).Text)
                Else
                    ListView8.Items(baris - 1).SubItems(3).Text = 0
                End If
                If ComboBox12.Text = "Real Time" Then

                    Chart1.Series(0).Points.AddXY(CDec(ListView8.Items(baris - 1).SubItems(1).Text.ToString), CDec(ListView8.Items(baris - 1).SubItems(2).Text))
                    If ComboBox17.Text = "Point" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                    ElseIf ComboBox17.Text = "Bar" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                    ElseIf ComboBox17.Text = "Area" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                    ElseIf ComboBox17.Text = "Fast Line" Then
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                    Else
                        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                    End If

                    If ComboBox13.Text = "Red" Then
                        Chart1.Series(0).Color = Color.Red
                    ElseIf ComboBox13.Text = "Green" Then
                        Chart1.Series(0).Color = Color.Green
                    ElseIf ComboBox13.Text = "Blue" Then
                        Chart1.Series(0).Color = Color.Blue
                    Else
                        Chart1.Series(0).Color = Color.Brown
                    End If
                    If ComboBox14.Text = "Dash" Then
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                        End With
                    Else
                        With Chart1.ChartAreas(0)
                            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                        End With
                    End If

                    wait(0.001)
                End If

                Chart1.ChartAreas("ChartArea1").AxisY.Title = ListView8.Columns(2).Text
                Chart1.ChartAreas("ChartArea1").AxisX.Title = ListView8.Columns(1).Text
                TextBox27.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum), "0.00E0")
                TextBox28.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum), "0.00E0")
                TextBox93.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum), "0.00E0")
                TextBox94.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum), "0.00E0")
            Next





            Button14.PerformClick()
            Button19.PerformClick()
            Button32.PerformClick()
            Button12.PerformClick()
            TextBox34.Text = Format(CDbl(TextBox34.Text), "0.00E0")
            TextBox35.Text = Format(CDbl(TextBox35.Text), "0.00E0")
            TextBox31.Text = Format(CDbl(TextBox31.Text), "0.00E0")
            TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
            TextBox45.Text = Format(CDbl(TextBox45.Text), "0.00E0")
            TextBox44.Text = Format(CDbl(TextBox44.Text), "0.00E0")
            TextBox43.Text = Format(CDbl(TextBox43.Text), "0.00E0")
            TextBox29.Text = Format(CDbl(TextBox29.Text), "0.00E0")
            TextBox30.Text = Format(CDbl(TextBox30.Text), "0.00E0")
            TextBox33.Text = Format(CDbl(TextBox33.Text), "0.00E0")

        Catch t As Exception
        End Try


        Me.Chart1.ChartAreas(0).AxisY.LabelStyle.Format = "0.00E0"
        Me.Chart1.ChartAreas(0).AxisX.LabelStyle.Format = "0.00E0"

    End Sub




    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Chart1.Series(0).Name = ""
        Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Maximum = 5
        Chart1.ChartAreas("ChartArea1").AxisY.Minimum = 0
        Chart1.ChartAreas("ChartArea1").AxisX.Minimum = 0
        Chart1.ChartAreas("ChartArea1").AxisX.Maximum = 5
        Chart1.Series(0).Points.AddXY(CDec(0), CDec(0))
        Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
        Chart1.Series(0).Color = Color.White
        TextBox32.Text = ""
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        If TabControl1.SelectedIndex = 2 Then
            Timer1.Enabled = False
            TextBox32.Text = "Mode : Fuel Cell Polarization"


        ElseIf TabControl1.SelectedIndex = 3 Then
            Timer1.Enabled = False
            TextBox32.Text = "Mode : Fuel Cell Power"

        ElseIf TabControl1.SelectedIndex = 4 Then
            Timer1.Enabled = False

            TextBox32.Text = "Mode : Hydrogen Usage"

        ElseIf TabControl1.SelectedIndex = 5 Then
            Timer1.Enabled = False
            TextBox32.Text = "Mode : Power Vs Hydrogen"


        ElseIf TabControl1.SelectedIndex = 6 Then
            Timer1.Enabled = False
            TextBox32.Text = "Mode : Power VS Partial Pressure"

        ElseIf TabControl1.SelectedIndex = 7 Then
            Timer1.Enabled = False
            TextBox32.Text = "Mode : Water Produced"
            wait(0.1)

        ElseIf TabControl1.SelectedIndex = 8 Then
            Timer1.Enabled = False
            TextBox32.Text = "Mode : Heat Generated"



        ElseIf TabControl1.SelectedIndex = 9 Then
            Timer1.Enabled = True
            TextBox32.Text = "Mode : Efficiency"

        Else
            Timer1.Enabled = True
            TextBox32.Text = "Mode : Efficiency"
        End If
        Timer1.Enabled = True
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TabPage3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub TextBox32_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox32.TextChanged
        TextBox27.Text = ""
        TextBox28.Text = ""
        TextBox93.Text = ""
        TextBox94.Text = ""

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        
        Timer1.Enabled = True
        TabControl1.SelectedIndex = 2
        Nam = " Volt"
        Button29.PerformClick()
       

        
        wait(5)
        TabControl1.SelectedIndex = 3
        Nam = " Watts"
        Button25.PerformClick()
        

      
        wait(5)
        TabControl1.SelectedIndex = 4
        Nam = " Watts"
        Button57.PerformClick()
      

        
        wait(5)
        TabControl1.SelectedIndex = 5
        Nam = " Watts"
        Button141.PerformClick()
        

       
          
        wait(5)
        TabControl1.SelectedIndex = 6
        Nam = " Watts"
        Button162.PerformClick()
        


       

            Timer1.Enabled = True

        wait(5)
        TabControl1.SelectedIndex = 7
        Nam = " litter/hour"
        Button78.PerformClick()
        

        
        wait(5)
        TabControl1.SelectedIndex = 8
        Nam = " Watts"
        Button99.PerformClick()
       


       
        wait(5)
        TabControl1.SelectedIndex = 9
        Nam = " Volts"
        Button120.PerformClick()
      



      
        wait(5)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TabControl1.SelectedIndex = 0
    End Sub

    Private Sub Label45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label45.Click

    End Sub

    Private Sub TabPage5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage5.Click

    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        SerialPort1.BaudRate = CmbBaud.SelectedItem
        SerialPort1.PortName = CmbScanPort.SelectedItem
        SerialPort1.Open()
        Dim i As Integer
        Dim baris As Integer
        Dim aryText() As String
        '  Timer1.Start()
        Try
            For baris = CInt(1) To 1000
                l = Me.ListView1.Items.Add("")
                For j As Integer = 1 To Me.ListView1.Columns.Count
                    l.SubItems.Add("")
                Next
                For Me.iterasi = 0 To 2




                    Dim k As String = SerialPort1.ReadExisting.ToString
                    aryText = k.Split(",")
                    'For i = 1 To UBound(aryText)
                    'If Not Val(k) = "999.000" Then
                    ListView1.Items(baris - 1).SubItems(2).Text = aryText(0)
                    ListView1.Items(baris - 1).SubItems(3).Text = aryText(1)
                    'Next i
                    ListView1.Items(baris - 1).SubItems(1).Text = Date.Now.ToString("HH:mm:ss")
                    ListView1.Items(baris - 1).SubItems(0).Text = baris
                    'End If

                    '  If Button2.Enabled = False Then
                    'Exit Sub

                    ' End If

                    wait(2)
                    ' SerialPort1.Close()
                Next
            Next
        Catch ex As Exception
        End Try
    End Sub
    ' End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        SerialPort1.Close()
        Timer1.Stop()
    End Sub

    Private Sub BtnScanPort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnScanPort.Click
        CmbScanPort.Items.Clear()
        Dim myPort As Array
        Dim i As Integer
        myPort = IO.Ports.SerialPort.GetPortNames
        CmbScanPort.Items.AddRange(myPort)
        i = CmbScanPort.Items.Count
        i = i - i
        Try
            CmbScanPort.SelectedIndex = i

        Catch ex As Exception
            Dim result As DialogResult
            result = MessageBox.Show("Com Port not detected", "Warning!!!", MessageBoxButtons.OK)
            CmbScanPort.Text = ""
            CmbScanPort.Items.Clear()
            Call Form1_Load(Me, e)
        End Try
        Button3.Enabled = True
        CmbScanPort.DroppedDown = True
    End Sub

    Private Sub Button31_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        Dim saveFileDialog1 As New SaveFileDialog()

        ' Sets the current file name filter string, which determines 
        ' the choices that appear in the "Save as file type" or 
        ' "Files of type" box in the dialog box.
        saveFileDialog1.Filter = "Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg|EMF (*.emf)|*.emf|PNG (*.png)|*.png|SVG (*.svg)|*.svg|GIF (*.gif)|*.gif|TIFF (*.tif)|*.tif"
        saveFileDialog1.FilterIndex = 2
        saveFileDialog1.RestoreDirectory = True

        ' Set image file format
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim format As ChartImageFormat = ChartImageFormat.Bmp

            If saveFileDialog1.FileName.EndsWith("bmp") Then
                format = ChartImageFormat.Bmp
            Else
                If saveFileDialog1.FileName.EndsWith("jpg") Then
                    format = ChartImageFormat.Jpeg
                Else
                    If saveFileDialog1.FileName.EndsWith("emf") Then
                        format = ChartImageFormat.Emf
                    Else
                        If saveFileDialog1.FileName.EndsWith("gif") Then
                            format = ChartImageFormat.Gif
                        Else
                            If saveFileDialog1.FileName.EndsWith("png") Then
                                format = ChartImageFormat.Png
                            Else
                                If saveFileDialog1.FileName.EndsWith("tif") Then
                                    format = ChartImageFormat.Tiff

                                End If
                            End If ' Save image
                        End If
                    End If
                End If
            End If
            Chart1.SaveImage(saveFileDialog1.FileName, format)
        End If
    End Sub

    Private Sub Chart1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Chart1.Click

    End Sub

    Private Sub Button36_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        With Chart1.ChartAreas(0)
            .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.DashDot
            '.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button35_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        With Chart1.ChartAreas(0)
            .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.DashDot
            '.AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
            '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
        End With
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
        Chart1.Series(0).Name = "V Out"
        'Dim tk As String
        Dim tk As String
        Dim P_H2 = TextBox4.Text


        Dim P_air As String = TextBox5.Text

        Dim alpha As String = TextBox9.Text
        Dim F As String = TextBox2.Text
        Dim R As String = TextBox1.Text

        Dim io As String = TextBox11.Text ^ TextBox12.Text

        Dim rin As String = TextBox8.Text

        Dim Bt As String = TextBox13.Text

        Dim Alpha1 As String = TextBox10.Text
        Dim k As String = TextBox17.Text
        Dim Gf_liq As String = TextBox14.Text
        Dim E_nernst As String
        Dim a As Integer
        Dim ag As String
        Dim bs As String
        Dim fx As String
        Dim fact As String
        Chart1.Series(0).Points.Clear()
        ListView10.Items.Clear()
        Nam = " Volt"
        Tc = TextBox3.Text
       

        '  N = TextBox24.Text / TextBox26.Text
        '    z = TextBox25.Text
        tk = TextBox3.Text + 273.15
        bs = 0
        fx = 1
        '  Try
        For Me.baris = 1 To CInt(100) 'Step 0.01


            ' For g As Integer = 1 To 30
            If baris - 1 = 0 Then

                fact = fx
            Else
                fx = Val((CInt(baris - 1) * fx))
                fact = Val(fx)
                'fx = fact
            End If
            ag = Val(-1 ^ (baris) * (-1) ^ (baris - 1) / (fact))
            bs = Val(bs) + Val(ag)
            'Next g

            'MsgBox(Val(bs))

            l = Me.ListView10.Items.Add("")
            For j As Integer = 1 To Me.ListView10.Columns.Count
                l.SubItems.Add("")
            Next
            ' Try
            '  For Me.iterasi = 2 To tipeA
            ListView10.Items(baris - 1).SubItems(1).Text = bs
            ListView10.Items(baris - 1).SubItems(0).Text = baris

            E_nernst = CStr((-Gf_liq / (2 * F)) - ((R * tk) * (1 + bs) / (2 * F)))
            V_out = CStr(E_nernst) ' - i * rin + Val(V_act) + Val(V_conc)
            ListView10.Items(baris - 1).SubItems(2).Text = (CStr(V_out)) 'V_out
            '  Else
            '      V_out = CDec(E_nernst) + CDec(V_ohmic) + CDec(V_act) + CDec(V_conc)
            'ListView10.Items(baris - 1).SubItems(3).Text = Math.Abs(Val(ListView10.Items(baris - 1).SubItems(2).Text) - Val(ListView10.Items(baris - 2).SubItems(2).Text))
            '  End If


            If ComboBox12.Text = "Real Time" Then

                Chart1.Series(0).Points.AddXY(CDec(ListView10.Items(baris - 1).SubItems(0).Text.ToString), CDec(ListView10.Items(baris - 1).SubItems(2).Text))
                Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
                Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
                Chart1.ChartAreas("ChartArea1").AxisY.Maximum = 1.23
                Chart1.ChartAreas("ChartArea1").AxisY.Minimum = 1.21
                Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
                Chart1.Series(0).Name = "V Out"
                If ComboBox17.Text = "Point" Then
                    Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                ElseIf ComboBox17.Text = "Bar" Then
                    Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                ElseIf ComboBox17.Text = "Area" Then
                    Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                ElseIf ComboBox17.Text = "Fast Line" Then
                    Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                Else
                    Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                End If

                If ComboBox13.Text = "Red" Then
                    Chart1.Series(0).Color = Color.Red
                ElseIf ComboBox13.Text = "Green" Then
                    Chart1.Series(0).Color = Color.Green
                ElseIf ComboBox13.Text = "Blue" Then
                    Chart1.Series(0).Color = Color.Blue
                Else
                    Chart1.Series(0).Color = Color.Brown
                End If
                If ComboBox14.Text = "Dash" Then
                    With Chart1.ChartAreas(0)
                        .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                    End With
                Else
                    With Chart1.ChartAreas(0)
                        .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                    End With
                End If

                wait(0.001)
            End If

            Chart1.ChartAreas("ChartArea1").AxisX.Title = ListView10.Columns(1).Text
            Chart1.ChartAreas("ChartArea1").AxisY.Title = ListView10.Columns(2).Text
            TextBox27.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum), "0.00E0")
            TextBox28.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum), "0.00E0")
            TextBox93.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum), "0.00E0")
            TextBox94.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum), "0.00E0")
        Next

        '  Catch t As Exception
        ' End Try

        ' Button14.PerformClick()
        ' Button19.PerformClick()
        'Button32.PerformClick()
        Button12.PerformClick()
        
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Maximum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisY.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
        Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
        Chart1.Series(0).Name = "V Out"
        'Dim tk As String
        Dim tk As String
        Dim P_H2 = TextBox4.Text


        Dim P_air As String = TextBox5.Text

        Dim alpha As String = TextBox9.Text
        Dim F As String = TextBox2.Text
        Dim R As String = TextBox1.Text

        Dim io As String = TextBox11.Text ^ TextBox12.Text

        Dim rin As String = TextBox8.Text

        Dim Bt As String = TextBox13.Text

        Dim Alpha1 As String = TextBox10.Text
        Dim k As String = TextBox17.Text
        Dim Gf_liq As String = TextBox14.Text
        Dim E_nernst As String
        Dim a As Integer
        Dim ag As String
        Dim bs As String
        Dim fx As String
        Dim fact As String
        Chart1.Series(0).Points.Clear()
        ListView10.Items.Clear()
        Nam = " Volt"
        Tc = TextBox3.Text
       

        '  N = TextBox24.Text / TextBox26.Text
        '    z = TextBox25.Text
        tk = TextBox3.Text + 273.15
        bs = 0
        fx = 1
        '  Try
        For Me.baris = 1 To CInt(100) 'Step 0.01


            ' For g As Integer = 1 To 30
            If baris - 1 = 0 Then

                fact = fx
            Else
                fx = Val((CInt(baris - 1) * fx))
                fact = Val(fx)
                'fx = fact
            End If
            ag = Val(-1 ^ (baris) * (-1) ^ (baris - 1) / (fact))
            bs = Val(bs) + Val(ag)
            'Next g

            'MsgBox(Val(bs))

            l = Me.ListView10.Items.Add("")
            For j As Integer = 1 To Me.ListView10.Columns.Count
                l.SubItems.Add("")
            Next
            ' Try
            '  For Me.iterasi = 2 To tipeA
            ListView10.Items(baris - 1).SubItems(1).Text = bs
            ListView10.Items(baris - 1).SubItems(0).Text = baris

            E_nernst = CStr((-Gf_liq / (2 * F)) - ((R * tk) * (1 + bs) / (2 * F)))
            V_out = CStr(E_nernst) ' - i * rin + Val(V_act) + Val(V_conc)
            ListView10.Items(baris - 1).SubItems(2).Text = (CStr(V_out)) 'V_out
            '  Else
            '      V_out = CDec(E_nernst) + CDec(V_ohmic) + CDec(V_act) + CDec(V_conc)
            'ListView10.Items(baris - 1).SubItems(3).Text = Math.Abs(Val(ListView10.Items(baris - 1).SubItems(2).Text) - Val(ListView10.Items(baris - 2).SubItems(2).Text))
            '  End If


            If ComboBox12.Text = "Real Time" Then

                Chart1.Series(0).Points.AddXY(CDec(ListView10.Items(baris - 1).SubItems(0).Text.ToString), CDec(ListView10.Items(baris - 1).SubItems(2).Text))
                Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Size = [Double].NaN
                Chart1.ChartAreas("ChartArea1").AxisY.ScaleView.Size = [Double].NaN
                Chart1.ChartAreas("ChartArea1").AxisY.Maximum = 1.23
                Chart1.ChartAreas("ChartArea1").AxisY.Minimum = 1.21
                Chart1.ChartAreas("ChartArea1").AxisX.Minimum = Double.NaN
                Chart1.ChartAreas("ChartArea1").AxisX.Maximum = Double.NaN
                Chart1.Series(0).Name = "V Out"
                If ComboBox17.Text = "Point" Then
                    Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Point
                ElseIf ComboBox17.Text = "Bar" Then
                    Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Bar
                ElseIf ComboBox17.Text = "Area" Then
                    Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Area
                ElseIf ComboBox17.Text = "Fast Line" Then
                    Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.FastLine

                Else
                    Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Spline
                End If

                If ComboBox13.Text = "Red" Then
                    Chart1.Series(0).Color = Color.Red
                ElseIf ComboBox13.Text = "Green" Then
                    Chart1.Series(0).Color = Color.Green
                ElseIf ComboBox13.Text = "Blue" Then
                    Chart1.Series(0).Color = Color.Blue
                Else
                    Chart1.Series(0).Color = Color.Brown
                End If
                If ComboBox14.Text = "Dash" Then
                    With Chart1.ChartAreas(0)
                        .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                    End With
                Else
                    With Chart1.ChartAreas(0)
                        .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Solid
                        '.AxisY2.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
                    End With
                End If

                wait(0.001)
            End If

            Chart1.ChartAreas("ChartArea1").AxisX.Title = ListView10.Columns(1).Text
            Chart1.ChartAreas("ChartArea1").AxisY.Title = ListView10.Columns(2).Text
            TextBox27.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Maximum), "0.00E0")
            TextBox28.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Maximum), "0.00E0")
            TextBox93.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisY.Minimum), "0.00E0")
            TextBox94.Text = Format(CDbl(Me.Chart1.ChartAreas("ChartArea1").AxisX.Minimum), "0.00E0")
        Next

        '  Catch t As Exception
        ' End Try

        ' Button14.PerformClick()
        ' Button19.PerformClick()
        'Button32.PerformClick()
        Button12.PerformClick()

    End Sub
End Class
